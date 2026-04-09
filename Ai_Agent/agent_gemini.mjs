
// ─── agent_gemini.mjs ───────────────────────────────────────────────────────
// Main entry point — the agentic loop that connects Gemini 2.5 Flash to ExcelJS.
//
// Usage:
//   node agent_gemini.mjs "create a monthly budget tracker" budget.xlsx
// ─────────────────────────────────────────────────────────────────────────────

import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";
dotenv.config({ quiet: true });
if (!process.env.GEMINI_API_KEY) {
  dotenv.config({ path: path.resolve(process.cwd(), '../.env'), quiet: true });
}

import { GoogleGenAI } from "@google/genai";
import ExcelJS from "exceljs";

import { toolDefinitions as rawToolDefs } from "./tools.mjs";
import { executeTool } from "./executor.mjs";
import {
  log,
  snapshotWorkbook,
  restoreSnapshot,
  buildFileContext,
  sheetSummary,
  fileSizeKB,
} from "./utils.mjs";

// ─── Constants ─────────────────────────────────────────────────────────────
const MODEL = "gemini-2.5-flash";
const MAX_TOKENS = 1024;
const MAX_ITERATIONS = 20;
const GLOBAL_TIMEOUT_MS = 120_000; // 2 minutes
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 5000;

// ─── Schema Converter for Gemini ───────────────────────────────────────────
// Gemini's tool schema requires uppercase property types (e.g. "OBJECT", "STRING") 
// according to SDK typings, though standard JSON schema usually works.
function convertSchema(schema) {
  if (!schema) return schema;
  const newSchema = { ...schema };
  if (newSchema.type) newSchema.type = newSchema.type.toUpperCase();
  if (newSchema.items) newSchema.items = convertSchema(newSchema.items);
  if (newSchema.properties) {
    newSchema.properties = { ...newSchema.properties };
    for (const key in newSchema.properties) {
      newSchema.properties[key] = convertSchema(newSchema.properties[key]);
    }
  }
  return newSchema;
}

const geminiTools = [
  {
    functionDeclarations: rawToolDefs.map((t) => ({
      name: t.name,
      description: t.description,
      parameters: convertSchema(t.input_schema),
    })),
  },
];

function getResponseText(response) {
  if (!response) return "";
  const parts = response?.candidates?.[0]?.content?.parts;
  if (!Array.isArray(parts)) return "";
  return parts
    .filter((part) => typeof part?.text === "string" && part.text.trim())
    .map((part) => part.text)
    .join("\n")
    .trim();
}

// ─── System prompt ─────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are an expert Excel automation agent. You create and edit Excel files by calling tools.
Rules:
- CONTEXT WINDOW: You are only sent a tiny preview (e.g., 3-5 rows) of any given sheet to protect API token limits. Do not assume the sheet only has 5 rows. Look at the total row counts in the context.
- Always call read_sheet first if editing an existing file to understand its current state.
- Plan all steps mentally before executing — think about what sheets, headers, data, formulas, and formatting are needed.
- Call tools one logical group at a time. After writing data, call read_sheet to verify before adding formulas.
- For formulas: use Excel formula syntax (SUM, AVERAGE, IF, VLOOKUP, etc.). Always apply formulas to the full range of data (e.g. A2:A800), not just the sample you can see.
- For charts: always place them at least 2 columns to the right of the data.
- Always apply_formula for totals/summaries — never hardcode computed values.
- When done, call read_sheet one final time to confirm everything looks correct.
- Be thorough — a professional Excel file has: clear headers (bold, colored), formatted numbers, totals row, and clean layout.`;

// ─── CLI Argument Parsing ──────────────────────────────────────────────────

const args = process.argv.slice(2);
if (args.length < 2) {
  console.error(
    "Usage: node agent_gemini.mjs \"<prompt>\" <file.xlsx>\n" +
      'Example: node agent_gemini.mjs "create a sales tracker" sales.xlsx'
  );
  process.exit(1);
}

const userPrompt = args[0];
const filePath = path.resolve(args[1]);

// ─── API key check ─────────────────────────────────────────────────────────

if (!process.env.GEMINI_API_KEY) {
  console.error(
    "[error]  GEMINI_API_KEY environment variable is not set.\n" +
      "         Ensure .env or environment has GEMINI_API_KEY exported."
  );
  process.exit(1);
}

// ─── Main agent function ───────────────────────────────────────────────────

async function runAgent() {
  const apiKey = process.env.GEMINI_API_KEY;
  log("agent", `Gemini API key loaded: ${apiKey.substring(0, 10)}... (${apiKey.length} chars)`);
  
  // Initialize the new Gemini SDK Client
  const ai = new GoogleGenAI({ apiKey });

  // ── 1. Load or create workbook ──────────────────────────────────────────
  const workbook = new ExcelJS.Workbook();
  const fileExists = fs.existsSync(filePath);

  if (fileExists) {
    await workbook.xlsx.readFile(filePath);
    log("agent", `Starting agent for: ${path.basename(filePath)}`);
    log("context", `Loaded ${workbook.worksheets.length} sheet(s): ${sheetSummary(workbook)}`);
  } else {
    log("agent", `Starting agent for: ${path.basename(filePath)} (new file)`);
    log("context", "No existing file — will create from scratch");
  }

  // ── 2. Snapshot for rollback ────────────────────────────────────────────
  const snapshot = await snapshotWorkbook(workbook);

  // ── 3. Build initial context ────────────────────────────────────────────
  let fileContext = "";
  if (fileExists) {
    fileContext =
      "\n\n## Current file contents\n" +
      "Below is a JSON summary of every sheet in the file:\n```json\n" +
      buildFileContext(workbook) +
      "\n```";
  }

  // Gemini uses parts-based message formatting
  const messages = [
    {
      role: "user",
      parts: [
        {
          text:
            `Task: ${userPrompt}\n\nFile: ${path.basename(filePath)}` +
            (fileExists ? " (existing file)" : " (new file — does not exist yet)") +
            fileContext,
        },
      ],
    },
  ];

  // ── 4. Global timeout via AbortController ───────────────────────────────
  const controller = new AbortController();
  const globalTimer = setTimeout(() => {
    controller.abort();
  }, GLOBAL_TIMEOUT_MS);

  // ── 5. Agentic loop ────────────────────────────────────────────────────
  let iterations = 0;

  try {
    while (iterations < MAX_ITERATIONS) {
      if (controller.signal.aborted) {
        log("warn", "Global timeout reached (120 s) — stopping agent");
        break;
      }

      // Context compression for Gemini: prune old read_sheet JSON payloads
      if (messages.length > 3) {
        for (let i = 0; i < messages.length - 2; i++) {
          const msg = messages[i];
          if (msg.role === "user" && Array.isArray(msg.parts)) {
            for (const part of msg.parts) {
              if (part.functionResponse && part.functionResponse.name === "read_sheet") {
                if (JSON.stringify(part.functionResponse.response).length > 150) {
                   part.functionResponse.response = { message: "[Data read successfully but omitted from history to save tokens. If needed, call read_sheet again.]" };
                }
              }
            }
          }
        }
      }

      iterations++;
      log("think", `Gemini is reasoning... (iteration ${iterations}/${MAX_ITERATIONS})`);

      let response;
      for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
          response = await ai.models.generateContent({
            model: MODEL,
            contents: messages,
            config: {
              systemInstruction: SYSTEM_PROMPT,
              tools: geminiTools,
              maxOutputTokens: MAX_TOKENS,
              temperature: 0.1
            }
          });
          break; // success
        } catch (apiErr) {
          const status = apiErr.status ?? "unknown";
          const isRetryable = status === 429 || status === 503 || status === 500;
          log("warn", `API call failed (attempt ${attempt}/${MAX_RETRIES}): ${status} — ${apiErr.message}`);
          if (isRetryable && attempt < MAX_RETRIES) {
            log("warn", `Retrying in ${RETRY_DELAY_MS / 1000}s...`);
            await new Promise((r) => setTimeout(r, RETRY_DELAY_MS * attempt));
            continue;
          }
          throw apiErr; // non-retryable or max retries hit
        }
      }

      // ── tool_use → execute tools ────────────────────────────────────────
      if (response.functionCalls && response.functionCalls.length > 0) {
        const functionCalls = response.functionCalls;
        
        // Print any standard thought text that came along with tool calls
        const responseText = getResponseText(response);
        if (responseText) {
          log("think", responseText);
        }

        // Push model response containing function calls back into messages exactly as sent
        messages.push({
          role: "model",
          parts: functionCalls.map(call => ({ functionCall: call }))
        });

        const toolResponses = [];

        // Execute sequentially
        for (const call of functionCalls) {
          const { name, args } = call;

          // Pretty-print tool invocation
          const argsPreview = JSON.stringify(args, null, 2)
            .split("\n")
            .map((line, i) => (i === 0 ? line : "          " + line))
            .join("\n");
          log("tool", `→ ${name.padEnd(14)} ${argsPreview}`);

          const result = await executeTool(name, args, workbook);
          
          if (typeof result === 'string' && (result.startsWith('{') || result.startsWith('['))) {
             log("result", "JSON payload received.");
          } else {
             log("result", result);
          }

          let jsonResult;
          try {
            // functionResponse.response strictly dictates an object structure for Gemini APIs
            jsonResult = JSON.parse(result);
          } catch (e) {
            jsonResult = { message: result };
          }

          toolResponses.push({
            functionResponse: {
              name,
              // Gemini strictly requires this to be an Object (Struct), it cannot be an Array.
              response: (typeof jsonResult === 'object' && jsonResult !== null && !Array.isArray(jsonResult)) 
                          ? jsonResult 
                          : { result: jsonResult }
            }
          });
        }

        // Return tool results back to the model
        messages.push({
          role: "user",
          parts: toolResponses
        });

        continue;
      }

      // ── done ────────────────────────────────────────────────────────────
      const responseText = getResponseText(response);
      if (responseText) {
        log("agent", responseText);
        break;
      }
      
      log("warn", "Unexpected Gemini response (no text, no function calls).");
      break;
    }

    if (iterations >= MAX_ITERATIONS) {
      log("warn", "Max iterations reached — saving current state");
    }

    // ── 6. Save workbook ──────────────────────────────────────────────────
    await workbook.xlsx.writeFile(filePath);
    log("save", `Saved → ${path.basename(filePath)} (${fileSizeKB(filePath)} KB)`);
    log("done", `Sheets: ${sheetSummary(workbook)}`);
  } catch (err) {
    // ── Rollback on error ─────────────────────────────────────────────────
    log("error", `Agent failed: ${err.message}`);

    try {
      await restoreSnapshot(workbook, snapshot);
      await workbook.xlsx.writeFile(filePath);
      log("error", "Rolled back to previous file state");
    } catch (rollbackErr) {
      log("error", `Rollback also failed: ${rollbackErr.message}`);
    }

    process.exitCode = 1;
  } finally {
    clearTimeout(globalTimer);
  }
}

// ─── Run ───────────────────────────────────────────────────────────────────
runAgent();
