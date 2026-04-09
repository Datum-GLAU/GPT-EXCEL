
// ─── agent.mjs ──────────────────────────────────────────────────────────────
// Main entry point — the agentic loop that connects Claude to ExcelJS.
//
// Usage:
//   node agent.mjs "create a monthly budget tracker" budget.xlsx
//   node agent.mjs "add SUM formulas to column D" data.xlsx
// ─────────────────────────────────────────────────────────────────────────────

import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";
dotenv.config({ quiet: true });
if (!process.env.ANTHROPIC_API_KEY) {
  dotenv.config({ path: path.resolve(process.cwd(), '../.env'), quiet: true });
}

import Anthropic from "@anthropic-ai/sdk";
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
const MODEL = "claude-haiku-4-5-20251001";
const MAX_TOKENS = 1024;
const MAX_ITERATIONS = 20;
const GLOBAL_TIMEOUT_MS = 120_000; // 2 minutes
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 5000;

// ─── Prompt Caching Setup ──────────────────────────────────────────────────
// Add ephemeral cache_control to the last tool to enable schema caching
const toolDefinitions = [...rawToolDefs];
if (toolDefinitions.length > 0) {
  toolDefinitions[toolDefinitions.length - 1] = {
    ...toolDefinitions[toolDefinitions.length - 1],
    cache_control: { type: "ephemeral" },
  };
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
    "Usage: node agent.mjs \"<prompt>\" <file.xlsx>\n" +
      'Example: node agent.mjs "create a sales tracker" sales.xlsx'
  );
  process.exit(1);
}

const userPrompt = args[0];
const filePath = path.resolve(args[1]);

// ─── API key check ─────────────────────────────────────────────────────────

if (!process.env.ANTHROPIC_API_KEY) {
  console.error(
    "[error]  ANTHROPIC_API_KEY environment variable is not set.\n" +
      "         Run: export ANTHROPIC_API_KEY=sk-ant-..."
  );
  process.exit(1);
}

// ─── Main agent function ───────────────────────────────────────────────────

async function runAgent() {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  log("agent", `API key loaded: ${apiKey.substring(0, 12)}...${apiKey.substring(apiKey.length - 4)} (${apiKey.length} chars)`);
  const anthropic = new Anthropic({ apiKey });

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

  const messages = [
    {
      role: "user",
      content:
        `Task: ${userPrompt}\n\nFile: ${path.basename(filePath)}` +
        (fileExists ? " (existing file)" : " (new file — does not exist yet)") +
        fileContext,
    },
  ];

  // ── 4. Global timeout via AbortController ───────────────────────────────
  const controller = new AbortController();
  const globalTimer = setTimeout(() => {
    controller.abort();
  }, GLOBAL_TIMEOUT_MS);

  // ── 5. Agentic loop ────────────────────────────────────────────────────
  let iterations = 0;
  const readSheetToolIds = new Set(); // Track read_sheet calls for compression

  try {
    while (iterations < MAX_ITERATIONS) {
      if (controller.signal.aborted) {
        log("warn", "Global timeout reached (120 s) — stopping agent");
        break;
      }

      // Context compression: prune old read_sheet payloads from history
      // We skip the last two messages (the most recent assistant+user turn)
      if (messages.length > 3) {
        for (let i = 0; i < messages.length - 2; i++) {
          const msg = messages[i];
          if (msg.role === "user" && Array.isArray(msg.content)) {
            for (const block of msg.content) {
              if (block.type === "tool_result" && readSheetToolIds.has(block.tool_use_id)) {
                if (block.content && block.content.startsWith("[")) {
                  block.content = "[Data read successfully but omitted from history to save tokens. If needed, call read_sheet again.]";
                }
              }
            }
          }
        }
      }

      iterations++;
      log("think", `Claude is reasoning... (iteration ${iterations}/${MAX_ITERATIONS})`);

      // Call Claude (with retry for transient errors)
      let response;
      for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
          response = await anthropic.messages.create({
            model: MODEL,
            max_tokens: MAX_TOKENS,
            system: [
              {
                type: "text",
                text: SYSTEM_PROMPT,
                cache_control: { type: "ephemeral" },
              },
            ],
            tools: toolDefinitions,
            messages,
          });
          break; // success
        } catch (apiErr) {
          const status = apiErr.status ?? apiErr.statusCode ?? "unknown";
          const isRetryable = [429, 529, 503, 500].includes(Number(status));
          log("warn", `API call failed (attempt ${attempt}/${MAX_RETRIES}): ${status} — ${apiErr.message}`);
          if (isRetryable && attempt < MAX_RETRIES) {
            log("warn", `Retrying in ${RETRY_DELAY_MS / 1000}s...`);
            await new Promise((r) => setTimeout(r, RETRY_DELAY_MS * attempt));
            continue;
          }
          throw apiErr; // non-retryable or max retries hit
        }
      }

      // ── end_turn → done ─────────────────────────────────────────────────
      if (response.stop_reason === "end_turn") {
        // Print any final text content
        for (const block of response.content) {
          if (block.type === "text" && block.text) {
            log("agent", block.text);
          }
        }
        break;
      }

      // ── tool_use → execute tools ────────────────────────────────────────
      if (response.stop_reason === "tool_use") {
        const toolBlocks = response.content.filter((b) => b.type === "tool_use");
        const toolResults = [];

        // Print any text blocks that came alongside the tool calls
        for (const block of response.content) {
          if (block.type === "text" && block.text) {
            log("think", block.text);
          }
        }

        for (const toolBlock of toolBlocks) {
          const { id, name, input } = toolBlock;

          if (name === "read_sheet") {
            readSheetToolIds.add(id);
          }

          // Pretty-print the tool call
          const argsPreview = JSON.stringify(input, null, 2)
            .split("\n")
            .map((line, i) => (i === 0 ? line : "          " + line))
            .join("\n");
          log("tool", `→ ${name.padEnd(14)} ${argsPreview}`);

          // Execute
          const result = await executeTool(name, input, workbook);
          log("result", result);

          toolResults.push({
            type: "tool_result",
            tool_use_id: id,
            content: result,
          });
        }

        // Push assistant response + all tool results back into messages
        messages.push({ role: "assistant", content: response.content });
        messages.push({ role: "user", content: toolResults });

        continue;
      }

      // ── Unexpected stop reason ──────────────────────────────────────────
      log("warn", `Unexpected stop_reason: ${response.stop_reason}`);
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
    const status = err.status ?? err.statusCode ?? "";
    const body = err.error ? JSON.stringify(err.error).substring(0, 300) : "";
    log("error", `Agent failed: ${status ? `[${status}] ` : ""}${err.message}`);
    if (body) log("error", `Response: ${body}`);

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
