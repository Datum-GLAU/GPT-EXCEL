
// ─── executor.mjs ───────────────────────────────────────────────────────────
// Tool implementations. Each function receives the ExcelJS workbook plus the
// tool arguments from Claude, performs the operation, and returns a
// human-readable result string (or error message).
// ─────────────────────────────────────────────────────────────────────────────

import { expandRange, parseCell, log } from "./utils.mjs";

// ─── 1. create_sheet ───────────────────────────────────────────────────────

/**
 * Creates a new worksheet and writes a styled header row.
 */
export function createSheet(workbook, { name, headers }) {
  if (!name) return { error: "Invalid args: name is required" };
  if (!headers || !Array.isArray(headers))
    return { error: "Invalid args: headers is required and must be an array" };

  // Guard against duplicate sheet names
  if (workbook.getWorksheet(name)) {
    return { error: `Sheet "${name}" already exists` };
  }

  const ws = workbook.addWorksheet(name);
  const headerRow = ws.addRow(headers);

  // Style: bold, blue fill (#4472C4), white text
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF4472C4" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });

  // Auto-fit-ish: set a reasonable minimum column width based on header length
  headers.forEach((h, i) => {
    ws.getColumn(i + 1).width = Math.max(12, String(h).length + 4);
  });

  return `✓ Sheet "${name}" created with ${headers.length} headers`;
}

// ─── 2. edit_cells ─────────────────────────────────────────────────────────

/**
 * Writes a 2-D data array starting at the given cell address.
 */
export function editCells(workbook, { sheet, startCell, data }) {
  if (!sheet) return { error: "Invalid args: sheet is required" };
  if (!startCell) return { error: "Invalid args: startCell is required" };
  if (!data || !Array.isArray(data))
    return { error: "Invalid args: data is required and must be a 2-D array" };

  const ws = workbook.getWorksheet(sheet);
  if (!ws) return { error: `Sheet "${sheet}" not found` };

  const { col: startCol, row: startRow } = parseCell(startCell);

  for (let r = 0; r < data.length; r++) {
    const rowData = data[r];
    if (!Array.isArray(rowData)) continue;
    for (let c = 0; c < rowData.length; c++) {
      ws.getCell(startRow + r, startCol + c).value = rowData[c];
    }
  }

  return `✓ Wrote ${data.length} rows starting at ${startCell}`;
}

// ─── 3. apply_formula ──────────────────────────────────────────────────────

/**
 * Writes Excel formula strings into the specified cells.
 */
export function applyFormula(workbook, { sheet, cells }) {
  if (!sheet) return { error: "Invalid args: sheet is required" };
  if (!cells || !Array.isArray(cells))
    return { error: "Invalid args: cells is required and must be an array" };

  const ws = workbook.getWorksheet(sheet);
  if (!ws) return { error: `Sheet "${sheet}" not found` };

  let count = 0;
  for (const { cell, formula } of cells) {
    if (!cell || !formula) {
      log("warn", `Skipping invalid formula entry: cell=${cell}, formula=${formula}`);
      continue;
    }
    // ExcelJS expects the formula without the leading "="
    const cleanFormula = formula.startsWith("=") ? formula.slice(1) : formula;
    ws.getCell(cell).value = { formula: cleanFormula };
    count++;
  }

  return `✓ Applied ${count} formula${count !== 1 ? "s" : ""}`;
}

// ─── 4. format_cells ───────────────────────────────────────────────────────

/**
 * Applies styling (bold, fontSize, bgColor, fontColor, numberFormat) to
 * every cell within the specified range.
 */
export function formatCells(workbook, { sheet, range, bold, fontSize, bgColor, fontColor, numberFormat }) {
  if (!sheet) return { error: "Invalid args: sheet is required" };
  if (!range) return { error: "Invalid args: range is required" };

  const ws = workbook.getWorksheet(sheet);
  if (!ws) return { error: `Sheet "${sheet}" not found` };

  const cellAddrs = expandRange(range);

  for (const addr of cellAddrs) {
    const cell = ws.getCell(addr);

    // Font properties — merge with existing font to avoid overwriting
    if (bold !== undefined || fontSize !== undefined || fontColor !== undefined) {
      const existingFont = cell.font ?? {};
      cell.font = {
        ...existingFont,
        ...(bold !== undefined && { bold }),
        ...(fontSize !== undefined && { size: fontSize }),
        ...(fontColor !== undefined && {
          color: { argb: `FF${fontColor}` },
        }),
      };
    }

    // Fill
    if (bgColor) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: `FF${bgColor}` },
      };
    }

    // Number format
    if (numberFormat) {
      cell.numFmt = numberFormat;
    }
  }

  return `✓ Formatted range ${range}`;
}

// ─── 5. add_chart ──────────────────────────────────────────────────────────

/**
 * Best-effort chart insertion using ExcelJS.  ExcelJS's chart support is
 * limited and experimental — if it fails we fall back to writing a comment
 * cell at the requested position so the user knows a chart was intended.
 */
export function addChart(workbook, { sheet, type, title, dataRange, position }) {
  if (!sheet) return { error: "Invalid args: sheet is required" };
  if (!type) return { error: "Invalid args: type is required" };
  if (!title) return { error: "Invalid args: title is required" };
  if (!dataRange) return { error: "Invalid args: dataRange is required" };
  if (!position) return { error: "Invalid args: position is required" };

  const ws = workbook.getWorksheet(sheet);
  if (!ws) return { error: `Sheet "${sheet}" not found` };

  // ExcelJS does not have a stable public chart API. We write a placeholder
  // cell so the user can add the chart manually in Excel.
  try {
    const cell = ws.getCell(position);
    cell.value = `📊 Chart: "${title}" (${type}) — Data: ${dataRange}. Open this file in Excel and insert the chart manually.`;
    cell.font = { italic: true, color: { argb: "FF888888" }, size: 10 };

    // Make column wide enough to show the message
    const { col } = parseCell(position);
    const currentWidth = ws.getColumn(col).width ?? 10;
    ws.getColumn(col).width = Math.max(currentWidth, 50);

    log("warn", `ExcelJS chart API is limited — wrote a placeholder at ${position}`);
    return `⚠ Chart "${title}" (${type}) — ExcelJS does not fully support charts. A placeholder note was written at ${position}. Please add the chart manually in Excel.`;
  } catch (err) {
    return `⚠ Chart placeholder failed: ${err.message}. Please add chart "${title}" manually.`;
  }
}

// ─── 6. read_sheet ─────────────────────────────────────────────────────────

/**
 * Returns current sheet data as a JSON string so the agent can inspect it.
 */
export function readSheet(workbook, { sheet, maxRows = 5 }) {
  if (!sheet) return { error: "Invalid args: sheet is required" };

  // Hard cap to protect context window size
  const limit = Math.min(maxRows, 20);

  const ws = workbook.getWorksheet(sheet);
  if (!ws) return { error: `Sheet "${sheet}" not found` };

  const rows = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber <= limit) {
      // Convert each cell to its display value (resolve formulas to their text)
      const vals = row.values.slice(1).map((v) => {
        if (v && typeof v === "object" && v.formula) {
          return `=${v.formula}` + (v.result !== undefined ? ` [=${v.result}]` : "");
        }
        return v;
      });
      rows.push(vals);
    }
  });

  return JSON.stringify(rows, null, 2);
}

// ─── Dispatcher ────────────────────────────────────────────────────────────
// Routes a tool call by name to the correct implementation. Validates that
// a handler exists and wraps every call in a timeout.

const TOOL_MAP = {
  create_sheet: createSheet,
  edit_cells: editCells,
  apply_formula: applyFormula,
  format_cells: formatCells,
  add_chart: addChart,
  read_sheet: readSheet,
};

const TOOL_TIMEOUT_MS = 10_000; // 10 seconds per tool execution

/**
 * Execute a tool by name with a per-call timeout.
 * @param {string} name       Tool name
 * @param {object} args       Tool arguments from Claude
 * @param {import('exceljs').Workbook} workbook  The active workbook
 * @returns {Promise<string>}  Result string or JSON error
 */
export async function executeTool(name, args, workbook) {
  const handler = TOOL_MAP[name];
  if (!handler) {
    return JSON.stringify({ error: `Unknown tool: ${name}` });
  }

  // Wrap in a timeout so a single runaway tool cannot hang the agent
  const result = await Promise.race([
    Promise.resolve().then(() => handler(workbook, args)),
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error("Tool timed out")), TOOL_TIMEOUT_MS)
    ),
  ]).catch((err) => {
    return JSON.stringify({ error: err.message });
  });

  // Normalise – always return a string
  return typeof result === "string" ? result : JSON.stringify(result);
}