
// ─── utils.mjs ──────────────────────────────────────────────────────────────
// Shared helpers: workbook snapshotting, context extraction, range parsing,
// and pretty-print logging.
// ─────────────────────────────────────────────────────────────────────────────

import fs from "node:fs";

// ─── Logging helpers ───────────────────────────────────────────────────────
// Colour codes for terminal output (ANSI)
const CYAN    = "\x1b[36m";
const GREEN   = "\x1b[32m";
const YELLOW  = "\x1b[33m";
const RED     = "\x1b[31m";
const DIM     = "\x1b[2m";
const BOLD    = "\x1b[1m";
const RESET   = "\x1b[0m";

export function log(tag, msg) {
  const colours = {
    agent:   CYAN,
    context: YELLOW,
    think:   DIM,
    tool:    BOLD,
    result:  GREEN,
    save:    GREEN,
    done:    CYAN,
    error:   RED,
    warn:    YELLOW,
  };
  const c = colours[tag] ?? "";
  console.log(`${c}[${tag}]${RESET}   ${msg}`);
}

// ─── Workbook snapshot (for rollback) ──────────────────────────────────────
// We serialise the entire workbook to a Buffer. On rollback we write this
// buffer back to disk. This is simpler and more reliable than trying to
// deep-clone the in-memory ExcelJS data model.

/**
 * Creates a binary snapshot of the workbook that can be restored later.
 * @param {import('exceljs').Workbook} workbook
 * @returns {Promise<Buffer>}
 */
export async function snapshotWorkbook(workbook) {
  return workbook.xlsx.writeBuffer();
}

/**
 * Restores a workbook from a previously captured snapshot buffer.
 * @param {import('exceljs').Workbook} workbook
 * @param {Buffer} buffer
 */
export async function restoreSnapshot(workbook, buffer) {
  await workbook.xlsx.load(buffer);
}

// ─── Sheet context for initial prompt ──────────────────────────────────────
// Reads every sheet in the workbook and returns a structured summary that
// Claude can use to understand the current file state.

/**
 * Builds a human-readable context string describing all sheets.
 * @param {import('exceljs').Workbook} workbook
 * @param {number} maxRows  Maximum sample rows to include per sheet.
 * @returns {string}
 */
export function buildFileContext(workbook, maxRows = 3) {
  const sheets = [];

  workbook.eachSheet((ws) => {
    const rowCount = ws.rowCount;
    const rows = [];

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber <= maxRows) {
        rows.push(row.values.slice(1)); // ExcelJS row.values is 1-indexed; drop index 0
      }
    });

    sheets.push({
      name: ws.name,
      rowCount,
      headers: rows[0] ?? [],
      sampleData: rows.slice(0, maxRows),
    });
  });

  return JSON.stringify(sheets, null, 2);
}

/**
 * One-line summary of sheets for console output.
 * @param {import('exceljs').Workbook} workbook
 * @returns {string}
 */
export function sheetSummary(workbook) {
  const parts = [];
  workbook.eachSheet((ws) => {
    parts.push(`${ws.name} (${ws.rowCount} rows)`);
  });
  return parts.join(", ");
}

// ─── Range parser ──────────────────────────────────────────────────────────
// Converts an Excel-style range string like "A1:D5" into an array of
// individual cell addresses so we can iterate over them.

/**
 * Parses a range string and yields every cell address inside it.
 * Example: "A1:C3" → ["A1","B1","C1","A2","B2","C2","A3","B3","C3"]
 * @param {string} range  e.g. "A1:D5"
 * @returns {string[]}
 */
export function expandRange(range) {
  const [start, end] = range.split(":");
  if (!end) return [start]; // single cell

  const { col: c1, row: r1 } = parseCell(start);
  const { col: c2, row: r2 } = parseCell(end);

  const cells = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      cells.push(colToLetter(c) + r);
    }
  }
  return cells;
}

/**
 * Parses a cell address like "AB12" into { col: number, row: number }.
 * col is 1-indexed (A=1, B=2, …).
 */
export function parseCell(addr) {
  const match = addr.match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error(`Invalid cell address: "${addr}"`);
  const letters = match[1].toUpperCase();
  const row = parseInt(match[2], 10);
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { col, row };
}

/**
 * Converts a 1-indexed column number to its letter representation.
 * 1 → "A", 26 → "Z", 27 → "AA"
 */
export function colToLetter(num) {
  let s = "";
  while (num > 0) {
    num--;
    s = String.fromCharCode(65 + (num % 26)) + s;
    num = Math.floor(num / 26);
  }
  return s;
}

// ─── File size helper ──────────────────────────────────────────────────────

/**
 * Returns the size of a file in KB, formatted to 1 decimal place.
 * @param {string} filePath
 * @returns {string}
 */
export function fileSizeKB(filePath) {
  const stats = fs.statSync(filePath);
  return (stats.size / 1024).toFixed(1);
}