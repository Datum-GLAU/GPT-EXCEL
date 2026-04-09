
// ─── tools.mjs ──────────────────────────────────────────────────────────────
// JSON-schema tool definitions sent to Claude so it knows which functions it
// can call and what arguments each one expects.
// ─────────────────────────────────────────────────────────────────────────────

export const toolDefinitions = [
  // ── 1. create_sheet ──────────────────────────────────────────────────────
  {
    name: "create_sheet",
    description:
      "Creates a new worksheet with column headers. The header row is automatically styled bold with a blue (#4472C4) background and white text.",
    input_schema: {
      type: "object",
      properties: {
        name: {
          type: "string",
          description: "Name of the new worksheet.",
        },
        headers: {
          type: "array",
          items: { type: "string" },
          description: "Column header labels for the first row.",
        },
      },
      required: ["name", "headers"],
    },
  },

  // ── 2. edit_cells ────────────────────────────────────────────────────────
  {
    name: "edit_cells",
    description:
      "Writes a 2-D array of values into a cell range starting at the given cell address.",
    input_schema: {
      type: "object",
      properties: {
        sheet: {
          type: "string",
          description: "Target worksheet name.",
        },
        startCell: {
          type: "string",
          description: 'Top-left cell address (e.g. "A2").',
        },
        data: {
          type: "array",
          items: {
            type: "array",
            items: {},
          },
          description:
            "2-D array of values. Each inner array is one row.",
        },
      },
      required: ["sheet", "startCell", "data"],
    },
  },

  // ── 3. apply_formula ─────────────────────────────────────────────────────
  {
    name: "apply_formula",
    description:
      "Writes Excel formula strings into one or more cells. Use standard Excel formula syntax (SUM, AVERAGE, IF, VLOOKUP, etc.).",
    input_schema: {
      type: "object",
      properties: {
        sheet: {
          type: "string",
          description: "Target worksheet name.",
        },
        cells: {
          type: "array",
          items: {
            type: "object",
            properties: {
              cell: {
                type: "string",
                description: 'Cell address (e.g. "D14").',
              },
              formula: {
                type: "string",
                description:
                  'Excel formula WITHOUT the leading "=" sign (e.g. "SUM(D2:D13)").',
              },
            },
            required: ["cell", "formula"],
          },
          description: "Array of cell-formula pairs to apply.",
        },
      },
      required: ["sheet", "cells"],
    },
  },

  // ── 4. format_cells ──────────────────────────────────────────────────────
  {
    name: "format_cells",
    description:
      "Applies styling (bold, font size, background colour, font colour, number format) to every cell in a given range.",
    input_schema: {
      type: "object",
      properties: {
        sheet: {
          type: "string",
          description: "Target worksheet name.",
        },
        range: {
          type: "string",
          description: 'Cell range (e.g. "A1:D1").',
        },
        bold: {
          type: "boolean",
          description: "Make the font bold.",
        },
        fontSize: {
          type: "number",
          description: "Font size in points.",
        },
        bgColor: {
          type: "string",
          description:
            'Background fill colour as hex WITHOUT the "#" (e.g. "4472C4").',
        },
        fontColor: {
          type: "string",
          description:
            'Font colour as hex WITHOUT the "#" (e.g. "FFFFFF").',
        },
        numberFormat: {
          type: "string",
          description:
            'Excel number format code (e.g. "#,##0.00", "0%", "$#,##0").',
        },
      },
      required: ["sheet", "range"],
    },
  },

  // ── 5. add_chart ─────────────────────────────────────────────────────────
  {
    name: "add_chart",
    description:
      "Adds a chart to a worksheet. Supported types: bar, line, pie, column. If the chart type is unsupported by ExcelJS the tool will write a placeholder comment cell instead.",
    input_schema: {
      type: "object",
      properties: {
        sheet: {
          type: "string",
          description: "Target worksheet name.",
        },
        type: {
          type: "string",
          enum: ["bar", "line", "pie", "column"],
          description: "Chart type.",
        },
        title: {
          type: "string",
          description: "Chart title.",
        },
        dataRange: {
          type: "string",
          description:
            'Data range for the chart (e.g. "A1:B12").',
        },
        position: {
          type: "string",
          description:
            'Top-left cell where the chart should be placed (e.g. "E2").',
        },
      },
      required: ["sheet", "type", "title", "dataRange", "position"],
    },
  },

  // ── 6. read_sheet ────────────────────────────────────────────────────────
  {
    name: "read_sheet",
    description:
      "Returns the current contents of a worksheet as a JSON array of rows. Use this to inspect or verify data before and after edits.",
    input_schema: {
      type: "object",
      properties: {
        sheet: {
          type: "string",
          description: "Target worksheet name.",
        },
        maxRows: {
          type: "number",
          description:
            "Maximum number of rows to return (default 5). Keep this small (e.g., 5 or 10) to avoid overloading the context window! Only request more if absolutely necessary.",
        },
      },
      required: ["sheet"],
    },
  },
];