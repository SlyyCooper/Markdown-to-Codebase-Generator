import { ChartType, PivotAggregationFunction } from "../taskpane/components/App";

export const tools = [
  {
    type: "function",
    function: {
      name: "write_to_excel",
      description: "Write values to a range of cells in Excel",
      parameters: {
        type: "object",
        properties: {
          startCell: {
            type: "string",
            description: "The starting cell address (e.g., 'A1')",
          },
          values: {
            type: "array",
            items: {
              type: "array",
              items: { type: ["string", "number", "boolean", "null"] },
            },
            description: "The values to write as a 2D array.",
          },
        },
        required: ["startCell", "values"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "read_from_excel",
      description: "Read a value from a specific cell in Excel",
      parameters: {
        type: "object",
        properties: {
          cellAddress: { type: "string", description: "The cell address to read from (e.g., 'A1')." },
        },
        required: ["cellAddress"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "format_cell",
      description: "Format a cell in Excel",
      parameters: {
        type: "object",
        properties: {
          cellAddress: { type: "string", description: "The cell address to format (e.g. 'A1')" },
          fontColor: { type: "string", description: "The font color (e.g. '#FF0000')" },
          backgroundColor: { type: "string", description: "The background color (e.g. '#FFFF00')" },
          bold: { type: "boolean", description: "Whether to make the text bold" },
        },
        required: ["cellAddress"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "add_chart",
      description: "Add a chart to the Excel worksheet",
      parameters: {
        type: "object",
        properties: {
          dataRange: {
            type: "string",
            description: "The range of cells containing chart data"
          },
          chartType: {
            type: "string",
            enum: Object.values(ChartType),
            description: "The type of chart to create"
          },
        },
        required: ["dataRange", "chartType"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "analyze_selected_range",
      description: "Analyze the data in the currently selected range in Excel",
      parameters: {
        type: "object",
        properties: {
          analysisType: {
            type: "string",
            enum: ["summary", "trend", "distribution"],
            description: "The type of analysis to perform."
          }
        },
        required: ["analysisType"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "write_to_selected_range",
      description: "Write values to the currently selected range in Excel (trim if larger)",
      parameters: {
        type: "object",
        properties: {
          values: {
            type: "array",
            items: {
              type: "array",
              items: { type: "string" }
            },
            description: "2D array of strings to write."
          }
        },
        required: ["values"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "read_range",
      description: "Read values from a specific range or the currently selected range in Excel",
      parameters: {
        type: "object",
        properties: {
          rangeAddress: {
            type: "string",
            description: "The range address to read from. If omitted, reads current selection."
          }
        },
        required: [],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "add_pivot",
      description: "Add a pivot table to the Excel worksheet",
      parameters: {
        type: "object",
        properties: {
          sourceDataRange: { type: "string", description: "Source data range (e.g. 'A1:D10')" },
          destinationCell: { type: "string", description: "Cell to place the pivot table (e.g. 'G1')" },
          rowFields: {
            type: "array",
            items: { type: "string" },
            description: "Field names for row labels."
          },
          columnFields: {
            type: "array",
            items: { type: "string" },
            description: "Field names for column labels."
          },
          dataFields: {
            type: "array",
            items: {
              type: "object",
              properties: {
                name: { type: "string" },
                function: { type: "string", enum: Object.values(PivotAggregationFunction) }
              },
              required: ["name", "function"],
              additionalProperties: false
            },
            description: "Array of data fields with aggregation functions."
          },
          filterFields: {
            type: "array",
            items: { type: "string" },
            description: "Optional array of fields to use as filters."
          }
        },
        required: ["sourceDataRange", "destinationCell", "rowFields", "columnFields", "dataFields"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "manage_worksheet",
      description: "Create or delete a worksheet in Excel",
      parameters: {
        type: "object",
        properties: {
          action: { type: "string", enum: ["create", "delete"], description: "Action to perform" },
          sheetName: { type: "string", description: "Name of the sheet" }
        },
        required: ["action", "sheetName"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "filter_data",
      description: "Filter data in Excel based on specified criteria",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string", description: "Range to filter" },
          column: { type: "string", description: "Column to apply the filter (e.g. 'A')" },
          filterType: {
            type: "string",
            enum: ["Equals", "GreaterThan", "LessThan", "Between", "Contains", "Values"],
            description: "Type of filter to apply"
          },
          criteria: { type: "object", description: "The criteria for the filter" }
        },
        required: ["column", "filterType", "criteria"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "sort_data",
      description: "Sort data in an Excel range",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string", description: "Range to sort" },
          sortFields: {
            type: "array",
            items: {
              type: "object",
              properties: {
                key: { type: "number", description: "Column index (0-based)" },
                ascending: { type: "boolean", description: "true for ascending, false for descending" },
                color: { type: "string", description: "Color to sort by" },
                dataOption: { type: "string", enum: ["normal", "textAsNumber"] }
              },
              required: ["key", "ascending"],
              additionalProperties: false
            },
            description: "Sort criteria"
          },
          matchCase: { type: "boolean" },
          hasHeaders: { type: "boolean" }
        },
        required: ["sortFields"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "merge_cells",
      description: "Merge cells in a specified range",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string" },
          across: { type: "boolean" }
        },
        required: ["range"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "unmerge_cells",
      description: "Unmerge cells in a specified range",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string" }
        },
        required: ["range"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "autofit_columns",
      description: "Auto-fit columns in a specified range",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string" }
        },
        required: ["range"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "autofit_rows",
      description: "Auto-fit rows in a specified range",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string" }
        },
        required: ["range"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "apply_conditional_format",
      description: "Apply a conditional format to a range in Excel",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string" },
          formatType: {
            type: "string",
            description: "Type of conditional format"
          },
          rule: { type: "object" },
          format: { type: "object" }
        },
        required: ["range", "formatType", "rule"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "clear_conditional_formats",
      description: "Clear all conditional formats from a range",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string" }
        },
        required: ["range"],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_worksheet_names",
      description: "Get the names of all worksheets in the current workbook",
      parameters: {
        type: "object",
        properties: {},
        required: [],
        additionalProperties: false
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_active_worksheet_name",
      description: "Get the name of the currently active worksheet",
      parameters: {
        type: "object",
        properties: {},
        required: [],
        additionalProperties: false
      }
    }
  }
];
