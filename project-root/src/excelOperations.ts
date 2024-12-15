/* global Excel, console */
import { Filters, PivotAggregationFunction, ChartType } from "./taskpane/components/App";

export async function writeToCellOperation(
  startCell: string,
  values: (string | number | boolean | null)[][]
): Promise<string> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(startCell).getResizedRange(values.length - 1, values[0].length - 1);
      range.values = values;
      await context.sync();

      range.load("address");
      await context.sync();
      return range.address;
    });
  } catch (error) {
    console.error("Error writing to Excel:", error);
    throw error;
  }
}

export async function readFromCellOperation(cellAddress: string): Promise<string> {
  return await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
    range.load("values");
    await context.sync();
    return range.values[0][0].toString();
  });
}

export async function formatCellOperation(
  cellAddress: string,
  format: {
    fontColor?: string;
    backgroundColor?: string;
    bold?: boolean;
  }
): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(cellAddress);
    if (format.fontColor) range.format.font.color = format.fontColor;
    if (format.backgroundColor) range.format.fill.color = format.backgroundColor;
    if (format.bold !== undefined) range.format.font.bold = format.bold;
    await context.sync();
  });
}

export async function addChartOperation(dataRange: string, chartType: Excel.ChartType): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(dataRange);
      const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
      chart.title.text = "New Chart";
      chart.legend.position = Excel.ChartLegendPosition.right;
      chart.height = 300;
      chart.width = 500;
      chart.setPosition("A15", "F30");
      await context.sync();
    });
  } catch (error) {
    console.error("Error in addChartOperation:", error);
    throw error;
  }
}

export async function getRangeData(rangeAddress?: string): Promise<{ address: string; values: any[][] }> {
  return await Excel.run(async (context) => {
    try {
      const range = rangeAddress
        ? context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress)
        : context.workbook.getSelectedRange();
      range.load(["address", "values"]);
      await context.sync();

      if (!range.values || range.values.length === 0) {
        throw new Error("The selected range is empty or contains no values.");
      }

      return {
        address: range.address,
        values: range.values,
      };
    } catch (error) {
      console.error("Error in getRangeData:", error);
      throw new Error("Failed to retrieve range data. Please ensure the range is valid and try again.");
    }
  });
}

export async function writeToSelectedRange(values: (string | number)[][]): Promise<string> {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["rowCount", "columnCount", "address"]);
    await context.sync();

    const trimmedValues = values.slice(0, range.rowCount).map((row) => row.slice(0, range.columnCount));
    range.values = trimmedValues;
    await context.sync();

    range.load("address");
    await context.sync();

    if (values.length > range.rowCount || values[0].length > range.columnCount) {
      return `${range.address} (Note: Data was trimmed to fit the range)`;
    }
    return range.address;
  });
}

export async function getSelectedRangeInfo(): Promise<{ address: string; rowCount: number; columnCount: number }> {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "rowCount", "columnCount", "values"]);
    await context.sync();
    return {
      address: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
    };
  });
}

export async function analyzeData(data: any[][], analysisType: string): Promise<string> {
  switch (analysisType) {
    case "summary":
      const allNumbers = data.flat().filter((val) => typeof val === "number");
      const sum = allNumbers.reduce((a, b) => a + b, 0);
      const avg = sum / allNumbers.length;
      const max = Math.max(...allNumbers);
      const min = Math.min(...allNumbers);
      return `Summary:\nSum: ${sum}\nAverage: ${avg}\nMax: ${max}\nMin: ${min}`;

    case "trend":
      const lastRow = data[data.length - 1].filter((val) => typeof val === "number");
      const firstRow = data[0].filter((val) => typeof val === "number");
      const trend = lastRow.map((val, index) => val - firstRow[index]);
      return `Trend (last value - first value):\n${trend.join(", ")}`;

    case "distribution":
      const flatData = data.flat();
      const distribution = flatData.reduce((acc: Record<string | number, number>, val) => {
        acc[val] = (acc[val] || 0) + 1;
        return acc;
      }, {});
      return `Distribution:\n${Object.entries(distribution)
        .map(([key, value]) => `${key}: ${value}`)
        .join("\n")}`;

    default:
      return "Unknown analysis type";
  }
}

export async function getCurrentSheetContent(
  options: {
    includeMetadata?: boolean;
    rowSeparator?: string;
    columnSeparator?: string;
    sheetName?: string;
  } = {}
): Promise<string> {
  const { includeMetadata = true, rowSeparator = "\n", columnSeparator = "\t", sheetName } = options;

  try {
    return await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["values", "address"]);
      sheet.load("name");
      await context.sync();

      const content = usedRange.values.map((row: any[]) => row.join(columnSeparator)).join(rowSeparator);
      if (includeMetadata) {
        return `Sheet ${sheet.name} (${usedRange.address}):${rowSeparator}${content}`;
      } else {
        return content;
      }
    });
  } catch (error) {
    console.error("Error getting sheet content:", error);
    throw error;
  }
}

export async function getWorksheetNames(): Promise<string[]> {
  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    return worksheets.items.map((sheet) => sheet.name);
  });
}

export async function ensureValidSelection(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();

    if (!range.address) {
      context.workbook.worksheets.getActiveWorksheet().getRange("A1").select();
      await context.sync();
    }
  });
}

export async function addPivotTableOperation(
  sourceDataRange: string,
  destinationCell: string,
  rowFields: string[],
  columnFields: string[],
  dataFields: Array<{ name: string; function: Excel.AggregationFunction }>,
  filterFields: string[] = []
): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceRange = sheet.getRange(sourceDataRange);
      const destinationRange = sheet.getRange(destinationCell);

      const pivotTable = sheet.pivotTables.add("NewPivotTable", sourceRange, destinationRange);
      rowFields.forEach((field) => pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field)));
      columnFields.forEach((field) => pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(field)));
      dataFields.forEach((field) => {
        const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(field.name));
        dataHierarchy.summarizeBy = field.function;
      });
      filterFields.forEach((field) => pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem(field)));

      await context.sync();
    });
  } catch (error) {
    console.error("Error in addPivotTableOperation:", error);
    throw error;
  }
}

export async function manageWorksheet(action: "create" | "delete", sheetName: string): Promise<string> {
  try {
    return await Excel.run(async (context) => {
      if (action === "create") {
        context.workbook.worksheets.add(sheetName);
        await context.sync();
        return `New worksheet "${sheetName}" has been created.`;
      } else if (action === "delete") {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.delete();
        await context.sync();
        return `Worksheet "${sheetName}" has been deleted.`;
      } else {
        throw new Error('Invalid action. Use "create" or "delete".');
      }
    });
  } catch (error) {
    console.error(`Error ${action}ing worksheet:`, error);
    throw error;
  }
}

export async function filterDataOperation(
  range: string | undefined,
  column: string,
  filterType: Filters,
  criteria: any
): Promise<{ range: string; filteredCount: number }> {
  return await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();
    let filterRange = range ? worksheet.getRange(range) : context.workbook.getSelectedRange();
    filterRange.load("address");
    await context.sync();

    let columnIndex = column.charCodeAt(0) - 65;
    let autoFilter = worksheet.autoFilter;
    autoFilter.apply(filterRange);

    switch (filterType) {
      case Filters.Equals:
        autoFilter.apply(filterRange, columnIndex, { filterOn: Excel.FilterOn.values, values: [criteria.value] });
        break;
      case Filters.GreaterThan:
        autoFilter.apply(filterRange, columnIndex, {
          filterOn: Excel.FilterOn.dynamic,
          dynamicCriteria: Excel.DynamicFilterCriteria.aboveAverage,
        });
        break;
      case Filters.LessThan:
        autoFilter.apply(filterRange, columnIndex, {
          filterOn: Excel.FilterOn.dynamic,
          dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage,
        });
        break;
      case Filters.Between:
        autoFilter.apply(filterRange, columnIndex, {
          filterOn: Excel.FilterOn.custom,
          criterion1: criteria.lowerBound.toString(),
          criterion2: criteria.upperBound.toString(),
          operator: Excel.FilterOperator.and,
        });
        break;
      case Filters.Contains:
        autoFilter.apply(filterRange, columnIndex, {
          filterOn: Excel.FilterOn.custom,
          criterion1: `*${criteria.value}*`,
        });
        break;
      case Filters.Values:
        autoFilter.apply(filterRange, columnIndex, { filterOn: Excel.FilterOn.values, values: criteria.values });
        break;
      default:
        throw new Error(`Unsupported filter type: ${filterType}`);
    }

    await context.sync();

    let visibleRange = filterRange.getVisibleView();
    visibleRange.load(["rowCount"]);
    await context.sync();

    return {
      range: filterRange.address,
      filteredCount: visibleRange.rowCount,
    };
  });
}

export async function sortDataOperation(
  range: string | undefined,
  sortFields: Excel.SortField[],
  matchCase?: boolean,
  hasHeaders?: boolean
): Promise<string> {
  return await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();
    let sortRange = range ? worksheet.getRange(range) : context.workbook.getSelectedRange();
    sortRange.load("address");
    await context.sync();

    sortRange.sort.apply(sortFields, matchCase, hasHeaders);
    await context.sync();
    return `Sorted range ${sortRange.address}`;
  });
}

export async function mergeCellsOperation(range: string, across: boolean = false): Promise<string> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeToMerge = sheet.getRange(range);
    rangeToMerge.merge(across);
    await context.sync();
    return `Merged cells in range ${range}`;
  });
}

export async function unmergeCellsOperation(range: string): Promise<string> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeToUnmerge = sheet.getRange(range);
    rangeToUnmerge.unmerge();
    await context.sync();
    return `Unmerged cells in range ${range}`;
  });
}

export async function autofitColumnsOperation(range: string): Promise<string> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeToAutofit = sheet.getRange(range);
    rangeToAutofit.format.autofitColumns();
    await context.sync();
    return `Auto-fitted columns in range ${range}`;
  });
}

export async function autofitRowsOperation(range: string): Promise<string> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeToAutofit = sheet.getRange(range);
    rangeToAutofit.format.autofitRows();
    await context.sync();
    return `Auto-fitted rows in range ${range}`;
  });
}

export async function applyConditionalFormat(
  range: string,
  formatType: Excel.ConditionalFormatType,
  rule: any,
  format?: Excel.ConditionalRangeFormat
): Promise<string> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeObject = sheet.getRange(range);
    const conditionalFormat = rangeObject.conditionalFormats.add(formatType);

    switch (formatType) {
      case Excel.ConditionalFormatType.cellValue:
        conditionalFormat.cellValue.rule = rule;
        if (format) Object.assign(conditionalFormat.cellValue.format, format);
        break;
      case Excel.ConditionalFormatType.colorScale:
        conditionalFormat.colorScale.criteria = rule;
        break;
      case Excel.ConditionalFormatType.dataBar:
        Object.assign(conditionalFormat.dataBar, rule);
        break;
      case Excel.ConditionalFormatType.iconSet:
        Object.assign(conditionalFormat.iconSet, rule);
        break;
      case Excel.ConditionalFormatType.presetCriteria:
        conditionalFormat.preset.rule = rule;
        if (format) Object.assign(conditionalFormat.preset.format, format);
        break;
      case Excel.ConditionalFormatType.containsText:
        conditionalFormat.textComparison.rule = rule;
        if (format) Object.assign(conditionalFormat.textComparison.format, format);
        break;
      case Excel.ConditionalFormatType.custom:
        conditionalFormat.custom.rule.formula = rule.formula;
        if (format) Object.assign(conditionalFormat.custom.format, format);
        break;
      default:
        break;
    }

    await context.sync();
    return `Conditional format applied to range ${range}`;
  });
}

export async function clearConditionalFormats(range: string): Promise<string> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeObject = sheet.getRange(range);
    rangeObject.conditionalFormats.clearAll();
    await context.sync();
    return `Conditional formats cleared from range ${range}`;
  });
}

export async function getActiveWorksheetName(): Promise<string> {
  return await Excel.run(async (context) => {
    const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load("name");
    await context.sync();
    return activeWorksheet.name;
  });
}
