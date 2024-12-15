/* global Excel, console, Office */

export async function insertText(text: string) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export function setTaskpaneDimensions() {
  Office.context.ui.displayDialogAsync("about:blank", { width: 30, height: 40 }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      result.value.close();
    }
  });
}
