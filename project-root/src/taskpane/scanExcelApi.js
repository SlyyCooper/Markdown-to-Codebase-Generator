"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");

function scanExcelApi() {
  var excelApiMethods = [];
  for (var key in Excel) {
    if (Excel.hasOwnProperty(key)) {
      var property = Excel[key];
      if (typeof property === "function") {
        excelApiMethods.push("Function: " + key);
      } else if (typeof property === "object") {
        excelApiMethods.push("Object: " + key);
        for (var subKey in property) {
          if (property.hasOwnProperty(subKey) && typeof property[subKey] === "function") {
            excelApiMethods.push("  - Method: " + subKey);
          }
        }
      }
    }
  }
  fs.writeFileSync("excelApiMethods.txt", excelApiMethods.join("\n"), "utf8");
  console.log("Excel API methods have been written to excelApiMethods.txt");
}
scanExcelApi();
