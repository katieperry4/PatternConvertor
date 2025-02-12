/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("convert-pattern").onclick = convertPattern;
  }
});

export async function convertPattern() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      range.load("values");
      await context.sync();

      let writtenPattern: string[] = [];
      let stitchTuple: [string, string, string] = ["BS", "SC", "DC"];

      const values = range.values;

      for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
        let currentRow: string[] = [];
        let window = 1;

        for (let colIndex = 0; colIndex < values[rowIndex].length; colIndex++) {
          let currentValue = values[rowIndex][colIndex].toString();
          let nextValue = colIndex + 1 < values[rowIndex].length ? values[rowIndex][colIndex + 1].toString() : null;
          if (currentValue == nextValue) {
            window++;
          } else {
            switch (currentValue) {
              case "x":
                currentRow.unshift(window + stitchTuple[2]);
                break;
              case "":
                currentRow.unshift(window + stitchTuple[1]);
                break;
              case "bs":
                currentRow.unshift(window + stitchTuple[0]);
                break;
            }
            window = 1;
          }
        }
        writtenPattern.unshift(currentRow.join(" , "));
      }
      console.log(writtenPattern.join("\n"));
      const outputRange = sheet.getRange("A54");
      outputRange.values = [[writtenPattern.join("\n")]];
      outputRange.format.autofitColumns();
    });
  } catch (error) {
    console.log("there was an error" + error);
  }
}
