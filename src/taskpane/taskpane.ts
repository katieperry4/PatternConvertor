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
      //I get the range that the user has selected
      const range = context.workbook.getSelectedRange();
      //get the sheet that the user is using
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      //get the values from the range
      range.load("values");
      //sync the context of the current sheet
      await context.sync();
      //define an array of strings to hold the pattern
      let writtenPattern: string[] = [];
      //the different crochet stitches since they're consistent
      let stitchTuple: [string, string, string] = ["BS", "SC", "DC"];
      //get the values from the selected range
      const values = range.values;
      //go through each row based on the length of the selection
      for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
        //string array to hold the stitches for the current row
        let currentRow: string[] = [];
        //define a sliding window at a length of 1 to begin with
        let window = 1;
        //go through each column in the selected row
        for (let colIndex = 0; colIndex < values[rowIndex].length; colIndex++) {
          //get the current value of the row/column combination
          let currentValue = values[rowIndex][colIndex].toString();
          //check out the value of the next column, if there isn't one, use null
          let nextValue = colIndex + 1 < values[rowIndex].length ? values[rowIndex][colIndex + 1].toString() : null;
          //if the value is the same, we extend the size of the window
          if (currentValue == nextValue) {
            window++;
          } else {
            //if the value is different we need to add the stitch (ex. 25SC) to the current row array
            //use unshift for how patterns are read
            switch (currentValue) {
              //if the current value is x it means DC
              case "x":
                currentRow.unshift(window + stitchTuple[2]);
                break;
              //if the current value is empty it means SC
              case "":
                currentRow.unshift(window + stitchTuple[1]);
                break;
              //if the current value is bs, it means border stitch
              case "bs":
                currentRow.unshift(window + stitchTuple[0]);
                break;
            }
            //reset the window back to 1 to start the next chunk of stitches
            window = 1;
          }
        }
        //after we go through a full row, we join the row as one string and add it to the
        //written pattern
        writtenPattern.unshift(currentRow.join(" , "));
      }
      //define where we want the output to be placed (I want it under the pattern)
      const outputStartRow = values.length + 5;
      //selects the start row/col and range where we want to start and where we want to end
      let outputRange = sheet.getRange(`B${outputStartRow}:B${outputStartRow + writtenPattern.length - 1}`);
      //selects the start row/col and range where we want the row counter to start and end
      let rowCounterRange = sheet.getRange(`A${outputStartRow}:A${outputStartRow + writtenPattern.length - 1}`);

      //map each row of the written pattern to a row
      let outputArray = writtenPattern.map((row) => [row]);
      //want each row to be numbered so I added row numbers and mapped them
      let rowNumbers = writtenPattern.map((_, index) => [index + 1]);

      //set the values of the range we defined earlier to the mapped outputArray
      outputRange.values = outputArray;
      //same for the range for the row numbers
      rowCounterRange.values = rowNumbers;

      //resize the rows and columns based on the output text size
      outputRange.format.autofitColumns();
      outputRange.format.autofitRows();
      rowCounterRange.format.autofitColumns();
      rowCounterRange.format.autofitRows();
    });
  } catch (error) {
    console.log("there was an error" + error);
  }
}
