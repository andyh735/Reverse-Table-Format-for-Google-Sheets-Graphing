/* When creating graphs in Sheets, you can noy have merged cells on only one column. This presents a challenge for data formatted in a table. If you have categorical/numerical descriptor of various data points merged in in column A, and matching data points in B, only the data point in the first row as the merged value in column A (where a new A value starts). Therefore, for data already formatted like this, you must unmerge all cells in column A, and manually fill in each empty cell to match A values with B values. 

In other words, if cells are merged in column A where the values of row(A:A)+1=row(A:A), and all cells in column B are unmerged, then only values in rows where row(A) satisfies the expression values (row(A) + 1 != row(A)).

This script completes this process automatically. At each row, the script checks if the current row's A value is blank. If it is not blank, the script copies the previous cell A value, and moves to the next row. If it is not blank, the script moves to the next row without making any changes. This script runs for a range that you specify (listed below). */

// NOTE: "Column A" and "Column B" in the notes refer to the items listed above, not the actual column A and B in the sheet. PLEASE read the above.

function unMerge() {

  // Helper functions definining the active sheet

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Enter parameters

  var unmergeentry = sheet.getRange('"insert range of values in column A as defined above"');
  var unmergerow = unmergeentry.getRow();

  // Unmerges all cells within the defined range in column A

  var unmerge = unmergeentry.breakApart();

  // Creates a cell counter

  var cellcounter = sheet.getRange('"insert Row number of first value in column A","insert Column number of first value in column A"').getRow();

  const lastrow = "insert row number of final value to check";

  // Gets first value for item in the top row defined in Column A

  var valuetocopy = sheet.getRange(cellcounter,unmergerow).getValue();
  
  // Moves the row down one to proceed to the next value
  
  var rowadd = cellcounter + 1;

  // Sets the variable setnewvalue with the value to check if the A values in subsequent cells are the same

  var setnewvalue = sheet.getRange(rowadd,unmergerow).setValue(valuetocopy);

  // Process repeats

  var cellcounter = cellcounter + 1;
  var checkr = cellcounter + 1;
  var checkrval = sheet.getRange(checkr,unmergerow).isBlank();
  console.log(checkrval);

  // Loop runs until last row has been reached

  while (cellcounter < lastrow) {

  // This loop checks to see if the succeeding cell is blank. If this the case, the loop repeats. if it is not, the value to check is then updated.
   
  while (checkrval == true) {
    var valuetocopy = sheet.getRange(cellcounter,unmergerow).getValue();
    var rowadd = cellcounter + 1;
    var setnewvalue = sheet.getRange(rowadd,unmergerow).setValue(valuetocopy);
    var cellcounter = cellcounter + 1;
    var checkr = cellcounter + 1;
    var checkrval = sheet.getRange(checkr,unmergerow).isBlank();
    console.log(checkrval);
  }
    var checkrvalvalue = sheet.getRange(checkr,unmergerow).getValue();
    console.log(checkrvalvalue);
    var refresh = sheet.getRange(checkr,unmergerow).setValue(checkrvalvalue);
    var cellcounter = cellcounter + 1;
    var checkrval = true;
  }

  // Task completed
  
console.log("Task Complete")
}