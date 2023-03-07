/**
* This function checks if the given input is a number or not.
*
* @param {Object} num The inputted argument, assumed to be a number.
* @return Returns a boolean reporting whether the input paramater is a number or not
* @author Jarren Ralf
*/
function removePrice()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1);
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++)
  {
    values[i][0] = values[i][0]
  }
}

/**
* This function 
*
* @author Jarren Ralf
*/
function addToItemList()
{
  const   SELECTION_COL  = 2;
  const        ITEM_COL  = 4;
  const         NUM_COLS = 1;
  const  NUM_HEADER_ROWS = 5;
  
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var finalDataRow = getLastRowSpecial(sheet.getRange("D:D").getValues()); // The final row of the items found in the search
  var activeRanges = sheet.getActiveRangeList().getRanges();               // The selected ranges corresponding to multiple item selections stored in an array
  var firstRows = [], lastRows = [], firstCols = [], lastCols = [], numRows = [];
  var itemVals = [[[]]];
  
  // Find the first & last row/col, number of rows, and item values in the set of all active ranges
  for (var r = 0; r < activeRanges.length; r++)
  {
    firstRows[r] = activeRanges[r].getRow();
    firstCols[r] = activeRanges[r].getColumn();
     lastRows[r] = activeRanges[r].getLastRow();
     lastCols[r] = activeRanges[r].getLastColumn();
      numRows[r] = lastRows[r] - firstRows[r] + 1;
     itemVals[r] = sheet.getRange(firstRows[r], SELECTION_COL, numRows[r], NUM_COLS).getValues();
  }
  
  var itemValues = [].concat.apply([], itemVals); // Concatenate all of the item values as a 2-D array
  var   numItems = itemValues.length;             // The number of items selected
  var        row = Math.min(...firstRows);        // This is the smallest starting row number out of all active ranges
  var        col = Math.min(...firstCols);        // This is the smallest starting col number out of all active ranges
  var    lastRow = Math.max( ...lastRows);        // This is the largest     final row number out of all active ranges
  var    lastCol = Math.max( ...lastCols);        // This is the largest     final col number out of all active ranges
  
  // Set the values if the selection is valid
  if (isInvalidSelection(row, col, lastRow, lastCol))
    ui.alert("Please make a selection in the first item list ONLY.")
  else
  {
    if (sheet.getRange("D6").isBlank()) // If the first position in the destination range is blank
      sheet.getRange(NUM_HEADER_ROWS + 1, ITEM_COL, numItems).setValues(itemValues);
    else if (isTooManyItems(finalDataRow, numItems))
      ui.alert("Number of items will exceed 100.\n\n Please select fewer items. ")
    else
      sheet.getRange(finalDataRow + 1, ITEM_COL, numItems).setValues(itemValues);
  }
}

/**
* This function checks if the user is properly selecting from the item list.
*
* @param {Number} row
* @param {Number} col
* @param {Number} lastRow
* @param {Number} lastCol
* @author Jarren Ralf
*/
function isInvalidSelection(row, col, lastRow, lastCol)
{
  return (row < 6 || col < 2 || lastRow > 105 || lastCol > 2);
}

/**
* This function checks if there are too many items that are being attempted to add to the list.
*
* @param {Number} lastRow
* @param {Number} numItems
* @author Jarren Ralf
*/
function isTooManyItems(lastRow, numItems)
{
  return (lastRow + numItems > 105);
}

/**
* This function clears the item list.
*
* @author Jarren Ralf
*/
function clear()
{
  SpreadsheetApp.getActive().getActiveSheet().getRange("D6:D105").clearContent();
}

/**
* This function 
*
* @author Jarren Ralf
*/
function generateLargeBarcodeLabels()
{
  const NUM_HEADER_ROWS = 5;
  const ITEM_COL = 4;
  const MAX_NUM_LABELS = 60;
  const NUM_LABELS_PER_PAGE = 6;
  const NUM_LABELS_PER_ROW = 2
  const   START_ROW = 2;
  const  LEFT_LABEL = 2;
  const RIGHT_LABEL = 4;
  const LABEL_JUMP  = 4; // The vertical translation of a label on the same piece of paper
  const   PAGE_JUMP = 1; // The vertical translation of the last row of labels to the first row on the next page
  
  var               ui = SpreadsheetApp.getUi();
  var      spreadsheet = SpreadsheetApp.getActive();
  var  itemSearchSheet = spreadsheet.getActiveSheet();
  var  largeLabelSheet = spreadsheet.getSheetByName("Large Barcode Labels")
  var rangeEntireSheet = largeLabelSheet.getRange("A:D");
  var       itemValues = itemSearchSheet.getRange("D6:D").getValues(); // The list of items to be printed on the chosen set of labels
  var  startOnLabelNum = itemSearchSheet.getRange("I2").getValue() - 1;
  
  if (itemSearchSheet.getRange("D6").isBlank())
    ui.alert("There are no items selected.")
  else
  {
    var lastRow_ItemCol = getLastRowSpecial(itemValues);
    var        numItems = lastRow_ItemCol;
    var          values = itemSearchSheet.getRange(NUM_HEADER_ROWS + 1, ITEM_COL, numItems, 8).getValues();
    var       numLabels = numItems + startOnLabelNum;
    var ii;
    
    if (numItems + startOnLabelNum > MAX_NUM_LABELS)
      ui.alert("Too many items selected.\n\nCan only generate 60 labels at one time.")
    else
    {
      largeLabelSheet.clear();
      
      // Reset the alignment of text and font parameters for the entire sheet
      rangeEntireSheet.setVerticalAlignment("middle");
      rangeEntireSheet.setHorizontalAlignment("center");
      rangeEntireSheet.setFontColor("black");
      rangeEntireSheet.setFontFamily("Arial");
      rangeEntireSheet.setFontLine("none");
      rangeEntireSheet.setFontSize(75);
      rangeEntireSheet.setFontStyle("normal");
      rangeEntireSheet.setFontWeight("normal");
      
      for (var i = startOnLabelNum; i < MAX_NUM_LABELS; i++)
      {
        ii = i - startOnLabelNum; // Re-indexed for readability
        
        if (i < numLabels)
        {
          if (i % NUM_LABELS_PER_ROW == 0) // If even index (Left Label)
            setBarcodeLabel(largeLabelSheet, START_ROW +       i/NUM_LABELS_PER_ROW*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE),  LEFT_LABEL, values[ii][0], values[ii][6], values[ii][1]); 
          else // If odd index (Right Label)
            setBarcodeLabel(largeLabelSheet, START_ROW + (i - 1)/NUM_LABELS_PER_ROW*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), RIGHT_LABEL, values[ii][0], values[ii][6], values[ii][1]);
        }
        else
          break;
      }
      
      // This a crude solution to fix a printing format issue for printing less than a full row of labels
      if (numLabels < NUM_LABELS_PER_ROW)
      {
        largeLabelSheet.getRange("D4").setFontColor("white");
        largeLabelSheet.getRange("D4").setValue("a");
      }
      
      // Take the user to the Quantity Record Labels sheet
      spreadsheet.setActiveSheet(largeLabelSheet.activate(), true);
    }
  }
}

/**
* This function 
*
* @author Jarren Ralf
*/
function generateQuantityRecordLabels()
{  
  const NUM_HEADER_ROWS = 5;
  const ITEM_COL = 4;
  const      MAX_NUM_LABELS = 60;
  const NUM_LABELS_PER_PAGE =  6;
  const NUM_LABELS_PER_ROW = 2;
  const   START_ROW =  2;
  const  LEFT_LABEL =  2;
  const RIGHT_LABEL =  8;
  const  LABEL_JUMP = 13; // The vertical translation of a label on the same piece of paper
  const   PAGE_JUMP =  1; // The vertical translation of the last row of labels to the first row on the next page
  
  var                  ui = SpreadsheetApp.getUi();
  var         spreadsheet = SpreadsheetApp.getActive();
  var     itemSearchSheet = spreadsheet.getActiveSheet();
  var qtyRecordLabelSheet = spreadsheet.getSheetByName("Quantity Record Labels")
  var    rangeEntireSheet = qtyRecordLabelSheet.getRange("A:L");
  var          itemValues = itemSearchSheet.getRange("D6:D").getValues(); // The list of items to be printed on the chosen set of labels
  
  if (itemSearchSheet.getRange("D6").isBlank())
    ui.alert("There are no items selected.")
  else
  {
    var lastRow_ItemCol = getLastRowSpecial(itemValues); // The last row of the list of items
    var numItems = lastRow_ItemCol;
    var values = itemSearchSheet.getRange(NUM_HEADER_ROWS + 1, ITEM_COL, numItems, 8).getValues();
    var numLabels = numItems;
    
    if (numLabels > MAX_NUM_LABELS)
      ui.alert("Too many items selected.\n\nCan only generate 60 labels at one time.")
    else
    {
      qtyRecordLabelSheet.clear(); // Clear all information on the sheet (Which includes all formatting)
      
      // Reset the alignment of text and font parameters for the entire sheet
      rangeEntireSheet.setVerticalAlignment("middle");
      rangeEntireSheet.setHorizontalAlignment("center");
      rangeEntireSheet.setFontColor("black");
      rangeEntireSheet.setFontFamily("Arial");
      rangeEntireSheet.setFontLine("none");
      rangeEntireSheet.setFontSize(10);
      rangeEntireSheet.setFontStyle("normal");
      rangeEntireSheet.setFontWeight("normal");
      
      // Set up the chosen number number of labels
      for (var i = 0; i < MAX_NUM_LABELS; i++)
      {
        if (i < numLabels)
        {
          if (i % NUM_LABELS_PER_ROW == 0) // If even index (Left Label)
            setQuantityRecordLabel(START_ROW +       i/NUM_LABELS_PER_ROW*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE),  LEFT_LABEL, values[i][6], values[i][1]); 
          else // If odd index (Right Label)
            setQuantityRecordLabel(START_ROW + (i - 1)/NUM_LABELS_PER_ROW*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), RIGHT_LABEL, values[i][6], values[i][1]);
        }
        else
          break;
      }
      
      // This a crude solution to fix a printing format issue for printing less than a full row of labels
      if (numLabels < NUM_LABELS_PER_ROW)
      {
        qtyRecordLabelSheet.getRange("L14").setFontColor("white");
        qtyRecordLabelSheet.getRange("L14").setValue("a");
      }
      
      // Take the user to the Quantity Record Labels sheet
      spreadsheet.setActiveSheet(qtyRecordLabelSheet.activate(), true);
    }
  }
}

/**
* This function 
*
* @author Jarren Ralf
*/
function generateSmallBarcodeLabels()
{
  const NUM_HEADER_ROWS = 5;
  const ITEM_COL = 4;
  const MAX_NUM_LABELS = 60;
  const NUM_LABELS_PER_ROW = 6;
  const   START_ROW = 2;
  const LABEL_JUMP  = 4; // The vertical translation of a label on the same piece of paper
  
  var              ui = SpreadsheetApp.getUi();
  var     spreadsheet = SpreadsheetApp.getActive();
  var itemSearchSheet = spreadsheet.getActiveSheet();
  var smallLabelSheet = spreadsheet.getSheetByName("Small Barcode Labels")
  var    rangeEntireSheet = smallLabelSheet.getRange("A:L");
  var          itemValues = itemSearchSheet.getRange("D6:D").getValues(); // The list of items to be printed on the chosen set of labels
  
  if (itemSearchSheet.getRange("D6").isBlank())
    ui.alert("There are no items selected.")
  else
  {
    var lastRow_ItemCol = getLastRowSpecial(itemValues);
    var        numItems = lastRow_ItemCol;
    var          values = itemSearchSheet.getRange(NUM_HEADER_ROWS + 1, ITEM_COL, numItems, 8).getValues();
    var       numLabels = numItems;
    
    if (numLabels > MAX_NUM_LABELS)
      ui.alert("Too many items selected.\n\nCan only generate 60 labels at one time.")
    else
    {
      smallLabelSheet.clear();
      
      // Reset the alignment of text and font parameters for the entire sheet
      rangeEntireSheet.setVerticalAlignment("middle");
      rangeEntireSheet.setHorizontalAlignment("center");
      rangeEntireSheet.setFontColor("black");
      rangeEntireSheet.setFontFamily("Arial");
      rangeEntireSheet.setFontLine("none");
      rangeEntireSheet.setFontSize(21);
      rangeEntireSheet.setFontStyle("normal");
      rangeEntireSheet.setFontWeight("normal");
      
      for (var i = 0; i < MAX_NUM_LABELS; i++)
      {
        if (i < numLabels)
        {
          if      (i % NUM_LABELS_PER_ROW == 0) 
            setBarcodeLabel(smallLabelSheet, START_ROW +       i/NUM_LABELS_PER_ROW*LABEL_JUMP,  2, values[i][0], values[i][6], values[i][1]);
          else if (i % NUM_LABELS_PER_ROW == 1)
            setBarcodeLabel(smallLabelSheet, START_ROW + (i - 1)/NUM_LABELS_PER_ROW*LABEL_JUMP,  4, values[i][0], values[i][6], values[i][1]); 
          else if (i % NUM_LABELS_PER_ROW == 2)
            setBarcodeLabel(smallLabelSheet, START_ROW + (i - 2)/NUM_LABELS_PER_ROW*LABEL_JUMP,  6, values[i][0], values[i][6], values[i][1]);
          else if (i % NUM_LABELS_PER_ROW == 3)
            setBarcodeLabel(smallLabelSheet, START_ROW + (i - 3)/NUM_LABELS_PER_ROW*LABEL_JUMP,  8, values[i][0], values[i][6], values[i][1]);
          else if (i % NUM_LABELS_PER_ROW == 4)
            setBarcodeLabel(smallLabelSheet, START_ROW + (i - 4)/NUM_LABELS_PER_ROW*LABEL_JUMP, 10, values[i][0], values[i][6], values[i][1]);
          else
            setBarcodeLabel(smallLabelSheet, START_ROW + (i - 5)/NUM_LABELS_PER_ROW*LABEL_JUMP, 12, values[i][0], values[i][6], values[i][1]);
        }
        else
          break;
      }
      
      // This a crude solution to fix a printing format issue for printing less than a full row of labels
      if (numLabels < NUM_LABELS_PER_ROW)
      {
        smallLabelSheet.getRange("L4").setFontColor("white");
        smallLabelSheet.getRange("L4").setValue("a");
      }
      
      // Take the user to the Small Barcode Labels sheet
      spreadsheet.setActiveSheet(smallLabelSheet.activate(), true);
    }
  }
}
  
/**
* Gets the last row number based on a selected column range values
*
* @param {Object[][]} range Takes a 2d array of a single column's values
* @returns {Number} The last row number with a value. 
*/
function getLastRowSpecial(range)
{
  var rowNum = 0;
  var blank = false;
  
  for(var row = 0; row < range.length; row++)
  {
    if(range[row][0] === "" && !blank)
    {
      rowNum = row;
      blank = true;
    }
    else if (range[row][0] !== "")
      blank = false;
  }
  return rowNum;
}

/**
* This function checks if the given input is a number or not.
*
* @param {Object} num The inputted argument, assumed to be a number.
* @return Returns a boolean reporting whether the input paramater is a number or not
* @author Jarren Ralf
*/
function isNumber(num)
{
  return !(isNaN(parseInt(num)));
}

// For somone editing the start on label cell
//function onEdit(e)
//{
//  var ui = SpreadsheetApp.getUi();
//  var startOnLabelRange = SpreadsheetApp.getActiveSheet().getRange("I2");
//  var startOnLabelValue = startOnLabelRange.getValue();
//  var editRange = e.range;
//  var editValues = editRange.getValues();
//  
//  if (startOnLabelValue == editRange) 
//  {
//    if (!isNumber(editValues))
//    {
//      ui.alert("Not a valid entry.");
//      startOnLabelRange.setValue(1);
//    }
//  }
//}

function onOpen(e)
{
  var startOnLabelRange = SpreadsheetApp.getActiveSheet().getRange("I2");
  var startOnLabelValue = startOnLabelRange.getValue();
  
  if (startOnLabelValue !== 1) 
    startOnLabelRange.setValue(1);
}

/**
 * Runs a BigQuery query and logs the results in a spreadsheet.
 */
function runQuery()
{
  // Replace this value with the project ID listed in the Google Cloud Platform project.
  var projectId = 'XXXXXXXX';

  var request = {query: 'SELECT TOP(word, 300) AS word, COUNT(*) AS word_count FROM publicdata:samples.shakespeare WHERE LENGTH(word) > 10;'};
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;

  // Check on status of the Query Job.
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete)
  {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  // Get all the rows of results.
  var rows = queryResults.rows;
  while (queryResults.pageToken)
  {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {pageToken: queryResults.pageToken});
    rows = rows.concat(queryResults.rows);
  }

  if (rows)
  {
    var spreadsheet = SpreadsheetApp.create('BiqQuery Results');
    var sheet = spreadsheet.getActiveSheet();

    // Append the headers.
    var headers = queryResults.schema.fields.map(function(field) {
      return field.name;
    });
    sheet.appendRow(headers);

    // Append the results.
    var data = new Array(rows.length);
    for (var i = 0; i < rows.length; i++)
    {
      var cols = rows[i].f;
      data[i] = new Array(cols.length);
      
      for (var j = 0; j < cols.length; j++)
        data[i][j] = cols[j].v;
    }
    sheet.getRange(2, 1, rows.length, headers.length).setValues(data);

    Logger.log('Results spreadsheet created: %s',
        spreadsheet.getUrl());
  }
  else
    Logger.log('No rows returned.');
  
}

/**
* This function prints one barcode label (Both Small and Large).
*
* @param     sheet     The label sheet
* @param      row      The row    that the label starts at
* @param      col      The column that the label starts at
* @param adagioDescrip The description of the item
* @param manualDescrip The description of the item
* @param   UPC_Code    The UPC code
* @author Jarren Ralf
*/
function setBarcodeLabel(sheet, row, col, adagioDescrip, manualDescrip, UPC_Code)
{
  sheet.getRange(row    , col).setWrap(true);
  sheet.getRange(row + 1, col).setWrap(true);
  sheet.getRange(row    , col).setValue(adagioDescrip); 
  sheet.getRange(row + 1, col).setValue(manualDescrip);
  
  // Set the barcode
  sheet.getRange(row + 2, col).setFormula(
    "=image(\"https://barcode.tec-it.com/barcode.ashx?data=" + 
     UPC_Code + 
    "&code=UPCA&multiplebarcodes=false&translate-esc=false&unit=Fit&dpi=96&imagetype=Gif&rotation=0&color=%23000000&bgcolor=%23ffffff&codepage=&qunit=Mm&quiet=0\", 2)");
}

/**
* This function prints one Quantity Record label.

* @param row      The row    that the label starts at
* @param col      The column that the label starts at
* @param descrip  The description of the item
* @param UPC_Code The UPC code of a particular item
* @author Jarren Ralf
*/
function setQuantityRecordLabel(row, col, descrip, UPC_Code)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName('Quantity Record Labels');
  
  // Merge all of the appropriate cells
  sheet.getRange(row    , col + 1,  1, 2).merge();       // The description cell
  sheet.getRange(row    , col + 3,  1, 2).merge();       // The barcode cell
  sheet.getRange(row + 1, col    , 11, 2).mergeAcross(); // The date cells
  
  // Set text wrap for the description cell
  sheet.getRange(row, col + 1).setWrap(true);
  
  // Set the borders of the label
  sheet.getRange(row + 1, col, 11, 5).setBorder(true, true, true, true, true, true);
  
  // Set the content of the label
  sheet.getRange(row    , col    ).setValue("Description:");
  sheet.getRange(row    , col + 1).setValue(descrip); 
  sheet.getRange(row + 1, col    ).setValue("DATE");
  sheet.getRange(row + 1, col + 2).setValue("IN"); 
  sheet.getRange(row + 1, col + 3).setValue("OUT"); 
  sheet.getRange(row + 1, col + 4).setValue("TOTAL");
  
  // Set the barcode
  sheet.getRange(row, col + 3).setFormula(
    "=image(\"https://barcode.tec-it.com/barcode.ashx?data=" + 
     UPC_Code + 
    "&code=UPCA&multiplebarcodes=false&translate-esc=false&unit=Fit&dpi=96&imagetype=Gif&rotation=0&color=%23000000&bgcolor=%23ffffff&codepage=&qunit=Mm&quiet=0\", 2)"); 
}

function updateData()
{
  var database = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');
  var csvData = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeLabels.jarren.csv").next().getBlob().getDataAsString());
  database.getRange('A:F').clearContent();
  database.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
}