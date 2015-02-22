/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  
  // Apparently this is the old way
  //var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //var entries = [{
  //  name : "Send Report",
  //  functionName : "createAndSendDocument"
  //}];
  //spreadsheet.addMenu("Nvent Menu", entries);
  
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Nvent Menu')
      .addItem('Send Report', 'createAndSendDocument')
      .addItem('Test Report', 'testReportAndAlert')
      .addToUi();
};


function testReportAndAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var testResults = '';
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control");
  if (sheet == null) {
    testResults = testResults + 'Sheet Named: "Control" is missing.\n';
  }
  var sendemail_hdr = 'SENDEMAIL';
  var datasheet_hdr = 'DATASHEET';
  var columns2sort_hdr = 'COLUMNS2SORT';
  var columns2write_hdr = 'COLUMNS2WRITE';
  var emails2send_hdr = 'EMAILS2SEND';
  var subject_hdr = 'SUBJECT';
  
  //TODO Add Tests To Check for Each Value
  //var settings = {};
  //settings.active = getValueAfterHeaderByName(sheet,sendemail_hdr);
  //settings.datasheet = getValueAfterHeaderByName(sheet,datasheet_hdr);
  //settings.columns2sort = getColumnAfterHeaderByName(sheet,columns2sort_hdr);
  //settings.columns2print = getColumnAfterHeaderByName(sheet,columns2write_hdr);
  //settings.emails = getColumnAfterHeaderByName(sheet,emails2send_hdr);
  //settings.subject = getValueAfterHeaderByName(sheet,subject_hdr);
  

  var result = ui.alert(testResults);

}

/**
 * Created by smorin on 2/20/15.
 */
//var columnBased = [['A',2,1,0,1],['B',0,1,2,3],['C','zebra','mellon','pie','apple']];
//var rowBased = [['A','B','C'],[2,0,'zebra'],[1,1,'mellon'],[0,2,'pie'],[1,3,'apple']];

// returns a column from an array of array that are
function getColumnAsArray(grid,columnIndex,offsetFromHead) {
    var colarray = [];
    if(offsetFromHead === undefined) {
        offsetFromHead = 0;
    }

    for(var i = 0; i < grid.length ; i++) {
        var value = grid[i][columnIndex];
        if(i >= offsetFromHead) {
            colarray.push(value);
        }
    }
    return colarray;
}

//var tmp = getColumnAsArray(columnBased,0);
////A,B,C
//tmp = getColumnAsArray(rowBased,0);
////A,2,1,1
//tmp = getColumnAsArray(rowBased,0,1);
////2,1,1

function arrayRemoveNulls(array) {
    var colarray = [];
    for(var i = 0; i < array.length ; i++) {
        var value = array[i];
        if(value != '') {
            colarray.push(value);
        }
    }
    return colarray;
}

function headerToIndex(headerArray,header) {
    for(var i = 0; i < headerArray.length ; i++) {
        var value = headerArray[i];
        if(value == header) {
           return i;
        }
    }
    return;
}

function headersToIndex(headerArray,headers) {
    var indexes = [];
    for(var headersIndex = 0; headersIndex < headers.length; headersIndex++) {
        for (var i = 0; i < headerArray.length; i++) {
            var value = headerArray[i];
            if (value == headers[headersIndex]) {
                indexes.push(i);
                break;
            }
        }
    }
    return indexes;
}

function removeHeader(rowBasedGrid) {
    return rowBasedGrid.slice(1);
}

function compareArrays(a,b,indexes) {
    if(indexes.length <= 0) {
        return 0;  // return equal do not change sort if all equal
    }
    var index = indexes[0];
    var first = a[index];
    var second = b[index];
    if(first == second) {
        if (indexes.length > 1) {
            return compareArrays(a, b, indexes.splice(1));
        } else {
            return 0;
        }
    } else {
        return first < second ? -1 : first > second ? 1 : 0;
    }
}

function sortGrid(grid,arrayOfIndexsToSort) {
    var sortedgrid = grid.sort(function(a, b)
    {
        return compareArrays(a,b,arrayOfIndexsToSort);
    });
    return sortedgrid;
}

function deepCopy(obj) {
    if (Object.prototype.toString.call(obj) === '[object Array]') {
        var out = [], i = 0, len = obj.length;
        for ( ; i < len; i++ ) {
            out[i] = arguments.callee(obj[i]);
        }
        return out;
    }
    if (Object.prototype.toString.call(obj) === '[object Date]') {
        return new Date(obj.getTime());
    }
    if (typeof obj === 'object') {
        var out = {}, i;
        for ( i in obj ) {
            out[i] = arguments.callee(obj[i]);
        }
        return out;
    }
    return obj;
}

function transposeGrid(grid) {
    var mygrid = deepCopy(grid);
    var transposed = [];

    for(var i = 0; i < mygrid.length; i++){
        for(var j = 0; j < mygrid[i].length; j++){
            if(j>=transposed.length) {
                transposed.push([]);
            }
            transposed[j].push(mygrid[i][j]);
        };
    };
    return transposed;
}

//var tmp = transposeGrid(rowBased);

function filterAndOrderRowBasedGridByIndex(grid,indexes) {
    var mygrid = transposeGrid(grid);
    var currentindex = 0;
    var indexmap = {};
    mygrid = mygrid.filter(function(element, index, array) {

        if(indexes.indexOf(index) != -1) {
            indexmap[index] = currentindex;
            currentindex++;
            return true;
        } else {
            return false;
        }
    });
    var finalgrid = [];
    for(var i = 0; i < indexes.length;i++) {
        finalgrid[i] = mygrid[indexmap[indexes[i]]];
    }
    // Move the rows according to the indexmap
    return transposeGrid(finalgrid);
}

//tmp = filterAndOrderRowBasedGridByIndex(rowBased,headersToIndex(rowBased[0],['C','B']));
//
//tmp = sortGrid(removeHeader(rowBased),[2,0]);
//var z = [1];

/*
End of File
 */



function getColumnIndex(Sheet,ColumnName) {
  var range = Sheet.getDataRange();
  var numRows = range.getNumRows();
  var gridvalues = range.getValues();
  if(numRows > 0) {
    for(var i = 0; i < gridvalues[0].length; i++) {
      if(ColumnName == gridvalues[0][i]){
        return i;
      }
    }
  } else {
    return null;
  }
}

//  getColumnAfterHeader
//  Arguments
//  - Sheet: A google Sheet object (rowBased)
//  - ColumnIndex: integer for the index to get
//  Returns:
function getColumnAfterHeader(Sheet,ColumnIndex) {
  var range = Sheet.getDataRange();
  var numRows = range.getNumRows();
  var gridvalues = range.getValues();
  return range.offset(1, ColumnIndex, numRows, 1).getValues();
}

function getColumnAsArray(grid,columnIndex,offsetFromHead) {
  var colarray = [];
  if(offsetFromHead === undefined) {
    offsetFromHead = 0;
  }
  
  for(var i = 0; i < grid.length ; i++) {
   var value = grid[i][0];
   if(value != '' && i >= offsetFromHead) {
     colarray.push(value);
   }
  }
  return colarray;
}

function getColumnAfterHeaderByName(Sheet,ColumnName) {
  var range = Sheet.getDataRange();
  var numRows = range.getNumRows();
  var gridvalues = range.getValues();
  var columnIndex = getColumnIndex(Sheet,ColumnName);
  var subgrid = range.offset(1, columnIndex, numRows, 1).getValues();
  var colarray = [];
   for(var i = 0; i < subgrid.length ; i++) {
    var value = subgrid[i][0];
    if(value != '') {
      colarray.push(value);
    }
  }
  return colarray;
}

function getValueAfterHeaderByName(Sheet,ColumnName) {
  var range = Sheet.getDataRange();
  var numRows = range.getNumRows();
  var gridvalues = range.getValues();
  var columnIndex = getColumnIndex(Sheet,ColumnName);
  var subgrid = range.offset(1, columnIndex, numRows, 1).getValues();
  var colarray = [];
   for(var i = 0; i < subgrid.length ; i++) {
    var value = subgrid[i][0];
    if(value != '') {
      colarray.push(value);
    }
  }
  if(colarray.length > 0) {
    return colarray[0];
  } else {
    return '';
  }
}

function getHeaders(grid) {
  var colarray = [];
  for(var i = 0; i < grid.length ; i++) {
  }
}

// Properties:
//  active - true/false
//  datasheet - String
//  columns2sort - Array<String(column names)>
//  columns2print - Array<String(column names)>
//  emails - Array<String(emails)>
//  subject - String
//  filters - Array<Object{column<String>,type<String(include,exclude)>,values<Array(Strings)>}>
function getControlSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control");
  var sendemail_hdr = 'SENDEMAIL';
  var datasheet_hdr = 'DATASHEET';
  var columns2sort_hdr = 'COLUMNS2SORT';
  var columns2write_hdr = 'COLUMNS2WRITE';
  var emails2send_hdr = 'EMAILS2SEND';
  var subject_hdr = 'SUBJECT';
  var filter_hdr = 'FILTER';
  
  var settings = {};
  
  settings.active = getValueAfterHeaderByName(sheet,sendemail_hdr);
  if(settings.active.toLocaleLowerCase() == 'no' || settings.active.toLocaleLowerCase() == 'false' || settings.active.toLocaleLowerCase() == 'off') {
    settings.active = false;
  } else {
    settings.active = true;
  }
  
  settings.datasheet = getValueAfterHeaderByName(sheet,datasheet_hdr);
  settings.columns2sort = getColumnAfterHeaderByName(sheet,columns2sort_hdr);
  settings.columns2print = getColumnAfterHeaderByName(sheet,columns2write_hdr);
  settings.emails = getColumnAfterHeaderByName(sheet,emails2send_hdr);
  settings.subject = getValueAfterHeaderByName(sheet,subject_hdr);
  settings.filters = getColumnAfterHeaderByName(sheet,filter_hdr);
  var filterobjs = [];
  var saveException = undefined;
  for(var i = 0; i < settings.filters.length; i++) {
    Logger.log('Parsing Filter:'+settings.filters[i]);
    try {
      filterobjs.push(JSON.parse(settings.filters[i]));
    } catch (e) {
      Logger.log('Had error:'+e);
      saveException = e;
      if(e instanceof SyntaxError) {
        try {
          Logger.log('Popping up UI.Alert');
          //var ui = SpreadsheetApp.getUi(); // Same variations.
          //var testResults = 'Problem Parsing Filter:"'+settings.filters[i]+'"\nMSG:'+e.message;
          //var result = ui.alert(testResults);
        } catch(e){
          Logger.log(e);
        }
      }
    }
    if(saveException != undefined) {
      Logger.log('Throwing Syntax Exception');
      throw new SyntaxError('Problem Parsing JSON from FILTER column in Control:'+saveException.message,saveException.filename,saveException.linenumber);
    } else {
      Logger.log('Parsing Filters');
    }
  }
  

  
  settings.filters = filterobjs;
  
  Logger.log(settings)
  return settings;
}

function applyFilter(array,indexMap,filter) {
  // {"column":"Status","type":"exclude","values":["Done"]}
  try {
    if(filter.type == 'exclude') {
      var value = array[indexMap[filter.column]];
      if ("values" in filter){
        //property exists
        for(var i = 0; i < filter.values.length;i++) {
          if(value == filter.values[i]) {
            return false;
          }
        }
      }
    } else {
      // Should be include
    }
      } catch(e) {
        Logger.log('Problem with a filter('+JSON.stringify(filter)+'):'+e);
    }
}

function applyFiltersRowBasedGridByIndex(grid,indexMap,filters) {
    var mygrid = deepCopy(grid);

    mygrid = mygrid.filter(function(element, index, array) {
      var keepRecord = true;
      var filterResult = undefined;
      for(var i = 0; i < filters.length; i++) {
        filterResult = applyFilter(element,indexMap,filters[i]);
        if(filterResult != undefined) {
          return filterResult;
        }
      }
      return keepRecord;
    });
    return mygrid;
}

function getSortedAndFilteredData(settings) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.datasheet);
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var gridvalues = range.getValues();
  //Task	Area	Type	Owner	Due Date	Status
  //PRINT: Owner, Status, Due Date, Task
  //SORT: Owner, Status, Due Date
  
  // First row is the headers
  var headers = deepCopy(gridvalues[0]);
  // create a map from headers to their index
  var headersMap = {};
  for(var i = 0; i < headers.length; i++) {
    headersMap[headers[i]]=i;
  }
  // Header Indexes to Print as an Array
  var headers_indexes = headersToIndex(headers,settings.columns2print);
  // Header Indexes to Sort as an Array
  var sort_indexes = headersToIndex(headers,settings.columns2sort);
  
  gridvalues = applyFiltersRowBasedGridByIndex(gridvalues,headersMap,settings.filters);
  gridvalues = sortGrid(removeHeader(gridvalues),sort_indexes);
  gridvalues = filterAndOrderRowBasedGridByIndex(gridvalues,headers_indexes);

  return gridvalues;
}


// ***** EXAMPLE CONTROL SHEET *****
//SUBJECT	SENDEMAIL	EMAILS2SEND	COLUMNS2WRITE	COLUMNS2SORT	DATASHEET	FILTER
//Daily Partner Task List Report	Yes	steve@nvent.solutions	Owner	Owner	Task list	{"column":"Status","type":"exclude","values":["Done"]}
//		steve.morin@gmail.com	Status	Status		
//			Due Date	Due Date		
//			Task			
function createAndSendDocument() {
  
  // Get a list of emails to send to
  var settings = getControlSettings();
  
  // Checks the settings to see if the report is disabled.
  if(settings.active == false) {
    return;
  }
  var data = getSortedAndFilteredData(settings);
  
  // Pseudo Code
  // get settings object
  // SUBJECT	SENDEMAIL	EMAILS2SEND	COLUMNS2WRITE	COLUMNS2SORT	DATASHEET
  //Task	Area	Type	Owner	Due Date	Status
  //PRINT: Owner, Status, Due Date, Task
  //SORT: Owner, Status, Due Date
  
  //create body
  var body = '';
  var bodyHtml = '';
  var maxCharPerColumn = [];

  //Initialize the Array for each column
  for(var columnIndex = 0; columnIndex < data[0].length; columnIndex++) {
    maxCharPerColumn[columnIndex] = 0;
  }
  
  //Find out the max string length for each row
  for(var rowIndex = 0; rowIndex < data.length; rowIndex++) {
    for(var columnIndex = 0; columnIndex < data[rowIndex].length; columnIndex++) {
      if(maxCharPerColumn[columnIndex] < data[rowIndex][columnIndex].toString().length) {
        maxCharPerColumn[columnIndex] = data[rowIndex][columnIndex].toString().length
      }
    }
  }
  
  for(var rowIndex = 0; rowIndex < data.length; rowIndex++) {
    var row = '';
    for(var columnIndex = 0; columnIndex < data[rowIndex].length; columnIndex++) {
      //Add data to each row
      var currentItem = data[rowIndex][columnIndex];
      if(typeof currentItem == 'object' && Object.prototype.toString.call(currentItem) === '[object Date]') {
        row = row + ' \t' + data[rowIndex][columnIndex].getMonth() + '/' + data[rowIndex][columnIndex].getDay() + '/' + data[rowIndex][columnIndex].getYear();
      } else {
        row = row + ' \t' + data[rowIndex][columnIndex];
      }
      //Add padding to that field
      //for(var spaceIndex = data[rowIndex][columnIndex].toString().length; spaceIndex < maxCharPerColumn[columnIndex]; spaceIndex++) {
      //  row = row + ' ';  
      //}
    }
    body = body + row + '\n';
  }
  
  bodyHtml = '<table>';
  for(var rowIndex = 0; rowIndex < data.length; rowIndex++) {
    var row = '';
    bodyHtml = bodyHtml +'<tr>';
    for(var columnIndex = 0; columnIndex < data[rowIndex].length; columnIndex++) {
      //Add data to each row
      var currentItem = data[rowIndex][columnIndex];
      if(typeof currentItem == 'object' && Object.prototype.toString.call(currentItem) === '[object Date]') {
        row = row + '<td>' + (currentItem.getMonth()+1) + '/' + currentItem.getDate() + '/' + currentItem.getYear() + '</td>';
        //row = row + '<td>' + currentItem.toLocaleDateString('en-US') + '</td>';
        
      } else {
        row = row + '<td>' + data[rowIndex][columnIndex] + '</td>';
      }
      //Add padding to that field
      //for(var spaceIndex = data[rowIndex][columnIndex].toString().length; spaceIndex < maxCharPerColumn[columnIndex]; spaceIndex++) {
      //  row = row + ' ';  
      //}
    }
    bodyHtml = bodyHtml + row + '\n';
    bodyHtml = bodyHtml +'</tr>';
  }
  bodyHtml = bodyHtml +'</table>';
  
  
  for(var i = 0; i < settings.emails.length ; i++) {    
    // Send yourself an email with a link to the document.
    GmailApp.sendEmail(settings.emails[i], settings.subject, body,{htmlBody:bodyHtml});
  }
};

