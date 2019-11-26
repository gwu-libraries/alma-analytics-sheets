/* Global variables */                 

// Alma Analytics API URL
var analyticsURL = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1/analytics/reports';

// Namespaces for the Analytics API XML
var rowNameSpace = XmlService.getNamespace('urn:schemas-microsoft-com:xml-analysis:rowset'),
  schemaNameSpace = XmlService.getNamespace('xsd', 'http://www.w3.org/2001/XMLSchema');

// Current (owning) spreadsheet --> Insert you spreadsheet ID here
var spreadsheet = SpreadsheetApp.openById('');
// Test version
//var spreadsheet = SpreadsheetApp.openById('');
// Google API class for throwing alerts to the user
//var ui = SpreadsheetApp.getUi();

// Keys used by the config objects -- should match the columns on the config tab of the calling (owning) spreadsheet.
var configKeys = ['apiKey', 'reportPath', 'spreadsheetTab'];

function encodePath(path) {
	// encodes an Alma Analytics path, replacing the ASCII characters for forward slash and whitespace with the appropriate code
    return path.replace(/\//g, '%2F').replace(/\s/g, '%20');
}

function errorHandler(e) { 
  Logger.log(e);
  throw e;
  //ui.alert(e);
}


function loadConfig() { 
/* Looks for a tab called config in the current (owning) spreadsheet, and loads the data into an array of JS objects.  
These objects will be used to fetch reports from Analytics and populate the data in the spreadsheet.*/
  try {
    var configSheet = spreadsheet.getSheetByName('config');
  }
  catch (e) { 
    throw 'Failed to find config tab in current spreadsheet.';
  }
  // data will be a 2-D array, where each element in the outer array is a row.
  // The first row should include the keys for the config obj.
  var data = configSheet.getDataRange().getValues(),
    columnKeys = data[0].filter(function(d) { 
        // make sure the columns correspond to the elements of configKeys
        return configKeys.indexOf(d) != -1;
    });
  if (columnKeys.length < configKeys.length) throw 'config tab is missing required columns.';
  // Populate an array of parameters for use in fetching and saving the reports. 
  try {
    data = data.slice(1);
    var reports = data.reduce(function(arr, row) { 
      // This function converts a 2-D array into an array of objects, with the object properties corresponding to the column names.
      var rowObj = {};
      columnKeys.forEach(function(key, i) { 
        // We only need one API key, so only the first row needs a value for this property.
        if ((row[i] == '') & (key == 'apiKey')) { 
          // Use the API key from the first row
          rowObj[key] = arr[0].apiKey;
        }
        else { 
          rowObj[key] = row[i];
        }
      });
      arr.push(rowObj);
      return arr;
    }, []);
   return reports;
  }
  catch (e) { 
    throw 'Error creating config obj. Please check config tab in the spreadsheet.';
  }
}


function main() {
  
  // Object to store the data for each sheet as we get it
  var tableObj = {};
  try {
    // Array of paramaters for fetching reports and writing to the current (owning) spreadsheet.
    var reports = loadConfig();
    // main loop ==> iterate over the reports to be extracted
    for (var i = 0; i<reports.length; i++) {
      try {
        var tableData = callAnalyticsAPI(reports[i]);
        tableObj[reports[i].spreadsheetTab] = tableData;
      }
      catch (e) { 
        if (e == 'Query failed!') {
          Logger.log('Query failed on ' + reports[i].spreadsheetTab);      
          continue;
        }
        else {
          throw e;
        }
      }
    }
  }
  catch(e) {
    errorHandler(e);
  }   
  // Now pass all the data to the spreadsheetApp at the same time
  sendToSpreadsheet(tableObj);
}

function callAnalyticsAPI(report) {
  /* Calls the Analtyics API to fetch a given report. Handles pagination when more than 1000 results are returned. Report should be a JS object with an API key and Analytics path as properties. */
  var headers = {Authorization: 'apikey ' + report.apiKey},
    params = {headers: headers,
                muteHttpExceptions: true},
    query = '?limit=1000&path=' + encodePath(report.reportPath),
    token = '', // initially, no token parameter needed
    data = [],
    columnMap = [],
    isFinished = false; // flag that will be set to true when the last page of results is reached
   
   while (!isFinished) { 
     var response = UrlFetchApp.fetch(analyticsURL + query + token, params);
     // Check for valid response
     if (response.getResponseCode() != 200) { 
     // throw exception  
       Logger.log(analyticsURL + query);
       Logger.log(response.getContentText());
       throw 'Query failed!'
     }
     try {
       var parsedResponse = parseXMLResponse(response.getContentText(), columnMap);
     }
     catch (e) { 
       Logger.log(e);
     }

     if (parsedResponse.isFinished == 'true') isFinished = true;
     
     if (parsedResponse.resumptionToken) token = '&token=' + parsedResponse.resumptionToken; // if we have a token, save it to use in subsequent requests
     
     data = data.concat(parsedResponse.data);
     columnMap = parsedResponse.columnMap; // Get the columnMap if we have it

    }
    var headerRow = columnMap.map(function(column) { // Extract the column headers as an array
           return column.columnHeading;
    }); 
    return [headerRow].concat(data); // Add the header row
  }

function parseXMLResponse(data, columnMap) { 
/* Iterate over the XML results, building up a 2-D array where each inner array corresponds to a  single row. columnMap should be non-null when we are retrieving results beyond the first page. Otherwise, it will be created
from the XML data.*/
  
  var document = XmlService.parse(data),
     result = document.getRootElement().getChild('QueryResult'),
     isFinished = result.getChild('IsFinished').getText(), // Is the report done, or are there more pages to return? 
     resumptionToken = result.getChild('ResumptionToken'),
     // rowset contains all the rows in the result (for this page of results)
     rows = result.getChild('ResultXml').getChild('rowset', rowNameSpace);
     if (columnMap.length == 0) { // If this is the first page of results, we need to get the column names
       Logger.log('Getting columns')
       columnMap = getColumnMap(rows).filter(function(column) { 
           return column.columnHeading != '0'; // Ignore columns with '0' as the heading -- these are spurious columns inserted by the Analytics API on some reports
         });
       // each row is its own element
       try {
         rows = rows.getChildren('Row', rowNameSpace);
       } catch (e) { 
         Logger.log('Problem fetching rows from the first page');
         throw e;
       }
     }
     else { 
       try {
         rows = result.getChild('ResultXml').getChild('rowset', rowNameSpace).getChildren('Row', rowNameSpace); // If it's not the first page of results, the rows won't have the namespace attribute
        } catch (e) { 
          Logger.log('Problem fetching rows from the second page');
          throw e;
        }
     }
     // in this reduce function, the first value is a 2-d array to be populated (for the DataRange of the  spreadsheet), the second is the particular row in the result set to be processed
     var table = rows.reduce(function(tableArray, rowElement) { 
                   // Maps the list of column names to their generic <Column> elements in the XML, parsing the data in each 
                   var row = columnMap.map(function(column) {
                     var cellValue = rowElement.getChildText(column.name, rowNameSpace);
                     if (!cellValue) cellValue = rowElement.getChildText(column.name); // see above -- namespaces are present only in the first page of results
                     return cellValue;
                   });                                 
                   tableArray.push(row);
                   return tableArray;
                 }, []);
 
     return {isFinished: isFinished,
             resumptionToken: resumptionToken ? result.getChild('ResumptionToken').getText() : null, // Token for fetching more pages -- present only on the first page
             data:table,
             columnMap: columnMap
           }; 
}



function getColumnMap(data) {
   // Because Analytics returns the columns in what appears to be arbitrary order, it is necessary to extract a mapping from the XML itself
   // This data is in enclosed in the following path //rowset/xsd:schema/xsd:complexTypes/xsd:sequence/
   // Using namespaces with XmlService is not straightforward, so the following is an inelegant method that avoids that problem
  
   // assumes that the <xsd:schema> element will always be the first child of <rowset>
   try {
     var sequence = data.getChildren()[0].getChildren()[0].getChildren()[0];
   } catch (e) {
     Logger.log('Problem fetching <sequence> element for column names');
     throw e;
   }
   // Returns a list of objects, each of which maps an element name to a columnHeading. 
   try {
     var result= sequence.getChildren().map(function (element) { 
       var attributes = element.getAttributes();
       return attributes.reduce(function (attrObj, attr) { 
         var attrName = attr.getName();
         // Looking for two particular headings
         if ((attrName == 'name') || (attrName == 'columnHeading')) {
           attrObj[attrName] = attr.getValue();
         }
         return attrObj;
       }, {});
     });
     return result;
    }
     catch (e) { 
       Logger.log('Error fetching children of <sequence> element');
     }
   
   

}


function sendToSpreadsheet(tableObj) {
 
  try {
    //loop over the reports for which we have data
    Object.keys(tableObj).forEach(function(tableKey) {
      var sheetName = tableKey; // From the main fn, each data object should have as its key a spreadsheet tab as defined in the config tab
       // get the sheet corresponding to this report
      var sheet = spreadsheet.getSheetByName(sheetName);
      if (sheetName == null) { 
        // if necessary, create the sheet
        sheet = spreadsheet.insertSheet(sheetName);
      }
      var data = tableObj[tableKey];
      sheet.clear();
      // Test for data that exceeds the current number of available rows
      var sheetLength = sheet.getMaxRows();
      if (sheetLength < data.length) { 
        var diff = data.length - sheetLength;
        sheet.insertRowsAfter(sheetLength, diff);
      }
      // Get the range of the required dimensions
      var range = sheet.getRange(1, 1, data.length, data[0].length);
      try { 
        range.setValues(data);
      }
      catch (e) {
        Logger.log(e);
      }
    });
  }
  catch (e) {
    //ui.alert("Loading spreadsheet failed!");
    Logger.log(e);
  }
}