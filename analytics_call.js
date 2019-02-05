                 
// For the Analytics API XML
var rowNameSpace = XmlService.getNamespace('urn:schemas-microsoft-com:xml-analysis:rowset'),
  schemaNameSpace = XmlService.getNamespace('xsd', 'http://www.w3.org/2001/XMLSchema');
// For throwing alerts to the user --> Doesn't work on a timed trigger
var ui = SpreadsheetApp.getUi();

// encodes an Alma Analytics path, replacing the ASCII characters for forward slash and whitespace with the appropriate code
function encodePath(path) {
  return path.replace(/\//g, '%2F').replace(/\s/g, '%20');
}

function main() {
  
  // Object to store the data for each sheet as we get it
  var tableObj = {};
  // main loop ==> iterate over the reports to be extracted
  for (var i = 0; i<configObj.reports.length; i++) {
    try {
      var tableData = callAnalyticsAPI(configObj.reports[i]);
      tableObj[configObj.reports[i].sheet] = tableData;
    }
    catch(e) {
      if (e == 'Query failed!') {
        ui.alert('Query failed on ' + configObj.reports[i].sheet);
        Logger.log('Query failed on ' + configObj.reports[i].sheet);
        continue;
      }
      else {
        ui.alert(e);
        Logger.log(e);
        continue;
      }
    }    
  }
  // Now pass all the data to the spreadsheetApp at the same time
  sendToSpreadsheet(tableObj)
}

function callAnalyticsAPI(report) {
  var headers = {Authorization: 'apikey ' + configObj.apikey},
    params = {headers: headers,
                muteHttpExceptions: true},
    query = '?limit=1000&path=' + encodePath(report.path); 
   
    var response = UrlFetchApp.fetch(configObj.url + query, params);
   
   // Check for valid response
     if (response.getResponseCode() != 200) { 
     // throw exception and 
       Logger.log(response.getContentText());
       throw 'Query failed!'
     }
     return parseXMLResponse(response.getContentText(), report.columns);
  }

function parseXMLResponse(data, columns) { 
// Iterate over the XML results, building up a 2-D array where each inner array corresponds to a  single row   
  
  var document = XmlService.parse(data),
     root = document.getRootElement(),
     // rowset contains all the rows in the result (for this page of results)
     rows = root.getChild('QueryResult').getChild('ResultXml').getChild('rowset', rowNameSpace);
     var columnMap = getColumnMap(rows);
     // each row is its own element
     rows = rows.getChildren('Row', rowNameSpace);
     // in this reduce function, the first value is a 2-d array to be populated (for the DataRange of the  spreadsheet), the second is the particular row in the result set to be processed
     var table = rows.reduce(function(tableArray, rowElement) { 
                   // Maps the list of column names to their generic Column elements in the XML, parsing the data in each 
                   var row = columnMap.filter(function(column) {
                                     // Ignore Column0, which contains no data
                                     return column.columnHeading != '0'
                                     }).map(function(column, i) {
                                       columns[i] = column.columnHeading;
                                       return rowElement.getChildText(column.name, rowNameSpace);
                                   });                                 
                   tableArray.push(row);
                   return tableArray
                 }, []);
     return table;
}



function getColumnMap(data) {
   // Because Analytics returns the columns in what appears to be arbitrary order, it is necessary to extract a mapping from the XML itself
   // This data is in enclosed in the following path //rowset/xsd:schema/xsd:complexTypes/xsd:sequence/
   // Using namespaces with XmlService is not straightforward, so the following is an inelegant method that avoids that problem
  
   // assumes that the <xsd:schema> element will always be the first child of <rowset>
   var sequence = data.getChildren()[0].getChildren()[0].getChildren()[0];
   // Returns a list of objects, each of which maps an element name to a columnHeading. The columnHeadings should match the column names in config.js
   return sequence.getChildren().map(function (element) { 
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
   
   

}


function sendToSpreadsheet(tableObj) {
 
  try {
    var ss = SpreadsheetApp.openById(configObj.spreadsheet);
    //loop over the reports for which we have data
    for (var i = 0; i < configObj.reports.length; i++) {
      var sheetName = configObj.reports[i].sheet;
      // skip if there's no data
      if (!tableObj.hasOwnProperty(sheetName)) continue;
      // get the sheet corresponding to this report
      var sheet = ss.getSheetByName(sheetName);
      // prepare the data by appending the columns
      var data = tableObj[sheetName],
          header = configObj.reports[i].columns;
      data = [header].concat(data)
      // Clear the current data
      sheet.clear();
      // Get the range of the required dimensions
      var range = sheet.getRange(1, 1, data.length, header.length);
      range.setValues(data);
    }
  }
  catch (e) {
    ui.alert("Loading spreadsheet failed!");
    Logger.log(e);
  }
}