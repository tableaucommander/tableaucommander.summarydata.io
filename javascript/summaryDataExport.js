
'use strict';

$(document).ready(function () {

  $("#pullWorksheetDataButton").click(function() {
  

      tableau.extensions.initializeAsync().then(function() {
          
          //function to Export individual worksheet summary data to single Excel Workbook
              exportToExcel(); 

      }, function(err) {
  
        // something went wrong in initialization
        $("#resultBox").html("Error while Initializing: " + err.toString());
      });
    });
})
  



function exportToExcel() {
  var dashboard = tableau.extensions.dashboardContent.dashboard;
  var excelWorksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
  var workbook = XLSX.utils.book_new();
  var sheetList = [];
  var totalSheets = excelWorksheets.length;
  var sheetCount = 0;  

  console.log(excelWorksheets.length);


  for (var b = 0; b < excelWorksheets.length; b++ ) {
      sheetList.push(excelWorksheets[b].name);
      
  }
  console.log(sheetList);

  for (var c = 0; c < excelWorksheets.length; c++ ) {

      var currentWorksheet = excelWorksheets[c];

      if (sheetList.indexOf(currentWorksheet.name) > -1) {




  //Get Worksheet Summary Data
  //-------------------------------------------------------------------------------------------------------------------

  currentWorksheet.getSummaryDataAsync().then(function (excelWorksheetData) {

      var excelColumns = excelWorksheetData.columns;
      var excelData = excelWorksheetData.data;

      


  //Pull Headers, utilizing two seperate arrays
  //-------------------------------------------------------------------------------------------------------------------
      var headerRows =[];

      for (var y = 0; y < 1; y++) {
          
          headerRows.push([]);
          
      }


      for (var k = 0; k < 1; k++) {
          for (var j = 0; j < excelColumns.length; j++) {

              headerRows[k].push(excelColumns[j].fieldName);                     
          
          }
      }


      

  //Build Main Data Table Array
  //-------------------------------------------------------------------------------------------------------------------

      var rows = [];

      for (var x = 0; x < excelData.length; x++) {
          
          rows.push([]);

      }

      
      for (var i = 0; i < excelData.length; i++) {

          for (var h = rows[i].length; h < excelColumns.length; h++) {

              rows[i].push(excelData[i][h].formattedValue);

          }   
      }



  //Export aggregated data to Excel Workbook
  //-------------------------------------------------------------------------------------------------------------------   
  

      headerRows.splice(excelColumns.length,excelColumns.length);
      var dbname = dashboard.name;
      var wsname = sheetList[sheetCount];
      sheetCount = sheetCount + 1; 
  
      var worksheet = XLSX.utils.json_to_sheet(rows, {origin: 'A2', skipHeader: true });
  
              XLSX.utils.sheet_add_aoa(worksheet, headerRows, { origin: 'A1' });
  
              XLSX.utils.book_append_sheet(workbook, worksheet, wsname);
  
              console.log(sheetCount);
              console.log(totalSheets);
              
              if (sheetCount == totalSheets) {
  
                      XLSX.writeFile(workbook, dbname + ".xlsx"); 
  
              }

          

  })

}
}
}