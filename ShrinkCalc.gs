// ShrinkCalc

// Processes and sorts sales and shrinkage data to identify problem areas

// by Grant Trebbin
function generateSummary() {

  // Sheet variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dumpsSheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Dumps"));
  var markdownsSheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Markdowns"));
  var salesSheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Sales"));
  var descriptionsSheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Descriptions"));
  var summarySheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Summary"));
  var analysisSheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Analysis"));
  
  // Clear sheets
  summarySheet.clearContents();
  
  // Add title row to report
  var titleArray = new Array();
  titleArray.push(["Item Number", "Description", "Sales", "Markdowns", "Dumps", "Loss", "Potential Sales", "Aggregate Loss (%)"])
  summarySheet.getRange(1, 1, 1, 8).setValues(titleArray);
  SpreadsheetApp.flush();
 
  
  // Get dumps data
  var dumpsData = dumpsSheet.getDataRange().getValues();

  // Get sales data
  var salesData = salesSheet.getDataRange().getValues();

  // Get markdowns data
  var markdownsData = markdownsSheet.getDataRange().getValues();
  
  // Get descriptions data
  var descriptionsData = descriptionsSheet.getDataRange().getValues();
  
  // Get analysis data
  var analysisData = analysisSheet.getDataRange().getValues();
  
  var itemNumbers = new Array();
  var temp = "";
  
  // get dumps item numbers
  for (var i = 1, ii= dumpsData.length; i < ii; i++) {
    temp = dumpsData[i][0];
    if (temp != ""){
      if (itemNumbers.indexOf(temp) ==  -1){
        itemNumbers.push(dumpsData[i][0]);
      }
    }
  }
  
  // get sales item numbers
  for (var i = 1, ii= salesData.length; i < ii; i++) {
    temp = salesData[i][0];
    if (temp != ""){
      if (itemNumbers.indexOf(temp) ==  -1){
        itemNumbers.push(salesData[i][0]);
      }
    }
  }
  
  // get markdowns item numbers
  for (var i = 1, ii= markdownsData.length; i < ii; i++) {
    temp = markdownsData[i][0];
    if (temp != ""){
      if (itemNumbers.indexOf(temp) ==  -1){
        itemNumbers.push(markdownsData[i][0]);
      }
    }
  }
  
  // Arrays to hold results
  var summaryArray = new Array();
  
  // Sort item numbers
  itemNumbers.sort(function(a,b){return a-b});

  // Assemble the data for each entry
  for (var i = 0; i<itemNumbers.length; i++){
    var refNo = itemNumbers[i];
    var Description = "";
    var Sales = 0;
    var Dumps = 0;
    var Markdowns = 0;
    
    // retrieve the sales
    for (var j = 1; j<salesData.length; j++){
      if(refNo == salesData[j][0]){
        Sales = salesData[j][1];
        break;
      }
    }

    // retrieve the description
    for (var j = 1; j<descriptionsData.length; j++){
      if(refNo == descriptionsData[j][0]){
        Description = descriptionsData[j][1];
        break;
      }
    }

    // retrieve the markdowns
    for (var j = 1; j<markdownsData.length; j++){
      if(refNo == markdownsData[j][0]){
        Markdowns = markdownsData[j][1];
        break;
      }
    }
    
    Dumps = 0;
    // retrieve the dumps
    for (var j = 1; j<dumpsData.length; j++){
      if(refNo == dumpsData[j][0]){
        Dumps = Dumps + dumpsData[j][2];
      }
    }

    
    // Calculate metrics. Limit decimal places for clarity
    var Loss = Dumps + Markdowns;
    var PotentialSales = Dumps + Sales + Markdowns;
    var AggregateLoss = (100 * (1-Sales/PotentialSales)).toFixed(2);
    
    // Add data row to report array
    summaryArray.push([refNo, Description, Sales, Markdowns, Dumps, Loss, PotentialSales, AggregateLoss]);
  }
  
  // Populate the summary sheet
  summarySheet.getRange(2,1,summaryArray.length,8).setValues(summaryArray);

  // Clear data from analysis sheet
  analysisSheet.getRange(4,2,analysisData.length,8).clear();
  // Clear borders from analysis sheet
  analysisSheet.getRange(1,1,analysisData.length,8).setBorder(false, false, false, false, false, false);
  finalAnalysisRow = analysisSheet.getLastRow();
  

  var AnalysisSales = 0
  var AnalysisMarkdowns = 0
  var AnalysisDumps = 0

  // Populate custom sheet
  for (var i = 3; i <finalAnalysisRow; i++){
    if (analysisData[i][0] != ""){
      var itemToFind = analysisData[i][0]
      for (var j = 0; j<summaryArray.length; j++){
        if (itemToFind == summaryArray[j][0]){
          AnalysisSales = AnalysisSales + summaryArray[j][2];
          AnalysisMarkdowns = AnalysisMarkdowns + summaryArray[j][3];
          AnalysisDumps = AnalysisDumps + summaryArray[j][4];
          analysisSheet.getRange(i+1,2,1,7).setValues([[summaryArray[j][1],summaryArray[j][2],
                                                        summaryArray[j][3],summaryArray[j][4],
                                                        summaryArray[j][5],summaryArray[j][6],
                                                        summaryArray[j][7]]]);
        }
      }
    }
  } 

  // Add totals to report
  var AnalysisLoss = AnalysisDumps+AnalysisMarkdowns;
  var AnalysisPotentialSales = AnalysisLoss + AnalysisSales
  var AnalysisAggregateLoss = (100 * (1-AnalysisSales/AnalysisPotentialSales)).toFixed(2);
  analysisSheet.getRange(finalAnalysisRow+1,2,1,1).setValue("Total");
  analysisSheet.getRange(finalAnalysisRow+1,3,1,1).setValue(AnalysisSales);
  analysisSheet.getRange(finalAnalysisRow+1,4,1,1).setValue(AnalysisMarkdowns);
  analysisSheet.getRange(finalAnalysisRow+1,5,1,1).setValue(AnalysisDumps);
  analysisSheet.getRange(finalAnalysisRow+1,6,1,1).setValue(AnalysisLoss);
  analysisSheet.getRange(finalAnalysisRow+1,7,1,1).setValue(AnalysisPotentialSales);
  analysisSheet.getRange(finalAnalysisRow+1,8,1,1).setValue(AnalysisAggregateLoss);
  analysisSheet.getRange(finalAnalysisRow+1,2,1,7).setFontWeight("bold");

  // redraw borders
  analysisSheet.getRange(2,1,finalAnalysisRow,8).setBorder(null, null, null, null, null, true);

}


// Sort the results data by Item Number
function sortItem(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Summary"));

  var range = summarySheet.getRange(2, 1, summarySheet.getLastRow()-1, summarySheet.getLastColumn());
  range.sort({column: 1, ascending: true});
}

// Sort the results data by Loss
function sortLoss(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Summary"));

  var range = summarySheet.getRange(2, 1, summarySheet.getLastRow()-1, summarySheet.getLastColumn());
  range.sort({column: 6, ascending: false});
}

// Sort the results data by Aggregate Loss
function sortAggregateLoss(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("Summary"));

  var range = summarySheet.getRange(2, 1, summarySheet.getLastRow()-1, summarySheet.getLastColumn());
  range.sort({column: 8, ascending: false});
}

// When the spreadsheet opens add a custom menu
function onOpen() {
  var activeSS = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [{name: "Generate Summary", functionName: "generateSummary"},
                   {name: "Sort by Item Number", functionName: "sortItem"},
                   {name: "Sort by Loss", functionName: "sortLoss"},
                   {name: "Sort by AggregateLoss", functionName: "sortAggregateLoss"}]
  activeSS.addMenu("WOW AutoStockR", menuItems);
}