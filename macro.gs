/** @OnlyCurrentDoc */

function RecordDay() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1').setValue("Undo 1");
  spreadsheet.getRange('D1').setValue("Undo 2");
  var date = spreadsheet.getRange('A1').getValue();
  spreadsheet.getRange('B1').setValue('24:00:00');
  spreadsheet.getRange('B1').setNumberFormat('[h]:mm:ss')
  var increment = 4;
  while (spreadsheet.getRange('A' + increment).getValue() != ""){
    // https://cloud.google.com/dataprep/docs/html/Comparison-Operators_57344676
    var isRecordableEvent1 = spreadsheet.getRange('D' + increment).getValue() > 0;
    var isRecordableEvent2 = spreadsheet.getRange('D' + increment).getValue() == "0";
    if (isRecordableEvent1 || isRecordableEvent2) {
      spreadsheet.getRange('A' + increment).setNumberFormat('[h]:mm:ss');
      spreadsheet.getRange('B' + increment).setNumberFormat('[h]:mm:ss');
      var start = spreadsheet.getRange('A' + increment).getValue();
      var end = spreadsheet.getRange('B' + increment).getValue();
      var event = spreadsheet.getRange('C' + increment).getValue();
      var amount = spreadsheet.getRange('D' + increment).getValue();

      if (end == ""){
        spreadsheet.getRange('A' + increment + ':E' + increment).setBackground('#ffff00');
      }
      else{

        var eventSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(event);

        if (eventSpreadsheet == null){
          spreadsheet.getRange('A' + increment + ':E' + increment).setBackground('#ffff00');
        }
        else{

          var increment2 = 3;
          while((eventSpreadsheet.getRange('A' + increment2).getValue() != "") && (eventSpreadsheet.getRange('D' + (increment2 + 1)).getValue() != "Tomorrow")){
            increment2++;
          }

          eventSpreadsheet.getRange('D' + (increment2 + 1)).setValue("Tomorrow");
          var sheetname = "'" + spreadsheet.getSheetName() + "'";
          eventSpreadsheet.getRange('F:F').setNumberFormat('[h]:mm:ss');
          eventSpreadsheet.getRange('B:B').setNumberFormat('[h]:mm:ss');
          eventSpreadsheet.getRange('D:D').setNumberFormat('[h]:mm:ss');
          eventSpreadsheet.getRange('H:H').setNumberFormat('[h]:mm:ss');

            // https://infoinspired.com/google-docs/spreadsheet/comparison-operators-in-google-sheets/#:~:text=%2C%22NO%22)-,Google%20Sheets%20Comparison%20Operator%20%E2%80%9C%3E%3D%E2%80%9D%20and%20Function%20GTE%20(,equal%20to%20the%20second%20value.
            if(start > end){
              spreadsheet.getRange('A1').copyTo(eventSpreadsheet.getRange('A' + increment2), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
              var previousFormula = eventSpreadsheet.getRange('B' + increment2).getFormula();
              eventSpreadsheet.getRange('B' + increment2).setFormula(previousFormula + " + " + sheetname + "!B1 - " + sheetname + "!A" + increment + " + " + sheetname + "!B" + increment);
              var previousAmount = eventSpreadsheet.getRange('C' + increment2).getFormula();
              eventSpreadsheet.getRange('C' + increment2).setFormula(previousAmount + " + " + sheetname + "!D" + increment);
            }
            else {
              spreadsheet.getRange('A1').copyTo(eventSpreadsheet.getRange('A' + increment2), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
              var previousFormula = eventSpreadsheet.getRange('B' + increment2).getFormula();
              eventSpreadsheet.getRange('B' + increment2).setFormula(previousFormula + " + " + sheetname + "!B" + increment + " - " + sheetname + "!A" + increment);
              var previousAmount = eventSpreadsheet.getRange('C' + increment2).getFormula();
              eventSpreadsheet.getRange('C' + increment2).setFormula(previousAmount + " + " + sheetname + "!D" + increment);
            }

          if (eventSpreadsheet.getRange('C' + increment2).getValue() != "0"){
            eventSpreadsheet.getRange('D' + increment2).setFormula('B' + increment2 + " / " + "C" + increment2);
          }
          eventSpreadsheet.getRange('F' + increment2).setFormula('F' + (increment2 - 1) + " + " + "B" + increment2);
          eventSpreadsheet.getRange('G' + increment2).setFormula('G' + (increment2 - 1) + " + " + "C" + increment2);
          if (eventSpreadsheet.getRange('G' + increment2).getValue() != "0"){
            eventSpreadsheet.getRange('H' + increment2).setFormula('F' + increment2 + " / " + "G" + increment2);
          }

          var increment3 = 4;
          while ((spreadsheet.getRange('G' + increment3).getValue() != "") && (spreadsheet.getRange('G' + increment3).getValue() != event)){
            increment3++;
          }

          spreadsheet.getRange('G' + increment3).setValue(event);


          var eventSheetName = "'" + eventSpreadsheet.getSheetName() + "'";

          spreadsheet.getRange('H' + increment3).setFormula(eventSheetName + '!B' + increment2);

          spreadsheet.getRange('I' + increment3).setFormula(eventSheetName + '!C' + increment2);

          spreadsheet.getRange('J' + increment3).setFormula(eventSheetName + '!D' + increment2);

          spreadsheet.getRange('K' + increment3).setFormula(eventSheetName + '!F' + increment2);

          spreadsheet.getRange('L' + increment3).setFormula(eventSheetName + '!G' + increment2);

          spreadsheet.getRange('M' + increment3).setFormula(eventSheetName + '!H' + increment2);

          spreadsheet.getRange('N' + increment3).setFormula(eventSheetName + '!H' + (increment2 - 1));
        }

      }

    }
    increment++;
  }

  var increment4 = 4;
  while (spreadsheet.getRange('G' + increment4).getValue() != ""){
    var clearTomorrowSpreadsheetName = spreadsheet.getRange('G' + increment4).getValue();
    var clearTomorrowSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(clearTomorrowSpreadsheetName);
    var increment5 = 3;
    while(clearTomorrowSpreadsheet.getRange('D' + increment5).getValue() != "Tomorrow"){
      increment5++;
    }
    clearTomorrowSpreadsheet.getRange('D' + increment5).setValue("");
    increment4++;
  }





      var increment5 = 4;
      while (spreadsheet.getRange('G' + increment5).getValue() != ""){
        increment5++;
      }
      increment5--;

      spreadsheet.getRange('G4:N' + increment5).sort({column: 8, ascending: false});

      spreadsheet.getRange('G3:H' + increment5).activate();
      var chart = spreadsheet.newChart()
      .asColumnChart()
      .addRange(spreadsheet.getRange('G3:H' + increment5))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'Time vs. Event')
      .setXAxisTitle('Event')
      .setYAxisTitle('Time')
      .setOption('height', 481)
      .setOption('width', 934)
      .setPosition(31, 6, 99, 0)
      .build();
      spreadsheet.insertChart(chart);






};

function UpdateBarGraphofAllActivities() {
  // https://webapps.stackexchange.com/questions/14112/in-google-spreadshets-how-can-you-loop-through-all-available-sheets-not-knowing
  // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheets
  var graphSpreadsheet = SpreadsheetApp.getActive();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var nextLocation = 2;
  for (var i = 0; i < sheets.length; i++){
    var currentSheet = sheets[i];
    if (currentSheet.getRange('A2').getValue() == "Date"){
      var j = 3;
      while(currentSheet.getRange('A' + j).getValue() != ""){
        j++;
      }
      var currentSheetName = "'" + currentSheet.getSheetName() + "'";
      graphSpreadsheet.getRange('A' + nextLocation).setFormula(currentSheetName + '!A1');
      graphSpreadsheet.getRange('C' + nextLocation).setFormula(currentSheetName + '!F' + (j - 1));
      nextLocation++;
    }
  }
  graphSpreadsheet.getRange('C:C').setNumberFormat('[h]:mm:ss');
  graphSpreadsheet.getRange('A2:C' + nextLocation).sort({column: 3, ascending: false});
  for (var k = 2; k < nextLocation; k++){
     // =(C2 / E5) * 100
    graphSpreadsheet.getRange('B' + k).setFormula("(C" + k + " / " + "E5) * 100");
  }
};

function UpdateAllCharts() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++){
    var currentSheet = sheets[i];

    if (currentSheet.getRange('A2').getValue() == "Date"){
      var DailyTimeChart = currentSheet.newChart()
      .asColumnChart()
      .addRange(currentSheet.getRange('A2:A'))
      .addRange(currentSheet.getRange('B2:B'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(-1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('title', 'Daily Time vs. Date')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setPosition(6, 3, 7, 18)
      .build();
      currentSheet.insertChart(DailyTimeChart);

      var DailyAmountChart = currentSheet.newChart()
      .asColumnChart()
      .addRange(currentSheet.getRange('A2:A'))
      .addRange(currentSheet.getRange('C2:C'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(-1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('title', 'Daily Amount vs. Date')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setPosition(6, 3, 7, 18)
      .build();
      currentSheet.insertChart(DailyAmountChart);

      var DailyAverageTimeChart = currentSheet.newChart()
      .asColumnChart()
      .addRange(currentSheet.getRange('A2:A'))
      .addRange(currentSheet.getRange('D2:D'))
      .addRange(currentSheet.getRange('H2:H'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(-1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('title', 'Daily Average Time vs. Date')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setPosition(6, 3, 7, 18)
      .build();
      currentSheet.insertChart(DailyAverageTimeChart);
    }
  }
};

function UpdatePieGraphsofSelectActivities() {
  var graphSpreadsheet = SpreadsheetApp.getActive();
  var sheet = graphSpreadsheet.getActiveSheet();
  var startOfNextPie = 1;

  while (graphSpreadsheet.getRange('A' + startOfNextPie).getValue() != ""){
    var i = startOfNextPie;
    while(graphSpreadsheet.getRange('A' + i).getValue() != ""){
      if (graphSpreadsheet.getRange('A' + i).getValue() != "Event"){
        var event = graphSpreadsheet.getRange('A' + i).getValue();
        var eventSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(event);

        var j = 3;
        while(eventSpreadsheet.getRange('A' + j).getValue() != ""){
          j++;
        }

        var eventSheetName = "'" + event + "'";
        graphSpreadsheet.getRange('C' + i).setFormula(eventSheetName + '!F' + (j - 1));
      }
      i++
    }
    i = i - 1;

    graphSpreadsheet.getRange('D' + (startOfNextPie + 1)).setFormula("SUM(C" + (startOfNextPie + 1) + ":C" + i + ")");
    var i = startOfNextPie;
    while(graphSpreadsheet.getRange('A' + i).getValue() != ""){
      if (graphSpreadsheet.getRange('A' + i).getValue() != "Event"){
        // =(C2 / D2) * 100
        graphSpreadsheet.getRange('B' + i).setFormula("(C" + i + " / D" + (startOfNextPie + 1) + ") * 100");
      }
      i++;
    }
    i = i - 1;

    var chart = sheet.newChart()
    .asPieChart()
    .addRange(graphSpreadsheet.getRange('A' + startOfNextPie + ':B' + i))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('bubble.stroke', '#000000')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('isStacked', 'false')
    .setOption('title', 'Percentage vs. Event')
    .setOption('annotations.domain.textStyle.color', '#808080')
    .setOption('textStyle.color', '#000000')
    .setOption('legend.textStyle.color', '#191919')
    .setOption('titleTextStyle.color', '#757575')
    .setOption('annotations.total.textStyle.color', '#808080')
    .setPosition(6, 3, 30, 18)
    .build();
    sheet.insertChart(chart);

    startOfNextPie = i + 2;
  }
};

function CreateGraphsforEvent() {
  var spreadsheet = SpreadsheetApp.getActive();
  var currentSheet = spreadsheet.getActiveSheet();

  var DailyTimeChart = currentSheet.newChart()
  .asColumnChart()
  .addRange(currentSheet.getRange('A2:A'))
  .addRange(currentSheet.getRange('B2:B'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('title', 'Daily Time vs. Date')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#191919')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setPosition(6, 3, 7, 18)
  .build();
  currentSheet.insertChart(DailyTimeChart);

  var DailyAmountChart = currentSheet.newChart()
  .asColumnChart()
  .addRange(currentSheet.getRange('A2:A'))
  .addRange(currentSheet.getRange('C2:C'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('title', 'Daily Amount vs. Date')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#191919')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setPosition(6, 3, 7, 18)
  .build();
  currentSheet.insertChart(DailyAmountChart);

  var DailyAverageTimeChart = currentSheet.newChart()
  .asLineChart()
  .addRange(currentSheet.getRange('A2:A'))
  .addRange(currentSheet.getRange('D2:D'))
  .addRange(currentSheet.getRange('H2:H'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('curveType', 'none')
  .setOption('domainAxis.direction', 1)
  .setOption('title', 'Daily Average Time vs. Date')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('legend.textStyle.color', '#191919')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('series.0.labelInLegend', 'Average Time for 1')
  .setOption('series.1.labelInLegend', 'Cumulative Average Time for 1')
  .setPosition(38, 10, 14, 8)
  .build();
  currentSheet.insertChart(DailyAverageTimeChart);
};

function ContinueRecordingDay() {
  //https://developers.google.com/apps-script/reference/base/prompt-response
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('What row should record day be continued from?');
  var rowToContinueFrom = response.getResponseText();



  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1').setValue("Undo 1");
  spreadsheet.getRange('D1').setValue("Undo 2");
  var date = spreadsheet.getRange('A1').getValue();
  spreadsheet.getRange('B1').setValue('24:00:00');
  spreadsheet.getRange('B1').setNumberFormat('[h]:mm:ss')
  var increment = rowToContinueFrom;
  while (spreadsheet.getRange('A' + increment).getValue() != ""){
    // https://cloud.google.com/dataprep/docs/html/Comparison-Operators_57344676
    var isRecordableEvent1 = spreadsheet.getRange('D' + increment).getValue() > 0;
    var isRecordableEvent2 = spreadsheet.getRange('D' + increment).getValue() == "0";
    if (isRecordableEvent1 || isRecordableEvent2) {
      spreadsheet.getRange('A' + increment).setNumberFormat('[h]:mm:ss');
      spreadsheet.getRange('B' + increment).setNumberFormat('[h]:mm:ss');
      var start = spreadsheet.getRange('A' + increment).getValue();
      var end = spreadsheet.getRange('B' + increment).getValue();
      var event = spreadsheet.getRange('C' + increment).getValue();
      var amount = spreadsheet.getRange('D' + increment).getValue();

      if (end == ""){
        spreadsheet.getRange('A' + increment + ':E' + increment).setBackground('#ffff00');
      }
      else{

        var eventSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(event);

        if (eventSpreadsheet == null){
          spreadsheet.getRange('A' + increment + ':E' + increment).setBackground('#ffff00');
        }
        else{

          var increment2 = 3;
          while((eventSpreadsheet.getRange('A' + increment2).getValue() != "") && (eventSpreadsheet.getRange('D' + (increment2 + 1)).getValue() != "Tomorrow")){
            increment2++;
          }

          eventSpreadsheet.getRange('D' + (increment2 + 1)).setValue("Tomorrow");
          var sheetname = "'" + spreadsheet.getSheetName() + "'";
          eventSpreadsheet.getRange('F:F').setNumberFormat('[h]:mm:ss');
          eventSpreadsheet.getRange('B:B').setNumberFormat('[h]:mm:ss');
          eventSpreadsheet.getRange('D:D').setNumberFormat('[h]:mm:ss');
          eventSpreadsheet.getRange('H:H').setNumberFormat('[h]:mm:ss');

            if(start > end){
              spreadsheet.getRange('A1').copyTo(eventSpreadsheet.getRange('A' + increment2), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
              var previousFormula = eventSpreadsheet.getRange('B' + increment2).getFormula();
              eventSpreadsheet.getRange('B' + increment2).setFormula(previousFormula + " + " + sheetname + "!B1 - " + sheetname + "!A" + increment + " + " + sheetname + "!B" + increment);
              var previousAmount = eventSpreadsheet.getRange('C' + increment2).getFormula();
              eventSpreadsheet.getRange('C' + increment2).setFormula(previousAmount + " + " + sheetname + "!D" + increment);
            }
            else {
              spreadsheet.getRange('A1').copyTo(eventSpreadsheet.getRange('A' + increment2), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
              var previousFormula = eventSpreadsheet.getRange('B' + increment2).getFormula();
              eventSpreadsheet.getRange('B' + increment2).setFormula(previousFormula + " + " + sheetname + "!B" + increment + " - " + sheetname + "!A" + increment);
              var previousAmount = eventSpreadsheet.getRange('C' + increment2).getFormula();
              eventSpreadsheet.getRange('C' + increment2).setFormula(previousAmount + " + " + sheetname + "!D" + increment);
            }

          if (eventSpreadsheet.getRange('C' + increment2).getValue() != "0"){
            eventSpreadsheet.getRange('D' + increment2).setFormula('B' + increment2 + " / " + "C" + increment2);
          }
          eventSpreadsheet.getRange('F' + increment2).setFormula('F' + (increment2 - 1) + " + " + "B" + increment2);
          eventSpreadsheet.getRange('G' + increment2).setFormula('G' + (increment2 - 1) + " + " + "C" + increment2);
          if (eventSpreadsheet.getRange('G' + increment2).getValue() != "0"){
            eventSpreadsheet.getRange('H' + increment2).setFormula('F' + increment2 + " / " + "G" + increment2);
          }

          var increment3 = 4;
          while ((spreadsheet.getRange('G' + increment3).getValue() != "") && (spreadsheet.getRange('G' + increment3).getValue() != event)){
            increment3++;
          }

          spreadsheet.getRange('G' + increment3).setValue(event);

          var eventSheetName = "'" + eventSpreadsheet.getSheetName() + "'";

          spreadsheet.getRange('H' + increment3).setFormula(eventSheetName + '!B' + increment2);

          spreadsheet.getRange('I' + increment3).setFormula(eventSheetName + '!C' + increment2);

          spreadsheet.getRange('J' + increment3).setFormula(eventSheetName + '!D' + increment2);

          spreadsheet.getRange('K' + increment3).setFormula(eventSheetName + '!F' + increment2);

          spreadsheet.getRange('L' + increment3).setFormula(eventSheetName + '!G' + increment2);

          spreadsheet.getRange('M' + increment3).setFormula(eventSheetName + '!H' + increment2);

          spreadsheet.getRange('N' + increment3).setFormula(eventSheetName + '!H' + (increment2 - 1));
        }

      }

    }
    increment++;
  }

  var increment4 = 4;
  while (spreadsheet.getRange('G' + increment4).getValue() != ""){
    var clearTomorrowSpreadsheetName = spreadsheet.getRange('G' + increment4).getValue();
    var clearTomorrowSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(clearTomorrowSpreadsheetName);
    var increment5 = 3;
    while(clearTomorrowSpreadsheet.getRange('D' + increment5).getValue() != "Tomorrow"){
      increment5++;
    }
    clearTomorrowSpreadsheet.getRange('D' + increment5).setValue("");
    increment4++;
  }




  var increment5 = 4;
  while (spreadsheet.getRange('G' + increment5).getValue() != ""){
    increment5++;
  }
  increment5--;

  spreadsheet.getRange('G4:N' + increment5).sort({column: 8, ascending: false});

  spreadsheet.getRange('G3:H' + increment5).activate();
  var chart = spreadsheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('G3:H' + increment5))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Time vs. Event')
  .setXAxisTitle('Event')
  .setYAxisTitle('Time')
  .setOption('height', 481)
  .setOption('width', 934)
  .setPosition(31, 6, 99, 0)
  .build();
  spreadsheet.insertChart(chart);




};

function ColumnCharttoLineChartforAllTasks() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++){
    var currentSheet = sheets[i];

    if (currentSheet.getRange('A2').getValue() == "Date"){

      charts = currentSheet.getCharts();
      chart = charts[charts.length - 1];
      currentSheet.removeChart(chart);

      var DailyAverageTimeChart = currentSheet.newChart()
      .asLineChart()
      .addRange(currentSheet.getRange('A2:A'))
      .addRange(currentSheet.getRange('D2:D'))
      .addRange(currentSheet.getRange('H2:H'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(-1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('curveType', 'none')
      .setOption('domainAxis.direction', 1)
      .setOption('title', 'Daily Average Time vs. Date')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('series.0.labelInLegend', 'Average Time for 1')
      .setOption('series.1.labelInLegend', 'Cumulative Average Time for 1')
      .setPosition(38, 10, 14, 8)
      .build();
      currentSheet.insertChart(DailyAverageTimeChart);
    }
  }
};

function CreateGraphsforallDays() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++){
    var currentSheet = sheets[i];

    if (currentSheet.getRange('C1').getValue() == "Undo 1"){

      var increment = 4;
      while (currentSheet.getRange('G' + increment).getValue() != ""){
        increment++;
      }
      increment--;

      currentSheet.getRange('G4:N' + increment).sort({column: 8, ascending: false});

      currentSheet.getRange('G3:H' + increment).activate();
      var chart = currentSheet.newChart()
      .asColumnChart()
      .addRange(currentSheet.getRange('G3:H' + increment))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'Time vs. Event')
      .setXAxisTitle('Event')
      .setYAxisTitle('Time')
      .setOption('height', 481)
      .setOption('width', 934)
      .setPosition(31, 6, 99, 0)
      .build();
      currentSheet.insertChart(chart);
    }
  }
};
