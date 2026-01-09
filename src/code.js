function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('RoB 2')
    .addItem('Assessment Form', 'showForm')
    .addItem('Summary', 'populateSummary')
    .addItem('Figures', 'generateFigures')
    .addItem('Generating a Print View', 'generatePrintView')
    .addItem('Discrepancy Check', 'showDiscrepancyCheckForm')
    .addToUi();
}

function showDiscrepancyCheckForm() {
  const html = HtmlService.createHtmlOutputFromFile('discrepancyCheckForm')
    .setWidth(1000)
    .setHeight(475);
  SpreadsheetApp.getUi().showModalDialog(html, 'RoB 2 discrepancy check');
}

function showForm() {
  const html = HtmlService.createHtmlOutputFromFile('form')
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'RoB 2 assessment for individual randomized, parallel group trials');
}

function populateIttTable(sheet, record, rowOffset) {
  
  const mapping = {
    1: {r: 1, c: 2}, 3: {r: 1, c: 4}, 2: {r: 1, c: 6}, 4: {r: 2, c: 2}, 9: {r: 2, c: 4},
    5: {r: 3, c: 2}, 6: {r: 3, c: 4}, 12: {r: 3, c: 6}, 7: {r: 4, c: 2}, 8: {r: 4, c: 4},
    11: {r: 4, c: 6}, 13: {r: 6, c: 5}, 14: {r: 7, c: 5}, 15: {r: 6, c: 6}, 16: {r: 8, c: 5},
    17: {r: 8, c: 6}, 19: {r: 9, c: 5}, 20: {r: 9, c: 6}, 23: {r: 10, c: 5}, 24: {r: 11, c: 5},
    25: {r: 10, c: 6}, 26: {r: 12, c: 5}, 27: {r: 12, c: 6}, 28: {r: 13, c: 5}, 29: {r: 13, c: 6},
    30: {r: 14, c: 5}, 31: {r: 14, c: 6}, 32: {r: 15, c: 5}, 33: {r: 15, c: 6}, 34: {r: 16, c: 5},
    35: {r: 16, c: 6}, 37: {r: 17, c: 5}, 38: {r: 17, c: 6}, 41: {r: 18, c: 5}, 42: {r: 18, c: 6},
    43: {r: 19, c: 5}, 44: {r: 19, c: 6}, 45: {r: 20, c: 5}, 46: {r: 20, c: 6}, 47: {r: 21, c: 5},
    49: {r: 22, c: 5}, 50: {r: 22, c: 6}, 53: {r: 23, c: 5}, 54: {r: 23, c: 6}, 55: {r: 24, c: 5},
    56: {r: 24, c: 6}, 57: {r: 25, c: 5}, 58: {r: 25, c: 6}, 59: {r: 26, c: 5}, 60: {r: 26, c: 6},
    61: {r: 27, c: 5}, 63: {r: 28, c: 5}, 64: {r: 28, c: 6}, 67: {r: 29, c: 5}, 68: {r: 29, c: 6},
    69: {r: 30, c: 5}, 70: {r: 30, c: 6}, 71: {r: 31, c: 5}, 72: {r: 31, c: 6}, 74: {r: 32, c: 5},
    75: {r: 32, c: 6}, 79: {r: 33, c: 5}, 80: {r: 33, c: 6}
  };

  for (const colIndex in mapping) {
    const target = mapping[colIndex];
    const value = record[colIndex];
    sheet.getRange(target.r + rowOffset, target.c).setValue(value);
  }
}

function populatePpTable(sheet, record, rowOffset) {
  const mapping = {
    1: {r: 1, c: 2}, 3: {r: 1, c: 4}, 2: {r: 1, c: 6}, 4: {r: 2, c: 2}, 9: {r: 2, c: 4},
    5: {r: 3, c: 2}, 6: {r: 3, c: 4}, 12: {r: 3, c: 6}, 7: {r: 4, c: 2}, 8: {r: 4, c: 4},
    11: {r: 4, c: 6}, 13: {r: 6, c: 5}, 14: {r: 7, c: 5}, 15: {r: 6, c: 6}, 16: {r: 8, c: 5},
    17: {r: 8, c: 6}, 19: {r: 9, c: 5}, 20: {r: 9, c: 6}, 23: {r: 10, c: 5}, 24: {r: 11, c: 5},
    25: {r: 10, c: 6}, 26: {r: 12, c: 5}, 27: {r: 12, c: 6}, 28: {r: 13, c: 5}, 29: {r: 13, c: 6},
    30: {r: 14, c: 5}, 31: {r: 14, c: 6}, 32: {r: 15, c: 5}, 33: {r: 15, c: 6}, 37: {r: 16, c: 5},
    38: {r: 16, c: 6}, 41: {r: 17, c: 5}, 42: {r: 17, c: 6}, 43: {r: 18, c: 5}, 44: {r: 18, c: 6},
    45: {r: 19, c: 5}, 46: {r: 19, c: 6}, 47: {r: 20, c: 5}, 49: {r: 21, c: 5}, 50: {r: 21, c: 6},
    53: {r: 22, c: 5}, 54: {r: 22, c: 6}, 55: {r: 23, c: 5}, 56: {r: 23, c: 6}, 57: {r: 24, c: 5},
    58: {r: 24, c: 6}, 59: {r: 25, c: 5}, 60: {r: 25, c: 6}, 61: {r: 26, c: 5}, 63: {r: 27, c: 5},
    64: {r: 27, c: 6}, 67: {r: 28, c: 5}, 68: {r: 28, c: 6}, 69: {r: 29, c: 5}, 70: {r: 29, c: 6},
    71: {r: 30, c: 5}, 72: {r: 30, c: 6}, 74: {r: 31, c: 5}, 75: {r: 31, c: 6}, 79: {r: 32, c: 5},
    80: {r: 32, c: 6}
  };

  for (const colIndex in mapping) {
    const target = mapping[colIndex];
    const value = record[colIndex];
    sheet.getRange(target.r + rowOffset, target.c).setValue(value);
  }
}

const ITT_SHEET_NAME = "Print (ITT)";
const PP_SHEET_NAME = "Print (PP)";
const RESULTS_SHEET_NAME = "Results";
const ITT_CRITERIA = "assignment to intervention (the 'intention-to-treat' effect)";
const PP_CRITERIA = "adhering to intervention (the 'per-protocol' effect)";


function generatePrintView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultsSheet = ss.getSheetByName(RESULTS_SHEET_NAME);
  const ittSheet = ss.getSheetByName(ITT_SHEET_NAME);
  const ppSheet = ss.getSheetByName(PP_SHEET_NAME);

  if (!resultsSheet || !ittSheet || !ppSheet) {
    SpreadsheetApp.getUi().alert("One or more required tabs (Results, Print (ITT), Print (PP)) were not found.");
    return;
  }


  if (ittSheet.getLastRow() > 33) {
    ittSheet.getRange(34, 1, ittSheet.getLastRow() - 33, ittSheet.getLastColumn()).clear();
  }

  if (ppSheet.getLastRow() > 32) {
    ppSheet.getRange(33, 1, ppSheet.getLastRow() - 32, ppSheet.getLastColumn()).clear();
  }
  SpreadsheetApp.flush();


  const resultsData = resultsSheet.getDataRange().getValues();
  resultsData.shift();

  const ittRecords = [];
  const ppRecords = [];
  

  for (const row of resultsData) {
    if (row[9] === ITT_CRITERIA) {
      ittRecords.push(row);
    } else if (row[9] === PP_CRITERIA) {
      ppRecords.push(row);
    }
  }


  if (ittRecords.length > 0) {
    const ittTemplateRange = ittSheet.getRange("A1:F33");
    const ittTableHeight = 33;
    const rowGap = 2;
    const ittOffset = ittTableHeight + rowGap;


    for (let i = 1; i < ittRecords.length; i++) {
      const destination = ittSheet.getRange(1 + (i * ittOffset), 1);
      ittTemplateRange.copyTo(destination);
    }


    ittRecords.forEach((record, index) => {
      const rowOffset = index * ittOffset;
      populateIttTable(ittSheet, record, rowOffset);
    });
  }


  if (ppRecords.length > 0) {
    const ppTemplateRange = ppSheet.getRange("A1:F32");
    const ppTableHeight = 32;
    const rowGap = 2;
    const ppOffset = ppTableHeight + rowGap;

    for (let i = 1; i < ppRecords.length; i++) {
      const destination = ppSheet.getRange(1 + (i * ppOffset), 1);
      ppTemplateRange.copyTo(destination);
    }

    ppRecords.forEach((record, index) => {
      const rowOffset = index * ppOffset;
      populatePpTable(ppSheet, record, rowOffset);
    });
  }
  
  SpreadsheetApp.getUi().alert("Print previews successfully generated!");
}


function resultSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('Summary');
  const resultsSheet = ss.getSheetByName('Results');

  if (!summarySheet || !resultsSheet) {
    SpreadsheetApp.getUi().alert('Error: Check if the "Summary" and "Results" tabs exist.');
    return;
  }

  const summaryStartRow = 2;
  const summaryDataRange = summarySheet.getRange(summaryStartRow, 1, summarySheet.getLastRow() - summaryStartRow + 1, 23);
  const summaryValues = summaryDataRange.getValues();

  const resultsStartRow = 3;
  const resultsDataRange = resultsSheet.getRange(resultsStartRow, 1, resultsSheet.getLastRow() - resultsStartRow + 1, resultsSheet.getLastColumn());
  const resultsValues = resultsDataRange.getValues();

  if (summaryValues.length === 0 || resultsValues.length === 0) {
    SpreadsheetApp.getUi().alert('There is no data to sync.');
    return;
  }
  
  const reverseMap = {
    0: 1,
    1: 2,
    2: 3,
    3: 4,
    4: 5,
    5: 6,
    6: 7,
    7: 8,
    8: 9,
    9: 10,
    10: 11,
    11: 19,
    12: 20,
    13: 37,
    14: 38,
    15: 49,
    16: 50,
    17: 63,
    18: 64,
    19: 74,
    20: 75,
    21: 79,
    22: 80
  };

  const rowsToSync = Math.min(summaryValues.length, resultsValues.length);

  for (let i = 0; i < rowsToSync; i++) {
    const summaryRow = summaryValues[i];

    for (const summaryColIndex in reverseMap) {
      const resultsColIndex = reverseMap[summaryColIndex];
      const valueToSync = summaryRow[summaryColIndex];
      
      resultsValues[i][resultsColIndex] = valueToSync;
    }
  }

  resultsDataRange.setValues(resultsValues);

  calculateRiskTable();
  
  SpreadsheetApp.getUi().alert(`Sync complete! ${rowsToSync} 6 rows were updated in 'Results' and the risk table was recalculated.`);
}

function calculateRiskTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('Summary');

  if (!summarySheet) {
    console.error('The "Summary" tab was not found.');
    return;
  }
  
  const dataRange = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 22);
  const summaryValues = dataRange.getValues();

  const GROUP_COL = 8;
  const WEIGHT_COL = 10;
  const DOMAIN_COLS = [11, 13, 15, 17, 19, 21];

  const ITT_GROUP = "assignment to intervention (the 'intention-to-treat' effect)";
  const PP_GROUP = "adhering to intervention (the 'per-protocol' effect)";

  let riskTotals = {
    itt: { "Low": [0,0,0,0,0,0], "Some concerns": [0,0,0,0,0,0], "High": [0,0,0,0,0,0] },
    pp:  { "Low": [0,0,0,0,0,0], "Some concerns": [0,0,0,0,0,0], "High": [0,0,0,0,0,0] }
  };
  
  let ittCount = 0;
  let ppCount = 0;

  summaryValues.forEach(row => {
    const group = row[GROUP_COL];
    const weight = parseFloat(row[WEIGHT_COL]);

    if (isNaN(weight)) {
      return; 
    }

    const valueToAdd = weight * 100;
    let targetGroup = null;

    if (group === ITT_GROUP) {
      targetGroup = riskTotals.itt;
      ittCount++;
    } else if (group === PP_GROUP) {
      targetGroup = riskTotals.pp;
      ppCount++;
    }

    if (targetGroup) {
      DOMAIN_COLS.forEach((colIndex, domainIndex) => {
        const riskLevel = row[colIndex];
        if (targetGroup[riskLevel]) {
          targetGroup[riskLevel][domainIndex] += valueToAdd;
        }
      });
    }
  });

  const ittHeaderText = `Assignment to intervention (the 'intention-to-treat' effect) - [total: ${ittCount}]`;
  const ppHeaderText = `Adhering to intervention (the 'per-protocol' effect) - [total: ${ppCount}]`;

  const ittResults = [
    riskTotals.itt["Low"],
    riskTotals.itt["Some concerns"],
    riskTotals.itt["High"]
  ];

  const ppResults = [
    riskTotals.pp["Low"],
    riskTotals.pp["Some concerns"],
    riskTotals.pp["High"]
  ];

  summarySheet.getRange('AB2').setValue(ittHeaderText);
  summarySheet.getRange('AB3:AG5').clearContent().setValues(ittResults);
  summarySheet.getRange('AB6').setValue(ppHeaderText);
  summarySheet.getRange('AB7:AG9').clearContent().setValues(ppResults);
}

function populateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultsSheet = ss.getSheetByName('Results');
  const summarySheet = ss.getSheetByName('Summary');

  if (!resultsSheet || !summarySheet) {
    SpreadsheetApp.getUi().alert('Error: Check if the "Results" and "Summary" tabs exist in your spreadsheet.');
    return;
  }

  const startRow = 3;
  const numRows = resultsSheet.getLastRow() - startRow + 1;

  if (numRows <= 0) {
    SpreadsheetApp.getUi().alert('There is no data to process in the "Results" tab from line 3 onwards.');
    return;
  }

  const resultsRange = resultsSheet.getRange(startRow, 1, numRows, resultsSheet.getLastColumn());
  const resultsValues = resultsRange.getValues();

  const columnMapping = [
    1,
    2,
    3,
    4,
    5,
    6,
    7,
    8,
    9,
    10,
    11,
    19,
    20,
    37,
    38,
    49,
    50,
    63,
    64,
    74,
    75,
    79,
    80
  ];

  const summaryData = resultsValues.map(originalRow => {
    return columnMapping.map(colIndex => originalRow[colIndex]);
  });

  if (summaryData.length > 0) {
    if (summarySheet.getLastRow() > 1) {
      summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 23).clearContent();
    }
    
    summarySheet.getRange(2, 1, summaryData.length, summaryData[0].length).setValues(summaryData);

    calculateRiskTable();

    SpreadsheetApp.getUi().alert('The "Summary" tab has been updated successfully!');
  }
}

function generateFigures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultsSheet = ss.getSheetByName('Results');
  const figureITTSheet = ss.getSheetByName('Figure (ITT)');
  const figurePPSheet = ss.getSheetByName('Figure (PP)');


  if (figureITTSheet.getLastRow() > 1) {
    figureITTSheet.getRange('A2:I' + figureITTSheet.getLastRow()).clearContent();
  }
  if (figurePPSheet.getLastRow() > 1) {
    figurePPSheet.getRange('A2:I' + figurePPSheet.getLastRow()).clearContent();
  }

  const resultsData = resultsSheet.getRange('B3:CB' + resultsSheet.getLastRow()).getValues();

  let nextIttRow = 2;
  let nextPpRow = 2;

  const lowRiskImg = 'K2';
  const someConcernsImg = 'K3';
  const highRiskImg = 'K4';

  for (let i = 0; i < resultsData.length; i++) {
    const row = resultsData[i];

    const studyId = row[0];
    const aod = row[6];
    const analysisType = row[8];
    const riskT = row[18];
    const riskAL = row[36];
    const riskAX = row[48];
    const riskBL = row[62];
    const riskBW = row[73];
    const riskCB = row[78];

    if (studyId && analysisType && aod && riskT && riskAL && riskAX && riskBL && riskBW && riskCB) {

      let targetSheet;
      let targetRow;

      // 5. Determinar a aba de destino
      if (analysisType === "assignment to intervention (the 'intention-to-treat' effect)") {
        targetSheet = figureITTSheet;
        targetRow = nextIttRow;
        nextIttRow++;
      } else if (analysisType === "adhering to intervention (the 'per-protocol' effect)") {
        targetSheet = figurePPSheet;
        targetRow = nextPpRow;
        nextPpRow++;
      }

      if (targetSheet) {

        targetSheet.getRange('B' + targetRow).setValue(studyId);
        targetSheet.getRange('C' + targetRow).setValue(aod);

        const riskValues = [riskT, riskAL, riskAX, riskBL, riskBW, riskCB];
        const targetColumns = ['D', 'E', 'F', 'G', 'H', 'I'];

        for (let j = 0; j < riskValues.length; j++) {
          let imageCell;
          if (riskValues[j] === 'Low') {
            imageCell = targetSheet.getRange(lowRiskImg);
          } else if (riskValues[j] === 'Some concerns') {
            imageCell = targetSheet.getRange(someConcernsImg);
          } else if (riskValues[j] === 'High') {
            imageCell = targetSheet.getRange(highRiskImg);
          }

          if (imageCell) {
            imageCell.copyTo(targetSheet.getRange(targetColumns[j] + targetRow));
          }
        }
      }
    }
  }
}

function getColumnOrder() {
  return [
    'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j',
    'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
    'u', 'v', 'w', 'x', 'y', 'z', 'aa', 'ab', 'ac', 'ad',
    'ae', 'af', 'ag', 'ah', 'ai', 'aj', 'ak', 'al', 'am',
    'an', 'ao', 'ap', 'aq', 'ar', 'as', 'at', 'au', 'av',
    'aw', 'ax', 'ay', 'az', 'ba', 'bb', 'bc', 'bd', 'be',
    'bf', 'bg', 'bh', 'bi', 'bj', 'bk', 'bl', 'bm', 'bn',
    'bo', 'bp', 'bq', 'br', 'bs', 'bt', 'bu', 'bv', 'bw',
    'bx', 'by', 'bz', 'ca', 'cb', 'cc', 'cd', 'ce'
  ];
}

function saveDataToSheet(data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Results";
    const sheet = spreadsheet.getSheetByName(sheetName);

    const columnOrder = getColumnOrder();
  
    const rowData = columnOrder.map(key => data[key] || "");

    sheet.appendRow(rowData);
  } catch (e) {
    Logger.log(`Error saving: ${e.message}`);
    throw e;
  }
}


function atualizarLinhaPorValor(data, valorParaLocalizar) {
  try {
    const sheetName = "Results";
    const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const primeiraLinhaDados = 3;
    const colunaBusca = 2;

    const valoresColunaB = planilha.getRange(primeiraLinhaDados, colunaBusca, planilha.getLastRow() - primeiraLinhaDados + 1, 1).getValues();

    let linhaEncontrada = -1;

    for (let i = 0; i < valoresColunaB.length; i++) {
      if (valoresColunaB[i][0] === valorParaLocalizar) {
        linhaEncontrada = primeiraLinhaDados + i;
        break;
      }
    }

    if (linhaEncontrada !== -1) {

      const columnOrder = getColumnOrder();

      const valoresParaDefinir = columnOrder.map(key => data[key] || "");

      const rangeParaAtualizar = planilha.getRange(linhaEncontrada, 1, 1, valoresParaDefinir.length);
      rangeParaAtualizar.setValues([valoresParaDefinir]);

      return true;
    } else {
      return false;
    }
  } catch (e) {
    Logger.log(`Error to update record: ${e.message}`);
    throw e;
  }
}

function getRecordIds(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`The tab "${sheetName}" was not found.`);
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return [];
  }

  let column = 83;
  const headerRow = sheet.getRange(2, 1, 1, column).getValues()[0];

  const dataRange = sheet.getRange(3, 1, lastRow - 2, column);
  const allValues = dataRange.getValues();

  const records = [];

  for (let i = 0; i < allValues.length; i++) {
    const row = allValues[i];

    const columnBValue = row[1];

    if (columnBValue !== null && String(columnBValue).trim() !== "") {
      const record = {};

      for (let j = 0; j < headerRow.length; j++) {

        const key = headerRow[j] ? String(headerRow[j]).trim() : `Coluna${j + 1}`;
        record[key] = row[j];
      }
      records.push(record);
    }

  }
  Logger.log(`DEBUG: ${JSON.stringify(records[10])}`);
  return JSON.stringify(records);
}