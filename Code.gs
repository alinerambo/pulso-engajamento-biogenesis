// Cole este código no editor do Google Apps Script (Extensões > Apps Script)
// da planilha Google Sheets que vai guardar as respostas.
// Depois: Implantar > Nova implantação > tipo "App da Web",
// executar como "Eu", acesso "Qualquer pessoa" — e copie a URL gerada.

const SHEET_NAME = 'Respostas';
const TOTAL_QUESTIONS = 9;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const header = ['timestamp'];
    for (let i = 1; i <= TOTAL_QUESTIONS; i++) header.push('q' + i);
    sheet.appendRow(header);
  }
  return sheet;
}

function doPost(e) {
  const sheet = getSheet_();
  const data = JSON.parse(e.postData.contents);
  const row = [new Date()];
  for (let i = 1; i <= TOTAL_QUESTIONS; i++) {
    row.push(Number(data['q' + i]));
  }
  sheet.appendRow(row);
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'results') {
    return ContentService
      .createTextOutput(JSON.stringify(computeResults_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function computeResults_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1); // remove cabeçalho
  const n = rows.length;

  const sums = new Array(TOTAL_QUESTIONS).fill(0);
  const counts = new Array(TOTAL_QUESTIONS).fill(0);
  rows.forEach(r => {
    for (let i = 0; i < TOTAL_QUESTIONS; i++) {
      const v = Number(r[i + 1]);
      if (!isNaN(v) && v > 0) {
        sums[i] += v;
        counts[i]++;
      }
    }
  });
  const questionAvgs = sums.map((s, i) => counts[i] ? s / counts[i] : null);

  const dims = {
    vigor: [0, 1, 2],
    dedicacao: [3, 4, 5],
    absorcao: [6, 7, 8]
  };
  const dimensionAvgs = {};
  Object.keys(dims).forEach(key => {
    const vals = dims[key].map(i => questionAvgs[i]).filter(v => v != null);
    dimensionAvgs[key] = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  });

  const dimVals = Object.values(dimensionAvgs).filter(v => v != null);
  const overallAvg = dimVals.length ? dimVals.reduce((a, b) => a + b, 0) / dimVals.length : null;

  return { count: n, questionAvgs, dimensionAvgs, overallAvg };
}
