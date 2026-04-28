function doGet() {
  return HtmlService.createHtmlOutputFromFile('metrics_dashboard')
    .setTitle('Otto Team Metrics Dashboard');
}

const SALES_SHEETS = {
  all: 'Sales - All',
  otto: 'Sales - Otto',
  weflex: 'Sales - Weflex',
  lcm: 'Sales - LCM',
  tfs: 'Sales - TFS',
  'triple-point': 'Sales - TP'
};

const OPS_SHEETS = {
  all: 'All',
  otto: 'Otto',
  weflex: 'WF',
  tfs: 'TFS',
  lcm: 'LCM',
  'triple-point': 'TP'
};

function getSalesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};

  Object.entries(SALES_SHEETS).forEach(([key, sheetName]) => {
    result[key] = readSheetRows_(ss, sheetName);
  });

  return result;
}

function getOpsData() {
  const opsSpreadsheetId = '15OzeUS0ZX5C3h6dN7EPWMEHG48rYIZl9_af2XRP3dTM';
  const opsSS = SpreadsheetApp.openById(opsSpreadsheetId);
  const result = {};

  Object.entries(OPS_SHEETS).forEach(([key, sheetName]) => {
    result[key] = readSheetRows_(opsSS, sheetName);
  });

  return result;
}

function getFreshdeskData() {
  const freshdeskSpreadsheetId = '1L-1DjkTWxLULB2GAaSJIxAz6xuGbwBHrksQrkEH05l8';
  const freshdeskSS = SpreadsheetApp.openById(freshdeskSpreadsheetId);

  return readSheetRows_(freshdeskSS, 'Freshdesk - all');
}

function getServicingData() {
  const servicingSpreadsheetId = '13dsrbWG55hc9F5spG3gusKk7E78IBkBoMRy6XX317oo';
  const servicingSS = SpreadsheetApp.openById(servicingSpreadsheetId);

  const servicingSheets = {
    all: 'SERVICING - All',
    otto: 'SERVICING - Otto',
    weflex: 'SERVICING - ALL WF',
    'triple-point': 'SERVICING - TP',
    tfs: 'SERVICING - TFS',
    lcm: 'SERVICING - LCM'
  };

  const result = {};

  Object.entries(servicingSheets).forEach(([key, sheetName]) => {
    result[key] = readSheetRows_(servicingSS, sheetName);
  });

  return result;
}

function readSheetRows_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    const availableSheets = ss.getSheets().map(s => s.getName()).join(', ');
    throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${availableSheets}`);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(header => String(header).trim());

  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const record = {};

      headers.forEach((header, index) => {
        const value = row[index];

        record[header] = value instanceof Date
          ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : value;
      });

      return record;
    });
}
