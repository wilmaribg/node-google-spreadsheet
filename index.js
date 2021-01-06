const GoogleSpreadsheet = require('./lib2/GoogleSpreadsheet');
const GoogleSpreadsheetWorksheet = require('./lib2/GoogleSpreadsheetWorksheet');
const GoogleSpreadsheetRow = require('./lib2/GoogleSpreadsheetRow');

const { GoogleSpreadsheetFormulaError } = require('./lib2/errors');

module.exports = {
  GoogleSpreadsheet,
  GoogleSpreadsheetWorksheet,
  GoogleSpreadsheetRow,

  GoogleSpreadsheetFormulaError,
};
