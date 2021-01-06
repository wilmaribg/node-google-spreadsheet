"use strict";

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var GoogleSpreadsheetFormulaError = function GoogleSpreadsheetFormulaError(errorInfo) {
  _classCallCheck(this, GoogleSpreadsheetFormulaError);

  this.type = errorInfo.type;
  this.message = errorInfo.message;
};

module.exports = {
  GoogleSpreadsheetFormulaError: GoogleSpreadsheetFormulaError
};