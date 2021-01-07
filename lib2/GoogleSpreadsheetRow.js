'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _require = require('./utils'),
    columnToLetter = _require.columnToLetter;

var GoogleSpreadsheetRow = function () {
  function GoogleSpreadsheetRow(parentSheet, rowNumber, data) {
    var _this = this;

    _classCallCheck(this, GoogleSpreadsheetRow);

    this._sheet = parentSheet; // the parent GoogleSpreadsheetWorksheet instance
    this._rowNumber = rowNumber; // the A1 row (1-indexed)
    this._rawData = data;

    var _loop = function _loop(i) {
      var propName = _this._sheet.headerValues[i];
      if (!propName) return 'continue'; // skip empty header
      Object.defineProperty(_this, propName, {
        get: function get() {
          return _this._rawData[i];
        },
        set: function set(newVal) {
          _this._rawData[i] = newVal;
        },
        enumerable: true
      });
    };

    for (var i = 0; i < this._sheet.headerValues.length; i++) {
      var _ret = _loop(i);

      if (_ret === 'continue') continue;
    }

    return this;
  }

  _createClass(GoogleSpreadsheetRow, [{
    key: 'save',
    value: async function save() {
      var options = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};

      if (this._deleted) throw new Error('This row has been deleted - call getRows again before making updates.');

      var response = await this._sheet._spreadsheet.axios.request({
        method: 'put',
        url: '/values/' + encodeURIComponent(this.a1Range),
        params: {
          valueInputOption: options.raw ? 'RAW' : 'USER_ENTERED',
          includeValuesInResponse: true
        },
        data: {
          range: this.a1Range,
          majorDimension: 'ROWS',
          values: [this._rawData]
        }
      });
      this._rawData = response.data.updatedData.values[0];
    }

    // delete this row

  }, {
    key: 'delete',
    value: async function _delete() {
      if (this._deleted) throw new Error('This row has been deleted - call getRows again before making updates.');

      var result = await this._sheet._makeSingleUpdateRequest('deleteRange', {
        range: {
          sheetId: this._sheet.sheetId,
          startRowIndex: this._rowNumber - 1, // this format is zero indexed, because of course...
          endRowIndex: this._rowNumber
        },
        shiftDimension: 'ROWS'
      });
      this._deleted = true;
      return result;
    }
  }, {
    key: 'del',
    value: async function del() {
      return this.delete();
    } // alias to mimic old version of this module

  }, {
    key: 'rowNumber',
    get: function get() {
      return this._rowNumber;
    }
    // TODO: deprecate rowIndex - the name implies it should be zero indexed :(

  }, {
    key: 'rowIndex',
    get: function get() {
      return this._rowNumber;
    }
  }, {
    key: 'a1Range',
    get: function get() {
      return [this._sheet.a1SheetName, '!', 'A' + this._rowNumber, ':', '' + columnToLetter(this._sheet.headerValues.length) + this._rowNumber].join('');
    }
  }]);

  return GoogleSpreadsheetRow;
}();

module.exports = GoogleSpreadsheetRow;