'use strict';

var _extends = Object.assign || function (target) { for (var i = 1; i < arguments.length; i++) { var source = arguments[i]; for (var key in source) { if (Object.prototype.hasOwnProperty.call(source, key)) { target[key] = source[key]; } } } return target; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _defineProperty(obj, key, value) { if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable: true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require('lodash');

var _require = require('./utils'),
    columnToLetter = _require.columnToLetter;

var _require2 = require('./errors'),
    GoogleSpreadsheetFormulaError = _require2.GoogleSpreadsheetFormulaError;

var GoogleSpreadsheetCell = function () {
  function GoogleSpreadsheetCell(parentSheet, rowIndex, columnIndex, cellData) {
    _classCallCheck(this, GoogleSpreadsheetCell);

    this._sheet = parentSheet; // the parent GoogleSpreadsheetWorksheet instance
    this._row = rowIndex;
    this._column = columnIndex;

    this._updateRawData(cellData);
    return this;
  }

  // newData can be undefined/null if the cell is totally empty and unformatted


  _createClass(GoogleSpreadsheetCell, [{
    key: '_updateRawData',
    value: function _updateRawData() {
      var newData = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};

      this._rawData = newData;
      this._draftData = {}; // stuff to save
      this._error = null;
      if (_.get(this._rawData, 'effectiveValue.errorValue')) {
        this._error = new GoogleSpreadsheetFormulaError(this._rawData.effectiveValue.errorValue);
      }
    }

    // CELL LOCATION/ADDRESS /////////////////////////////////////////////////////////////////////////

  }, {
    key: '_getFormatParam',
    value: function _getFormatParam(param) {
      // we freeze the object so users don't change nested props accidentally
      // TODO: figure out something that would throw an error if you try to update it?
      if (_.get(this._draftData, 'userEnteredFormat.' + param)) {
        throw new Error('User format is unsaved - save the cell to be able to read it again');
      }
      return Object.freeze(this._rawData.userEnteredFormat[param]);
    }
  }, {
    key: '_setFormatParam',
    value: function _setFormatParam(param, newVal) {
      if (_.isEqual(newVal, _.get(this._rawData, 'userEnteredFormat.' + param))) {
        _.unset(this._draftData, 'userEnteredFormat.' + param);
      } else {
        _.set(this._draftData, 'userEnteredFormat.' + param, newVal);
        this._draftData.clearFormat = false;
      }
    }

    // format getters

  }, {
    key: 'clearAllFormatting',
    value: function clearAllFormatting() {
      // need to track this separately since by setting/unsetting things, we may end up with
      // this._draftData.userEnteredFormat as an empty object, but not an intent to clear it
      this._draftData.clearFormat = true;
      delete this._draftData.userEnteredFormat;
    }

    // SAVING + UTILS ////////////////////////////////////////////////////////////////////////////////

    // returns true if there are any updates that have not been saved yet

  }, {
    key: 'discardUnsavedChanges',
    value: function discardUnsavedChanges() {
      this._draftData = {};
    }
  }, {
    key: 'save',
    value: async function save() {
      await this._sheet.saveUpdatedCells([this]);
    }

    // used by worksheet when saving cells
    // returns an individual batchUpdate request to update the cell

  }, {
    key: '_getUpdateRequest',
    value: function _getUpdateRequest() {
      // this logic should match the _isDirty logic above
      // but we need it broken up to build the request below
      var isValueUpdated = this._draftData.value !== undefined;
      var isNoteUpdated = this._draftData.note !== undefined;
      var isFormatUpdated = !!_.keys(this._draftData.userEnteredFormat || {}).length;
      var isFormatCleared = this._draftData.clearFormat;

      // if no updates, we return null, which we can filter out later before sending requests
      if (!_.some([isValueUpdated, isNoteUpdated, isFormatUpdated, isFormatCleared])) {
        return null;
      }

      // build up the formatting object, which has some quirks...
      var format = _extends({}, this._rawData.userEnteredFormat, this._draftData.userEnteredFormat);
      // if background color already set, cell has backgroundColor and backgroundColorStyle
      // but backgroundColorStyle takes precendence so we must remove to set the color
      // see https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#CellFormat
      if (_.get(this._draftData, 'userEnteredFormat.backgroundColor')) {
        delete format.backgroundColorStyle;
      }

      return {
        updateCells: {
          rows: [{
            values: [_extends({}, isValueUpdated && {
              userEnteredValue: _defineProperty({}, this._draftData.valueType, this._draftData.value)
            }, isNoteUpdated && {
              note: this._draftData.note
            }, isFormatUpdated && {
              userEnteredFormat: format
            }, isFormatCleared && {
              userEnteredFormat: {}
            })]
          }],
          // turns into a string of which fields to update ex "note,userEnteredFormat"
          fields: _.keys(_.pickBy({
            userEnteredValue: isValueUpdated,
            note: isNoteUpdated,
            userEnteredFormat: isFormatUpdated || isFormatCleared
          })).join(','),
          start: {
            sheetId: this._sheet.sheetId,
            rowIndex: this.rowIndex,
            columnIndex: this.columnIndex
          }
        }
      };
    }
  }, {
    key: 'rowIndex',
    get: function get() {
      return this._row;
    }
  }, {
    key: 'columnIndex',
    get: function get() {
      return this._column;
    }
  }, {
    key: 'a1Column',
    get: function get() {
      return columnToLetter(this._column + 1);
    }
  }, {
    key: 'a1Row',
    get: function get() {
      return this._row + 1;
    } // a1 row numbers start at 1 instead of 0

  }, {
    key: 'a1Address',
    get: function get() {
      return '' + this.a1Column + this.a1Row;
    }

    // CELL CONTENTS - VALUE/FORMULA/NOTES ///////////////////////////////////////////////////////////

  }, {
    key: 'value',
    get: function get() {
      // const typeKey = _.keys(this._rawData.effectiveValue)[0];
      if (this._draftData.value !== undefined) throw new Error('Value has been changed');
      if (this._error) return this._error;
      if (!this._rawData.effectiveValue) return null;
      return _.values(this._rawData.effectiveValue)[0];
    },
    set: function set(newValue) {
      if (_.isBoolean(newValue)) {
        this._draftData.valueType = 'boolValue';
      } else if (_.isString(newValue)) {
        if (newValue.substr(0, 1) === '=') this._draftData.valueType = 'formulaValue';else this._draftData.valueType = 'stringValue';
      } else if (_.isFinite(newValue)) {
        this._draftData.valueType = 'numberValue';
      } else if (_.isNil(newValue)) {
        // null or undefined
        this._draftData.valueType = 'stringValue';
        newValue = '';
      } else {
        throw new Error('Set value to boolean, string, or number');
      }
      this._draftData.value = newValue;
    }
  }, {
    key: 'valueType',
    get: function get() {
      // an error only happens with a formula
      if (this._error) return 'errorValue';
      if (!this._rawData.effectiveValue) return null;
      return _.keys(this._rawData.effectiveValue)[0];
    }
  }, {
    key: 'formattedValue',
    get: function get() {
      return this._rawData.formattedValue || null;
    },
    set: function set(newVal) {
      throw new Error('You cannot modify the formatted value directly');
    }
  }, {
    key: 'formula',
    get: function get() {
      return _.get(this._rawData, 'userEnteredValue.formulaValue', null);
    },
    set: function set(newValue) {
      if (newValue.substr(0, 1) !== '=') throw new Error('formula must begin with "="');
      this.value = newValue; // use existing value setter
    }
  }, {
    key: 'formulaError',
    get: function get() {
      return this._error;
    }
  }, {
    key: 'hyperlink',
    get: function get() {
      if (this._draftData.value) throw new Error('Save cell to be able to read hyperlink');
      return this._rawData.hyperlink;
    },
    set: function set(val) {
      throw new Error('Do not set hyperlink directly. Instead set cell.formula, for example `cell.formula = \'=HYPERLINK("http://google.com", "Google")\'`');
    }
  }, {
    key: 'note',
    get: function get() {
      return this._draftData.note !== undefined ? this._draftData.note : this._rawData.note;
    },
    set: function set(newVal) {
      if (newVal === null || newVal === undefined) newVal = '';
      if (!_.isString(newVal)) throw new Error('Note must be a string');
      if (newVal === this._rawData.note) delete this._draftData.note;else this._draftData.note = newVal;
    }

    // CELL FORMATTING ///////////////////////////////////////////////////////////////////////////////

  }, {
    key: 'userEnteredFormat',
    get: function get() {
      return this._rawData.userEnteredFormat;
    },
    set: function set(newVal) {
      throw new Error('Do not modify directly, instead use format properties');
    }
  }, {
    key: 'effectiveFormat',
    get: function get() {
      return this._rawData.effectiveFormat;
    },
    set: function set(newVal) {
      throw new Error('Read-only');
    }
  }, {
    key: 'numberFormat',
    get: function get() {
      return this._getFormatParam('numberFormat');
    },


    // format setters
    set: function set(newVal) {
      return this._setFormatParam('numberFormat', newVal);
    }
  }, {
    key: 'backgroundColor',
    get: function get() {
      return this._getFormatParam('backgroundColor');
    },
    set: function set(newVal) {
      return this._setFormatParam('backgroundColor', newVal);
    }
  }, {
    key: 'borders',
    get: function get() {
      return this._getFormatParam('borders');
    },
    set: function set(newVal) {
      return this._setFormatParam('borders', newVal);
    }
  }, {
    key: 'padding',
    get: function get() {
      return this._getFormatParam('padding');
    },
    set: function set(newVal) {
      return this._setFormatParam('padding', newVal);
    }
  }, {
    key: 'horizontalAlignment',
    get: function get() {
      return this._getFormatParam('horizontalAlignment');
    },
    set: function set(newVal) {
      return this._setFormatParam('horizontalAlignment', newVal);
    }
  }, {
    key: 'verticalAlignment',
    get: function get() {
      return this._getFormatParam('verticalAlignment');
    },
    set: function set(newVal) {
      return this._setFormatParam('verticalAlignment', newVal);
    }
  }, {
    key: 'wrapStrategy',
    get: function get() {
      return this._getFormatParam('wrapStrategy');
    },
    set: function set(newVal) {
      return this._setFormatParam('wrapStrategy', newVal);
    }
  }, {
    key: 'textDirection',
    get: function get() {
      return this._getFormatParam('textDirection');
    },
    set: function set(newVal) {
      return this._setFormatParam('textDirection', newVal);
    }
  }, {
    key: 'textFormat',
    get: function get() {
      return this._getFormatParam('textFormat');
    },
    set: function set(newVal) {
      return this._setFormatParam('textFormat', newVal);
    }
  }, {
    key: 'hyperlinkDisplayType',
    get: function get() {
      return this._getFormatParam('hyperlinkDisplayType');
    },
    set: function set(newVal) {
      return this._setFormatParam('hyperlinkDisplayType', newVal);
    }
  }, {
    key: 'textRotation',
    get: function get() {
      return this._getFormatParam('textRotation');
    },
    set: function set(newVal) {
      return this._setFormatParam('textRotation', newVal);
    }
  }, {
    key: '_isDirty',
    get: function get() {
      // have to be careful about checking undefined rather than falsy
      // in case a new value is empty string or 0 or false
      if (this._draftData.note !== undefined) return true;
      if (_.keys(this._draftData.userEnteredFormat).length) return true;
      if (this._draftData.clearFormat) return true;
      if (this._draftData.value !== undefined) return true;
      return false;
    }
  }]);

  return GoogleSpreadsheetCell;
}();

module.exports = GoogleSpreadsheetCell;