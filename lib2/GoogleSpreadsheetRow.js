"use strict";

function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }

function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); return Constructor; }

var _require = require('./utils'),
    columnToLetter = _require.columnToLetter;

var GoogleSpreadsheetRow = /*#__PURE__*/function () {
  function GoogleSpreadsheetRow(parentSheet, rowNumber, data) {
    var _this = this;

    _classCallCheck(this, GoogleSpreadsheetRow);

    this._sheet = parentSheet; // the parent GoogleSpreadsheetWorksheet instance

    this._rowNumber = rowNumber; // the A1 row (1-indexed)

    this._rawData = data;

    var _loop = function _loop(i) {
      var propName = _this._sheet.headerValues[i];
      if (!propName) return "continue"; // skip empty header

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

      if (_ret === "continue") continue;
    }

    return this;
  }

  _createClass(GoogleSpreadsheetRow, [{
    key: "save",
    value: function () {
      var _save = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee() {
        var options,
            response,
            _args = arguments;
        return regeneratorRuntime.wrap(function _callee$(_context) {
          while (1) {
            switch (_context.prev = _context.next) {
              case 0:
                options = _args.length > 0 && _args[0] !== undefined ? _args[0] : {};

                if (!this._deleted) {
                  _context.next = 3;
                  break;
                }

                throw new Error('This row has been deleted - call getRows again before making updates.');

              case 3:
                _context.next = 5;
                return this._sheet._spreadsheet.axios.request({
                  method: 'put',
                  url: "/values/".concat(encodeURIComponent(this.a1Range)),
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

              case 5:
                response = _context.sent;
                this._rawData = response.data.updatedData.values[0];

              case 7:
              case "end":
                return _context.stop();
            }
          }
        }, _callee, this);
      }));

      function save() {
        return _save.apply(this, arguments);
      }

      return save;
    }() // delete this row

  }, {
    key: "delete",
    value: function () {
      var _delete2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee2() {
        var result;
        return regeneratorRuntime.wrap(function _callee2$(_context2) {
          while (1) {
            switch (_context2.prev = _context2.next) {
              case 0:
                if (!this._deleted) {
                  _context2.next = 2;
                  break;
                }

                throw new Error('This row has been deleted - call getRows again before making updates.');

              case 2:
                _context2.next = 4;
                return this._sheet._makeSingleUpdateRequest('deleteRange', {
                  range: {
                    sheetId: this._sheet.sheetId,
                    startRowIndex: this._rowNumber - 1,
                    // this format is zero indexed, because of course...
                    endRowIndex: this._rowNumber
                  },
                  shiftDimension: 'ROWS'
                });

              case 4:
                result = _context2.sent;
                this._deleted = true;
                return _context2.abrupt("return", result);

              case 7:
              case "end":
                return _context2.stop();
            }
          }
        }, _callee2, this);
      }));

      function _delete() {
        return _delete2.apply(this, arguments);
      }

      return _delete;
    }()
  }, {
    key: "del",
    value: function () {
      var _del = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee3() {
        return regeneratorRuntime.wrap(function _callee3$(_context3) {
          while (1) {
            switch (_context3.prev = _context3.next) {
              case 0:
                return _context3.abrupt("return", this["delete"]());

              case 1:
              case "end":
                return _context3.stop();
            }
          }
        }, _callee3, this);
      }));

      function del() {
        return _del.apply(this, arguments);
      }

      return del;
    }() // alias to mimic old version of this module

  }, {
    key: "rowNumber",
    get: function get() {
      return this._rowNumber;
    } // TODO: deprecate rowIndex - the name implies it should be zero indexed :(

  }, {
    key: "rowIndex",
    get: function get() {
      return this._rowNumber;
    }
  }, {
    key: "a1Range",
    get: function get() {
      return [this._sheet.a1SheetName, '!', "A".concat(this._rowNumber), ':', "".concat(columnToLetter(this._sheet.headerValues.length)).concat(this._rowNumber)].join('');
    }
  }]);

  return GoogleSpreadsheetRow;
}();

module.exports = GoogleSpreadsheetRow;