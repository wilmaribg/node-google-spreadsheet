"use strict";

function _toConsumableArray(arr) { return _arrayWithoutHoles(arr) || _iterableToArray(arr) || _unsupportedIterableToArray(arr) || _nonIterableSpread(); }

function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }

function _unsupportedIterableToArray(o, minLen) { if (!o) return; if (typeof o === "string") return _arrayLikeToArray(o, minLen); var n = Object.prototype.toString.call(o).slice(8, -1); if (n === "Object" && o.constructor) n = o.constructor.name; if (n === "Map" || n === "Set") return Array.from(o); if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen); }

function _iterableToArray(iter) { if (typeof Symbol !== "undefined" && Symbol.iterator in Object(iter)) return Array.from(iter); }

function _arrayWithoutHoles(arr) { if (Array.isArray(arr)) return _arrayLikeToArray(arr); }

function _arrayLikeToArray(arr, len) { if (len == null || len > arr.length) len = arr.length; for (var i = 0, arr2 = new Array(len); i < len; i++) { arr2[i] = arr[i]; } return arr2; }

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); if (enumerableOnly) symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; }); keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = arguments[i] != null ? arguments[i] : {}; if (i % 2) { ownKeys(Object(source), true).forEach(function (key) { _defineProperty(target, key, source[key]); }); } else if (Object.getOwnPropertyDescriptors) { Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)); } else { ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } } return target; }

function _defineProperty(obj, key, value) { if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable: true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }

function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }

function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); return Constructor; }

var _ = require('lodash');

var GoogleSpreadsheetRow = require('./GoogleSpreadsheetRow');

var GoogleSpreadsheetCell = require('./GoogleSpreadsheetCell');

var _require = require('./utils'),
    getFieldMask = _require.getFieldMask,
    columnToLetter = _require.columnToLetter,
    letterToColumn = _require.letterToColumn;

function checkForDuplicateHeaders(headers) {
  // check for duplicate headers
  var checkForDupes = _.groupBy(headers); // { c1: ['c1'], c2: ['c2', 'c2' ]}


  _.each(checkForDupes, function (grouped, header) {
    if (!header) return; // empty columns are skipped, so multiple is ok

    if (grouped.length > 1) {
      throw new Error("Duplicate header detected: \"".concat(header, "\". Please make sure all non-empty headers are unique"));
    }
  });
}

var GoogleSpreadsheetWorksheet = /*#__PURE__*/function () {
  function GoogleSpreadsheetWorksheet(parentSpreadsheet, _ref) {
    var properties = _ref.properties,
        data = _ref.data;

    _classCallCheck(this, GoogleSpreadsheetWorksheet);

    this._spreadsheet = parentSpreadsheet; // the parent GoogleSpreadsheet instance
    // basic properties

    this._rawProperties = properties;
    this._cells = []; // we will use a 2d sparse array to store cells;

    this._rowMetadata = []; // 1d sparse array

    this._columnMetadata = [];
    if (data) this._fillCellData(data);
    return this;
  } // INTERNAL UTILITY FUNCTIONS ////////////////////////////////////////////////////////////////////


  _createClass(GoogleSpreadsheetWorksheet, [{
    key: "_makeSingleUpdateRequest",
    value: function () {
      var _makeSingleUpdateRequest2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee(requestType, requestParams) {
        return regeneratorRuntime.wrap(function _callee$(_context) {
          while (1) {
            switch (_context.prev = _context.next) {
              case 0:
                return _context.abrupt("return", this._spreadsheet._makeSingleUpdateRequest(requestType, _objectSpread({}, requestParams)));

              case 1:
              case "end":
                return _context.stop();
            }
          }
        }, _callee, this);
      }));

      function _makeSingleUpdateRequest(_x, _x2) {
        return _makeSingleUpdateRequest2.apply(this, arguments);
      }

      return _makeSingleUpdateRequest;
    }()
  }, {
    key: "_ensureInfoLoaded",
    value: function _ensureInfoLoaded() {
      if (!this._rawProperties) {
        throw new Error('You must call `doc.loadInfo()` again before accessing this property');
      }
    }
  }, {
    key: "resetLocalCache",
    value: function resetLocalCache(dataOnly) {
      if (!dataOnly) this._rawProperties = null;
      this.headerValues = null;
      this._cells = [];
    }
  }, {
    key: "_fillCellData",
    value: function _fillCellData(dataRanges) {
      var _this = this;

      _.each(dataRanges, function (range) {
        var startRow = range.startRow || 0;
        var startColumn = range.startColumn || 0;
        var numRows = range.rowMetadata.length;
        var numColumns = range.columnMetadata.length; // update cell data for entire range

        for (var i = 0; i < numRows; i++) {
          var actualRow = startRow + i;

          for (var j = 0; j < numColumns; j++) {
            var actualColumn = startColumn + j; // if the row has not been initialized yet, do it

            if (!_this._cells[actualRow]) _this._cells[actualRow] = []; // see if the response includes some info for the cell

            var cellData = _.get(range, "rowData[".concat(i, "].values[").concat(j, "]")); // update the cell object or create it


            if (_this._cells[actualRow][actualColumn]) {
              _this._cells[actualRow][actualColumn]._updateRawData(cellData);
            } else {
              _this._cells[actualRow][actualColumn] = new GoogleSpreadsheetCell(_this, actualRow, actualColumn, cellData);
            }
          }
        } // update row metadata


        for (var _i = 0; _i < range.rowMetadata.length; _i++) {
          _this._rowMetadata[startRow + _i] = range.rowMetadata[_i];
        } // update column metadata


        for (var _i2 = 0; _i2 < range.columnMetadata.length; _i2++) {
          _this._columnMetadata[startColumn + _i2] = range.columnMetadata[_i2];
        }
      });
    } // PROPERTY GETTERS //////////////////////////////////////////////////////////////////////////////

  }, {
    key: "_getProp",
    value: function _getProp(param) {
      this._ensureInfoLoaded();

      return this._rawProperties[param];
    }
  }, {
    key: "_setProp",
    value: function _setProp(param, newVal) {
      // eslint-disable-line no-unused-vars
      throw new Error('Do not update directly - use `updateProperties()`');
    }
  }, {
    key: "getCellByA1",
    value: function getCellByA1(a1Address) {
      var split = a1Address.match(/([A-Z]+)([0-9]+)/);
      var columnIndex = letterToColumn(split[1]);
      var rowIndex = parseInt(split[2]);
      return this.getCell(rowIndex - 1, columnIndex - 1);
    }
  }, {
    key: "getCell",
    value: function getCell(rowIndex, columnIndex) {
      if (rowIndex < 0 || columnIndex < 0) throw new Error('Min coordinate is 0, 0');

      if (rowIndex >= this.rowCount || columnIndex >= this.columnCount) {
        throw new Error("Out of bounds, sheet is ".concat(this.rowCount, " by ").concat(this.columnCount));
      }

      if (!_.get(this._cells, "[".concat(rowIndex, "][").concat(columnIndex, "]"))) {
        throw new Error('This cell has not been loaded yet');
      }

      return this._cells[rowIndex][columnIndex];
    }
  }, {
    key: "loadCells",
    value: function () {
      var _loadCells = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee2(sheetFilters) {
        var _this2 = this;

        var filtersArray;
        return regeneratorRuntime.wrap(function _callee2$(_context2) {
          while (1) {
            switch (_context2.prev = _context2.next) {
              case 0:
                if (sheetFilters) {
                  _context2.next = 2;
                  break;
                }

                return _context2.abrupt("return", this._spreadsheet.loadCells(this.a1SheetName));

              case 2:
                filtersArray = _.isArray(sheetFilters) ? sheetFilters : [sheetFilters];
                filtersArray = _.map(filtersArray, function (filter) {
                  // add sheet name to A1 ranges
                  if (_.isString(filter)) {
                    if (filter.startsWith(_this2.a1SheetName)) return filter;
                    return "".concat(_this2.a1SheetName, "!").concat(filter);
                  }

                  if (_.isObject(filter)) {
                    // TODO: detect and support DeveloperMetadata filters
                    if (!filter.sheetId) {
                      return _objectSpread({
                        sheetId: _this2.sheetId
                      }, filter);
                    }

                    if (filter.sheetId !== _this2.sheetId) {
                      throw new Error('Leave sheet ID blank or set to matching ID of this sheet');
                    } else {
                      return filter;
                    }
                  } else {
                    throw new Error('Each filter must be a A1 range string or gridrange object');
                  }
                });
                return _context2.abrupt("return", this._spreadsheet.loadCells(filtersArray));

              case 5:
              case "end":
                return _context2.stop();
            }
          }
        }, _callee2, this);
      }));

      function loadCells(_x3) {
        return _loadCells.apply(this, arguments);
      }

      return loadCells;
    }()
  }, {
    key: "saveUpdatedCells",
    value: function () {
      var _saveUpdatedCells = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee3() {
        var cellsToSave;
        return regeneratorRuntime.wrap(function _callee3$(_context3) {
          while (1) {
            switch (_context3.prev = _context3.next) {
              case 0:
                cellsToSave = _.filter(_.flatten(this._cells), {
                  _isDirty: true
                });

                if (!cellsToSave.length) {
                  _context3.next = 4;
                  break;
                }

                _context3.next = 4;
                return this.saveCells(cellsToSave);

              case 4:
              case "end":
                return _context3.stop();
            }
          }
        }, _callee3, this);
      }));

      function saveUpdatedCells() {
        return _saveUpdatedCells.apply(this, arguments);
      }

      return saveUpdatedCells;
    }()
  }, {
    key: "saveCells",
    value: function () {
      var _saveCells = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee4(cellsToUpdate) {
        var _this3 = this;

        var requests, responseRanges;
        return regeneratorRuntime.wrap(function _callee4$(_context4) {
          while (1) {
            switch (_context4.prev = _context4.next) {
              case 0:
                // we send an individual "updateCells" request for each cell
                // because the fields that are udpated for each group are the same
                // and we dont want to accidentally overwrite something
                requests = _.map(cellsToUpdate, function (cell) {
                  return cell._getUpdateRequest();
                });
                responseRanges = _.map(cellsToUpdate, function (c) {
                  return "".concat(_this3.a1SheetName, "!").concat(c.a1Address);
                }); // if nothing is being updated the request returned is just `null`
                // so we make sure at least 1 request is valid - otherwise google throws a 400

                if (_.compact(requests).length) {
                  _context4.next = 4;
                  break;
                }

                throw new Error('At least one cell must have something to update');

              case 4:
                _context4.next = 6;
                return this._spreadsheet._makeBatchUpdateRequest(requests, responseRanges);

              case 6:
              case "end":
                return _context4.stop();
            }
          }
        }, _callee4, this);
      }));

      function saveCells(_x4) {
        return _saveCells.apply(this, arguments);
      }

      return saveCells;
    }() // SAVING THIS FOR FUTURE USE
    // puts the cells that need updating into batches
    // async updateCellsByBatches() {
    //   // saving this code, but it's problematic because each group must have the same update fields
    //   const cellsByRow = _.groupBy(cellsToUpdate, 'rowIndex');
    //   const groupsToSave = [];
    //   _.each(cellsByRow, (cells, rowIndex) => {
    //     let cellGroup = [];
    //     _.each(cells, (c) => {
    //       if (!cellGroup.length) {
    //         cellGroup.push(c);
    //       } else if (
    //         cellGroup[cellGroup.length - 1].columnIndex ===
    //         c.columnIndex - 1
    //       ) {
    //         cellGroup.push(c);
    //       } else {
    //         groupsToSave.push(cellGroup);
    //         cellGroup = [];
    //       }
    //     });
    //     groupsToSave.push(cellGroup);
    //   });
    //   const requests = _.map(groupsToSave, (cellGroup) => ({
    //     updateCells: {
    //       rows: [
    //         {
    //           values: _.map(cellGroup, (cell) => ({
    //             ...cell._draftData.value && {
    //               userEnteredValue: { [cell._draftData.valueType]: cell._draftData.value },
    //             },
    //             ...cell._draftData.note !== undefined && {
    //               note: cell._draftData.note ,
    //             },
    //             ...cell._draftData.userEnteredFormat && {
    //               userEnteredValue: cell._draftData.userEnteredFormat,
    //             },
    //           })),
    //         },
    //       ],
    //       fields: 'userEnteredValue,note,userEnteredFormat',
    //       start: {
    //         sheetId: this.sheetId,
    //         rowIndex: cellGroup[0].rowIndex,
    //         columnIndex: cellGroup[0].columnIndex,
    //       },
    //     },
    //   }));
    //   const responseRanges = _.map(groupsToSave, (cellGroup) => {
    //     let a1Range = cellGroup[0].a1Address;
    //     if (cellGroup.length > 1)
    //       a1Range += `:${cellGroup[cellGroup.length - 1].a1Address}`;
    //     return `${cellGroup[0]._sheet.a1SheetName}!${a1Range}`;
    //   });
    // }
    // ROW BASED FUNCTIONS ///////////////////////////////////////////////////////////////////////////

  }, {
    key: "loadHeaderRow",
    value: function () {
      var _loadHeaderRow = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee5() {
        var rows;
        return regeneratorRuntime.wrap(function _callee5$(_context5) {
          while (1) {
            switch (_context5.prev = _context5.next) {
              case 0:
                _context5.next = 2;
                return this.getCellsInRange("A1:".concat(this.lastColumnLetter, "1"));

              case 2:
                rows = _context5.sent;

                if (rows) {
                  _context5.next = 5;
                  break;
                }

                throw new Error('No values in the header row - fill the first row with header values before trying to interact with rows');

              case 5:
                this.headerValues = _.map(rows[0], function (header) {
                  return header.trim();
                });

                if (_.compact(this.headerValues).length) {
                  _context5.next = 8;
                  break;
                }

                throw new Error('All your header cells are blank - fill the first row with header values before trying to interact with rows');

              case 8:
                checkForDuplicateHeaders(this.headerValues);

              case 9:
              case "end":
                return _context5.stop();
            }
          }
        }, _callee5, this);
      }));

      function loadHeaderRow() {
        return _loadHeaderRow.apply(this, arguments);
      }

      return loadHeaderRow;
    }()
  }, {
    key: "setHeaderRow",
    value: function () {
      var _setHeaderRow = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee6(headerValues) {
        var trimmedHeaderValues, response;
        return regeneratorRuntime.wrap(function _callee6$(_context6) {
          while (1) {
            switch (_context6.prev = _context6.next) {
              case 0:
                if (headerValues) {
                  _context6.next = 2;
                  break;
                }

                return _context6.abrupt("return");

              case 2:
                if (!(headerValues.length > this.columnCount)) {
                  _context6.next = 4;
                  break;
                }

                throw new Error("Sheet is not large enough to fit ".concat(headerValues.length, " columns. Resize the sheet first."));

              case 4:
                trimmedHeaderValues = _.map(headerValues, function (h) {
                  return h.trim();
                });
                checkForDuplicateHeaders(trimmedHeaderValues);

                if (_.compact(trimmedHeaderValues).length) {
                  _context6.next = 8;
                  break;
                }

                throw new Error('All your header cells are blank -');

              case 8:
                _context6.next = 10;
                return this._spreadsheet.axios.request({
                  method: 'put',
                  url: "/values/".concat(this.encodedA1SheetName, "!1:1"),
                  params: {
                    valueInputOption: 'USER_ENTERED',
                    // other option is RAW
                    includeValuesInResponse: true
                  },
                  data: {
                    range: "".concat(this.a1SheetName, "!1:1"),
                    majorDimension: 'ROWS',
                    values: [[].concat(_toConsumableArray(trimmedHeaderValues), _toConsumableArray(_.times(this.columnCount - trimmedHeaderValues.length, function () {
                      return '';
                    })))]
                  }
                });

              case 10:
                response = _context6.sent;
                this.headerValues = response.data.updatedData.values[0];

              case 12:
              case "end":
                return _context6.stop();
            }
          }
        }, _callee6, this);
      }));

      function setHeaderRow(_x5) {
        return _setHeaderRow.apply(this, arguments);
      }

      return setHeaderRow;
    }()
  }, {
    key: "addRows",
    value: function () {
      var _addRows = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee7(rows) {
        var _this4 = this;

        var options,
            rowsAsArrays,
            response,
            updatedRange,
            rowNumber,
            _args7 = arguments;
        return regeneratorRuntime.wrap(function _callee7$(_context7) {
          while (1) {
            switch (_context7.prev = _context7.next) {
              case 0:
                options = _args7.length > 1 && _args7[1] !== undefined ? _args7[1] : {};

                if (!this.title.includes(':')) {
                  _context7.next = 3;
                  break;
                }

                throw new Error('Please remove the ":" from your sheet title. There is a bug with the google API which breaks appending rows if any colons are in the sheet title.');

              case 3:
                if (_.isArray(rows)) {
                  _context7.next = 5;
                  break;
                }

                throw new Error('You must pass in an array of row values to append');

              case 5:
                if (this.headerValues) {
                  _context7.next = 8;
                  break;
                }

                _context7.next = 8;
                return this.loadHeaderRow();

              case 8:
                // convert each row into an array of cell values rather than the key/value object
                rowsAsArrays = [];

                _.each(rows, function (row) {
                  var rowAsArray;

                  if (_.isArray(row)) {
                    rowAsArray = row;
                  } else if (_.isObject(row)) {
                    rowAsArray = [];

                    for (var i = 0; i < _this4.headerValues.length; i++) {
                      var propName = _this4.headerValues[i];
                      rowAsArray[i] = row[propName];
                    }
                  } else {
                    throw new Error('Each row must be an object or an array');
                  }

                  rowsAsArrays.push(rowAsArray);
                });

                _context7.next = 12;
                return this._spreadsheet.axios.request({
                  method: 'post',
                  url: "/values/".concat(this.encodedA1SheetName, "!A1:append"),
                  params: {
                    valueInputOption: options.raw ? 'RAW' : 'USER_ENTERED',
                    insertDataOption: options.insert ? 'INSERT_ROWS' : 'OVERWRITE',
                    includeValuesInResponse: true
                  },
                  data: {
                    values: rowsAsArrays
                  }
                });

              case 12:
                response = _context7.sent;
                // extract the new row number from the A1-notation data range in the response
                // ex: in "'Sheet8!A2:C2" -- we want the `2`
                updatedRange = response.data.updates.updatedRange;
                rowNumber = updatedRange.match(/![A-Z]+([0-9]+):?/)[1];
                rowNumber = parseInt(rowNumber); // if new rows were added, we need update sheet.rowRount

                if (options.insert) {
                  this._rawProperties.gridProperties.rowCount += rows.length;
                } else if (rowNumber + rows.length > this.rowCount) {
                  // have to subtract 1 since one row was inserted at rowNumber
                  this._rawProperties.gridProperties.rowCount = rowNumber + rows.length - 1;
                }

                return _context7.abrupt("return", _.map(response.data.updates.updatedData.values, function (rowValues) {
                  var row = new GoogleSpreadsheetRow(_this4, rowNumber++, rowValues);
                  return row;
                }));

              case 18:
              case "end":
                return _context7.stop();
            }
          }
        }, _callee7, this);
      }));

      function addRows(_x6) {
        return _addRows.apply(this, arguments);
      }

      return addRows;
    }()
  }, {
    key: "addRow",
    value: function () {
      var _addRow = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee8(rowValues, options) {
        var rows;
        return regeneratorRuntime.wrap(function _callee8$(_context8) {
          while (1) {
            switch (_context8.prev = _context8.next) {
              case 0:
                _context8.next = 2;
                return this.addRows([rowValues], options);

              case 2:
                rows = _context8.sent;
                return _context8.abrupt("return", rows[0]);

              case 4:
              case "end":
                return _context8.stop();
            }
          }
        }, _callee8, this);
      }));

      function addRow(_x7, _x8) {
        return _addRow.apply(this, arguments);
      }

      return addRow;
    }()
  }, {
    key: "getRows",
    value: function () {
      var _getRows = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee9() {
        var options,
            firstRow,
            lastRow,
            lastColumn,
            rawRows,
            rows,
            rowNum,
            i,
            _args9 = arguments;
        return regeneratorRuntime.wrap(function _callee9$(_context9) {
          while (1) {
            switch (_context9.prev = _context9.next) {
              case 0:
                options = _args9.length > 0 && _args9[0] !== undefined ? _args9[0] : {};
                // https://developers.google.com/sheets/api/guides/migration
                // v4 API does not have equivalents for the row-order query parameters provided
                // Reverse-order is trivial; simply process the returned values array in reverse order.
                // Order by column is not supported for reads, but it is possible to sort the data then read
                // v4 API does not currently have a direct equivalent for the Sheets API v3 structured queries
                // However, you can retrieve the relevant data and sort through it as needed in your application
                // options
                // - offset
                // - limit
                options.offset = options.offset || 0;
                options.limit = options.limit || this.rowCount - 1;

                if (this.headerValues) {
                  _context9.next = 6;
                  break;
                }

                _context9.next = 6;
                return this.loadHeaderRow();

              case 6:
                firstRow = 2 + options.offset; // skip first row AND not zero indexed

                lastRow = firstRow + options.limit - 1; // inclusive so we subtract 1

                lastColumn = columnToLetter(this.headerValues.length);
                _context9.next = 11;
                return this.getCellsInRange("A".concat(firstRow, ":").concat(lastColumn).concat(lastRow));

              case 11:
                rawRows = _context9.sent;

                if (rawRows) {
                  _context9.next = 14;
                  break;
                }

                return _context9.abrupt("return", []);

              case 14:
                rows = [];
                rowNum = firstRow;

                for (i = 0; i < rawRows.length; i++) {
                  rows.push(new GoogleSpreadsheetRow(this, rowNum++, rawRows[i]));
                }

                return _context9.abrupt("return", rows);

              case 18:
              case "end":
                return _context9.stop();
            }
          }
        }, _callee9, this);
      }));

      function getRows() {
        return _getRows.apply(this, arguments);
      }

      return getRows;
    }() // BASIC PROPS ///////////////////////////////////////////////////////////////////////////////////

  }, {
    key: "updateProperties",
    value: function () {
      var _updateProperties = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee10(properties) {
        return regeneratorRuntime.wrap(function _callee10$(_context10) {
          while (1) {
            switch (_context10.prev = _context10.next) {
              case 0:
                return _context10.abrupt("return", this._makeSingleUpdateRequest('updateSheetProperties', {
                  properties: _objectSpread({
                    sheetId: this.sheetId
                  }, properties),
                  fields: getFieldMask(properties)
                }));

              case 1:
              case "end":
                return _context10.stop();
            }
          }
        }, _callee10, this);
      }));

      function updateProperties(_x9) {
        return _updateProperties.apply(this, arguments);
      }

      return updateProperties;
    }()
  }, {
    key: "updateGridProperties",
    value: function () {
      var _updateGridProperties = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee11(gridProperties) {
        return regeneratorRuntime.wrap(function _callee11$(_context11) {
          while (1) {
            switch (_context11.prev = _context11.next) {
              case 0:
                return _context11.abrupt("return", this.updateProperties({
                  gridProperties: gridProperties
                }));

              case 1:
              case "end":
                return _context11.stop();
            }
          }
        }, _callee11, this);
      }));

      function updateGridProperties(_x10) {
        return _updateGridProperties.apply(this, arguments);
      }

      return updateGridProperties;
    }() // just a shortcut because resize makes more sense to change rowCount / columnCount

  }, {
    key: "resize",
    value: function () {
      var _resize = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee12(gridProperties) {
        return regeneratorRuntime.wrap(function _callee12$(_context12) {
          while (1) {
            switch (_context12.prev = _context12.next) {
              case 0:
                return _context12.abrupt("return", this.updateGridProperties(gridProperties));

              case 1:
              case "end":
                return _context12.stop();
            }
          }
        }, _callee12, this);
      }));

      function resize(_x11) {
        return _resize.apply(this, arguments);
      }

      return resize;
    }()
  }, {
    key: "updateDimensionProperties",
    value: function () {
      var _updateDimensionProperties = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee13(columnsOrRows, properties, bounds) {
        return regeneratorRuntime.wrap(function _callee13$(_context13) {
          while (1) {
            switch (_context13.prev = _context13.next) {
              case 0:
                return _context13.abrupt("return", this._makeSingleUpdateRequest('updateDimensionProperties', {
                  range: _objectSpread({
                    sheetId: this.sheetId,
                    dimension: columnsOrRows
                  }, bounds && {
                    startIndex: bounds.startIndex,
                    endIndex: bounds.endIndex
                  }),
                  properties: properties,
                  fields: getFieldMask(properties)
                }));

              case 1:
              case "end":
                return _context13.stop();
            }
          }
        }, _callee13, this);
      }));

      function updateDimensionProperties(_x12, _x13, _x14) {
        return _updateDimensionProperties.apply(this, arguments);
      }

      return updateDimensionProperties;
    }() // OTHER /////////////////////////////////////////////////////////////////////////////////////////
    // this uses the "values" getter and does not give all the info about the cell contents
    // it is used internally when loading header cells

  }, {
    key: "getCellsInRange",
    value: function () {
      var _getCellsInRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee14(a1Range, options) {
        var response;
        return regeneratorRuntime.wrap(function _callee14$(_context14) {
          while (1) {
            switch (_context14.prev = _context14.next) {
              case 0:
                _context14.next = 2;
                return this._spreadsheet.axios.get("/values/".concat(this.encodedA1SheetName, "!").concat(a1Range), {
                  params: options
                });

              case 2:
                response = _context14.sent;
                return _context14.abrupt("return", response.data.values);

              case 4:
              case "end":
                return _context14.stop();
            }
          }
        }, _callee14, this);
      }));

      function getCellsInRange(_x15, _x16) {
        return _getCellsInRange.apply(this, arguments);
      }

      return getCellsInRange;
    }()
  }, {
    key: "updateNamedRange",
    value: function () {
      var _updateNamedRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee15() {
        return regeneratorRuntime.wrap(function _callee15$(_context15) {
          while (1) {
            switch (_context15.prev = _context15.next) {
              case 0:
              case "end":
                return _context15.stop();
            }
          }
        }, _callee15);
      }));

      function updateNamedRange() {
        return _updateNamedRange.apply(this, arguments);
      }

      return updateNamedRange;
    }()
  }, {
    key: "addNamedRange",
    value: function () {
      var _addNamedRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee16() {
        return regeneratorRuntime.wrap(function _callee16$(_context16) {
          while (1) {
            switch (_context16.prev = _context16.next) {
              case 0:
              case "end":
                return _context16.stop();
            }
          }
        }, _callee16);
      }));

      function addNamedRange() {
        return _addNamedRange.apply(this, arguments);
      }

      return addNamedRange;
    }()
  }, {
    key: "deleteNamedRange",
    value: function () {
      var _deleteNamedRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee17() {
        return regeneratorRuntime.wrap(function _callee17$(_context17) {
          while (1) {
            switch (_context17.prev = _context17.next) {
              case 0:
              case "end":
                return _context17.stop();
            }
          }
        }, _callee17);
      }));

      function deleteNamedRange() {
        return _deleteNamedRange.apply(this, arguments);
      }

      return deleteNamedRange;
    }()
  }, {
    key: "repeatCell",
    value: function () {
      var _repeatCell = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee18() {
        return regeneratorRuntime.wrap(function _callee18$(_context18) {
          while (1) {
            switch (_context18.prev = _context18.next) {
              case 0:
              case "end":
                return _context18.stop();
            }
          }
        }, _callee18);
      }));

      function repeatCell() {
        return _repeatCell.apply(this, arguments);
      }

      return repeatCell;
    }()
  }, {
    key: "autoFill",
    value: function () {
      var _autoFill = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee19() {
        return regeneratorRuntime.wrap(function _callee19$(_context19) {
          while (1) {
            switch (_context19.prev = _context19.next) {
              case 0:
              case "end":
                return _context19.stop();
            }
          }
        }, _callee19);
      }));

      function autoFill() {
        return _autoFill.apply(this, arguments);
      }

      return autoFill;
    }()
  }, {
    key: "cutPaste",
    value: function () {
      var _cutPaste = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee20() {
        return regeneratorRuntime.wrap(function _callee20$(_context20) {
          while (1) {
            switch (_context20.prev = _context20.next) {
              case 0:
              case "end":
                return _context20.stop();
            }
          }
        }, _callee20);
      }));

      function cutPaste() {
        return _cutPaste.apply(this, arguments);
      }

      return cutPaste;
    }()
  }, {
    key: "copyPaste",
    value: function () {
      var _copyPaste = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee21() {
        return regeneratorRuntime.wrap(function _callee21$(_context21) {
          while (1) {
            switch (_context21.prev = _context21.next) {
              case 0:
              case "end":
                return _context21.stop();
            }
          }
        }, _callee21);
      }));

      function copyPaste() {
        return _copyPaste.apply(this, arguments);
      }

      return copyPaste;
    }()
  }, {
    key: "mergeCells",
    value: function () {
      var _mergeCells = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee22(range) {
        var mergeType,
            _args22 = arguments;
        return regeneratorRuntime.wrap(function _callee22$(_context22) {
          while (1) {
            switch (_context22.prev = _context22.next) {
              case 0:
                mergeType = _args22.length > 1 && _args22[1] !== undefined ? _args22[1] : 'MERGE_ALL';

                if (!(range.sheetId && range.sheetId !== this.sheetId)) {
                  _context22.next = 3;
                  break;
                }

                throw new Error('Leave sheet ID blank or set to matching ID of this sheet');

              case 3:
                _context22.next = 5;
                return this._makeSingleUpdateRequest('mergeCells', {
                  mergeType: mergeType,
                  range: _objectSpread(_objectSpread({}, range), {}, {
                    sheetId: this.sheetId
                  })
                });

              case 5:
              case "end":
                return _context22.stop();
            }
          }
        }, _callee22, this);
      }));

      function mergeCells(_x17) {
        return _mergeCells.apply(this, arguments);
      }

      return mergeCells;
    }()
  }, {
    key: "unmergeCells",
    value: function () {
      var _unmergeCells = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee23(range) {
        return regeneratorRuntime.wrap(function _callee23$(_context23) {
          while (1) {
            switch (_context23.prev = _context23.next) {
              case 0:
                if (!(range.sheetId && range.sheetId !== this.sheetId)) {
                  _context23.next = 2;
                  break;
                }

                throw new Error('Leave sheet ID blank or set to matching ID of this sheet');

              case 2:
                _context23.next = 4;
                return this._makeSingleUpdateRequest('unmergeCells', {
                  range: _objectSpread(_objectSpread({}, range), {}, {
                    sheetId: this.sheetId
                  })
                });

              case 4:
              case "end":
                return _context23.stop();
            }
          }
        }, _callee23, this);
      }));

      function unmergeCells(_x18) {
        return _unmergeCells.apply(this, arguments);
      }

      return unmergeCells;
    }()
  }, {
    key: "updateBorders",
    value: function () {
      var _updateBorders = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee24() {
        return regeneratorRuntime.wrap(function _callee24$(_context24) {
          while (1) {
            switch (_context24.prev = _context24.next) {
              case 0:
              case "end":
                return _context24.stop();
            }
          }
        }, _callee24);
      }));

      function updateBorders() {
        return _updateBorders.apply(this, arguments);
      }

      return updateBorders;
    }()
  }, {
    key: "addFilterView",
    value: function () {
      var _addFilterView = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee25() {
        return regeneratorRuntime.wrap(function _callee25$(_context25) {
          while (1) {
            switch (_context25.prev = _context25.next) {
              case 0:
              case "end":
                return _context25.stop();
            }
          }
        }, _callee25);
      }));

      function addFilterView() {
        return _addFilterView.apply(this, arguments);
      }

      return addFilterView;
    }()
  }, {
    key: "appendCells",
    value: function () {
      var _appendCells = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee26() {
        return regeneratorRuntime.wrap(function _callee26$(_context26) {
          while (1) {
            switch (_context26.prev = _context26.next) {
              case 0:
              case "end":
                return _context26.stop();
            }
          }
        }, _callee26);
      }));

      function appendCells() {
        return _appendCells.apply(this, arguments);
      }

      return appendCells;
    }()
  }, {
    key: "clearBasicFilter",
    value: function () {
      var _clearBasicFilter = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee27() {
        return regeneratorRuntime.wrap(function _callee27$(_context27) {
          while (1) {
            switch (_context27.prev = _context27.next) {
              case 0:
              case "end":
                return _context27.stop();
            }
          }
        }, _callee27);
      }));

      function clearBasicFilter() {
        return _clearBasicFilter.apply(this, arguments);
      }

      return clearBasicFilter;
    }()
  }, {
    key: "deleteDimension",
    value: function () {
      var _deleteDimension = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee28() {
        return regeneratorRuntime.wrap(function _callee28$(_context28) {
          while (1) {
            switch (_context28.prev = _context28.next) {
              case 0:
              case "end":
                return _context28.stop();
            }
          }
        }, _callee28);
      }));

      function deleteDimension() {
        return _deleteDimension.apply(this, arguments);
      }

      return deleteDimension;
    }()
  }, {
    key: "deleteEmbeddedObject",
    value: function () {
      var _deleteEmbeddedObject = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee29() {
        return regeneratorRuntime.wrap(function _callee29$(_context29) {
          while (1) {
            switch (_context29.prev = _context29.next) {
              case 0:
              case "end":
                return _context29.stop();
            }
          }
        }, _callee29);
      }));

      function deleteEmbeddedObject() {
        return _deleteEmbeddedObject.apply(this, arguments);
      }

      return deleteEmbeddedObject;
    }()
  }, {
    key: "deleteFilterView",
    value: function () {
      var _deleteFilterView = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee30() {
        return regeneratorRuntime.wrap(function _callee30$(_context30) {
          while (1) {
            switch (_context30.prev = _context30.next) {
              case 0:
              case "end":
                return _context30.stop();
            }
          }
        }, _callee30);
      }));

      function deleteFilterView() {
        return _deleteFilterView.apply(this, arguments);
      }

      return deleteFilterView;
    }()
  }, {
    key: "duplicateFilterView",
    value: function () {
      var _duplicateFilterView = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee31() {
        return regeneratorRuntime.wrap(function _callee31$(_context31) {
          while (1) {
            switch (_context31.prev = _context31.next) {
              case 0:
              case "end":
                return _context31.stop();
            }
          }
        }, _callee31);
      }));

      function duplicateFilterView() {
        return _duplicateFilterView.apply(this, arguments);
      }

      return duplicateFilterView;
    }()
  }, {
    key: "duplicateSheet",
    value: function () {
      var _duplicateSheet = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee32() {
        return regeneratorRuntime.wrap(function _callee32$(_context32) {
          while (1) {
            switch (_context32.prev = _context32.next) {
              case 0:
              case "end":
                return _context32.stop();
            }
          }
        }, _callee32);
      }));

      function duplicateSheet() {
        return _duplicateSheet.apply(this, arguments);
      }

      return duplicateSheet;
    }()
  }, {
    key: "findReplace",
    value: function () {
      var _findReplace = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee33() {
        return regeneratorRuntime.wrap(function _callee33$(_context33) {
          while (1) {
            switch (_context33.prev = _context33.next) {
              case 0:
              case "end":
                return _context33.stop();
            }
          }
        }, _callee33);
      }));

      function findReplace() {
        return _findReplace.apply(this, arguments);
      }

      return findReplace;
    }()
  }, {
    key: "insertDimension",
    value: function () {
      var _insertDimension = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee34() {
        return regeneratorRuntime.wrap(function _callee34$(_context34) {
          while (1) {
            switch (_context34.prev = _context34.next) {
              case 0:
              case "end":
                return _context34.stop();
            }
          }
        }, _callee34);
      }));

      function insertDimension() {
        return _insertDimension.apply(this, arguments);
      }

      return insertDimension;
    }()
  }, {
    key: "insertRange",
    value: function () {
      var _insertRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee35() {
        return regeneratorRuntime.wrap(function _callee35$(_context35) {
          while (1) {
            switch (_context35.prev = _context35.next) {
              case 0:
              case "end":
                return _context35.stop();
            }
          }
        }, _callee35);
      }));

      function insertRange() {
        return _insertRange.apply(this, arguments);
      }

      return insertRange;
    }()
  }, {
    key: "moveDimension",
    value: function () {
      var _moveDimension = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee36() {
        return regeneratorRuntime.wrap(function _callee36$(_context36) {
          while (1) {
            switch (_context36.prev = _context36.next) {
              case 0:
              case "end":
                return _context36.stop();
            }
          }
        }, _callee36);
      }));

      function moveDimension() {
        return _moveDimension.apply(this, arguments);
      }

      return moveDimension;
    }()
  }, {
    key: "updateEmbeddedObjectPosition",
    value: function () {
      var _updateEmbeddedObjectPosition = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee37() {
        return regeneratorRuntime.wrap(function _callee37$(_context37) {
          while (1) {
            switch (_context37.prev = _context37.next) {
              case 0:
              case "end":
                return _context37.stop();
            }
          }
        }, _callee37);
      }));

      function updateEmbeddedObjectPosition() {
        return _updateEmbeddedObjectPosition.apply(this, arguments);
      }

      return updateEmbeddedObjectPosition;
    }()
  }, {
    key: "pasteData",
    value: function () {
      var _pasteData = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee38() {
        return regeneratorRuntime.wrap(function _callee38$(_context38) {
          while (1) {
            switch (_context38.prev = _context38.next) {
              case 0:
              case "end":
                return _context38.stop();
            }
          }
        }, _callee38);
      }));

      function pasteData() {
        return _pasteData.apply(this, arguments);
      }

      return pasteData;
    }()
  }, {
    key: "textToColumns",
    value: function () {
      var _textToColumns = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee39() {
        return regeneratorRuntime.wrap(function _callee39$(_context39) {
          while (1) {
            switch (_context39.prev = _context39.next) {
              case 0:
              case "end":
                return _context39.stop();
            }
          }
        }, _callee39);
      }));

      function textToColumns() {
        return _textToColumns.apply(this, arguments);
      }

      return textToColumns;
    }()
  }, {
    key: "updateFilterView",
    value: function () {
      var _updateFilterView = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee40() {
        return regeneratorRuntime.wrap(function _callee40$(_context40) {
          while (1) {
            switch (_context40.prev = _context40.next) {
              case 0:
              case "end":
                return _context40.stop();
            }
          }
        }, _callee40);
      }));

      function updateFilterView() {
        return _updateFilterView.apply(this, arguments);
      }

      return updateFilterView;
    }()
  }, {
    key: "deleteRange",
    value: function () {
      var _deleteRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee41() {
        return regeneratorRuntime.wrap(function _callee41$(_context41) {
          while (1) {
            switch (_context41.prev = _context41.next) {
              case 0:
              case "end":
                return _context41.stop();
            }
          }
        }, _callee41);
      }));

      function deleteRange() {
        return _deleteRange.apply(this, arguments);
      }

      return deleteRange;
    }()
  }, {
    key: "appendDimension",
    value: function () {
      var _appendDimension = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee42() {
        return regeneratorRuntime.wrap(function _callee42$(_context42) {
          while (1) {
            switch (_context42.prev = _context42.next) {
              case 0:
              case "end":
                return _context42.stop();
            }
          }
        }, _callee42);
      }));

      function appendDimension() {
        return _appendDimension.apply(this, arguments);
      }

      return appendDimension;
    }()
  }, {
    key: "addConditionalFormatRule",
    value: function () {
      var _addConditionalFormatRule = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee43() {
        return regeneratorRuntime.wrap(function _callee43$(_context43) {
          while (1) {
            switch (_context43.prev = _context43.next) {
              case 0:
              case "end":
                return _context43.stop();
            }
          }
        }, _callee43);
      }));

      function addConditionalFormatRule() {
        return _addConditionalFormatRule.apply(this, arguments);
      }

      return addConditionalFormatRule;
    }()
  }, {
    key: "updateConditionalFormatRule",
    value: function () {
      var _updateConditionalFormatRule = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee44() {
        return regeneratorRuntime.wrap(function _callee44$(_context44) {
          while (1) {
            switch (_context44.prev = _context44.next) {
              case 0:
              case "end":
                return _context44.stop();
            }
          }
        }, _callee44);
      }));

      function updateConditionalFormatRule() {
        return _updateConditionalFormatRule.apply(this, arguments);
      }

      return updateConditionalFormatRule;
    }()
  }, {
    key: "deleteConditionalFormatRule",
    value: function () {
      var _deleteConditionalFormatRule = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee45() {
        return regeneratorRuntime.wrap(function _callee45$(_context45) {
          while (1) {
            switch (_context45.prev = _context45.next) {
              case 0:
              case "end":
                return _context45.stop();
            }
          }
        }, _callee45);
      }));

      function deleteConditionalFormatRule() {
        return _deleteConditionalFormatRule.apply(this, arguments);
      }

      return deleteConditionalFormatRule;
    }()
  }, {
    key: "sortRange",
    value: function () {
      var _sortRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee46() {
        return regeneratorRuntime.wrap(function _callee46$(_context46) {
          while (1) {
            switch (_context46.prev = _context46.next) {
              case 0:
              case "end":
                return _context46.stop();
            }
          }
        }, _callee46);
      }));

      function sortRange() {
        return _sortRange.apply(this, arguments);
      }

      return sortRange;
    }()
  }, {
    key: "setDataValidation",
    value: function () {
      var _setDataValidation = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee47() {
        return regeneratorRuntime.wrap(function _callee47$(_context47) {
          while (1) {
            switch (_context47.prev = _context47.next) {
              case 0:
              case "end":
                return _context47.stop();
            }
          }
        }, _callee47);
      }));

      function setDataValidation() {
        return _setDataValidation.apply(this, arguments);
      }

      return setDataValidation;
    }()
  }, {
    key: "setBasicFilter",
    value: function () {
      var _setBasicFilter = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee48() {
        return regeneratorRuntime.wrap(function _callee48$(_context48) {
          while (1) {
            switch (_context48.prev = _context48.next) {
              case 0:
              case "end":
                return _context48.stop();
            }
          }
        }, _callee48);
      }));

      function setBasicFilter() {
        return _setBasicFilter.apply(this, arguments);
      }

      return setBasicFilter;
    }()
  }, {
    key: "addProtectedRange",
    value: function () {
      var _addProtectedRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee49() {
        return regeneratorRuntime.wrap(function _callee49$(_context49) {
          while (1) {
            switch (_context49.prev = _context49.next) {
              case 0:
              case "end":
                return _context49.stop();
            }
          }
        }, _callee49);
      }));

      function addProtectedRange() {
        return _addProtectedRange.apply(this, arguments);
      }

      return addProtectedRange;
    }()
  }, {
    key: "updateProtectedRange",
    value: function () {
      var _updateProtectedRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee50() {
        return regeneratorRuntime.wrap(function _callee50$(_context50) {
          while (1) {
            switch (_context50.prev = _context50.next) {
              case 0:
              case "end":
                return _context50.stop();
            }
          }
        }, _callee50);
      }));

      function updateProtectedRange() {
        return _updateProtectedRange.apply(this, arguments);
      }

      return updateProtectedRange;
    }()
  }, {
    key: "deleteProtectedRange",
    value: function () {
      var _deleteProtectedRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee51() {
        return regeneratorRuntime.wrap(function _callee51$(_context51) {
          while (1) {
            switch (_context51.prev = _context51.next) {
              case 0:
              case "end":
                return _context51.stop();
            }
          }
        }, _callee51);
      }));

      function deleteProtectedRange() {
        return _deleteProtectedRange.apply(this, arguments);
      }

      return deleteProtectedRange;
    }()
  }, {
    key: "autoResizeDimensions",
    value: function () {
      var _autoResizeDimensions = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee52() {
        return regeneratorRuntime.wrap(function _callee52$(_context52) {
          while (1) {
            switch (_context52.prev = _context52.next) {
              case 0:
              case "end":
                return _context52.stop();
            }
          }
        }, _callee52);
      }));

      function autoResizeDimensions() {
        return _autoResizeDimensions.apply(this, arguments);
      }

      return autoResizeDimensions;
    }()
  }, {
    key: "addChart",
    value: function () {
      var _addChart = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee53() {
        return regeneratorRuntime.wrap(function _callee53$(_context53) {
          while (1) {
            switch (_context53.prev = _context53.next) {
              case 0:
              case "end":
                return _context53.stop();
            }
          }
        }, _callee53);
      }));

      function addChart() {
        return _addChart.apply(this, arguments);
      }

      return addChart;
    }()
  }, {
    key: "updateChartSpec",
    value: function () {
      var _updateChartSpec = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee54() {
        return regeneratorRuntime.wrap(function _callee54$(_context54) {
          while (1) {
            switch (_context54.prev = _context54.next) {
              case 0:
              case "end":
                return _context54.stop();
            }
          }
        }, _callee54);
      }));

      function updateChartSpec() {
        return _updateChartSpec.apply(this, arguments);
      }

      return updateChartSpec;
    }()
  }, {
    key: "updateBanding",
    value: function () {
      var _updateBanding = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee55() {
        return regeneratorRuntime.wrap(function _callee55$(_context55) {
          while (1) {
            switch (_context55.prev = _context55.next) {
              case 0:
              case "end":
                return _context55.stop();
            }
          }
        }, _callee55);
      }));

      function updateBanding() {
        return _updateBanding.apply(this, arguments);
      }

      return updateBanding;
    }()
  }, {
    key: "addBanding",
    value: function () {
      var _addBanding = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee56() {
        return regeneratorRuntime.wrap(function _callee56$(_context56) {
          while (1) {
            switch (_context56.prev = _context56.next) {
              case 0:
              case "end":
                return _context56.stop();
            }
          }
        }, _callee56);
      }));

      function addBanding() {
        return _addBanding.apply(this, arguments);
      }

      return addBanding;
    }()
  }, {
    key: "deleteBanding",
    value: function () {
      var _deleteBanding = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee57() {
        return regeneratorRuntime.wrap(function _callee57$(_context57) {
          while (1) {
            switch (_context57.prev = _context57.next) {
              case 0:
              case "end":
                return _context57.stop();
            }
          }
        }, _callee57);
      }));

      function deleteBanding() {
        return _deleteBanding.apply(this, arguments);
      }

      return deleteBanding;
    }()
  }, {
    key: "createDeveloperMetadata",
    value: function () {
      var _createDeveloperMetadata = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee58() {
        return regeneratorRuntime.wrap(function _callee58$(_context58) {
          while (1) {
            switch (_context58.prev = _context58.next) {
              case 0:
              case "end":
                return _context58.stop();
            }
          }
        }, _callee58);
      }));

      function createDeveloperMetadata() {
        return _createDeveloperMetadata.apply(this, arguments);
      }

      return createDeveloperMetadata;
    }()
  }, {
    key: "updateDeveloperMetadata",
    value: function () {
      var _updateDeveloperMetadata = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee59() {
        return regeneratorRuntime.wrap(function _callee59$(_context59) {
          while (1) {
            switch (_context59.prev = _context59.next) {
              case 0:
              case "end":
                return _context59.stop();
            }
          }
        }, _callee59);
      }));

      function updateDeveloperMetadata() {
        return _updateDeveloperMetadata.apply(this, arguments);
      }

      return updateDeveloperMetadata;
    }()
  }, {
    key: "deleteDeveloperMetadata",
    value: function () {
      var _deleteDeveloperMetadata = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee60() {
        return regeneratorRuntime.wrap(function _callee60$(_context60) {
          while (1) {
            switch (_context60.prev = _context60.next) {
              case 0:
              case "end":
                return _context60.stop();
            }
          }
        }, _callee60);
      }));

      function deleteDeveloperMetadata() {
        return _deleteDeveloperMetadata.apply(this, arguments);
      }

      return deleteDeveloperMetadata;
    }()
  }, {
    key: "randomizeRange",
    value: function () {
      var _randomizeRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee61() {
        return regeneratorRuntime.wrap(function _callee61$(_context61) {
          while (1) {
            switch (_context61.prev = _context61.next) {
              case 0:
              case "end":
                return _context61.stop();
            }
          }
        }, _callee61);
      }));

      function randomizeRange() {
        return _randomizeRange.apply(this, arguments);
      }

      return randomizeRange;
    }()
  }, {
    key: "addDimensionGroup",
    value: function () {
      var _addDimensionGroup = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee62() {
        return regeneratorRuntime.wrap(function _callee62$(_context62) {
          while (1) {
            switch (_context62.prev = _context62.next) {
              case 0:
              case "end":
                return _context62.stop();
            }
          }
        }, _callee62);
      }));

      function addDimensionGroup() {
        return _addDimensionGroup.apply(this, arguments);
      }

      return addDimensionGroup;
    }()
  }, {
    key: "deleteDimensionGroup",
    value: function () {
      var _deleteDimensionGroup = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee63() {
        return regeneratorRuntime.wrap(function _callee63$(_context63) {
          while (1) {
            switch (_context63.prev = _context63.next) {
              case 0:
              case "end":
                return _context63.stop();
            }
          }
        }, _callee63);
      }));

      function deleteDimensionGroup() {
        return _deleteDimensionGroup.apply(this, arguments);
      }

      return deleteDimensionGroup;
    }()
  }, {
    key: "updateDimensionGroup",
    value: function () {
      var _updateDimensionGroup = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee64() {
        return regeneratorRuntime.wrap(function _callee64$(_context64) {
          while (1) {
            switch (_context64.prev = _context64.next) {
              case 0:
              case "end":
                return _context64.stop();
            }
          }
        }, _callee64);
      }));

      function updateDimensionGroup() {
        return _updateDimensionGroup.apply(this, arguments);
      }

      return updateDimensionGroup;
    }()
  }, {
    key: "trimWhitespace",
    value: function () {
      var _trimWhitespace = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee65() {
        return regeneratorRuntime.wrap(function _callee65$(_context65) {
          while (1) {
            switch (_context65.prev = _context65.next) {
              case 0:
              case "end":
                return _context65.stop();
            }
          }
        }, _callee65);
      }));

      function trimWhitespace() {
        return _trimWhitespace.apply(this, arguments);
      }

      return trimWhitespace;
    }()
  }, {
    key: "deleteDuplicates",
    value: function () {
      var _deleteDuplicates = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee66() {
        return regeneratorRuntime.wrap(function _callee66$(_context66) {
          while (1) {
            switch (_context66.prev = _context66.next) {
              case 0:
              case "end":
                return _context66.stop();
            }
          }
        }, _callee66);
      }));

      function deleteDuplicates() {
        return _deleteDuplicates.apply(this, arguments);
      }

      return deleteDuplicates;
    }()
  }, {
    key: "addSlicer",
    value: function () {
      var _addSlicer = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee67() {
        return regeneratorRuntime.wrap(function _callee67$(_context67) {
          while (1) {
            switch (_context67.prev = _context67.next) {
              case 0:
              case "end":
                return _context67.stop();
            }
          }
        }, _callee67);
      }));

      function addSlicer() {
        return _addSlicer.apply(this, arguments);
      }

      return addSlicer;
    }()
  }, {
    key: "updateSlicerSpec",
    value: function () {
      var _updateSlicerSpec = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee68() {
        return regeneratorRuntime.wrap(function _callee68$(_context68) {
          while (1) {
            switch (_context68.prev = _context68.next) {
              case 0:
              case "end":
                return _context68.stop();
            }
          }
        }, _callee68);
      }));

      function updateSlicerSpec() {
        return _updateSlicerSpec.apply(this, arguments);
      }

      return updateSlicerSpec;
    }() // delete this worksheet

  }, {
    key: "delete",
    value: function () {
      var _delete2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee69() {
        return regeneratorRuntime.wrap(function _callee69$(_context69) {
          while (1) {
            switch (_context69.prev = _context69.next) {
              case 0:
                return _context69.abrupt("return", this._spreadsheet.deleteSheet(this.sheetId));

              case 1:
              case "end":
                return _context69.stop();
            }
          }
        }, _callee69, this);
      }));

      function _delete() {
        return _delete2.apply(this, arguments);
      }

      return _delete;
    }()
  }, {
    key: "del",
    value: function () {
      var _del = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee70() {
        return regeneratorRuntime.wrap(function _callee70$(_context70) {
          while (1) {
            switch (_context70.prev = _context70.next) {
              case 0:
                return _context70.abrupt("return", this["delete"]());

              case 1:
              case "end":
                return _context70.stop();
            }
          }
        }, _callee70, this);
      }));

      function del() {
        return _del.apply(this, arguments);
      }

      return del;
    }() // alias to mimic old interface
    // copies this worksheet into another document/spreadsheet

  }, {
    key: "copyToSpreadsheet",
    value: function () {
      var _copyToSpreadsheet = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee71(destinationSpreadsheetId) {
        return regeneratorRuntime.wrap(function _callee71$(_context71) {
          while (1) {
            switch (_context71.prev = _context71.next) {
              case 0:
                return _context71.abrupt("return", this._spreadsheet.axios.post("/sheets/".concat(this.sheetId, ":copyTo"), {
                  destinationSpreadsheetId: destinationSpreadsheetId
                }));

              case 1:
              case "end":
                return _context71.stop();
            }
          }
        }, _callee71, this);
      }));

      function copyToSpreadsheet(_x19) {
        return _copyToSpreadsheet.apply(this, arguments);
      }

      return copyToSpreadsheet;
    }()
  }, {
    key: "clear",
    value: function () {
      var _clear = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee72() {
        return regeneratorRuntime.wrap(function _callee72$(_context72) {
          while (1) {
            switch (_context72.prev = _context72.next) {
              case 0:
                _context72.next = 2;
                return this._spreadsheet.axios.post("/values/".concat(this.encodedA1SheetName, ":clear"));

              case 2:
                this.resetLocalCache(true);

              case 3:
              case "end":
                return _context72.stop();
            }
          }
        }, _callee72, this);
      }));

      function clear() {
        return _clear.apply(this, arguments);
      }

      return clear;
    }()
  }, {
    key: "sheetId",
    get: function get() {
      return this._getProp('sheetId');
    },
    set: function set(newVal) {
      return this._setProp('sheetId', newVal);
    }
  }, {
    key: "title",
    get: function get() {
      return this._getProp('title');
    },
    set: function set(newVal) {
      return this._setProp('title', newVal);
    }
  }, {
    key: "index",
    get: function get() {
      return this._getProp('index');
    },
    set: function set(newVal) {
      return this._setProp('index', newVal);
    }
  }, {
    key: "sheetType",
    get: function get() {
      return this._getProp('sheetType');
    },
    set: function set(newVal) {
      return this._setProp('sheetType', newVal);
    }
  }, {
    key: "gridProperties",
    get: function get() {
      return this._getProp('gridProperties');
    },
    set: function set(newVal) {
      return this._setProp('gridProperties', newVal);
    }
  }, {
    key: "hidden",
    get: function get() {
      return this._getProp('hidden');
    },
    set: function set(newVal) {
      return this._setProp('hidden', newVal);
    }
  }, {
    key: "tabColor",
    get: function get() {
      return this._getProp('tabColor');
    },
    set: function set(newVal) {
      return this._setProp('tabColor', newVal);
    }
  }, {
    key: "rightToLeft",
    get: function get() {
      return this._getProp('rightToLeft');
    },
    set: function set(newVal) {
      return this._setProp('rightToLeft', newVal);
    }
  }, {
    key: "rowCount",
    get: function get() {
      this._ensureInfoLoaded();

      return this.gridProperties.rowCount;
    },
    set: function set(newVal) {
      throw new Error('Do not update directly. Use resize()');
    }
  }, {
    key: "columnCount",
    get: function get() {
      this._ensureInfoLoaded();

      return this.gridProperties.columnCount;
    },
    set: function set(newVal) {
      throw new Error('Do not update directly. Use resize()');
    }
  }, {
    key: "colCount",
    get: function get() {
      throw new Error('`colCount` is deprecated - please use `columnCount` instead.');
    }
  }, {
    key: "a1SheetName",
    get: function get() {
      return "'".concat(this.title.replace(/'/g, "''"), "'");
    }
  }, {
    key: "encodedA1SheetName",
    get: function get() {
      return encodeURIComponent(this.a1SheetName);
    }
  }, {
    key: "lastColumnLetter",
    get: function get() {
      return columnToLetter(this.columnCount);
    } // CELLS-BASED INTERACTIONS //////////////////////////////////////////////////////////////////////

  }, {
    key: "cellStats",
    get: function get() {
      var allCells = _.flatten(this._cells);

      allCells = _.compact(allCells);
      return {
        nonEmpty: _.filter(allCells, function (c) {
          return c.value;
        }).length,
        loaded: allCells.length,
        total: this.rowCount * this.columnCount
      };
    }
  }]);

  return GoogleSpreadsheetWorksheet;
}();

module.exports = GoogleSpreadsheetWorksheet;