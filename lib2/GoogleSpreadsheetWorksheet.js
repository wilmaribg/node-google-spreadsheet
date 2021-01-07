'use strict';

var _extends = Object.assign || function (target) { for (var i = 1; i < arguments.length; i++) { var source = arguments[i]; for (var key in source) { if (Object.prototype.hasOwnProperty.call(source, key)) { target[key] = source[key]; } } } return target; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _toConsumableArray(arr) { if (Array.isArray(arr)) { for (var i = 0, arr2 = Array(arr.length); i < arr.length; i++) { arr2[i] = arr[i]; } return arr2; } else { return Array.from(arr); } }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

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
      throw new Error('Duplicate header detected: "' + header + '". Please make sure all non-empty headers are unique');
    }
  });
}

var GoogleSpreadsheetWorksheet = function () {
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
  }

  // INTERNAL UTILITY FUNCTIONS ////////////////////////////////////////////////////////////////////


  _createClass(GoogleSpreadsheetWorksheet, [{
    key: '_makeSingleUpdateRequest',
    value: async function _makeSingleUpdateRequest(requestType, requestParams) {
      // pass the call up to the parent
      return this._spreadsheet._makeSingleUpdateRequest(requestType, _extends({}, requestParams));
    }
  }, {
    key: '_ensureInfoLoaded',
    value: function _ensureInfoLoaded() {
      if (!this._rawProperties) {
        throw new Error('You must call `doc.loadInfo()` again before accessing this property');
      }
    }
  }, {
    key: 'resetLocalCache',
    value: function resetLocalCache(dataOnly) {
      if (!dataOnly) this._rawProperties = null;
      this.headerValues = null;
      this._cells = [];
    }
  }, {
    key: '_fillCellData',
    value: function _fillCellData(dataRanges) {
      var _this = this;

      _.each(dataRanges, function (range) {
        var startRow = range.startRow || 0;
        var startColumn = range.startColumn || 0;
        var numRows = range.rowMetadata.length;
        var numColumns = range.columnMetadata.length;

        // update cell data for entire range
        for (var i = 0; i < numRows; i++) {
          var actualRow = startRow + i;
          for (var j = 0; j < numColumns; j++) {
            var actualColumn = startColumn + j;

            // if the row has not been initialized yet, do it
            if (!_this._cells[actualRow]) _this._cells[actualRow] = [];

            // see if the response includes some info for the cell
            var cellData = _.get(range, 'rowData[' + i + '].values[' + j + ']');

            // update the cell object or create it
            if (_this._cells[actualRow][actualColumn]) {
              _this._cells[actualRow][actualColumn]._updateRawData(cellData);
            } else {
              _this._cells[actualRow][actualColumn] = new GoogleSpreadsheetCell(_this, actualRow, actualColumn, cellData);
            }
          }
        }

        // update row metadata
        for (var _i = 0; _i < range.rowMetadata.length; _i++) {
          _this._rowMetadata[startRow + _i] = range.rowMetadata[_i];
        }
        // update column metadata
        for (var _i2 = 0; _i2 < range.columnMetadata.length; _i2++) {
          _this._columnMetadata[startColumn + _i2] = range.columnMetadata[_i2];
        }
      });
    }

    // PROPERTY GETTERS //////////////////////////////////////////////////////////////////////////////

  }, {
    key: '_getProp',
    value: function _getProp(param) {
      this._ensureInfoLoaded();
      return this._rawProperties[param];
    }
  }, {
    key: '_setProp',
    value: function _setProp(param, newVal) {
      // eslint-disable-line no-unused-vars
      throw new Error('Do not update directly - use `updateProperties()`');
    }
  }, {
    key: 'getCellByA1',
    value: function getCellByA1(a1Address) {
      var split = a1Address.match(/([A-Z]+)([0-9]+)/);
      var columnIndex = letterToColumn(split[1]);
      var rowIndex = parseInt(split[2]);
      return this.getCell(rowIndex - 1, columnIndex - 1);
    }
  }, {
    key: 'getCell',
    value: function getCell(rowIndex, columnIndex) {
      if (rowIndex < 0 || columnIndex < 0) throw new Error('Min coordinate is 0, 0');
      if (rowIndex >= this.rowCount || columnIndex >= this.columnCount) {
        throw new Error('Out of bounds, sheet is ' + this.rowCount + ' by ' + this.columnCount);
      }

      if (!_.get(this._cells, '[' + rowIndex + '][' + columnIndex + ']')) {
        throw new Error('This cell has not been loaded yet');
      }
      return this._cells[rowIndex][columnIndex];
    }
  }, {
    key: 'loadCells',
    value: async function loadCells(sheetFilters) {
      var _this2 = this;

      // load the whole sheet
      if (!sheetFilters) return this._spreadsheet.loadCells(this.a1SheetName);

      var filtersArray = _.isArray(sheetFilters) ? sheetFilters : [sheetFilters];
      filtersArray = _.map(filtersArray, function (filter) {
        // add sheet name to A1 ranges
        if (_.isString(filter)) {
          if (filter.startsWith(_this2.a1SheetName)) return filter;
          return _this2.a1SheetName + '!' + filter;
        }
        if (_.isObject(filter)) {
          // TODO: detect and support DeveloperMetadata filters
          if (!filter.sheetId) {
            return _extends({ sheetId: _this2.sheetId }, filter);
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
      return this._spreadsheet.loadCells(filtersArray);
    }
  }, {
    key: 'saveUpdatedCells',
    value: async function saveUpdatedCells() {
      var cellsToSave = _.filter(_.flatten(this._cells), { _isDirty: true });
      if (cellsToSave.length) {
        await this.saveCells(cellsToSave);
      }
      // TODO: do we want to return stats? or the cells that got updated?
    }
  }, {
    key: 'saveCells',
    value: async function saveCells(cellsToUpdate) {
      var _this3 = this;

      // we send an individual "updateCells" request for each cell
      // because the fields that are udpated for each group are the same
      // and we dont want to accidentally overwrite something
      var requests = _.map(cellsToUpdate, function (cell) {
        return cell._getUpdateRequest();
      });
      var responseRanges = _.map(cellsToUpdate, function (c) {
        return _this3.a1SheetName + '!' + c.a1Address;
      });

      // if nothing is being updated the request returned is just `null`
      // so we make sure at least 1 request is valid - otherwise google throws a 400
      if (!_.compact(requests).length) {
        throw new Error('At least one cell must have something to update');
      }

      await this._spreadsheet._makeBatchUpdateRequest(requests, responseRanges);
    }

    // SAVING THIS FOR FUTURE USE
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
    key: 'loadHeaderRow',
    value: async function loadHeaderRow() {
      var rows = await this.getCellsInRange('A1:' + this.lastColumnLetter + '1');
      if (!rows) {
        throw new Error('No values in the header row - fill the first row with header values before trying to interact with rows');
      }
      this.headerValues = _.map(rows[0], function (header) {
        return header.trim();
      });
      if (!_.compact(this.headerValues).length) {
        throw new Error('All your header cells are blank - fill the first row with header values before trying to interact with rows');
      }
      checkForDuplicateHeaders(this.headerValues);
    }
  }, {
    key: 'setHeaderRow',
    value: async function setHeaderRow(headerValues) {
      if (!headerValues) return;
      if (headerValues.length > this.columnCount) {
        throw new Error('Sheet is not large enough to fit ' + headerValues.length + ' columns. Resize the sheet first.');
      }
      var trimmedHeaderValues = _.map(headerValues, function (h) {
        return h.trim();
      });
      checkForDuplicateHeaders(trimmedHeaderValues);

      if (!_.compact(trimmedHeaderValues).length) {
        throw new Error('All your header cells are blank -');
      }

      var response = await this._spreadsheet.axios.request({
        method: 'put',
        url: '/values/' + this.encodedA1SheetName + '!1:1',
        params: {
          valueInputOption: 'USER_ENTERED', // other option is RAW
          includeValuesInResponse: true
        },
        data: {
          range: this.a1SheetName + '!1:1',
          majorDimension: 'ROWS',
          values: [[].concat(_toConsumableArray(trimmedHeaderValues), _toConsumableArray(_.times(this.columnCount - trimmedHeaderValues.length, function () {
            return '';
          })))]
        }
      });
      this.headerValues = response.data.updatedData.values[0];
    }
  }, {
    key: 'addRows',
    value: async function addRows(rows) {
      var _this4 = this;

      var options = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};

      // adds multiple rows in one API interaction using the append endpoint

      // each row can be an array or object
      // an array is just cells
      // ex: ['column 1', 'column 2', 'column 3']
      // an object must use the header row values as keys
      // ex: { col1: 'column 1', col2: 'column 2', col3: 'column 3' }

      // google bug that does not handle colons in names
      // see https://issuetracker.google.com/issues/150373119
      if (this.title.includes(':')) {
        throw new Error('Please remove the ":" from your sheet title. There is a bug with the google API which breaks appending rows if any colons are in the sheet title.');
      }

      if (!_.isArray(rows)) throw new Error('You must pass in an array of row values to append');

      if (!this.headerValues) await this.loadHeaderRow();

      // convert each row into an array of cell values rather than the key/value object
      var rowsAsArrays = [];
      _.each(rows, function (row) {
        var rowAsArray = void 0;
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

      var response = await this._spreadsheet.axios.request({
        method: 'post',
        url: '/values/' + this.encodedA1SheetName + '!A1:append',
        params: {
          valueInputOption: options.raw ? 'RAW' : 'USER_ENTERED',
          insertDataOption: options.insert ? 'INSERT_ROWS' : 'OVERWRITE',
          includeValuesInResponse: true
        },
        data: {
          values: rowsAsArrays
        }
      });

      // extract the new row number from the A1-notation data range in the response
      // ex: in "'Sheet8!A2:C2" -- we want the `2`
      var updatedRange = response.data.updates.updatedRange;

      var rowNumber = updatedRange.match(/![A-Z]+([0-9]+):?/)[1];
      rowNumber = parseInt(rowNumber);

      // if new rows were added, we need update sheet.rowRount
      if (options.insert) {
        this._rawProperties.gridProperties.rowCount += rows.length;
      } else if (rowNumber + rows.length > this.rowCount) {
        // have to subtract 1 since one row was inserted at rowNumber
        this._rawProperties.gridProperties.rowCount = rowNumber + rows.length - 1;
      }

      return _.map(response.data.updates.updatedData.values, function (rowValues) {
        var row = new GoogleSpreadsheetRow(_this4, rowNumber++, rowValues);
        return row;
      });
    }
  }, {
    key: 'addRow',
    value: async function addRow(rowValues, options) {
      var rows = await this.addRows([rowValues], options);
      return rows[0];
    }
  }, {
    key: 'getRows',
    value: async function getRows() {
      var options = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};

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

      if (!this.headerValues) await this.loadHeaderRow();

      var firstRow = 2 + options.offset; // skip first row AND not zero indexed
      var lastRow = firstRow + options.limit - 1; // inclusive so we subtract 1
      var lastColumn = columnToLetter(this.headerValues.length);
      var rawRows = await this.getCellsInRange('A' + firstRow + ':' + lastColumn + lastRow);

      if (!rawRows) return [];

      var rows = [];
      var rowNum = firstRow;
      for (var i = 0; i < rawRows.length; i++) {
        rows.push(new GoogleSpreadsheetRow(this, rowNum++, rawRows[i]));
      }
      return rows;
    }

    // BASIC PROPS ///////////////////////////////////////////////////////////////////////////////////

  }, {
    key: 'updateProperties',
    value: async function updateProperties(properties) {
      // Request type = `updateSheetProperties`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateSheetPropertiesRequest

      // properties
      // - title (string)
      // - index (number)
      // - gridProperties ({ object (GridProperties) } - https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#gridproperties
      // - hidden (boolean)
      // - tabColor ({ object (Color) } - https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#Color
      // - rightToLeft (boolean)

      return this._makeSingleUpdateRequest('updateSheetProperties', {
        properties: _extends({
          sheetId: this.sheetId
        }, properties),
        fields: getFieldMask(properties)
      });
    }
  }, {
    key: 'updateGridProperties',
    value: async function updateGridProperties(gridProperties) {
      // just passes the call through to update gridProperties
      // see https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#GridProperties

      // gridProperties
      // - rowCount
      // - columnCount
      // - frozenRowCount
      // - frozenColumnCount
      // - hideGridLines
      return this.updateProperties({ gridProperties: gridProperties });
    }

    // just a shortcut because resize makes more sense to change rowCount / columnCount

  }, {
    key: 'resize',
    value: async function resize(gridProperties) {
      return this.updateGridProperties(gridProperties);
    }
  }, {
    key: 'updateDimensionProperties',
    value: async function updateDimensionProperties(columnsOrRows, properties, bounds) {
      // Request type = `updateDimensionProperties`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#updatedimensionpropertiesrequest

      // columnsOrRows = COLUMNS|ROWS
      // properties
      // - pixelSize
      // - hiddenByUser
      // - developerMetadata
      // bounds
      // - startIndex
      // - endIndex

      return this._makeSingleUpdateRequest('updateDimensionProperties', {
        range: _extends({
          sheetId: this.sheetId,
          dimension: columnsOrRows
        }, bounds && {
          startIndex: bounds.startIndex,
          endIndex: bounds.endIndex
        }),
        properties: properties,
        fields: getFieldMask(properties)
      });
    }

    // OTHER /////////////////////////////////////////////////////////////////////////////////////////

    // this uses the "values" getter and does not give all the info about the cell contents
    // it is used internally when loading header cells

  }, {
    key: 'getCellsInRange',
    value: async function getCellsInRange(a1Range, options) {
      var response = await this._spreadsheet.axios.get('/values/' + this.encodedA1SheetName + '!' + a1Range, {
        params: options
      });
      return response.data.values;
    }
  }, {
    key: 'updateNamedRange',
    value: async function updateNamedRange() {
      // Request type = `updateNamedRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateNamedRangeRequest
    }
  }, {
    key: 'addNamedRange',
    value: async function addNamedRange() {
      // Request type = `addNamedRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddNamedRangeRequest
    }
  }, {
    key: 'deleteNamedRange',
    value: async function deleteNamedRange() {
      // Request type = `deleteNamedRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteNamedRangeRequest
    }
  }, {
    key: 'repeatCell',
    value: async function repeatCell() {
      // Request type = `repeatCell`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#RepeatCellRequest
    }
  }, {
    key: 'autoFill',
    value: async function autoFill() {
      // Request type = `autoFill`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AutoFillRequest
    }
  }, {
    key: 'cutPaste',
    value: async function cutPaste() {
      // Request type = `cutPaste`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#CutPasteRequest
    }
  }, {
    key: 'copyPaste',
    value: async function copyPaste() {
      // Request type = `copyPaste`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#CopyPasteRequest
    }
  }, {
    key: 'mergeCells',
    value: async function mergeCells(range) {
      var mergeType = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 'MERGE_ALL';

      // Request type = `mergeCells`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#MergeCellsRequest
      if (range.sheetId && range.sheetId !== this.sheetId) {
        throw new Error('Leave sheet ID blank or set to matching ID of this sheet');
      }
      await this._makeSingleUpdateRequest('mergeCells', {
        mergeType: mergeType,
        range: _extends({}, range, {
          sheetId: this.sheetId
        })
      });
    }
  }, {
    key: 'unmergeCells',
    value: async function unmergeCells(range) {
      // Request type = `unmergeCells`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UnmergeCellsRequest
      if (range.sheetId && range.sheetId !== this.sheetId) {
        throw new Error('Leave sheet ID blank or set to matching ID of this sheet');
      }
      await this._makeSingleUpdateRequest('unmergeCells', {
        range: _extends({}, range, {
          sheetId: this.sheetId
        })
      });
    }
  }, {
    key: 'updateBorders',
    value: async function updateBorders() {
      // Request type = `updateBorders`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateBordersRequest
    }
  }, {
    key: 'addFilterView',
    value: async function addFilterView() {
      // Request type = `addFilterView`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddFilterViewRequest
    }
  }, {
    key: 'appendCells',
    value: async function appendCells() {
      // Request type = `appendCells`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AppendCellsRequest
    }
  }, {
    key: 'clearBasicFilter',
    value: async function clearBasicFilter() {
      // Request type = `clearBasicFilter`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#ClearBasicFilterRequest
    }
  }, {
    key: 'deleteDimension',
    value: async function deleteDimension() {
      // Request type = `deleteDimension`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteDimensionRequest
    }
  }, {
    key: 'deleteEmbeddedObject',
    value: async function deleteEmbeddedObject() {
      // Request type = `deleteEmbeddedObject`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteEmbeddedObjectRequest
    }
  }, {
    key: 'deleteFilterView',
    value: async function deleteFilterView() {
      // Request type = `deleteFilterView`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteFilterViewRequest
    }
  }, {
    key: 'duplicateFilterView',
    value: async function duplicateFilterView() {
      // Request type = `duplicateFilterView`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DuplicateFilterViewRequest
    }
  }, {
    key: 'duplicateSheet',
    value: async function duplicateSheet() {
      // Request type = `duplicateSheet`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DuplicateSheetRequest
    }
  }, {
    key: 'findReplace',
    value: async function findReplace() {
      // Request type = `findReplace`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#FindReplaceRequest
    }
  }, {
    key: 'insertDimension',
    value: async function insertDimension() {
      // Request type = `insertDimension`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#InsertDimensionRequest
    }
  }, {
    key: 'insertRange',
    value: async function insertRange() {
      // Request type = `insertRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#InsertRangeRequest
    }
  }, {
    key: 'moveDimension',
    value: async function moveDimension() {
      // Request type = `moveDimension`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#MoveDimensionRequest
    }
  }, {
    key: 'updateEmbeddedObjectPosition',
    value: async function updateEmbeddedObjectPosition() {
      // Request type = `updateEmbeddedObjectPosition`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateEmbeddedObjectPositionRequest
    }
  }, {
    key: 'pasteData',
    value: async function pasteData() {
      // Request type = `pasteData`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#PasteDataRequest
    }
  }, {
    key: 'textToColumns',
    value: async function textToColumns() {
      // Request type = `textToColumns`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#TextToColumnsRequest
    }
  }, {
    key: 'updateFilterView',
    value: async function updateFilterView() {
      // Request type = `updateFilterView`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateFilterViewRequest
    }
  }, {
    key: 'deleteRange',
    value: async function deleteRange() {
      // Request type = `deleteRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteRangeRequest
    }
  }, {
    key: 'appendDimension',
    value: async function appendDimension() {
      // Request type = `appendDimension`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AppendDimensionRequest
    }
  }, {
    key: 'addConditionalFormatRule',
    value: async function addConditionalFormatRule() {
      // Request type = `addConditionalFormatRule`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddConditionalFormatRuleRequest
    }
  }, {
    key: 'updateConditionalFormatRule',
    value: async function updateConditionalFormatRule() {
      // Request type = `updateConditionalFormatRule`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateConditionalFormatRuleRequest
    }
  }, {
    key: 'deleteConditionalFormatRule',
    value: async function deleteConditionalFormatRule() {
      // Request type = `deleteConditionalFormatRule`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteConditionalFormatRuleRequest
    }
  }, {
    key: 'sortRange',
    value: async function sortRange() {
      // Request type = `sortRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#SortRangeRequest
    }
  }, {
    key: 'setDataValidation',
    value: async function setDataValidation() {
      // Request type = `setDataValidation`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#SetDataValidationRequest
    }
  }, {
    key: 'setBasicFilter',
    value: async function setBasicFilter() {
      // Request type = `setBasicFilter`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#SetBasicFilterRequest
    }
  }, {
    key: 'addProtectedRange',
    value: async function addProtectedRange() {
      // Request type = `addProtectedRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddProtectedRangeRequest
    }
  }, {
    key: 'updateProtectedRange',
    value: async function updateProtectedRange() {
      // Request type = `updateProtectedRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateProtectedRangeRequest
    }
  }, {
    key: 'deleteProtectedRange',
    value: async function deleteProtectedRange() {
      // Request type = `deleteProtectedRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteProtectedRangeRequest
    }
  }, {
    key: 'autoResizeDimensions',
    value: async function autoResizeDimensions() {
      // Request type = `autoResizeDimensions`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AutoResizeDimensionsRequest
    }
  }, {
    key: 'addChart',
    value: async function addChart() {
      // Request type = `addChart`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddChartRequest
    }
  }, {
    key: 'updateChartSpec',
    value: async function updateChartSpec() {
      // Request type = `updateChartSpec`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateChartSpecRequest
    }
  }, {
    key: 'updateBanding',
    value: async function updateBanding() {
      // Request type = `updateBanding`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateBandingRequest
    }
  }, {
    key: 'addBanding',
    value: async function addBanding() {
      // Request type = `addBanding`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddBandingRequest
    }
  }, {
    key: 'deleteBanding',
    value: async function deleteBanding() {
      // Request type = `deleteBanding`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteBandingRequest
    }
  }, {
    key: 'createDeveloperMetadata',
    value: async function createDeveloperMetadata() {
      // Request type = `createDeveloperMetadata`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#CreateDeveloperMetadataRequest
    }
  }, {
    key: 'updateDeveloperMetadata',
    value: async function updateDeveloperMetadata() {
      // Request type = `updateDeveloperMetadata`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateDeveloperMetadataRequest
    }
  }, {
    key: 'deleteDeveloperMetadata',
    value: async function deleteDeveloperMetadata() {
      // Request type = `deleteDeveloperMetadata`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteDeveloperMetadataRequest
    }
  }, {
    key: 'randomizeRange',
    value: async function randomizeRange() {
      // Request type = `randomizeRange`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#RandomizeRangeRequest
    }
  }, {
    key: 'addDimensionGroup',
    value: async function addDimensionGroup() {
      // Request type = `addDimensionGroup`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddDimensionGroupRequest
    }
  }, {
    key: 'deleteDimensionGroup',
    value: async function deleteDimensionGroup() {
      // Request type = `deleteDimensionGroup`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteDimensionGroupRequest
    }
  }, {
    key: 'updateDimensionGroup',
    value: async function updateDimensionGroup() {
      // Request type = `updateDimensionGroup`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateDimensionGroupRequest
    }
  }, {
    key: 'trimWhitespace',
    value: async function trimWhitespace() {
      // Request type = `trimWhitespace`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#TrimWhitespaceRequest
    }
  }, {
    key: 'deleteDuplicates',
    value: async function deleteDuplicates() {
      // Request type = `deleteDuplicates`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteDuplicatesRequest
    }
  }, {
    key: 'addSlicer',
    value: async function addSlicer() {
      // Request type = `addSlicer`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddSlicerRequest
    }
  }, {
    key: 'updateSlicerSpec',
    value: async function updateSlicerSpec() {}
    // Request type = `updateSlicerSpec`
    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#UpdateSlicerSpecRequest


    // delete this worksheet

  }, {
    key: 'delete',
    value: async function _delete() {
      return this._spreadsheet.deleteSheet(this.sheetId);
    }
  }, {
    key: 'del',
    value: async function del() {
      return this.delete();
    } // alias to mimic old interface

    // copies this worksheet into another document/spreadsheet

  }, {
    key: 'copyToSpreadsheet',
    value: async function copyToSpreadsheet(destinationSpreadsheetId) {
      return this._spreadsheet.axios.post('/sheets/' + this.sheetId + ':copyTo', {
        destinationSpreadsheetId: destinationSpreadsheetId
      });
    }
  }, {
    key: 'clear',
    value: async function clear() {
      // clears all the data in the sheet
      // sheet name without ie 'sheet1' rather than 'sheet1'!A1:B5 is all cells
      await this._spreadsheet.axios.post('/values/' + this.encodedA1SheetName + ':clear');
      this.resetLocalCache(true);
    }
  }, {
    key: 'sheetId',
    get: function get() {
      return this._getProp('sheetId');
    },
    set: function set(newVal) {
      return this._setProp('sheetId', newVal);
    }
  }, {
    key: 'title',
    get: function get() {
      return this._getProp('title');
    },
    set: function set(newVal) {
      return this._setProp('title', newVal);
    }
  }, {
    key: 'index',
    get: function get() {
      return this._getProp('index');
    },
    set: function set(newVal) {
      return this._setProp('index', newVal);
    }
  }, {
    key: 'sheetType',
    get: function get() {
      return this._getProp('sheetType');
    },
    set: function set(newVal) {
      return this._setProp('sheetType', newVal);
    }
  }, {
    key: 'gridProperties',
    get: function get() {
      return this._getProp('gridProperties');
    },
    set: function set(newVal) {
      return this._setProp('gridProperties', newVal);
    }
  }, {
    key: 'hidden',
    get: function get() {
      return this._getProp('hidden');
    },
    set: function set(newVal) {
      return this._setProp('hidden', newVal);
    }
  }, {
    key: 'tabColor',
    get: function get() {
      return this._getProp('tabColor');
    },
    set: function set(newVal) {
      return this._setProp('tabColor', newVal);
    }
  }, {
    key: 'rightToLeft',
    get: function get() {
      return this._getProp('rightToLeft');
    },
    set: function set(newVal) {
      return this._setProp('rightToLeft', newVal);
    }
  }, {
    key: 'rowCount',
    get: function get() {
      this._ensureInfoLoaded();
      return this.gridProperties.rowCount;
    },
    set: function set(newVal) {
      throw new Error('Do not update directly. Use resize()');
    }
  }, {
    key: 'columnCount',
    get: function get() {
      this._ensureInfoLoaded();
      return this.gridProperties.columnCount;
    },
    set: function set(newVal) {
      throw new Error('Do not update directly. Use resize()');
    }
  }, {
    key: 'colCount',
    get: function get() {
      throw new Error('`colCount` is deprecated - please use `columnCount` instead.');
    }
  }, {
    key: 'a1SheetName',
    get: function get() {
      return '\'' + this.title.replace(/'/g, "''") + '\'';
    }
  }, {
    key: 'encodedA1SheetName',
    get: function get() {
      return encodeURIComponent(this.a1SheetName);
    }
  }, {
    key: 'lastColumnLetter',
    get: function get() {
      return columnToLetter(this.columnCount);
    }

    // CELLS-BASED INTERACTIONS //////////////////////////////////////////////////////////////////////

  }, {
    key: 'cellStats',
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