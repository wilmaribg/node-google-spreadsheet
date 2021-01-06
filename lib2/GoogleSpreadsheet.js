"use strict";

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); if (enumerableOnly) symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; }); keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = arguments[i] != null ? arguments[i] : {}; if (i % 2) { ownKeys(Object(source), true).forEach(function (key) { _defineProperty(target, key, source[key]); }); } else if (Object.getOwnPropertyDescriptors) { Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)); } else { ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } } return target; }

function _defineProperty(obj, key, value) { if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable: true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }

function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

function _typeof(obj) { "@babel/helpers - typeof"; if (typeof Symbol === "function" && typeof Symbol.iterator === "symbol") { _typeof = function _typeof(obj) { return typeof obj; }; } else { _typeof = function _typeof(obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }; } return _typeof(obj); }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }

function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); return Constructor; }

var _ = require('lodash');

var _require = require('google-auth-library'),
    JWT = _require.JWT;

var Axios = require('axios');

var GoogleSpreadsheetWorksheet = require('./GoogleSpreadsheetWorksheet');

var _require2 = require('./utils'),
    getFieldMask = _require2.getFieldMask;

var GOOGLE_AUTH_SCOPES = ['https://www.googleapis.com/auth/spreadsheets' // the list from the sheets v4 auth for spreadsheets.get
// 'https://www.googleapis.com/auth/drive',
// 'https://www.googleapis.com/auth/drive.readonly',
// 'https://www.googleapis.com/auth/drive.file',
// 'https://www.googleapis.com/auth/spreadsheets',
// 'https://www.googleapis.com/auth/spreadsheets.readonly',
];
var AUTH_MODES = {
  JWT: 'JWT',
  API_KEY: 'API_KEY',
  RAW_ACCESS_TOKEN: 'RAW_ACCESS_TOKEN',
  OAUTH: 'OAUTH'
};

var GoogleSpreadsheet = /*#__PURE__*/function () {
  function GoogleSpreadsheet(sheetId) {
    _classCallCheck(this, GoogleSpreadsheet);

    this.spreadsheetId = sheetId;
    this.authMode = null;
    this._rawSheets = {};
    this._rawProperties = null; // create an axios instance with sheet root URL and interceptors to handle auth

    this.axios = Axios.create({
      baseURL: "https://sheets.googleapis.com/v4/spreadsheets/".concat(sheetId || ''),
      // send arrays in params with duplicate keys - ie `?thing=1&thing=2` vs `?thing[]=1...`
      // solution taken from https://github.com/axios/axios/issues/604
      paramsSerializer: function paramsSerializer(params) {
        var options = '';

        _.keys(params).forEach(function (key) {
          var isParamTypeObject = _typeof(params[key]) === 'object';
          var isParamTypeArray = isParamTypeObject && params[key].length >= 0;
          if (!isParamTypeObject) options += "".concat(key, "=").concat(encodeURIComponent(params[key]), "&");

          if (isParamTypeObject && isParamTypeArray) {
            _.each(params[key], function (val) {
              options += "".concat(key, "=").concat(encodeURIComponent(val), "&");
            });
          }
        });

        return options ? options.slice(0, -1) : options;
      }
    }); // have to use bind here or the functions dont have access to `this` :(

    this.axios.interceptors.request.use(this._setAxiosRequestAuth.bind(this));
    this.axios.interceptors.response.use(this._handleAxiosResponse.bind(this), this._handleAxiosErrors.bind(this));
    return this;
  } // CREATE NEW DOC ////////////////////////////////////////////////////////////////////////////////


  _createClass(GoogleSpreadsheet, [{
    key: "createNewSpreadsheetDocument",
    value: function () {
      var _createNewSpreadsheetDocument = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee(properties) {
        var _this = this;

        var response;
        return regeneratorRuntime.wrap(function _callee$(_context) {
          while (1) {
            switch (_context.prev = _context.next) {
              case 0:
                if (!this.spreadsheetId) {
                  _context.next = 2;
                  break;
                }

                throw new Error('Only call `createNewSpreadsheetDocument()` on a GoogleSpreadsheet object that has no spreadsheetId set');

              case 2:
                _context.next = 4;
                return this.axios.post(this.url, {
                  properties: properties
                });

              case 4:
                response = _context.sent;
                this.spreadsheetId = response.data.spreadsheetId;
                this.axios.defaults.baseURL += this.spreadsheetId;
                this._rawProperties = response.data.properties;

                _.each(response.data.sheets, function (s) {
                  return _this._updateOrCreateSheet(s);
                });

              case 9:
              case "end":
                return _context.stop();
            }
          }
        }, _callee, this);
      }));

      function createNewSpreadsheetDocument(_x) {
        return _createNewSpreadsheetDocument.apply(this, arguments);
      }

      return createNewSpreadsheetDocument;
    }() // AUTH RELATED FUNCTIONS ////////////////////////////////////////////////////////////////////////

  }, {
    key: "useApiKey",
    value: function () {
      var _useApiKey = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee2(key) {
        return regeneratorRuntime.wrap(function _callee2$(_context2) {
          while (1) {
            switch (_context2.prev = _context2.next) {
              case 0:
                this.authMode = AUTH_MODES.API_KEY;
                this.apiKey = key;

              case 2:
              case "end":
                return _context2.stop();
            }
          }
        }, _callee2, this);
      }));

      function useApiKey(_x2) {
        return _useApiKey.apply(this, arguments);
      }

      return useApiKey;
    }() // token must be created and managed (refreshed) elsewhere

  }, {
    key: "useRawAccessToken",
    value: function () {
      var _useRawAccessToken = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee3(token) {
        return regeneratorRuntime.wrap(function _callee3$(_context3) {
          while (1) {
            switch (_context3.prev = _context3.next) {
              case 0:
                this.authMode = AUTH_MODES.RAW_ACCESS_TOKEN;
                this.accessToken = token;

              case 2:
              case "end":
                return _context3.stop();
            }
          }
        }, _callee3, this);
      }));

      function useRawAccessToken(_x3) {
        return _useRawAccessToken.apply(this, arguments);
      }

      return useRawAccessToken;
    }()
  }, {
    key: "useOAuth2Client",
    value: function () {
      var _useOAuth2Client = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee4(oAuth2Client) {
        return regeneratorRuntime.wrap(function _callee4$(_context4) {
          while (1) {
            switch (_context4.prev = _context4.next) {
              case 0:
                this.authMode = AUTH_MODES.OAUTH;
                this.oAuth2Client = oAuth2Client;

              case 2:
              case "end":
                return _context4.stop();
            }
          }
        }, _callee4, this);
      }));

      function useOAuth2Client(_x4) {
        return _useOAuth2Client.apply(this, arguments);
      }

      return useOAuth2Client;
    }() // creds should be an object obtained by loading the json file google gives you
    // impersonateAs is an email of any user in the G Suite domain
    // (only works if service account has domain-wide delegation enabled)

  }, {
    key: "useServiceAccountAuth",
    value: function () {
      var _useServiceAccountAuth = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee5(creds) {
        var impersonateAs,
            _args5 = arguments;
        return regeneratorRuntime.wrap(function _callee5$(_context5) {
          while (1) {
            switch (_context5.prev = _context5.next) {
              case 0:
                impersonateAs = _args5.length > 1 && _args5[1] !== undefined ? _args5[1] : null;
                this.jwtClient = new JWT({
                  email: creds.client_email,
                  key: creds.private_key,
                  scopes: GOOGLE_AUTH_SCOPES,
                  subject: impersonateAs
                });
                _context5.next = 4;
                return this.renewJwtAuth();

              case 4:
              case "end":
                return _context5.stop();
            }
          }
        }, _callee5, this);
      }));

      function useServiceAccountAuth(_x5) {
        return _useServiceAccountAuth.apply(this, arguments);
      }

      return useServiceAccountAuth;
    }()
  }, {
    key: "renewJwtAuth",
    value: function () {
      var _renewJwtAuth = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee6() {
        return regeneratorRuntime.wrap(function _callee6$(_context6) {
          while (1) {
            switch (_context6.prev = _context6.next) {
              case 0:
                this.authMode = AUTH_MODES.JWT;
                _context6.next = 3;
                return this.jwtClient.authorize();

              case 3:
              case "end":
                return _context6.stop();
            }
          }
        }, _callee6, this);
      }));

      function renewJwtAuth() {
        return _renewJwtAuth.apply(this, arguments);
      }

      return renewJwtAuth;
    }() // TODO: provide mechanism to share single JWT auth between docs?
    // INTERNAL UTILITY FUNCTIONS ////////////////////////////////////////////////////////////////////

  }, {
    key: "_setAxiosRequestAuth",
    value: function () {
      var _setAxiosRequestAuth2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee7(config) {
        var credentials;
        return regeneratorRuntime.wrap(function _callee7$(_context7) {
          while (1) {
            switch (_context7.prev = _context7.next) {
              case 0:
                if (!(this.authMode === AUTH_MODES.JWT)) {
                  _context7.next = 8;
                  break;
                }

                if (this.jwtClient) {
                  _context7.next = 3;
                  break;
                }

                throw new Error('JWT auth is not set up properly');

              case 3:
                _context7.next = 5;
                return this.jwtClient.authorize();

              case 5:
                config.headers.Authorization = "Bearer ".concat(this.jwtClient.credentials.access_token);
                _context7.next = 29;
                break;

              case 8:
                if (!(this.authMode === AUTH_MODES.RAW_ACCESS_TOKEN)) {
                  _context7.next = 14;
                  break;
                }

                if (this.accessToken) {
                  _context7.next = 11;
                  break;
                }

                throw new Error('Invalid access token');

              case 11:
                config.headers.Authorization = "Bearer ".concat(this.accessToken);
                _context7.next = 29;
                break;

              case 14:
                if (!(this.authMode === AUTH_MODES.API_KEY)) {
                  _context7.next = 21;
                  break;
                }

                if (this.apiKey) {
                  _context7.next = 17;
                  break;
                }

                throw new Error('Please set API key');

              case 17:
                config.params = config.params || {};
                config.params.key = this.apiKey;
                _context7.next = 29;
                break;

              case 21:
                if (!(this.authMode === AUTH_MODES.OAUTH)) {
                  _context7.next = 28;
                  break;
                }

                _context7.next = 24;
                return this.oAuth2Client.getAccessToken();

              case 24:
                credentials = _context7.sent;
                config.headers.Authorization = "Bearer ".concat(credentials.token);
                _context7.next = 29;
                break;

              case 28:
                throw new Error('You must initialize some kind of auth before making any requests');

              case 29:
                return _context7.abrupt("return", config);

              case 30:
              case "end":
                return _context7.stop();
            }
          }
        }, _callee7, this);
      }));

      function _setAxiosRequestAuth(_x6) {
        return _setAxiosRequestAuth2.apply(this, arguments);
      }

      return _setAxiosRequestAuth;
    }()
  }, {
    key: "_handleAxiosResponse",
    value: function () {
      var _handleAxiosResponse2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee8(response) {
        return regeneratorRuntime.wrap(function _callee8$(_context8) {
          while (1) {
            switch (_context8.prev = _context8.next) {
              case 0:
                return _context8.abrupt("return", response);

              case 1:
              case "end":
                return _context8.stop();
            }
          }
        }, _callee8);
      }));

      function _handleAxiosResponse(_x7) {
        return _handleAxiosResponse2.apply(this, arguments);
      }

      return _handleAxiosResponse;
    }()
  }, {
    key: "_handleAxiosErrors",
    value: function () {
      var _handleAxiosErrors2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee9(error) {
        var _error$response$data$, code, message;

        return regeneratorRuntime.wrap(function _callee9$(_context9) {
          while (1) {
            switch (_context9.prev = _context9.next) {
              case 0:
                if (!(error.response && error.response.data)) {
                  _context9.next = 6;
                  break;
                }

                if (error.response.data.error) {
                  _context9.next = 3;
                  break;
                }

                throw error;

              case 3:
                _error$response$data$ = error.response.data.error, code = _error$response$data$.code, message = _error$response$data$.message;
                error.message = "Google API error - [".concat(code, "] ").concat(message);
                throw error;

              case 6:
                if (!(_.get(error, 'response.status') === 403)) {
                  _context9.next = 9;
                  break;
                }

                if (!(this.authMode === AUTH_MODES.API_KEY)) {
                  _context9.next = 9;
                  break;
                }

                throw new Error('Sheet is private. Use authentication or make public. (see https://github.com/theoephraim/node-google-spreadsheet#a-note-on-authentication for details)');

              case 9:
                throw error;

              case 10:
              case "end":
                return _context9.stop();
            }
          }
        }, _callee9, this);
      }));

      function _handleAxiosErrors(_x8) {
        return _handleAxiosErrors2.apply(this, arguments);
      }

      return _handleAxiosErrors;
    }()
  }, {
    key: "_makeSingleUpdateRequest",
    value: function () {
      var _makeSingleUpdateRequest2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee10(requestType, requestParams) {
        var _this2 = this;

        var response;
        return regeneratorRuntime.wrap(function _callee10$(_context10) {
          while (1) {
            switch (_context10.prev = _context10.next) {
              case 0:
                _context10.next = 2;
                return this.axios.post(':batchUpdate', {
                  requests: [_defineProperty({}, requestType, requestParams)],
                  includeSpreadsheetInResponse: true // responseRanges: [string]
                  // responseIncludeGridData: true

                });

              case 2:
                response = _context10.sent;

                this._updateRawProperties(response.data.updatedSpreadsheet.properties);

                _.each(response.data.updatedSpreadsheet.sheets, function (s) {
                  return _this2._updateOrCreateSheet(s);
                }); // console.log('API RESPONSE', response.data.replies[0][requestType]);


                return _context10.abrupt("return", response.data.replies[0][requestType]);

              case 6:
              case "end":
                return _context10.stop();
            }
          }
        }, _callee10, this);
      }));

      function _makeSingleUpdateRequest(_x9, _x10) {
        return _makeSingleUpdateRequest2.apply(this, arguments);
      }

      return _makeSingleUpdateRequest;
    }()
  }, {
    key: "_makeBatchUpdateRequest",
    value: function () {
      var _makeBatchUpdateRequest2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee11(requests, responseRanges) {
        var _this3 = this;

        var response;
        return regeneratorRuntime.wrap(function _callee11$(_context11) {
          while (1) {
            switch (_context11.prev = _context11.next) {
              case 0:
                _context11.next = 2;
                return this.axios.post(':batchUpdate', _objectSpread({
                  requests: requests,
                  includeSpreadsheetInResponse: true
                }, responseRanges && _objectSpread({
                  responseIncludeGridData: true
                }, responseRanges !== '*' && {
                  responseRanges: responseRanges
                })));

              case 2:
                response = _context11.sent;

                this._updateRawProperties(response.data.updatedSpreadsheet.properties);

                _.each(response.data.updatedSpreadsheet.sheets, function (s) {
                  return _this3._updateOrCreateSheet(s);
                });

              case 5:
              case "end":
                return _context11.stop();
            }
          }
        }, _callee11, this);
      }));

      function _makeBatchUpdateRequest(_x11, _x12) {
        return _makeBatchUpdateRequest2.apply(this, arguments);
      }

      return _makeBatchUpdateRequest;
    }()
  }, {
    key: "_ensureInfoLoaded",
    value: function _ensureInfoLoaded() {
      if (!this._rawProperties) throw new Error('You must call `doc.loadInfo()` before accessing this property');
    }
  }, {
    key: "_updateRawProperties",
    value: function _updateRawProperties(newProperties) {
      this._rawProperties = newProperties;
    }
  }, {
    key: "_updateOrCreateSheet",
    value: function _updateOrCreateSheet(_ref2) {
      var properties = _ref2.properties,
          data = _ref2.data;
      var sheetId = properties.sheetId;

      if (!this._rawSheets[sheetId]) {
        this._rawSheets[sheetId] = new GoogleSpreadsheetWorksheet(this, {
          properties: properties,
          data: data
        });
      } else {
        this._rawSheets[sheetId]._rawProperties = properties;

        this._rawSheets[sheetId]._fillCellData(data);
      }
    } // BASIC PROPS //////////////////////////////////////////////////////////////////////////////

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
    key: "updateProperties",
    value: function () {
      var _updateProperties = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee12(properties) {
        return regeneratorRuntime.wrap(function _callee12$(_context12) {
          while (1) {
            switch (_context12.prev = _context12.next) {
              case 0:
                _context12.next = 2;
                return this._makeSingleUpdateRequest('updateSpreadsheetProperties', {
                  properties: properties,
                  fields: getFieldMask(properties)
                });

              case 2:
              case "end":
                return _context12.stop();
            }
          }
        }, _callee12, this);
      }));

      function updateProperties(_x13) {
        return _updateProperties.apply(this, arguments);
      }

      return updateProperties;
    }() // BASIC INFO ////////////////////////////////////////////////////////////////////////////////////

  }, {
    key: "loadInfo",
    value: function () {
      var _loadInfo = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee13(includeCells) {
        var _this4 = this;

        var response;
        return regeneratorRuntime.wrap(function _callee13$(_context13) {
          while (1) {
            switch (_context13.prev = _context13.next) {
              case 0:
                _context13.next = 2;
                return this.axios.get('/', {
                  params: _objectSpread({}, includeCells && {
                    includeGridData: true
                  })
                });

              case 2:
                response = _context13.sent;
                this._rawProperties = response.data.properties;

                _.each(response.data.sheets, function (s) {
                  return _this4._updateOrCreateSheet(s);
                });

              case 5:
              case "end":
                return _context13.stop();
            }
          }
        }, _callee13, this);
      }));

      function loadInfo(_x14) {
        return _loadInfo.apply(this, arguments);
      }

      return loadInfo;
    }()
  }, {
    key: "getInfo",
    value: function () {
      var _getInfo = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee14() {
        return regeneratorRuntime.wrap(function _callee14$(_context14) {
          while (1) {
            switch (_context14.prev = _context14.next) {
              case 0:
                return _context14.abrupt("return", this.loadInfo());

              case 1:
              case "end":
                return _context14.stop();
            }
          }
        }, _callee14, this);
      }));

      function getInfo() {
        return _getInfo.apply(this, arguments);
      }

      return getInfo;
    }() // alias to mimic old version

  }, {
    key: "resetLocalCache",
    value: function resetLocalCache() {
      this._rawProperties = null;
      this._rawSheets = {};
    } // WORKSHEETS ////////////////////////////////////////////////////////////////////////////////////

  }, {
    key: "addSheet",
    value: function () {
      var _addSheet = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee15() {
        var properties,
            response,
            newSheetId,
            newSheet,
            _args15 = arguments;
        return regeneratorRuntime.wrap(function _callee15$(_context15) {
          while (1) {
            switch (_context15.prev = _context15.next) {
              case 0:
                properties = _args15.length > 0 && _args15[0] !== undefined ? _args15[0] : {};
                _context15.next = 3;
                return this._makeSingleUpdateRequest('addSheet', {
                  properties: _.omit(properties, 'headers', 'headerValues')
                });

              case 3:
                response = _context15.sent;
                // _makeSingleUpdateRequest already adds the sheet
                newSheetId = response.properties.sheetId;
                newSheet = this.sheetsById[newSheetId]; // allow it to work with `.headers` but `.headerValues` is the real prop

                if (!(properties.headerValues || properties.headers)) {
                  _context15.next = 9;
                  break;
                }

                _context15.next = 9;
                return newSheet.setHeaderRow(properties.headerValues || properties.headers);

              case 9:
                return _context15.abrupt("return", newSheet);

              case 10:
              case "end":
                return _context15.stop();
            }
          }
        }, _callee15, this);
      }));

      function addSheet() {
        return _addSheet.apply(this, arguments);
      }

      return addSheet;
    }()
  }, {
    key: "addWorksheet",
    value: function () {
      var _addWorksheet = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee16(properties) {
        return regeneratorRuntime.wrap(function _callee16$(_context16) {
          while (1) {
            switch (_context16.prev = _context16.next) {
              case 0:
                return _context16.abrupt("return", this.addSheet(properties));

              case 1:
              case "end":
                return _context16.stop();
            }
          }
        }, _callee16, this);
      }));

      function addWorksheet(_x15) {
        return _addWorksheet.apply(this, arguments);
      }

      return addWorksheet;
    }() // alias to mimic old version

  }, {
    key: "deleteSheet",
    value: function () {
      var _deleteSheet = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee17(sheetId) {
        return regeneratorRuntime.wrap(function _callee17$(_context17) {
          while (1) {
            switch (_context17.prev = _context17.next) {
              case 0:
                _context17.next = 2;
                return this._makeSingleUpdateRequest('deleteSheet', {
                  sheetId: sheetId
                });

              case 2:
                delete this._rawSheets[sheetId];

              case 3:
              case "end":
                return _context17.stop();
            }
          }
        }, _callee17, this);
      }));

      function deleteSheet(_x16) {
        return _deleteSheet.apply(this, arguments);
      }

      return deleteSheet;
    }() // NAMED RANGES //////////////////////////////////////////////////////////////////////////////////

  }, {
    key: "addNamedRange",
    value: function () {
      var _addNamedRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee18(name, range, namedRangeId) {
        return regeneratorRuntime.wrap(function _callee18$(_context18) {
          while (1) {
            switch (_context18.prev = _context18.next) {
              case 0:
                return _context18.abrupt("return", this._makeSingleUpdateRequest('addNamedRange', {
                  name: name,
                  range: range,
                  namedRangeId: namedRangeId
                }));

              case 1:
              case "end":
                return _context18.stop();
            }
          }
        }, _callee18, this);
      }));

      function addNamedRange(_x17, _x18, _x19) {
        return _addNamedRange.apply(this, arguments);
      }

      return addNamedRange;
    }()
  }, {
    key: "deleteNamedRange",
    value: function () {
      var _deleteNamedRange = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee19(namedRangeId) {
        return regeneratorRuntime.wrap(function _callee19$(_context19) {
          while (1) {
            switch (_context19.prev = _context19.next) {
              case 0:
                return _context19.abrupt("return", this._makeSingleUpdateRequest('deleteNamedRange', {
                  namedRangeId: namedRangeId
                }));

              case 1:
              case "end":
                return _context19.stop();
            }
          }
        }, _callee19, this);
      }));

      function deleteNamedRange(_x20) {
        return _deleteNamedRange.apply(this, arguments);
      }

      return deleteNamedRange;
    }() // LOADING CELLS /////////////////////////////////////////////////////////////////////////////////

  }, {
    key: "loadCells",
    value: function () {
      var _loadCells = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee20(filters) {
        var _this5 = this;

        var readOnlyMode, filtersArray, dataFilters, result, sheets;
        return regeneratorRuntime.wrap(function _callee20$(_context20) {
          while (1) {
            switch (_context20.prev = _context20.next) {
              case 0:
                // you can pass in a single filter or an array of filters
                // strings are treated as a1 ranges
                // objects are treated as GridRange objects
                // TODO: make it support DeveloperMetadataLookup objects
                // TODO: switch to this mode if using a read-only auth token?
                readOnlyMode = this.authMode === AUTH_MODES.API_KEY;
                filtersArray = _.isArray(filters) ? filters : [filters];
                dataFilters = _.map(filtersArray, function (filter) {
                  if (_.isString(filter)) {
                    return readOnlyMode ? filter : {
                      a1Range: filter
                    };
                  }

                  if (_.isObject(filter)) {
                    if (readOnlyMode) {
                      throw new Error('Only A1 ranges are supported when fetching cells with read-only access (using only an API key)');
                    } // TODO: make this support Developer Metadata filters


                    return {
                      gridRange: filter
                    };
                  }

                  throw new Error('Each filter must be an A1 range string or a gridrange object');
                });

                if (!(this.authMode === AUTH_MODES.API_KEY)) {
                  _context20.next = 9;
                  break;
                }

                _context20.next = 6;
                return this.axios.get('/', {
                  params: {
                    includeGridData: true,
                    ranges: dataFilters
                  }
                });

              case 6:
                result = _context20.sent;
                _context20.next = 12;
                break;

              case 9:
                _context20.next = 11;
                return this.axios.post(':getByDataFilter', {
                  includeGridData: true,
                  dataFilters: dataFilters
                });

              case 11:
                result = _context20.sent;

              case 12:
                sheets = result.data.sheets;

                _.each(sheets, function (sheet) {
                  _this5._updateOrCreateSheet(sheet);
                });

              case 14:
              case "end":
                return _context20.stop();
            }
          }
        }, _callee20, this);
      }));

      function loadCells(_x21) {
        return _loadCells.apply(this, arguments);
      }

      return loadCells;
    }()
  }, {
    key: "title",
    get: function get() {
      return this._getProp('title');
    },
    set: function set(newVal) {
      this._setProp('title', newVal);
    }
  }, {
    key: "locale",
    get: function get() {
      return this._getProp('locale');
    },
    set: function set(newVal) {
      this._setProp('locale', newVal);
    }
  }, {
    key: "timeZone",
    get: function get() {
      return this._getProp('timeZone');
    },
    set: function set(newVal) {
      this._setProp('timeZone', newVal);
    }
  }, {
    key: "autoRecalc",
    get: function get() {
      return this._getProp('autoRecalc');
    },
    set: function set(newVal) {
      this._setProp('autoRecalc', newVal);
    }
  }, {
    key: "defaultFormat",
    get: function get() {
      return this._getProp('defaultFormat');
    },
    set: function set(newVal) {
      this._setProp('defaultFormat', newVal);
    }
  }, {
    key: "spreadsheetTheme",
    get: function get() {
      return this._getProp('spreadsheetTheme');
    },
    set: function set(newVal) {
      this._setProp('spreadsheetTheme', newVal);
    }
  }, {
    key: "iterativeCalculationSettings",
    get: function get() {
      return this._getProp('iterativeCalculationSettings');
    },
    set: function set(newVal) {
      this._setProp('iterativeCalculationSettings', newVal);
    }
  }, {
    key: "sheetCount",
    get: function get() {
      this._ensureInfoLoaded();

      return _.values(this._rawSheets).length;
    }
  }, {
    key: "sheetsById",
    get: function get() {
      this._ensureInfoLoaded();

      return this._rawSheets;
    }
  }, {
    key: "sheetsByIndex",
    get: function get() {
      this._ensureInfoLoaded();

      return _.sortBy(this._rawSheets, 'index');
    }
  }, {
    key: "sheetsByTitle",
    get: function get() {
      this._ensureInfoLoaded();

      return _.keyBy(this._rawSheets, 'title');
    }
  }]);

  return GoogleSpreadsheet;
}();

module.exports = GoogleSpreadsheet;