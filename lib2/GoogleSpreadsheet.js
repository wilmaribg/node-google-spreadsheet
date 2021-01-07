'use strict';

var _extends = Object.assign || function (target) { for (var i = 1; i < arguments.length; i++) { var source = arguments[i]; for (var key in source) { if (Object.prototype.hasOwnProperty.call(source, key)) { target[key] = source[key]; } } } return target; };

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _defineProperty(obj, key, value) { if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable: true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require('lodash');

var _require = require('google-auth-library'),
    JWT = _require.JWT;

var Axios = require('axios');

var GoogleSpreadsheetWorksheet = require('./GoogleSpreadsheetWorksheet');

var _require2 = require('./utils'),
    getFieldMask = _require2.getFieldMask;

var GOOGLE_AUTH_SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

var AUTH_MODES = {
  JWT: 'JWT',
  API_KEY: 'API_KEY',
  RAW_ACCESS_TOKEN: 'RAW_ACCESS_TOKEN',
  OAUTH: 'OAUTH'
};

var GoogleSpreadsheet = function () {
  function GoogleSpreadsheet(sheetId) {
    _classCallCheck(this, GoogleSpreadsheet);

    this.spreadsheetId = sheetId;
    this.authMode = null;
    this._rawSheets = {};
    this._rawProperties = null;

    // create an axios instance with sheet root URL and interceptors to handle auth
    this.axios = Axios.create({
      baseURL: 'https://sheets.googleapis.com/v4/spreadsheets/' + (sheetId || ''),
      // send arrays in params with duplicate keys - ie `?thing=1&thing=2` vs `?thing[]=1...`
      // solution taken from https://github.com/axios/axios/issues/604
      paramsSerializer: function paramsSerializer(params) {
        var options = '';
        _.keys(params).forEach(function (key) {
          var isParamTypeObject = _typeof(params[key]) === 'object';
          var isParamTypeArray = isParamTypeObject && params[key].length >= 0;
          if (!isParamTypeObject) options += key + '=' + encodeURIComponent(params[key]) + '&';
          if (isParamTypeObject && isParamTypeArray) {
            _.each(params[key], function (val) {
              options += key + '=' + encodeURIComponent(val) + '&';
            });
          }
        });
        return options ? options.slice(0, -1) : options;
      }
    });
    // have to use bind here or the functions dont have access to `this` :(
    this.axios.interceptors.request.use(this._setAxiosRequestAuth.bind(this));
    this.axios.interceptors.response.use(this._handleAxiosResponse.bind(this), this._handleAxiosErrors.bind(this));

    return this;
  }

  // CREATE NEW DOC ////////////////////////////////////////////////////////////////////////////////


  _createClass(GoogleSpreadsheet, [{
    key: 'createNewSpreadsheetDocument',
    value: async function createNewSpreadsheetDocument(properties) {
      var _this = this;

      // see updateProperties for more info about available properties

      if (this.spreadsheetId) {
        throw new Error('Only call `createNewSpreadsheetDocument()` on a GoogleSpreadsheet object that has no spreadsheetId set');
      }
      var response = await this.axios.post(this.url, {
        properties: properties
      });
      this.spreadsheetId = response.data.spreadsheetId;
      this.axios.defaults.baseURL += this.spreadsheetId;

      this._rawProperties = response.data.properties;
      _.each(response.data.sheets, function (s) {
        return _this._updateOrCreateSheet(s);
      });
    }

    // AUTH RELATED FUNCTIONS ////////////////////////////////////////////////////////////////////////

  }, {
    key: 'useApiKey',
    value: async function useApiKey(key) {
      this.authMode = AUTH_MODES.API_KEY;
      this.apiKey = key;
    }

    // token must be created and managed (refreshed) elsewhere

  }, {
    key: 'useRawAccessToken',
    value: async function useRawAccessToken(token) {
      this.authMode = AUTH_MODES.RAW_ACCESS_TOKEN;
      this.accessToken = token;
    }
  }, {
    key: 'useOAuth2Client',
    value: async function useOAuth2Client(oAuth2Client) {
      this.authMode = AUTH_MODES.OAUTH;
      this.oAuth2Client = oAuth2Client;
    }

    // creds should be an object obtained by loading the json file google gives you
    // impersonateAs is an email of any user in the G Suite domain
    // (only works if service account has domain-wide delegation enabled)

  }, {
    key: 'useServiceAccountAuth',
    value: async function useServiceAccountAuth(creds) {
      var impersonateAs = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;

      this.jwtClient = new JWT({
        email: creds.client_email,
        key: creds.private_key,
        scopes: GOOGLE_AUTH_SCOPES,
        subject: impersonateAs
      });
      await this.renewJwtAuth();
    }
  }, {
    key: 'renewJwtAuth',
    value: async function renewJwtAuth() {
      this.authMode = AUTH_MODES.JWT;
      await this.jwtClient.authorize();
      /*
      returned token looks like
        {
          access_token: 'secret-token...',
          token_type: 'Bearer',
          expiry_date: 1576005020000,
          id_token: undefined,
          refresh_token: 'jwt-placeholder'
        }
      */
    }

    // TODO: provide mechanism to share single JWT auth between docs?

    // INTERNAL UTILITY FUNCTIONS ////////////////////////////////////////////////////////////////////

  }, {
    key: '_setAxiosRequestAuth',
    value: async function _setAxiosRequestAuth(config) {
      // TODO: check auth mode, if valid, renew if expired, etc
      if (this.authMode === AUTH_MODES.JWT) {
        if (!this.jwtClient) throw new Error('JWT auth is not set up properly');
        // this seems to do the right thing and only renew the token if expired
        await this.jwtClient.authorize();
        config.headers.Authorization = 'Bearer ' + this.jwtClient.credentials.access_token;
      } else if (this.authMode === AUTH_MODES.RAW_ACCESS_TOKEN) {
        if (!this.accessToken) throw new Error('Invalid access token');
        config.headers.Authorization = 'Bearer ' + this.accessToken;
      } else if (this.authMode === AUTH_MODES.API_KEY) {
        if (!this.apiKey) throw new Error('Please set API key');
        config.params = config.params || {};
        config.params.key = this.apiKey;
      } else if (this.authMode === AUTH_MODES.OAUTH) {
        var credentials = await this.oAuth2Client.getAccessToken();
        config.headers.Authorization = 'Bearer ' + credentials.token;
      } else {
        throw new Error('You must initialize some kind of auth before making any requests');
      }
      return config;
    }
  }, {
    key: '_handleAxiosResponse',
    value: async function _handleAxiosResponse(response) {
      return response;
    }
  }, {
    key: '_handleAxiosErrors',
    value: async function _handleAxiosErrors(error) {
      // console.log(error);
      if (error.response && error.response.data) {
        // usually the error has a code and message, but occasionally not
        if (!error.response.data.error) throw error;

        var _error$response$data$ = error.response.data.error,
            code = _error$response$data$.code,
            message = _error$response$data$.message;

        error.message = 'Google API error - [' + code + '] ' + message;
        throw error;
      }

      if (_.get(error, 'response.status') === 403) {
        if (this.authMode === AUTH_MODES.API_KEY) {
          throw new Error('Sheet is private. Use authentication or make public. (see https://github.com/theoephraim/node-google-spreadsheet#a-note-on-authentication for details)');
        }
      }
      throw error;
    }
  }, {
    key: '_makeSingleUpdateRequest',
    value: async function _makeSingleUpdateRequest(requestType, requestParams) {
      var _this2 = this;

      var response = await this.axios.post(':batchUpdate', {
        requests: [_defineProperty({}, requestType, requestParams)],
        includeSpreadsheetInResponse: true
        // responseRanges: [string]
        // responseIncludeGridData: true
      });

      this._updateRawProperties(response.data.updatedSpreadsheet.properties);
      _.each(response.data.updatedSpreadsheet.sheets, function (s) {
        return _this2._updateOrCreateSheet(s);
      });
      // console.log('API RESPONSE', response.data.replies[0][requestType]);
      return response.data.replies[0][requestType];
    }
  }, {
    key: '_makeBatchUpdateRequest',
    value: async function _makeBatchUpdateRequest(requests, responseRanges) {
      var _this3 = this;

      // this is used for updating batches of cells
      var response = await this.axios.post(':batchUpdate', _extends({
        requests: requests,
        includeSpreadsheetInResponse: true
      }, responseRanges && _extends({
        responseIncludeGridData: true
      }, responseRanges !== '*' && { responseRanges: responseRanges })));

      this._updateRawProperties(response.data.updatedSpreadsheet.properties);
      _.each(response.data.updatedSpreadsheet.sheets, function (s) {
        return _this3._updateOrCreateSheet(s);
      });
    }
  }, {
    key: '_ensureInfoLoaded',
    value: function _ensureInfoLoaded() {
      if (!this._rawProperties) throw new Error('You must call `doc.loadInfo()` before accessing this property');
    }
  }, {
    key: '_updateRawProperties',
    value: function _updateRawProperties(newProperties) {
      this._rawProperties = newProperties;
    }
  }, {
    key: '_updateOrCreateSheet',
    value: function _updateOrCreateSheet(_ref2) {
      var properties = _ref2.properties,
          data = _ref2.data;
      var sheetId = properties.sheetId;

      if (!this._rawSheets[sheetId]) {
        this._rawSheets[sheetId] = new GoogleSpreadsheetWorksheet(this, { properties: properties, data: data });
      } else {
        this._rawSheets[sheetId]._rawProperties = properties;
        this._rawSheets[sheetId]._fillCellData(data);
      }
    }

    // BASIC PROPS //////////////////////////////////////////////////////////////////////////////

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
    key: 'updateProperties',
    value: async function updateProperties(properties) {
      // updateSpreadsheetProperties
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#SpreadsheetProperties

      /*
        title (string) - title of the spreadsheet
        locale (string) - ISO code
        autoRecalc (enum) - ON_CHANGE|MINUTE|HOUR
        timeZone (string) - timezone code
        iterativeCalculationSettings (object) - see https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#IterativeCalculationSettings
       */

      await this._makeSingleUpdateRequest('updateSpreadsheetProperties', {
        properties: properties,
        fields: getFieldMask(properties)
      });
    }

    // BASIC INFO ////////////////////////////////////////////////////////////////////////////////////

  }, {
    key: 'loadInfo',
    value: async function loadInfo(includeCells) {
      var _this4 = this;

      var response = await this.axios.get('/', {
        params: _extends({}, includeCells && { includeGridData: true })
      });
      this._rawProperties = response.data.properties;
      _.each(response.data.sheets, function (s) {
        return _this4._updateOrCreateSheet(s);
      });
    }
  }, {
    key: 'getInfo',
    value: async function getInfo() {
      return this.loadInfo();
    } // alias to mimic old version

  }, {
    key: 'resetLocalCache',
    value: function resetLocalCache() {
      this._rawProperties = null;
      this._rawSheets = {};
    }

    // WORKSHEETS ////////////////////////////////////////////////////////////////////////////////////

  }, {
    key: 'addSheet',
    value: async function addSheet() {
      var properties = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};

      // Request type = `addSheet`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddSheetRequest

      var response = await this._makeSingleUpdateRequest('addSheet', {
        properties: _.omit(properties, 'headers', 'headerValues')
      });
      // _makeSingleUpdateRequest already adds the sheet
      var newSheetId = response.properties.sheetId;
      var newSheet = this.sheetsById[newSheetId];

      // allow it to work with `.headers` but `.headerValues` is the real prop
      if (properties.headerValues || properties.headers) {
        await newSheet.setHeaderRow(properties.headerValues || properties.headers);
      }

      return newSheet;
    }
  }, {
    key: 'addWorksheet',
    value: async function addWorksheet(properties) {
      return this.addSheet(properties);
    } // alias to mimic old version

  }, {
    key: 'deleteSheet',
    value: async function deleteSheet(sheetId) {
      // Request type = `deleteSheet`
      // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteSheetRequest
      await this._makeSingleUpdateRequest('deleteSheet', { sheetId: sheetId });
      delete this._rawSheets[sheetId];
    }

    // NAMED RANGES //////////////////////////////////////////////////////////////////////////////////

  }, {
    key: 'addNamedRange',
    value: async function addNamedRange(name, range, namedRangeId) {
      // namedRangeId is optional
      return this._makeSingleUpdateRequest('addNamedRange', {
        name: name,
        range: range,
        namedRangeId: namedRangeId
      });
    }
  }, {
    key: 'deleteNamedRange',
    value: async function deleteNamedRange(namedRangeId) {
      return this._makeSingleUpdateRequest('deleteNamedRange', { namedRangeId: namedRangeId });
    }

    // LOADING CELLS /////////////////////////////////////////////////////////////////////////////////

  }, {
    key: 'loadCells',
    value: async function loadCells(filters) {
      var _this5 = this;

      // you can pass in a single filter or an array of filters
      // strings are treated as a1 ranges
      // objects are treated as GridRange objects
      // TODO: make it support DeveloperMetadataLookup objects

      // TODO: switch to this mode if using a read-only auth token?
      var readOnlyMode = this.authMode === AUTH_MODES.API_KEY;

      var filtersArray = _.isArray(filters) ? filters : [filters];
      var dataFilters = _.map(filtersArray, function (filter) {
        if (_.isString(filter)) {
          return readOnlyMode ? filter : { a1Range: filter };
        }
        if (_.isObject(filter)) {
          if (readOnlyMode) {
            throw new Error('Only A1 ranges are supported when fetching cells with read-only access (using only an API key)');
          }
          // TODO: make this support Developer Metadata filters
          return { gridRange: filter };
        }
        throw new Error('Each filter must be an A1 range string or a gridrange object');
      });

      var result = void 0;
      // when using an API key only, we must use the regular get endpoint
      // because :getByDataFilter requires higher access
      if (this.authMode === AUTH_MODES.API_KEY) {
        result = await this.axios.get('/', {
          params: {
            includeGridData: true,
            ranges: dataFilters
          }
        });
        // otherwise we use the getByDataFilter endpoint because it is more flexible
      } else {
        result = await this.axios.post(':getByDataFilter', {
          includeGridData: true,
          dataFilters: dataFilters
        });
      }

      var sheets = result.data.sheets;

      _.each(sheets, function (sheet) {
        _this5._updateOrCreateSheet(sheet);
      });
    }
  }, {
    key: 'title',
    get: function get() {
      return this._getProp('title');
    },
    set: function set(newVal) {
      this._setProp('title', newVal);
    }
  }, {
    key: 'locale',
    get: function get() {
      return this._getProp('locale');
    },
    set: function set(newVal) {
      this._setProp('locale', newVal);
    }
  }, {
    key: 'timeZone',
    get: function get() {
      return this._getProp('timeZone');
    },
    set: function set(newVal) {
      this._setProp('timeZone', newVal);
    }
  }, {
    key: 'autoRecalc',
    get: function get() {
      return this._getProp('autoRecalc');
    },
    set: function set(newVal) {
      this._setProp('autoRecalc', newVal);
    }
  }, {
    key: 'defaultFormat',
    get: function get() {
      return this._getProp('defaultFormat');
    },
    set: function set(newVal) {
      this._setProp('defaultFormat', newVal);
    }
  }, {
    key: 'spreadsheetTheme',
    get: function get() {
      return this._getProp('spreadsheetTheme');
    },
    set: function set(newVal) {
      this._setProp('spreadsheetTheme', newVal);
    }
  }, {
    key: 'iterativeCalculationSettings',
    get: function get() {
      return this._getProp('iterativeCalculationSettings');
    },
    set: function set(newVal) {
      this._setProp('iterativeCalculationSettings', newVal);
    }
  }, {
    key: 'sheetCount',
    get: function get() {
      this._ensureInfoLoaded();
      return _.values(this._rawSheets).length;
    }
  }, {
    key: 'sheetsById',
    get: function get() {
      this._ensureInfoLoaded();
      return this._rawSheets;
    }
  }, {
    key: 'sheetsByIndex',
    get: function get() {
      this._ensureInfoLoaded();
      return _.sortBy(this._rawSheets, 'index');
    }
  }, {
    key: 'sheetsByTitle',
    get: function get() {
      this._ensureInfoLoaded();
      return _.keyBy(this._rawSheets, 'title');
    }
  }]);

  return GoogleSpreadsheet;
}();

module.exports = GoogleSpreadsheet;