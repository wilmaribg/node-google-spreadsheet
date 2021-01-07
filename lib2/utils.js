'use strict';

var _ = require('lodash');

function getFieldMask(obj) {
  return _.keys(obj).join(',');
}

function columnToLetter(column) {
  var temp = void 0;
  var letter = '';
  var col = column;
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter) {
  var column = 0;
  var length = letter.length;

  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * 26 ** (length - i - 1);
  }
  return column;
}

module.exports = {
  getFieldMask: getFieldMask,
  columnToLetter: columnToLetter,
  letterToColumn: letterToColumn
};