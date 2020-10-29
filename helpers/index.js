"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.getColumnSymbols = void 0;

var _xlsx = _interopRequireDefault(require("xlsx"));

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }

var getColumnSymbols = function getColumnSymbols(ref) {
  var columnCount = _xlsx["default"].utils.decode_range(ref).e.c + 1;
  var columnSymbols = Array.from({
    length: columnCount
  });

  for (var index = 0; index < columnCount; index++) {
    columnSymbols[index] = _xlsx["default"].utils.encode_col(index);
  }

  return columnSymbols;
};

exports.getColumnSymbols = getColumnSymbols;