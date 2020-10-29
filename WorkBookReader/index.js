"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports["default"] = void 0;

var _xlsx = _interopRequireDefault(require("xlsx"));

var _helpers = require("../helpers");

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }

function _toConsumableArray(arr) { return _arrayWithoutHoles(arr) || _iterableToArray(arr) || _unsupportedIterableToArray(arr) || _nonIterableSpread(); }

function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }

function _unsupportedIterableToArray(o, minLen) { if (!o) return; if (typeof o === "string") return _arrayLikeToArray(o, minLen); var n = Object.prototype.toString.call(o).slice(8, -1); if (n === "Object" && o.constructor) n = o.constructor.name; if (n === "Map" || n === "Set") return Array.from(o); if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen); }

function _iterableToArray(iter) { if (typeof Symbol !== "undefined" && Symbol.iterator in Object(iter)) return Array.from(iter); }

function _arrayWithoutHoles(arr) { if (Array.isArray(arr)) return _arrayLikeToArray(arr); }

function _arrayLikeToArray(arr, len) { if (len == null || len > arr.length) len = arr.length; for (var i = 0, arr2 = new Array(len); i < len; i++) { arr2[i] = arr[i]; } return arr2; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }

function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); return Constructor; }

function _defineProperty(obj, key, value) { if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable: true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }

var WorkBookReader = function () {
  function WorkBookReader() {
    _classCallCheck(this, WorkBookReader);

    _defineProperty(this, "workBook", void 0);
  }

  _createClass(WorkBookReader, [{
    key: "readFile",
    value: function readFile(filename, opts) {
      this.workBook = _xlsx["default"].readFile(filename, opts);
      return this.workBook;
    }
  }, {
    key: "read",
    value: function read(data, opts) {
      this.workBook = _xlsx["default"].read(data, opts);
      return this.workBook;
    }
  }, {
    key: "setWorkBook",
    value: function setWorkBook(workBook) {
      this.workBook = workBook;
      return this.workBook;
    }
  }, {
    key: "forEachRowEx",
    value: function forEachRowEx(sheetName, cb) {
      var options = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};
      var columnNames = [];
      return this.forEachRow(sheetName, function (row, rowIndex, range) {
        if (rowIndex === range.s.r) {
          columnNames = _toConsumableArray(row);

          if (options.getModifiedColumnNames) {
            columnNames = options.getModifiedColumnNames(columnNames);
          }

          return;
        }

        var r = {};
        row.forEach(function (c, i) {
          return r[columnNames[i]] = c;
        });
        return cb(r, rowIndex, range);
      });
    }
  }, {
    key: "forEachRow",
    value: function forEachRow(sheetName, cb) {
      var ws = this.workBook.Sheets[sheetName];

      if (!ws) {
        return new Error("sheet not found: ".concat(sheetName));
      }

      var ref = ws['!ref'];

      var range = _xlsx["default"].utils.decode_range(ref);

      var columnSize = range.e.c - range.s.c + 1;

      for (var r = range.s.r; r <= range.e.r; ++r) {
        var _row = Array.from({
          length: columnSize
        });

        for (var c = range.s.c; c <= range.e.c; ++c) {
          var cell_address = {
            c: c,
            r: r
          };

          var cell_ref = _xlsx["default"].utils.encode_cell(cell_address);

          var cell = ws[cell_ref];
          _row[c - range.s.c] = cell && cell.v;
        }

        var keepGoing = cb(_row, r, range);

        if (keepGoing === false) {
          break;
        }
      }
    }
  }, {
    key: "test",
    value: function test() {
      var sheetIndex = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : 0;
      var wb = this.workBook;
      console.log('wb.SheetNames :', wb.SheetNames);
      var ws = wb.Sheets[wb.SheetNames[sheetIndex]];
      console.log('ws["!ref"] :', ws['!ref']);
      var ref = ws['!ref'];
      console.log('ref :', ref);
      var columnSymbols = (0, _helpers.getColumnSymbols)(ref);
      console.log('columnSymbols :', columnSymbols);

      var data = _xlsx["default"].utils.sheet_to_json(ws);

      columnSymbols.forEach(function (columnSymbol) {});
      this.forEachRowEx(wb.SheetNames[sheetIndex], function (row) {
        console.log('row :', row);
      }, {
        getModifiedColumnNames: function getModifiedColumnNames(cols) {
          cols[cols.length - 2] = 'sizes';
          cols[cols.length - 1] = 'fits';
          return cols;
        }
      });
    }
  }]);

  return WorkBookReader;
}();

exports["default"] = WorkBookReader;