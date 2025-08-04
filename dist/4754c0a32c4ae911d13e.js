function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _regeneratorValues(e) { if (null != e) { var t = e["function" == typeof Symbol && Symbol.iterator || "@@iterator"], r = 0; if (t) return t.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) return { next: function next() { return e && r >= e.length && (e = void 0), { value: e && e[r++], done: !e }; } }; } throw new TypeError(_typeof(e) + " is not iterable"); }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { if (r) i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n;else { var o = function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); }; o("next", 0), o("throw", 1), o("return", 2); } }, _regeneratorDefine2(e, r, n, t); }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
import { handleCellChange } from '../display/display.js';
import { openSettings } from '../settings/settings.js';
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("generate-ordering-report").onclick = function () {
      return tryCatch(generateOrderingReport);
    };
    document.getElementById("generate-inventory-report").onclick = function () {
      return tryCatch(generateInventoryReport);
    };
    document.getElementById('start-date').addEventListener('input', checkDatesAndClearMessage);
    document.getElementById('end-date').addEventListener('input', checkDatesAndClearMessage);
    document.getElementById("order-filtering").addEventListener('change', filteringDropdown);
    document.getElementById("inventory-filtering").addEventListener('change', invFilteringDropdown);
    document.getElementById("settings-button").onclick = function () {
      return tryCatch(openSettings);
    };
  }
});

// Global Variable inits     
var filter = [];
var invFilter = [];
var earlyDateMap = new Map();
var startDateMap = new Map();
var orderingWorksheet;
var orderingTable;
var outputJobs = new Set();
var allData = [];
export var matchingData = [];
function generateOrderingReport() {
  return _generateOrderingReport.apply(this, arguments);
}
function _generateOrderingReport() {
  _generateOrderingReport = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2() {
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          _context2.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(context) {
              var startDateValue, endDateValue;
              return _regenerator().w(function (_context) {
                while (1) switch (_context.n) {
                  case 0:
                    resetGenerateOrdering();
                    _context.n = 1;
                    return context.sync();
                  case 1:
                    orderingWorksheet = context.workbook.worksheets.add("Ordering");
                    orderingTable = orderingWorksheet.tables.add("A1:G1", true);
                    orderingTable.name = "OrderingTable";
                    orderingTable.getHeaderRowRange().values = [["Case #", "Demand", "Current Inventory", "On Order", "Required Amount", "Buy or Make", "Earliest Start Date"]];
                    orderingTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
                    orderingTable.getRange().format.autofitColumns();
                    orderingTable.getRange().format.autofitRows();
                    startDateValue = document.getElementById('start-date').value;
                    endDateValue = document.getElementById('end-date').value;
                    if (!(!startDateValue || !endDateValue)) {
                      _context.n = 2;
                      break;
                    }
                    document.getElementById('message-area').textContent = "Please enter the dates";
                    return _context.a(2);
                  case 2:
                    document.getElementById('message-area').textContent = " ";
                    dateFilter();
                    _context.n = 3;
                    return context.sync();
                  case 3:
                    importColumnData();
                    _context.n = 4;
                    return context.sync();
                  case 4:
                    orderingWorksheet.onSelectionChanged.add(displayData);
                    _context.n = 5;
                    return context.sync();
                  case 5:
                    return _context.a(2);
                }
              }, _callee);
            }));
            return function (_x4) {
              return _ref.apply(this, arguments);
            };
          }());
        case 1:
          return _context2.a(2);
      }
    }, _callee2);
  }));
  return _generateOrderingReport.apply(this, arguments);
}
function generateInventoryReport() {
  return _generateInventoryReport.apply(this, arguments);
}
function _generateInventoryReport() {
  _generateInventoryReport = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4() {
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.n) {
        case 0:
          _context4.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref2 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3(context) {
              var inventoryWorksheet, inventoryTable, inventoryReadout, inventoryReadoutTable, startDateValue, endDateValue;
              return _regenerator().w(function (_context3) {
                while (1) switch (_context3.n) {
                  case 0:
                    resetGenerateInventory();
                    _context3.n = 1;
                    return context.sync();
                  case 1:
                    inventoryWorksheet = context.workbook.worksheets.add("Inventory At");
                    inventoryTable = inventoryWorksheet.tables.add("A1:J1", true);
                    inventoryReadout = context.workbook.worksheets.add("Inventory Request");
                    inventoryReadoutTable = inventoryReadout.tables.add("A1:D1", true);
                    inventoryTable.name = "InventoryAtTable";
                    inventoryTable.style = "TableStyleMedium10";
                    inventoryReadoutTable.name = "InventoryReadoutTable";
                    inventoryReadoutTable.style = "TableStyleMedium10";
                    inventoryReadoutTable.getHeaderRowRange().values = [["Item Code", "Inventory Date", "Inventory Ref", "Inventory Qty"]];
                    inventoryTable.getHeaderRowRange().values = [["Case #", "Demand", "Qty MEB", "Qty EFW", "Total MEB + EFW", "On Order", "Start Date", "Release Date", "Qty Needed (MEB)", "Notes"]];
                    inventoryTable.columns.getItemAt(2).getRange().numberFormat = [["\u20AC#,##0.00"]];
                    inventoryTable.getRange().format.autofitColumns();
                    inventoryTable.getRange().format.autofitRows();
                    startDateValue = document.getElementById('start-date').value;
                    endDateValue = document.getElementById('end-date').value;
                    if (!(!startDateValue || !endDateValue)) {
                      _context3.n = 2;
                      break;
                    }
                    document.getElementById('message-area').textContent = "Please enter the dates";
                    return _context3.a(2);
                  case 2:
                    document.getElementById('message-area').textContent = " ";
                    otherDateFilter();
                    _context3.n = 3;
                    return context.sync();
                  case 3:
                    importOtherColumnData();
                    _context3.n = 4;
                    return context.sync();
                  case 4:
                    _context3.n = 5;
                    return importColumnData();
                  case 5:
                    readoutData();
                  case 6:
                    _context3.n = 7;
                    return context.sync();
                  case 7:
                    return _context3.a(2);
                }
              }, _callee3);
            }));
            return function (_x5) {
              return _ref2.apply(this, arguments);
            };
          }());
        case 1:
          return _context4.a(2);
      }
    }, _callee4);
  }));
  return _generateInventoryReport.apply(this, arguments);
}
function tryCatch(_x) {
  return _tryCatch.apply(this, arguments);
}
function _tryCatch() {
  _tryCatch = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5(callback) {
    var messageArea, _t;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.n) {
        case 0:
          _context5.p = 0;
          _context5.n = 1;
          return callback();
        case 1:
          _context5.n = 3;
          break;
        case 2:
          _context5.p = 2;
          _t = _context5.v;
          messageArea = document.getElementById('message-area');
          if (messageArea) {
            messageArea.textContent = "Error: ".concat(_t.message || _t);
            messageArea.style.color = 'red';
            setTimeout(function () {
              messageArea.textContent = '';
              messageArea.style.color = '';
            }, 5000);
          }
          console.error(_t);
        case 3:
          return _context5.a(2);
      }
    }, _callee5, null, [[0, 2]]);
  }));
  return _tryCatch.apply(this, arguments);
}
function importColumnData() {
  return _importColumnData.apply(this, arguments);
}
function _importColumnData() {
  _importColumnData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee7() {
    return _regenerator().w(function (_context7) {
      while (1) switch (_context7.n) {
        case 0:
          _context7.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref3 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6(context) {
              var inventoryReportWorksheet, inventoryUsedRange, dynamicWorksheet, dynamicUsedRange, openPOsWorksheet, openPOsUsedRange, orderingWorksheet, dynamicHeaders, dynItemCodeIdx, dynItemQtyIdx, dynWorkIdx, dynItemCodeColumn, dynItemQtyColumn, dynWorkColumn, openPOsHeaders, openPOItemCodeIdx, openPOItemQtyIdx, openPOItemCodeColumn, openPOItemQtyColumn, inventoryHeaders, invItemCodeIdx, invItemQtyIdx, invRepItemCodeColumn, invRepItemQtyColumn, dynamic, dynamicQR, dynamicICR, dynamicWork, inventoryICR, inventoryQR, openPOs, openPOsICR, openPOsQR, filteredICR, filteredQR, fullDynamicMap, dynamicMap, inventoryMap, openPOsMap, allItemCodes, result, _iterator, _step, _code, dynamicQty, inventoryQty, openPOsQty, toOrder, caseNumbers, requiredAmounts, sell, _iterator2, _step2, _code2, _dynamicQty, _openPOsQty, overBuy, orderingUsedRange, orderingValues, startArray, i, itemCode, start, dateOnly, caseOrder, demand, _iterator3, _step3, _code3, demandQty, demandOutput, currentInventory, _iterator4, _step4, _code4, currentInvQty, currentInventoryOutput, onOrder, _iterator5, _step5, _code5, onOrderQty, onOrderOutput, orderOrMakeMap, _i, code, work, orderOrMake, _iterator6, _step6, _code6, workCentersSet, _workCenters, orderOrMakeOutput, orderOrMakeCategory, _i2, workCenters, usedRange, borders, lastRow, highlight;
              return _regenerator().w(function (_context6) {
                while (1) switch (_context6.n) {
                  case 0:
                    inventoryReportWorksheet = context.workbook.worksheets.getItem("Inventory");
                    inventoryUsedRange = inventoryReportWorksheet.getUsedRange().load("values");
                    dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
                    dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
                    openPOsWorksheet = context.workbook.worksheets.getItem("Open PO's");
                    openPOsUsedRange = openPOsWorksheet.getUsedRange().load("values");
                    orderingWorksheet = context.workbook.worksheets.getItem("Ordering");
                    _context6.n = 1;
                    return context.sync();
                  case 1:
                    //Dynamic fluid Placement
                    dynamicHeaders = dynamicUsedRange.values[0];
                    dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
                    dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
                    dynWorkIdx = dynamicHeaders.indexOf("Work Center");
                    dynItemCodeColumn = "".concat(colIdxToLetter(dynItemCodeIdx), ":").concat(colIdxToLetter(dynItemCodeIdx));
                    dynItemQtyColumn = "".concat(colIdxToLetter(dynItemQtyIdx), ":").concat(colIdxToLetter(dynItemQtyIdx));
                    dynWorkColumn = "".concat(colIdxToLetter(dynWorkIdx), ":").concat(colIdxToLetter(dynWorkIdx)); //Open PO's fluid Placement
                    openPOsHeaders = openPOsUsedRange.values[0];
                    openPOItemCodeIdx = openPOsHeaders.indexOf("Item Code");
                    openPOItemQtyIdx = openPOsHeaders.indexOf("Outstanding Qty");
                    openPOItemCodeColumn = "".concat(colIdxToLetter(openPOItemCodeIdx), ":").concat(colIdxToLetter(openPOItemCodeIdx));
                    openPOItemQtyColumn = "".concat(colIdxToLetter(openPOItemQtyIdx), ":").concat(colIdxToLetter(openPOItemQtyIdx)); //Inventory Report Fluid Placement
                    inventoryHeaders = inventoryUsedRange.values[0];
                    invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
                    invItemQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
                    invRepItemCodeColumn = "".concat(colIdxToLetter(invItemCodeIdx), ":").concat(colIdxToLetter(invItemCodeIdx));
                    invRepItemQtyColumn = "".concat(colIdxToLetter(invItemQtyIdx), ":").concat(colIdxToLetter(invItemQtyIdx)); //Quanity and Item Code from Dynamic, Inventory Report, and Open PO's sheets
                    dynamic = context.workbook.worksheets.getItem("Dynamic");
                    dynamicQR = dynamic.getRange(dynItemQtyColumn).getUsedRange().load("values");
                    dynamicICR = dynamic.getRange(dynItemCodeColumn).getUsedRange().load("values");
                    dynamicWork = dynamic.getRange(dynWorkColumn).getUsedRange().load("values");
                    _context6.n = 2;
                    return context.sync();
                  case 2:
                    inventoryICR = inventoryReportWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values");
                    inventoryQR = inventoryReportWorksheet.getRange(invRepItemQtyColumn).getUsedRange().load("values");
                    openPOs = context.workbook.worksheets.getItem("Open PO's");
                    openPOsICR = openPOs.getRange(openPOItemCodeColumn).getUsedRange().load("values");
                    openPOsQR = openPOs.getRange(openPOItemQtyColumn).getUsedRange().load("values");
                    _context6.n = 3;
                    return context.sync();
                  case 3:
                    //Date Filtering
                    filteredICR = filter.map(function (item) {
                      return [item.itemCode];
                    });
                    filteredQR = filter.map(function (item) {
                      return [item.qty];
                    }); //Sum Map Building
                    fullDynamicMap = buildSumMap(dynamicICR.values, dynamicQR.values);
                    dynamicMap = buildSumMap(filteredICR, filteredQR);
                    inventoryMap = buildSumMap(inventoryICR.values, inventoryQR.values);
                    openPOsMap = buildSumMap(openPOsICR.values, openPOsQR.values);
                    allItemCodes = new Set([].concat(_toConsumableArray(dynamicMap.keys()), _toConsumableArray(inventoryMap.keys()), _toConsumableArray(openPOsMap.keys())));
                    result = [["Case #", "Required Amount"]];
                    _iterator = _createForOfIteratorHelper(allItemCodes);
                    try {
                      for (_iterator.s(); !(_step = _iterator.n()).done;) {
                        _code = _step.value;
                        dynamicQty = dynamicMap.get(_code) || 0;
                        inventoryQty = inventoryMap.get(_code) || 0;
                        openPOsQty = openPOsMap.get(_code) || 0;
                        toOrder = dynamicQty - inventoryQty - openPOsQty;
                        if (toOrder > 0) {
                          result.push([_code, toOrder]);
                        }
                      }
                    } catch (err) {
                      _iterator.e(err);
                    } finally {
                      _iterator.f();
                    }
                    caseNumbers = result.map(function (row) {
                      return [row[0]];
                    });
                    orderingWorksheet.getRange("A1:A".concat(caseNumbers.length)).values = caseNumbers;
                    requiredAmounts = result.map(function (row) {
                      return [row[1]];
                    });
                    orderingWorksheet.getRange("E1:E".concat(requiredAmounts.length)).values = requiredAmounts;
                    _context6.n = 4;
                    return context.sync();
                  case 4:
                    // Remove From Order
                    sell = [["Case #", "Remove From Order"]];
                    _iterator2 = _createForOfIteratorHelper(allItemCodes);
                    try {
                      for (_iterator2.s(); !(_step2 = _iterator2.n()).done;) {
                        _code2 = _step2.value;
                        _dynamicQty = Number(fullDynamicMap.get(_code2)) || 0;
                        _openPOsQty = Number(openPOsMap.get(_code2)) || 0;
                        overBuy = _openPOsQty - _dynamicQty;
                        if (!isNaN(_dynamicQty) && !isNaN(_openPOsQty)) {
                          if (String(_code2).includes("COR") && _openPOsQty > _dynamicQty) {
                            sell.push([_code2, overBuy]);
                          }
                        }
                      }
                    } catch (err) {
                      _iterator2.e(err);
                    } finally {
                      _iterator2.f();
                    }
                    console.log("Sell these", sell);

                    //Importing the Planned Start Date
                    orderingUsedRange = orderingWorksheet.getUsedRange().load("values");
                    _context6.n = 5;
                    return context.sync();
                  case 5:
                    orderingValues = orderingUsedRange.values;
                    startArray = [["Earliest Start Date"]];
                    for (i = 1; i < orderingValues.length; i++) {
                      itemCode = String(orderingValues[i][0]).trim();
                      start = String(earlyDateMap.get(itemCode)) || "No Start Date Established";
                      dateOnly = start.split(' ').slice(0, 4).join(' ');
                      startArray.push([dateOnly]);
                    }
                    orderingWorksheet.getRange("G1:G".concat(startArray.length)).values = startArray;

                    //Importing Demand, Current Inventory, and On Order
                    caseOrder = result.map(function (row) {
                      return row[0];
                    });
                    demand = [["Demand"]];
                    _iterator3 = _createForOfIteratorHelper(caseOrder);
                    try {
                      for (_iterator3.s(); !(_step3 = _iterator3.n()).done;) {
                        _code3 = _step3.value;
                        demandQty = dynamicMap.get(_code3) || 0;
                        if (demandQty > 0) {
                          demand.push([demandQty]);
                        }
                      }
                    } catch (err) {
                      _iterator3.e(err);
                    } finally {
                      _iterator3.f();
                    }
                    demandOutput = demand.map(function (row) {
                      return [row[0]];
                    });
                    orderingWorksheet.getRange("B1:B".concat(demandOutput.length)).values = demandOutput;
                    currentInventory = [["Current Inventory"]];
                    _iterator4 = _createForOfIteratorHelper(caseOrder.slice(1));
                    try {
                      for (_iterator4.s(); !(_step4 = _iterator4.n()).done;) {
                        _code4 = _step4.value;
                        currentInvQty = inventoryMap.get(_code4) || 0;
                        currentInventory.push([currentInvQty]);
                      }
                    } catch (err) {
                      _iterator4.e(err);
                    } finally {
                      _iterator4.f();
                    }
                    currentInventoryOutput = currentInventory.map(function (row) {
                      return [row[0]];
                    });
                    orderingWorksheet.getRange("C1:C".concat(currentInventoryOutput.length)).values = currentInventoryOutput;
                    onOrder = [["On Order"]];
                    _iterator5 = _createForOfIteratorHelper(caseOrder.slice(1));
                    try {
                      for (_iterator5.s(); !(_step5 = _iterator5.n()).done;) {
                        _code5 = _step5.value;
                        onOrderQty = openPOsMap.get(_code5) || 0;
                        onOrder.push([onOrderQty]);
                      }
                    } catch (err) {
                      _iterator5.e(err);
                    } finally {
                      _iterator5.f();
                    }
                    onOrderOutput = onOrder.map(function (row) {
                      return [row[0]];
                    });
                    orderingWorksheet.getRange("D1:D".concat(onOrderOutput.length)).values = onOrderOutput;

                    //Buy or Make Logic
                    orderOrMakeMap = new Map();
                    for (_i = 1; _i < dynamicICR.values.length; _i++) {
                      code = String(dynamicICR.values[_i][0]).trim();
                      work = dynamicWork.values[_i] ? String(dynamicWork.values[_i][0]).trim() : "";
                      if (code && work) {
                        if (!orderOrMakeMap.has(code)) {
                          orderOrMakeMap.set(code, new Set());
                        }
                        orderOrMakeMap.get(code).add(work);
                      }
                    }
                    _context6.n = 6;
                    return context.sync();
                  case 6:
                    orderOrMake = [["Buy or Make"]];
                    _iterator6 = _createForOfIteratorHelper(caseOrder.slice(1));
                    try {
                      for (_iterator6.s(); !(_step6 = _iterator6.n()).done;) {
                        _code6 = _step6.value;
                        workCentersSet = orderOrMakeMap.get(_code6);
                        _workCenters = workCentersSet ? Array.from(workCentersSet).join(", ") : "";
                        orderOrMake.push([_workCenters]);
                      }
                    } catch (err) {
                      _iterator6.e(err);
                    } finally {
                      _iterator6.f();
                    }
                    orderOrMakeOutput = orderOrMake.map(function (row) {
                      return [row[0]];
                    });
                    orderOrMakeCategory = [["Buy or Make"]];
                    _i2 = 1;
                  case 7:
                    if (!(_i2 < orderOrMakeOutput.length)) {
                      _context6.n = 9;
                      break;
                    }
                    workCenters = orderOrMakeOutput[_i2][0];
                    if (workCenters.includes("40FGAL3A") || workCenters.includes("40FGAL3B") || workCenters.includes("40FGAL3C") || workCenters.includes("40FGSI2A") || workCenters.includes("40AIFG2B")) {
                      orderOrMakeCategory.push(["Must Buy"]);
                    } else if (Number(requiredAmounts[_i2][0]) >= 300) {
                      orderOrMakeCategory.push(["Can Buy"]);
                    } else {
                      orderOrMakeCategory.push(["Can Make"]);
                    }
                    _context6.n = 8;
                    return context.sync();
                  case 8:
                    _i2++;
                    _context6.n = 7;
                    break;
                  case 9:
                    orderingWorksheet.getRange("F1:F".concat(orderOrMakeCategory.length)).values = orderOrMakeCategory;

                    // Table Formatting
                    orderingWorksheet.getRange("A:G").format.autofitColumns();
                    orderingWorksheet.getRange("A:H").format.horizontalAlignment = "Center";
                    orderingWorksheet.getRange("A:H").format.verticalAlignment = "Center";
                    orderingWorksheet.getRange("D:D").numberFormat = [['General']];
                    orderingWorksheet.getRange("A:A").format.columnWidth = 150;
                    orderingWorksheet.getRange("B:B").format.columnWidth = 90;
                    orderingWorksheet.getRange("C:C").format.columnWidth = 120;
                    orderingWorksheet.getRange("D:D").format.columnWidth = 90;
                    orderingWorksheet.getRange("E:E").format.columnWidth = 130;
                    orderingWorksheet.getRange("F:F").format.columnWidth = 100;
                    orderingWorksheet.getRange("B:B").format.columnWidth = 130;
                    orderingWorksheet.getUsedRange().format.rowHeight = 30;
                    orderingWorksheet.freezePanes.freezeRows(1);
                    orderingWorksheet.getRange("E1:E1").format.fill.color = "#BE5014";
                    orderingWorksheet.getRange("E1:E1").format.font.color = "yellow";

                    //All border lines
                    usedRange = orderingWorksheet.getUsedRange();
                    borders = usedRange.format.borders;
                    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach(function (edge) {
                      borders.getItem(edge).style = "Continuous";
                      borders.getItem(edge).weight = "Thin";
                      borders.getItem(edge).color = "#000000";
                    });
                    //Bold Outline Lines
                    lastRow = demandOutput.length;
                    highlight = orderingWorksheet.getRange("E1:E".concat(lastRow)).format.borders;
                    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(function (side) {
                      highlight.getItem(side).style = "Continuous";
                      highlight.getItem(side).weight = "Thick";
                      highlight.getItem(side).color = "#BE5014";
                    });
                    _context6.n = 10;
                    return context.sync();
                  case 10:
                    return _context6.a(2);
                }
              }, _callee6);
            }));
            return function (_x6) {
              return _ref3.apply(this, arguments);
            };
          }());
        case 1:
          return _context7.a(2);
      }
    }, _callee7);
  }));
  return _importColumnData.apply(this, arguments);
}
function importOtherColumnData(_x2) {
  return _importOtherColumnData.apply(this, arguments);
}
function _importOtherColumnData() {
  _importOtherColumnData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee9(event) {
    return _regenerator().w(function (_context9) {
      while (1) switch (_context9.n) {
        case 0:
          _context9.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref4 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee8(context) {
              var inventoryReportWorksheet, inventoryUsedRange, dynamicWorksheet, dynamicUsedRange, openPOsWorksheet, openPOsUsedRange, inventoryWorksheet, dynamicHeaders, dynItemCodeIdx, dynItemQtyIdx, dynItemCodeColumn, dynItemQtyColumn, openPOsHeaders, openPOItemCodeIdx, openPOItemQtyIdx, openPOItemCodeColumn, openPOItemQtyColumn, inventoryHeaders, invItemCodeIdx, invItemQtyIdx, invLocationIdx, invRepItemCodeColumn, invRepItemQtyColumn, invRepLocationColumn, inventoryICR, inventoryQR, inventoryLOC, openPOs, openPOsICR, openPOsQR, invFilterICR, invFilterQR, initialEntry, result, _iterator7, _step7, _code7, demandQty, caseNumbers, requiredAmounts, mebArray, efwArray, i, code, found, j, invCode, location, qty, isMeb, isEFW, mebSumMap, mebAmounts, efwSumMap, efwAmounts, total, openPOsMap, onOrder, _iterator8, _step8, _code8, onOrderQty, startArray, releaseArray, _i3, itemCode, start, dateOnly, releaseDate, adjustedRelease, qtyNeeded, filteredData, _i4, caseNumbersFiltered, demandFiltered, mebQtyFiltered, efwQtyFiltered, totalQtyFiltered, onOrderFiltered, startDateFiltered, releaseDateFiltered, qtyNeededFiltered, notesFiltered, usedRange, borders, lastRow, highlight;
              return _regenerator().w(function (_context8) {
                while (1) switch (_context8.n) {
                  case 0:
                    inventoryReportWorksheet = context.workbook.worksheets.getItem("Inventory");
                    inventoryUsedRange = inventoryReportWorksheet.getUsedRange().load("values");
                    dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
                    dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
                    openPOsWorksheet = context.workbook.worksheets.getItem("Open PO's");
                    openPOsUsedRange = openPOsWorksheet.getUsedRange().load("values");
                    inventoryWorksheet = context.workbook.worksheets.getItem("Inventory At");
                    _context8.n = 1;
                    return context.sync();
                  case 1:
                    //Dynamic fluid Placement
                    dynamicHeaders = dynamicUsedRange.values[0];
                    dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
                    dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
                    dynItemCodeColumn = "".concat(colIdxToLetter(dynItemCodeIdx), ":").concat(colIdxToLetter(dynItemCodeIdx));
                    dynItemQtyColumn = "".concat(colIdxToLetter(dynItemQtyIdx), ":").concat(colIdxToLetter(dynItemQtyIdx)); //Open PO's fluid Placement
                    openPOsHeaders = openPOsUsedRange.values[0];
                    openPOItemCodeIdx = openPOsHeaders.indexOf("Item Code");
                    openPOItemQtyIdx = openPOsHeaders.indexOf("Outstanding Qty");
                    openPOItemCodeColumn = "".concat(colIdxToLetter(openPOItemCodeIdx), ":").concat(colIdxToLetter(openPOItemCodeIdx));
                    openPOItemQtyColumn = "".concat(colIdxToLetter(openPOItemQtyIdx), ":").concat(colIdxToLetter(openPOItemQtyIdx)); //Inventory Report Fluid Placement
                    inventoryHeaders = inventoryUsedRange.values[0];
                    invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
                    invItemQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
                    invLocationIdx = inventoryHeaders.indexOf("Location");
                    invRepItemCodeColumn = "".concat(colIdxToLetter(invItemCodeIdx), ":").concat(colIdxToLetter(invItemCodeIdx));
                    invRepItemQtyColumn = "".concat(colIdxToLetter(invItemQtyIdx), ":").concat(colIdxToLetter(invItemQtyIdx));
                    invRepLocationColumn = "".concat(colIdxToLetter(invLocationIdx), ":").concat(colIdxToLetter(invLocationIdx));
                    _context8.n = 2;
                    return context.sync();
                  case 2:
                    //Quanity and Item Code from Dynamic, Inventory Report, and Open PO's sheets
                    inventoryICR = inventoryReportWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values");
                    inventoryQR = inventoryReportWorksheet.getRange(invRepItemQtyColumn).getUsedRange().load("values");
                    inventoryLOC = inventoryReportWorksheet.getRange(invRepLocationColumn).getUsedRange().load("values");
                    openPOs = context.workbook.worksheets.getItem("Open PO's");
                    openPOsICR = openPOs.getRange(openPOItemCodeColumn).getUsedRange().load("values");
                    openPOsQR = openPOs.getRange(openPOItemQtyColumn).getUsedRange().load("values");
                    _context8.n = 3;
                    return context.sync();
                  case 3:
                    //Date Filtering
                    invFilterICR = invFilter.map(function (item) {
                      return [item.itemCode];
                    });
                    invFilterQR = invFilter.map(function (item) {
                      return [item.qty];
                    }); //Sum Map Building
                    initialEntry = buildSumMap(invFilterICR, invFilterQR);
                    result = [["Case #", "Demand"]];
                    _iterator7 = _createForOfIteratorHelper(initialEntry.keys());
                    try {
                      for (_iterator7.s(); !(_step7 = _iterator7.n()).done;) {
                        _code7 = _step7.value;
                        demandQty = initialEntry.get(_code7) || 0;
                        if (demandQty > 0) {
                          result.push([_code7, demandQty]);
                        }
                      }
                    } catch (err) {
                      _iterator7.e(err);
                    } finally {
                      _iterator7.f();
                    }
                    _context8.n = 4;
                    return context.sync();
                  case 4:
                    caseNumbers = result.map(function (row) {
                      return row[0];
                    });
                    requiredAmounts = result.map(function (row) {
                      return [row[1]];
                    });
                    console.log("Required Amounts", requiredAmounts);

                    //Mebane-EFW Inventory Map
                    mebArray = [];
                    efwArray = [];
                    for (i = 0; i < caseNumbers.length; i++) {
                      code = String(caseNumbers[i]).trim();
                      found = false;
                      for (j = 0; j < inventoryICR.values.length; j++) {
                        invCode = String(inventoryICR.values[j][0]).trim();
                        location = String(inventoryLOC.values[j][0]).trim();
                        qty = Number(inventoryQR.values[j][0]);
                        if (code === invCode) {
                          isMeb = location.includes("MEB");
                          isEFW = location.includes("EFW");
                          mebArray.push([code, isMeb ? qty : 0]);
                          efwArray.push([code, isEFW ? qty : 0]);
                          found = true;
                        }
                      }
                      if (!found) {
                        mebArray.push([code, 0]);
                        efwArray.push([code, 0]);
                      }
                    }
                    _context8.n = 5;
                    return context.sync();
                  case 5:
                    mebSumMap = buildSumMap(mebArray.map(function (item) {
                      return [item[0]];
                    }), mebArray.map(function (item) {
                      return [item[1]];
                    }));
                    mebAmounts = Array.from(mebSumMap.entries()).map(function (row) {
                      return [row[1]];
                    });
                    efwSumMap = buildSumMap(efwArray.map(function (item) {
                      return [item[0]];
                    }), efwArray.map(function (item) {
                      return [item[1]];
                    }));
                    efwAmounts = Array.from(efwSumMap.entries()).map(function (row) {
                      return [row[1]];
                    });
                    total = mebAmounts.map(function (value, index) {
                      return Number(value) + efwAmounts[index][0];
                    }); //Inventory On Order 
                    openPOsMap = buildSumMap(openPOsICR.values, openPOsQR.values);
                    onOrder = [["On Order"]];
                    _iterator8 = _createForOfIteratorHelper(caseNumbers.slice(1));
                    try {
                      for (_iterator8.s(); !(_step8 = _iterator8.n()).done;) {
                        _code8 = _step8.value;
                        onOrderQty = openPOsMap.get(_code8) || 0;
                        onOrder.push([onOrderQty]);
                      }

                      // Importing the Start and Release Date
                    } catch (err) {
                      _iterator8.e(err);
                    } finally {
                      _iterator8.f();
                    }
                    startArray = [["Earliest Start Date"]];
                    releaseArray = [["Release Date"]];
                    for (_i3 = 1; _i3 < caseNumbers.length; _i3++) {
                      itemCode = String(caseNumbers[_i3]).trim();
                      start = String(startDateMap.get(itemCode)) || "No Start Date Established";
                      dateOnly = start.split(' ').slice(0, 4).join(' ');
                      startArray.push([dateOnly]);
                      releaseDate = new Date(start);
                      if (!isNaN(releaseDate)) {
                        releaseDate.setDate(releaseDate.getDate() - 10);
                        adjustedRelease = releaseDate.toDateString();
                        releaseArray.push([adjustedRelease]);
                      } else {
                        releaseArray.push(["Invalid Release Date"]);
                      }
                    }
                    qtyNeeded = requiredAmounts.slice(1).map(function (value, index) {
                      return Number(value) - mebAmounts[index];
                    });
                    filteredData = [];
                    filteredData.push({
                      caseNumber: "Case #",
                      demand: "Demand",
                      mebQty: "Qty MEB",
                      efwQty: "Qty EFW",
                      totalQty: "Total MEB + EFW",
                      onOrder: "On Order",
                      startDate: "Start Date",
                      releaseDate: "Release Date",
                      qtyNeeded: "Qty Needed (MEB)",
                      notes: "Notes"
                    });
                    for (_i4 = 0; _i4 < caseNumbers.slice(1).length; _i4++) {
                      if (qtyNeeded[_i4] > 0) {
                        filteredData.push({
                          caseNumber: caseNumbers[_i4 + 1],
                          demand: requiredAmounts[_i4 + 1][0],
                          mebQty: mebAmounts[_i4][0],
                          efwQty: efwAmounts[_i4][0],
                          totalQty: total[_i4],
                          onOrder: onOrder[_i4 + 1][0],
                          startDate: startArray[_i4 + 1][0],
                          releaseDate: releaseArray[_i4 + 1][0],
                          qtyNeeded: qtyNeeded[_i4],
                          notes: ""
                        });
                      }
                    }
                    if (filteredData.length > 1) {
                      caseNumbersFiltered = filteredData.map(function (row) {
                        return [row.caseNumber];
                      });
                      demandFiltered = filteredData.map(function (row) {
                        return [row.demand];
                      });
                      mebQtyFiltered = filteredData.map(function (row) {
                        return [row.mebQty];
                      });
                      efwQtyFiltered = filteredData.map(function (row) {
                        return [row.efwQty];
                      });
                      totalQtyFiltered = filteredData.map(function (row) {
                        return [row.totalQty];
                      });
                      onOrderFiltered = filteredData.map(function (row) {
                        return [row.onOrder];
                      });
                      startDateFiltered = filteredData.map(function (row) {
                        return [row.startDate];
                      });
                      releaseDateFiltered = filteredData.map(function (row) {
                        return [row.releaseDate];
                      });
                      qtyNeededFiltered = filteredData.map(function (row) {
                        return [row.qtyNeeded];
                      });
                      notesFiltered = filteredData.map(function (row) {
                        return [row.notes];
                      });
                      inventoryWorksheet.getRange("A1:A".concat(caseNumbersFiltered.length)).values = caseNumbersFiltered;
                      inventoryWorksheet.getRange("B1:B".concat(demandFiltered.length)).values = demandFiltered;
                      inventoryWorksheet.getRange("C1:C".concat(mebQtyFiltered.length)).values = mebQtyFiltered;
                      inventoryWorksheet.getRange("D1:D".concat(efwQtyFiltered.length)).values = efwQtyFiltered;
                      inventoryWorksheet.getRange("E1:E".concat(totalQtyFiltered.length)).values = totalQtyFiltered;
                      inventoryWorksheet.getRange("F1:F".concat(onOrderFiltered.length)).values = onOrderFiltered;
                      inventoryWorksheet.getRange("G1:G".concat(startDateFiltered.length)).values = startDateFiltered;
                      inventoryWorksheet.getRange("H1:H".concat(releaseDateFiltered.length)).values = releaseDateFiltered;
                      inventoryWorksheet.getRange("I1:I".concat(qtyNeededFiltered.length)).values = qtyNeededFiltered;
                      inventoryWorksheet.getRange("J1:J".concat(notesFiltered.length)).values = notesFiltered;
                    }
                    _context8.n = 6;
                    return context.sync();
                  case 6:
                    //Inventory formatting
                    inventoryWorksheet.getRange("A:J").format.autofitColumns();
                    inventoryWorksheet.getRange("A:J").format.horizontalAlignment = "Center";
                    inventoryWorksheet.getRange("A:J").format.verticalAlignment = "Center";
                    inventoryWorksheet.getRange("C:C").numberFormat = [['General']];
                    inventoryWorksheet.getRange("A:A").format.columnWidth = 120;
                    inventoryWorksheet.getRange("B:B").format.columnWidth = 70;
                    inventoryWorksheet.getRange("C:C").format.columnWidth = 70;
                    inventoryWorksheet.getRange("D:D").format.columnWidth = 70;
                    inventoryWorksheet.getRange("E:E").format.columnWidth = 100;
                    inventoryWorksheet.getRange("F:F").format.columnWidth = 75;
                    inventoryWorksheet.getRange("G:G").format.columnWidth = 90;
                    inventoryWorksheet.getRange("H:H").format.columnWidth = 90;
                    inventoryWorksheet.getRange("I:I").format.columnWidth = 130;
                    inventoryWorksheet.getRange("J:J").format.columnWidth = 150;
                    inventoryWorksheet.getUsedRange().format.rowHeight = 20;
                    inventoryWorksheet.freezePanes.freezeRows(1);
                    inventoryWorksheet.getRange("I1:I1").format.fill.color = "#00008B";
                    inventoryWorksheet.getRange("I1:I1").format.font.color = "yellow";

                    //All border lines
                    usedRange = inventoryWorksheet.getUsedRange();
                    borders = usedRange.format.borders;
                    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach(function (edge) {
                      borders.getItem(edge).style = "Continuous";
                      borders.getItem(edge).weight = "Thin";
                      borders.getItem(edge).color = "#000000";
                    });

                    //Bold Outline Lines
                    lastRow = filteredData.length;
                    highlight = inventoryWorksheet.getRange("I1:I".concat(lastRow)).format.borders;
                    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(function (side) {
                      highlight.getItem(side).style = "Continuous";
                      highlight.getItem(side).weight = "Thick";
                      highlight.getItem(side).color = "#00008B";
                    });
                    _context8.n = 7;
                    return context.sync();
                  case 7:
                    return _context8.a(2);
                }
              }, _callee8);
            }));
            return function (_x7) {
              return _ref4.apply(this, arguments);
            };
          }());
        case 1:
          return _context9.a(2);
      }
    }, _callee9);
  }));
  return _importOtherColumnData.apply(this, arguments);
}
function readoutData() {
  return _readoutData.apply(this, arguments);
}
function _readoutData() {
  _readoutData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee1() {
    return _regenerator().w(function (_context1) {
      while (1) switch (_context1.n) {
        case 0:
          _context1.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref5 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee0(context) {
              var inventoryAtWorksheet, inventoryAtUsedRange, inventoryWorksheet, inventoryUsedRange, inventoryRequestWorksheet, inventoryAtHeaders, invAtItemCodeIdx, invAtQtyNeededIdx, invAtQtyEFWIdx, invAtItemCodeColumn, invAtQtyNeededColumn, invAtQtyEFWColumn, inventoryHeaders, invItemCodeIdx, invDateIdx, invRefIdx, invQtyIdx, invLocIdx, invRepItemCodeColumn, invRepDateColumn, invRepRefColumn, invRepQtyColumn, invLocColumn, invAtItemCodes, invAtQtyNeeded, invAtQtyEFW, invItemCodes, invDates, invRefs, invQtys, invLoc, filteredData, i, qtyNeeded, qtyEFW, inventoryDataMap, _i5, itemCode, dateStr, start, date, loc, ref, qty, _iterator9, _step9, _step9$value, _itemCode2, data, readoutResult, _i6, _filteredData, filteredItem, _itemCode, _qtyNeeded, inventoryItems, totalPulled, palletsPulled, _iterator0, _step0, invItem, itemCodes, dates, refs, qtys, usedRange, borders, _t2;
              return _regenerator().w(function (_context0) {
                while (1) switch (_context0.n) {
                  case 0:
                    inventoryAtWorksheet = context.workbook.worksheets.getItem("Inventory At");
                    inventoryAtUsedRange = inventoryAtWorksheet.getUsedRange().load("values");
                    inventoryWorksheet = context.workbook.worksheets.getItem("Inventory");
                    inventoryUsedRange = inventoryWorksheet.getUsedRange().load("values");
                    inventoryRequestWorksheet = context.workbook.worksheets.getItem("Inventory Request");
                    _context0.n = 1;
                    return context.sync();
                  case 1:
                    //Inventory At fluid Placement
                    inventoryAtHeaders = inventoryAtUsedRange.values[0];
                    invAtItemCodeIdx = inventoryAtHeaders.indexOf("Case #");
                    invAtQtyNeededIdx = inventoryAtHeaders.indexOf("Qty Needed (MEB)");
                    invAtQtyEFWIdx = inventoryAtHeaders.indexOf("Qty EFW");
                    invAtItemCodeColumn = "".concat(colIdxToLetter(invAtItemCodeIdx), ":").concat(colIdxToLetter(invAtItemCodeIdx));
                    invAtQtyNeededColumn = "".concat(colIdxToLetter(invAtQtyNeededIdx), ":").concat(colIdxToLetter(invAtQtyNeededIdx));
                    invAtQtyEFWColumn = "".concat(colIdxToLetter(invAtQtyEFWIdx), ":").concat(colIdxToLetter(invAtQtyEFWIdx)); //Inventory Report fluid Placement
                    inventoryHeaders = inventoryUsedRange.values[0];
                    invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
                    invDateIdx = inventoryHeaders.indexOf("Inventory Date");
                    invRefIdx = inventoryHeaders.indexOf("Inventory Ref");
                    invQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
                    invLocIdx = inventoryHeaders.indexOf("Location");
                    invRepItemCodeColumn = "".concat(colIdxToLetter(invItemCodeIdx), ":").concat(colIdxToLetter(invItemCodeIdx));
                    invRepDateColumn = "".concat(colIdxToLetter(invDateIdx), ":").concat(colIdxToLetter(invDateIdx));
                    invRepRefColumn = "".concat(colIdxToLetter(invRefIdx), ":").concat(colIdxToLetter(invRefIdx));
                    invRepQtyColumn = "".concat(colIdxToLetter(invQtyIdx), ":").concat(colIdxToLetter(invQtyIdx));
                    invLocColumn = "".concat(colIdxToLetter(invLocIdx), ":").concat(colIdxToLetter(invLocIdx));
                    _context0.n = 2;
                    return context.sync();
                  case 2:
                    //Get data from Inventory At sheet
                    invAtItemCodes = inventoryAtWorksheet.getRange(invAtItemCodeColumn).getUsedRange().load("values");
                    invAtQtyNeeded = inventoryAtWorksheet.getRange(invAtQtyNeededColumn).getUsedRange().load("values");
                    invAtQtyEFW = inventoryAtWorksheet.getRange(invAtQtyEFWColumn).getUsedRange().load("values"); //Get data from Inventory sheet
                    invItemCodes = inventoryWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values");
                    invDates = inventoryWorksheet.getRange(invRepDateColumn).getUsedRange().load("values");
                    invRefs = inventoryWorksheet.getRange(invRepRefColumn).getUsedRange().load("values");
                    invQtys = inventoryWorksheet.getRange(invRepQtyColumn).getUsedRange().load("values");
                    invLoc = inventoryWorksheet.getRange(invLocColumn).getUsedRange().load("values");
                    _context0.n = 3;
                    return context.sync();
                  case 3:
                    filteredData = [];
                    for (i = 1; i < invAtItemCodes.values.length; i++) {
                      qtyNeeded = Number(invAtQtyNeeded.values[i][0]);
                      qtyEFW = Number(invAtQtyEFW.values[i][0]);
                      if (qtyNeeded > 300 && qtyEFW > 0) {
                        filteredData.push({
                          itemCode: String(invAtItemCodes.values[i][0]).trim(),
                          qtyNeeded: qtyNeeded
                        });
                      }
                    }

                    //Build inventory data map for each item code
                    inventoryDataMap = new Map();
                    for (_i5 = 1; _i5 < invItemCodes.values.length; _i5++) {
                      itemCode = String(invItemCodes.values[_i5][0]).trim();
                      dateStr = invDates.values[_i5] ? String(invDates.values[_i5][0]).trim() : "";
                      start = ExcelDateToJSDate(dateStr);
                      start.setHours(0, 0, 0, 0);
                      date = formatDate(start);
                      loc = invLoc.values[_i5] ? String(invLoc.values[_i5][0]).trim() : "";
                      ref = invRefs.values[_i5] ? String(invRefs.values[_i5][0]).trim() : "";
                      qty = Number(invQtys.values[_i5][0]);
                      if (loc.includes("EFW")) {
                        if (itemCode && !isNaN(qty) && qty > 0) {
                          if (!inventoryDataMap.has(itemCode)) {
                            inventoryDataMap.set(itemCode, []);
                          }
                          inventoryDataMap.get(itemCode).push({
                            date: date,
                            ref: ref,
                            qty: qty
                          });
                        }
                      }
                    }

                    //Sort inventory data by date (oldest first) for each item code
                    _iterator9 = _createForOfIteratorHelper(inventoryDataMap);
                    try {
                      for (_iterator9.s(); !(_step9 = _iterator9.n()).done;) {
                        _step9$value = _slicedToArray(_step9.value, 2), _itemCode2 = _step9$value[0], data = _step9$value[1];
                        data.sort(function (a, b) {
                          var dateA = new Date(a.date);
                          var dateB = new Date(b.date);
                          return dateA - dateB;
                        });
                      }

                      //Generate readout data
                    } catch (err) {
                      _iterator9.e(err);
                    } finally {
                      _iterator9.f();
                    }
                    readoutResult = [["Item Code", "Inventory Date", "Inventory Ref", "Inventory Qty"]];
                    _i6 = 0, _filteredData = filteredData;
                  case 4:
                    if (!(_i6 < _filteredData.length)) {
                      _context0.n = 13;
                      break;
                    }
                    filteredItem = _filteredData[_i6];
                    _itemCode = filteredItem.itemCode;
                    _qtyNeeded = filteredItem.qtyNeeded;
                    if (!inventoryDataMap.has(_itemCode)) {
                      _context0.n = 12;
                      break;
                    }
                    inventoryItems = inventoryDataMap.get(_itemCode);
                    totalPulled = 0;
                    palletsPulled = 0;
                    _iterator0 = _createForOfIteratorHelper(inventoryItems);
                    _context0.p = 5;
                    _iterator0.s();
                  case 6:
                    if ((_step0 = _iterator0.n()).done) {
                      _context0.n = 9;
                      break;
                    }
                    invItem = _step0.value;
                    if (!(totalPulled >= _qtyNeeded)) {
                      _context0.n = 7;
                      break;
                    }
                    return _context0.a(3, 9);
                  case 7:
                    readoutResult.push([_itemCode, invItem.date, invItem.ref, invItem.qty]);
                    totalPulled += invItem.qty;
                    palletsPulled++;
                  case 8:
                    _context0.n = 6;
                    break;
                  case 9:
                    _context0.n = 11;
                    break;
                  case 10:
                    _context0.p = 10;
                    _t2 = _context0.v;
                    _iterator0.e(_t2);
                  case 11:
                    _context0.p = 11;
                    _iterator0.f();
                    return _context0.f(11);
                  case 12:
                    _i6++;
                    _context0.n = 4;
                    break;
                  case 13:
                    //Write data to Inventory Request table
                    if (readoutResult.length > 1) {
                      itemCodes = readoutResult.map(function (row) {
                        return [row[0]];
                      });
                      dates = readoutResult.map(function (row) {
                        return [row[1]];
                      });
                      refs = readoutResult.map(function (row) {
                        return [row[2]];
                      });
                      qtys = readoutResult.map(function (row) {
                        return [row[3]];
                      });
                      inventoryRequestWorksheet.getRange("A1:A".concat(itemCodes.length)).values = itemCodes;
                      inventoryRequestWorksheet.getRange("B1:B".concat(dates.length)).values = dates;
                      inventoryRequestWorksheet.getRange("C1:C".concat(refs.length)).values = refs;
                      inventoryRequestWorksheet.getRange("D1:D".concat(qtys.length)).values = qtys;
                    }
                    _context0.n = 14;
                    return context.sync();
                  case 14:
                    //Formatting
                    inventoryRequestWorksheet.getRange("A:D").format.autofitColumns();
                    inventoryRequestWorksheet.getRange("A:D").format.horizontalAlignment = "Center";
                    inventoryRequestWorksheet.getRange("A:D").format.verticalAlignment = "Center";
                    inventoryRequestWorksheet.getRange("A:A").format.columnWidth = 120;
                    inventoryRequestWorksheet.getRange("B:B").format.columnWidth = 100;
                    inventoryRequestWorksheet.getRange("C:C").format.columnWidth = 100;
                    inventoryRequestWorksheet.getRange("D:D").format.columnWidth = 100;
                    inventoryRequestWorksheet.getUsedRange().format.rowHeight = 20;
                    inventoryRequestWorksheet.freezePanes.freezeRows(1);

                    //All border lines
                    usedRange = inventoryRequestWorksheet.getUsedRange();
                    borders = usedRange.format.borders;
                    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach(function (edge) {
                      borders.getItem(edge).style = "Continuous";
                      borders.getItem(edge).weight = "Thin";
                      borders.getItem(edge).color = "#000000";
                    });
                    _context0.n = 15;
                    return context.sync();
                  case 15:
                    return _context0.a(2);
                }
              }, _callee0, null, [[5, 10, 11, 12]]);
            }));
            return function (_x8) {
              return _ref5.apply(this, arguments);
            };
          }());
        case 1:
          return _context1.a(2);
      }
    }, _callee1);
  }));
  return _readoutData.apply(this, arguments);
}
function resetGenerateOrdering() {
  return _resetGenerateOrdering.apply(this, arguments);
}
function _resetGenerateOrdering() {
  _resetGenerateOrdering = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee11() {
    return _regenerator().w(function (_context11) {
      while (1) switch (_context11.n) {
        case 0:
          _context11.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref6 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee10(context) {
              var sheets;
              return _regenerator().w(function (_context10) {
                while (1) switch (_context10.n) {
                  case 0:
                    sheets = context.workbook.worksheets;
                    sheets.getItemOrNullObject("Ordering").delete();
                    filter = [];
                    earlyDateMap.clear();
                    _context10.n = 1;
                    return context.sync();
                  case 1:
                    return _context10.a(2);
                }
              }, _callee10);
            }));
            return function (_x9) {
              return _ref6.apply(this, arguments);
            };
          }());
        case 1:
          return _context11.a(2);
      }
    }, _callee11);
  }));
  return _resetGenerateOrdering.apply(this, arguments);
}
function resetGenerateInventory() {
  return _resetGenerateInventory.apply(this, arguments);
}
function _resetGenerateInventory() {
  _resetGenerateInventory = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee13() {
    return _regenerator().w(function (_context13) {
      while (1) switch (_context13.n) {
        case 0:
          _context13.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref7 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee12(context) {
              var sheets;
              return _regenerator().w(function (_context12) {
                while (1) switch (_context12.n) {
                  case 0:
                    sheets = context.workbook.worksheets;
                    sheets.getItemOrNullObject("Inventory At").delete();
                    sheets.getItemOrNullObject("Inventory Request").delete();
                    _context12.n = 1;
                    return context.sync();
                  case 1:
                    return _context12.a(2);
                }
              }, _callee12);
            }));
            return function (_x0) {
              return _ref7.apply(this, arguments);
            };
          }());
        case 1:
          return _context13.a(2);
      }
    }, _callee13);
  }));
  return _resetGenerateInventory.apply(this, arguments);
}
function colIdxToLetter(idx) {
  var letter = "";
  while (idx >= 0) {
    letter = String.fromCharCode(idx % 26 + 65) + letter;
    idx = Math.floor(idx / 26) - 1;
  }
  return letter;
}
function buildSumMap(itemCodes, qtys) {
  var map = new Map();
  for (var i = 1; i < itemCodes.length; i++) {
    var code = itemCodes[i][0];
    var qty = Number(qtys[i][0]);
    if (code && !isNaN(qty)) {
      map.set(code, (map.get(code) || 0) + qty);
    }
  }
  return map;
}
function dateFilter() {
  return _dateFilter.apply(this, arguments);
}
function _dateFilter() {
  _dateFilter = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee15() {
    return _regenerator().w(function (_context15) {
      while (1) switch (_context15.n) {
        case 0:
          _context15.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref8 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee14(context) {
              var startDateInput, endDateInput, startDate, endDate, dynamicWorksheet, dynamicUsedRange, dynamicHeaders, dynItemCodeIdx, dynStartIdx, dynItemQtyIdx, dynJobIdx, dynStartColumn, dynItemQtyColumn, dynItemCodeColumn, dynJobColumn, dynamic, dynamicICR, dynamicQR, plannedStart, jobNumber, jobLatestMap, i, itemCode, dateStr, date, job, qty, _iterator1, _step1, entry, _itemCode4, _date2, _i7, _itemCode3, _dateStr, _date, _job, _qty;
              return _regenerator().w(function (_context14) {
                while (1) switch (_context14.n) {
                  case 0:
                    startDateInput = document.getElementById('start-date').value;
                    endDateInput = document.getElementById('end-date').value;
                    startDate = inputDateParse(startDateInput);
                    endDate = inputDateParse(endDateInput);
                    dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
                    dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
                    _context14.n = 1;
                    return context.sync();
                  case 1:
                    dynamicHeaders = dynamicUsedRange.values[0];
                    dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
                    dynStartIdx = dynamicHeaders.indexOf("Planned Start");
                    dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
                    dynJobIdx = dynamicHeaders.indexOf("Order Number");
                    dynStartColumn = "".concat(colIdxToLetter(dynStartIdx), ":").concat(colIdxToLetter(dynStartIdx));
                    dynItemQtyColumn = "".concat(colIdxToLetter(dynItemQtyIdx), ":").concat(colIdxToLetter(dynItemQtyIdx));
                    dynItemCodeColumn = "".concat(colIdxToLetter(dynItemCodeIdx), ":").concat(colIdxToLetter(dynItemCodeIdx));
                    dynJobColumn = "".concat(colIdxToLetter(dynJobIdx), ":").concat(colIdxToLetter(dynJobIdx));
                    dynamic = context.workbook.worksheets.getItem("Dynamic");
                    dynamicICR = dynamic.getRange(dynItemCodeColumn).getUsedRange().load("values");
                    dynamicQR = dynamic.getRange(dynItemQtyColumn).getUsedRange().load("values");
                    plannedStart = dynamicWorksheet.getRange(dynStartColumn).getUsedRange().load("values");
                    jobNumber = dynamicWorksheet.getRange(dynJobColumn).getUsedRange().load("values");
                    _context14.n = 2;
                    return context.sync();
                  case 2:
                    jobLatestMap = new Map();
                    for (i = 1; i < dynamicICR.values.length; i++) {
                      itemCode = String(dynamicICR.values[i][0]).trim();
                      dateStr = plannedStart.values[i] ? String(plannedStart.values[i][0]).trim() : "";
                      date = ExcelDateToJSDate(dateStr);
                      date.setHours(0, 0, 0, 0);
                      job = String(jobNumber.values[i][0]).trim();
                      qty = Number(dynamicQR.values[i][0]);
                      if (itemCode && date >= startDate && date <= endDate) {
                        if (!jobLatestMap.has(job) || date > jobLatestMap.get(job).date) {
                          jobLatestMap.set(job, {
                            itemCode: itemCode,
                            qty: qty,
                            date: date,
                            job: job
                          });
                        }
                      }
                    }
                    filter = Array.from(jobLatestMap.values());
                    earlyDateMap.clear();
                    _iterator1 = _createForOfIteratorHelper(filter);
                    try {
                      for (_iterator1.s(); !(_step1 = _iterator1.n()).done;) {
                        entry = _step1.value;
                        _itemCode4 = entry.itemCode, _date2 = entry.date;
                        if (!earlyDateMap.has(_itemCode4) || _date2 < earlyDateMap.get(_itemCode4)) {
                          earlyDateMap.set(_itemCode4, _date2);
                        }
                      }
                    } catch (err) {
                      _iterator1.e(err);
                    } finally {
                      _iterator1.f();
                    }
                    filter.sort(function (a, b) {
                      return a.date - b.date;
                    });
                    _context14.n = 3;
                    return context.sync();
                  case 3:
                    for (_i7 = 1; _i7 < dynamicICR.values.length; _i7++) {
                      _itemCode3 = String(dynamicICR.values[_i7][0]).trim();
                      _dateStr = plannedStart.values[_i7] ? String(plannedStart.values[_i7][0]).trim() : "";
                      _date = ExcelDateToJSDate(_dateStr);
                      _date.setHours(0, 0, 0, 0);
                      _job = String(jobNumber.values[_i7][0]).trim();
                      _qty = Number(dynamicQR.values[_i7][0]);
                      if (_itemCode3) {
                        allData.push({
                          itemCode: _itemCode3,
                          qty: _qty,
                          job: _job,
                          date: _date
                        });
                      }
                    }
                  case 4:
                    return _context14.a(2);
                }
              }, _callee14);
            }));
            return function (_x1) {
              return _ref8.apply(this, arguments);
            };
          }());
        case 1:
          return _context15.a(2);
      }
    }, _callee15);
  }));
  return _dateFilter.apply(this, arguments);
}
function otherDateFilter() {
  return _otherDateFilter.apply(this, arguments);
}
function _otherDateFilter() {
  _otherDateFilter = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee17() {
    return _regenerator().w(function (_context17) {
      while (1) switch (_context17.n) {
        case 0:
          _context17.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref9 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee16(context) {
              var startDateInput, endDateInput, startDate, endDate, dynamicWorksheet, dynamicUsedRange, dynamicHeaders, dynItemCodeIdx, dynStartIdx, dynItemQtyIdx, dynJobIdx, dynStartColumn, dynItemQtyColumn, dynItemCodeColumn, dynJobColumn, dynamic, dynamicICR, dynamicQR, plannedStart, jobNumber, jobLatestMap, i, itemCode, dateStr, date, job, qty, _iterator10, _step10, entry, _itemCode5, _date3;
              return _regenerator().w(function (_context16) {
                while (1) switch (_context16.n) {
                  case 0:
                    startDateInput = document.getElementById('start-date').value;
                    endDateInput = document.getElementById('end-date').value;
                    startDate = inputDateParse(startDateInput);
                    endDate = inputDateParse(endDateInput);
                    dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
                    dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
                    _context16.n = 1;
                    return context.sync();
                  case 1:
                    dynamicHeaders = dynamicUsedRange.values[0];
                    dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
                    dynStartIdx = dynamicHeaders.indexOf("Planned Start");
                    dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
                    dynJobIdx = dynamicHeaders.indexOf("Order Number");
                    dynStartColumn = "".concat(colIdxToLetter(dynStartIdx), ":").concat(colIdxToLetter(dynStartIdx));
                    dynItemQtyColumn = "".concat(colIdxToLetter(dynItemQtyIdx), ":").concat(colIdxToLetter(dynItemQtyIdx));
                    dynItemCodeColumn = "".concat(colIdxToLetter(dynItemCodeIdx), ":").concat(colIdxToLetter(dynItemCodeIdx));
                    dynJobColumn = "".concat(colIdxToLetter(dynJobIdx), ":").concat(colIdxToLetter(dynJobIdx));
                    dynamic = context.workbook.worksheets.getItem("Dynamic");
                    dynamicICR = dynamic.getRange(dynItemCodeColumn).getUsedRange().load("values");
                    dynamicQR = dynamic.getRange(dynItemQtyColumn).getUsedRange().load("values");
                    plannedStart = dynamicWorksheet.getRange(dynStartColumn).getUsedRange().load("values");
                    jobNumber = dynamicWorksheet.getRange(dynJobColumn).getUsedRange().load("values");
                    _context16.n = 2;
                    return context.sync();
                  case 2:
                    jobLatestMap = new Map();
                    for (i = 1; i < dynamicICR.values.length; i++) {
                      itemCode = String(dynamicICR.values[i][0]).trim();
                      dateStr = plannedStart.values[i] ? String(plannedStart.values[i][0]).trim() : "";
                      date = ExcelDateToJSDate(dateStr);
                      date.setHours(0, 0, 0, 0);
                      job = String(jobNumber.values[i][0]).trim();
                      qty = Number(dynamicQR.values[i][0]);
                      if (itemCode && date >= startDate && date <= endDate) {
                        if (!jobLatestMap.has(job) || date > jobLatestMap.get(job).date) {
                          jobLatestMap.set(job, {
                            itemCode: itemCode,
                            qty: qty,
                            date: date,
                            job: job
                          });
                        }
                      }
                    }
                    invFilter = Array.from(jobLatestMap.values());
                    startDateMap.clear();
                    _iterator10 = _createForOfIteratorHelper(invFilter);
                    try {
                      for (_iterator10.s(); !(_step10 = _iterator10.n()).done;) {
                        entry = _step10.value;
                        _itemCode5 = entry.itemCode, _date3 = entry.date;
                        if (!startDateMap.has(_itemCode5) || _date3 < startDateMap.get(_itemCode5)) {
                          startDateMap.set(_itemCode5, _date3);
                        }
                      }
                    } catch (err) {
                      _iterator10.e(err);
                    } finally {
                      _iterator10.f();
                    }
                    invFilter.sort(function (a, b) {
                      return a.date - b.date;
                    });
                    _context16.n = 3;
                    return context.sync();
                  case 3:
                    return _context16.a(2);
                }
              }, _callee16);
            }));
            return function (_x10) {
              return _ref9.apply(this, arguments);
            };
          }());
        case 1:
          return _context17.a(2);
      }
    }, _callee17);
  }));
  return _otherDateFilter.apply(this, arguments);
}
function inputDateParse(str) {
  var _str$split$map = str.split('-').map(Number),
    _str$split$map2 = _slicedToArray(_str$split$map, 3),
    year = _str$split$map2[0],
    month = _str$split$map2[1],
    day = _str$split$map2[2];
  return new Date(year, month - 1, day);
}
function ExcelDateToJSDate(excelDate) {
  var utcDate = new Date((excelDate - 25569) * 86400 * 1000);
  return new Date(utcDate.getTime() + utcDate.getTimezoneOffset() * 60000);
}
function checkDatesAndClearMessage() {
  var startDateValue = document.getElementById('start-date').value;
  var endDateValue = document.getElementById('end-date').value;
  if (startDateValue && endDateValue) {
    document.getElementById('message-area').textContent = "";
  }
}
function displayData(_x3) {
  return _displayData.apply(this, arguments);
}
function _displayData() {
  _displayData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee19(event) {
    return _regenerator().w(function (_context20) {
      while (1) switch (_context20.n) {
        case 0:
          _context20.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref0 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee18(context) {
              var sheet, range, match, allDataICR, allDatajob, allDataQR, allDatadate, _loop, i;
              return _regenerator().w(function (_context19) {
                while (1) switch (_context19.n) {
                  case 0:
                    sheet = context.workbook.worksheets.getItem("Ordering");
                    range = sheet.getRange(event.address);
                    range.load(["columnIndex", "values", "address"]);
                    _context19.n = 1;
                    return context.sync();
                  case 1:
                    console.log("Index Number", range.columnIndex);
                    outputJobs.clear();
                    if (!(range.columnIndex == 0)) {
                      _context19.n = 5;
                      break;
                    }
                    matchingData.length = 0;
                    match = range.values[0];
                    allDataICR = allData.map(function (item) {
                      return [item.itemCode];
                    });
                    allDatajob = allData.map(function (item) {
                      return [item.job];
                    });
                    allDataQR = allData.map(function (item) {
                      return [item.qty];
                    });
                    allDatadate = allData.map(function (item) {
                      return [item.date];
                    });
                    _loop = /*#__PURE__*/_regenerator().m(function _loop() {
                      var code, job, qty, date, fDate, duplicateDate, idx;
                      return _regenerator().w(function (_context18) {
                        while (1) switch (_context18.n) {
                          case 0:
                            code = allDataICR[i][0];
                            job = allDatajob[i][0];
                            qty = allDataQR[i][0];
                            date = allDatadate[i][0];
                            fDate = formatDate(date);
                            if (match == code) {
                              if (!outputJobs.has(job)) {
                                matchingData.push({
                                  code: code,
                                  job: job,
                                  qty: qty,
                                  fDate: fDate
                                });
                                outputJobs.add(job);
                              } else {
                                duplicateDate = earlyDateMap.get(code);
                                idx = matchingData.findIndex(function (entry) {
                                  return entry.job === job && entry.code === code;
                                });
                                if (idx !== -1) {
                                  matchingData[idx].date = duplicateDate;
                                }
                              }
                            }
                          case 1:
                            return _context18.a(2);
                        }
                      }, _loop);
                    });
                    i = 0;
                  case 2:
                    if (!(i < allDataICR.length)) {
                      _context19.n = 4;
                      break;
                    }
                    return _context19.d(_regeneratorValues(_loop()), 3);
                  case 3:
                    i++;
                    _context19.n = 2;
                    break;
                  case 4:
                    console.log("intial finding of Matching Data", matchingData);
                    handleCellChange([].concat(matchingData));
                    matchingData.sort(function (a, b) {
                      return a.date - b.date;
                    });
                    localStorage.setItem("matchingData", JSON.stringify(matchingData));
                    _context19.n = 6;
                    break;
                  case 5:
                    console.log("Not in range");
                  case 6:
                    return _context19.a(2);
                }
              }, _callee18);
            }));
            return function (_x11) {
              return _ref0.apply(this, arguments);
            };
          }());
        case 1:
          return _context20.a(2);
      }
    }, _callee19);
  }));
  return _displayData.apply(this, arguments);
}
function filteringDropdown() {
  return _filteringDropdown.apply(this, arguments);
}
function _filteringDropdown() {
  _filteringDropdown = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee21() {
    return _regenerator().w(function (_context22) {
      while (1) switch (_context22.n) {
        case 0:
          _context22.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref1 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee20(context) {
              var orderingWorksheet, orderingTable, amountFilter, buyOrMakeFilter, _t3;
              return _regenerator().w(function (_context21) {
                while (1) switch (_context21.n) {
                  case 0:
                    orderingWorksheet = context.workbook.worksheets.getItem("Ordering");
                    orderingTable = orderingWorksheet.tables.getItem("OrderingTable");
                    amountFilter = orderingTable.columns.getItem("Required Amount").filter;
                    buyOrMakeFilter = orderingTable.columns.getItem("Buy or Make").filter;
                    _t3 = document.getElementById('order-filtering').value;
                    _context21.n = _t3 === "Intial" ? 1 : _t3 === "over-300" ? 2 : _t3 === "Must-buy" ? 3 : _t3 === "Can-buy" ? 4 : _t3 === "Can-make" ? 5 : 6;
                    break;
                  case 1:
                    console.log("no changes made");
                    amountFilter.clear();
                    buyOrMakeFilter.clear();
                    return _context21.a(3, 7);
                  case 2:
                    amountFilter.clear();
                    buyOrMakeFilter.clear();
                    amountFilter.applyCustomFilter(">=300");
                    return _context21.a(3, 7);
                  case 3:
                    amountFilter.clear();
                    buyOrMakeFilter.clear();
                    buyOrMakeFilter.applyValuesFilter(["Must Buy"]);
                    return _context21.a(3, 7);
                  case 4:
                    amountFilter.clear();
                    buyOrMakeFilter.clear();
                    buyOrMakeFilter.applyValuesFilter(["Can Buy"]);
                    return _context21.a(3, 7);
                  case 5:
                    amountFilter.clear();
                    buyOrMakeFilter.clear();
                    buyOrMakeFilter.applyValuesFilter(["Can Make"]);
                    return _context21.a(3, 7);
                  case 6:
                    console.log("No valid filter selected");
                    return _context21.a(3, 7);
                  case 7:
                    _context21.n = 8;
                    return context.sync();
                  case 8:
                    return _context21.a(2);
                }
              }, _callee20);
            }));
            return function (_x12) {
              return _ref1.apply(this, arguments);
            };
          }());
        case 1:
          return _context22.a(2);
      }
    }, _callee21);
  }));
  return _filteringDropdown.apply(this, arguments);
}
function invFilteringDropdown() {
  return _invFilteringDropdown.apply(this, arguments);
}
function _invFilteringDropdown() {
  _invFilteringDropdown = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee23() {
    return _regenerator().w(function (_context24) {
      while (1) switch (_context24.n) {
        case 0:
          _context24.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref10 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee22(context) {
              var inventoryWorksheet, inventoryTable, qtyNeededFilter, _t4;
              return _regenerator().w(function (_context23) {
                while (1) switch (_context23.n) {
                  case 0:
                    inventoryWorksheet = context.workbook.worksheets.getItem("Inventory At");
                    inventoryTable = inventoryWorksheet.tables.getItem("InventoryAtTable");
                    qtyNeededFilter = inventoryTable.columns.getItem("Qty Needed (MEB)").filter;
                    _t4 = document.getElementById('inventory-filtering').value;
                    _context23.n = _t4 === "Intial case" ? 1 : _t4 === "over-300" ? 2 : 3;
                    break;
                  case 1:
                    console.log("no changes made");
                    qtyNeededFilter.clear();
                    return _context23.a(3, 4);
                  case 2:
                    qtyNeededFilter.clear();
                    qtyNeededFilter.applyCustomFilter(">=300");
                    return _context23.a(3, 4);
                  case 3:
                    console.log("No valid filter selected");
                    return _context23.a(3, 4);
                  case 4:
                    _context23.n = 5;
                    return context.sync();
                  case 5:
                    return _context23.a(2);
                }
              }, _callee22);
            }));
            return function (_x13) {
              return _ref10.apply(this, arguments);
            };
          }());
        case 1:
          return _context24.a(2);
      }
    }, _callee23);
  }));
  return _invFilteringDropdown.apply(this, arguments);
}
function formatDate(dateString) {
  var date = new Date(dateString);
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).replace(/\//g, '-');
}