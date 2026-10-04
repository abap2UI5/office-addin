// A stand-in for Office.js in the smoke test: the tests answer
// https://officeapis.public.onecdn.static.microsoft/1/office.js with this
// file. It models only what the host page and z2ui5.cc.ExcelBridge use -
// Office.onReady( ), Office.HostType, Office.context.host and the parts of
// Excel.run( ) that write a range, read the selection and subscribe to
// selection changes - and records every call in window.__excel, where the
// tests read it.
//
// window.__stubHost (set by the test before the page loads) is what
// Office.onReady( ) reports: "Excel" by default, null for "opened outside
// Office". Like the real Office.js it sets history.pushState and
// replaceState to null - office-init.js has to put them back.
(function () {
  "use strict";
  var host = window.__stubHost === undefined ? "Excel" : window.__stubHost;

  window.history.pushState = null;
  window.history.replaceState = null;

  function column(n) {
    var s = "";
    for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) {
      s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    }
    return s;
  }
  function sheetRef(name) {
    return /^[A-Za-z0-9_]+$/.test(name) ? name : "'" + name + "'";
  }

  var state = {
    sheets: ["Sheet1"],
    writes: [],
    tables: [],
    handlers: [],
    failNext: null,
    selection: {
      address: "Sheet1!A1:C2",
      values: [
        ["MAT-1", 1, 45292],
        ["MAT-2", 2.5, 45293],
      ],
      numberFormat: [
        ["@", "General", "yyyy-mm-dd"],
        ["@", "General", "yyyy-mm-dd"],
      ],
    },
    // the user selects something else in the sheet
    select: function (address) {
      state.selection.address = address;
      state.handlers.forEach(function (h) {
        if (!h.removed) h.fn({ address: address });
      });
    },
  };
  window.__excel = state;

  function context() {
    var pending = [];
    var ctx = {
      sync: function () {
        var run = pending;
        pending = [];
        return Promise.resolve().then(function () {
          run.forEach(function (fn) {
            fn();
          });
        });
      },
    };
    function rangeOn(sheet, row, col) {
      return {
        worksheet: { name: sheet },
        getResizedRange: function (rows, cols) {
          var range = {
            numberFormat: null,
            values: null,
            address: "",
            format: { autofitColumns: function () {} },
            load: function () {
              pending.push(function () {
                if (state.failNext) {
                  var e = state.failNext;
                  state.failNext = null;
                  throw e;
                }
                range.address =
                  sheetRef(sheet) +
                  "!" +
                  column(col) +
                  (row + 1) +
                  ":" +
                  column(col + cols) +
                  (row + rows + 1);
                state.writes.push({
                  sheet: sheet,
                  address: range.address,
                  values: range.values,
                  numberFormat: range.numberFormat,
                });
              });
            },
          };
          return range;
        },
      };
    }
    function sheetObject(name) {
      return {
        name: name,
        activate: function () {},
        getRange: function () {
          return rangeOn(name, 0, 0);
        },
        tables: {
          add: function (range, hasHeaders) {
            state.tables.push({ sheet: name, hasHeaders: hasHeaders });
          },
        },
      };
    }
    function selectedRange() {
      var sel = state.selection;
      var range = {
        address: sel.address,
        rowCount: sel.values.length,
        columnCount: sel.values[0].length,
        values: sel.values,
        numberFormat: sel.numberFormat,
        text: sel.values.map(function (r) {
          return r.map(String);
        }),
        isNullObject: false,
        load: function () {},
        getCell: function () {
          return { worksheet: sheetObject("Sheet1") };
        },
        getUsedRangeOrNullObject: function () {
          return range;
        },
      };
      return range;
    }
    ctx.workbook = {
      worksheets: {
        add: function (name) {
          var sheetName = name || "Sheet" + (state.sheets.length + 1);
          state.sheets.push(sheetName);
          return sheetObject(sheetName);
        },
        getItem: sheetObject,
        getActiveWorksheet: function () {
          return sheetObject("Sheet1");
        },
      },
      getSelectedRange: selectedRange,
      onSelectionChanged: {
        add: function (fn) {
          var handler = {
            fn: fn,
            context: ctx,
            removed: false,
            remove: function () {
              handler.removed = true;
            },
          };
          state.handlers.push(handler);
          return handler;
        },
      },
    };
    return ctx;
  }

  window.Excel = {
    run: function (a, b) {
      var fn = typeof a === "function" ? a : b;
      var ctx = typeof a === "function" ? context() : a;
      return Promise.resolve().then(function () {
        return fn(ctx);
      });
    },
  };

  window.Office = {
    HostType: { Excel: "Excel", Word: "Word" },
    PlatformType: { PC: "PC", OfficeOnline: "OfficeOnline", Mac: "Mac" },
    context: { host: host },
    onReady: function (cb) {
      var info = { host: host, platform: host ? "PC" : null };
      if (cb) cb(info);
      return Promise.resolve(info);
    },
  };
})();
