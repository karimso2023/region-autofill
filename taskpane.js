/* RegionAutofill task pane.
   - Watches edits in المحافظه and العنوان on sheets that look like the shipping sheet.
   - Fills اسم المنطقه only on a confident single match; otherwise writes a review note.
   - Auto-filled values are tracked by a hidden per-row marker column (RegionAutoKey):
     the marker lives on the same row, so sorting / inserting / deleting rows keeps
     the auto-vs-manual distinction attached to the row's own content.
   - Editing an address after an auto-fill updates the region; if nothing matches
     anymore, the stale auto value is cleared and a review note is written instead.
   - A region typed by hand is never overwritten; a conflict gets a review note.
   - Columns are located by their header text on every run, so inserted or moved
     columns do not break the mapping.
   - Works fully offline inside the workbook; no workbook data leaves the document.
*/
(function () {
  "use strict";

  var REGIONS_SHEET = "الأقاليم";
  var NOTE_HDR = "ملاحظة مراجعة المنطقة";
  var MARKER_HDR = "RegionAutoKey";
  var watched = {}; // sheetId -> true

  var statusEl = document.getElementById("status");
  var logEl = document.getElementById("log");

  function setStatus(text, cls) {
    statusEl.textContent = text;
    statusEl.className = cls || "";
  }
  function log(text) {
    var d = document.createElement("div");
    d.textContent = new Date().toLocaleTimeString("ar-EG") + " - " + text;
    logEl.insertBefore(d, logEl.firstChild);
  }

  function colLetter(i) { // 0-based -> A1 letter
    var s = "";
    i = i + 1;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }

  function isShippingHeaders(h) {
    // h: row-1 values array; the shipping sheet's fixed E/F/J headers
    var L = RegionLogic.norm;
    return L(h[4]) === L("المحافظه") && L(h[5] || "").indexOf(L("المنطقه")) !== -1 && L(h[9]) === L("العنوان");
  }

  // ---- column resolution --------------------------------------------------

  // Locates the working columns by header text. When create is true, missing
  // note/marker columns are appended after the last used header and the marker
  // column is hidden (best effort). Returns 0-based indexes + last used row.
  function resolveColumns(ctx, sheet, create) {
    var used = sheet.getUsedRange(true);
    used.load("values,rowCount");
    return ctx.sync().then(function () {
      var vals = used.values || [];
      var h = (vals[0] || []).slice();
      var L = RegionLogic.norm;
      function find(pred) {
        for (var i = 0; i < h.length; i++) {
          if (h[i] !== null && h[i] !== undefined && h[i] !== "" && pred(String(h[i]))) return i;
        }
        return -1;
      }
      var cols = {
        gov: find(function (x) { return L(x) === L("المحافظه"); }),
        region: find(function (x) { return L(x) === L("اسم المنطقه"); }),
        addr: find(function (x) { return L(x) === L("العنوان"); }), // first match: recipient J, not sender AB
        note: find(function (x) { return L(x) === L(NOTE_HDR); }),
        marker: find(function (x) { return x === MARKER_HDR; }),
        lastRow: used.rowCount
      };
      if (!create || cols.gov < 0 || cols.region < 0 || cols.addr < 0) return cols;
      var writes = [];
      if (cols.note < 0) { cols.note = h.length; h.push(NOTE_HDR); writes.push([cols.note, NOTE_HDR]); }
      if (cols.marker < 0) { cols.marker = h.length; h.push(MARKER_HDR); writes.push([cols.marker, MARKER_HDR]); }
      if (!writes.length) return cols;
      writes.forEach(function (w) { sheet.getRange(colLetter(w[0]) + "1").values = [[w[1]]]; });
      return ctx.sync().then(function () {
        // best effort: hide the marker column; harmless if the host rejects it
        try {
          sheet.getRange(colLetter(cols.marker) + ":" + colLetter(cols.marker)).columnHidden = true;
          return ctx.sync().then(function () { return cols; }, function () { return cols; });
        } catch (e) { return cols; }
      });
    });
  }

  // ---- Excel wiring -------------------------------------------------------

  function qualifyAndWatch() {
    return Excel.run(function (ctx) {
      var sheets = ctx.workbook.worksheets;
      sheets.load("items/name,id");
      return ctx.sync().then(function () {
        var hasRegions = sheets.items.some(function (s) { return s.name === REGIONS_SHEET; });
        if (!hasRegions) {
          setStatus("الملف ده مفيهوش صفحة «" + REGIONS_SHEET + "» - الإضافة بتشتغل مع شيت الشحن بس.", "warn");
          return;
        }
        var chain = Promise.resolve();
        var found = 0;
        sheets.items.forEach(function (s) {
          chain = chain.then(function () {
            return Excel.run(function (c2) {
              var sh = c2.workbook.worksheets.getItem(s.id);
              var hdr = sh.getRange("A1:J1");
              hdr.load("values");
              return c2.sync().then(function () {
                if (isShippingHeaders(hdr.values[0])) {
                  sh.onChanged.add(onSheetChanged);
                  watched[s.id] = true;
                  found++;
                  log("مراقبة شيت: " + s.name);
                }
              });
            });
          });
        });
        return chain.then(function () {
          if (found > 0) setStatus("شغّالة: أي تعديل في المحافظة أو العنوان هيملى المنطقة تلقائيًا. الحالات الغامضة هيتكتب لها ملاحظة في عمود «" + NOTE_HDR + "».", "ok");
          else setStatus("الشيتات المفتوحة مش بشكل شيت الشحن (المحافظه/اسم المنطقه/العنوان). افتح شيت الشحن وجرب زر التعبئة.", "warn");
        });
      });
    }).catch(showError);
  }

  function loadRegionMap(ctx) {
    var reg = ctx.workbook.worksheets.getItem(REGIONS_SHEET);
    var used = reg.getUsedRange(true);
    used.load("values");
    return ctx.sync().then(function () {
      var vals = used.values;
      var map = {};
      if (!vals.length) return map;
      var hdr = vals[0];
      for (var j = 0; j < hdr.length; j++) {
        var h = hdr[j];
        if (h && typeof h === "string") {
          var list = [];
          for (var i = 1; i < vals.length; i++) {
            var v = vals[i][j];
            if (v !== null && v !== undefined && String(v).trim() !== "") list.push(String(v).trim());
          }
          map[h.trim()] = list;
        }
      }
      return map;
    });
  }

  function processRows(sheetName, rowNumbers) {
    return Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getItem(sheetName);
      var regionMapP = loadRegionMap(ctx);
      var colsP = resolveColumns(ctx, sheet, true);
      return Promise.all([regionMapP, colsP]).then(function (rs) {
        var regionMap = rs[0], cols = rs[1];
        if (cols.gov < 0 || cols.region < 0 || cols.addr < 0) {
          setStatus("مش لاقي أعمدة المحافظه/اسم المنطقه/العنوان في شيت «" + sheetName + "».", "warn");
          return;
        }
        var gL = colLetter(cols.gov), rL = colLetter(cols.region), aL = colLetter(cols.addr),
            nL = colLetter(cols.note), mL = colLetter(cols.marker);
        var chain = Promise.resolve();
        var counts = { fill: 0, note: 0, keep: 0 };
        rowNumbers.forEach(function (row) {
          chain = chain.then(function () {
            var g = sheet.getRange(gL + row), a = sheet.getRange(aL + row),
                f = sheet.getRange(rL + row), m = sheet.getRange(mL + row);
            g.load("values"); a.load("values"); f.load("values"); m.load("values");
            return ctx.sync().then(function () {
              var d = RegionLogic.decide(regionMap, g.values[0][0], a.values[0][0], f.values[0][0], m.values[0][0]);
              var note = sheet.getRange(nL + row);
              if (d.action === "fill") {
                f.values = [[d.value]];
                m.values = [[d.value]]; // marker: this row's value is ours
                note.clear(Excel.ClearApplyTo.contents);
                counts.fill++;
                log("صف " + row + ": اتملت «" + d.value + "»");
              } else if (d.action === "note-clear") {
                f.clear(Excel.ClearApplyTo.contents); // stale auto value removed, never kept silently
                m.clear(Excel.ClearApplyTo.contents);
                note.values = [[d.value]];
                counts.note++;
                log("صف " + row + ": القيمة القديمة اتمسحت وللمراجعة - " + d.value);
              } else if (d.action === "note") {
                m.clear(Excel.ClearApplyTo.contents);
                note.values = [[d.value]];
                counts.note++;
                log("صف " + row + ": للمراجعة - " + d.value);
              } else if (d.action === "keep-clear") {
                note.clear(Excel.ClearApplyTo.contents);
                m.clear(Excel.ClearApplyTo.contents); // value is manual now
                counts.keep++;
              } else if (d.action === "keep-silent") {
                m.clear(Excel.ClearApplyTo.contents);
                counts.keep++;
              }
              return ctx.sync();
            });
          });
        });
        return chain.then(function () {
          log("خلصت: " + counts.fill + " اتملوا، " + counts.note + " للمراجعة، " + counts.keep + " يدوي اتساب.");
        });
      });
    }).catch(showError);
  }

  function onSheetChanged(ev) {
    if (!watched[ev.worksheetId]) return;
    return Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getItem(ev.worksheetId);
      sheet.load("name");
      var addr = sheet.getRange(ev.address);
      addr.load("rowIndex,columnIndex,rowCount,columnCount");
      return ctx.sync().then(function () {
        var c0 = addr.columnIndex, c1 = addr.columnIndex + addr.columnCount - 1;
        var r0 = addr.rowIndex + 1, r1 = addr.rowIndex + addr.rowCount; // 1-based
        return resolveColumns(ctx, sheet, false).then(function (cols) {
          if (cols.gov < 0 || cols.addr < 0) return;
          var touches = (c0 <= cols.gov && c1 >= cols.gov) || (c0 <= cols.addr && c1 >= cols.addr);
          if (!touches) return;
          var rows = {};
          for (var r = Math.max(2, r0); r <= r1; r++) rows[r] = true;
          var list = Object.keys(rows).map(Number);
          if (!list.length) return;
          return processRows(sheet.name, list);
        });
      });
    }).catch(showError);
  }

  function fillAll() {
    return Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      return ctx.sync().then(function () {
        return resolveColumns(ctx, sheet, false).then(function (cols) {
          if (cols.gov < 0 || cols.region < 0 || cols.addr < 0) {
            setStatus("الشيت الحالي مش شيت الشحن.", "warn");
            return;
          }
          var rows = [];
          for (var r = 2; r <= cols.lastRow; r++) rows.push(r);
          return processRows(sheet.name, rows);
        });
      });
    }).catch(showError);
  }

  function showError(err) {
    setStatus("حصل خطأ: " + (err && err.message ? err.message : err), "warn");
  }

  document.getElementById("fillAll").addEventListener("click", function () {
    setStatus("بشتغل…", "");
    fillAll().then(function () { setStatus("تمت التعبئة. راجع عمود الملاحظات.", "ok"); });
  });

  // Prepare a carrier-upload copy: some carrier portals reject extra columns,
  // so this deletes the helper note/marker columns from the CURRENT sheet.
  // Meant for the upload copy only; on the master workbook it also drops
  // auto-fill tracking and review notes, so it asks first.
  function cleanupHelpers() {
    if (!confirm("ده هيحذف عمود «" + NOTE_HDR + "» وعمود العلامة من الشيت الحالي نهائيًا، والتتبع التلقائي هيقف للصفوف المتملية. استخدمه على نسخة الرفع لشركة الشحن بس، مش على ملفك الأساسي. متأكد؟")) return;
    setStatus("بحذف أعمدة المساعدة…", "");
    Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      return resolveColumns(ctx, sheet, false).then(function (cols) {
        var toDelete = [cols.note, cols.marker].filter(function (c) { return c >= 0; }).sort(function (a, b) { return b - a; });
        if (!toDelete.length) { setStatus("مفيش أعمدة مساعدة في الشيت الحالي.", "warn"); return; }
        toDelete.forEach(function (c) {
          var L = colLetter(c);
          sheet.getRange(L + ":" + L).delete(Excel.DeleteShiftDirection.left);
        });
        return ctx.sync().then(function () {
          log("اتحذفت أعمدة المساعدة من الشيت الحالي - جاهز للرفع.");
          setStatus("تم: الشيت الحالي بقى بأعمدة شركة الشحن بس.", "ok");
        });
      });
    }).catch(showError);
  }
  var cleanupBtn = document.getElementById("cleanup");
  if (cleanupBtn) cleanupBtn.addEventListener("click", cleanupHelpers);

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function () {
      qualifyAndWatch();
    });
  } else if (typeof Excel !== "undefined") {
    // demo/test harness with a stubbed Excel API
    qualifyAndWatch();
  } else {
    setStatus("الصفحة دي بتشتغل جوّه Excel بس.", "warn");
  }
})();
