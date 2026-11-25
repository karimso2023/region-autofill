// RegionAutofill core matching logic - pure JS, no Office dependencies.
// Mirrors the validated pilot rules exactly.
(function (global) {
  "use strict";

  function norm(s) {
    if (typeof s !== "string") s = s == null ? "" : String(s);
    s = s.trim();
    s = s.split("أ").join("ا").split("إ").join("ا").split("آ").join("ا");
    s = s.split("ة").join("ه").split("ى").join("ي");
    s = s.split("ئ").join("ي").split("ؤ").join("و");
    s = s.split("ث").join("ت");
    return s.replace(/\s+/g, " ");
  }

  function stripAl(s) { return s.indexOf("ال") === 0 ? s.slice(2) : s; }
  function tailVariant(s) { return s.endsWith("ا") ? s.slice(0, -1) + "ي" : ""; }

  // paddedAddr must be " " + norm(address) + " "
  function matchRegion(paddedAddr, region) {
    var r = norm(region);
    if (!r) return false;
    if (paddedAddr.indexOf(r) !== -1) return true;
    var r2 = stripAl(r);
    if (r2 !== r && paddedAddr.indexOf(" " + r2 + " ") !== -1) return true;
    var t = tailVariant(r);
    if (t && paddedAddr.indexOf(t) !== -1) return true;
    var t2 = tailVariant(r2);
    if (t2 && paddedAddr.indexOf(" " + t2 + " ") !== -1) return true;
    return false;
  }

  // regionMap: { governorateName: [regionName, ...] } exactly as in the الأقاليم sheet
  function findGovKey(regionMap, gov) {
    var g = norm(gov);
    for (var k in regionMap) {
      if (!Object.prototype.hasOwnProperty.call(regionMap, k)) continue;
      var kk = norm(k);
      if (kk === g || kk === stripAl(g) || stripAl(kk) === stripAl(g)) return k;
    }
    return null;
  }

  // Returns { action: 'fill'|'note'|'note-clear'|'keep-clear'|'keep-silent'|'skip', value }
  // existingF: current region cell content; stateValue: the add-in's per-row marker
  // (the value it auto-filled into this row, stored in a hidden marker column), or null.
  // 'note-clear': current F is a stale auto-fill - the caller clears F and writes the note.
  function decide(regionMap, gov, addr, existingF, stateValue) {
    if (!norm(addr)) return { action: "skip", value: null };
    var existing = norm(existingF);
    var isAuto = stateValue != null && norm(stateValue) === existing && existing !== "";
    function flag(text) { return { action: isAuto ? "note-clear" : "note", value: text }; }
    var key = findGovKey(regionMap, gov);
    if (key === null) return flag("المحافظة غير موجودة في صفحة الأقاليم. راجع اسم المحافظة.");
    var padded = " " + norm(addr) + " ";
    var hits = regionMap[key].filter(function (r) { return matchRegion(padded, r); });
    var g = norm(gov);
    var echo = hits.filter(function (h) { return norm(h) === g || g.indexOf(norm(h)) !== -1; });
    var rest = hits.filter(function (h) { return echo.indexOf(h) === -1; });
    var winners = rest.filter(function (h) {
      return !rest.some(function (h2) { return h2 !== h && norm(h2).indexOf(norm(h)) !== -1; });
    });

    if (existing && !isAuto) {
      // Manually chosen region: never overwrite
      var match = winners.some(function (w) { return norm(w) === existing; });
      if (match) return { action: "keep-clear", value: null };
      if (winners.length > 0) {
        return { action: "note", value: "المنطقة الحالية مختلفة عن مطابقة العنوان: " + winners.join(" أو ") + ". راجع." };
      }
      return { action: "keep-silent", value: null };
    }

    if (winners.length === 1) return { action: "fill", value: winners[0] };
    if (winners.length > 1) {
      return flag("غامض: العنوان طابق أكتر من منطقة: " + winners.join(" أو ") + ". راجع أيهما المقصود.");
    }
    if (echo.length > 0) {
      return flag("المطابقة الوحيدة هي اسم المحافظة نفسها (" + echo[0] + "). لو العنوان فيه منطقة تانية، راجع يدويًا.");
    }
    return flag("العنوان لم يطابق أي منطقة من قائمة المحافظة المعتمدة. راجع يدويًا.");
  }

  var api = { norm: norm, stripAl: stripAl, matchRegion: matchRegion, findGovKey: findGovKey, decide: decide };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else global.RegionLogic = api;
})(typeof window !== "undefined" ? window : globalThis);
