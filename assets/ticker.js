(function () {
  "use strict";
  var BASE = "https://data912.com";

  function formatPrice(p) {
    if (p == null || isNaN(p)) return "—";
    if (p >= 1000) return p.toLocaleString("es-AR", { maximumFractionDigits: 0, minimumFractionDigits: 0 });
    if (p >= 1) return p.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return p.toLocaleString("es-AR", { minimumFractionDigits: 4, maximumFractionDigits: 4 });
  }

  function formatPct(p) {
    if (p == null || isNaN(p)) return "";
    var s = (p >= 0 ? "+" : "") + p.toFixed(2) + "%";
    return s;
  }

  function buildItems(bonds, stocks, mepMark) {
    var items = [];
    var seen = {};
    function add(symbol, price, pct, label) {
      if (seen[symbol] || price == null || isNaN(price)) return;
      seen[symbol] = true;
      items.push({ symbol: symbol, price: price, pct: pct, label: label || symbol });
    }
    if (bonds && bonds.length) {
      var main = ["AL30", "GD30", "AL29", "GD29", "AL35", "GD35", "GD38", "GD41"];
      main.forEach(function (s) {
        var b = bonds.find(function (x) { return (x.symbol || x.ticker) === s; });
        if (b) add(b.symbol || b.ticker, b.c != null ? b.c : b.px_bid, b.pct_change);
      });
      bonds.slice(0, 12).forEach(function (b) {
        var s = b.symbol || b.ticker;
        if (!main.includes(s)) add(s, b.c != null ? b.c : b.px_bid, b.pct_change);
      });
    }
    if (stocks && stocks.length) {
      ["GGAL", "PAMP", "YPFD", "ALUA", "BMA", "TXAR", "SUPV"].forEach(function (s) {
        var b = stocks.find(function (x) { return (x.symbol || x.ticker) === s; });
        if (b) add(b.symbol || b.ticker, b.c != null ? b.c : b.px_bid, b.pct_change);
      });
    }
    if (mepMark != null && !isNaN(mepMark)) {
      items.push({ symbol: "MEP", price: mepMark, pct: null, label: "MEP" });
    }
    return items;
  }

  function render(barEl, items) {
    if (!items.length) {
      barEl.innerHTML = "<span class=\"ticker-item\"><span class=\"ticker-symbol\">Cargando…</span></span>";
      return;
    }
    var html = "";
    items.forEach(function (it) {
      var pctClass = it.pct != null ? (it.pct >= 0 ? "positive" : "negative") : "";
      var pctStr = it.pct != null ? "<span class=\"ticker-pct " + pctClass + "\">" + formatPct(it.pct) + "</span>" : "";
      html += "<span class=\"ticker-item\"><span class=\"ticker-symbol\">" + (it.label || it.symbol) + "</span><span class=\"ticker-price\">" + formatPrice(it.price) + "</span>" + pctStr + "</span>";
    });
    barEl.innerHTML = html + html;
  }

  function run() {
    var barEl = document.getElementById("ticker-bar");
    if (!barEl) return;

    var wrap = barEl.closest(".ticker-wrap");
    if (wrap) document.body.classList.add("has-ticker");

    function done(bonds, stocks, mepMark) {
      var items = buildItems(bonds, stocks, mepMark);
      render(barEl, items);
    }

    Promise.all([
      fetch(BASE + "/live/arg_bonds").then(function (r) { return r.ok ? r.json() : []; }).catch(function () { return []; }),
      fetch(BASE + "/live/arg_stocks").then(function (r) { return r.ok ? r.json() : []; }).catch(function () { return []; }),
      fetch(BASE + "/live/mep").then(function (r) { return r.ok ? r.json() : []; }).catch(function () { return []; })
    ]).then(function (results) {
      var mepMark = (results[2] && results[2][0] && results[2][0].mark != null) ? results[2][0].mark : null;
      done(results[0], results[1], mepMark);
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", run);
  } else {
    run();
  }
})();
