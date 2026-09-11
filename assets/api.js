(function () {
  "use strict";
  var origin = window.location.origin;
  var isPages = /\.github\.io$/i.test(window.location.hostname || "");
  if (window.location.protocol === "file:" || !origin || origin === "null") {
    window.SGR_API = "http://127.0.0.1:5055";
  } else {
    window.SGR_API = "";
  }

  var STATIC_GET = {
    "/api/rendimientos/resultados": "data/rendimientos.json",
    "/api/contingente/resultados": "data/contingente_procesado.json",
    "/api/calendario-rentas/resultados": "data/calendario_rentas.json",
    "/api/calendario-global/resultados": "data/calendario_global.json",
    "/api/conciliacion/resultados": "data/conciliacion.json"
  };

  window.sgrFetch = function (path, options) {
    options = options || {};
    if (isPages) {
      var method = String(options.method || "GET").toUpperCase();
      var clean = String(path || "").split("?")[0];
      var rel = STATIC_GET[clean];
      if (method === "GET" && rel) {
        return fetch(rel, { cache: "no-store" });
      }
      return Promise.resolve(
        new Response(
          JSON.stringify({
            error: "En el link público se puede consultar. Procesar y exportar se hace en la app de la PC."
          }),
          { status: 400, headers: { "Content-Type": "application/json" } }
        )
      );
    }
    return fetch((window.SGR_API || "") + path, options);
  };
})();
