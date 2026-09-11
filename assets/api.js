(function () {
  "use strict";
  /** Base URL del backend Flask. Si abrís el HTML directo (file://), apunta a localhost. */
  var origin = window.location.origin;
  if (window.location.protocol === "file:" || !origin || origin === "null") {
    window.SGR_API = "http://127.0.0.1:5055";
  } else {
    window.SGR_API = "";
  }
  window.sgrFetch = function (path, options) {
    return fetch((window.SGR_API || "") + path, options);
  };
})();
