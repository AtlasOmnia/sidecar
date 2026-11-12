/* Classic, parser-blocking script: resolve the theme before styles or body paint. */
(function () {
  "use strict";
  var palettes = ["teal", "blue", "gold", "red"];
  var modes = ["auto", "dark", "light"];
  var media = typeof window.matchMedia === "function" ? window.matchMedia("(prefers-color-scheme: dark)") : null;
  var settings;
  var bound = false;
  var officeTimer;

  function normalize(value) {
    value = value && typeof value === "object" ? value : {};
    return {
      palette: palettes.indexOf(value.palette) >= 0 ? value.palette : "teal",
      mode: modes.indexOf(value.mode) >= 0 ? value.mode : "auto"
    };
  }

  function read() {
    try { return normalize(JSON.parse(window.localStorage.getItem("omnia.theme"))); }
    catch (_) { return normalize(null); }
  }

  function officeScheme() {
    try {
      var theme = window.Office && window.Office.context && window.Office.context.officeTheme;
      if (!theme) return null;
      if (typeof theme.isDarkTheme === "boolean") return theme.isDarkTheme ? "dark" : "light";
      var color = theme.bodyBackgroundColor;
      if (typeof color !== "string") return null;
      color = color.replace(/^#/, "");
      if (/^[a-f\d]{3}$/i.test(color)) color = color.split("").map(function (digit) { return digit + digit; }).join("");
      if (!/^[a-f\d]{6}$/i.test(color)) return null;
      var rgb = [0, 2, 4].map(function (offset) {
        var channel = parseInt(color.slice(offset, offset + 2), 16) / 255;
        return channel <= 0.04045 ? channel / 12.92 : Math.pow((channel + 0.055) / 1.055, 2.4);
      });
      return rgb[0] * 0.2126 + rgb[1] * 0.7152 + rgb[2] * 0.0722 < 0.179 ? "dark" : "light";
    } catch (_) { return null; }
  }

  function syncControls() {
    ["palette", "mode"].forEach(function (kind) {
      document.querySelectorAll("[data-set-" + kind + "]").forEach(function (button) {
        var checked = button.getAttribute("data-set-" + kind) === settings[kind];
        button.setAttribute("aria-checked", String(checked));
        button.tabIndex = checked ? 0 : -1;
      });
    });
  }

  function refresh() {
    var scheme = settings.mode === "auto" ? officeScheme() || (media && media.matches ? "dark" : "light") : settings.mode;
    var root = document.documentElement;
    root.dataset.palette = settings.palette;
    root.dataset.scheme = scheme;
    syncControls();
  }

  function setSettings(value) {
    settings = normalize(value);
    try { window.localStorage.setItem("omnia.theme", JSON.stringify(settings)); }
    catch (_) { /* Theme remains usable in memory if WebView storage is unavailable. */ }
    refresh();
  }

  function bindControls() {
    if (bound) return;
    bound = true;
    ["palette", "mode"].forEach(function (kind) {
      var buttons = Array.from(document.querySelectorAll("[data-set-" + kind + "]"));
      function select(button) {
        var next = normalize(settings);
        next[kind] = button.getAttribute("data-set-" + kind);
        setSettings(next);
      }
      buttons.forEach(function (button, index) {
        button.addEventListener("click", function () { select(button); });
        button.addEventListener("keydown", function (event) {
          var target;
          if (event.key === "ArrowRight" || event.key === "ArrowDown") target = (index + 1) % buttons.length;
          else if (event.key === "ArrowLeft" || event.key === "ArrowUp") target = (index + buttons.length - 1) % buttons.length;
          else if (event.key === "Home") target = 0;
          else if (event.key === "End") target = buttons.length - 1;
          else return;
          event.preventDefault();
          select(buttons[target]);
          buttons[target].focus();
        });
      });
    });
    refresh();
  }

  // PowerPoint has no supported OfficeThemeChanged event. Re-read the inexpensive
  // context property while visible, and on focus/resume, without changing Office state.
  function observeOffice() {
    refresh();
    if (!officeTimer) officeTimer = window.setInterval(function () {
      if (settings.mode === "auto" && !document.hidden) refresh();
    }, 1000);
  }
  function stopObserving() {
    if (officeTimer) window.clearInterval(officeTimer);
    officeTimer = undefined;
  }

  settings = read();
  window.OmniaTheme = {
    getSettings: function () { return normalize(settings); },
    setSettings: setSettings,
    refresh: refresh,
    observeOffice: observeOffice
  };
  refresh();
  if (media && media.addEventListener) media.addEventListener("change", refresh);
  window.addEventListener("focus", refresh);
  document.addEventListener("visibilitychange", refresh);
  window.addEventListener("pagehide", stopObserving);
  window.addEventListener("pageshow", function () { if (window.Office) observeOffice(); });
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", bindControls, { once: true });
  else bindControls();
})();
