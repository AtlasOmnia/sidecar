/* Synchronous, dependency-free pre-paint bootstrap shared by startup and settings. */
(function (win) {
  "use strict";
  var palettes = ["teal", "blue", "gold", "red"];
  var modes = ["auto", "dark", "light"];
  var current;
  var bound = false;

  function normalize(value) {
    value = value && typeof value === "object" ? value : {};
    return {
      palette: palettes.indexOf(value.palette) >= 0 ? value.palette : "teal",
      mode: modes.indexOf(value.mode) >= 0 ? value.mode : "auto"
    };
  }
  function load() {
    try { return normalize(JSON.parse(win.localStorage.getItem("omnia.theme"))); }
    catch (_) { return normalize(); } // Storage may be blocked by the Office host.
  }
  function scheme(theme) {
    if (theme.mode !== "auto") return theme.mode;
    try {
      var office = win.Office && win.Office.context && win.Office.context.officeTheme;
      if (office && typeof office.isDarkTheme === "boolean") return office.isDarkTheme ? "dark" : "light";
      var color = office && office.bodyBackgroundColor;
      if (typeof color === "string" && /^#?(?:[0-9a-f]{3}|[0-9a-f]{6})$/i.test(color)) {
        color = color.replace(/^#/, "");
        if (color.length === 3) color = color.split("").map(function (c) { return c + c; }).join("");
        var rgb = [0, 2, 4].map(function (offset) {
          var channel = parseInt(color.slice(offset, offset + 2), 16) / 255;
          return channel <= 0.04045 ? channel / 12.92 : Math.pow((channel + 0.055) / 1.055, 2.4);
        });
        return 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2] < 0.5 ? "dark" : "light";
      }
    } catch (_) { /* Office theme can throw before Office.onReady. */ }
    try { return win.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light"; }
    catch (_) { return "light"; }
  }
  function syncControls() {
    ["palette", "mode"].forEach(function (kind) {
      win.document.querySelectorAll("[data-set-" + kind + "]").forEach(function (button) {
        var selected = button.getAttribute("data-set-" + kind) === current[kind];
        button.setAttribute("aria-checked", String(selected));
        button.tabIndex = selected ? 0 : -1;
      });
    });
  }
  function apply(theme) {
    current = normalize(theme);
    win.document.documentElement.dataset.palette = current.palette;
    win.document.documentElement.dataset.scheme = scheme(current);
    syncControls();
    return current;
  }
  function refresh() { return apply(current || load()); }
  function save(theme) {
    apply(theme);
    try { win.localStorage.setItem("omnia.theme", JSON.stringify(current)); }
    catch (_) { /* Keep the user's choice for this pane even if persistence is denied. */ }
    return current;
  }
  function bindControls() {
    if (bound) return;
    bound = true;
    ["palette", "mode"].forEach(function (kind) {
      var buttons = Array.from(win.document.querySelectorAll("[data-set-" + kind + "]"));
      buttons.forEach(function (button, index) {
        button.addEventListener("click", function () {
          var next = { palette: current.palette, mode: current.mode };
          next[kind] = button.getAttribute("data-set-" + kind);
          save(next);
        });
        button.addEventListener("keydown", function (event) {
          var next;
          if (event.key === "ArrowRight" || event.key === "ArrowDown") next = (index + 1) % buttons.length;
          else if (event.key === "ArrowLeft" || event.key === "ArrowUp") next = (index + buttons.length - 1) % buttons.length;
          else if (event.key === "Home") next = 0;
          else if (event.key === "End") next = buttons.length - 1;
          else return;
          event.preventDefault();
          buttons[next].click();
          buttons[next].focus();
        });
      });
    });
    try {
      var media = win.matchMedia("(prefers-color-scheme: dark)");
      if (media.addEventListener) media.addEventListener("change", refresh);
      else if (media.addListener) media.addListener(refresh);
    } catch (_) { /* Older hosts may not implement matchMedia. */ }
    // Word has no OfficeThemeChanged event (that event is Outlook-only).
    // Refresh while visible and on focus, without changing any Office registration.
    win.addEventListener("focus", refresh);
    win.document.addEventListener("visibilitychange", refresh);
    var timer = win.setInterval(function () {
      if (!win.document.hidden && current.mode === "auto") refresh();
    }, 1000);
    win.addEventListener("pagehide", function () { win.clearInterval(timer); }, { once: true });
    refresh();
  }
  win.OmniaTheme = { load: load, apply: apply, save: save, refresh: refresh, bindControls: bindControls };
  apply(load());
})(window);
