#!/usr/bin/env node

import { spawn } from "node:child_process";
import { constants as fsConstants } from "node:fs";
import * as fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(SCRIPT_DIR, "..");
const DEFAULT_BROWSER = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const RESULT_META_NAME = "hermes-taskpane-layout-result";
const RESULT_VERSION = 1;
const RESULT_MARKER_PATTERN = new RegExp(`<meta\\s+name=["']${RESULT_META_NAME}["']\\s+content=["']([A-Za-z0-9+/=]+)["']`, "u");
const TIMEOUT_MS = 30_000;

const CASES = [
  {
    id: "empty-local-280x600-hidden",
    width: 280,
    height: 600,
    content: "empty",
    privacy: "local",
    checkpoint: false,
    maxTextarea: false,
    requiresScroll: false,
  },
  {
    id: "long-cloud-320x480-hidden",
    width: 320,
    height: 480,
    content: "long",
    privacy: "cloud",
    checkpoint: false,
    maxTextarea: false,
    requiresScroll: false,
  },
  {
    id: "unbroken-external-480x720-hidden",
    width: 480,
    height: 720,
    content: "unbroken",
    privacy: "external",
    checkpoint: false,
    maxTextarea: false,
    requiresScroll: false,
  },
  {
    id: "activity-local-320x480-hidden",
    width: 320,
    height: 480,
    content: "activity",
    privacy: "local",
    checkpoint: false,
    maxTextarea: false,
    requiresScroll: false,
  },
  {
    id: "checkpoint-cloud-280x600-shown",
    width: 280,
    height: 600,
    content: "long",
    privacy: "cloud",
    checkpoint: true,
    maxTextarea: false,
    requiresScroll: false,
  },
  {
    id: "checkpoint-external-320x480-shown",
    width: 320,
    height: 480,
    content: "activity",
    privacy: "external",
    checkpoint: true,
    maxTextarea: false,
    requiresScroll: false,
  },
  {
    id: "max-textarea-local-320x300-hidden",
    width: 320,
    height: 300,
    content: "empty",
    privacy: "local",
    checkpoint: false,
    maxTextarea: true,
    requiresScroll: true,
  },
  {
    id: "short-checkpoint-cloud-320x300-shown",
    width: 320,
    height: 300,
    content: "activity",
    privacy: "cloud",
    checkpoint: true,
    maxTextarea: true,
    requiresScroll: true,
  },
  {
    id: "unbroken-external-280x600-hidden",
    width: 280,
    height: 600,
    content: "unbroken",
    privacy: "external",
    checkpoint: false,
    maxTextarea: false,
    requiresScroll: false,
  },
  {
    id: "settings-overlay-local-280x300-open",
    width: 280,
    height: 300,
    content: "empty",
    privacy: "local",
    checkpoint: false,
    maxTextarea: false,
    requiresScroll: false,
    settingsOpen: true,
  },
  {
    id: "empty-local-480x720-hidden",
    width: 480,
    height: 720,
    content: "empty",
    privacy: "local",
    checkpoint: false,
    maxTextarea: false,
    requiresScroll: false,
  },
];

const REQUIRED_SETTINGS_ASSERTIONS = [
  "settings-dialog-open",
  "settings-dialog-horizontally-contained",
  "settings-panel-horizontally-contained",
  "settings-overlay-no-horizontal-clipping",
  "settings-evidence-is-dialog-not-page-chrome",
  "settings-close-reachable-at-dialog-top",
  "settings-menu-and-save-reachable-after-internal-scroll",
  "settings-dialog-content-reaches-bottom",
  "settings-scroll-does-not-substitute-page-chrome",
];

const HARNESS_SCRIPT = String.raw`
(() => {
  const RESULT_META_NAME = "hermes-taskpane-layout-result";
  const RESULT_VERSION = 1;
  const payloadMeta = document.querySelector('meta[name="hermes-taskpane-layout-payload"]');
  const payload = decodeJson(payloadMeta && payloadMeta.content);
  const results = [];
  const browserErrors = [];
  let currentIframe = null;

  window.addEventListener("error", (event) => {
    browserErrors.push({ type: "error", message: String(event.message || "window error") });
  });
  window.addEventListener("unhandledrejection", (event) => {
    browserErrors.push({ type: "unhandledrejection", message: String(event.reason || "unhandled rejection") });
  });

  function decodeJson(encoded) {
    if (!encoded) throw new Error("missing harness payload");
    const binary = atob(encoded);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return JSON.parse(new TextDecoder().decode(bytes));
  }

  function encodeJson(value) {
    const bytes = new TextEncoder().encode(JSON.stringify(value));
    let binary = "";
    for (const byte of bytes) binary += String.fromCharCode(byte);
    return btoa(binary);
  }

  function emitResult(result) {
    const oldMarker = document.querySelector('meta[name="' + RESULT_META_NAME + '"]');
    if (oldMarker) oldMarker.remove();
    const marker = document.createElement("meta");
    marker.name = RESULT_META_NAME;
    marker.content = encodeJson(result);
    document.head.appendChild(marker);
  }

  function assertion(assertions, name, passed, details) {
    assertions.push({ name, passed: Boolean(passed), details: details || {} });
  }

  function describeElement(element) {
    return element.id || element.className || element.tagName.toLowerCase();
  }

  function rectDetails(rect) {
    return {
      left: round(rect.left),
      top: round(rect.top),
      right: round(rect.right),
      bottom: round(rect.bottom),
      width: round(rect.width),
      height: round(rect.height),
    };
  }

  function round(value) {
    return Math.round(value * 100) / 100;
  }

  function isVisible(element, win) {
    const style = win.getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== "none" && style.visibility !== "hidden" && style.visibility !== "collapse" && rect.width > 0 && rect.height > 0;
  }

  function visibleControls(doc, win) {
    return Array.from(doc.querySelectorAll("button, input, textarea")).filter((element) => isVisible(element, win));
  }

  function focusableControls(doc, win) {
    return visibleControls(doc, win).filter((element) => !element.disabled);
  }

  // Keep the fixture path synchronous. The parent owns all mutations and
  // layout is read immediately after the same-document DOM/style writes.
  function settle() {}

  function childElement(doc, selector) {
    const element = doc.querySelector(selector);
    if (!element) throw new Error("missing child fixture element " + selector);
    return element;
  }

  function appendMessage(doc, role, text, lastMarker) {
    const root = doc.createElement("article");
    root.className = "message " + role;
    const bubble = doc.createElement("div");
    bubble.className = "bubble";
    bubble.textContent = text;
    root.appendChild(bubble);
    if (lastMarker) {
      root.dataset.layoutLast = "true";
      bubble.dataset.layoutLast = "true";
    }
    childElement(doc, "#messages").appendChild(root);
    return root;
  }

  function appendCodeMessage(doc, text, lastMarker) {
    const root = doc.createElement("article");
    root.className = "message assistant";
    const bubble = doc.createElement("div");
    bubble.className = "bubble";
    const paragraph = doc.createElement("p");
    paragraph.textContent = "Generated code:";
    const pre = doc.createElement("pre");
    const code = doc.createElement("code");
    code.textContent = text;
    pre.appendChild(code);
    bubble.append(paragraph, pre);
    root.appendChild(bubble);
    if (lastMarker) {
      root.dataset.layoutLast = "true";
      bubble.dataset.layoutLast = "true";
    }
    childElement(doc, "#messages").appendChild(root);
    return root;
  }

  function appendActivityMessage(doc, lastMarker) {
    const root = appendMessage(doc, "assistant", "I inspected the presentation and prepared the requested operation.", false);
    const activity = doc.createElement("details");
    activity.className = "activity";
    activity.open = true;
    const summary = doc.createElement("summary");
    summary.textContent = "Activity";
    const list = doc.createElement("div");
    list.className = "activity-list";
    const labels = [
      "Reading selected slides and preserving the current presentation context",
      "Checking the active bridge session before sending the operation",
      "Formatting a deliberately long status line to exercise narrow layouts",
      "Operation completed without exposing prompt or presentation content",
    ];
    for (let index = 0; index < labels.length; index += 1) {
      const row = doc.createElement("div");
      row.className = "activity-item" + (index === labels.length - 1 ? "" : "");
      const marker = doc.createElement("span");
      marker.textContent = index === labels.length - 1 ? "✓" : "…";
      const label = doc.createElement("span");
      label.textContent = labels[index];
      row.append(marker, label);
      list.appendChild(row);
    }
    activity.append(summary, list);
    root.appendChild(activity);
    if (lastMarker) {
      root.dataset.layoutLast = "true";
      activity.dataset.layoutLast = "true";
    }
    return root;
  }

  function populateConversation(doc, spec) {
    const messages = childElement(doc, "#messages");
    messages.replaceChildren();
    if (spec.content === "long") {
      for (let index = 0; index < 10; index += 1) {
        appendMessage(doc, index % 2 === 0 ? "user" : "assistant", "Message " + (index + 1) + " — this deliberately long conversation exercises independent message scrolling while the composer remains part of the pane chrome. ".repeat(3), false);
      }
      appendMessage(doc, "assistant", "The final assistant response is still reachable after scrolling through the conversation.", true);
    } else if (spec.content === "unbroken") {
      appendMessage(doc, "user", "A".repeat(620), false);
      appendCodeMessage(doc, "constPowerPointOperationPayload=" + "x".repeat(720) + ";", true);
    } else if (spec.content === "activity") {
      appendActivityMessage(doc, false);
      appendMessage(doc, "user", "A follow-up request with enough text to keep the conversation populated.", true);
    }
  }

  function setupCase(doc, spec) {
    const badge = childElement(doc, "#privacy-badge");
    const privacy = payload.privacyLabels[spec.privacy];
    if (!privacy) throw new Error("missing privacy fixture " + spec.privacy);
    badge.textContent = privacy.label;
    badge.title = privacy.tooltip;
    badge.dataset.privacy = spec.privacy;

    childElement(doc, "#bridge-status").textContent = "Office bridge connected";
    childElement(doc, "#bridge-status").className = "bridge-status bridge-connected";
    childElement(doc, "#bridge-session-id").textContent = "bridge-session-0123456789abcdef0123456789abcdef";
    childElement(doc, "#copy-bridge-session").disabled = false;
    childElement(doc, "#checkpoint-bar").classList.toggle("hidden", !spec.checkpoint);
    childElement(doc, "#prompt-input").value = spec.maxTextarea ? "line one\nline two\nline three\nline four\nline five\nline six" : "";
    childElement(doc, "#prompt-input").style.height = spec.maxTextarea ? "180px" : "";
    populateConversation(doc, spec);
    if (spec.settingsOpen) {
      const dialog = childElement(doc, "#settings-dialog");
      if (typeof dialog.showModal !== "function") throw new Error("settings dialog does not support showModal");
      dialog.showModal();
    }
  }

  function basicGeometryAssertions(doc, win, spec, assertions) {
    const root = doc.documentElement;
    const body = doc.body;
    const width = spec.width;
    const height = spec.height;
    const horizontalOverflow = Math.max(root.scrollWidth - root.clientWidth, body.scrollWidth - body.clientWidth);
    assertion(assertions, "no-horizontal-document-overflow", horizontalOverflow <= 1, {
      rootScrollWidth: root.scrollWidth,
      rootClientWidth: root.clientWidth,
      bodyScrollWidth: body.scrollWidth,
      bodyClientWidth: body.clientWidth,
      overflow: round(horizontalOverflow),
    });

    const controls = visibleControls(doc, win).filter((element) => !spec.settingsOpen || !element.closest("#settings-dialog"));
    const outOfBounds = controls
      .map((element) => ({ element, rect: element.getBoundingClientRect() }))
      .filter(({ rect }) => rect.left < -1 || rect.right > width + 1)
      .map(({ element, rect }) => ({ element: describeElement(element), rect: rectDetails(rect) }));
    assertion(assertions, "all-visible-controls-horizontally-contained", outOfBounds.length === 0, { outOfBounds });

    const overlaps = [];
    const controlRects = controls.map((element) => ({ element, rect: element.getBoundingClientRect() }));
    for (let first = 0; first < controlRects.length; first += 1) {
      for (let second = first + 1; second < controlRects.length; second += 1) {
        const a = controlRects[first].rect;
        const b = controlRects[second].rect;
        const overlapWidth = Math.min(a.right, b.right) - Math.max(a.left, b.left);
        const overlapHeight = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
        if (overlapWidth > 1 && overlapHeight > 1) {
          overlaps.push({
            first: describeElement(controlRects[first].element),
            second: describeElement(controlRects[second].element),
            overlapWidth: round(overlapWidth),
            overlapHeight: round(overlapHeight),
          });
        }
      }
    }
    assertion(assertions, "no-visible-control-overlap", overlaps.length === 0, { overlaps });

    const messages = childElement(doc, "#messages");
    const composer = childElement(doc, ".composer");
    const messageRect = messages.getBoundingClientRect();
    const composerRect = composer.getBoundingClientRect();
    const normalHeight = height >= 480;
    if (normalHeight) {
      assertion(assertions, "messages-have-positive-normal-height", messageRect.height > 0, { rect: rectDetails(messageRect) });
      assertion(assertions, "composer-visible-at-normal-height", composerRect.height > 0 && composerRect.top >= -1 && composerRect.bottom <= height + 1, {
        rect: rectDetails(composerRect),
      });
    }

    const checkpoint = childElement(doc, "#checkpoint-bar");
    const checkpointRect = checkpoint.getBoundingClientRect();
    const checkpointHidden = win.getComputedStyle(checkpoint).display === "none";
    assertion(assertions, "privacy-label-rendered", childElement(doc, "#privacy-badge").textContent === payload.privacyLabels[spec.privacy].label, {
      expected: payload.privacyLabels[spec.privacy].label,
      actual: childElement(doc, "#privacy-badge").textContent,
    });
    if (spec.checkpoint) {
      const checkpointControls = ["#checkpoint-stop", "#checkpoint-continue"].map((selector) => childElement(doc, selector));
      assertion(assertions, "shown-checkpoint-controls-visible", !checkpointHidden && checkpointRect.width > 0 && checkpointRect.height > 0 && checkpointControls.every((control) => isVisible(control, win)), {
        checkpoint: rectDetails(checkpointRect),
        controls: checkpointControls.map((control) => rectDetails(control.getBoundingClientRect())),
      });
    } else {
      assertion(assertions, "hidden-checkpoint-leaves-no-flex-spacer", checkpointHidden && checkpointRect.width <= 1 && checkpointRect.height <= 1, {
        checkpoint: rectDetails(checkpointRect),
      });
      if (normalHeight) {
        assertion(assertions, "hidden-checkpoint-adjoins-composer", Math.abs(messageRect.bottom - composerRect.top) <= 1, {
          messagesBottom: round(messageRect.bottom),
          composerTop: round(composerRect.top),
        });
      }
    }

    if (spec.content === "long" || spec.content === "unbroken" || spec.content === "activity") {
      const last = childElement(doc, "[data-layout-last]");
      const beforeScrollTop = messages.scrollTop;
      const hasIndependentOverflow = messages.scrollHeight > messages.clientHeight + 1;
      messages.scrollTop = messages.scrollHeight;
      const lastRect = last.getBoundingClientRect();
      const afterScrollTop = messages.scrollTop;
      const settledMessageRect = messages.getBoundingClientRect();
      assertion(assertions, "long-content-scrolls-independently", hasIndependentOverflow && afterScrollTop >= beforeScrollTop && afterScrollTop > 0, {
        beforeScrollTop: round(beforeScrollTop),
        afterScrollTop: round(afterScrollTop),
        scrollHeight: messages.scrollHeight,
        clientHeight: messages.clientHeight,
      });
      assertion(assertions, "last-content-reachable-inside-message-viewport", lastRect.bottom >= settledMessageRect.top - 1 && lastRect.bottom <= settledMessageRect.bottom + 1, {
        last: rectDetails(lastRect),
        messages: rectDetails(settledMessageRect),
      });
    }

    if (spec.maxTextarea) {
      const textarea = childElement(doc, "#prompt-input");
      const computed = win.getComputedStyle(textarea);
      assertion(assertions, "textarea-retains-existing-maximum", computed.maxHeight === "180px" && textarea.getBoundingClientRect().height <= 180 + 1, {
        computedMaxHeight: computed.maxHeight,
        rect: rectDetails(textarea.getBoundingClientRect()),
      });
    }

    if (normalHeight) {
      const composerAfterMessageScroll = composer.getBoundingClientRect();
      assertion(assertions, "composer-remains-visible-after-message-scroll", composerAfterMessageScroll.top >= -1 && composerAfterMessageScroll.bottom <= height + 1, {
        rect: rectDetails(composerAfterMessageScroll),
      });
    }
  }

  function shortHeightAssertions(doc, win, spec, assertions) {
    if (spec.height > 300 || spec.settingsOpen) return;
    const shell = childElement(doc, ".app-shell");
    const body = doc.body;
    const root = doc.documentElement;
    shell.scrollTop = 0;
    body.scrollTop = 0;
    root.scrollTop = 0;
    const shellScrollHeight = shell.scrollHeight;
    const shellClientHeight = shell.clientHeight;
    const documentScrollHeight = Math.max(body.scrollHeight, root.scrollHeight);
    const documentClientHeight = Math.max(body.clientHeight, root.clientHeight);
    const hasScrollSurface = shellScrollHeight > shellClientHeight + 1 || documentScrollHeight > documentClientHeight + 1;
    assertion(assertions, "short-height-has-deliberate-scroll-surface", !spec.requiresScroll || hasScrollSurface, {
      shellScrollHeight,
      shellClientHeight,
      documentScrollHeight,
      documentClientHeight,
    });

    const focusResults = [];
    for (const control of focusableControls(doc, win).filter((element) => !spec.settingsOpen || !element.closest("#settings-dialog"))) {
      control.focus();
      control.scrollIntoView({ block: "nearest", inline: "nearest" });
      settle();
      const rect = control.getBoundingClientRect();
      const focused = doc.activeElement === control;
      const reachable = focused && rect.top >= -1 && rect.bottom <= spec.height + 1 && rect.left >= -1 && rect.right <= spec.width + 1;
      focusResults.push({ control: describeElement(control), focused, reachable, rect: rectDetails(rect) });
    }
    assertion(assertions, "short-height-controls-focus-and-scroll-into-view", focusResults.every((result) => result.reachable), { focusResults });
  }

  function settingsOverlayAssertions(doc, win, spec, assertions) {
    if (!spec.settingsOpen) return;

    const dialog = childElement(doc, "#settings-dialog");
    const panel = childElement(dialog, ".settings-panel");
    const close = childElement(dialog, "#close-settings-button");
    const menu = childElement(dialog, "menu");
    const save = childElement(dialog, "#save-settings-button");
    const topbar = childElement(doc, ".topbar");
    const composer = childElement(doc, ".composer");
    const initialDialogRect = dialog.getBoundingClientRect();
    const initialPanelRect = panel.getBoundingClientRect();
    const open = dialog.open && isVisible(dialog, win);
    const pageChromeBeforeScroll = [topbar, composer].map((element) => rectDetails(element.getBoundingClientRect()));

    assertion(assertions, "settings-dialog-open", open, {
      open: dialog.open,
      dialog: rectDetails(initialDialogRect),
    });
    assertion(assertions, "settings-dialog-horizontally-contained", open && initialDialogRect.left >= -1 && initialDialogRect.right <= spec.width + 1, {
      dialog: rectDetails(initialDialogRect),
      viewport: { width: spec.width, height: spec.height },
    });
    assertion(assertions, "settings-panel-horizontally-contained", open && initialPanelRect.left >= initialDialogRect.left - 1 && initialPanelRect.right <= initialDialogRect.right + 1, {
      dialog: rectDetails(initialDialogRect),
      panel: rectDetails(initialPanelRect),
    });
    assertion(assertions, "settings-overlay-no-horizontal-clipping", open && dialog.scrollWidth <= dialog.clientWidth + 1 && panel.scrollWidth <= panel.clientWidth + 1 && Math.max(doc.documentElement.scrollWidth - doc.documentElement.clientWidth, doc.body.scrollWidth - doc.body.clientWidth) <= 1, {
      dialogScrollWidth: dialog.scrollWidth,
      dialogClientWidth: dialog.clientWidth,
      panelScrollWidth: panel.scrollWidth,
      panelClientWidth: panel.clientWidth,
      documentScrollWidth: Math.max(doc.documentElement.scrollWidth, doc.body.scrollWidth),
      documentClientWidth: Math.max(doc.documentElement.clientWidth, doc.body.clientWidth),
    });
    assertion(assertions, "settings-evidence-is-dialog-not-page-chrome", dialog.contains(panel) && menu.closest("#settings-dialog") === dialog && !dialog.contains(topbar) && !dialog.contains(composer), {
      dialogId: dialog.id,
      panelWithinDialog: dialog.contains(panel),
      menuWithinDialog: menu.closest("#settings-dialog") === dialog,
      dialogContainsTopbar: dialog.contains(topbar),
      dialogContainsComposer: dialog.contains(composer),
    });

    dialog.scrollTop = 0;
    const closeRect = close.getBoundingClientRect();
    const viewportAtTop = dialog.getBoundingClientRect();
    assertion(assertions, "settings-close-reachable-at-dialog-top", open && isVisible(close, win) && closeRect.left >= viewportAtTop.left - 1 && closeRect.right <= viewportAtTop.right + 1 && closeRect.top >= -1 && closeRect.bottom <= spec.height + 1, {
      close: rectDetails(closeRect),
      dialog: rectDetails(viewportAtTop),
    });

    const maxScrollTop = Math.max(0, dialog.scrollHeight - dialog.clientHeight);
    const hasInternalVerticalOverflow = maxScrollTop > 1;
    dialog.scrollTop = dialog.scrollHeight;
    const afterScrollTop = dialog.scrollTop;
    const scrolledDialogRect = dialog.getBoundingClientRect();
    const menuRect = menu.getBoundingClientRect();
    const saveRect = save.getBoundingClientRect();
    const bottomReachable = menuRect.bottom >= scrolledDialogRect.top - 1 && menuRect.bottom <= scrolledDialogRect.bottom + 1 && saveRect.bottom >= scrolledDialogRect.top - 1 && saveRect.bottom <= scrolledDialogRect.bottom + 1 && saveRect.top >= -1 && saveRect.bottom <= spec.height + 1;
    assertion(assertions, "settings-menu-and-save-reachable-after-internal-scroll", open && hasInternalVerticalOverflow && afterScrollTop >= maxScrollTop - 1 && isVisible(menu, win) && isVisible(save, win) && bottomReachable, {
      maxScrollTop,
      afterScrollTop,
      dialog: rectDetails(scrolledDialogRect),
      menu: rectDetails(menuRect),
      save: rectDetails(saveRect),
    });
    assertion(assertions, "settings-dialog-content-reaches-bottom", open && hasInternalVerticalOverflow && afterScrollTop >= maxScrollTop - 1 && menuRect.bottom <= scrolledDialogRect.bottom + 1, {
      scrollHeight: dialog.scrollHeight,
      clientHeight: dialog.clientHeight,
      maxScrollTop,
      afterScrollTop,
      menu: rectDetails(menuRect),
      dialog: rectDetails(scrolledDialogRect),
    });
    const pageChromeAfterScroll = [topbar, composer].map((element) => rectDetails(element.getBoundingClientRect()));
    assertion(assertions, "settings-scroll-does-not-substitute-page-chrome", JSON.stringify(pageChromeBeforeScroll) === JSON.stringify(pageChromeAfterScroll) && !dialog.contains(topbar) && !dialog.contains(composer), {
      before: pageChromeBeforeScroll,
      after: pageChromeAfterScroll,
      dialog: rectDetails(scrolledDialogRect),
    });
  }

  function runCase(spec) {
    const result = { id: spec.id, width: spec.width, height: spec.height, assertions: [], passed: false };
    try {
      if (currentIframe) currentIframe.remove();
      const iframe = document.createElement("iframe");
      currentIframe = iframe;
      iframe.setAttribute("sandbox", "allow-same-origin");
      iframe.setAttribute("aria-hidden", "true");
      iframe.style.cssText = "display:block; position:absolute; left:0; top:0; opacity:0; pointer-events:none; width:" + spec.width + "px; height:" + spec.height + "px; border:0;";
      document.body.appendChild(iframe);
      const childDoc = iframe.contentDocument;
      const childWin = iframe.contentWindow;
      if (!childDoc || !childWin) throw new Error("sandboxed same-origin iframe is not accessible");
      childDoc.open();
      childDoc.write(payload.html);
      childDoc.close();
      const style = childDoc.createElement("style");
      style.textContent = payload.css;
      style.dataset.layoutHarness = "actual-taskpane-css";
      childDoc.head.appendChild(style);
      setupCase(childDoc, spec);
      settle();
      basicGeometryAssertions(childDoc, childWin, spec, result.assertions);
      shortHeightAssertions(childDoc, childWin, spec, result.assertions);
      settingsOverlayAssertions(childDoc, childWin, spec, result.assertions);
      result.passed = result.assertions.length > 0 && result.assertions.every((item) => item.passed);
    } catch (error) {
      result.error = error instanceof Error ? error.message : String(error);
      result.passed = false;
    }
    results.push(result);
  }

  function run() {
    try {
      if (!payload || payload.version !== 1 || !Array.isArray(payload.cases)) throw new Error("invalid harness payload");
      for (const spec of payload.cases) runCase(spec);
      const expectedIds = payload.cases.map((spec) => spec.id);
      const actualIds = results.map((result) => result.id);
      const complete = results.length === expectedIds.length && expectedIds.every((id, index) => actualIds[index] === id) && results.every((result) => Array.isArray(result.assertions) && result.assertions.length > 0);
      const failedAssertions = results.flatMap((result) => result.assertions.filter((item) => !item.passed).map((item) => ({ caseId: result.id, ...item })));
      emitResult({
        version: RESULT_VERSION,
        complete,
        expectedCaseIds: expectedIds,
        cases: results,
        browserErrors,
        summary: {
          expectedCases: expectedIds.length,
          completedCases: results.length,
          passedCases: results.filter((result) => result.passed).length,
          failedCases: results.filter((result) => !result.passed).length,
          failedAssertions: failedAssertions.length,
        },
      });
    } catch (error) {
      emitResult({
        version: RESULT_VERSION,
        complete: false,
        expectedCaseIds: Array.isArray(payload && payload.cases) ? payload.cases.map((spec) => spec.id) : [],
        cases: results,
        browserErrors,
        error: error instanceof Error ? error.message : String(error),
        summary: {
          expectedCases: Array.isArray(payload && payload.cases) ? payload.cases.length : 0,
          completedCases: results.length,
          passedCases: 0,
          failedCases: results.length,
          failedAssertions: results.reduce((count, result) => count + result.assertions.filter((item) => !item.passed).length, 0),
        },
      });
    }
  }

  run();
})();
`;

function parsePrivacyLabels(source) {
  const labels = {};
  const block = source.match(/PRIVACY_LABELS[\s\S]*?\};/u)?.[0] || "";
  const pattern = /\b(local|cloud|external):\s*\{\s*label:\s*["']([^"']+)["']/gu;
  for (const match of block.matchAll(pattern)) labels[match[1]] = { label: match[2], tooltip: "" };
  if (Object.keys(labels).length !== 3 || !labels.local || !labels.cloud || !labels.external) {
    throw new Error("could not parse all privacy labels from src/privacy.ts");
  }
  return labels;
}

function buildHarness(payload) {
  const payloadBase64 = Buffer.from(JSON.stringify(payload), "utf8").toString("base64");
  return `<!doctype html><html><head><meta charset="utf-8"><meta name="hermes-taskpane-layout-payload" content="${payloadBase64}"></head><body><script>${HARNESS_SCRIPT}</script></body></html>`;
}

function parseResult(output) {
  const pattern = new RegExp(`<meta\\s+name=["']${RESULT_META_NAME}["']\\s+content=["']([A-Za-z0-9+/=]+)["']`, "u");
  const match = output.match(pattern);
  if (!match) return null;
  try {
    return JSON.parse(Buffer.from(match[1], "base64").toString("utf8"));
  } catch {
    return null;
  }
}

function validateResult(result) {
  const expectedIds = CASES.map((spec) => spec.id);
  if (!result || result.version !== RESULT_VERSION) return { ok: false, reason: "missing or invalid result marker", result };
  if (result.complete !== true) return { ok: false, reason: result.error || "incomplete result marker", result };
  if (!Array.isArray(result.expectedCaseIds) || result.expectedCaseIds.join("\n") !== expectedIds.join("\n")) return { ok: false, reason: "result case list does not match expected case list", result };
  if (!Array.isArray(result.cases) || result.cases.length !== expectedIds.length) return { ok: false, reason: "result case count is incomplete", result };
  for (let index = 0; index < expectedIds.length; index += 1) {
    const item = result.cases[index];
    if (!item || item.id !== expectedIds[index] || !Array.isArray(item.assertions) || item.assertions.length === 0) {
      return { ok: false, reason: "case or assertion result is incomplete at index " + index, result };
    }
  }
  const settingsCaseIndex = CASES.findIndex((spec) => spec.settingsOpen === true);
  if (settingsCaseIndex >= 0) {
    const settingsAssertionNames = new Set(result.cases[settingsCaseIndex].assertions.map((item) => item.name));
    const missingSettingsAssertions = REQUIRED_SETTINGS_ASSERTIONS.filter((name) => !settingsAssertionNames.has(name));
    if (missingSettingsAssertions.length > 0) return { ok: false, reason: "settings assertion contract is incomplete: " + missingSettingsAssertions.join(", "), result };
  }
  const failedAssertions = result.cases.flatMap((item) => item.assertions.filter((assertion) => !assertion.passed).map((assertion) => ({ caseId: item.id, ...assertion })));
  if (failedAssertions.length > 0) return { ok: false, reason: "layout assertions failed", result, failedAssertions };
  if (Array.isArray(result.browserErrors) && result.browserErrors.length > 0) return { ok: false, reason: "browser error reported", result, failedAssertions: result.browserErrors };
  return { ok: true, result };
}

function formatFailures(validation) {
  const result = validation.result;
  const lines = ["TASKPANE_LAYOUT_RESULT", JSON.stringify(result.summary || {}, null, 2)];
  if (validation.failedAssertions) {
    for (const failure of validation.failedAssertions) lines.push(JSON.stringify(failure));
  }
  if (result && result.error) lines.push("HARNESS_ERROR: " + result.error);
  return lines.join("\n");
}

async function assertExecutable(browserBin) {
  try {
    await fs.access(browserBin, fsConstants.X_OK);
  } catch {
    throw new Error("browser executable is not runnable: " + browserBin);
  }
}

async function runBrowser(browserBin, harnessPath, profilePath) {
  const args = [
    "--headless",
    "--disable-gpu",
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-background-networking",
    "--allow-file-access-from-files",
    "--force-device-scale-factor=1",
    "--dump-dom",
    "--virtual-time-budget=10000",
    "--user-data-dir=" + profilePath,
    "file://" + harnessPath,
  ];
  return new Promise((resolve, reject) => {
    const child = spawn(browserBin, args, { detached: true, stdio: ["ignore", "pipe", "pipe"] });
    const stdout = [];
    const stderr = [];
    let settled = false;
    let markerSeen = false;
    let terminationTimer;
    const terminate = (signal) => {
      if (!child.pid) return;
      try {
        process.kill(-child.pid, signal);
      } catch {
        child.kill(signal);
      }
    };
    const timer = setTimeout(() => {
      if (settled) return;
      terminate("SIGKILL");
      settled = true;
      reject(new Error("browser timed out after " + TIMEOUT_MS + "ms"));
    }, TIMEOUT_MS);
    child.stdout.on("data", (chunk) => {
      stdout.push(chunk);
      if (!markerSeen && RESULT_MARKER_PATTERN.test(Buffer.concat(stdout).toString("utf8"))) {
        markerSeen = true;
        terminate("SIGTERM");
        terminationTimer = setTimeout(() => terminate("SIGKILL"), 1000);
      }
    });
    child.stderr.on("data", (chunk) => stderr.push(chunk));
    child.once("error", (error) => {
      if (settled) return;
      clearTimeout(timer);
      if (terminationTimer) clearTimeout(terminationTimer);
      settled = true;
      reject(new Error("browser failed to start: " + error.message));
    });
    child.once("close", (code, signal) => {
      if (settled) return;
      clearTimeout(timer);
      if (terminationTimer) clearTimeout(terminationTimer);
      settled = true;
      resolve({
        code,
        signal,
        markerSeen,
        stdout: Buffer.concat(stdout).toString("utf8"),
        stderr: Buffer.concat(stderr).toString("utf8"),
      });
    });
  });
}

async function main() {
  let tempDir;
  try {
    const [html, css, privacySource] = await Promise.all([
      fs.readFile(path.join(REPO_ROOT, "src/taskpane/taskpane.html"), "utf8"),
      fs.readFile(path.join(REPO_ROOT, "src/taskpane/taskpane.css"), "utf8"),
      fs.readFile(path.join(REPO_ROOT, "src/privacy.ts"), "utf8"),
    ]);
    const privacyLabels = parsePrivacyLabels(privacySource);
    const sanitizedHtml = html.replace(/<script\b[^>]*>[\s\S]*?<\/script\s*>/giu, "");
    if (/<script\b/iu.test(sanitizedHtml)) throw new Error("source fixture still contains executable script markup");

    const scratchRoot = process.env.TMPDIR || os.tmpdir();
    tempDir = await fs.mkdtemp(path.join(scratchRoot, "hermes-taskpane-layout-"));
    const harnessPath = path.join(tempDir, "harness.html");
    const profilePath = path.join(tempDir, "chrome-profile");
    const payload = { version: 1, html: sanitizedHtml, css, privacyLabels, cases: CASES };
    await fs.writeFile(harnessPath, buildHarness(payload), "utf8");

    const browserBin = process.env.BROWSER_BIN || DEFAULT_BROWSER;
    await assertExecutable(browserBin);
    const browser = await runBrowser(browserBin, harnessPath, profilePath);
    const result = parseResult(browser.stdout);
    const validation = validateResult(result);
    if (browser.code !== 0 && !browser.markerSeen) {
      console.error("browser exit code: " + browser.code + (browser.signal ? " signal: " + browser.signal : ""));
      if (browser.stderr.trim()) console.error(browser.stderr.trim());
      if (result) console.error(formatFailures(validation));
      else if (browser.stdout.trim()) console.error(browser.stdout.trim());
      throw new Error("browser process failed");
    }
    if (!validation.ok) {
      console.error(formatFailures(validation));
      throw new Error(validation.reason);
    }
    console.log("TASKPANE_LAYOUT_PASS " + JSON.stringify(result.summary));
  } finally {
    if (tempDir) {
      if (process.env.KEEP_LAYOUT_TEMP === "1") console.error("KEEP_LAYOUT_TEMP=" + tempDir);
      else await fs.rm(tempDir, { recursive: true, force: true });
    }
  }
}

main().catch((error) => {
  console.error("test-taskpane-layout: " + (error instanceof Error ? error.message : String(error)));
  process.exitCode = 1;
});
