import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import path from "node:path";

const css = ["omnia-pane.css", "taskpane.css"].map(name => readFileSync(path.resolve(process.cwd(), "src/taskpane", name), "utf8")).join("\n");

const parentBaselineCss = `html,
body {
  height: 100%;
  margin: 0;
}

body {
  min-width: 320px;
}`;

const rootRuleCoverageCss = `${parentBaselineCss}

html {
  min-width: 280px;
}

body {
  min-width: 290px;
}

@media (max-width: 340px) {
  html,
  body {
    min-width: 999px;
  }
}`;

function rule(selector: string): string {
  const escapedSelector = selector.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = css.match(new RegExp(`(?:^|})\\s*${escapedSelector}\\s*\\{([^{}]*)\\}`, "m"));
  expect(match, `Missing CSS rule for ${selector}`).not.toBeNull();
  return match?.[1] ?? "";
}

function mediaBlock(query: RegExp): string | null {
  const match = css.match(query);
  if (!match || match.index === undefined) return null;

  const openBrace = css.indexOf("{", match.index);
  let depth = 0;
  for (let index = openBrace; index < css.length; index += 1) {
    if (css[index] === "{") depth += 1;
    if (css[index] === "}") {
      depth -= 1;
      if (depth === 0) return css.slice(openBrace + 1, index);
    }
  }

  return null;
}

function findBlockStart(cssText: string, start: number): number {
  let quote: string | null = null;
  for (let index = start; index < cssText.length; index += 1) {
    const character = cssText[index];
    if (quote) {
      if (character === "\\") {
        index += 1;
      } else if (character === quote) {
        quote = null;
      }
      continue;
    }

    if (character === "/" && cssText[index + 1] === "*") {
      const commentEnd = cssText.indexOf("*/", index + 2);
      if (commentEnd === -1) return -1;
      index = commentEnd + 1;
      continue;
    }

    if (character === "'" || character === '"') {
      quote = character;
    } else if (character === "{") {
      return index;
    }
  }
  return -1;
}

function findBlockEnd(cssText: string, openBrace: number): number {
  let depth = 1;
  let quote: string | null = null;
  for (let index = openBrace + 1; index < cssText.length; index += 1) {
    const character = cssText[index];
    if (quote) {
      if (character === "\\") {
        index += 1;
      } else if (character === quote) {
        quote = null;
      }
      continue;
    }

    if (character === "/" && cssText[index + 1] === "*") {
      const commentEnd = cssText.indexOf("*/", index + 2);
      if (commentEnd === -1) return -1;
      index = commentEnd + 1;
      continue;
    }

    if (character === "'" || character === '"') {
      quote = character;
    } else if (character === "{") {
      depth += 1;
    } else if (character === "}") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}

function rootBodyDeclarations(cssText: string): string[] {
  const declarations: string[] = [];
  let searchStart = 0;
  while (searchStart < cssText.length) {
    const openBrace = findBlockStart(cssText, searchStart);
    if (openBrace === -1) break;

    const closeBrace = findBlockEnd(cssText, openBrace);
    if (closeBrace === -1) break;

    const prelude = cssText.slice(searchStart, openBrace).trim();
    if (!prelude.startsWith("@")) {
      const selectors = prelude.split(",").map((selector) => selector.trim());
      if (selectors.some((selector) => selector === "html" || selector === "body")) {
        declarations.push(cssText.slice(openBrace + 1, closeBrace));
      }
    }

    searchStart = closeBrace + 1;
  }
  return declarations;
}

function oversizedRootMinWidths(cssText: string): number[] {
  return rootBodyDeclarations(cssText)
    .flatMap((declarations) =>
      [...declarations.matchAll(/min-width\s*:\s*(\d+(?:\.\d+)?)px/gi)].map((match) => Number(match[1])),
    )
    .filter((width) => width > 260);
}

describe("Word taskpane narrow layout contract", () => {
  it("detects the parent standalone body minimum width", () => {
    expect(oversizedRootMinWidths(parentBaselineCss)).toEqual([320]);
    expect(oversizedRootMinWidths(rootRuleCoverageCss)).toEqual([320, 280, 290]);
  });

  it("does not impose an oversized minimum width on the document roots", () => {
    expect(oversizedRootMinWidths(css)).toEqual([]);
  });

  it("makes the shared flex conversation shrink instead of forcing overflow", () => {
    expect(rule(".app-shell")).toMatch(/display\s*:\s*flex\s*;/);
    expect(rule(".app-shell")).toMatch(/flex-direction\s*:\s*column\s*;/);
    expect(rule(".messages")).toMatch(/min-height\s*:\s*0\s*;/);
    expect(rule(".messages")).toMatch(/flex\s*:\s*1\s+1\s+auto\s*;/);
  });

  it("keeps a 100vh fallback before the dynamic viewport height", () => {
    expect(rule(".app-shell")).toMatch(
      /height\s*:\s*100vh\s*;[\s\S]*height\s*:\s*100dvh\s*;/,
    );
  });

  it("lets the primary grid and flex children shrink", () => {
    expect(rule(".topbar > *")).toMatch(/min-width\s*:\s*0\s*;/);
    expect(rule(".messages")).toMatch(/min-width\s*:\s*0\s*;/);
    expect(rule(".composer")).toMatch(/min-width\s*:\s*0\s*;/);
    expect(rule(".composer-actions")).toMatch(/min-width\s*:\s*0\s*;/);
  });

  it("makes the bridge row wrap and ellipsize around the copy control", () => {
    const strip = rule(".bridge-strip");
    const copy = rule(".bridge-copy");
    const status = rule(".bridge-status");
    const session = rule(".bridge-session-id");
    const copyButton = rule(".bridge-copy-button");

    expect({
      stripWraps: /flex-wrap\s*:\s*wrap\s*;/.test(strip),
      copyCanShrink: /min-width\s*:\s*0\s*;/.test(copy),
      statusEllipsizes:
        /overflow\s*:\s*hidden\s*;/.test(status) &&
        /text-overflow\s*:\s*ellipsis\s*;/.test(status) &&
        /white-space\s*:\s*nowrap\s*;/.test(status),
      sessionFits:
        /max-width\s*:\s*100%\s*;/.test(session) &&
        /overflow\s*:\s*hidden\s*;/.test(session) &&
        /text-overflow\s*:\s*ellipsis\s*;/.test(session) &&
        /white-space\s*:\s*nowrap\s*;/.test(session),
      copyButtonStaysVisible: /flex\s*:\s*none\s*;/.test(copyButton),
    }).toEqual({
      stripWraps: true,
      copyCanShrink: true,
      statusEllipsizes: true,
      sessionFits: true,
      copyButtonStaysVisible: true,
    });
  });

  it("tightens spacing and composer height for panes at or below 340px", () => {
    const narrowMedia = mediaBlock(/@media\s*\(\s*max-width\s*:\s*340px\s*\)/);
    expect(narrowMedia).not.toBeNull();
    if (!narrowMedia) return;

    expect(narrowMedia).toMatch(/\.topbar\s*\{[^}]*\b(?:gap|padding)\s*:/s);
    expect(narrowMedia).toMatch(/\.bridge-strip\s*\{[^}]*\b(?:gap|padding)\s*:/s);
    expect(rule("textarea")).toMatch(/min-height\s*:\s*82px\s*;/);
    expect(rule("textarea")).toMatch(/resize\s*:\s*none\s*;/);
  });
});
