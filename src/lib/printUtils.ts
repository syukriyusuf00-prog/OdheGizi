/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import html2canvas from "html2canvas";
import jsPDF from "jspdf";

// Cache for converted colors
const colorCache = new Map<string, string>();

let dummyEl: HTMLDivElement | null = null;

function getDummyEl(): HTMLDivElement | null {
  if (typeof document === "undefined") return null;
  if (!dummyEl && document.body) {
    dummyEl = document.createElement("div");
    dummyEl.style.display = "none";
    document.body.appendChild(dummyEl);
  }
  return dummyEl;
}

/**
 * Converts an oklab(...) color string into an rgb(...) or rgba(...) color string.
 */
function oklabToRgb(oklabStr: string): string {
  if (colorCache.has(oklabStr)) {
    return colorCache.get(oklabStr)!;
  }

  const dummy = getDummyEl();
  if (dummy) {
    try {
      dummy.style.color = "";
      dummy.style.color = oklabStr;
      const computed = getComputedStyle(dummy).color;
      if (computed && !computed.includes("oklab") && !computed.includes("oklch")) {
        colorCache.set(oklabStr, computed);
        return computed;
      }
    } catch {
      // Fallback
    }
  }

  const match = oklabStr.match(/oklab\(\s*([\d.%]+)[\s,]+([\d.-]+)[\s,]+([\d.-]+)(?:[\s,/]+([\d.%]+))?\s*\)/i);
  if (!match) return "rgb(0, 0, 0)";

  let l = parseFloat(match[1]);
  if (match[1].endsWith("%")) l /= 100;
  const a = parseFloat(match[2]);
  const b = parseFloat(match[3]);
  let alpha = match[4] !== undefined ? parseFloat(match[4]) : 1;
  if (match[4] && match[4].endsWith("%")) alpha /= 100;

  const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
  const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
  const s_ = l - 0.0894841775 * a - 1.291485548 * b;

  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;

  const rLin = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLin = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLin = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.707614701 * s3;

  const toGamma = (x: number) => {
    if (x <= 0) return 0;
    if (x >= 1) return 255;
    const g = x <= 0.0031308 ? 12.92 * x : 1.055 * Math.pow(x, 1 / 2.4) - 0.055;
    return Math.min(255, Math.max(0, Math.round(g * 255)));
  };

  const r = toGamma(rLin);
  const g = toGamma(gLin);
  const bComp = toGamma(bLin);

  const res = alpha < 1 ? `rgba(${r}, ${g}, ${bComp}, ${alpha})` : `rgb(${r}, ${g}, ${bComp})`;
  colorCache.set(oklabStr, res);
  return res;
}

/**
 * Converts an oklch(...) color string into an rgb(...) or rgba(...) color string.
 */
function oklchToRgb(oklchStr: string): string {
  if (colorCache.has(oklchStr)) {
    return colorCache.get(oklchStr)!;
  }

  const dummy = getDummyEl();
  if (dummy) {
    try {
      dummy.style.color = "";
      dummy.style.color = oklchStr;
      const computed = getComputedStyle(dummy).color;
      if (computed && !computed.includes("oklch") && !computed.includes("oklab")) {
        colorCache.set(oklchStr, computed);
        return computed;
      }
    } catch {
      // Fallback
    }
  }

  const match = oklchStr.match(/oklch\(\s*([\d.%]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.%]+))?\s*\)/i);
  if (!match) return "rgb(0, 0, 0)";

  let l = parseFloat(match[1]);
  if (match[1].endsWith("%")) l /= 100;
  const c = parseFloat(match[2]);
  const h = parseFloat(match[3]);
  let alpha = match[4] !== undefined ? parseFloat(match[4]) : 1;
  if (match[4] && match[4].endsWith("%")) alpha /= 100;

  const hRad = (h * Math.PI) / 180;
  const a = c * Math.cos(hRad);
  const b = c * Math.sin(hRad);

  const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
  const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
  const s_ = l - 0.0894841775 * a - 1.291485548 * b;

  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;

  const rLin = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLin = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLin = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.707614701 * s3;

  const toGamma = (x: number) => {
    if (x <= 0) return 0;
    if (x >= 1) return 255;
    const g = x <= 0.0031308 ? 12.92 * x : 1.055 * Math.pow(x, 1 / 2.4) - 0.055;
    return Math.min(255, Math.max(0, Math.round(g * 255)));
  };

  const r = toGamma(rLin);
  const g = toGamma(gLin);
  const bComp = toGamma(bLin);

  const res = alpha < 1 ? `rgba(${r}, ${g}, ${bComp}, ${alpha})` : `rgb(${r}, ${g}, ${bComp})`;
  colorCache.set(oklchStr, res);
  return res;
}

/**
 * Converts a CSS color(...) function string (e.g., color(srgb 0.1 0.2 0.3 / 0.8)) into an rgb(...) or rgba(...) string.
 */
function parseColorFunctionToRgb(colorStr: string): string {
  if (colorCache.has(colorStr)) {
    return colorCache.get(colorStr)!;
  }
  const match = colorStr.match(/color\(\s*([\w-]+)\s+([\d.%-]+)[\s,]+([\d.%-]+)[\s,]+([\d.%-]+)(?:[\s,/]+([\d.%-]+))?\s*\)/i);
  if (!match) return "rgb(0, 0, 0)";

  let r = parseFloat(match[2]);
  if (match[2].endsWith("%")) r = (r / 100) * 255;
  else if (r <= 1) r = r * 255;

  let g = parseFloat(match[3]);
  if (match[3].endsWith("%")) g = (g / 100) * 255;
  else if (g <= 1) g = g * 255;

  let b = parseFloat(match[4]);
  if (match[4].endsWith("%")) b = (b / 100) * 255;
  else if (b <= 1) b = b * 255;

  let alpha = match[5] !== undefined ? parseFloat(match[5]) : 1;
  if (match[5] && match[5].endsWith("%")) alpha /= 100;

  r = Math.min(255, Math.max(0, Math.round(r)));
  g = Math.min(255, Math.max(0, Math.round(g)));
  b = Math.min(255, Math.max(0, Math.round(b)));

  const res = alpha < 1 ? `rgba(${r}, ${g}, ${b}, ${alpha})` : `rgb(${r}, ${g}, ${b})`;
  colorCache.set(colorStr, res);
  return res;
}

let canvas2dCtx: CanvasRenderingContext2D | null = null;

function convertCssColorToRgb(colorStr: string): string {
  if (colorCache.has(colorStr)) {
    return colorCache.get(colorStr)!;
  }
  try {
    if (!canvas2dCtx && typeof document !== "undefined") {
      const canvas = document.createElement("canvas");
      canvas.width = 1;
      canvas.height = 1;
      canvas2dCtx = canvas.getContext("2d");
    }
    if (canvas2dCtx) {
      canvas2dCtx.fillStyle = "#000000";
      canvas2dCtx.fillStyle = colorStr;
      const converted = canvas2dCtx.fillStyle;
      if (
        converted &&
        !converted.includes("oklab") &&
        !converted.includes("oklch") &&
        !converted.includes("color(") &&
        !converted.includes("color-mix")
      ) {
        colorCache.set(colorStr, converted);
        return converted;
      }
    }
  } catch {
    // Fallback
  }

  if (colorStr.includes("color(")) return parseColorFunctionToRgb(colorStr);
  if (colorStr.includes("oklab")) return oklabToRgb(colorStr);
  if (colorStr.includes("oklch")) return oklchToRgb(colorStr);

  return colorStr;
}

const HAS_UNSUPPORTED_COLOR = (s: string) =>
  s.includes("oklab") || s.includes("oklch") || s.includes("light-dark") || s.includes("color-mix") || s.includes("color(");

export function sanitizeOklchString(str: string): string {
  if (!str) return str;
  if (!HAS_UNSUPPORTED_COLOR(str)) {
    return str;
  }

  let result = str;

  if (result.includes("light-dark")) {
    result = result.replace(/light-dark\(\s*([^,\)]+)\s*,\s*([^,\)]+)\s*\)/gi, "$1");
  }
  if (result.includes("color(")) {
    result = result.replace(/color\([^()]+\)/gi, (m) => convertCssColorToRgb(m));
  }
  if (result.includes("oklab")) {
    result = result.replace(/oklab\([^()]+\)/gi, (m) => convertCssColorToRgb(m));
  }
  if (result.includes("oklch")) {
    result = result.replace(/oklch\([^()]+\)/gi, (m) => convertCssColorToRgb(m));
  }
  if (result.includes("color-mix")) {
    result = result.replace(/color-mix\([^()]+\)/gi, (m) => convertCssColorToRgb(m));
  }

  return result;
}

let cachedSanitizedDocCss: { key: string; css: string } | null = null;

/**
 * Collects and sanitizes current document stylesheets once per stylesheet signature
 * so opening print previews or triggering smart print never stalls the main thread.
 */
export function getCachedSanitizedDocumentStyles(): string {
  if (typeof document === "undefined") return "";
  const nodes = Array.from(document.querySelectorAll("style, link[rel='stylesheet']"));
  const sig = nodes
    .map((el) =>
      el.tagName.toLowerCase() === "style"
        ? `s:${(el.textContent || "").length}`
        : `l:${(el as HTMLLinkElement).href || ""}`
    )
    .join("|");

  if (cachedSanitizedDocCss && cachedSanitizedDocCss.key === sig) {
    return cachedSanitizedDocCss.css;
  }

  let combinedStyles = "";
  nodes.forEach((styleEl) => {
    if (styleEl.tagName.toLowerCase() === "style") {
      combinedStyles += (styleEl.textContent || "") + "\n";
    } else if (styleEl.tagName.toLowerCase() === "link") {
      try {
        const sheet = (styleEl as HTMLLinkElement).sheet;
        if (sheet) {
          Array.from(sheet.cssRules).forEach((rule) => {
            combinedStyles += rule.cssText + "\n";
          });
        }
      } catch {
        // Ignore CORS stylesheets
      }
    }
  });

  const sanitized = sanitizeOklchString(combinedStyles);
  cachedSanitizedDocCss = { key: sig, css: sanitized };
  return sanitized;
}

function createSanitizedComputedStyle(origFn: typeof window.getComputedStyle) {
  return function (elt: Element, pseudoElt?: string | null): CSSStyleDeclaration {
    const style = origFn(elt, pseudoElt);
    return new Proxy(style, {
      get(target, prop) {
        if (prop === "getPropertyValue") {
          return function (propertyName: string) {
            const raw = target.getPropertyValue(propertyName);
            if (raw && typeof raw === "string" && HAS_UNSUPPORTED_COLOR(raw)) {
              return sanitizeOklchString(raw);
            }
            return raw;
          };
        }

        let val: any;
        try {
          val = target[prop as keyof CSSStyleDeclaration];
        } catch {
          val = Reflect.get(target, prop);
        }

        if (typeof val === "string" && HAS_UNSUPPORTED_COLOR(val)) {
          return sanitizeOklchString(val);
        }
        if (typeof val === "function") {
          return val.bind(target);
        }
        return val;
      },
    });
  };
}

export interface OptimalPrintLayout {
  orientation: "portrait" | "landscape";
  adaptiveFontPt: number;
  maxCols: number;
  contentWidthPx: number;
}

export interface SmartPrintOptions {
  title?: string;
  paperSize?: "A4" | "F4";
  orientation?: "auto" | "portrait" | "landscape";
  scalePercent?: number;
}

export interface TidyExtractResult {
  paperSize: "A4" | "F4";
  orientation: "portrait" | "landscape";
  adaptiveFontPt: number;
  maxCols: number;
  tableCount: number;
  fittedCleanly: boolean;
  recommendedScalePercent: number;
  summaryLabel: string;
}

const MM_TO_PX = 96 / 25.4; // ~3.779527559 px/mm at standard CSS 96 DPI

export function getPaperDimensionsMm(
  paperSize: "A4" | "F4" = "A4",
  orientation: "portrait" | "landscape" = "portrait"
): { widthMm: number; heightMm: number } {
  let widthMm = paperSize === "F4" ? 215 : 210;
  let heightMm = paperSize === "F4" ? 330 : 297;
  if (orientation === "landscape") {
    const tmp = widthMm;
    widthMm = heightMm;
    heightMm = tmp;
  }
  return { widthMm, heightMm };
}

export function getPaperPrintablePx(
  paperSize: "A4" | "F4" = "A4",
  orientation: "portrait" | "landscape" = "portrait",
  marginMm: number = 8
): { widthPx: number; heightPx: number } {
  const { widthMm, heightMm } = getPaperDimensionsMm(paperSize, orientation);
  return {
    widthPx: Math.round((widthMm - marginMm * 2) * MM_TO_PX),
    heightPx: Math.round((heightMm - marginMm * 2) * MM_TO_PX),
  };
}

function isTransparentColor(color?: string | null): boolean {
  if (!color) return true;
  const c = color.trim().toLowerCase();
  if (!c || c === "transparent" || c === "none" || c === "initial" || c === "inherit") {
    return true;
  }
  const rgbaMatch = c.match(/rgba?\(\s*[\d.]+\s*,\s*[\d.]+\s*,\s*[\d.]+\s*(?:,\s*([\d.]+)\s*)?\)/);
  if (rgbaMatch && rgbaMatch[1] !== undefined && parseFloat(rgbaMatch[1]) === 0) {
    return true;
  }
  return false;
}

/**
 * Inspects a worksheet DOM element and determines whether it fits cleanly in Portrait
 * or should automatically switch to Landscape with adaptive Arial typography.
 */
export function detectOptimalPrintLayout(
  element: HTMLElement | null,
  paperSize: "A4" | "F4" = "A4"
): OptimalPrintLayout {
  if (!element) {
    return {
      orientation: "portrait",
      adaptiveFontPt: 9.5,
      maxCols: 0,
      contentWidthPx: 0,
    };
  }

  // Printable width in CSS px (~96dpi) with ~16mm total horizontal margin
  const portraitPrintablePx = getPaperPrintablePx(paperSize, "portrait", 8).widthPx;

  let maxCols = 0;
  let widestTablePx = 0;

  element.querySelectorAll("table").forEach((table) => {
    if (table.closest(".no-print")) return;
    const htmlTable = table as HTMLElement;
    const ownerWin = htmlTable.ownerDocument?.defaultView || window;
    const style = ownerWin.getComputedStyle(htmlTable);
    if (style.display === "none") return;

    const tblWidth = Math.max(htmlTable.scrollWidth || 0, htmlTable.offsetWidth || 0);
    if (tblWidth > widestTablePx) widestTablePx = tblWidth;

    table.querySelectorAll("tr").forEach((tr) => {
      if (tr.closest(".no-print")) return;
      let colCount = 0;
      Array.from(tr.children).forEach((cell) => {
        if ((cell as HTMLElement).classList?.contains("no-print")) return;
        const cs = parseInt((cell as HTMLTableCellElement).getAttribute("colspan") || "1", 10);
        colCount += Number.isFinite(cs) && cs > 0 ? cs : 1;
      });
      if (colCount > maxCols) maxCols = colCount;
    });
  });

  const rectWidth = element.getBoundingClientRect?.().width || 0;
  const contentWidthPx = Math.max(widestTablePx, element.scrollWidth || 0, rectWidth);

  // Switch to landscape when the table has many columns (>= 8) or naturally overflows Portrait width
  const needsLandscape =
    maxCols >= 8 ||
    widestTablePx > portraitPrintablePx * 1.04 ||
    (maxCols >= 7 && widestTablePx > portraitPrintablePx * 0.96);

  const orientation: "portrait" | "landscape" = needsLandscape ? "landscape" : "portrait";

  let adaptiveFontPt = 9.5;
  if (maxCols >= 14) {
    adaptiveFontPt = 7.4;
  } else if (maxCols >= 11) {
    adaptiveFontPt = 8.0;
  } else if (maxCols >= 9) {
    adaptiveFontPt = 8.6;
  } else if (maxCols >= 7) {
    adaptiveFontPt = 9.0;
  }

  return {
    orientation,
    adaptiveFontPt,
    maxCols,
    contentWidthPx,
  };
}

/**
 * Deep-clones a DOM subtree while accurately copying live user-typed values from
 * <input>, <textarea>, and <select> (which native cloneNode(true) otherwise drops).
 */
export function cloneElementWithLiveFormValues(
  sourceEl: HTMLElement,
  targetDoc: Document = document
): HTMLElement {
  const rawClone = sourceEl.cloneNode(true) as HTMLElement;

  const srcInputs = sourceEl.querySelectorAll("input");
  const dstInputs = rawClone.querySelectorAll("input");
  srcInputs.forEach((src, i) => {
    const dst = dstInputs[i];
    if (!dst) return;
    dst.value = src.value;
    dst.defaultValue = src.value;
    dst.checked = src.checked;
    dst.defaultChecked = src.checked;
  });

  const srcTextareas = sourceEl.querySelectorAll("textarea");
  const dstTextareas = rawClone.querySelectorAll("textarea");
  srcTextareas.forEach((src, i) => {
    const dst = dstTextareas[i];
    if (!dst) return;
    dst.value = src.value;
    dst.textContent = src.value;
  });

  const srcSelects = sourceEl.querySelectorAll("select");
  const dstSelects = rawClone.querySelectorAll("select");
  srcSelects.forEach((src, i) => {
    const dst = dstSelects[i];
    if (!dst) return;
    dst.selectedIndex = src.selectedIndex;
    dst.value = src.value;
  });

  if (rawClone.ownerDocument !== targetDoc) {
    return targetDoc.importNode(rawClone, true) as HTMLElement;
  }
  return rawClone;
}

/**
 * Extracts semantic text alignment & font-weight tokens from a form control's class list
 * without carrying over form padding/border/focus classes that shift cell alignment in PDFs.
 */
function pickSemanticTextClasses(className: string): string {
  if (!className) return "";
  const keep = className.split(/\s+/).filter((cls) => {
    return (
      cls === "text-center" ||
      cls === "text-right" ||
      cls === "text-left" ||
      cls === "font-bold" ||
      cls === "font-semibold" ||
      cls === "font-medium" ||
      cls === "italic" ||
      cls === "uppercase" ||
      cls === "tabular-nums"
    );
  });
  return keep.join(" ");
}

/**
 * Replaces form inputs, textareas, and selects inside a cloned DOM tree with clean static Arial
 * text blocks so printed sheets and exported PDFs render cleanly with zero clipping or baseline drift.
 */
export function convertFormControlsToStaticText(root: HTMLElement) {
  const doc = root.ownerDocument || document;

  // Remove datalists and buttons inside print body
  root.querySelectorAll("datalist").forEach((n) => n.remove());

  root.querySelectorAll("input, textarea, select").forEach((ctrl) => {
    if (ctrl.closest(".no-print")) {
      ctrl.remove();
      return;
    }

    const insideTableCell = Boolean(ctrl.closest("td, th"));
    const tag = ctrl.tagName.toLowerCase();

    if (tag === "input") {
      const inp = ctrl as HTMLInputElement;
      if (inp.type === "hidden" || inp.type === "file") {
        inp.remove();
        return;
      }
      if (inp.type === "checkbox" || inp.type === "radio") {
        const mark = doc.createElement("span");
        mark.textContent = inp.checked ? "✓" : "☐";
        mark.style.fontFamily = "Arial, Helvetica, sans-serif";
        mark.style.fontWeight = "700";
        mark.style.display = "inline";
        inp.replaceWith(mark);
        return;
      }

      const el = doc.createElement(insideTableCell ? "div" : "span");
      const semanticCls = pickSemanticTextClasses(inp.className);
      if (semanticCls) el.className = semanticCls;
      const val = (inp.value || "").trim();
      el.textContent = val || (insideTableCell ? "\u00A0" : "");
      el.style.display = insideTableCell ? "block" : "inline";
      if (insideTableCell) {
        el.style.width = "100%";
      }
      el.style.margin = "0";
      el.style.padding = "0";
      el.style.border = "none";
      el.style.background = "transparent";
      el.style.boxShadow = "none";
      el.style.fontFamily = "Arial, Helvetica, sans-serif";
      el.style.lineHeight = "1.35";
      el.style.whiteSpace = "pre-wrap";
      el.style.wordBreak = "break-word";
      el.style.overflowWrap = "break-word";
      inp.replaceWith(el);
      return;
    }

    if (tag === "textarea") {
      const ta = ctrl as HTMLTextAreaElement;
      const div = doc.createElement("div");
      const semanticCls = pickSemanticTextClasses(ta.className);
      if (semanticCls) div.className = semanticCls;
      const val = (ta.value || ta.textContent || "").replace(/\r\n/g, "\n").trimEnd();
      div.textContent = val || "\u00A0";
      div.style.display = "block";
      div.style.width = "100%";
      div.style.margin = "0";
      div.style.padding = "0";
      div.style.border = "none";
      div.style.background = "transparent";
      div.style.boxShadow = "none";
      div.style.fontFamily = "Arial, Helvetica, sans-serif";
      div.style.lineHeight = "1.35";
      div.style.whiteSpace = "pre-wrap";
      div.style.wordBreak = "break-word";
      div.style.overflowWrap = "break-word";
      ta.replaceWith(div);
      return;
    }

    if (tag === "select") {
      const sel = ctrl as HTMLSelectElement;
      const selectedOpt = sel.options[sel.selectedIndex];
      const textVal = (selectedOpt ? selectedOpt.text : sel.value || "").trim();
      const el = doc.createElement(insideTableCell ? "div" : "span");
      const semanticCls = pickSemanticTextClasses(sel.className);
      if (semanticCls) el.className = semanticCls;
      el.textContent = textVal || (insideTableCell ? "\u00A0" : "");
      el.style.display = insideTableCell ? "block" : "inline";
      if (insideTableCell) {
        el.style.width = "100%";
      }
      el.style.margin = "0";
      el.style.padding = "0";
      el.style.border = "none";
      el.style.background = "transparent";
      el.style.boxShadow = "none";
      el.style.fontFamily = "Arial, Helvetica, sans-serif";
      el.style.lineHeight = "1.35";
      sel.replaceWith(el);
    }
  });
}

/**
 * Fixes the classic html2canvas bug where <tr>/<thead> background colors paint OVER rowSpan cells
 * from previous rows (e.g., slicing "Kode Identitas Makanan" in half!).
 * Moves section/row backgrounds safely onto individual <th>/<td> cells and clears <tr>/<thead> bg.
 */
function propagateRowBackgroundsToCells(root: HTMLElement) {
  const ownerWin = root.ownerDocument?.defaultView || window;

  root.querySelectorAll("table").forEach((table) => {
    const sectionBgMap = new Map<Element, { bg: string; color: string }>();

    table.querySelectorAll("thead, tbody, tfoot").forEach((section) => {
      const secEl = section as HTMLElement;
      const cs = ownerWin.getComputedStyle(secEl);
      const rawBg = sanitizeOklchString(cs.backgroundColor || "");
      const rawColor = sanitizeOklchString(cs.color || "");
      if (!isTransparentColor(rawBg)) {
        sectionBgMap.set(section, { bg: rawBg, color: rawColor });
        secEl.style.setProperty("background-color", "transparent", "important");
      }
    });

    table.querySelectorAll("tr").forEach((tr) => {
      const trEl = tr as HTMLElement;
      const trCs = ownerWin.getComputedStyle(trEl);
      const trBgRaw = sanitizeOklchString(trCs.backgroundColor || "");
      const trColorRaw = sanitizeOklchString(trCs.color || "");
      const parentSec = tr.parentElement ? sectionBgMap.get(tr.parentElement) : undefined;

      const effectiveBg = !isTransparentColor(trBgRaw) ? trBgRaw : parentSec?.bg || "";
      const effectiveColor = trColorRaw || parentSec?.color || "";

      if (effectiveBg) {
        Array.from(tr.children).forEach((child) => {
          const tag = child.tagName.toLowerCase();
          if (tag !== "th" && tag !== "td") return;
          const cellEl = child as HTMLElement;
          const cellCs = ownerWin.getComputedStyle(cellEl);
          const cellBgRaw = sanitizeOklchString(cellCs.backgroundColor || "");
          if (isTransparentColor(cellBgRaw)) {
            cellEl.style.setProperty("background-color", effectiveBg, "important");
          } else if (HAS_UNSUPPORTED_COLOR(cellCs.backgroundColor || "")) {
            cellEl.style.setProperty("background-color", cellBgRaw, "important");
          }
          if (effectiveColor && !cellEl.style.color) {
            const cellColorRaw = sanitizeOklchString(cellCs.color || "");
            if (cellColorRaw) {
              cellEl.style.setProperty("color", cellColorRaw, "important");
            }
          }
        });
      }

      // Always ensure <tr> never paints a background rectangle across rowSpan cells
      trEl.style.setProperty("background-color", "transparent", "important");
    });
  });
}

/**
 * Auto-Tidy Document Normalizer:
 * Cleans up wrappers, enforces uniform Arial typography, normalizes cell padding & vertical alignment,
 * kills rowSpan bg overlap, and ensures tables auto-fit cleanly to the exact paper width with zero right-edge clipping.
 */
export function tidyExtractedDocumentDom(
  root: HTMLElement,
  options: {
    targetWidthPx: number;
    adaptiveFontPt: number;
    maxCols: number;
  }
): { finalFontPt: number; fittedCleanly: boolean; tableCount: number; naturalWidthPx: number } {
  const { targetWidthPx, maxCols } = options;

  // 1. Strip interactive/non-print elements & editing attributes
  root.querySelectorAll(".no-print, button:not(.print-keep), datalist").forEach((n) => n.remove());
  root.removeAttribute("contenteditable");
  root.querySelectorAll("[contenteditable]").forEach((el) => el.removeAttribute("contenteditable"));

  // 2. Normalize root wrapper to fill target paper width cleanly with zero chrome
  root.style.margin = "0 auto";
  root.style.boxShadow = "none";
  root.style.border = "none";
  root.style.borderRadius = "0";
  root.style.transform = "none";
  root.style.zoom = "1";
  root.style.width = "100%";
  root.style.maxWidth = "100%";
  root.style.minWidth = "0";
  root.style.backgroundColor = "#ffffff";
  root.style.fontFamily = "Arial, Helvetica, sans-serif";

  // Remove screen-only padding on outer print-area container so paper margins stay consistent
  if (root.id && root.id.startsWith("print-area-")) {
    root.style.padding = "0";
  }

  // 3. Relax scroll & min-width wrappers so nothing clips horizontally
  root.querySelectorAll("*").forEach((el) => {
    if (!(el instanceof HTMLElement)) return;

    const styleAttr = el.getAttribute("style");
    if (styleAttr && HAS_UNSUPPORTED_COLOR(styleAttr)) {
      el.setAttribute("style", sanitizeOklchString(styleAttr));
    }
    if (el.style.zoom) {
      el.style.zoom = "1";
    }

    const cls = el.className && typeof el.className === "string" ? el.className : "";
    if (
      cls.includes("overflow-x-auto") ||
      cls.includes("overflow-auto") ||
      cls.includes("overflow-x-scroll")
    ) {
      el.style.overflow = "visible";
      el.style.width = "100%";
      el.style.maxWidth = "100%";
    }

    // Remove constraining screen min-w-[...] on container divs (not tables/images)
    if (el.tagName.toLowerCase() === "div" && cls.includes("min-w-[")) {
      el.style.minWidth = "0";
      el.style.maxWidth = "100%";
      el.style.width = "100%";
    }
  });

  // 4. Convert all form controls to static Arial text cleanly
  convertFormControlsToStaticText(root);

  // 5. Propagate <tr>/<thead> backgrounds down onto <th>/<td> cells (kills rowSpan overlap bug!)
  propagateRowBackgroundsToCells(root);

  // 6. Normalize every table for crisp, well-proportioned Arial output
  const tables = Array.from(root.querySelectorAll("table"));
  let currentFontPt = options.adaptiveFontPt;

  const applyTableTypography = (fontPt: number, denseMode: boolean) => {
    const headerPad = denseMode ? "4px 4px" : "5px 6px";
    const cellPad = denseMode ? "3.5px 4px" : "4.5px 6px";

    tables.forEach((tbl) => {
      const htmlTbl = tbl as HTMLElement;
      // Preserve narrow standalone signature boxes (e.g., w-64 verification box)
      const isNarrowSigBox =
        htmlTbl.classList.contains("w-64") ||
        htmlTbl.classList.contains("w-56") ||
        htmlTbl.classList.contains("w-72");

      if (!isNarrowSigBox) {
        htmlTbl.style.width = "100%";
        htmlTbl.style.maxWidth = "100%";
      }
      htmlTbl.style.borderCollapse = "collapse";
      htmlTbl.style.fontFamily = "Arial, Helvetica, sans-serif";

      tbl.querySelectorAll("th, td").forEach((cell) => {
        const cellEl = cell as HTMLElement;
        const isHeader = cell.tagName.toLowerCase() === "th";
        const parentTr = cell.parentElement;
        const keepTopAlign =
          cellEl.classList.contains("align-top") ||
          Boolean(parentTr && parentTr.classList.contains("align-top"));

        cellEl.style.fontFamily = "Arial, Helvetica, sans-serif";
        cellEl.style.fontSize = `${fontPt}pt`;
        cellEl.style.lineHeight = "1.35";
        cellEl.style.verticalAlign = keepTopAlign ? "top" : "middle";
        cellEl.style.wordBreak = "break-word";
        cellEl.style.overflowWrap = "break-word";

        // Don't crush explicit signature spacer cells (h-16 / h-20)
        const isSignatureGap =
          cellEl.classList.contains("h-16") ||
          cellEl.classList.contains("h-20") ||
          cellEl.classList.contains("h-14");
        if (!isSignatureGap) {
          cellEl.style.padding = isHeader ? headerPad : cellPad;
        }
      });
    });
  };

  applyTableTypography(currentFontPt, maxCols >= 10);

  // 7. Auto-fit width loop: if any table still overflows targetWidthPx, relax min-widths & step down font size
  let naturalWidthPx = Math.max(root.scrollWidth || 0, targetWidthPx);
  if (naturalWidthPx > targetWidthPx + 4) {
    // Pass 1: relax cell min-widths so browser distributes columns across 100% paper width
    tables.forEach((tbl) => {
      tbl.querySelectorAll("th, td").forEach((cell) => {
        const cellEl = cell as HTMLElement;
        if (cellEl.className && typeof cellEl.className === "string" && cellEl.className.includes("min-w-[")) {
          cellEl.style.minWidth = "0";
        }
      });
    });
    naturalWidthPx = Math.max(root.scrollWidth || 0, targetWidthPx);
  }

  // Pass 2: smoothly step down font size if table still exceeds paper width
  while (naturalWidthPx > targetWidthPx + 4 && currentFontPt > 6.6) {
    currentFontPt = Math.max(6.6, Math.round((currentFontPt - 0.4) * 10) / 10);
    applyTableTypography(currentFontPt, true);
    naturalWidthPx = Math.max(root.scrollWidth || 0, targetWidthPx);
  }

  return {
    finalFontPt: currentFontPt,
    fittedCleanly: naturalWidthPx <= targetWidthPx + 6,
    tableCount: tables.length,
    naturalWidthPx,
  };
}

/**
 * Extracts a clean, auto-tidied copy of a worksheet ready for preview, printing, or Word/PDF export.
 */
export function extractAndTidyDocument(
  elementTarget: string | HTMLElement,
  options?: {
    paperSize?: "A4" | "F4";
    orientation?: "auto" | "portrait" | "landscape";
    targetDoc?: Document;
  }
): {
  clonedEl: HTMLElement | null;
  meta: TidyExtractResult;
} {
  const sourceEl =
    typeof elementTarget === "string" ? document.getElementById(elementTarget) : elementTarget;

  const paperSize: "A4" | "F4" = options?.paperSize === "F4" ? "F4" : "A4";
  if (!sourceEl) {
    return {
      clonedEl: null,
      meta: {
        paperSize,
        orientation: "portrait",
        adaptiveFontPt: 9.5,
        maxCols: 0,
        tableCount: 0,
        fittedCleanly: true,
        recommendedScalePercent: 100,
        summaryLabel: `${paperSize} Potret · Arial 9.5pt`,
      },
    };
  }

  const detected = detectOptimalPrintLayout(sourceEl, paperSize);
  const resolvedOrientation: "portrait" | "landscape" =
    options?.orientation && options.orientation !== "auto"
      ? options.orientation
      : detected.orientation;

  const { widthPx } = getPaperPrintablePx(paperSize, resolvedOrientation, 8);
  const clonedEl = cloneElementWithLiveFormValues(sourceEl, options?.targetDoc || document);

  // Run static cleanup immediately (full computed-style bg propagation runs once mounted in stage/iframe)
  const preTidy = tidyExtractedDocumentDom(clonedEl, {
    targetWidthPx: widthPx,
    adaptiveFontPt: detected.adaptiveFontPt,
    maxCols: detected.maxCols,
  });

  const recommendedScalePercent =
    preTidy.naturalWidthPx > widthPx + 6
      ? Math.max(70, Math.floor((widthPx / preTidy.naturalWidthPx) * 100))
      : 100;

  return {
    clonedEl,
    meta: {
      paperSize,
      orientation: resolvedOrientation,
      adaptiveFontPt: preTidy.finalFontPt,
      maxCols: detected.maxCols,
      tableCount: preTidy.tableCount,
      fittedCleanly: preTidy.fittedCleanly,
      recommendedScalePercent,
      summaryLabel: `${paperSize} ${resolvedOrientation === "landscape" ? "Lansekap" : "Potret"} · Arial ${preTidy.finalFontPt}pt`,
    },
  };
}

/**
 * Waits for all <img> elements inside a container to complete loading (with bounded timeout).
 */
async function waitForImagesLoaded(container: HTMLElement, maxWaitMs = 2500): Promise<void> {
  const imgs = Array.from(container.querySelectorAll("img"));
  if (imgs.length === 0) return;

  await Promise.race([
    Promise.all(
      imgs.map(
        (img) =>
          new Promise<void>((resolve) => {
            if (img.complete) {
              resolve();
              return;
            }
            img.addEventListener("load", () => resolve(), { once: true });
            img.addEventListener("error", () => resolve(), { once: true });
          })
      )
    ),
    new Promise<void>((resolve) => setTimeout(resolve, maxWaitMs)),
  ]);
}

interface CapturedStageResult {
  canvas: HTMLCanvasElement;
  renderScale: number;
  stageWidthPx: number;
  stageHeightPx: number;
  finalOrientation: "portrait" | "landscape";
  finalFontPt: number;
  forcedBreaksPx: number[];
  safeRowCutsPx: number[];
}

/**
 * Renders any worksheet inside a dedicated, isolated, exact-paper-width off-screen staging container
 * so exports NEVER clip on the right edge, NEVER suffer Tailwind v4 html2canvas baseline drop,
 * NEVER overlap rowSpan header cells, and expose exact row Y-coordinates for clean PDF pagination.
 */
async function captureTidyDocumentCanvas(
  element: HTMLElement,
  options: {
    paperSize: "A4" | "F4";
    orientation: "auto" | "portrait" | "landscape";
    highRes?: boolean;
  }
): Promise<CapturedStageResult> {
  const detected = detectOptimalPrintLayout(element, options.paperSize);
  const resolvedOrientation: "portrait" | "landscape" =
    options.orientation === "auto" ? detected.orientation : options.orientation;

  const { widthPx: paperPrintablePx } = getPaperPrintablePx(
    options.paperSize,
    resolvedOrientation,
    8
  );

  // Create isolated off-screen staging host mounted directly on document.body
  const stageHost = document.createElement("div");
  stageHost.id = `sispermbg-export-stage-${Date.now()}`;
  stageHost.setAttribute("aria-hidden", "true");
  stageHost.style.position = "fixed";
  stageHost.style.left = "-20000px";
  stageHost.style.top = "0";
  stageHost.style.width = `${paperPrintablePx}px`;
  stageHost.style.backgroundColor = "#ffffff";
  stageHost.style.color = "#000000";
  stageHost.style.margin = "0";
  stageHost.style.padding = "0";
  stageHost.style.zIndex = "-9999";
  stageHost.style.pointerEvents = "none";
  stageHost.style.overflow = "visible";
  stageHost.style.fontFamily = "Arial, Helvetica, sans-serif";

  // Inject temporary scoped stage style + global html2canvas FontMetrics baseline protection
  const stageStyle = document.createElement("style");
  stageStyle.textContent = `
    body > div[style*="visibility: hidden"] > img[width="1"][height="1"] {
      display: inline !important;
      vertical-align: baseline !important;
    }
    #${stageHost.id}, #${stageHost.id} * {
      box-sizing: border-box !important;
      font-family: Arial, Helvetica, sans-serif !important;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    #${stageHost.id} .no-print,
    #${stageHost.id} button:not(.print-keep),
    #${stageHost.id} datalist {
      display: none !important;
    }
    #${stageHost.id} .overflow-x-auto,
    #${stageHost.id} .overflow-auto {
      overflow: visible !important;
      max-width: 100% !important;
    }
    #${stageHost.id} table {
      border-collapse: collapse !important;
    }
  `;

  const clonedContent = cloneElementWithLiveFormValues(element, document);
  // If coming from PrintPreviewModal's #paper-sheet-content, strip its visual preview padding/border/zoom
  if (clonedContent.id === "paper-sheet-content" || clonedContent.classList.contains("paper-sheet")) {
    clonedContent.style.padding = "0";
    clonedContent.style.border = "none";
    clonedContent.style.boxShadow = "none";
    clonedContent.style.width = "100%";
    clonedContent.style.minHeight = "auto";
    clonedContent.style.zoom = "1";
  }

  stageHost.appendChild(clonedContent);
  document.head.appendChild(stageStyle);
  document.body.appendChild(stageHost);

  const origWindowGetComputedStyle = window.getComputedStyle;

  try {
    // Wait for images & browser layout
    await waitForImagesLoaded(stageHost, 2500);
    await new Promise<void>((r) => requestAnimationFrame(() => requestAnimationFrame(() => r())));

    // Run full Auto-Tidy DOM normalization while mounted so computed styles & widths are exact!
    const tidyStats = tidyExtractedDocumentDom(clonedContent, {
      targetWidthPx: paperPrintablePx,
      adaptiveFontPt: detected.adaptiveFontPt,
      maxCols: detected.maxCols,
    });

    // Guarantee zero right-edge clipping: if a super-wide table still needs extra width, expand stage cleanly
    const finalStageWidth = Math.ceil(
      Math.max(paperPrintablePx, stageHost.scrollWidth, clonedContent.scrollWidth)
    );
    if (finalStageWidth > paperPrintablePx) {
      stageHost.style.width = `${finalStageWidth}px`;
      await new Promise<void>((r) => requestAnimationFrame(() => r()));
    }

    const finalStageHeight = Math.ceil(
      Math.max(stageHost.scrollHeight, clonedContent.scrollHeight, 100)
    );

    // Measure exact row & page-break Y-coordinates BEFORE rasterization for clean multi-page PDF breaks
    const stageRect = stageHost.getBoundingClientRect();
    const forcedBreaksSet = new Set<number>();
    const safeCutsSet = new Set<number>();

    stageHost.querySelectorAll("*").forEach((node) => {
      if (!(node instanceof HTMLElement)) return;
      const cls = typeof node.className === "string" ? node.className : "";
      const tag = node.tagName.toLowerCase();

      if (
        cls.includes("break-before-page") ||
        cls.includes("page-break-before") ||
        node.style.pageBreakBefore === "always" ||
        node.style.breakBefore === "page"
      ) {
        const y = Math.round(node.getBoundingClientRect().top - stageRect.top);
        if (y > 24 && y < finalStageHeight - 24) {
          forcedBreaksSet.add(y);
          safeCutsSet.add(y);
        }
      }

      if (
        tag === "tr" ||
        tag === "table" ||
        cls.includes("break-inside-avoid") ||
        cls.includes("signature")
      ) {
        const rect = node.getBoundingClientRect();
        const topY = Math.round(rect.top - stageRect.top);
        const botY = Math.round(rect.bottom - stageRect.top);
        if (topY > 12 && topY < finalStageHeight - 12) safeCutsSet.add(topY);
        if (botY > 12 && botY < finalStageHeight - 12) safeCutsSet.add(botY);
      }
    });

    const area = finalStageWidth * finalStageHeight;
    const baseScale = options.highRes ? 2 : 2;
    const renderScale =
      finalStageHeight > 5200 || area > 5_800_000
        ? 1.35
        : finalStageHeight > 3000 || area > 3_200_000
        ? 1.6
        : baseScale;

    window.getComputedStyle = createSanitizedComputedStyle(origWindowGetComputedStyle);

    const canvas = await html2canvas(stageHost, {
      scale: renderScale,
      useCORS: true,
      allowTaint: true,
      backgroundColor: "#ffffff",
      logging: false,
      width: finalStageWidth,
      height: finalStageHeight,
      windowWidth: finalStageWidth + 40,
      windowHeight: finalStageHeight + 40,
      scrollX: 0,
      scrollY: 0,
      onclone: (clonedDoc) => {
        const clonedWin = clonedDoc.defaultView || window;
        if (clonedWin) {
          clonedWin.getComputedStyle = createSanitizedComputedStyle(
            clonedWin.getComputedStyle.bind(clonedWin)
          );
        }
        // Sanitize any oklch/oklab colors inside cloned stylesheets so html2canvas CSS parser never throws
        clonedDoc.querySelectorAll("style").forEach((s) => {
          const txt = s.textContent || "";
          if (HAS_UNSUPPORTED_COLOR(txt)) {
            s.textContent = sanitizeOklchString(txt);
          }
        });
        // Ensure html2canvas FontMetrics baseline probe img in clonedDoc also uses display: inline
        const fixStyle = clonedDoc.createElement("style");
        fixStyle.textContent = `
          body > div[style*="visibility: hidden"] > img[width="1"][height="1"] {
            display: inline !important;
            vertical-align: baseline !important;
          }
        `;
        clonedDoc.head.appendChild(fixStyle);
      },
    });

    const forcedBreaksPx = Array.from(forcedBreaksSet)
      .map((y) => Math.round(y * renderScale))
      .sort((a, b) => a - b);
    const safeRowCutsPx = Array.from(safeCutsSet)
      .map((y) => Math.round(y * renderScale))
      .sort((a, b) => a - b);

    return {
      canvas,
      renderScale,
      stageWidthPx: finalStageWidth,
      stageHeightPx: finalStageHeight,
      finalOrientation: resolvedOrientation,
      finalFontPt: tidyStats.finalFontPt,
      forcedBreaksPx,
      safeRowCutsPx,
    };
  } finally {
    window.getComputedStyle = origWindowGetComputedStyle;
    stageHost.remove();
    stageStyle.remove();
  }
}

/**
 * Downloads a specific DOM element as a high-resolution, auto-tidied PNG image.
 */
export const downloadElementAsImage = async (
  elementTarget: string | HTMLElement,
  filename: string,
  onProgress?: (msg: string | null) => void,
  options?: { paperSize?: "A4" | "F4"; orientation?: "auto" | "portrait" | "landscape" }
) => {
  const element =
    typeof elementTarget === "string" ? document.getElementById(elementTarget) : elementTarget;
  if (!element) {
    console.error("Element target not found.");
    if (onProgress) onProgress("Elemen tidak ditemukan.");
    return;
  }

  const paperSize = options?.paperSize || "A4";
  const orientation = options?.orientation || "auto";

  if (onProgress) onProgress("Mengekstrak & merapikan tata letak otomatis...");

  try {
    const captured = await captureTidyDocumentCanvas(element, {
      paperSize,
      orientation,
      highRes: true,
    });

    if (onProgress) onProgress("Mengunduh gambar PNG presisi...");

    const imgData = captured.canvas.toDataURL("image/png");
    const link = document.createElement("a");
    link.download = `${filename}.png`;
    link.href = imgData;
    link.click();

    if (onProgress) onProgress(null);
  } catch (error) {
    console.error("Gagal mengunduh gambar:", error);
    if (onProgress) onProgress("Gagal mengunduh gambar. Silakan gunakan Unduh PDF.");
    setTimeout(() => {
      if (onProgress) onProgress(null);
    }, 4000);
  }
};

/**
 * Copies computed visual styles (background color, text color, alignment, weight, borders)
 * from the live DOM onto a detached clone so Word (.doc) exports retain full table shading and borders.
 */
function snapshotLiveStylesForWord(sourceEl: HTMLElement, clonedEl: HTMLElement) {
  const ownerWin = sourceEl.ownerDocument?.defaultView || window;
  const srcNodes = Array.from(sourceEl.querySelectorAll("table, thead, tbody, tfoot, tr, th, td, h1, h2, h3, h4, p, div"));
  const dstNodes = Array.from(clonedEl.querySelectorAll("table, thead, tbody, tfoot, tr, th, td, h1, h2, h3, h4, p, div"));
  const count = Math.min(srcNodes.length, dstNodes.length);

  for (let i = 0; i < count; i++) {
    const src = srcNodes[i] as HTMLElement;
    const dst = dstNodes[i] as HTMLElement;
    const tag = src.tagName.toLowerCase();
    const cs = ownerWin.getComputedStyle(src);

    if (tag === "th" || tag === "td") {
      const bg = sanitizeOklchString(cs.backgroundColor || "");
      if (!isTransparentColor(bg)) {
        dst.style.backgroundColor = bg;
      }
      const color = sanitizeOklchString(cs.color || "");
      if (color) dst.style.color = color;
      if (cs.textAlign) dst.style.textAlign = cs.textAlign;
      if (cs.fontWeight && parseInt(cs.fontWeight, 10) >= 600) {
        dst.style.fontWeight = cs.fontWeight;
      }
      const borderWidth = parseFloat(cs.borderTopWidth || "0");
      if (borderWidth > 0) {
        const bColor = sanitizeOklchString(cs.borderTopColor || "#000000");
        dst.style.border = `${Math.max(1, Math.round(borderWidth))}px solid ${bColor}`;
      } else {
        dst.style.border = "none";
      }
    } else if (tag === "tr" || tag === "thead" || tag === "tbody" || tag === "tfoot") {
      const bg = sanitizeOklchString(cs.backgroundColor || "");
      if (!isTransparentColor(bg)) {
        dst.querySelectorAll(":scope > th, :scope > td, :scope > tr > th, :scope > tr > td").forEach((cell) => {
          const cellEl = cell as HTMLElement;
          if (!cellEl.style.backgroundColor || isTransparentColor(cellEl.style.backgroundColor)) {
            cellEl.style.backgroundColor = bg;
          }
        });
      }
      dst.style.backgroundColor = "transparent";
    } else if (tag === "h1" || tag === "h2" || tag === "h3" || tag === "h4" || tag === "p") {
      if (cs.textAlign) dst.style.textAlign = cs.textAlign;
      if (cs.fontWeight && parseInt(cs.fontWeight, 10) >= 600) {
        dst.style.fontWeight = cs.fontWeight;
      }
    }
  }
}

/**
 * Exports any worksheet directly to a clean, editable Microsoft Word (.doc) file
 * with exact A4/F4 page orientation, uniform Arial typography, crisp borders, KOP, and TTD.
 */
export const exportToWordDoc = (
  elementTarget: string | HTMLElement,
  filename: string,
  options?: { paperSize?: "A4" | "F4"; orientation?: "auto" | "portrait" | "landscape"; title?: string }
) => {
  const sourceEl =
    typeof elementTarget === "string" ? document.getElementById(elementTarget) : elementTarget;
  if (!sourceEl) return;

  const paperSize: "A4" | "F4" = options?.paperSize === "F4" ? "F4" : "A4";
  const detected = detectOptimalPrintLayout(sourceEl, paperSize);
  const resolvedOrientation: "portrait" | "landscape" =
    options?.orientation && options.orientation !== "auto"
      ? options.orientation
      : detected.orientation;

  const { widthPx } = getPaperPrintablePx(paperSize, resolvedOrientation, 8);
  const clonedEl = cloneElementWithLiveFormValues(sourceEl, document);

  // Snapshot live computed cell backgrounds, colors, borders & text alignment BEFORE stripping controls
  snapshotLiveStylesForWord(sourceEl, clonedEl);

  const tidyStats = tidyExtractedDocumentDom(clonedEl, {
    targetWidthPx: widthPx,
    adaptiveFontPt: detected.adaptiveFontPt,
    maxCols: detected.maxCols,
  });

  const { widthMm, heightMm } = getPaperDimensionsMm(paperSize, resolvedOrientation);
  const safeTitle = (options?.title || filename || "Dokumen SISPERMBG").replace(/</g, "&lt;");

  const wordHtml = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8" />
    <title>${safeTitle}</title>
    <!--[if gte mso 9]>
    <xml>
      <w:WordDocument>
        <w:View>Print</w:View>
        <w:Zoom>100</w:Zoom>
        <w:DoNotOptimizeForBrowser/>
      </w:WordDocument>
    </xml>
    <![endif]-->
    <style>
      @page WordSection1 {
        size: ${widthMm}mm ${heightMm}mm;
        mso-page-orientation: ${resolvedOrientation};
        margin: 10mm 10mm 10mm 10mm;
      }
      div.WordSection1 {
        page: WordSection1;
      }
      body {
        font-family: Arial, Helvetica, sans-serif;
        color: #000000;
        background: #ffffff;
        font-size: ${tidyStats.finalFontPt}pt;
        line-height: 1.35;
      }
      table {
        border-collapse: collapse;
        width: 100%;
        font-family: Arial, Helvetica, sans-serif;
      }
      th, td {
        padding: 4px 6px;
        font-family: Arial, Helvetica, sans-serif;
        font-size: ${tidyStats.finalFontPt}pt;
        vertical-align: middle;
      }
      .text-center { text-align: center; }
      .text-right { text-align: right; }
      .text-left { text-align: left; }
      .font-bold { font-weight: bold; }
      .font-semibold { font-weight: 600; }
      .italic { font-style: italic; }
      .uppercase { text-transform: uppercase; }
      .underline { text-decoration: underline; }
      img {
        max-width: 100%;
        height: auto;
      }
    </style>
  </head>
  <body>
    <div class="WordSection1">
      ${clonedEl.outerHTML}
    </div>
  </body>
</html>`;

  const blob = new Blob(["\ufeff", wordHtml], { type: "application/msword;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${filename}.doc`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
};

/**
 * Triggers standard browser printing for a specific target.
 */
export const triggerPrint = () => {
  window.print();
};

/**
 * Precision hidden-iframe printing (no popups) with uniform Arial typography,
 * Auto-Tidy extraction, adaptive font sizing, and automatic Portrait vs Landscape selection.
 */
export const printElementSmart = (
  elementTarget: string | HTMLElement,
  options: SmartPrintOptions = {}
) => {
  const element =
    typeof elementTarget === "string" ? document.getElementById(elementTarget) : elementTarget;
  if (!element) {
    window.print();
    return;
  }

  const paperSize = options.paperSize || "A4";
  const detected = detectOptimalPrintLayout(element, paperSize);
  const resolvedOrientation: "portrait" | "landscape" =
    options.orientation && options.orientation !== "auto"
      ? options.orientation
      : detected.orientation;
  const { widthPx: printableWidthPx } = getPaperPrintablePx(paperSize, resolvedOrientation, 8);
  const pageSizeCss =
    paperSize === "F4" ? `215mm 330mm ${resolvedOrientation}` : `A4 ${resolvedOrientation}`;
  const rawClone = cloneElementWithLiveFormValues(element, document);

  const sanitizedCss = getCachedSanitizedDocumentStyles();

  // Remove any stale hidden print iframe
  const existingFrame = document.getElementById("sispermbg-smart-print-frame");
  if (existingFrame) existingFrame.remove();

  const iframe = document.createElement("iframe");
  iframe.id = "sispermbg-smart-print-frame";
  iframe.setAttribute("aria-hidden", "true");
  iframe.style.position = "fixed";
  iframe.style.left = "-12000px";
  iframe.style.top = "0";
  iframe.style.width = `${printableWidthPx}px`;
  iframe.style.height = "1000px";
  iframe.style.border = "0";
  iframe.style.opacity = "0";
  iframe.style.pointerEvents = "none";
  document.body.appendChild(iframe);

  const iframeDoc = iframe.contentDocument || iframe.contentWindow?.document;
  if (!iframeDoc) {
    iframe.remove();
    window.print();
    return;
  }

  const safeTitle = (options.title || document.title || "Dokumen SISPERMBG").replace(/</g, "&lt;");

  iframeDoc.open();
  iframeDoc.write(`<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>${safeTitle}</title>
    <style>
      ${sanitizedCss}

      *, *::before, *::after {
        box-sizing: border-box !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        font-family: Arial, Helvetica, sans-serif !important;
      }

      @page {
        size: ${pageSizeCss};
        margin: 8mm;
      }

      html, body {
        margin: 0 !important;
        padding: 0 !important;
        width: 100% !important;
        background: #ffffff !important;
        color: #000000 !important;
        font-family: Arial, Helvetica, sans-serif !important;
      }

      .smart-print-root {
        width: 100% !important;
        margin: 0 auto !important;
        font-family: Arial, Helvetica, sans-serif !important;
      }

      .no-print, header, nav, footer, button:not(.print-keep), .tab-bar {
        display: none !important;
      }

      .overflow-x-auto, .overflow-auto {
        overflow: visible !important;
      }

      table {
        border-collapse: collapse !important;
        page-break-inside: auto;
        font-family: Arial, Helvetica, sans-serif !important;
      }

      thead {
        display: table-header-group;
      }

      tr {
        page-break-inside: avoid;
        break-inside: avoid;
      }

      th, td {
        font-family: Arial, Helvetica, sans-serif !important;
        line-height: 1.32 !important;
        vertical-align: middle;
      }

      img {
        max-width: 100% !important;
      }
    </style>
  </head>
  <body>
    <div id="smart-print-root" class="smart-print-root"></div>
  </body>
</html>`);
  iframeDoc.close();

  const printRoot = iframeDoc.getElementById("smart-print-root");
  if (printRoot) {
    const imported = iframeDoc.importNode(rawClone, true);
    printRoot.appendChild(imported);
    const tidyStats = tidyExtractedDocumentDom(imported, {
      targetWidthPx: printableWidthPx,
      adaptiveFontPt: detected.adaptiveFontPt,
      maxCols: detected.maxCols,
    });
    const autoScale =
      options.scalePercent ||
      (!tidyStats.fittedCleanly && tidyStats.naturalWidthPx > printableWidthPx + 6
        ? Math.max(70, Math.floor((printableWidthPx / tidyStats.naturalWidthPx) * 100))
        : 100);
    printRoot.style.zoom = `${autoScale}%`;
  }

  const runPrint = () => {
    try {
      iframe.contentWindow?.focus();
      iframe.contentWindow?.print();
    } catch {
      window.print();
    }
    setTimeout(() => {
      iframe.remove();
    }, 1500);
  };

  const imgs = Array.from(iframeDoc.querySelectorAll("img"));
  if (imgs.length === 0) {
    setTimeout(runPrint, 250);
  } else {
    let loaded = 0;
    let triggered = false;
    const triggerOnce = () => {
      if (triggered) return;
      triggered = true;
      setTimeout(runPrint, 180);
    };
    imgs.forEach((img) => {
      if (img.complete) {
        loaded += 1;
      } else {
        img.onload = () => {
          loaded += 1;
          if (loaded >= imgs.length) triggerOnce();
        };
        img.onerror = () => {
          loaded += 1;
          if (loaded >= imgs.length) triggerOnce();
        };
      }
    });
    if (loaded >= imgs.length) {
      triggerOnce();
    } else {
      setTimeout(triggerOnce, 1200);
    }
  }
};

/**
 * Exports any worksheet directly to a razor-sharp, Auto-Tidied PDF file:
 * - Renders inside an isolated exact-paper-width staging sheet (ZERO right-edge clipping!)
 * - Fixes html2canvas Tailwind v4 baseline drop & rowSpan background overlap universally
 * - Slices multi-page documents strictly on table row / page-break boundaries (ZERO chopped rows!)
 */
export const exportToPDF = async (
  elementTarget: string | HTMLElement,
  filename: string,
  scalePercent: number = 100,
  onProgress?: (msg: string | null) => void,
  options?: { paperSize?: "A4" | "F4"; orientation?: "auto" | "portrait" | "landscape" }
) => {
  const element =
    typeof elementTarget === "string" ? document.getElementById(elementTarget) : elementTarget;
  if (!element) {
    console.error("Element target not found.");
    if (onProgress) onProgress("Elemen tidak ditemukan.");
    return;
  }

  const paperSize = options?.paperSize || "A4";
  const requestedOrientation = options?.orientation || "auto";
  const detected = detectOptimalPrintLayout(element, paperSize);
  const previewOrientation: "portrait" | "landscape" =
    requestedOrientation === "auto" ? detected.orientation : requestedOrientation;

  if (onProgress) {
    onProgress(
      `Ekstrak & rapikan otomatis (${paperSize} ${
        previewOrientation === "landscape" ? "Lansekap" : "Potret"
      } · Arial)...`
    );
  }

  try {
    // Yield a frame so progress message paints immediately
    await new Promise<void>((r) => requestAnimationFrame(() => r()));

    const captured = await captureTidyDocumentCanvas(element, {
      paperSize,
      orientation: requestedOrientation,
      highRes: true,
    });

    if (onProgress) onProgress("Menyusun halaman PDF presisi tanpa terpotong...");
    await new Promise<void>((r) => setTimeout(r, 0));

    const { canvas, finalOrientation, forcedBreaksPx, safeRowCutsPx } = captured;
    const imgWidth = Math.max(canvas.width, 1);
    const imgHeight = Math.max(canvas.height, 1);

    const pdfFormat: "a4" | [number, number] = paperSize === "F4" ? [215, 330] : "a4";

    const pdf = new jsPDF({
      orientation: finalOrientation,
      unit: "mm",
      format: pdfFormat,
      compress: true,
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();

    const marginX = 8;
    const marginY = 8;
    const availWidth = Math.max(10, pdfWidth - marginX * 2);
    const availHeight = Math.max(10, pdfHeight - marginY * 2);

    const safeScaleRatio = Math.min(1, Math.max(0.55, (Number(scalePercent) || 100) / 100));
    const imgAspect = imgHeight / imgWidth;
    let printWidth = availWidth * safeScaleRatio;
    let printHeight = printWidth * imgAspect;

    // Single-page auto-fit when there are no forced page breaks and content slightly exceeds 1 page (<= 10%)
    if (
      forcedBreaksPx.length === 0 &&
      printHeight > availHeight &&
      printHeight <= availHeight * 1.1
    ) {
      printHeight = availHeight;
      printWidth = printHeight / imgAspect;
    }

    const xPos = marginX + Math.max(0, (availWidth - printWidth) / 2);

    if (forcedBreaksPx.length === 0 && printHeight <= availHeight) {
      const pageImg = canvas.toDataURL("image/jpeg", 0.94);
      pdf.addImage(pageImg, "JPEG", xPos, marginY, printWidth, printHeight, undefined, "FAST");
    } else {
      // Smart row-boundary-aware pagination so NO table row, header, photo, or signature is ever sliced in half!
      const pxPerMm = imgWidth / printWidth;
      const maxSlicePx = Math.max(120, Math.floor(availHeight * pxPerMm));
      const sliceCanvas = document.createElement("canvas");
      sliceCanvas.width = imgWidth;
      const sliceCtx = sliceCanvas.getContext("2d", { alpha: false });

      let offsetY = 0;
      let pageIndex = 0;
      const maxPages = 50;

      while (offsetY < imgHeight - 2 && pageIndex < maxPages) {
        const remainingPx = imgHeight - offsetY;
        const rawEndY = Math.min(imgHeight, offsetY + maxSlicePx);

        // 1. Check if an explicit page break falls within this page's vertical span
        const forcedInSpan = forcedBreaksPx.find(
          (y) => y > offsetY + 40 && y <= rawEndY
        );

        let cutEndY = rawEndY;
        if (forcedInSpan !== undefined) {
          cutEndY = forcedInSpan;
        } else if (remainingPx > maxSlicePx) {
          // 2. Snap to the cleanest row/block boundary <= rawEndY (advancing at least 45% of a page)
          const minAcceptableEndY = offsetY + Math.floor(maxSlicePx * 0.45);
          let bestRowCut: number | null = null;
          for (let i = safeRowCutsPx.length - 1; i >= 0; i--) {
            const candidate = safeRowCutsPx[i];
            if (candidate <= rawEndY && candidate >= minAcceptableEndY) {
              bestRowCut = candidate;
              break;
            }
          }
          if (bestRowCut !== null) {
            cutEndY = bestRowCut;
          }
        }

        const currentSlicePx = Math.max(10, Math.min(imgHeight - offsetY, cutEndY - offsetY));
        sliceCanvas.height = currentSlicePx;

        if (sliceCtx) {
          sliceCtx.fillStyle = "#ffffff";
          sliceCtx.fillRect(0, 0, imgWidth, currentSlicePx);
          sliceCtx.drawImage(
            canvas,
            0,
            offsetY,
            imgWidth,
            currentSlicePx,
            0,
            0,
            imgWidth,
            currentSlicePx
          );
        }

        const sliceMmHeight = currentSlicePx / pxPerMm;
        const sliceData = sliceCanvas.toDataURL("image/jpeg", 0.92);
        if (pageIndex > 0) pdf.addPage();
        pdf.addImage(sliceData, "JPEG", xPos, marginY, printWidth, sliceMmHeight, undefined, "FAST");

        offsetY += currentSlicePx;
        pageIndex += 1;

        if (offsetY < imgHeight) {
          await new Promise<void>((r) => setTimeout(r, 0));
        }
      }
    }

    pdf.save(`${filename}.pdf`);
    if (onProgress) onProgress(null);
  } catch (err) {
    console.error("Gagal mengekspor PDF:", err);
    if (onProgress) onProgress("Gagal membuat PDF. Mengalihkan ke Cetak Presisi...");
    setTimeout(() => {
      if (onProgress) onProgress(null);
      printElementSmart(element, {
        title: filename,
        paperSize,
        orientation: requestedOrientation,
      });
    }, 1200);
  }
};
