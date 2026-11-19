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
  const s_ = l - 0.0894841775 * a - 1.2914855480 * b;

  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;

  const rLin = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLin = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLin = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.7076147010 * s3;

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
  const s_ = l - 0.0894841775 * a - 1.2914855480 * b;

  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;

  const rLin = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLin = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLin = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.7076147010 * s3;

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
      if (converted && !converted.includes("oklab") && !converted.includes("oklch") && !converted.includes("color(") && !converted.includes("color-mix")) {
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
      get(target, prop, receiver) {
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
      }
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
  const portraitPrintablePx = paperSize === "F4" ? 736 : 718;

  let maxCols = 0;
  let widestTablePx = 0;

  element.querySelectorAll("table").forEach((table) => {
    if (table.closest(".no-print")) return;
    const htmlTable = table as HTMLElement;
    const style = window.getComputedStyle(htmlTable);
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

  // Switch to landscape when the table has many columns (>= 9) or naturally overflows Portrait width
  const needsLandscape =
    maxCols >= 9 ||
    widestTablePx > portraitPrintablePx * 1.06 ||
    (maxCols >= 7 && widestTablePx > portraitPrintablePx * 0.98);

  const orientation: "portrait" | "landscape" = needsLandscape ? "landscape" : "portrait";

  let adaptiveFontPt = 9.5;
  if (maxCols >= 14) {
    adaptiveFontPt = 7.2;
  } else if (maxCols >= 11) {
    adaptiveFontPt = 7.8;
  } else if (maxCols >= 9) {
    adaptiveFontPt = 8.4;
  } else if (maxCols >= 7) {
    adaptiveFontPt = 8.8;
  }

  return {
    orientation,
    adaptiveFontPt,
    maxCols,
    contentWidthPx,
  };
}

/**
 * Replaces form inputs, textareas, and selects inside a cloned DOM tree with crisp static text
 * spans so printed sheets and exported PDFs render cleanly in Arial without clipping or borders.
 */
export function convertFormControlsToStaticText(root: HTMLElement) {
  const doc = root.ownerDocument || document;

  root.querySelectorAll("input, textarea, select").forEach((ctrl) => {
    if (ctrl.closest(".no-print")) {
      ctrl.remove();
      return;
    }

    if (ctrl instanceof HTMLInputElement) {
      if (ctrl.type === "hidden" || ctrl.type === "file") {
        ctrl.remove();
        return;
      }
      if (ctrl.type === "checkbox" || ctrl.type === "radio") {
        const mark = doc.createElement("span");
        mark.textContent = ctrl.checked ? "✓" : "☐";
        mark.style.fontFamily = "Arial, Helvetica, sans-serif";
        mark.style.fontWeight = "700";
        ctrl.replaceWith(mark);
        return;
      }

      const span = doc.createElement("span");
      span.className = ctrl.className;
      span.textContent = ctrl.value || "";
      span.style.display = "inline-block";
      span.style.border = "none";
      span.style.background = "transparent";
      span.style.boxShadow = "none";
      span.style.padding = "0";
      span.style.fontFamily = "Arial, Helvetica, sans-serif";
      span.style.whiteSpace = "pre-wrap";
      span.style.wordBreak = "break-word";
      ctrl.replaceWith(span);
      return;
    }

    if (ctrl instanceof HTMLTextAreaElement) {
      const div = doc.createElement("div");
      div.className = ctrl.className;
      div.textContent = ctrl.value || "";
      div.style.border = "none";
      div.style.background = "transparent";
      div.style.boxShadow = "none";
      div.style.padding = "0";
      div.style.fontFamily = "Arial, Helvetica, sans-serif";
      div.style.whiteSpace = "pre-wrap";
      div.style.wordBreak = "break-word";
      ctrl.replaceWith(div);
      return;
    }

    if (ctrl instanceof HTMLSelectElement) {
      const selectedOpt = ctrl.options[ctrl.selectedIndex];
      const span = doc.createElement("span");
      span.className = ctrl.className;
      span.textContent = selectedOpt ? selectedOpt.text : ctrl.value || "";
      span.style.display = "inline-block";
      span.style.border = "none";
      span.style.background = "transparent";
      span.style.boxShadow = "none";
      span.style.padding = "0";
      span.style.fontFamily = "Arial, Helvetica, sans-serif";
      ctrl.replaceWith(span);
    }
  });
}

/**
 * Applies onclone DOM transformations and getComputedStyle patch for html2canvas
 */
function prepareClonedDoc(
  clonedDoc: Document,
  elementTarget: string | HTMLElement,
  adaptiveFontPt: number = 9
) {
  const clonedWin = clonedDoc.defaultView || window;
  if (clonedWin) {
    clonedWin.getComputedStyle = createSanitizedComputedStyle(clonedWin.getComputedStyle.bind(clonedWin));
  }

  // 1. Unwrap all @media print { ... } CSS rules into regular top-level CSS so html2canvas applies print styles
  let printCssContent = "";
  clonedDoc.querySelectorAll("style").forEach((styleEl) => {
    const text = styleEl.textContent || "";
    if (text.includes("@media")) {
      const mediaPrintMatches = text.match(/@media\s+print\s*\{([\s\S]*?)\}/gi);
      if (mediaPrintMatches) {
        mediaPrintMatches.forEach((m) => {
          const innerCss = m.replace(/^@media\s+print\s*\{/i, "").replace(/\}\s*$/, "");
          printCssContent += "\n" + innerCss;
        });
      }
    }
    if (HAS_UNSUPPORTED_COLOR(text)) {
      styleEl.textContent = sanitizeOklchString(text);
    }
  });

  if (printCssContent.trim()) {
    const newStyle = clonedDoc.createElement("style");
    newStyle.setAttribute("type", "text/css");
    newStyle.textContent = sanitizeOklchString(printCssContent);
    clonedDoc.head.appendChild(newStyle);
  }

  // Inject global print overrides to enforce uniform Arial typography & crisp paper rendering
  const overrideStyle = clonedDoc.createElement("style");
  overrideStyle.setAttribute("type", "text/css");
  overrideStyle.textContent = `
    * {
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
      box-sizing: border-box !important;
      font-family: Arial, Helvetica, sans-serif !important;
    }
    body, html {
      background-color: #ffffff !important;
      color: #000000 !important;
      margin: 0 !important;
      padding: 0 !important;
      width: 100% !important;
      font-family: Arial, Helvetica, sans-serif !important;
    }
    .no-print, header, nav, footer, button, .tab-bar {
      display: none !important;
    }
    .paper-sheet {
      box-shadow: none !important;
      border: none !important;
      border-radius: 0 !important;
      zoom: 1 !important;
      padding: 0 !important;
      margin: 0 auto !important;
      width: 100% !important;
    }
    table {
      border-collapse: collapse !important;
      width: 100% !important;
      font-family: Arial, Helvetica, sans-serif !important;
    }
    th, td {
      font-family: Arial, Helvetica, sans-serif !important;
      font-size: ${adaptiveFontPt}pt !important;
      line-height: 1.28 !important;
    }
    .overflow-x-auto, .overflow-auto {
      overflow: visible !important;
    }
  `;
  clonedDoc.head.appendChild(overrideStyle);

  // 2. Sanitize rules in document.styleSheets
  try {
    Array.from(clonedDoc.styleSheets).forEach((sheet) => {
      try {
        if (sheet.cssRules) {
          Array.from(sheet.cssRules).forEach((rule) => {
            if (rule.cssText && HAS_UNSUPPORTED_COLOR(rule.cssText)) {
              if ("style" in rule && (rule as CSSStyleRule).style) {
                const style = (rule as CSSStyleRule).style;
                for (let i = 0; i < style.length; i++) {
                  const prop = style[i];
                  const val = style.getPropertyValue(prop);
                  if (val && HAS_UNSUPPORTED_COLOR(val)) {
                    style.setProperty(prop, sanitizeOklchString(val));
                  }
                }
              }
            }
          });
        }
      } catch {
        // Ignore CORS stylesheet errors
      }
    });
  } catch {
    // Ignore stylesheet access errors
  }

  // 3. Sanitize inline styles on all elements and remove zoom
  clonedDoc.querySelectorAll("*").forEach((el) => {
    if (el instanceof HTMLElement || el instanceof SVGElement) {
      const styleAttr = el.getAttribute("style");
      if (styleAttr && HAS_UNSUPPORTED_COLOR(styleAttr)) {
        el.setAttribute("style", sanitizeOklchString(styleAttr));
      }
      if (el instanceof HTMLElement && el.style.zoom) {
        el.style.zoom = "1";
      }
    }
  });

  // 4. Adjust target element layout & convert form controls cleanly
  let clonedEl: HTMLElement | null = null;
  if (typeof elementTarget === "string") {
    clonedEl = clonedDoc.getElementById(elementTarget);
  } else if (elementTarget) {
    if (elementTarget.id) {
      clonedEl = clonedDoc.getElementById(elementTarget.id);
    }
    if (!clonedEl) {
      clonedEl = (clonedDoc.querySelector(".paper-sheet") || clonedDoc.body?.firstElementChild) as HTMLElement;
    }
  }

  if (clonedEl) {
    clonedEl.querySelectorAll(".no-print").forEach((node) => node.remove());
    convertFormControlsToStaticText(clonedEl);
    clonedEl.style.transform = "none";
    clonedEl.style.margin = "0 auto";
    clonedEl.style.boxShadow = "none";
    clonedEl.style.border = "none";
    clonedEl.style.width = "100%";
    clonedEl.style.zoom = "1";
    clonedEl.style.fontFamily = "Arial, Helvetica, sans-serif";
  }
}

/**
 * Downloads a specific DOM element as a high-resolution PNG image.
 * Uses html2canvas with uniform Arial typography, cross-origin support, and OKLCH/OKLAB sanitization.
 */
export const downloadElementAsImage = async (
  elementTarget: string | HTMLElement,
  filename: string,
  onProgress?: (msg: string | null) => void
) => {
  const element = typeof elementTarget === "string" ? document.getElementById(elementTarget) : elementTarget;
  if (!element) {
    console.error(`Element target not found.`);
    if (onProgress) onProgress("Elemen tidak ditemukan.");
    return;
  }

  if (onProgress) onProgress("Merapikan tata letak Arial & memproses gambar...");

  const origWindowGetComputedStyle = window.getComputedStyle;
  const targetDoc = element.ownerDocument || document;
  const targetWin = targetDoc.defaultView || window;
  const layout = detectOptimalPrintLayout(element, "A4");

  try {
    window.getComputedStyle = createSanitizedComputedStyle(origWindowGetComputedStyle);

    const canvas = await html2canvas(element, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      backgroundColor: "#ffffff",
      logging: false,
      scrollX: 0,
      scrollY: 0,
      windowWidth: targetDoc.documentElement ? targetDoc.documentElement.scrollWidth : targetWin.innerWidth,
      windowHeight: targetDoc.documentElement ? targetDoc.documentElement.scrollHeight : targetWin.innerHeight,
      onclone: (clonedDoc) => {
        prepareClonedDoc(clonedDoc, elementTarget, layout.adaptiveFontPt);
      }
    });

    if (onProgress) onProgress("Mengunduh gambar...");

    const imgData = canvas.toDataURL("image/png");
    const link = document.createElement("a");
    link.download = `${filename}.png`;
    link.href = imgData;
    link.click();

    if (onProgress) onProgress(null);
  } catch (error) {
    console.error("Gagal mengunduh gambar:", error);
    if (onProgress) onProgress("Gagal mengunduh gambar. Silakan gunakan Cetak PDF.");
    setTimeout(() => {
      if (onProgress) onProgress(null);
    }, 4000);
  } finally {
    window.getComputedStyle = origWindowGetComputedStyle;
  }
};

/**
 * Triggers standard browser printing for a specific target.
 */
export const triggerPrint = () => {
  window.print();
};

/**
 * Precision hidden-iframe printing (no popups) with uniform Arial typography,
 * adaptive font sizing, and automatic Portrait vs Landscape selection.
 */
export const printElementSmart = (
  elementTarget: string | HTMLElement,
  options: SmartPrintOptions = {}
) => {
  const element = typeof elementTarget === "string" ? document.getElementById(elementTarget) : elementTarget;
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

  const pageSizeCss = paperSize === "F4" ? `215mm 330mm ${resolvedOrientation}` : `A4 ${resolvedOrientation}`;
  const scalePercent = options.scalePercent || 100;

  // Reuse cached sanitized stylesheets from current document
  const sanitizedCss = getCachedSanitizedDocumentStyles();

  // Clone element cleanly and replace inputs/selects with static Arial text
  const cloned = element.cloneNode(true) as HTMLElement;
  cloned.querySelectorAll(".no-print, button").forEach((n) => n.remove());
  convertFormControlsToStaticText(cloned);
  cloned.style.margin = "0 auto";
  cloned.style.boxShadow = "none";
  cloned.style.border = "none";
  cloned.style.width = "100%";
  cloned.style.maxWidth = "100%";
  cloned.style.transform = "none";

  // Remove any stale hidden print iframe
  const existingFrame = document.getElementById("sispermbg-smart-print-frame");
  if (existingFrame) existingFrame.remove();

  const iframe = document.createElement("iframe");
  iframe.id = "sispermbg-smart-print-frame";
  iframe.setAttribute("aria-hidden", "true");
  iframe.style.position = "fixed";
  iframe.style.right = "0";
  iframe.style.bottom = "0";
  iframe.style.width = "0";
  iframe.style.height = "0";
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
        margin: 7mm 8mm;
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
        zoom: ${scalePercent}%;
        font-family: Arial, Helvetica, sans-serif !important;
      }

      .no-print, header, nav, footer, button, .tab-bar {
        display: none !important;
      }

      .overflow-x-auto, .overflow-auto {
        overflow: visible !important;
      }

      table {
        width: 100% !important;
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
        font-size: ${detected.adaptiveFontPt}pt !important;
        line-height: 1.25 !important;
      }

      img {
        max-width: 100% !important;
      }
    </style>
  </head>
  <body>
    <div class="smart-print-root">
      ${cloned.outerHTML}
    </div>
  </body>
</html>`);
  iframeDoc.close();

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

  // Wait briefly for all images inside cloned document to finish loading
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
 * Exports a specific DOM element directly to a PDF file using html2canvas and jsPDF,
 * automatically enforcing uniform Arial typography, F4/A4 sizing, and auto Portrait/Landscape selection.
 */
export const exportToPDF = async (
  elementTarget: string | HTMLElement,
  filename: string,
  _scalePercent: number = 100,
  onProgress?: (msg: string | null) => void,
  options?: { paperSize?: "A4" | "F4"; orientation?: "auto" | "portrait" | "landscape" }
) => {
  const element = typeof elementTarget === "string" ? document.getElementById(elementTarget) : elementTarget;
  if (!element) {
    console.error(`Element target not found.`);
    if (onProgress) onProgress("Elemen tidak ditemukan.");
    return;
  }

  const paperSize = options?.paperSize || "A4";
  const detected = detectOptimalPrintLayout(element, paperSize);
  const requestedOrientation = options?.orientation || "auto";
  const resolvedOrientation: "portrait" | "landscape" =
    requestedOrientation === "auto" ? detected.orientation : requestedOrientation;

  if (onProgress) {
    onProgress(
      `Merapikan format Arial (${paperSize} ${resolvedOrientation === "landscape" ? "Lansekap" : "Potret"})...`
    );
  }

  const origWindowGetComputedStyle = window.getComputedStyle;
  const targetDoc = element.ownerDocument || document;
  const targetWin = targetDoc.defaultView || window;

  try {
    window.getComputedStyle = createSanitizedComputedStyle(origWindowGetComputedStyle);

    // Yield a frame so progress message paints before html2canvas starts
    await new Promise<void>((r) => requestAnimationFrame(() => r()));

    const elW = Math.max(element.scrollWidth || 0, element.offsetWidth || 800);
    const elH = Math.max(element.scrollHeight || 0, element.offsetHeight || 600);
    const area = elW * elH;
    // Dynamically protect Chrome memory on tall photo-heavy sheets (e.g., 20-30 day IKSS photo tables)
    const renderScale =
      elH > 4500 || area > 5_000_000 ? 1.25 : elH > 2600 || area > 2_600_000 ? 1.5 : 2;

    const canvas = await html2canvas(element, {
      scale: renderScale,
      useCORS: true,
      allowTaint: true,
      backgroundColor: "#ffffff",
      logging: false,
      scrollX: 0,
      scrollY: 0,
      windowWidth: targetDoc.documentElement ? targetDoc.documentElement.scrollWidth : targetWin.innerWidth,
      windowHeight: targetDoc.documentElement ? targetDoc.documentElement.scrollHeight : targetWin.innerHeight,
      onclone: (clonedDoc) => {
        prepareClonedDoc(clonedDoc, elementTarget, detected.adaptiveFontPt);
      }
    });

    if (onProgress) onProgress("Menyusun halaman PDF presisi...");
    await new Promise<void>((r) => setTimeout(r, 0));

    const imgWidth = Math.max(canvas.width, 1);
    const imgHeight = Math.max(canvas.height, 1);

    const finalOrientation: "portrait" | "landscape" =
      requestedOrientation === "auto"
        ? detected.orientation === "landscape" || imgWidth / imgHeight > 1.25
          ? "landscape"
          : "portrait"
        : resolvedOrientation;

    const pdfFormat: "a4" | [number, number] = paperSize === "F4" ? [215, 330] : "a4";

    const pdf = new jsPDF({
      orientation: finalOrientation,
      unit: "mm",
      format: pdfFormat,
      compress: true,
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();

    const marginX = 7;
    const marginY = 7;
    const availWidth = Math.max(10, pdfWidth - marginX * 2);
    const availHeight = Math.max(10, pdfHeight - marginY * 2);

    const imgAspect = imgHeight / imgWidth;
    let printWidth = availWidth;
    let printHeight = printWidth * imgAspect;

    // If page content slightly exceeds single page height, fit neatly onto one page
    if (printHeight > availHeight && printHeight <= availHeight * 1.15) {
      printHeight = availHeight;
      printWidth = printHeight / imgAspect;
    }

    const xPos = marginX + Math.max(0, (availWidth - printWidth) / 2);

    if (printHeight <= availHeight) {
      const pageImg = canvas.toDataURL("image/jpeg", 0.92);
      pdf.addImage(pageImg, "JPEG", xPos, marginY, printWidth, printHeight, undefined, "FAST");
    } else {
      // Slice canvas per page so multi-page PDFs never duplicate huge full-document images in memory
      const pxPerMm = imgWidth / printWidth;
      const pageSlicePx = Math.max(100, Math.floor(availHeight * pxPerMm));
      const sliceCanvas = document.createElement("canvas");
      sliceCanvas.width = imgWidth;
      const sliceCtx = sliceCanvas.getContext("2d", { alpha: false });

      let offsetY = 0;
      let pageIndex = 0;
      const maxPages = 40;

      while (offsetY < imgHeight && pageIndex < maxPages) {
        const remainingPx = imgHeight - offsetY;
        const currentSlicePx = Math.min(pageSlicePx, remainingPx);
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
        const sliceData = sliceCanvas.toDataURL("image/jpeg", 0.9);
        if (pageIndex > 0) pdf.addPage();
        pdf.addImage(sliceData, "JPEG", xPos, marginY, printWidth, sliceMmHeight, undefined, "FAST");

        offsetY += currentSlicePx;
        pageIndex += 1;

        // Yield cleanly between pages so Chrome never marks the page unresponsive
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
  } finally {
    window.getComputedStyle = origWindowGetComputedStyle;
  }
};



