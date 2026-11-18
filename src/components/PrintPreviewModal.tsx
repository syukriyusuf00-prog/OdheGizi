import React, { useState, useEffect, useRef } from "react";
import {
  X,
  Printer,
  FileText,
  Image as ImageIcon,
  ZoomIn,
  ZoomOut,
  RotateCw,
  Eye,
  Sliders,
  Maximize2,
  Edit3,
  RotateCcw,
} from "lucide-react";
import {
  exportToPDF,
  downloadElementAsImage,
  sanitizeOklchString,
  detectOptimalPrintLayout,
  convertFormControlsToStaticText,
} from "../lib/printUtils";

interface PrintPreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  elementId: string;
  filename?: string;
  defaultScale?: number;
  defaultPaperSize?: "A4" | "F4";
  defaultOrientation?: "auto" | "portrait" | "landscape";
}

export const PrintPreviewModal: React.FC<PrintPreviewModalProps> = ({
  isOpen,
  onClose,
  title,
  elementId,
  filename = "Dokumen_Cetak",
  defaultScale = 100,
  defaultPaperSize = "A4",
  defaultOrientation = "auto",
}) => {
  const [scale, setScale] = useState<number>(defaultScale);
  const [paperSize, setPaperSize] = useState<"A4" | "F4">(defaultPaperSize);
  const [orientationMode, setOrientationMode] = useState<"auto" | "portrait" | "landscape">(
    defaultOrientation
  );
  const [orientation, setOrientation] = useState<"portrait" | "landscape">(
    defaultOrientation === "landscape" ? "landscape" : "portrait"
  );
  const [adaptiveFontPt, setAdaptiveFontPt] = useState<number>(9.2);
  const [autoPrint, setAutoPrint] = useState<boolean>(false);
  const [isEditable, setIsEditable] = useState<boolean>(true);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusMsg, setStatusMsg] = useState<string | null>(null);

  const iframeRef = useRef<HTMLIFrameElement | null>(null);
  const hasAutoPrintedRef = useRef<boolean>(false);

  useEffect(() => {
    if (!isOpen) return;
    const safePaper: "A4" | "F4" = defaultPaperSize === "F4" ? "F4" : "A4";
    setScale(defaultScale);
    setPaperSize(safePaper);
    setOrientationMode(defaultOrientation);
    hasAutoPrintedRef.current = false;

    const targetEl = document.getElementById(elementId);
    const detected = detectOptimalPrintLayout(targetEl, safePaper);
    setAdaptiveFontPt(detected.adaptiveFontPt);
    if (defaultOrientation === "auto") {
      setOrientation(detected.orientation);
    } else {
      setOrientation(defaultOrientation);
    }
  }, [isOpen, defaultScale, defaultPaperSize, defaultOrientation, elementId]);

  // Re-detect optimal Portrait vs Landscape when paperSize changes under Auto mode
  useEffect(() => {
    if (!isOpen) return;
    if (orientationMode === "auto") {
      const targetEl = document.getElementById(elementId);
      const detected = detectOptimalPrintLayout(targetEl, paperSize);
      setOrientation(detected.orientation);
      setAdaptiveFontPt(detected.adaptiveFontPt);
    }
  }, [isOpen, paperSize, orientationMode, elementId]);

  // Initial document clone when modal opens or target element changes
  useEffect(() => {
    if (!isOpen) return;

    const timer = setTimeout(() => {
      initIframeContent();
    }, 150);

    return () => clearTimeout(timer);
  }, [isOpen, elementId]);

  // Dynamically update paper size, orientation, zoom scale & edit mode without resetting user edits
  useEffect(() => {
    if (!isOpen) return;
    updateIframeStyles();
  }, [scale, paperSize, orientation, isEditable, adaptiveFontPt]);

  const initIframeContent = () => {
    const iframe = iframeRef.current;
    if (!iframe) return;

    const targetEl = document.getElementById(elementId);
    if (!targetEl) return;

    const doc = iframe.contentDocument || iframe.contentWindow?.document;
    if (!doc) return;

    const detected = detectOptimalPrintLayout(targetEl, paperSize);
    const resolvedOrientation: "portrait" | "landscape" =
      orientationMode === "auto" ? detected.orientation : orientation;
    const fontPt = detected.adaptiveFontPt;

    // Collect and sanitize all stylesheets from parent document
    let combinedStyles = "";
    document.querySelectorAll("style, link[rel='stylesheet']").forEach((styleEl) => {
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
          // CORS stylesheet bypass
        }
      }
    });

    const sanitizedCss = sanitizeOklchString(combinedStyles);

    let sheetWidthMm = paperSize === "A4" ? 210 : 215;
    let sheetHeightMm = paperSize === "A4" ? 297 : 330;

    if (resolvedOrientation === "landscape") {
      const temp = sheetWidthMm;
      sheetWidthMm = sheetHeightMm;
      sheetHeightMm = temp;
    }

    const pageSizeCss =
      paperSize === "F4" ? `215mm 330mm ${resolvedOrientation}` : `A4 ${resolvedOrientation}`;

    // Clone element and convert form controls to crisp static Arial text
    const clonedEl = targetEl.cloneNode(true) as HTMLElement;
    clonedEl.querySelectorAll(".no-print, button").forEach((el) => el.remove());
    convertFormControlsToStaticText(clonedEl);
    clonedEl.style.margin = "0 auto";
    clonedEl.style.boxShadow = "none";
    clonedEl.style.border = "none";
    clonedEl.style.width = "100%";
    clonedEl.style.transform = "none";
    clonedEl.style.fontFamily = "Arial, Helvetica, sans-serif";

    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <title>${title}</title>
          <style id="dynamic-preview-styles">
            ${sanitizedCss}

            *, *::before, *::after {
              box-sizing: border-box;
              -webkit-print-color-adjust: exact !important;
              print-color-adjust: exact !important;
              color-adjust: exact !important;
              font-family: Arial, Helvetica, sans-serif !important;
            }

            html, body {
              margin: 0;
              padding: 0;
              background-color: #f1f5f9;
              font-family: Arial, Helvetica, sans-serif !important;
              color: #0f172a;
              -webkit-print-color-adjust: exact !important;
              print-color-adjust: exact !important;
            }

            .no-print, header, nav, footer, button {
              display: none !important;
            }

            .overflow-x-auto, .overflow-auto {
              overflow: visible !important;
            }

            .preview-wrapper {
              display: flex;
              justify-content: center;
              padding: 24px;
              min-height: 100vh;
              box-sizing: border-box;
            }

            .paper-sheet {
              background: #ffffff;
              width: ${sheetWidthMm}mm;
              min-height: ${sheetHeightMm}mm;
              padding: 10mm;
              box-sizing: border-box;
              box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
              border-radius: 4px;
              border: 1px solid #cbd5e1;
              zoom: ${scale}%;
              transition: zoom 0.15s ease-in-out;
              outline: none;
              margin: 0 auto;
              font-family: Arial, Helvetica, sans-serif !important;
            }

            .paper-sheet table {
              border-collapse: collapse !important;
              width: 100% !important;
              font-family: Arial, Helvetica, sans-serif !important;
            }

            .paper-sheet th, .paper-sheet td {
              font-family: Arial, Helvetica, sans-serif !important;
              font-size: ${fontPt}pt !important;
              line-height: 1.26 !important;
            }

            .paper-sheet[contenteditable="true"] *:hover {
              outline: 1px dashed #6366f1 !important;
              cursor: text;
            }

            .paper-sheet[contenteditable="true"] *:focus {
              outline: 2px solid #4f46e5 !important;
              outline-offset: 2px;
              background-color: rgba(238, 242, 255, 0.6) !important;
              border-radius: 2px;
            }

            @media print {
              html, body {
                background: #ffffff !important;
                padding: 0 !important;
                margin: 0 !important;
                width: 100% !important;
                font-family: Arial, Helvetica, sans-serif !important;
              }
              .preview-wrapper {
                padding: 0 !important;
                margin: 0 auto !important;
                display: flex !important;
                justify-content: center !important;
                width: 100% !important;
              }
              .paper-sheet {
                box-shadow: none !important;
                border: none !important;
                border-radius: 0 !important;
                width: 100% !important;
                max-width: 100% !important;
                min-height: auto !important;
                padding: 0 !important;
                margin: 0 auto !important;
                zoom: ${scale}% !important;
                outline: none !important;
              }
              .paper-sheet * {
                outline: none !important;
              }
              table {
                border-collapse: collapse !important;
                width: 100% !important;
              }
              thead {
                display: table-header-group;
              }
              tr {
                page-break-inside: avoid;
                break-inside: avoid;
              }
              img {
                max-width: 100% !important;
              }
            }
          </style>
          <style id="dynamic-page-rule">
            @page {
              size: ${pageSizeCss};
              margin: 7mm 8mm;
            }
          </style>
        </head>
        <body>
          <div class="preview-wrapper">
            <div id="paper-sheet-content" class="paper-sheet" contenteditable="${isEditable ? "true" : "false"}" suppresscontenteditablewarning="true">
              ${clonedEl.outerHTML}
            </div>
          </div>
        </body>
      </html>
    `;

    doc.open();
    doc.write(htmlContent);
    doc.close();

    if (autoPrint && !hasAutoPrintedRef.current) {
      hasAutoPrintedRef.current = true;
      setIsProcessing(true);
      setStatusMsg("Membuka dialog cetak presisi...");
      setTimeout(() => {
        if (iframeRef.current && iframeRef.current.contentWindow) {
          try {
            iframeRef.current.contentWindow.focus();
            iframeRef.current.contentWindow.print();
          } catch {
            window.print();
          }
        } else {
          window.print();
        }
        setIsProcessing(false);
        setStatusMsg(null);
      }, 380);
    }
  };

  const updateIframeStyles = () => {
    const iframe = iframeRef.current;
    if (!iframe) return;
    const doc = iframe.contentDocument || iframe.contentWindow?.document;
    if (!doc) return;

    let sheetWidthMm = paperSize === "A4" ? 210 : 215;
    let sheetHeightMm = paperSize === "A4" ? 297 : 330;
    if (orientation === "landscape") {
      const temp = sheetWidthMm;
      sheetWidthMm = sheetHeightMm;
      sheetHeightMm = temp;
    }

    const sheet = doc.getElementById("paper-sheet-content");
    if (sheet) {
      sheet.style.width = `${sheetWidthMm}mm`;
      sheet.style.minHeight = `${sheetHeightMm}mm`;
      sheet.style.zoom = `${scale}%`;
      sheet.setAttribute("contenteditable", isEditable ? "true" : "false");
    }

    const pageRuleEl = doc.getElementById("dynamic-page-rule");
    if (pageRuleEl) {
      const pageSizeCss = paperSize === "F4" ? `215mm 330mm ${orientation}` : `A4 ${orientation}`;
      pageRuleEl.textContent = `@page { size: ${pageSizeCss}; margin: 7mm 8mm; }`;
    }
  };

  const handleSelectOrientation = (mode: "auto" | "portrait" | "landscape") => {
    setOrientationMode(mode);
    if (mode === "auto") {
      const targetEl = document.getElementById(elementId);
      const detected = detectOptimalPrintLayout(targetEl, paperSize);
      setOrientation(detected.orientation);
      setAdaptiveFontPt(detected.adaptiveFontPt);
    } else {
      setOrientation(mode);
    }
  };

  const handlePrintDirect = () => {
    setIsProcessing(true);
    setStatusMsg("Mempersiapkan pencetakan Arial presisi...");
    const iframeDoc = iframeRef.current?.contentDocument;
    if (iframeDoc?.activeElement instanceof HTMLElement) {
      iframeDoc.activeElement.blur();
    }
    setTimeout(() => {
      if (iframeRef.current && iframeRef.current.contentWindow) {
        try {
          iframeRef.current.contentWindow.focus();
          iframeRef.current.contentWindow.print();
        } catch {
          window.print();
        }
      } else {
        window.print();
      }
      setIsProcessing(false);
      setStatusMsg(null);
    }, 260);
  };

  const handlePDFExport = async () => {
    setIsProcessing(true);
    const iframeDoc = iframeRef.current?.contentDocument;
    const paperSheetEl = iframeDoc?.getElementById("paper-sheet-content") as HTMLElement;
    const pdfOpts = { paperSize, orientation };
    if (paperSheetEl) {
      paperSheetEl.setAttribute("contenteditable", "false");
      if (iframeDoc?.activeElement instanceof HTMLElement) {
        iframeDoc.activeElement.blur();
      }
      await exportToPDF(paperSheetEl, filename, scale, (msg) => setStatusMsg(msg), pdfOpts);
      if (isEditable) paperSheetEl.setAttribute("contenteditable", "true");
    } else {
      await exportToPDF(elementId, filename, scale, (msg) => setStatusMsg(msg), pdfOpts);
    }
    setIsProcessing(false);
  };

  const handlePNGExport = async () => {
    setIsProcessing(true);
    const iframeDoc = iframeRef.current?.contentDocument;
    const paperSheetEl = iframeDoc?.getElementById("paper-sheet-content") as HTMLElement;
    if (paperSheetEl) {
      paperSheetEl.setAttribute("contenteditable", "false");
      if (iframeDoc?.activeElement instanceof HTMLElement) {
        iframeDoc.activeElement.blur();
      }
      await downloadElementAsImage(paperSheetEl, filename, (msg) => setStatusMsg(msg));
      if (isEditable) paperSheetEl.setAttribute("contenteditable", "true");
    } else {
      await downloadElementAsImage(elementId, filename, (msg) => setStatusMsg(msg));
    }
    setIsProcessing(false);
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[9999] bg-slate-900/70 backdrop-blur-sm flex flex-col justify-between p-2 sm:p-4 animate-fade-in no-print">
      {/* Top Header & Toolbar */}
      <div className="bg-white rounded-2xl shadow-xl border border-slate-200 p-3 sm:p-4 flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3 shrink-0">
        {/* Left: Title & Format Badge */}
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-indigo-50 text-indigo-600 rounded-xl shrink-0">
            <Eye className="w-5 h-5" />
          </div>
          <div>
            <div className="flex flex-wrap items-center gap-2">
              <h3 className="text-base font-bold text-slate-800">{title}</h3>
              <span className="px-2 py-0.5 text-[11px] font-semibold bg-emerald-50 text-emerald-800 border border-emerald-200 rounded-md">
                Font: Arial Presisi ({adaptiveFontPt}pt)
              </span>
            </div>
            <p className="text-xs text-slate-500">
              Tata letak seragam Arial otomatis menyesuaikan Potret atau Lansekap
            </p>
          </div>
        </div>

        {/* Center: Controls for Paper Size, Orientation, Scale */}
        <div className="flex flex-wrap items-center gap-2.5 bg-slate-50 p-2 rounded-xl border border-slate-200/80">
          {/* Paper Size Selector */}
          <div className="flex items-center gap-1 bg-white p-1 rounded-lg border border-slate-200 text-xs font-semibold">
            <span className="px-1.5 text-slate-500 flex items-center gap-1">
              <FileText className="w-3.5 h-3.5" /> Kertas:
            </span>
            <button
              type="button"
              onClick={() => setPaperSize("A4")}
              className={`px-2 py-1 rounded-md transition ${
                paperSize === "A4"
                  ? "bg-indigo-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              A4
            </button>
            <button
              type="button"
              onClick={() => setPaperSize("F4")}
              className={`px-2 py-1 rounded-md transition ${
                paperSize === "F4"
                  ? "bg-indigo-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              F4 / Folio
            </button>
          </div>

          {/* Orientation Toggle (Auto / Potret / Lansekap) */}
          <div className="flex items-center gap-1 bg-white p-1 rounded-lg border border-slate-200 text-xs font-semibold">
            <span className="px-1.5 text-slate-500 flex items-center gap-1">
              <RotateCw className="w-3.5 h-3.5" /> Orientasi:
            </span>
            <button
              type="button"
              onClick={() => handleSelectOrientation("auto")}
              className={`px-2 py-1 rounded-md transition ${
                orientationMode === "auto"
                  ? "bg-emerald-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
              title="Otomatis Potret bila muat, berpindah Lansekap bila tabel lebar"
            >
              Auto ({orientation === "portrait" ? "Potret" : "Lansekap"})
            </button>
            <button
              type="button"
              onClick={() => handleSelectOrientation("portrait")}
              className={`px-2 py-1 rounded-md transition ${
                orientationMode === "portrait"
                  ? "bg-indigo-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              Potret
            </button>
            <button
              type="button"
              onClick={() => handleSelectOrientation("landscape")}
              className={`px-2 py-1 rounded-md transition ${
                orientationMode === "landscape"
                  ? "bg-indigo-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              Lansekap
            </button>
          </div>

          {/* Zoom Scale Selector */}
          <div className="flex items-center gap-1 bg-white p-1 px-2 rounded-lg border border-slate-200 text-xs font-semibold">
            <span className="text-slate-500 flex items-center gap-1">
              <Sliders className="w-3.5 h-3.5" />
            </span>
            <button
              type="button"
              onClick={() => setScale((s) => Math.max(50, s - 5))}
              className="p-1 text-slate-600 hover:bg-slate-100 rounded"
              title="Perkecil Skala"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span className="w-11 text-center font-bold text-indigo-700 bg-indigo-50 px-1 py-0.5 rounded tabular-nums">
              {scale}%
            </span>
            <button
              type="button"
              onClick={() => setScale((s) => Math.min(150, s + 5))}
              className="p-1 text-slate-600 hover:bg-slate-100 rounded"
              title="Perbesar Skala"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>

            <div className="flex items-center gap-1 ml-1 border-l border-slate-200 pl-1.5">
              {[80, 90, 100].map((preset) => (
                <button
                  key={preset}
                  type="button"
                  onClick={() => setScale(preset)}
                  className={`px-1.5 py-0.5 text-[10px] rounded ${
                    scale === preset ? "bg-slate-800 text-white" : "text-slate-500 hover:bg-slate-100"
                  }`}
                >
                  {preset}%
                </button>
              ))}
            </div>
          </div>

          {/* Live Text Editing Mode Toggle */}
          <button
            type="button"
            onClick={() => setIsEditable((prev) => !prev)}
            className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer ${
              isEditable
                ? "bg-indigo-50 border-indigo-300 text-indigo-800"
                : "bg-white border-slate-200 text-slate-500 hover:bg-slate-100"
            }`}
            title="Klik teks pada kertas pratinjau untuk langsung menyunting sebelum cetak"
          >
            <Edit3 className={`w-3.5 h-3.5 ${isEditable ? "text-indigo-600" : "text-slate-400"}`} />
            <span>Edit Teks: <strong>{isEditable ? "ON" : "OFF"}</strong></span>
          </button>

          {/* Reset Content Button */}
          <button
            type="button"
            onClick={initIframeContent}
            className="px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-600 bg-white hover:bg-slate-100 transition cursor-pointer flex items-center gap-1.5"
            title="Kembalikan teks kertas ke susunan asli dari sistem"
          >
            <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
            <span>Reset</span>
          </button>

          {/* Auto-Print Toggle */}
          <button
            type="button"
            onClick={() => setAutoPrint((prev) => !prev)}
            className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer ${
              autoPrint
                ? "bg-emerald-50 border-emerald-300 text-emerald-800"
                : "bg-white border-slate-200 text-slate-500 hover:bg-slate-100"
            }`}
            title="Buka dialog cetak otomatis saat pratinjau dibuka"
          >
            <Printer className={`w-3.5 h-3.5 ${autoPrint ? "text-emerald-600" : "text-slate-400"}`} />
            <span>Auto-Print: <strong>{autoPrint ? "ON" : "OFF"}</strong></span>
          </button>
        </div>

        {/* Right: Actions & Close */}
        <div className="flex items-center gap-2 justify-end shrink-0">
          <button
            type="button"
            onClick={handlePrintDirect}
            disabled={isProcessing}
            className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 active:scale-95 text-white text-xs font-bold rounded-xl shadow-sm transition flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
          >
            <Printer className="w-4 h-4" />
            <span>Cetak Langsung</span>
          </button>

          <button
            type="button"
            onClick={handlePDFExport}
            disabled={isProcessing}
            className="px-3.5 py-2 bg-rose-600 hover:bg-rose-700 active:scale-95 text-white text-xs font-bold rounded-xl shadow-sm transition flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
          >
            <FileText className="w-4 h-4" />
            <span className="hidden sm:inline">Unduh PDF</span>
          </button>

          <button
            type="button"
            onClick={handlePNGExport}
            disabled={isProcessing}
            className="px-3.5 py-2 bg-slate-800 hover:bg-slate-900 active:scale-95 text-white text-xs font-bold rounded-xl shadow-sm transition flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
          >
            <ImageIcon className="w-4 h-4" />
            <span className="hidden sm:inline">PNG</span>
          </button>

          <button
            type="button"
            onClick={onClose}
            className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition cursor-pointer ml-1"
            title="Tutup Pratinjau"
          >
            <X className="w-5 h-5" />
          </button>
        </div>
      </div>

      {/* Status banner if processing */}
      {statusMsg && (
        <div className="my-2 p-2.5 bg-indigo-600 text-white text-xs font-semibold rounded-xl text-center shadow-md">
          {statusMsg}
        </div>
      )}

      {/* Center: Live Iframe Container */}
      <div className="flex-1 my-2 bg-slate-800/50 rounded-2xl border border-slate-700 shadow-inner overflow-hidden relative">
        <iframe
          ref={iframeRef}
          title="Print Preview Frame"
          className="w-full h-full border-none rounded-2xl bg-slate-100"
        />
      </div>

      {/* Bottom Footer Tip */}
      <div className="bg-white/95 backdrop-blur-md rounded-xl p-2.5 px-4 border border-slate-200 text-xs text-slate-600 flex flex-wrap items-center justify-between gap-2 shrink-0">
        <div className="flex items-center gap-2">
          <Maximize2 className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
          <span>
            Format: <strong>{paperSize} ({orientation === "portrait" ? "Potret" : "Lansekap"})</strong> — Font:{" "}
            <strong>Arial ({adaptiveFontPt}pt)</strong> — Skala: <strong>{scale}%</strong>
          </span>
        </div>
        <div className="flex items-center gap-2 text-[11px] text-indigo-700 bg-indigo-50 px-2.5 py-1 rounded-lg border border-indigo-100 font-semibold">
          <Edit3 className="w-3 h-3 text-indigo-600 shrink-0" />
          <span>
            <strong>Tips:</strong> Klik teks pada kertas di atas bila ingin menyesuaikan sebelum mencetak atau mengunduh PDF.
          </span>
        </div>
      </div>
    </div>
  );
};

export default PrintPreviewModal;

