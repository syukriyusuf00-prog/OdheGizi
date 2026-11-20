import React, { useState, useEffect, useRef, useCallback } from "react";
import {
  X,
  Printer,
  FileText,
  FileDown,
  Image as ImageIcon,
  ZoomIn,
  ZoomOut,
  RotateCw,
  Eye,
  Sliders,
  Maximize2,
  Edit3,
  RotateCcw,
  Sparkles,
  CheckCircle2,
} from "lucide-react";
import {
  exportToPDF,
  exportToWordDoc,
  downloadElementAsImage,
  getCachedSanitizedDocumentStyles,
  detectOptimalPrintLayout,
  cloneElementWithLiveFormValues,
  tidyExtractedDocumentDom,
  getPaperDimensionsMm,
  getPaperPrintablePx,
} from "../lib/printUtils";

interface PrintPreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  elementId: string;
  filename?: string;
  defaultScale?: number;
  defaultPaperSize?: "A4" | "F4";
  paperSize?: "A4" | "F4";
  defaultOrientation?: "auto" | "portrait" | "landscape";
}

export const PrintPreviewModal: React.FC<PrintPreviewModalProps> = ({
  isOpen,
  onClose,
  title,
  elementId,
  filename = "Dokumen_Cetak",
  defaultScale = 100,
  defaultPaperSize,
  paperSize: paperSizeProp,
  defaultOrientation = "auto",
}) => {
  const initialPaper: "A4" | "F4" =
    (defaultPaperSize || paperSizeProp) === "F4" ? "F4" : "A4";

  const [scale, setScale] = useState<number>(defaultScale);
  const [paperSize, setPaperSize] = useState<"A4" | "F4">(initialPaper);
  const [orientationMode, setOrientationMode] = useState<"auto" | "portrait" | "landscape">(
    defaultOrientation
  );
  const [orientation, setOrientation] = useState<"portrait" | "landscape">(
    defaultOrientation === "landscape" ? "landscape" : "portrait"
  );
  const [adaptiveFontPt, setAdaptiveFontPt] = useState<number>(9.0);
  const [autoTidyMeta, setAutoTidyMeta] = useState<{
    fittedCleanly: boolean;
    tableCount: number;
    maxCols: number;
  }>({ fittedCleanly: true, tableCount: 0, maxCols: 0 });
  const [autoPrint, setAutoPrint] = useState<boolean>(false);
  const [isEditable, setIsEditable] = useState<boolean>(true);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusMsg, setStatusMsg] = useState<string | null>(null);
  const [tidySuccessBanner, setTidySuccessBanner] = useState<string | null>(null);

  const iframeRef = useRef<HTMLIFrameElement | null>(null);
  const hasAutoPrintedRef = useRef<boolean>(false);

  useEffect(() => {
    if (!isOpen) return;
    const safePaper: "A4" | "F4" =
      (defaultPaperSize || paperSizeProp) === "F4" ? "F4" : "A4";
    setScale(defaultScale);
    setPaperSize(safePaper);
    setOrientationMode(defaultOrientation);
    setTidySuccessBanner(null);
    hasAutoPrintedRef.current = false;

    const targetEl = document.getElementById(elementId);
    const detected = detectOptimalPrintLayout(targetEl, safePaper);
    setAdaptiveFontPt(detected.adaptiveFontPt);
    if (defaultOrientation === "auto") {
      setOrientation(detected.orientation);
    } else {
      setOrientation(defaultOrientation);
    }
  }, [isOpen, defaultScale, defaultPaperSize, paperSizeProp, defaultOrientation, elementId]);

  const buildAndRenderIframe = useCallback(
    (opts?: {
      overridePaper?: "A4" | "F4";
      overrideOrientationMode?: "auto" | "portrait" | "landscape";
      overrideOrientation?: "portrait" | "landscape";
      preserveEdits?: boolean;
      showToast?: boolean;
    }) => {
      const iframe = iframeRef.current;
      if (!iframe) return;

      const doc = iframe.contentDocument || iframe.contentWindow?.document;
      if (!doc) return;

      const activePaper = opts?.overridePaper || paperSize;
      const activeOrientMode = opts?.overrideOrientationMode || orientationMode;
      const existingSheet = doc.getElementById("paper-sheet-content") as HTMLElement | null;
      const sourceEl =
        opts?.preserveEdits && existingSheet
          ? existingSheet
          : document.getElementById(elementId);
      if (!sourceEl) return;

      const detected = detectOptimalPrintLayout(sourceEl, activePaper);
      const resolvedOrientation: "portrait" | "landscape" =
        opts?.overrideOrientation
          ? opts.overrideOrientation
          : activeOrientMode === "auto"
          ? detected.orientation
          : orientation;

      const { widthMm: sheetWidthMm, heightMm: sheetHeightMm } = getPaperDimensionsMm(
        activePaper,
        resolvedOrientation
      );
      const { widthPx: printableWidthPx } = getPaperPrintablePx(
        activePaper,
        resolvedOrientation,
        8
      );
      const pageSizeCss =
        activePaper === "F4"
          ? `215mm 330mm ${resolvedOrientation}`
          : `A4 ${resolvedOrientation}`;

      // Clone source while preserving all live input/textarea/select values
      const clonedEl = cloneElementWithLiveFormValues(sourceEl, document);
      if (clonedEl.id === "paper-sheet-content") {
        // Unwrap inner children if re-tidying already-mounted preview content
        clonedEl.removeAttribute("id");
        clonedEl.classList.remove("paper-sheet");
        clonedEl.style.padding = "0";
        clonedEl.style.minHeight = "auto";
      }

      const sanitizedCss = getCachedSanitizedDocumentStyles();
      const safeTitle = (title || "Pratinjau Dokumen").replace(/</g, "&lt;");

      const htmlContent = `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <title>${safeTitle}</title>
    <style id="dynamic-preview-styles">
      ${sanitizedCss}

      *, *::before, *::after {
        box-sizing: border-box !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
        font-family: Arial, Helvetica, sans-serif !important;
      }

      html, body {
        margin: 0;
        padding: 0;
        background-color: #e2e8f0;
        font-family: Arial, Helvetica, sans-serif !important;
        color: #000000;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }

      .no-print, header, nav, footer, button:not(.print-keep), datalist {
        display: none !important;
      }

      .overflow-x-auto, .overflow-auto {
        overflow: visible !important;
        max-width: 100% !important;
      }

      .preview-wrapper {
        display: flex;
        justify-content: center;
        padding: 20px;
        min-height: 100vh;
        box-sizing: border-box;
      }

      .paper-sheet {
        background: #ffffff;
        width: ${sheetWidthMm}mm;
        min-height: ${sheetHeightMm}mm;
        padding: 8mm;
        box-sizing: border-box;
        box-shadow: 0 12px 30px -8px rgba(15, 23, 42, 0.18);
        border-radius: 3px;
        border: 1px solid #cbd5e1;
        zoom: ${scale}%;
        transition: zoom 0.12s ease-in-out;
        outline: none;
        margin: 0 auto;
        font-family: Arial, Helvetica, sans-serif !important;
        color: #000000;
      }

      .paper-sheet table {
        border-collapse: collapse !important;
        font-family: Arial, Helvetica, sans-serif !important;
      }

      .paper-sheet th, .paper-sheet td {
        font-family: Arial, Helvetica, sans-serif !important;
        line-height: 1.35 !important;
        vertical-align: middle;
        word-break: break-word;
        overflow-wrap: break-word;
      }

      .paper-sheet[contenteditable="true"] *:hover {
        outline: 1px dashed #6366f1 !important;
        cursor: text;
      }

      .paper-sheet[contenteditable="true"] *:focus {
        outline: 2px solid #4f46e5 !important;
        outline-offset: 1px;
        background-color: rgba(238, 242, 255, 0.55) !important;
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
          display: block !important;
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
        margin: 8mm;
      }
    </style>
  </head>
  <body>
    <div class="preview-wrapper">
      <div id="paper-sheet-content" class="paper-sheet" contenteditable="${
        isEditable ? "true" : "false"
      }" suppresscontenteditablewarning="true"></div>
    </div>
  </body>
</html>`;

      doc.open();
      doc.write(htmlContent);
      doc.close();

      // Mount cloned content into live iframe sheet and run full Auto-Tidy DOM normalizer
      const sheetHost = doc.getElementById("paper-sheet-content");
      if (sheetHost) {
        const imported = doc.importNode(clonedEl, true);
        sheetHost.appendChild(imported);

        // Run Auto-Tidy while live inside the iframe so computed styles & widths resolve accurately
        const tidyStats = tidyExtractedDocumentDom(imported, {
          targetWidthPx: printableWidthPx,
          adaptiveFontPt: detected.adaptiveFontPt,
          maxCols: detected.maxCols,
        });

        setOrientation(resolvedOrientation);
        setAdaptiveFontPt(tidyStats.finalFontPt);
        setAutoTidyMeta({
          fittedCleanly: tidyStats.fittedCleanly,
          tableCount: tidyStats.tableCount,
          maxCols: detected.maxCols,
        });

        // Auto-fit zoom slightly if an ultra-wide table still exceeded paper width
        if (!tidyStats.fittedCleanly && tidyStats.naturalWidthPx > printableWidthPx + 6) {
          const autoFitPct = Math.max(
            75,
            Math.floor((printableWidthPx / tidyStats.naturalWidthPx) * 100)
          );
          setScale(autoFitPct);
          sheetHost.style.zoom = `${autoFitPct}%`;
        } else if (opts?.showToast) {
          setScale(100);
          sheetHost.style.zoom = "100%";
        }

        if (opts?.showToast) {
          setTidySuccessBanner(
            `Ekstrak rapi otomatis selesai · ${activePaper} ${
              resolvedOrientation === "landscape" ? "Lansekap" : "Potret"
            } · Arial ${tidyStats.finalFontPt}pt · Siap unduh PDF/Word/Cetak`
          );
          window.setTimeout(() => setTidySuccessBanner(null), 4000);
        }
      }

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
    },
    [elementId, paperSize, orientationMode, orientation, scale, isEditable, title, autoPrint]
  );

  // Initial document extraction & auto-tidy when modal opens or target element changes
  useEffect(() => {
    if (!isOpen) return;
    const timer = setTimeout(() => {
      buildAndRenderIframe({ preserveEdits: false, showToast: false });
    }, 100);
    return () => clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, elementId]);

  // Dynamically update paper dimensions, orientation, zoom scale & edit mode without wiping live edits
  useEffect(() => {
    if (!isOpen) return;
    const iframe = iframeRef.current;
    if (!iframe) return;
    const doc = iframe.contentDocument || iframe.contentWindow?.document;
    if (!doc) return;

    const { widthMm: sheetWidthMm, heightMm: sheetHeightMm } = getPaperDimensionsMm(
      paperSize,
      orientation
    );

    const sheet = doc.getElementById("paper-sheet-content");
    if (sheet) {
      sheet.style.width = `${sheetWidthMm}mm`;
      sheet.style.minHeight = `${sheetHeightMm}mm`;
      sheet.style.zoom = `${scale}%`;
      sheet.setAttribute("contenteditable", isEditable ? "true" : "false");
    }

    const pageRuleEl = doc.getElementById("dynamic-page-rule");
    if (pageRuleEl) {
      const pageSizeCss =
        paperSize === "F4" ? `215mm 330mm ${orientation}` : `A4 ${orientation}`;
      pageRuleEl.textContent = `@page { size: ${pageSizeCss}; margin: 8mm; }`;
    }
  }, [isOpen, scale, paperSize, orientation, isEditable]);

  const handleSelectPaperSize = (nextPaper: "A4" | "F4") => {
    setPaperSize(nextPaper);
    let nextOrient = orientation;
    if (orientationMode === "auto") {
      const targetEl = document.getElementById(elementId);
      const detected = detectOptimalPrintLayout(targetEl, nextPaper);
      nextOrient = detected.orientation;
      setOrientation(nextOrient);
      setAdaptiveFontPt(detected.adaptiveFontPt);
    }
    buildAndRenderIframe({
      overridePaper: nextPaper,
      overrideOrientationMode: orientationMode,
      overrideOrientation: nextOrient,
      preserveEdits: true,
      showToast: false,
    });
  };

  const handleSelectOrientation = (mode: "auto" | "portrait" | "landscape") => {
    setOrientationMode(mode);
    if (mode === "auto") {
      const targetEl = document.getElementById(elementId);
      const detected = detectOptimalPrintLayout(targetEl, paperSize);
      setOrientation(detected.orientation);
      setAdaptiveFontPt(detected.adaptiveFontPt);
      buildAndRenderIframe({
        overrideOrientationMode: "auto",
        overrideOrientation: detected.orientation,
        preserveEdits: true,
        showToast: false,
      });
    } else {
      setOrientation(mode);
      buildAndRenderIframe({
        overrideOrientationMode: mode,
        overrideOrientation: mode,
        preserveEdits: true,
        showToast: false,
      });
    }
  };

  const handleAutoTidyNow = () => {
    buildAndRenderIframe({
      overrideOrientationMode: orientationMode,
      preserveEdits: true,
      showToast: true,
    });
  };

  const handleResetOriginal = () => {
    buildAndRenderIframe({
      overrideOrientationMode: "auto",
      preserveEdits: false,
      showToast: true,
    });
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
    }, 240);
  };

  const handlePDFExport = async () => {
    setIsProcessing(true);
    const iframeDoc = iframeRef.current?.contentDocument;
    const paperSheetEl = iframeDoc?.getElementById("paper-sheet-content") as HTMLElement | null;
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

  const handleWordExport = () => {
    const iframeDoc = iframeRef.current?.contentDocument;
    const paperSheetEl = iframeDoc?.getElementById("paper-sheet-content") as HTMLElement | null;
    if (paperSheetEl && iframeDoc?.activeElement instanceof HTMLElement) {
      iframeDoc.activeElement.blur();
    }
    exportToWordDoc(paperSheetEl || elementId, filename, {
      paperSize,
      orientation,
      title,
    });
    setTidySuccessBanner("File Word (.doc) rapi otomatis berhasil diunduh.");
    window.setTimeout(() => setTidySuccessBanner(null), 3500);
  };

  const handlePNGExport = async () => {
    setIsProcessing(true);
    const iframeDoc = iframeRef.current?.contentDocument;
    const paperSheetEl = iframeDoc?.getElementById("paper-sheet-content") as HTMLElement | null;
    const imgOpts = { paperSize, orientation };
    if (paperSheetEl) {
      paperSheetEl.setAttribute("contenteditable", "false");
      if (iframeDoc?.activeElement instanceof HTMLElement) {
        iframeDoc.activeElement.blur();
      }
      await downloadElementAsImage(paperSheetEl, filename, (msg) => setStatusMsg(msg), imgOpts);
      if (isEditable) paperSheetEl.setAttribute("contenteditable", "true");
    } else {
      await downloadElementAsImage(elementId, filename, (msg) => setStatusMsg(msg), imgOpts);
    }
    setIsProcessing(false);
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[9999] bg-slate-900/75 backdrop-blur-sm flex flex-col justify-between p-2 sm:p-4 animate-fade-in no-print">
      {/* Top Header & Toolbar */}
      <div className="bg-white rounded-2xl shadow-xl border border-slate-200 p-3 sm:p-4 flex flex-col xl:flex-row items-stretch xl:items-center justify-between gap-3 shrink-0">
        {/* Left: Title & Auto-Tidy Status */}
        <div className="flex items-center gap-3 min-w-0">
          <div className="p-2.5 bg-indigo-50 text-indigo-600 rounded-xl shrink-0">
            <Eye className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <div className="flex flex-wrap items-center gap-2">
              <h3 className="text-sm sm:text-base font-bold text-slate-900 truncate">{title}</h3>
              <span className="px-2 py-0.5 text-[11px] font-semibold bg-emerald-50 text-emerald-800 border border-emerald-200 rounded-md inline-flex items-center gap-1 shrink-0">
                <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                Ekstrak Rapi · Arial {adaptiveFontPt}pt
              </span>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Tata letak dirapikan otomatis (anti terpotong, sejajar tengah, {paperSize}{" "}
              {orientation === "landscape" ? "Lansekap" : "Potret"}) sebelum diunduh
            </p>
          </div>
        </div>

        {/* Center: Controls for Paper Size, Orientation, Scale & Auto-Tidy */}
        <div className="flex flex-wrap items-center gap-2 bg-slate-50 p-2 rounded-xl border border-slate-200/80">
          {/* Paper Size Selector */}
          <div className="flex items-center gap-1 bg-white p-1 rounded-lg border border-slate-200 text-xs font-semibold">
            <span className="px-1.5 text-slate-500 flex items-center gap-1">
              <FileText className="w-3.5 h-3.5" /> Kertas:
            </span>
            <button
              type="button"
              onClick={() => handleSelectPaperSize("A4")}
              className={`px-2 py-1 rounded-md transition cursor-pointer ${
                paperSize === "A4"
                  ? "bg-indigo-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              A4
            </button>
            <button
              type="button"
              onClick={() => handleSelectPaperSize("F4")}
              className={`px-2 py-1 rounded-md transition cursor-pointer ${
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
              className={`px-2 py-1 rounded-md transition cursor-pointer ${
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
              className={`px-2 py-1 rounded-md transition cursor-pointer ${
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
              className={`px-2 py-1 rounded-md transition cursor-pointer ${
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
              className="p-1 text-slate-600 hover:bg-slate-100 rounded cursor-pointer"
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
              className="p-1 text-slate-600 hover:bg-slate-100 rounded cursor-pointer"
              title="Perbesar Skala"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>

            <div className="flex items-center gap-1 ml-1 border-l border-slate-200 pl-1.5">
              {[85, 95, 100].map((preset) => (
                <button
                  key={preset}
                  type="button"
                  onClick={() => setScale(preset)}
                  className={`px-1.5 py-0.5 text-[10px] rounded cursor-pointer ${
                    scale === preset
                      ? "bg-slate-800 text-white"
                      : "text-slate-500 hover:bg-slate-100"
                  }`}
                >
                  {preset}%
                </button>
              ))}
            </div>
          </div>

          {/* Auto-Tidy Action Button */}
          <button
            type="button"
            onClick={handleAutoTidyNow}
            className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-xs transition cursor-pointer"
            title="Ekstrak ulang & rapikan otomatis seluruh tabel, kolom, font Arial, dan batas kertas sebelum unduh"
          >
            <Sparkles className="w-3.5 h-3.5" />
            <span>Rapikan Otomatis</span>
          </button>

          {/* Live Text Editing Mode Toggle */}
          <button
            type="button"
            onClick={() => setIsEditable((prev) => !prev)}
            className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer ${
              isEditable
                ? "bg-indigo-50 border-indigo-300 text-indigo-800"
                : "bg-white border-slate-200 text-slate-500 hover:bg-slate-100"
            }`}
            title="Klik teks pada kertas pratinjau untuk langsung menyunting sebelum cetak/unduh"
          >
            <Edit3 className={`w-3.5 h-3.5 ${isEditable ? "text-indigo-600" : "text-slate-400"}`} />
            <span>
              Edit: <strong>{isEditable ? "ON" : "OFF"}</strong>
            </span>
          </button>

          {/* Reset Content Button */}
          <button
            type="button"
            onClick={handleResetOriginal}
            className="px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-600 bg-white hover:bg-slate-100 transition cursor-pointer flex items-center gap-1.5"
            title="Muat ulang susunan asli dari lembar kerja"
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
            <span>
              Auto-Print: <strong>{autoPrint ? "ON" : "OFF"}</strong>
            </span>
          </button>
        </div>

        {/* Right: Precision Export & Print Actions */}
        <div className="flex flex-wrap items-center gap-2 justify-end shrink-0">
          <button
            type="button"
            onClick={handlePrintDirect}
            disabled={isProcessing}
            className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 active:scale-95 text-white text-xs font-bold rounded-xl shadow-sm transition flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
          >
            <Printer className="w-4 h-4" />
            <span>Cetak</span>
          </button>

          <button
            type="button"
            onClick={handlePDFExport}
            disabled={isProcessing}
            className="px-3.5 py-2 bg-rose-600 hover:bg-rose-700 active:scale-95 text-white text-xs font-bold rounded-xl shadow-sm transition flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
          >
            <FileDown className="w-4 h-4" />
            <span>Unduh PDF</span>
          </button>

          <button
            type="button"
            onClick={handleWordExport}
            disabled={isProcessing}
            className="px-3 py-2 bg-blue-700 hover:bg-blue-800 active:scale-95 text-white text-xs font-bold rounded-xl shadow-sm transition flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
            title="Unduh versi dokumen Word (.doc) rapi yang bisa langsung disunting"
          >
            <FileText className="w-4 h-4" />
            <span className="hidden sm:inline">Word</span>
          </button>

          <button
            type="button"
            onClick={handlePNGExport}
            disabled={isProcessing}
            className="px-3 py-2 bg-slate-800 hover:bg-slate-900 active:scale-95 text-white text-xs font-bold rounded-xl shadow-sm transition flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
          >
            <ImageIcon className="w-4 h-4" />
            <span className="hidden sm:inline">PNG</span>
          </button>

          <button
            type="button"
            onClick={onClose}
            className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition cursor-pointer ml-0.5"
            title="Tutup Pratinjau"
          >
            <X className="w-5 h-5" />
          </button>
        </div>
      </div>

      {/* Status / Success Banner */}
      {statusMsg && (
        <div className="my-2 p-2.5 bg-indigo-600 text-white text-xs font-semibold rounded-xl text-center shadow-md">
          {statusMsg}
        </div>
      )}
      {!statusMsg && tidySuccessBanner && (
        <div className="my-2 p-2.5 bg-emerald-600 text-white text-xs font-semibold rounded-xl text-center shadow-md flex items-center justify-center gap-2">
          <CheckCircle2 className="w-4 h-4 shrink-0" />
          <span>{tidySuccessBanner}</span>
        </div>
      )}

      {/* Center: Live Tidied Paper Iframe Container */}
      <div className="flex-1 my-2 bg-slate-800/50 rounded-2xl border border-slate-700 shadow-inner overflow-hidden relative">
        <iframe
          ref={iframeRef}
          title="Print Preview Frame"
          className="w-full h-full border-none rounded-2xl bg-slate-200"
        />
      </div>

      {/* Bottom Footer Info */}
      <div className="bg-white/95 backdrop-blur-md rounded-xl p-2.5 px-4 border border-slate-200 text-xs text-slate-600 flex flex-wrap items-center justify-between gap-2 shrink-0">
        <div className="flex flex-wrap items-center gap-2">
          <Maximize2 className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
          <span>
            Format:{" "}
            <strong>
              {paperSize} ({orientation === "portrait" ? "Potret" : "Lansekap"})
            </strong>{" "}
            — Font: <strong>Arial ({adaptiveFontPt}pt)</strong> — Skala: <strong>{scale}%</strong>
            {autoTidyMeta.tableCount > 0 && (
              <>
                {" "}
                —{" "}
                <strong className="text-emerald-700">
                  {autoTidyMeta.fittedCleanly ? "Muat 100% Lebar Kertas" : "Skala Otomatis Aktif"}
                </strong>
              </>
            )}
          </span>
        </div>
        <div className="flex items-center gap-2 text-[11px] text-indigo-700 bg-indigo-50 px-2.5 py-1 rounded-lg border border-indigo-100 font-semibold">
          <Sparkles className="w-3 h-3 text-indigo-600 shrink-0" />
          <span>
            Klik teks untuk sunting langsung, atau tekan <strong>Rapikan Otomatis</strong> sebelum
            mengunduh PDF / Word / PNG.
          </span>
        </div>
      </div>
    </div>
  );
};

export default PrintPreviewModal;
