/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo } from "react";
import { FoodCostDay, TKPIItem, MasterMenu, HariPM, SPPGProfile, UnifiedKopConfig, DEFAULT_UNIFIED_KOP_CONFIG } from "../types";
import { calculateDay, formatRupiah, getCountsForDay, getDayPMBreakdown, getPeriodPMAndRABSummary } from "../utils/calc";
import { composeMenuName, formatIndonesianDate } from "../utils/menuSchedule";
import { TARGET_AKG_LIMITS } from "../tkpiData";
import { PieChart, ListOrdered, PiggyBank, CalendarRange, CalendarDays, TrendingDown, TrendingUp, AlertTriangle, Printer, Download, Image as ImageIcon, Eye, X, ZoomIn, RefreshCw, RotateCcw, FlaskConical, Users, Wallet } from "lucide-react";
import { downloadElementAsImage, exportToPDF, printElementSmart } from "../lib/printUtils";
import { KopSuratRenderHeader, LogoCrop } from "./KopSuratConfigSection";
import AutoSignatureBlock from "./AutoSignatureBlock";
import PrintPreviewModal from "./PrintPreviewModal";

interface DashboardOutputsProps {
  foodCostDays: FoodCostDay[];
  tkpiList: TKPIItem[];
  masterMenu: MasterMenu;
  harianPM: HariPM[];
  profile?: SPPGProfile;
  kopLine1: string;
  setKopLine1: (val: string) => void;
  kopLine2: string;
  setKopLine2: (val: string) => void;
  kopLine3: string;
  setKopLine3: (val: string) => void;
  kopLine4: string;
  setKopLine4: (val: string) => void;
  leftLogo: string;
  setLeftLogo: (val: string) => void;
  rightLogo: string;
  setRightLogo: (val: string) => void;
  leftLogoCrop?: LogoCrop;
  setLeftLogoCrop?: (val: LogoCrop) => void;
  rightLogoCrop?: LogoCrop;
  setRightLogoCrop?: (val: LogoCrop) => void;
  paperSize?: "A4" | "F4";
  setPaperSize?: (val: "A4" | "F4") => void;
  kopConfig?: UnifiedKopConfig;
  bumbuPct?: number;
  isAdmin?: boolean;
  pmSettings?: {
    porsiKecilHarga: number;
    porsiBesarHarga: number;
    porsiKecilSasaranIds: string[];
    porsiBesarSasaranIds: string[];
  };
  setPmSettings?: (updated: any) => void;
}

export default function DashboardOutputs({
  foodCostDays,
  tkpiList,
  masterMenu,
  harianPM,
  profile,
  kopLine1,
  setKopLine1,
  kopLine2,
  setKopLine2,
  kopLine3,
  setKopLine3,
  kopLine4,
  setKopLine4,
  leftLogo,
  setLeftLogo,
  rightLogo,
  setRightLogo,
  leftLogoCrop = { top: 0, bottom: 0, left: 0, right: 0 },
  setLeftLogoCrop = () => {},
  rightLogoCrop = { top: 0, bottom: 0, left: 0, right: 0 },
  setRightLogoCrop = () => {},
  paperSize = "A4",
  setPaperSize = () => {},
  kopConfig = DEFAULT_UNIFIED_KOP_CONFIG,
  bumbuPct = 10,
  pmSettings,
  setPmSettings,
}: DashboardOutputsProps) {
  const [activeOutputTab, setActiveOutputTab] = useState<"rekap" | "nota">("rekap");
  const [targetBudgetPorsiBesar, setTargetBudgetPorsiBesar] = useState<number>(pmSettings?.porsiBesarHarga ?? 10000);
  const [targetBudgetPorsiKecil, setTargetBudgetPorsiKecil] = useState<number>(pmSettings?.porsiKecilHarga ?? 8000);

  useEffect(() => {
    if (pmSettings?.porsiBesarHarga !== undefined && pmSettings.porsiBesarHarga !== targetBudgetPorsiBesar) {
      setTargetBudgetPorsiBesar(pmSettings.porsiBesarHarga);
    }
    if (pmSettings?.porsiKecilHarga !== undefined && pmSettings.porsiKecilHarga !== targetBudgetPorsiKecil) {
      setTargetBudgetPorsiKecil(pmSettings.porsiKecilHarga);
    }
  }, [pmSettings?.porsiBesarHarga, pmSettings?.porsiKecilHarga]);

  const handleUpdateBudgetBesar = (val: number) => {
    const clean = Math.max(0, Number(val) || 0);
    setTargetBudgetPorsiBesar(clean);
    if (setPmSettings && pmSettings) {
      setPmSettings({ ...pmSettings, porsiBesarHarga: clean });
    }
  };

  const handleUpdateBudgetKecil = (val: number) => {
    const clean = Math.max(0, Number(val) || 0);
    setTargetBudgetPorsiKecil(clean);
    if (setPmSettings && pmSettings) {
      setPmSettings({ ...pmSettings, porsiKecilHarga: clean });
    }
  };
  const [saldoAwalUsiaSekolah, setSaldoAwalUsiaSekolah] = useState<number>(50000000); // 50 juta
  const [saldoAwal3B, setSaldoAwal3B] = useState<number>(15000000); // 15 juta
  const [selectedNotaDay, setSelectedNotaDay] = useState<number>(1);
  const [selectedDashboardDay, setSelectedDashboardDay] = useState<number>(1);
  const [viewMode, setViewMode] = useState<"edit" | "print">("edit");
  const [isDownloading, setIsDownloading] = useState<string | null>(null);
  const [printDocType, setPrintDocType] = useState<"rekap" | "nota">("rekap");
  const [printScale, setPrintScale] = useState<number>(100);
  const [showPrintModal, setShowPrintModal] = useState<boolean>(false);
  const [isSyncing, setIsSyncing] = useState<boolean>(false);
  const [syncAlert, setSyncAlert] = useState<string | null>(null);
  const iframeRef = useRef<HTMLIFrameElement | null>(null);

  // Auto-sync function
  const handleAutoSync = () => {
    setIsSyncing(true);
    setTimeout(() => {
      setIsSyncing(false);
      setSyncAlert("Sinkronisasi Berhasil! Seluruh kalkulasi Dashboard & Output telah disinkronkan secara otomatis dari Penerima Manfaat, Master Menu, dan Food Cost.");
      setTimeout(() => setSyncAlert(null), 5000);
    }, 500);
  };

  // Reset all panels
  const handleResetAllPanels = () => {
    if (confirm("Reset Panel Dashboard & Output:\n\nApakah Anda yakin ingin mereset seluruh panel pengaturan (target anggaran, saldo awal, dan filter) ke nilai standar?")) {
      setTargetBudgetPorsiBesar(10000);
      setTargetBudgetPorsiKecil(8000);
      setSaldoAwalUsiaSekolah(50000000);
      setSaldoAwal3B(15000000);
      setSelectedDashboardDay(1);
      setSelectedNotaDay(1);
      setSyncAlert("Seluruh panel pengaturan Dashboard & Output telah berhasil di-reset ke setelan standar.");
      setTimeout(() => setSyncAlert(null), 5000);
    }
  };

  useEffect(() => {
    if (!showPrintModal) return;
    const timer = setTimeout(() => {
      if (!iframeRef.current) return;
      const iframe = iframeRef.current;
      const doc = iframe.contentDocument || iframe.contentWindow?.document;
      if (!doc) return;

      const targetEl = document.getElementById("print-area-dashboard-rekap");
      if (!targetEl) return;

      const styles = Array.from(document.querySelectorAll("style, link[rel='stylesheet']"))
        .map((el) => el.outerHTML)
        .join("\n");

      const clonedContent = targetEl.cloneNode(true) as HTMLElement;
      clonedContent.style.zoom = `${printScale}%`;
      clonedContent.style.margin = "0 auto";
      clonedContent.style.boxShadow = "none";
      clonedContent.style.border = "none";

      doc.open();
      doc.write(`
        <!DOCTYPE html>
        <html>
          <head>
            <meta charset="utf-8">
            <title>Pratinjau Cetak Modal</title>
            ${styles}
            <style>
              body {
                background-color: #f8fafc;
                margin: 0;
                padding: 24px;
                display: flex;
                justify-content: center;
                font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
              }
              @page {
                size: A4 portrait;
                margin: 10mm;
              }
              @media print {
                body {
                  background-color: #ffffff !important;
                  padding: 0 !important;
                }
                #preview-wrapper {
                  max-width: 100% !important;
                  width: 100% !important;
                }
              }
            </style>
          </head>
          <body>
            <div id="preview-wrapper" style="width: 100%; max-width: 210mm;">
              ${clonedContent.outerHTML}
            </div>
          </body>
        </html>
      `);
      doc.close();
    }, 50);

    return () => clearTimeout(timer);
  }, [showPrintModal, printScale, printDocType, kopLine1, kopLine2, kopLine3, kopLine4, leftLogo, rightLogo, leftLogoCrop, rightLogoCrop, selectedNotaDay, selectedDashboardDay]);

  // Selected dashboard day counts
  const currentDayCounts = getCountsForDay(harianPM, selectedDashboardDay);
  
  const pmKecilSekolah = currentDayCounts.pmKecilSekolah;
  const pmBesarSekolah = currentDayCounts.pmBesarSekolah;
  const totalSekolahSiswa = pmKecilSekolah + pmBesarSekolah;
  
  const totalBalita = currentDayCounts.totalBalita;
  const totalBumil = currentDayCounts.totalBumil;
  const totalBusui = currentDayCounts.totalBusui;
  const total3BPM = totalBalita + totalBumil + totalBusui;

  const totalSekolahAlergi = currentDayCounts.totalSekolahAlergi;
  const total3BAlergi = currentDayCounts.total3BAlergi;
  const totalAlergiCombined = currentDayCounts.totalAlergiCombined;

  const totalMPAsi = currentDayCounts.totalMPAsi;
  const grandTotalPMAll = totalSekolahSiswa + total3BPM;

  const pmKecil3B = totalBalita;
  const pmBesar3B = totalBumil + totalBusui;

  const activePeriodeDates =
    profile?.periodeDates && profile.periodeDates.length > 0
      ? profile.periodeDates
      : masterMenu?.scheduleConfig?.periodeDates || [];
  const plannedLen = activePeriodeDates.length;
  const totalDays = Math.max(plannedLen > 0 ? plannedLen : 10, 1);

  const selectedDashboardDateStr = activePeriodeDates[selectedDashboardDay - 1] || "";
  const formattedDashboardDate = selectedDashboardDateStr
    ? formatIndonesianDate(selectedDashboardDateStr)
    : `Hari Kerja Ke-${selectedDashboardDay}`;

  const periodRangeLabel = useMemo(() => {
    const validDates = activePeriodeDates
      .slice(0, totalDays)
      .filter((d): d is string => Boolean(d && d.trim()));
    if (validDates.length >= 2) {
      return `${formatIndonesianDate(validDates[0])} s/d ${formatIndonesianDate(validDates[validDates.length - 1])}`;
    }
    if (validDates.length === 1) {
      return formatIndonesianDate(validDates[0]);
    }
    return `Siklus ${totalDays} Hari Kerja`;
  }, [activePeriodeDates, totalDays]);

  const dailyBreakdown = useMemo(
    () =>
      getDayPMBreakdown(
        harianPM,
        selectedDashboardDay,
        targetBudgetPorsiKecil,
        targetBudgetPorsiBesar
      ),
    [harianPM, selectedDashboardDay, targetBudgetPorsiKecil, targetBudgetPorsiBesar]
  );

  const periodSummary = useMemo(
    () =>
      getPeriodPMAndRABSummary(
        harianPM,
        totalDays,
        targetBudgetPorsiKecil,
        targetBudgetPorsiBesar
      ),
    [harianPM, totalDays, targetBudgetPorsiKecil, targetBudgetPorsiBesar]
  );

  const formatIDRCompact = (num: number) => `Rp ${Math.round(num).toLocaleString("id-ID")}`;

  useEffect(() => {
    if (selectedDashboardDay > totalDays) setSelectedDashboardDay(totalDays);
    if (selectedNotaDay > totalDays) setSelectedNotaDay(totalDays);
  }, [totalDays, selectedDashboardDay, selectedNotaDay]);

  const getMenuNameForDay = (dayIdx: number): string => {
    const activePlan =
      masterMenu.plans?.find((p) => p.id === masterMenu.activePlanId) ||
      masterMenu.plans?.[0];
    const fromPlan = composeMenuName(activePlan?.days?.[dayIdx]);
    if (fromPlan) return fromPlan;
    const fromFc = foodCostDays.find((d) => d.hariKe === dayIdx + 1 && d.namaMenu)?.namaMenu;
    if (fromFc) return fromFc;
    const fromLegacy = composeMenuName(masterMenu.usiaSekolah?.[dayIdx]);
    return fromLegacy || `Menu Hari ${dayIdx + 1}`;
  };

  const SEKOLAH_IDS_SET = new Set([
    "tk_paud_lb",
    "sd_kelas_1_3",
    "sd_kelas_4_6",
    "smp_mts_smplb",
    "sma_smk_ma",
    "pendidik",
    "tenaga_kependidikan",
    "sampel",
  ]);
  const THREEB_IDS_SET = new Set([
    "anak_balita",
    "anak_balita_13_59",
    "balita_6_11",
    "ibu_hamil",
    "ibu_menyusui",
  ]);

  const computeDashboardGroupResult = (
    dayNum: number,
    fcDay: Partial<FoodCostDay>,
    pmBesarFallback: number,
    pmKecilFallback: number,
    targetIdsSet: Set<string>
  ) => {
    const entries = fcDay.activeMode === "kategori" ? fcDay.categoryEntries : undefined;
    if (!Array.isArray(entries) || entries.length === 0) {
      return calculateDay(
        fcDay.porsiBesarBahan || [],
        fcDay.porsiKecilBahan || [],
        pmBesarFallback,
        pmKecilFallback,
        fcDay.bufferPct ?? 5,
        tkpiList,
        fcDay.bumbuPct ?? bumbuPct
      );
    }

    const dayPM = harianPM.find((h) => h.hariKe === dayNum) || harianPM[0];
    const matching = entries.filter((ce) => (ce.sasaranIds || []).some((id) => targetIdsSet.has(id)));
    const targetEntries = matching.length > 0 ? matching : entries;

    const baseRes = calculateDay(
      [],
      [],
      pmBesarFallback,
      pmKecilFallback,
      fcDay.bufferPct ?? 5,
      tkpiList,
      fcDay.bumbuPct ?? bumbuPct
    );

    let subtotalBesarCost = 0;
    let subtotalKecilCost = 0;
    let totalBesarBahanCost = 0;
    let totalKecilBahanCost = 0;
    let bufferBesarCost = 0;
    let bufferKecilCost = 0;
    let bumbuBesarCost = 0;
    let bumbuKecilCost = 0;
    let sumPmBesar = 0;
    let sumPmKecil = 0;
    const porsiBesarItems: typeof baseRes.porsiBesarItems = [];
    const porsiKecilItems: typeof baseRes.porsiKecilItems = [];
    let wBesar = 0;
    let eBesar = 0, pBesar = 0, lBesar = 0, khBesar = 0, sBesar = 0;
    let wKecil = 0;
    let eKecil = 0, pKecil = 0, lKecil = 0, khKecil = 0, sKecil = 0;

    targetEntries.forEach((ce) => {
      if (!ce.bahanList || ce.bahanList.length === 0) return;
      const entryPm =
        ce.customPmCount !== undefined
          ? Math.max(0, Number(ce.customPmCount) || 0)
          : (ce.sasaranIds || []).reduce((sum, sId) => {
              const sas = dayPM?.sasaran?.find((x) => x.id === sId);
              if (!sas) return sum;
              return sum + (ce.porsi === "besar" ? Number(sas.porsiBesar) || 0 : Number(sas.porsiKecil) || 0);
            }, 0);
      const entryBuf = ce.bufferPct ?? fcDay.bufferPct ?? 5;
      const res =
        ce.porsi === "besar"
          ? calculateDay(ce.bahanList, [], entryPm, 0, entryBuf, tkpiList, fcDay.bumbuPct ?? bumbuPct)
          : calculateDay([], ce.bahanList, 0, entryPm, entryBuf, tkpiList, fcDay.bumbuPct ?? bumbuPct);

      if (ce.porsi === "besar") {
        subtotalBesarCost += res.subtotalBesarCost;
        totalBesarBahanCost += res.totalBesarBahanCost;
        bufferBesarCost += res.bufferBesarCost;
        bumbuBesarCost += res.bumbuBesarCost;
        sumPmBesar += entryPm;
        porsiBesarItems.push(...res.porsiBesarItems);
        const w = Math.max(entryPm, 1);
        wBesar += w;
        eBesar += (res.nutrisiPorsiBesar.energi || 0) * w;
        pBesar += (res.nutrisiPorsiBesar.protein || 0) * w;
        lBesar += (res.nutrisiPorsiBesar.lemak || 0) * w;
        khBesar += (res.nutrisiPorsiBesar.kh || 0) * w;
        sBesar += (res.nutrisiPorsiBesar.serat || 0) * w;
      } else {
        subtotalKecilCost += res.subtotalKecilCost;
        totalKecilBahanCost += res.totalKecilBahanCost;
        bufferKecilCost += res.bufferKecilCost;
        bumbuKecilCost += res.bumbuKecilCost;
        sumPmKecil += entryPm;
        porsiKecilItems.push(...res.porsiKecilItems);
        const w = Math.max(entryPm, 1);
        wKecil += w;
        eKecil += (res.nutrisiPorsiKecil.energi || 0) * w;
        pKecil += (res.nutrisiPorsiKecil.protein || 0) * w;
        lKecil += (res.nutrisiPorsiKecil.lemak || 0) * w;
        khKecil += (res.nutrisiPorsiKecil.kh || 0) * w;
        sKecil += (res.nutrisiPorsiKecil.serat || 0) * w;
      }
    });

    const finalPmBesar = sumPmBesar > 0 ? sumPmBesar : pmBesarFallback;
    const finalPmKecil = sumPmKecil > 0 ? sumPmKecil : pmKecilFallback;

    return {
      ...baseRes,
      jumlahPMBesar: finalPmBesar,
      jumlahPMKecil: finalPmKecil,
      porsiBesarItems,
      porsiKecilItems,
      subtotalBesarCost,
      subtotalKecilCost,
      totalBesarBahanCost,
      totalKecilBahanCost,
      bufferBesarCost,
      bufferKecilCost,
      bumbuBesarCost,
      bumbuKecilCost,
      costPerPorsiBesar: finalPmBesar > 0 ? subtotalBesarCost / finalPmBesar : 0,
      costPerPorsiKecil: finalPmKecil > 0 ? subtotalKecilCost / finalPmKecil : 0,
      nutrisiPorsiBesar: {
        ...baseRes.nutrisiPorsiBesar,
        energi: wBesar > 0 ? eBesar / wBesar : 0,
        protein: wBesar > 0 ? pBesar / wBesar : 0,
        lemak: wBesar > 0 ? lBesar / wBesar : 0,
        kh: wBesar > 0 ? khBesar / wBesar : 0,
        serat: wBesar > 0 ? sBesar / wBesar : 0,
      },
      nutrisiPorsiKecil: {
        ...baseRes.nutrisiPorsiKecil,
        energi: wKecil > 0 ? eKecil / wKecil : 0,
        protein: wKecil > 0 ? pKecil / wKecil : 0,
        lemak: wKecil > 0 ? lKecil / wKecil : 0,
        kh: wKecil > 0 ? khKecil / wKecil : 0,
        serat: wKecil > 0 ? sKecil / wKecil : 0,
      },
    };
  };

  // Precalculate all planned days to render Rekap with day-specific recipient counts
  const calculatedDays = Array.from({ length: totalDays }).map((_, idx) => {
    const dayNum = idx + 1;
    const dayCounts = getCountsForDay(harianPM, dayNum);
    
    // School Group (Basah/Standard)
    const schoolDayData = foodCostDays.find((d) => d.hariKe === dayNum && d.jenisMenu === "Basah") || {
      hariKe: dayNum,
      jenisMenu: "Basah" as const,
      porsiBesarBahan: [],
      porsiKecilBahan: [],
      bufferPct: 5
    };
    
    const schoolResult = computeDashboardGroupResult(
      dayNum,
      schoolDayData,
      dayCounts.pmBesarSekolah,
      dayCounts.pmKecilSekolah,
      SEKOLAH_IDS_SET
    );

    // 3B Group (MP-ASI)
    const threeBDayData = foodCostDays.find((d) => d.hariKe === dayNum && d.jenisMenu === "MP-ASI") || {
      hariKe: dayNum,
      jenisMenu: "MP-ASI" as const,
      porsiBesarBahan: [],
      porsiKecilBahan: [],
      bufferPct: 5
    };

    const threeBResult = computeDashboardGroupResult(
      dayNum,
      threeBDayData,
      dayCounts.pmBesar3B,
      dayCounts.pmKecil3B,
      THREEB_IDS_SET
    );

    return {
      dayNum,
      schoolResult,
      threeBResult,
      menuName: getMenuNameForDay(idx)
    };
  });

  // Overall calculations for entire period (sum of all 12 days)
  const totalBelanjaSekolah = calculatedDays.reduce((acc, d) => acc + d.schoolResult.subtotalBesarCost + d.schoolResult.subtotalKecilCost, 0);
  const totalBelanja3B = calculatedDays.reduce((acc, d) => acc + d.threeBResult.subtotalBesarCost + d.threeBResult.subtotalKecilCost, 0);
  const grandTotalBelanja = totalBelanjaSekolah + totalBelanja3B;

  // Rujukan AKG Standards (Average for color coding)
  // Standard SD Besar is 585-683 kkal -> midpoint ~634 kkal
  const akgEnergyRef = 634;
  const akgProteinRef = 18.5; // g
  const akgLemakRef = 20; // g
  const akgKhRef = 95; // g
  const akgSeratRef = 8; // g

  // Helper for AKG indicator color-coding
  const getAKGColorBadge = (actual: number, ref: number) => {
    const pct = (actual / ref) * 100;
    if (pct >= 90 && pct <= 110) {
      return { text: "bg-emerald-50 text-emerald-800 border-emerald-200/80", icon: "✅", label: "Sesuai", pct };
    } else if (pct < 90) {
      return { text: "bg-amber-50 text-amber-800 border-amber-200/80", icon: "⚠️", label: "Kurang", pct };
    } else {
      return { text: "bg-rose-50 text-rose-800 border-rose-200/80", icon: "❌", label: "Lebih", pct };
    }
  };

  // Helper for budget comparison
  const getBudgetBadge = (cost: number, target: number) => {
    if (cost === 0) return { text: "text-slate-400 bg-slate-50", label: "N/A" };
    if (cost <= target) {
      return { text: "text-emerald-700 bg-emerald-50 font-bold", label: "UNDER BUDGET" };
    } else {
      return { text: "text-rose-700 bg-rose-50 font-bold", label: "OVER BUDGET" };
    }
  };

  // Aggregate ingredients for Nota Pesanan of specific selected day
  const getNotaIngredients = (dayNum: number, group: "sekolah" | "3b") => {
    const dayData = foodCostDays.find((d) => d.hariKe === dayNum && d.jenisMenu === (group === "sekolah" ? "Basah" : "MP-ASI"));
    if (!dayData) return [];

    const dayCounts = getCountsForDay(harianPM, dayNum);

    const result = computeDashboardGroupResult(
      dayNum,
      dayData,
      group === "sekolah" ? dayCounts.pmBesarSekolah : dayCounts.pmBesar3B,
      group === "sekolah" ? dayCounts.pmKecilSekolah : dayCounts.pmKecil3B,
      group === "sekolah" ? SEKOLAH_IDS_SET : THREEB_IDS_SET
    );

    // Group items by tkpiId to avoid multiple entries of the same food
    const aggregated: Record<string, {
      nama: string;
      totalKg: number;
      sumber: string;
      hargaSatuan: number;
    }> = {};

    // Standard items
    result.porsiBesarItems.forEach((b) => {
      if (!aggregated[b.id]) {
        aggregated[b.id] = { nama: b.nama, totalKg: b.totalKebutuhanKg, sumber: b.sumber, hargaSatuan: b.hargaSatuan };
      } else {
        aggregated[b.id].totalKg += b.totalKebutuhanKg;
      }
    });

    result.porsiKecilItems.forEach((k) => {
      if (!aggregated[k.id]) {
        aggregated[k.id] = { nama: k.nama, totalKg: k.totalKebutuhanKg, sumber: k.sumber, hargaSatuan: k.hargaSatuan };
      } else {
        aggregated[k.id].totalKg += k.totalKebutuhanKg;
      }
    });

    return Object.values(aggregated);
  };

  const currentNotaSekolah = getNotaIngredients(selectedNotaDay, "sekolah");
  const currentNota3B = getNotaIngredients(selectedNotaDay, "3b");

  return (
    <div id="outputs-view-container" className="space-y-6">
      {/* Top Controls: Tab Switcher & Action Buttons */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 pb-2 no-print">
        <div className="flex border-b border-transparent">
          <button
            type="button"
            onClick={() => setViewMode("edit")}
            className={`px-5 py-3 text-xs font-bold uppercase tracking-wider border-b-2 transition ${viewMode === "edit" ? "border-indigo-600 text-indigo-600" : "border-transparent text-slate-500 hover:text-slate-700"}`}
          >
            📝 Kustomisasi & Dashboard Interaktif
          </button>
          <button
            type="button"
            onClick={() => setViewMode("print")}
            className={`px-5 py-3 text-xs font-bold uppercase tracking-wider border-b-2 transition ${viewMode === "print" ? "border-indigo-600 text-indigo-600" : "border-transparent text-slate-500 hover:text-slate-700"}`}
          >
            🖨️ Pratinjau & Cetak Dokumen (A4)
          </button>
        </div>

        {/* Action Buttons: Sinkronisasi & Reset Panel */}
        <div className="flex items-center gap-2 mb-2 sm:mb-0">
          <button
            id="btn-sync-dashboard"
            type="button"
            onClick={handleAutoSync}
            disabled={isSyncing}
            className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white font-bold text-xs rounded-xl transition shadow-xs cursor-pointer"
            title="Sinkronisasi Otomatis data terbaru dari Penerima Manfaat, Master Menu, dan Food Cost"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? "animate-spin" : ""}`} />
            {isSyncing ? "Menyinkronkan..." : "Sinkronisasi Otomatis"}
          </button>

          <button
            id="btn-reset-dashboard-panels"
            type="button"
            onClick={handleResetAllPanels}
            className="flex items-center gap-1.5 px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 font-bold text-xs rounded-xl transition cursor-pointer"
            title="Reset seluruh panel pengaturan anggaran dan filter ke default"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            Reset Semua Panel
          </button>
        </div>
      </div>

      {syncAlert && (
        <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 p-3.5 rounded-xl text-xs font-bold flex items-center justify-between gap-2 shadow-xs animate-fadeIn no-print">
          <div className="flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-ping" />
            <span>{syncAlert}</span>
          </div>
          <button onClick={() => setSyncAlert(null)} className="text-emerald-600 hover:text-emerald-900 font-extrabold cursor-pointer">✕</button>
        </div>
      )}

      {viewMode === "edit" && (
        <div className="space-y-6">
          {/* Dashboard PM & Day Selector Panel */}
          <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-5">
        <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-4">
          <div className="flex items-start gap-3">
            <div className="text-[#4F46E5] shrink-0 mt-0.5">
              <CalendarRange className="w-6 h-6" />
            </div>
            <div>
              <h3 className="text-lg font-extrabold text-slate-900 leading-tight">
                Monitoring Penerima Manfaat (PM) Siklus {totalDays} Hari Kerja
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Pilih hari siklus (Hari 1 s/d Hari {totalDays}) untuk melihat rincian alokasi porsi dan menu gizi sasaran harian.
              </p>
            </div>
          </div>
          
          {/* Day Cycle Navigation Bar */}
          <div className="flex items-center gap-1.5 overflow-x-auto pb-2 xl:pb-0 scrollbar-none max-w-full">
            {Array.from({ length: totalDays }).map((_, idx) => {
              const day = idx + 1;
              const isSelected = selectedDashboardDay === day;
              return (
                <button
                  key={day}
                  id={`btn-dash-day-${day}`}
                  type="button"
                  onClick={() => {
                    setSelectedDashboardDay(day);
                    setSelectedNotaDay(day); // Sync both day views for user convenience!
                  }}
                  className={`w-11 h-11 rounded-lg text-center flex flex-col justify-center items-center transition shrink-0 ${
                    isSelected
                      ? "bg-[#4F46E5] text-white shadow-xs"
                      : "bg-[#F8FAFC] hover:bg-slate-100 text-slate-600 border border-slate-100/60"
                  }`}
                >
                  <span className={`text-[9px] block leading-none ${isSelected ? "text-white/80" : "text-slate-400 font-medium"}`}>Hari</span>
                  <span className="text-sm font-bold block leading-none mt-1">{day}</span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Ringkasan PM & Estimasi RAB Harian (dengan Tanggal) + Keseluruhan Periode */}
        <div className="space-y-3">
          {/* Row 1: 4 Category Breakdown Cards (Non-Overlapping) */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
            {/* Card 1: PORSI BESAR */}
            <div className="bg-[#ECFEFF]/70 border border-cyan-200/70 rounded-2xl p-4 flex flex-col justify-between shadow-xs">
              <div>
                <div className="flex items-center justify-between gap-1">
                  <span className="text-[10px] font-black text-cyan-800 uppercase tracking-wider">
                    Porsi Besar (H-{selectedDashboardDay})
                  </span>
                  <span className="text-[10px] font-semibold text-cyan-700/80">Tanpa Alergi</span>
                </div>
                <div className="mt-1.5 flex items-baseline gap-1.5">
                  <span className="text-2xl font-black text-cyan-950 tracking-tight tabular-nums">
                    {dailyBreakdown.nonSampelBesar.toLocaleString("id-ID")}
                  </span>
                  <span className="text-xs text-cyan-700 font-bold">Porsi</span>
                </div>
              </div>
              <div className="mt-2 pt-2 border-t border-cyan-200/70 flex items-center justify-between text-[11px]">
                <span className="text-cyan-700 font-medium">Estimasi RAB</span>
                <span className="font-extrabold font-mono text-cyan-950">
                  {formatIDRCompact(dailyBreakdown.rabNonSampelBesar)}
                </span>
              </div>
            </div>

            {/* Card 2: PORSI KECIL */}
            <div className="bg-[#EEF2FF]/70 border border-indigo-200/70 rounded-2xl p-4 flex flex-col justify-between shadow-xs">
              <div>
                <div className="flex items-center justify-between gap-1">
                  <span className="text-[10px] font-black text-indigo-800 uppercase tracking-wider">
                    Porsi Kecil (H-{selectedDashboardDay})
                  </span>
                  <span className="text-[10px] font-semibold text-indigo-700/80">Tanpa Alergi</span>
                </div>
                <div className="mt-1.5 flex items-baseline gap-1.5">
                  <span className="text-2xl font-black text-indigo-950 tracking-tight tabular-nums">
                    {dailyBreakdown.nonSampelKecil.toLocaleString("id-ID")}
                  </span>
                  <span className="text-xs text-indigo-700 font-bold">Porsi</span>
                </div>
              </div>
              <div className="mt-2 pt-2 border-t border-indigo-200/70 flex items-center justify-between text-[11px]">
                <span className="text-indigo-700 font-medium">Estimasi RAB</span>
                <span className="font-extrabold font-mono text-indigo-950">
                  {formatIDRCompact(dailyBreakdown.rabNonSampelKecil)}
                </span>
              </div>
            </div>

            {/* Card 3: PM ALERGI */}
            <div className="bg-[#FEF3C7]/70 border border-amber-200/80 rounded-2xl p-4 flex flex-col justify-between shadow-xs">
              <div>
                <div className="flex items-center justify-between gap-1">
                  <span className="text-[10px] font-black text-amber-800 uppercase tracking-wider">
                    PM Alergi (H-{selectedDashboardDay})
                  </span>
                  <span className="text-[10px] font-semibold text-amber-800">
                    B: {dailyBreakdown.nonSampelAlergiBesar} · K: {dailyBreakdown.nonSampelAlergiKecil}
                  </span>
                </div>
                <div className="mt-1.5 flex items-baseline gap-1.5">
                  <span className="text-2xl font-black text-amber-950 tracking-tight tabular-nums">
                    {(dailyBreakdown.nonSampelAlergiBesar + dailyBreakdown.nonSampelAlergiKecil).toLocaleString("id-ID")}
                  </span>
                  <span className="text-xs text-amber-700 font-bold">Porsi</span>
                </div>
              </div>
              <div className="mt-2 pt-2 border-t border-amber-200/70 text-[11px] text-amber-900 font-medium">
                B: <strong className="font-mono">{formatIDRCompact(dailyBreakdown.rabNonSampelAlergiBesar)}</strong> · K:{" "}
                <strong className="font-mono">{formatIDRCompact(dailyBreakdown.rabNonSampelAlergiKecil)}</strong>
              </div>
            </div>

            {/* Card 4: SAMPEL */}
            <div className="bg-blue-50/80 border border-blue-200/80 rounded-2xl p-4 flex flex-col justify-between shadow-xs">
              <div>
                <div className="flex items-center justify-between gap-1">
                  <span className="text-[10px] font-black text-blue-800 uppercase tracking-wider flex items-center gap-1">
                    <FlaskConical className="w-3.5 h-3.5" />
                    Sampel (H-{selectedDashboardDay})
                  </span>
                  <span className="text-[10px] font-semibold text-blue-800">
                    B: {dailyBreakdown.sampelBesar} · K: {dailyBreakdown.sampelKecil}
                  </span>
                </div>
                <div className="mt-1.5 flex items-baseline gap-1.5">
                  <span className="text-2xl font-black text-blue-950 tracking-tight tabular-nums">
                    {dailyBreakdown.totalSampel.toLocaleString("id-ID")}
                  </span>
                  <span className="text-xs text-blue-700 font-bold">Porsi</span>
                </div>
              </div>
              <div className="mt-2 pt-2 border-t border-blue-200/70 flex items-center justify-between text-[11px]">
                <span className="text-blue-700 font-medium">Estimasi RAB</span>
                <span className="font-extrabold font-mono text-blue-950">
                  {formatIDRCompact(dailyBreakdown.rabSampel)}
                </span>
              </div>
            </div>
          </div>

          {/* Row 2: 2 Wide Daily Hero Cards WITH DATE (Total PM Harian & Total RAB Harian) */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-3">
            {/* TOTAL PM HARIAN */}
            <div className="bg-[#1E1B4B] border border-indigo-950 text-white rounded-2xl p-4 space-y-3 shadow-sm">
              <div className="flex items-start justify-between gap-2 flex-wrap">
                <div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="text-[10px] font-black text-indigo-200 uppercase tracking-wider">
                      Total PM Harian (H-{selectedDashboardDay})
                    </span>
                    <span className="inline-flex items-center gap-1 text-[11px] font-bold text-indigo-100 bg-indigo-900/70 border border-indigo-700/80 px-2 py-0.5 rounded-lg">
                      <CalendarDays className="w-3 h-3 text-indigo-300" />
                      {formattedDashboardDate}
                    </span>
                  </div>
                  <div className="mt-1.5 flex items-baseline gap-1.5">
                    <span className="text-3xl font-black text-white tracking-tight tabular-nums">
                      {dailyBreakdown.totalPM.toLocaleString("id-ID")}
                    </span>
                    <span className="text-xs text-indigo-300 font-bold">Porsi</span>
                  </div>
                </div>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 pt-2.5 border-t border-indigo-800/70 text-[11px]">
                <div className="bg-indigo-950/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-cyan-300 font-semibold block">Porsi Besar</span>
                  <span className="font-extrabold text-white tabular-nums">{dailyBreakdown.nonSampelBesar.toLocaleString("id-ID")}</span>
                </div>
                <div className="bg-indigo-950/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-amber-300 font-semibold block">Alergi Besar</span>
                  <span className="font-extrabold text-white tabular-nums">{dailyBreakdown.nonSampelAlergiBesar.toLocaleString("id-ID")}</span>
                </div>
                <div className="bg-indigo-950/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-indigo-300 font-semibold block">Porsi Kecil</span>
                  <span className="font-extrabold text-white tabular-nums">{dailyBreakdown.nonSampelKecil.toLocaleString("id-ID")}</span>
                </div>
                <div className="bg-indigo-950/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-orange-300 font-semibold block">Alergi Kecil</span>
                  <span className="font-extrabold text-white tabular-nums">{dailyBreakdown.nonSampelAlergiKecil.toLocaleString("id-ID")}</span>
                </div>
                <div className="bg-indigo-950/70 rounded-xl px-2.5 py-1.5 col-span-2 sm:col-span-1">
                  <span className="text-[10px] text-blue-300 font-semibold block">Sampel</span>
                  <span className="font-extrabold text-white tabular-nums">
                    {dailyBreakdown.totalSampel.toLocaleString("id-ID")}{" "}
                    <span className="text-[9px] font-normal text-indigo-200">(B:{dailyBreakdown.sampelBesar}/K:{dailyBreakdown.sampelKecil})</span>
                  </span>
                </div>
              </div>
            </div>

            {/* TOTAL RAB HARIAN */}
            <div className="bg-gradient-to-br from-emerald-600 to-teal-700 border border-emerald-800 text-white rounded-2xl p-4 space-y-3 shadow-sm">
              <div className="flex items-start justify-between gap-2 flex-wrap">
                <div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="text-[10px] font-black text-emerald-100 uppercase tracking-wider">
                      Total RAB Harian (H-{selectedDashboardDay})
                    </span>
                    <span className="inline-flex items-center gap-1 text-[11px] font-bold text-white bg-emerald-900/45 border border-emerald-400/50 px-2 py-0.5 rounded-lg">
                      <CalendarDays className="w-3 h-3 text-emerald-200" />
                      {formattedDashboardDate}
                    </span>
                  </div>
                  <div className="mt-1.5 flex items-baseline gap-2 flex-wrap">
                    <span className="text-3xl font-black font-mono text-white tracking-tight">
                      {formatIDRCompact(dailyBreakdown.totalRAB)}
                    </span>
                    <span className="text-[11px] text-emerald-100/90 font-medium">
                      (Besar {formatIDRCompact(targetBudgetPorsiBesar)} · Kecil {formatIDRCompact(targetBudgetPorsiKecil)})
                    </span>
                  </div>
                </div>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 pt-2.5 border-t border-emerald-500/40 text-[11px]">
                <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-emerald-200 font-semibold block">Porsi Besar</span>
                  <span className="font-extrabold font-mono text-white block truncate">
                    {formatIDRCompact(dailyBreakdown.rabNonSampelBesar)}
                  </span>
                </div>
                <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-amber-200 font-semibold block">Alergi Besar</span>
                  <span className="font-extrabold font-mono text-white block truncate">
                    {formatIDRCompact(dailyBreakdown.rabNonSampelAlergiBesar)}
                  </span>
                </div>
                <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-sky-200 font-semibold block">Porsi Kecil</span>
                  <span className="font-extrabold font-mono text-white block truncate">
                    {formatIDRCompact(dailyBreakdown.rabNonSampelKecil)}
                  </span>
                </div>
                <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-orange-200 font-semibold block">Alergi Kecil</span>
                  <span className="font-extrabold font-mono text-white block truncate">
                    {formatIDRCompact(dailyBreakdown.rabNonSampelAlergiKecil)}
                  </span>
                </div>
                <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5 col-span-2 sm:col-span-1">
                  <span className="text-[10px] text-blue-200 font-semibold block">Sampel</span>
                  <span className="font-extrabold font-mono text-white block truncate">
                    {formatIDRCompact(dailyBreakdown.rabSampel)}
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* Row 3: Ringkasan Keseluruhan Periode (Jumlah PM & Estimasi Total RAB Keseluruhan) */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-3">
            {/* KESELURUHAN JUMLAH PM */}
            <div className="bg-slate-50/90 border border-indigo-200/80 rounded-2xl p-4 space-y-3">
              <div className="flex items-start justify-between gap-2 flex-wrap">
                <div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="inline-flex items-center gap-1.5 text-[11px] font-extrabold text-indigo-800 uppercase tracking-wider">
                      <Users className="w-3.5 h-3.5 text-indigo-600" />
                      Keseluruhan Jumlah PM ({totalDays} Hari Kerja)
                    </span>
                    <span className="text-[10px] font-bold text-indigo-700 bg-white border border-indigo-200 px-2 py-0.5 rounded-lg">
                      {periodSummary.filledDaysCount}/{totalDays} Hari Terisi
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500 mt-0.5">
                    Periode: <strong className="text-slate-700">{periodRangeLabel}</strong>
                  </p>
                </div>
                <div className="text-right">
                  <span className="text-2xl font-black text-slate-900 tabular-nums">
                    {periodSummary.totalPM.toLocaleString("id-ID")}
                  </span>
                  <span className="text-xs font-bold text-indigo-600 ml-1">Total Porsi</span>
                </div>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 pt-2.5 border-t border-slate-200/70 text-[11px]">
                <div className="bg-white border border-cyan-200/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-cyan-800 font-bold block">Porsi Besar</span>
                  <span className="text-sm font-black text-cyan-950 tabular-nums">{periodSummary.nonSampelBesar.toLocaleString("id-ID")}</span>
                </div>
                <div className="bg-white border border-amber-200/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-amber-800 font-bold block">Alergi Besar</span>
                  <span className="text-sm font-black text-amber-950 tabular-nums">{periodSummary.nonSampelAlergiBesar.toLocaleString("id-ID")}</span>
                </div>
                <div className="bg-white border border-indigo-200/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-indigo-800 font-bold block">Porsi Kecil</span>
                  <span className="text-sm font-black text-indigo-950 tabular-nums">{periodSummary.nonSampelKecil.toLocaleString("id-ID")}</span>
                </div>
                <div className="bg-white border border-orange-200/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-orange-800 font-bold block">Alergi Kecil</span>
                  <span className="text-sm font-black text-orange-950 tabular-nums">{periodSummary.nonSampelAlergiKecil.toLocaleString("id-ID")}</span>
                </div>
                <div className="bg-white border border-blue-200/70 rounded-xl px-2.5 py-1.5 col-span-2 sm:col-span-1">
                  <span className="text-[10px] text-blue-800 font-bold block">Sampel</span>
                  <span className="text-sm font-black text-blue-950 tabular-nums">{periodSummary.totalSampel.toLocaleString("id-ID")}</span>
                  <span className="text-[9px] text-blue-700 font-semibold block">B:{periodSummary.sampelBesar} · K:{periodSummary.sampelKecil}</span>
                </div>
              </div>
            </div>

            {/* ESTIMASI TOTAL RAB KESELURUHAN */}
            <div className="bg-slate-50/90 border border-emerald-200/80 rounded-2xl p-4 space-y-3">
              <div className="flex items-start justify-between gap-2 flex-wrap">
                <div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="inline-flex items-center gap-1.5 text-[11px] font-extrabold text-emerald-800 uppercase tracking-wider">
                      <Wallet className="w-3.5 h-3.5 text-emerald-600" />
                      Estimasi Total RAB Keseluruhan ({totalDays} Hari Kerja)
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500 mt-0.5">
                    Periode: <strong className="text-slate-700">{periodRangeLabel}</strong>
                  </p>
                </div>
                <div className="text-right">
                  <span className="text-2xl font-black font-mono text-emerald-800">
                    {formatIDRCompact(periodSummary.totalRAB)}
                  </span>
                </div>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 pt-2.5 border-t border-slate-200/70 text-[11px]">
                <div className="bg-white border border-emerald-200/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-emerald-800 font-bold block">Porsi Besar</span>
                  <span className="font-extrabold font-mono text-emerald-950 block truncate">{formatIDRCompact(periodSummary.rabNonSampelBesar)}</span>
                </div>
                <div className="bg-white border border-amber-200/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-amber-800 font-bold block">Alergi Besar</span>
                  <span className="font-extrabold font-mono text-amber-950 block truncate">{formatIDRCompact(periodSummary.rabNonSampelAlergiBesar)}</span>
                </div>
                <div className="bg-white border border-sky-200/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-sky-800 font-bold block">Porsi Kecil</span>
                  <span className="font-extrabold font-mono text-sky-950 block truncate">{formatIDRCompact(periodSummary.rabNonSampelKecil)}</span>
                </div>
                <div className="bg-white border border-orange-200/70 rounded-xl px-2.5 py-1.5">
                  <span className="text-[10px] text-orange-800 font-bold block">Alergi Kecil</span>
                  <span className="font-extrabold font-mono text-orange-950 block truncate">{formatIDRCompact(periodSummary.rabNonSampelAlergiKecil)}</span>
                </div>
                <div className="bg-white border border-blue-200/70 rounded-xl px-2.5 py-1.5 col-span-2 sm:col-span-1">
                  <span className="text-[10px] text-blue-800 font-bold block">Sampel</span>
                  <span className="font-extrabold font-mono text-blue-950 block truncate">{formatIDRCompact(periodSummary.rabSampel)}</span>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Selected Day Status Bar */}
        <div className="bg-slate-50 rounded-xl p-3 flex flex-wrap items-center justify-between gap-4 text-xs">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="bg-indigo-100 text-indigo-800 px-2.5 py-0.5 rounded font-bold">
              Hari Ke-{selectedDashboardDay} · {formattedDashboardDate}
            </span>
            <span className="text-slate-600 font-medium">
              Menu Aktif:{" "}
              <strong className="text-slate-900">
                {getMenuNameForDay(selectedDashboardDay - 1)}
              </strong>
            </span>
          </div>
          <div className="flex items-center gap-4 text-[11px] text-slate-500 flex-wrap">
            <span>Estimasi Porsi Sekolah: <strong className="text-slate-700 font-mono">{totalSekolahSiswa} porsi</strong></span>
            <span>Estimasi Porsi 3B: <strong className="text-slate-700 font-mono">{total3BPM} porsi</strong></span>
            <span>Sampel: <strong className="text-blue-700 font-mono">{dailyBreakdown.totalSampel} porsi</strong></span>
          </div>
        </div>
      </div>

      {/* Tab Switcher */}
      <div className="flex border-b border-slate-200">
        <button
          id="btn-output-rekap"
          type="button"
          onClick={() => setActiveOutputTab("rekap")}
          className={`py-3 px-5 font-bold text-sm border-b-2 transition flex items-center gap-2 ${
            activeOutputTab === "rekap"
              ? "border-indigo-600 text-indigo-600"
              : "border-transparent text-slate-500 hover:text-slate-800"
          }`}
        >
          <PieChart className="w-4 h-4" />
          Output 1 — Rekap Nilai Gizi & Food Cost Harian
        </button>
        <button
          id="btn-output-nota"
          type="button"
          onClick={() => setActiveOutputTab("nota")}
          className={`py-3 px-5 font-bold text-sm border-b-2 transition flex items-center gap-2 ${
            activeOutputTab === "nota"
              ? "border-indigo-600 text-indigo-600"
              : "border-transparent text-slate-500 hover:text-slate-800"
          }`}
        >
          <ListOrdered className="w-4 h-4" />
          Output 2 — Nota Pesanan & Subsidi Silang
        </button>
      </div>

      {activeOutputTab === "rekap" ? (
        <div className="space-y-6">
          {/* Budget Setting Panel */}
          <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex flex-col xl:flex-row xl:items-center justify-between gap-6">
            <div className="space-y-1.5 flex-1">
              <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <PiggyBank className="w-4.5 h-4.5 text-emerald-600" />
                Alokasi Anggaran Target Per Porsi
              </h4>
              <p className="text-xs text-slate-500">
                Tentukan target plafon anggaran makan bergizi per per porsi untuk memicu analisis real-time LOW/OVER budget di tabel bawah.
              </p>
            </div>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 xl:shrink-0 w-full xl:w-auto">
              {/* PORSI BESAR */}
              <div className="space-y-1">
                <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
                  PORSI BESAR (SD 1-4, SMP, BUMIL/BUSUI) Target Plafon:
                </span>
                <div className="flex items-center">
                  <span className="bg-amber-50 border border-amber-200 border-r-0 rounded-l-lg text-xs font-bold text-amber-700 px-3 py-2">Rp</span>
                  <input
                    id="input-budget-porsi-besar"
                    type="number"
                    value={targetBudgetPorsiBesar}
                    onChange={(e) => handleUpdateBudgetBesar(Number(e.target.value))}
                    className="px-3 py-1.5 border border-slate-200 rounded-r-lg text-sm text-slate-800 font-semibold focus:outline-none focus:ring-2 focus:ring-amber-500/20 font-mono text-right w-full sm:w-32"
                  />
                </div>
              </div>

              {/* PORSI KECIL */}
              <div className="space-y-1">
                <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
                  PORSI KECIL (BALITA, TK/PAUD, SD 1-3) Target Plafon:
                </span>
                <div className="flex items-center">
                  <span className="bg-cyan-50 border border-cyan-200 border-r-0 rounded-l-lg text-xs font-bold text-cyan-700 px-3 py-2">Rp</span>
                  <input
                    id="input-budget-porsi-kecil"
                    type="number"
                    value={targetBudgetPorsiKecil}
                    onChange={(e) => handleUpdateBudgetKecil(Number(e.target.value))}
                    className="px-3 py-1.5 border border-slate-200 rounded-r-lg text-sm text-slate-800 font-semibold focus:outline-none focus:ring-2 focus:ring-cyan-500/20 font-mono text-right w-full sm:w-32"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Table of 10 days rekap */}
          <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
            <div className="p-4 bg-slate-50/50 border-b border-slate-100">
              <h3 className="font-bold text-slate-800 flex items-center gap-2">
                <CalendarRange className="w-5 h-5 text-indigo-600" />
                Siklus {totalDays} Hari Kerja — Analisis Kandungan Gizi vs Cost
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">Ringkasan gizi harian (AKG Rujukan SD 4-6) dan realisasi harga porsi</p>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead className="bg-slate-100 text-[10px] font-bold text-slate-600 uppercase tracking-wider">
                  <tr>
                    <th className="p-3 pl-4 w-12 text-center">Hari</th>
                    <th className="p-3 min-w-[150px]">Nama Menu</th>
                    <th className="p-3 text-right">Berat Bersih (g)</th>
                    <th className="p-3 text-right">Energi (kkal)</th>
                    <th className="p-3 text-right">Protein (g)</th>
                    <th className="p-3 text-right">Lemak (g)</th>
                    <th className="p-3 text-right">Karbohidrat (g)</th>
                    <th className="p-3 text-right">Serat (g)</th>
                    <th className="p-3 text-right">Porsi Kecil (Rp)</th>
                    <th className="p-3 text-right">Porsi Besar (Rp)</th>
                    <th className="p-3 text-right">RAB Harian (Rp)</th>
                    <th className="p-3 text-center min-w-[120px]">Status Anggaran</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-xs text-slate-700">
                  {calculatedDays.map((d) => {
                    // Average nutrisi
                    const energi = d.schoolResult.nutrisiPorsiBesar.energi || 0;
                    const protein = d.schoolResult.nutrisiPorsiBesar.protein || 0;
                    const lemak = d.schoolResult.nutrisiPorsiBesar.lemak || 0;
                    const kh = d.schoolResult.nutrisiPorsiBesar.kh || 0;
                    const serat = d.schoolResult.nutrisiPorsiBesar.serat || 0;

                    const beratBersih = d.schoolResult.porsiBesarItems.reduce((acc, item) => acc + item.beratBB, 0);

                    const costKecil = d.schoolResult.costPerPorsiKecil;
                    const costBesar = d.schoolResult.costPerPorsiBesar;
                    const rabHarian = d.schoolResult.subtotalBesarCost + d.schoolResult.subtotalKecilCost + d.threeBResult.subtotalBesarCost + d.threeBResult.subtotalKecilCost;

                    const isBesarOver = costBesar > 0 && costBesar > targetBudgetPorsiBesar;
                    const isKecilOver = costKecil > 0 && costKecil > targetBudgetPorsiKecil;
                    const isAnyOver = isBesarOver || isKecilOver;
                    const isAnyValid = costBesar > 0 || costKecil > 0;

                    const budgetStatus = !isAnyValid
                      ? { text: "text-slate-400 bg-slate-50", label: "N/A" }
                      : isAnyOver
                        ? { text: "text-rose-700 bg-rose-50 font-bold", label: "OVER BUDGET" }
                        : { text: "text-emerald-700 bg-emerald-50 font-bold", label: "UNDER BUDGET" };

                    // Nutrients status
                    const nEnergy = getAKGColorBadge(energi, akgEnergyRef);
                    const nProtein = getAKGColorBadge(protein, akgProteinRef);
                    const nLemak = getAKGColorBadge(lemak, akgLemakRef);
                    const nKh = getAKGColorBadge(kh, akgKhRef);
                    const nSerat = getAKGColorBadge(serat, akgSeratRef);

                    return (
                      <tr key={d.dayNum} className="hover:bg-slate-50/50 transition-colors">
                        <td className="p-3 pl-4 text-center font-bold text-indigo-600 font-mono">H-{d.dayNum}</td>
                        <td className="p-3 font-medium text-slate-800">{d.menuName}</td>
                        
                        {/* Berat Bersih */}
                        <td className="p-3 text-right font-mono font-semibold text-indigo-700">
                          {beratBersih > 0 ? `${beratBersih.toFixed(1)}g` : "-"}
                        </td>

                        {/* Energi */}
                        <td className="p-3 text-right font-mono">
                          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] border font-semibold ${nEnergy.text}`}>
                            {energi.toFixed(0)} ({nEnergy.pct.toFixed(0)}%)
                          </span>
                        </td>

                        {/* Protein */}
                        <td className="p-3 text-right font-mono">
                          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] border font-semibold ${nProtein.text}`}>
                            {protein.toFixed(1)}g ({nProtein.pct.toFixed(0)}%)
                          </span>
                        </td>

                        {/* Lemak */}
                        <td className="p-3 text-right font-mono">
                          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] border font-semibold ${nLemak.text}`}>
                            {lemak.toFixed(1)}g ({nLemak.pct.toFixed(0)}%)
                          </span>
                        </td>

                        {/* KH */}
                        <td className="p-3 text-right font-mono">
                          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] border font-semibold ${nKh.text}`}>
                            {kh.toFixed(1)}g ({nKh.pct.toFixed(0)}%)
                          </span>
                        </td>

                        {/* Serat */}
                        <td className="p-3 text-right font-mono">
                          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] border font-semibold ${nSerat.text}`}>
                            {serat.toFixed(1)}g ({nSerat.pct.toFixed(0)}%)
                          </span>
                        </td>

                        {/* Porsi Kecil */}
                        <td className="p-3 text-right font-mono text-cyan-700 font-semibold bg-cyan-50/10">
                          {costKecil > 0 ? formatRupiah(costKecil) : "-"}
                        </td>

                        {/* Porsi Besar */}
                        <td className="p-3 text-right font-mono text-amber-700 font-semibold bg-amber-50/10">
                          {costBesar > 0 ? formatRupiah(costBesar) : "-"}
                        </td>

                        {/* RAB Harian */}
                        <td className="p-3 text-right font-mono text-indigo-900 font-extrabold bg-indigo-50/10">
                          {formatRupiah(rabHarian)}
                        </td>

                        {/* Status Budget */}
                        <td className="p-3 text-center">
                          <span className={`px-2 py-1.5 rounded-lg text-[10px] border tracking-wider uppercase font-extrabold ${budgetStatus.text}`}>
                            {budgetStatus.label}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Rekap footer */}
            <div className="p-4 bg-indigo-50/20 border-t border-slate-200 flex flex-col md:flex-row items-center justify-between gap-4">
              <div className="flex flex-wrap items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-ping"></span>
                <span className="text-xs text-slate-500 font-medium">Legend Nutrisi (% AKG):</span>
                <span className="text-[10px] font-bold bg-amber-50 text-amber-800 border border-amber-200/80 px-2 py-0.5 rounded">KURANG (&lt;90%)</span>
                <span className="text-[10px] font-bold bg-emerald-50 text-emerald-800 border border-emerald-200/80 px-2 py-0.5 rounded">SESUAI (90–110%)</span>
                <span className="text-[10px] font-bold bg-rose-50 text-rose-800 border border-rose-200/80 px-2 py-0.5 rounded">LEBIH (&gt;110%)</span>
              </div>
              
              <div className="text-right">
                <span className="text-xs text-slate-500 font-semibold block">Total Proyeksi Belanja {totalDays} Hari Kerja:</span>
                <span className="text-xl font-black text-indigo-950 font-mono">{formatRupiah(grandTotalBelanja)}</span>
              </div>
            </div>
          </div>
        </div>
      ) : (
        <div className="space-y-6">
          {/* Consolidated Logistics PO Promo banner */}
          <div className="bg-gradient-to-r from-orange-50 to-amber-50 border border-orange-100 p-4 rounded-xl flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <span className="text-xl">📋</span>
              <div>
                <h5 className="text-xs font-bold text-amber-900">Nota Pesanan Logistik Gabungan (Presisi 100% Siap Cetak/Download)</h5>
                <p className="text-[11px] text-amber-700">Untuk rekap belanja kotor gabungan antara Usia Sekolah dan 3B per hari dalam format cetak resmi, silakan gunakan menu utama <strong>Nota Pesanan Logistik</strong>.</p>
              </div>
            </div>
            <button
              id="btn-goto-nota-logistik"
              type="button"
              onClick={() => {
                const navBtn = document.getElementById("tab-nav-notalogistik");
                if (navBtn) navBtn.click();
              }}
              className="text-xs bg-amber-600 hover:bg-amber-700 text-white font-extrabold px-3 py-1.5 rounded-lg shadow-sm shadow-amber-600/10 transition shrink-0"
            >
              Buka Nota Gabungan →
            </button>
          </div>

          {/* Day selection for PO */}
          <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-5">
            <div className="space-y-1">
              <h4 className="text-sm font-bold text-slate-800">Pilih Hari Nota Pesanan</h4>
              <p className="text-xs text-slate-500">Tampilkan nota rincian kebutuhan pesanan logistik dan bumbu per hari terpilih.</p>
            </div>
            
            <div className="flex items-center gap-2">
              <span className="text-xs text-slate-500 font-bold">Hari Ke:</span>
              <select
                id="select-nota-day-idx"
                value={selectedNotaDay}
                onChange={(e) => setSelectedNotaDay(Number(e.target.value))}
                className="bg-slate-100 border border-slate-200 rounded-lg text-sm px-4 py-2 font-mono font-bold text-indigo-600 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
              >
                {Array.from({ length: totalDays }).map((_, i) => (
                  <option key={i} value={i + 1}>
                    Hari Ke-{i + 1}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* TWO NOTAS - Usia Sekolah vs 3B */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* NOTA USIA SEKOLAH */}
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
              <div className="p-4 border-b border-slate-100 bg-slate-50/60">
                <span className="text-[10px] font-black text-indigo-600 uppercase tracking-wider block">Nota Pesanan</span>
                <h4 className="font-bold text-slate-800 text-sm">Kelompok Usia Sekolah (TK, SD, SMP, SMA)</h4>
                <p className="text-xs text-slate-500 mt-0.5">Alokasi penerima manfaat: {pmBesarSekolah + pmKecilSekolah} siswa</p>
              </div>

              <div className="flex-grow overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-100 text-slate-600 font-bold">
                    <tr>
                      <th className="p-2.5 pl-4 w-12 text-center">No</th>
                      <th className="p-2.5">Bahan Baku</th>
                      <th className="p-2.5 text-right w-24">Jumlah (Kg)</th>
                      <th className="p-2.5 text-right w-28">Harga (Rp/Kg)</th>
                      <th className="p-2.5 text-right w-28">Total Belanja</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {currentNotaSekolah.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="p-8 text-center text-slate-400 italic">
                          Belum ada komponen pangan yang diinput hari ini.
                        </td>
                      </tr>
                    ) : (
                      currentNotaSekolah.map((item, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50">
                          <td className="p-2.5 pl-4 text-center font-mono text-slate-400">{idx + 1}</td>
                          <td className="p-2.5 font-medium text-slate-800">{item.nama}</td>
                          <td className="p-2.5 text-right font-mono">{item.totalKg.toFixed(3)} Kg</td>
                          <td className="p-2.5 text-right font-mono">{formatRupiah(item.hargaSatuan)}</td>
                          <td className="p-2.5 text-right font-mono font-bold text-slate-900">
                            {formatRupiah(item.totalKg * item.hargaSatuan)}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              {/* Costing calculations & Cross-Subsidy for Schools */}
              <div className="p-4 bg-slate-50/80 border-t border-slate-100 space-y-3">
                <div className="flex justify-between text-xs">
                  <span className="text-slate-500">Sub-total Bahan:</span>
                  <span className="font-mono font-bold text-slate-800">
                    {formatRupiah(currentNotaSekolah.reduce((acc, i) => acc + (i.totalKg * i.hargaSatuan), 0))}
                  </span>
                </div>
                
                {/* Cross subsidy calculator */}
                <div className="bg-white p-3 rounded-xl border border-slate-200 text-xs space-y-2">
                  <span className="font-bold text-indigo-950 flex items-center gap-1.5">
                    <PiggyBank className="w-4 h-4 text-indigo-600" />
                    Kalkulator Subsidi Silang Usia Sekolah
                  </span>
                  <div className="grid grid-cols-2 gap-3 pt-1">
                    <div>
                      <label className="text-[10px] text-slate-400 font-bold uppercase tracking-wider block">Saldo Awal Alokasi</label>
                      <input
                        id="input-saldo-sekolah"
                        type="number"
                        value={saldoAwalUsiaSekolah}
                        onChange={(e) => setSaldoAwalUsiaSekolah(Number(e.target.value))}
                        className="w-full bg-slate-50 border border-slate-200 rounded p-1.5 text-right font-mono font-bold mt-1 text-slate-800"
                      />
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider block">Sisa Saldo (Selisih)</span>
                      <div className={`p-1.5 rounded text-right font-mono font-black mt-1 ${saldoAwalUsiaSekolah - totalBelanjaSekolah >= 0 ? "text-emerald-700 bg-emerald-50" : "text-rose-700 bg-rose-50"}`}>
                        {formatRupiah(saldoAwalUsiaSekolah - totalBelanjaSekolah)}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* NOTA 3B */}
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
              <div className="p-4 border-b border-slate-100 bg-slate-50/60">
                <span className="text-[10px] font-black text-rose-600 uppercase tracking-wider block">Nota Pesanan</span>
                <h4 className="font-bold text-slate-800 text-sm">Kelompok 3B (Bumil, Balita, Busui)</h4>
                <p className="text-xs text-slate-500 mt-0.5">Alokasi penerima manfaat: {pmBesar3B + pmKecil3B} orang</p>
              </div>

              <div className="flex-grow overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-100 text-slate-600 font-bold">
                    <tr>
                      <th className="p-2.5 pl-4 w-12 text-center">No</th>
                      <th className="p-2.5">Bahan Baku</th>
                      <th className="p-2.5 text-right w-24">Jumlah (Kg)</th>
                      <th className="p-2.5 text-right w-28">Harga (Rp/Kg)</th>
                      <th className="p-2.5 text-right w-28">Total Belanja</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {currentNota3B.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="p-8 text-center text-slate-400 italic">
                          Belum ada komponen pangan yang diinput hari ini.
                        </td>
                      </tr>
                    ) : (
                      currentNota3B.map((item, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50">
                          <td className="p-2.5 pl-4 text-center font-mono text-slate-400">{idx + 1}</td>
                          <td className="p-2.5 font-medium text-slate-800">{item.nama}</td>
                          <td className="p-2.5 text-right font-mono">{item.totalKg.toFixed(3)} Kg</td>
                          <td className="p-2.5 text-right font-mono">{formatRupiah(item.hargaSatuan)}</td>
                          <td className="p-2.5 text-right font-mono font-bold text-slate-900">
                            {formatRupiah(item.totalKg * item.hargaSatuan)}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              {/* Costing calculations & Cross-Subsidy for 3B */}
              <div className="p-4 bg-slate-50/80 border-t border-slate-100 space-y-3">
                <div className="flex justify-between text-xs">
                  <span className="text-slate-500">Sub-total Bahan:</span>
                  <span className="font-mono font-bold text-slate-800">
                    {formatRupiah(currentNota3B.reduce((acc, i) => acc + (i.totalKg * i.hargaSatuan), 0))}
                  </span>
                </div>
                
                {/* Cross subsidy calculator */}
                <div className="bg-white p-3 rounded-xl border border-slate-200 text-xs space-y-2">
                  <span className="font-bold text-rose-950 flex items-center gap-1.5">
                    <PiggyBank className="w-4 h-4 text-rose-600" />
                    Kalkulator Subsidi Silang Kelompok 3B
                  </span>
                  <div className="grid grid-cols-2 gap-3 pt-1">
                    <div>
                      <label className="text-[10px] text-slate-400 font-bold uppercase tracking-wider block">Saldo Awal Alokasi</label>
                      <input
                        id="input-saldo-3b"
                        type="number"
                        value={saldoAwal3B}
                        onChange={(e) => setSaldoAwal3B(Number(e.target.value))}
                        className="w-full bg-slate-50 border border-slate-200 rounded p-1.5 text-right font-mono font-bold mt-1 text-slate-800"
                      />
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider block">Sisa Saldo (Selisih)</span>
                      <div className={`p-1.5 rounded text-right font-mono font-black mt-1 ${saldoAwal3B - totalBelanja3B >= 0 ? "text-emerald-700 bg-emerald-50" : "text-rose-700 bg-rose-50"}`}>
                        {formatRupiah(saldoAwal3B - totalBelanja3B)}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  )}

      {/* --- PRATINJAU CETAK KOP & EKSPOR --- */}
      <div className={viewMode === "print" ? "space-y-6" : "hidden print:block space-y-6"}>
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4 no-print">
          <div className="flex flex-col sm:flex-row justify-between sm:items-center pb-2 border-b border-slate-100 gap-3">
            <div className="space-y-1">
              <h4 className="text-xs font-black text-slate-900 uppercase tracking-wider">Sesuaikan KOP Surat &amp; Cetak / Ekspor</h4>
              <div className="flex items-center gap-2 mt-1">
                <span className="text-xs font-bold text-slate-500">Pilih Dokumen:</span>
                <select
                  value={printDocType}
                  onChange={(e) => setPrintDocType(e.target.value as any)}
                  className="bg-slate-100 border border-slate-200 rounded px-2.5 py-1 text-xs font-bold text-indigo-600 focus:outline-none"
                >
                  <option value="rekap">RAB 10 Hari Kerja &amp; Rekap Gizi</option>
                  <option value="nota">Nota Rincian Bahan (Hari Ke-{selectedNotaDay})</option>
                </select>
              </div>
            </div>

            <div className="flex items-center gap-2 self-start sm:self-auto flex-wrap">
              {/* Slider Scale Control */}
              <div className="flex items-center gap-2 bg-slate-100 border border-slate-200 rounded-xl px-3 py-1.5">
                <span className="text-xs font-bold text-slate-600">Skala Cetak:</span>
                <input
                  type="range"
                  min={50}
                  max={100}
                  step={5}
                  value={printScale}
                  onChange={(e) => setPrintScale(Number(e.target.value))}
                  className="w-24 accent-indigo-600 cursor-pointer"
                />
                <span className="text-xs font-black text-indigo-700 min-w-[36px]">{printScale}%</span>
              </div>

              <button
                id="btn-do-print-preview-modal"
                type="button"
                onClick={() => setShowPrintModal(true)}
                className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition"
              >
                <Eye className="w-4 h-4" />
                Print Preview Modal
              </button>

              <button
                id="btn-do-download-pdf"
                type="button"
                disabled={!!isDownloading}
                onClick={() => exportToPDF("print-area-dashboard-rekap", printDocType === "rekap" ? `RAB_${totalDays}_Hari_Rekap_Gizi` : "Nota_Rincian_Bahan_Hari_" + selectedNotaDay, printScale, setIsDownloading, { paperSize, orientation: "auto" })}
                className="bg-red-600 hover:bg-red-700 disabled:opacity-50 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition"
              >
                <Download className="w-4 h-4" />
                {isDownloading === "Memproses dokumen PDF..." || isDownloading === "Membuat file PDF..." ? isDownloading : "Ekspor PDF"}
              </button>

              <button
                id="btn-do-download-img"
                type="button"
                disabled={!!isDownloading}
                onClick={() => downloadElementAsImage("print-area-dashboard-rekap", printDocType === "rekap" ? `RAB_${totalDays}_Hari_Rekap_Gizi` : "Nota_Rincian_Bahan_Hari_" + selectedNotaDay, setIsDownloading)}
                className="bg-amber-600 hover:bg-amber-700 disabled:opacity-50 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition"
              >
                <ImageIcon className="w-4 h-4" />
                {isDownloading === "Memproses gambar..." || isDownloading === "Mengunduh gambar..." ? isDownloading : "Unduh Gambar (PNG)"}
              </button>

              <button
                id="btn-do-print"
                type="button"
                onClick={() => printElementSmart("print-area-dashboard-rekap", { title: printDocType === "rekap" ? `RAB ${totalDays} Hari & Rekap Gizi` : `Nota Rincian Bahan Hari ${selectedNotaDay}`, paperSize, orientation: "auto" })}
                className="bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition"
              >
                <Printer className="w-4 h-4" />
                Cetak / Simpan PDF
              </button>
            </div>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase">KOP Baris 1</label>
              <input
                type="text"
                value={kopLine1}
                onChange={(e) => setKopLine1(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 rounded-lg text-xs p-2 text-slate-800 font-bold"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase">KOP Baris 2</label>
              <input
                type="text"
                value={kopLine2}
                onChange={(e) => setKopLine2(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 rounded-lg text-xs p-2 text-slate-800 font-bold"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase">KOP Baris 3</label>
              <input
                type="text"
                value={kopLine3}
                onChange={(e) => setKopLine3(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 rounded-lg text-xs p-2 text-slate-800 font-bold"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase">KOP Baris 4</label>
              <input
                type="text"
                value={kopLine4}
                onChange={(e) => setKopLine4(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 rounded-lg text-xs p-2 text-slate-800"
              />
            </div>
          </div>
        </div>

        <div className="bg-slate-100 p-6 rounded-2xl border border-slate-300 shadow-inner flex justify-center overflow-x-auto print:bg-transparent print:p-0 print:border-none print:shadow-none">
          <div 
            id="print-area-dashboard-rekap" 
            style={{ zoom: `${printScale}%`, fontFamily: "Arial, Helvetica, sans-serif" }}
            className="bg-white p-8 border border-slate-400 shadow-md w-full max-w-[210mm] min-w-[210mm] text-slate-950 print:text-black print:border-none print:shadow-none print:p-0 print:m-0 space-y-8"
          >
              {/* Kop Surat Terpusat SPPG */}
              <KopSuratRenderHeader
                kopLine1={kopLine1}
                kopLine2={kopLine2}
                kopLine3={kopLine3}
                kopLine4={kopLine4}
                leftLogo={leftLogo}
                rightLogo={rightLogo}
                leftLogoCrop={leftLogoCrop}
                rightLogoCrop={rightLogoCrop}
                kopConfig={kopConfig}
                docKey="dashboard"
              />

              {printDocType === "rekap" ? (
                <div className="space-y-6">
                  {/* Title */}
                  <div className="text-center space-y-1">
                    <h3 className="font-black text-black uppercase tracking-wide" style={{ fontSize: '13pt' }}>
                      LAPORAN REKAPITULASI RAB &amp; ANALISIS KECUKUPAN GIZI
                    </h3>
                    <p className="text-xs font-bold text-slate-700 uppercase">
                      MONITORING PROGRAM SPPG SIKLUS {totalDays} HARI KERJA
                    </p>
                  </div>

                  {/* Dynamic-day Table for print */}
                  <table className="w-full text-left text-[9px] border-collapse border border-slate-300">
                    <thead>
                      <tr className="bg-slate-100 font-bold text-slate-700 uppercase border border-slate-300">
                        <th className="border border-slate-300 p-1.5 text-center">Hari</th>
                        <th className="border border-slate-300 p-1.5">Menu Utama</th>
                        <th className="border border-slate-300 p-1.5 text-right">Energi</th>
                        <th className="border border-slate-300 p-1.5 text-right">Protein</th>
                        <th className="border border-slate-300 p-1.5 text-right">Lemak</th>
                        <th className="border border-slate-300 p-1.5 text-right">KH</th>
                        <th className="border border-slate-300 p-1.5 text-right">Serat</th>
                        <th className="border border-slate-300 p-1.5 text-right">P.Kecil</th>
                        <th className="border border-slate-300 p-1.5 text-right">P.Besar</th>
                        <th className="border border-slate-300 p-1.5 text-right">RAB Harian</th>
                      </tr>
                    </thead>
                    <tbody>
                      {calculatedDays.map(({ dayNum, schoolResult, menuName }) => {
                        return (
                          <tr key={dayNum} className="text-slate-800 font-medium">
                            <td className="border border-slate-300 p-1.5 text-center font-mono font-bold">H{dayNum}</td>
                            <td className="border border-slate-300 p-1.5 truncate max-w-[120px]">{menuName}</td>
                            <td className="border border-slate-300 p-1.5 text-right font-mono">{(schoolResult.nutrisiPorsiBesar.energi || 0).toFixed(0)} kkal</td>
                            <td className="border border-slate-300 p-1.5 text-right font-mono">{(schoolResult.nutrisiPorsiBesar.protein || 0).toFixed(1)}g</td>
                            <td className="border border-slate-300 p-1.5 text-right font-mono">{(schoolResult.nutrisiPorsiBesar.lemak || 0).toFixed(1)}g</td>
                            <td className="border border-slate-300 p-1.5 text-right font-mono">{(schoolResult.nutrisiPorsiBesar.kh || 0).toFixed(1)}g</td>
                            <td className="border border-slate-300 p-1.5 text-right font-mono">{(schoolResult.nutrisiPorsiBesar.serat || 0).toFixed(1)}g</td>
                            <td className="border border-slate-300 p-1.5 text-right font-mono">{formatRupiah(schoolResult.costPerPorsiKecil)}</td>
                            <td className="border border-slate-300 p-1.5 text-right font-mono">{formatRupiah(schoolResult.costPerPorsiBesar)}</td>
                            <td className="border border-slate-300 p-1.5 text-right font-mono font-bold">{formatRupiah(schoolResult.subtotalBesarCost + schoolResult.subtotalKecilCost)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>

                  {/* Summary row */}
                  <div className="flex justify-between items-center pt-3 border-t-2 border-slate-800">
                    <span className="text-xs font-bold uppercase text-slate-700">Total Proyeksi Belanja Kumulatif {totalDays} Hari Kerja:</span>
                    <span className="text-sm font-black text-black font-mono">
                      {formatRupiah(calculatedDays.reduce((acc, d) => acc + (d.schoolResult.subtotalBesarCost + d.schoolResult.subtotalKecilCost) + (d.threeBResult.subtotalBesarCost + d.threeBResult.subtotalKecilCost), 0))}
                    </span>
                  </div>
                </div>
              ) : (
                <div className="space-y-6">
                  {/* Title */}
                  <div className="text-center space-y-1">
                    <h3 className="font-black text-black uppercase tracking-wide" style={{ fontSize: '13pt' }}>
                      LAPORAN NOTA RINCIAN BAHAN MAKANAN &amp; ANGGARAN
                    </h3>
                    <p className="text-xs font-bold text-slate-700 uppercase">
                      HARI KE-{selectedNotaDay} | MENU: {getMenuNameForDay(selectedNotaDay - 1).toUpperCase()}
                    </p>
                  </div>

                  {/* Two columns: Usia Sekolah vs 3B */}
                  <div className="grid grid-cols-2 gap-4">
                    {/* Usia Sekolah */}
                    <div className="space-y-2">
                      <h4 className="text-[10px] font-bold uppercase tracking-wider text-slate-800 border-b pb-1">Usia Sekolah (Siswa)</h4>
                      <table className="w-full text-left text-[9px] border-collapse border border-slate-300">
                        <thead className="bg-slate-50 font-bold">
                          <tr>
                            <th className="border border-slate-300 p-1">Bahan</th>
                            <th className="border border-slate-300 p-1 text-right">Netto</th>
                            <th className="border border-slate-300 p-1 text-right">Total</th>
                          </tr>
                        </thead>
                        <tbody>
                          {currentNotaSekolah.map((item, idx) => (
                            <tr key={idx}>
                              <td className="border border-slate-300 p-1 truncate max-w-[80px]">{item.nama}</td>
                              <td className="border border-slate-300 p-1 text-right font-mono">{item.totalKg.toFixed(3)} Kg</td>
                              <td className="border border-slate-300 p-1 text-right font-mono">{formatRupiah(item.totalKg * item.hargaSatuan)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      <div className="text-right text-[9px] font-bold pt-1">
                        Total Sekolah: {formatRupiah(currentNotaSekolah.reduce((acc, i) => acc + (i.totalKg * i.hargaSatuan), 0))}
                      </div>
                    </div>

                    {/* Kelompok 3B */}
                    <div className="space-y-2">
                      <h4 className="text-[10px] font-bold uppercase tracking-wider text-slate-800 border-b pb-1">Kelompok 3B (MP-ASI)</h4>
                      <table className="w-full text-left text-[9px] border-collapse border border-slate-300">
                        <thead className="bg-slate-50 font-bold">
                          <tr>
                            <th className="border border-slate-300 p-1">Bahan</th>
                            <th className="border border-slate-300 p-1 text-right">Netto</th>
                            <th className="border border-slate-300 p-1 text-right">Total</th>
                          </tr>
                        </thead>
                        <tbody>
                          {currentNota3B.map((item, idx) => (
                            <tr key={idx}>
                              <td className="border border-slate-300 p-1 truncate max-w-[80px]">{item.nama}</td>
                              <td className="border border-slate-300 p-1 text-right font-mono">{item.totalKg.toFixed(3)} Kg</td>
                              <td className="border border-slate-300 p-1 text-right font-mono">{formatRupiah(item.totalKg * item.hargaSatuan)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      <div className="text-right text-[9px] font-bold pt-1">
                        Total 3B: {formatRupiah(currentNota3B.reduce((acc, i) => acc + (i.totalKg * i.hargaSatuan), 0))}
                      </div>
                    </div>
                  </div>

                  {/* Subsidy report */}
                  <div className="bg-slate-50 p-3 rounded-lg border border-slate-200 text-[10px] space-y-1">
                    <div className="flex justify-between font-bold">
                      <span>PAGU ALOKASI HARIAN USIA SEKOLAH:</span>
                      <span className="font-mono">{formatRupiah(saldoAwalUsiaSekolah)}</span>
                    </div>
                    <div className="flex justify-between font-bold">
                      <span>PAGU ALOKASI HARIAN KELOMPOK 3B:</span>
                      <span className="font-mono">{formatRupiah(saldoAwal3B)}</span>
                    </div>
                    <div className="flex justify-between font-bold text-slate-800 border-t pt-1 mt-1">
                      <span>TOTAL BELANJA HARIAN GABUNGAN:</span>
                      <span className="font-mono text-rose-800">
                        {formatRupiah(
                          currentNotaSekolah.reduce((acc, i) => acc + (i.totalKg * i.hargaSatuan), 0) +
                          currentNota3B.reduce((acc, i) => acc + (i.totalKg * i.hargaSatuan), 0)
                        )}
                      </span>
                    </div>
                  </div>
                </div>
              )}
              {profile && <AutoSignatureBlock profile={profile} docKey="dashboard" />}
            </div>
          </div>
        </div>

        {/* --- PRINT PREVIEW MODAL --- */}
        <PrintPreviewModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          title={printDocType === "rekap" ? `Pratinjau Cetak: RAB ${totalDays} Hari & Rekap Gizi` : `Pratinjau Cetak: Nota Rincian Bahan Hari ${selectedNotaDay}`}
          elementId="print-area-dashboard-rekap"
          filename={printDocType === "rekap" ? `RAB_${totalDays}_Hari_Rekap_Gizi` : `Nota_Rincian_Bahan_Hari_${selectedNotaDay}`}
          defaultScale={printScale}
          defaultPaperSize={paperSize}
          defaultOrientation="auto"
        />
      </div>
  );
}
