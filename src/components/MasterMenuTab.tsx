/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useEffect, useMemo, useState } from "react";
import {
  MasterMenu,
  MasterMenuDietType,
  MasterMenuDurationType,
  MasterMenuPlan,
  MasterMenuScheduleConfig,
  MenuItem,
  SPPGProfile,
  UnifiedKopConfig,
  DEFAULT_UNIFIED_KOP_CONFIG,
} from "../types";
import { downloadElementAsImage, exportToPDF, printElementSmart } from "../lib/printUtils";
import { KopSuratRenderHeader, LogoCrop } from "./KopSuratConfigSection";
import AutoSignatureBlock from "./AutoSignatureBlock";
import PrintPreviewModal from "./PrintPreviewModal";
import {
  SASARAN_KATEGORI_LIST,
  SEKOLAH_SASARAN_IDS,
  TIGAB_SASARAN_IDS,
  MPASI_SASARAN_IDS,
  buildCategoryLabelFromIds,
  composeMenuName,
  computeNextPeriodStart,
  createEmptyMenuItem,
  ensureMasterMenuPlans,
  formatIndonesianDate,
  formatIndonesianMonthYear,
  generateScheduleDates,
  getDurationBadgeLabel,
  groupDatesIntoWeeks,
  normalizeMenuItem,
  resizeMenuDays,
  syncMasterMenuToLegacy,
} from "../utils/menuSchedule";
import {
  ArrowRight,
  Calendar,
  Check,
  CheckCircle2,
  Clipboard,
  Copy,
  Download,
  Eye,
  Image as ImageIcon,
  Layers,
  LayoutGrid,
  Plus,
  Printer,
  RotateCcw,
  Send,
  Table2,
  Trash2,
  Users,
} from "lucide-react";

interface MasterMenuTabProps {
  menu: MasterMenu;
  onChange: (updated: MasterMenu) => void;
  profile?: SPPGProfile;
  onProfileChange?: (updated: SPPGProfile) => void;
  onSendToFoodCost?: (updatedMenu: MasterMenu, dates: string[]) => void;
  onNavigateToFoodCost?: () => void;
  customLogo?: string;
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
  isAdmin?: boolean;
}

const DURATION_OPTIONS: { value: MasterMenuDurationType; label: string; hint: string }[] = [
  { value: "1_hari", label: "1 Hari", hint: "Harian" },
  { value: "1_minggu", label: "1 Minggu", hint: "1 Pekan" },
  { value: "2_minggu", label: "2 Minggu", hint: "2 Pekan" },
  { value: "3_minggu", label: "3 Minggu", hint: "3 Pekan" },
  { value: "4_minggu", label: "4 Minggu", hint: "4 Pekan" },
  { value: "1_bulan", label: "1 Bulan Penuh", hint: "Kalender" },
];

const DIET_OPTIONS: { value: MasterMenuDietType; label: string }[] = [
  { value: "Standar", label: "Standar (Basah)" },
  { value: "Alergi", label: "Diet Alergi" },
  { value: "Kering", label: "Kering / Kudapan" },
  { value: "MP-ASI", label: "MP-ASI (6-11 Bln)" },
];

const MENU_COMPONENT_ROWS: {
  key: keyof Pick<MenuItem, "karbohidrat" | "laukHewani" | "laukNabati" | "sayur" | "buah" | "susu">;
  label: string;
  placeholder: string;
}[] = [
  { key: "karbohidrat", label: "Makanan Pokok", placeholder: "Nasi Putih" },
  { key: "laukHewani", label: "Lauk Hewani", placeholder: "Ayam / Telur / Ikan" },
  { key: "laukNabati", label: "Lauk Nabati", placeholder: "Tempe / Tahu" },
  { key: "sayur", label: "Sayuran", placeholder: "Tumis / Sup Sayur" },
  { key: "buah", label: "Buah", placeholder: "Pisang / Semangka / Jeruk" },
  { key: "susu", label: "Susu", placeholder: "Susu UHT / (opsional)" },
];

export default function MasterMenuTab({
  menu,
  onChange,
  profile,
  onProfileChange,
  onSendToFoodCost,
  onNavigateToFoodCost,
  customLogo,
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
  isAdmin = false,
}: MasterMenuTabProps) {
  const normalizedMenu = useMemo(
    () => ensureMasterMenuPlans(menu, profile?.periodeDates),
    [menu, profile?.periodeDates]
  );

  const plans = normalizedMenu.plans || [];
  const activePlan =
    plans.find((p) => p.id === normalizedMenu.activePlanId) || plans[0];

  const scheduleConfig: MasterMenuScheduleConfig = normalizedMenu.scheduleConfig!;
  const plannedDates = useMemo(() => {
    if (scheduleConfig.periodeDates && scheduleConfig.periodeDates.length > 0) {
      return scheduleConfig.periodeDates;
    }
    return generateScheduleDates(scheduleConfig);
  }, [scheduleConfig]);

  const totalDays = Math.max(1, plannedDates.length);
  const weekBlocks = useMemo(() => groupDatesIntoWeeks(plannedDates), [plannedDates]);

  const [viewMode, setViewMode] = useState<"table" | "cards" | "print">("table");
  const [showCustomDates, setShowCustomDates] = useState<boolean>(false);
  const [showPrintModal, setShowPrintModal] = useState<boolean>(false);
  const [isDownloading, setIsDownloading] = useState<string | null>(null);
  const [copiedItem, setCopiedItem] = useState<{ sourceLabel: string; item: MenuItem } | null>(null);
  const [justPastedDay, setJustPastedDay] = useState<number | null>(null);
  const [sendFeedback, setSendFeedback] = useState<string | null>(null);
  const [confirmModal, setConfirmModal] = useState<{
    title: string;
    message: string;
    confirmLabel: string;
    onConfirm: () => void;
  } | null>(null);

  // Sync KOP defaults when profile updates and no custom localStorage exists
  useEffect(() => {
    if (profile) {
      if (!localStorage.getItem("kop_line3") && profile.namaLembaga) {
        setKopLine3(profile.namaLembaga);
      }
      if (!localStorage.getItem("kop_line4") && profile.alamat) {
        setKopLine4(`Alamat : ${profile.alamat}`);
      }
    }
  }, [profile]);

  const commitMenuUpdate = (nextMenu: MasterMenu) => {
    const synced = syncMasterMenuToLegacy(nextMenu);
    onChange(synced);
  };

  const applyScheduleConfig = (nextConfig: MasterMenuScheduleConfig, explicitDates?: string[]) => {
    const newDates = explicitDates ?? generateScheduleDates(nextConfig);
    const updatedConfig: MasterMenuScheduleConfig = {
      ...nextConfig,
      periodeDates: newDates,
    };
    const resizedPlans = plans.map((p) => ({
      ...p,
      days: resizeMenuDays(p.days, newDates.length),
    }));

    const nextMenu: MasterMenu = {
      ...normalizedMenu,
      scheduleConfig: updatedConfig,
      plans: resizedPlans,
    };
    commitMenuUpdate(nextMenu);

    if (onProfileChange && profile) {
      const nextPeriodStart = computeNextPeriodStart(
        newDates,
        updatedConfig.workDaysPerWeek
      );
      onProfileChange({
        ...profile,
        periodeDates: newDates,
        awalPeriodeBerikutnya: nextPeriodStart || profile.awalPeriodeBerikutnya,
      });
    }
  };

  const handleSelectDuration = (durationType: MasterMenuDurationType) => {
    const calendarMonth =
      scheduleConfig.calendarMonth ||
      (scheduleConfig.startDate ? scheduleConfig.startDate.slice(0, 7) : "2026-11");
    applyScheduleConfig({
      ...scheduleConfig,
      durationType,
      calendarMonth,
    });
  };

  const handleStartDateChange = (newStartDate: string) => {
    if (!newStartDate) return;
    applyScheduleConfig({
      ...scheduleConfig,
      startDate: newStartDate,
      calendarMonth: newStartDate.slice(0, 7),
    });
  };

  const handleCalendarMonthChange = (newMonth: string) => {
    if (!newMonth) return;
    applyScheduleConfig({
      ...scheduleConfig,
      calendarMonth: newMonth,
      startDate: `${newMonth}-01`,
    });
  };

  const handleWorkDaysChange = (workDaysPerWeek: 5 | 6 | 7) => {
    applyScheduleConfig({
      ...scheduleConfig,
      workDaysPerWeek,
    });
  };

  const handleSingleDateChange = (idx: number, val: string) => {
    const updatedDates = [...plannedDates];
    updatedDates[idx] = val;
    applyScheduleConfig(scheduleConfig, updatedDates);
  };

  const updateActivePlan = (updater: (plan: MasterMenuPlan) => MasterMenuPlan) => {
    if (!activePlan) return;
    const nextPlans = plans.map((p) => (p.id === activePlan.id ? updater(p) : p));
    commitMenuUpdate({
      ...normalizedMenu,
      plans: nextPlans,
    });
  };

  const handleSelectPlan = (planId: string) => {
    commitMenuUpdate({
      ...normalizedMenu,
      activePlanId: planId,
    });
  };

  const handleAddPlan = () => {
    const newId = `plan_${Date.now()}`;
    const newPlan: MasterMenuPlan = {
      id: newId,
      title: `Paket Menu ${plans.length + 1}`,
      sasaranIds: ["sd_kelas_1_3"],
      dietType: "Standar",
      days: Array.from({ length: totalDays }, () => createEmptyMenuItem()),
    };
    commitMenuUpdate({
      ...normalizedMenu,
      plans: [...plans, newPlan],
      activePlanId: newId,
    });
  };

  const handleDuplicatePlan = () => {
    if (!activePlan) return;
    const newId = `plan_${Date.now()}`;
    const clonedPlan: MasterMenuPlan = {
      ...activePlan,
      id: newId,
      title: `${activePlan.title} (Salinan)`,
      sasaranIds: [...activePlan.sasaranIds],
      days: activePlan.days.map((d) => ({ ...d })),
    };
    commitMenuUpdate({
      ...normalizedMenu,
      plans: [...plans, clonedPlan],
      activePlanId: newId,
    });
  };

  const handleDeleteActivePlan = () => {
    if (!activePlan || plans.length <= 1) return;
    setConfirmModal({
      title: "Hapus Paket Master Menu?",
      message: `Paket "${activePlan.title}" akan dihapus beserta seluruh menu hariannya. Tindakan ini tidak dapat dibatalkan.`,
      confirmLabel: "Ya, Hapus Paket",
      onConfirm: () => {
        const remaining = plans.filter((p) => p.id !== activePlan.id);
        commitMenuUpdate({
          ...normalizedMenu,
          plans: remaining,
          activePlanId: remaining[0]?.id,
        });
        setConfirmModal(null);
      },
    });
  };

  const handleToggleSasaran = (sasaranId: string) => {
    if (!activePlan) return;
    const exists = activePlan.sasaranIds.includes(sasaranId);
    const nextIds = exists
      ? activePlan.sasaranIds.filter((id) => id !== sasaranId)
      : [...activePlan.sasaranIds, sasaranId];

    // If current title still equals the previous auto-label, keep it cleanly updated
    const prevAuto = buildCategoryLabelFromIds(activePlan.sasaranIds);
    const shouldAutoUpdateTitle =
      !activePlan.title || activePlan.title === prevAuto || activePlan.title === "BELUM ADA KATEGORI";

    updateActivePlan((p) => ({
      ...p,
      sasaranIds: nextIds,
      title: shouldAutoUpdateTitle ? buildCategoryLabelFromIds(nextIds) : p.title,
    }));
  };

  const handleApplySasaranPreset = (ids: string[], title: string, diet?: MasterMenuDietType) => {
    updateActivePlan((p) => ({
      ...p,
      sasaranIds: [...ids],
      title,
      dietType: diet || p.dietType,
    }));
  };

  const handleDayFieldChange = (dayIndex: number, field: keyof MenuItem, value: string) => {
    if (!activePlan) return;
    const nextDays = resizeMenuDays(activePlan.days, totalDays);
    const current = normalizeMenuItem(nextDays[dayIndex]);
    const updatedItem: MenuItem = {
      ...current,
      [field]: value,
    };

    if (field === "buah" || field === "susu") {
      const buahVal = field === "buah" ? value : updatedItem.buah || "";
      const susuVal = field === "susu" ? value : updatedItem.susu || "";
      updatedItem.buahSusu = [buahVal.trim(), susuVal.trim()].filter(Boolean).join(" + ");
    }

    nextDays[dayIndex] = updatedItem;
    updateActivePlan((p) => ({ ...p, days: nextDays }));
  };

  const getDayMenuItem = (dayIndex: number): MenuItem => {
    if (!activePlan) return createEmptyMenuItem();
    return normalizeMenuItem(activePlan.days[dayIndex]);
  };

  const handleCopyDay = (dayIdx: number) => {
    const item = getDayMenuItem(dayIdx);
    setCopiedItem({
      sourceLabel: `${activePlan?.title || "Menu"} • Hari ${dayIdx + 1}`,
      item: { ...item },
    });
  };

  const handlePasteDay = (dayIdx: number) => {
    if (!copiedItem || !activePlan) return;
    const nextDays = resizeMenuDays(activePlan.days, totalDays);
    nextDays[dayIdx] = { ...normalizeMenuItem(copiedItem.item) };
    updateActivePlan((p) => ({ ...p, days: nextDays }));
    setJustPastedDay(dayIdx);
    setTimeout(() => setJustPastedDay(null), 1500);
  };

  const handleResetDay = (dayIdx: number) => {
    setConfirmModal({
      title: `Kosongkan Menu Hari Ke-${dayIdx + 1}?`,
      message: `Seluruh isian Makanan Pokok, Lauk, Sayuran, Buah, dan Susu pada Hari Ke-${dayIdx + 1} (${formatIndonesianDate(plannedDates[dayIdx]) || ""}) akan dikosongkan.`,
      confirmLabel: "Kosongkan Hari Ini",
      onConfirm: () => {
        if (!activePlan) return;
        const nextDays = resizeMenuDays(activePlan.days, totalDays);
        nextDays[dayIdx] = createEmptyMenuItem();
        updateActivePlan((p) => ({ ...p, days: nextDays }));
        setConfirmModal(null);
      },
    });
  };

  const handleResetActivePlanDays = () => {
    if (!activePlan) return;
    setConfirmModal({
      title: `Reset Semua ${totalDays} Hari pada Paket Aktif?`,
      message: `Seluruh isian menu (${totalDays} hari) pada paket "${activePlan.title}" akan dikosongkan.`,
      confirmLabel: "Ya, Reset Paket Ini",
      onConfirm: () => {
        updateActivePlan((p) => ({
          ...p,
          days: Array.from({ length: totalDays }, () => createEmptyMenuItem()),
        }));
        setConfirmModal(null);
      },
    });
  };

  const handleSendMenuToFoodCost = () => {
    const syncedMenu = syncMasterMenuToLegacy({
      ...normalizedMenu,
      lastSentAt: new Date().toISOString(),
    });
    onChange(syncedMenu);

    if (onProfileChange && profile) {
      const nextPeriodStart = computeNextPeriodStart(
        plannedDates,
        scheduleConfig.workDaysPerWeek
      );
      onProfileChange({
        ...profile,
        periodeDates: plannedDates,
        awalPeriodeBerikutnya: nextPeriodStart || profile.awalPeriodeBerikutnya,
      });
    }

    if (onSendToFoodCost) {
      onSendToFoodCost(syncedMenu, plannedDates);
    }

    setSendFeedback(
      `Perencanaan menu ${totalDays} hari (${plans.length} paket kategori) dan tanggal kalender telah dikirim otomatis ke panel Food Cost.`
    );
  };

  const defaultLogoSrc = customLogo || "/assets/images/app_logo_sppg.png";

  const bannerCategoryTitle = useMemo(() => {
    if (!activePlan) return "";
    const titlePart = (activePlan.title || buildCategoryLabelFromIds(activePlan.sasaranIds)).toUpperCase();
    const dietSuffix =
      activePlan.dietType && activePlan.dietType !== "Standar"
        ? ` (${activePlan.dietType.toUpperCase()})`
        : "";
    return `${titlePart}${dietSuffix}`;
  }, [activePlan]);

  const masterBannerHeadline = useMemo(() => {
    if (scheduleConfig.durationType === "1_bulan") {
      const monthName = formatIndonesianMonthYear(scheduleConfig.calendarMonth).toUpperCase();
      return `MASTER MENU ${monthName} (${totalDays} HARI) ${bannerCategoryTitle}`.trim();
    }
    return `MASTER MENU ${totalDays} HARI ${bannerCategoryTitle}`.trim();
  }, [scheduleConfig.durationType, scheduleConfig.calendarMonth, totalDays, bannerCategoryTitle]);

  const activeKopLeftSrc = leftLogo || defaultLogoSrc;

  return (
    <div id="master-menu-tab-container" className="space-y-5 font-sans">
      {/* SECTION 1: KONTROL JADWAL & KONFIGURASI TANGGAL TERPUSAT */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 space-y-4 no-print">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-4 border-b border-slate-100">
          <div className="space-y-1">
            <div className="flex items-center gap-2 text-slate-900">
              <Calendar className="w-5 h-5 text-emerald-700 shrink-0" />
              <h2 className="text-base font-bold tracking-tight">
                Perencanaan Master Menu & Jadwal Terpusat
              </h2>
            </div>
            <p className="text-xs text-slate-500">
              Semua tanggal di Profil SPPG, Penerima Manfaat, Food Cost, dan Nota Pesanan mengikuti konfigurasi jadwal di sini.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              id="btn-send-menu-to-foodcost"
              onClick={handleSendMenuToFoodCost}
              className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-2 transition cursor-pointer whitespace-nowrap shrink-0 shadow-xs"
            >
              <Send className="w-3.5 h-3.5" />
              Kirim Perencanaan Menu ke Food Cost
            </button>
          </div>
        </div>

        {/* Duration + Start Date + Workdays */}
        <div className="grid grid-cols-1 xl:grid-cols-12 gap-4 items-end">
          <div className="xl:col-span-6 space-y-1.5">
            <label className="text-[11px] font-bold text-slate-600 block">
              Rentang Tampilan & Siklus Kalender
            </label>
            <div className="grid grid-cols-3 sm:grid-cols-6 gap-1.5 bg-slate-100 p-1 rounded-xl border border-slate-200">
              {DURATION_OPTIONS.map((opt) => {
                const isActive = scheduleConfig.durationType === opt.value;
                return (
                  <button
                    key={opt.value}
                    type="button"
                    id={`btn-duration-${opt.value}`}
                    onClick={() => handleSelectDuration(opt.value)}
                    className={`px-2 py-2 rounded-lg text-xs font-bold transition text-center cursor-pointer whitespace-nowrap ${
                      isActive
                        ? "bg-white text-slate-900 shadow-xs border border-slate-200"
                        : "text-slate-600 hover:text-slate-900 hover:bg-white/60"
                    }`}
                  >
                    {opt.label}
                  </button>
                );
              })}
            </div>
          </div>

          <div className="xl:col-span-3 space-y-1.5">
            <label className="text-[11px] font-bold text-slate-600 block">
              {scheduleConfig.durationType === "1_bulan"
                ? "Pilih Bulan Kalender"
                : "Tanggal Mulai Perencanaan"}
            </label>
            {scheduleConfig.durationType === "1_bulan" ? (
              <input
                type="month"
                id="input-master-menu-month"
                value={scheduleConfig.calendarMonth || scheduleConfig.startDate.slice(0, 7)}
                onChange={(e) => handleCalendarMonthChange(e.target.value)}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
              />
            ) : (
              <input
                type="date"
                id="input-master-menu-start-date"
                value={scheduleConfig.startDate}
                onChange={(e) => handleStartDateChange(e.target.value)}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
              />
            )}
          </div>

          <div className="xl:col-span-3 space-y-1.5">
            <label className="text-[11px] font-bold text-slate-600 block">
              Hari Operasional per Minggu
            </label>
            <div className="flex items-center gap-1.5">
              <select
                id="select-master-workdays"
                value={scheduleConfig.workDaysPerWeek}
                onChange={(e) => handleWorkDaysChange(Number(e.target.value) as 5 | 6 | 7)}
                className="flex-1 px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
              >
                <option value={5}>5 Hari (Senin – Jum&apos;at)</option>
                <option value={6}>6 Hari (Senin – Sabtu)</option>
                <option value={7}>7 Hari (Semua Hari)</option>
              </select>
              <button
                type="button"
                onClick={() => setShowCustomDates(!showCustomDates)}
                className={`px-3 py-2 rounded-xl text-xs font-semibold border transition cursor-pointer whitespace-nowrap ${
                  showCustomDates
                    ? "bg-slate-900 text-white border-slate-900"
                    : "bg-white text-slate-700 border-slate-200 hover:bg-slate-50"
                }`}
                title="Sesuaikan tanggal satu per satu jika ada hari libur nasional"
              >
                Ubah Satuan
              </button>
            </div>
          </div>
        </div>

        {/* Summary text metadata (Zero-Pill clean inline format) */}
        <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500 pt-1">
          <div className="flex flex-wrap items-center gap-2">
            <span className="font-semibold text-slate-700">
              {getDurationBadgeLabel(scheduleConfig.durationType, totalDays)}
            </span>
            <span aria-hidden="true">·</span>
            <span className="font-mono tabular-nums">
              {formatIndonesianDate(plannedDates[0])} s.d.{" "}
              {formatIndonesianDate(plannedDates[plannedDates.length - 1])}
            </span>
            <span aria-hidden="true">·</span>
            <span>{weekBlocks.length} Blok Minggu</span>
          </div>
          {normalizedMenu.lastSentAt && (
            <span className="text-emerald-700 font-medium font-mono tabular-nums">
              Terakhir dikirim ke Food Cost:{" "}
              {new Date(normalizedMenu.lastSentAt).toLocaleTimeString("id-ID", {
                hour: "2-digit",
                minute: "2-digit",
              })}
            </span>
          )}
        </div>

        {/* Collapsible Individual Date Customizer */}
        {showCustomDates && (
          <div className="pt-3 border-t border-slate-100">
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold text-slate-700">
                Penyesuaian Tanggal Perencanaan Satuan ({totalDays} Hari)
              </span>
              <button
                type="button"
                onClick={() => applyScheduleConfig({ ...scheduleConfig })}
                className="text-xs font-semibold text-indigo-600 hover:underline cursor-pointer"
              >
                Hitung Ulang Otomatis dari Kalender
              </button>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-7 gap-2.5">
              {plannedDates.map((dStr, idx) => (
                <div key={idx} className="bg-slate-50 p-2 rounded-lg border border-slate-200 space-y-1">
                  <span className="text-[10px] font-bold text-slate-500 block font-mono">
                    MENU {idx + 1}
                  </span>
                  <input
                    type="date"
                    value={dStr}
                    onChange={(e) => handleSingleDateChange(idx, e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded px-1.5 py-1 text-[11px] font-mono text-slate-800"
                  />
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Feedback Banner when sent to Food Cost */}
      {sendFeedback && (
        <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 no-print">
          <div className="flex items-center gap-2.5 text-emerald-900 text-xs font-semibold">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>{sendFeedback}</span>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            {onNavigateToFoodCost && (
              <button
                type="button"
                onClick={onNavigateToFoodCost}
                className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 cursor-pointer transition"
              >
                Buka Panel Food Cost
                <ArrowRight className="w-3.5 h-3.5" />
              </button>
            )}
            <button
              type="button"
              onClick={() => setSendFeedback(null)}
              className="px-2.5 py-1.5 text-xs font-semibold text-emerald-700 hover:bg-emerald-100 rounded-lg cursor-pointer"
            >
              Tutup
            </button>
          </div>
        </div>
      )}

      {/* SECTION 2: PAKET MASTER MENU & PEMILIHAN KATEGORI SASARAN (GAMBAR 2) */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 space-y-4 no-print">
        {/* Plans tabs */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 pb-3 border-b border-slate-100">
          <div className="flex items-center gap-2">
            <Layers className="w-4 h-4 text-indigo-600 shrink-0" />
            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Paket Menu & Pengelompokan Kategori Sasaran
              </h3>
              <p className="text-xs text-slate-500">
                Pilih 1 kategori atau centang bebas beberapa kategori sekaligus jika menggunakan Master Menu yang sama.
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              type="button"
              id="btn-add-menu-plan"
              onClick={handleAddPlan}
              className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 cursor-pointer transition whitespace-nowrap"
            >
              <Plus className="w-3.5 h-3.5" />
              Buat Paket Baru
            </button>
            <button
              type="button"
              onClick={handleDuplicatePlan}
              className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold flex items-center gap-1.5 cursor-pointer transition whitespace-nowrap"
              title="Salin paket aktif menjadi paket baru"
            >
              <Copy className="w-3.5 h-3.5" />
              Duplikat
            </button>
            {plans.length > 1 && (
              <button
                type="button"
                onClick={handleDeleteActivePlan}
                className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg text-xs font-semibold flex items-center gap-1.5 cursor-pointer transition whitespace-nowrap"
              >
                <Trash2 className="w-3.5 h-3.5" />
                Hapus Paket
              </button>
            )}
          </div>
        </div>

        {/* Plan Switcher Row */}
        <div className="flex flex-wrap gap-2">
          {plans.map((p) => {
            const isActive = activePlan?.id === p.id;
            const catCount = p.sasaranIds.length;
            return (
              <button
                key={p.id}
                type="button"
                onClick={() => handleSelectPlan(p.id)}
                className={`px-3.5 py-2 rounded-xl text-xs font-bold border transition text-left cursor-pointer ${
                  isActive
                    ? "bg-slate-900 text-white border-slate-900 shadow-xs"
                    : "bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100"
                }`}
              >
                <div className="flex items-center gap-2 whitespace-nowrap">
                  <span>{p.title || "Paket Tanpa Judul"}</span>
                  <span
                    className={`text-[10px] font-mono ${
                      isActive ? "text-emerald-300" : "text-slate-400"
                    }`}
                  >
                    {p.dietType} · {catCount} Kategori
                  </span>
                </div>
              </button>
            );
          })}
        </div>

        {/* Active Plan Config: Title + Diet + Multi-Category Sasaran Grid (Gambar 2) */}
        {activePlan && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 pt-2">
            {/* Left col: Title, Diet & Quick Presets */}
            <div className="lg:col-span-4 space-y-4">
              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-slate-600 block">
                  Label Kategori pada Header Tabel (cth: 1-3 SD)
                </label>
                <input
                  type="text"
                  id="input-plan-title"
                  value={activePlan.title}
                  onChange={(e) =>
                    updateActivePlan((p) => ({ ...p, title: e.target.value }))
                  }
                  placeholder="Misal: 1-3 SD atau TK/PAUD & SD 1-3"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-slate-600 block">
                  Jenis Diet / Sinkronisasi Food Cost
                </label>
                <select
                  id="select-plan-diet"
                  value={activePlan.dietType}
                  onChange={(e) =>
                    updateActivePlan((p) => ({
                      ...p,
                      dietType: e.target.value as MasterMenuDietType,
                    }))
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
                >
                  {DIET_OPTIONS.map((d) => (
                    <option key={d.value} value={d.value}>
                      {d.label}
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-slate-600 block">
                  Preset Cepat Gabungan Kategori
                </label>
                <div className="flex flex-wrap gap-1.5">
                  <button
                    type="button"
                    onClick={() =>
                      handleApplySasaranPreset(["tk_paud_lb", "sd_kelas_1_3"], "1-3 SD & TK/PAUD")
                    }
                    className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[11px] font-semibold cursor-pointer"
                  >
                    TK & 1-3 SD
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      handleApplySasaranPreset(
                        ["sd_kelas_4_6", "smp_mts_smplb", "sma_smk_ma", "pendidik", "tenaga_kependidikan"],
                        "SD 4-6, SMP, SMA & Pendidik"
                      )
                    }
                    className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[11px] font-semibold cursor-pointer"
                  >
                    SD 4-6 s/d SMA
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      handleApplySasaranPreset(SEKOLAH_SASARAN_IDS, "Usia Sekolah & Pendidik")
                    }
                    className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[11px] font-semibold cursor-pointer"
                  >
                    Semua Sekolah
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      handleApplySasaranPreset(
                        ["anak_balita", "anak_balita_13_59", "ibu_hamil", "ibu_menyusui"],
                        "3B (Balita, Bumil & Busui)"
                      )
                    }
                    className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[11px] font-semibold cursor-pointer"
                  >
                    Kelompok 3B
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      handleApplySasaranPreset(
                        MPASI_SASARAN_IDS,
                        "Balita 6-11 Bulan (MP-ASI)",
                        "MP-ASI"
                      )
                    }
                    className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[11px] font-semibold cursor-pointer"
                  >
                    Balita 6-11 Bln
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      handleApplySasaranPreset(
                        SASARAN_KATEGORI_LIST.map((s) => s.id),
                        "Semua Kelompok Sasaran"
                      )
                    }
                    className="px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 rounded-lg text-[11px] font-semibold cursor-pointer"
                  >
                    Gabung Semua (12)
                  </button>
                </div>
              </div>
            </div>

            {/* Right col: 12 Kelompok Sasaran Multi-Select matching Gambar 2 */}
            <div className="lg:col-span-8">
              <div className="border border-slate-300 rounded-xl overflow-hidden">
                <div
                  className="px-3.5 py-2 border-b border-slate-300 flex items-center justify-between"
                  style={{ backgroundColor: "#92D050" }}
                >
                  <div className="flex items-center gap-2 text-xs font-extrabold text-slate-950">
                    <Users className="w-3.5 h-3.5" />
                    <span>Kelompok Sasaran (Pilih Bebas Satu atau Beberapa Kategori)</span>
                  </div>
                  <span className="text-[11px] font-bold text-slate-900 font-mono">
                    {activePlan.sasaranIds.length} / {SASARAN_KATEGORI_LIST.length} dipilih
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 divide-y sm:divide-y-0 bg-white">
                  {SASARAN_KATEGORI_LIST.map((sasaran) => {
                    const checked = activePlan.sasaranIds.includes(sasaran.id);
                    return (
                      <label
                        key={sasaran.id}
                        className={`flex items-center gap-2.5 px-3 py-2.5 text-xs border-b border-r border-slate-200 cursor-pointer transition select-none ${
                          checked
                            ? "bg-emerald-50/70 text-slate-950 font-semibold"
                            : "bg-white text-slate-700 hover:bg-slate-50"
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => handleToggleSasaran(sasaran.id)}
                          className="w-4 h-4 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer shrink-0"
                        />
                        <span className="truncate" title={sasaran.label}>
                          {sasaran.label}
                        </span>
                      </label>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* SECTION 3: TOOLBAR TAMPILAN, EKSPOR & KOP SURAT */}
      <div className="bg-white rounded-2xl border border-slate-200 p-4 flex flex-col lg:flex-row lg:items-center justify-between gap-4 no-print">
        {/* View mode toggle */}
        <div className="flex flex-wrap items-center gap-2">
          <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200">
            <button
              type="button"
              id="btn-mode-table"
              onClick={() => setViewMode("table")}
              className={`px-3.5 py-2 text-xs font-bold rounded-lg transition flex items-center gap-1.5 cursor-pointer whitespace-nowrap ${
                viewMode === "table"
                  ? "bg-white text-slate-900 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <Table2 className="w-3.5 h-3.5" />
              Tabel Master Menu (Edit Langsung)
            </button>
            <button
              type="button"
              id="btn-mode-cards"
              onClick={() => setViewMode("cards")}
              className={`px-3.5 py-2 text-xs font-bold rounded-lg transition flex items-center gap-1.5 cursor-pointer whitespace-nowrap ${
                viewMode === "cards"
                  ? "bg-white text-slate-900 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <LayoutGrid className="w-3.5 h-3.5" />
              Kartu Input & Nama Menu
            </button>
            <button
              type="button"
              id="btn-mode-print"
              onClick={() => setViewMode("print")}
              className={`px-3.5 py-2 text-xs font-bold rounded-lg transition flex items-center gap-1.5 cursor-pointer whitespace-nowrap ${
                viewMode === "print"
                  ? "bg-white text-slate-900 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <Eye className="w-3.5 h-3.5" />
              Pratinjau Cetak Resmi
            </button>
          </div>
        </div>

        {/* Actions: Reset + Export + Print */}
        <div className="flex flex-wrap items-center gap-2">
          {copiedItem && (
            <span className="text-[11px] font-semibold text-indigo-700 mr-1">
              Tersalin: {copiedItem.sourceLabel}
            </span>
          )}
          <button
            type="button"
            onClick={handleResetActivePlanDays}
            className="px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-xl text-xs font-semibold flex items-center gap-1.5 cursor-pointer transition whitespace-nowrap"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            Reset Paket Ini
          </button>
          <button
            type="button"
            onClick={() => setShowPrintModal(true)}
            className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-xl text-xs font-semibold flex items-center gap-1.5 cursor-pointer transition whitespace-nowrap"
          >
            <Eye className="w-3.5 h-3.5" />
            Pratinjau Modal
          </button>
          <button
            type="button"
            disabled={isDownloading === "pdf"}
            onClick={async () => {
              setIsDownloading("pdf");
              await exportToPDF(
                "print-area-master-menu",
                `Master_Menu_${totalDays}_Hari_${activePlan?.title || "SPPG"}`,
                100,
                undefined,
                { paperSize, orientation: "auto" }
              );
              setIsDownloading(null);
            }}
            className="px-3 py-2 bg-rose-600 hover:bg-rose-700 disabled:opacity-50 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 cursor-pointer transition whitespace-nowrap"
          >
            <Download className="w-3.5 h-3.5" />
            {isDownloading === "pdf" ? "Memproses..." : "Ekspor PDF"}
          </button>
          <button
            type="button"
            disabled={isDownloading === "png"}
            onClick={async () => {
              setIsDownloading("png");
              await downloadElementAsImage(
                "print-area-master-menu",
                `Master_Menu_${totalDays}_Hari_${activePlan?.title || "SPPG"}`
              );
              setIsDownloading(null);
            }}
            className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 cursor-pointer transition whitespace-nowrap"
          >
            <ImageIcon className="w-3.5 h-3.5" />
            {isDownloading === "png" ? "Memproses..." : "Unduh PNG"}
          </button>
          <button
            type="button"
            onClick={() =>
              printElementSmart("print-area-master-menu", {
                title: masterBannerHeadline,
                paperSize,
                orientation: "auto",
              })
            }
            className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 cursor-pointer transition whitespace-nowrap"
          >
            <Printer className="w-3.5 h-3.5" />
            Cetak
          </button>
        </div>
      </div>

      {/* SECTION 4A: CARD INPUT VIEW (for detailed card editing + explicit Nama Menu Food Cost) */}
      {viewMode === "cards" && activePlan && (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4 no-print">
          {plannedDates.map((dateStr, idx) => {
            const item = getDayMenuItem(idx);
            const autoComposed = composeMenuName(item);
            return (
              <div
                key={idx}
                className="bg-white rounded-2xl border border-slate-200 p-4 space-y-3 flex flex-col justify-between"
              >
                <div className="space-y-3">
                  <div className="flex items-start justify-between gap-2 pb-2.5 border-b border-slate-100">
                    <div>
                      <span className="text-[11px] font-extrabold text-emerald-700 font-mono block">
                        MENU {idx + 1} · {formatIndonesianDate(dateStr)}
                      </span>
                      <span className="text-xs font-bold text-slate-800 block mt-0.5 truncate max-w-[230px]">
                        {autoComposed || "Belum diisi"}
                      </span>
                    </div>
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        onClick={() => handleCopyDay(idx)}
                        className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-600 cursor-pointer"
                        title="Salin menu hari ini"
                      >
                        <Copy className="w-3.5 h-3.5" />
                      </button>
                      {copiedItem && (
                        <button
                          type="button"
                          onClick={() => handlePasteDay(idx)}
                          className={`p-1.5 rounded-lg cursor-pointer ${
                            justPastedDay === idx
                              ? "bg-emerald-100 text-emerald-700"
                              : "hover:bg-indigo-50 text-indigo-600"
                          }`}
                          title="Tempel menu tersalin"
                        >
                          {justPastedDay === idx ? (
                            <Check className="w-3.5 h-3.5" />
                          ) : (
                            <Clipboard className="w-3.5 h-3.5" />
                          )}
                        </button>
                      )}
                      <button
                        type="button"
                        onClick={() => handleResetDay(idx)}
                        className="p-1.5 rounded-lg hover:bg-rose-50 text-rose-500 cursor-pointer"
                        title="Kosongkan hari ini"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>

                  <div className="space-y-2 text-xs">
                    <div>
                      <label className="text-[10px] font-bold text-indigo-700 block mb-0.5">
                        Nama Menu (Otomatis ke Panel Food Cost)
                      </label>
                      <input
                        type="text"
                        value={item.namaMenu}
                        onChange={(e) => handleDayFieldChange(idx, "namaMenu", e.target.value)}
                        placeholder={autoComposed || "Otomatis dari gabungan komponen"}
                        className="w-full px-2.5 py-1.5 bg-indigo-50/40 border border-indigo-200 rounded-lg text-xs font-bold text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                      />
                    </div>

                    {MENU_COMPONENT_ROWS.map((row) => (
                      <div key={row.key} className="grid grid-cols-12 items-center gap-2">
                        <span className="col-span-4 text-[11px] font-semibold text-slate-600">
                          {row.label}
                        </span>
                        <input
                          type="text"
                          value={item[row.key] || ""}
                          onChange={(e) => handleDayFieldChange(idx, row.key, e.target.value)}
                          placeholder={row.placeholder}
                          className="col-span-8 px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 focus:bg-white"
                        />
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* SECTION 4B: TABEL MASTER MENU RESMI (GAMBAR 1 - ALWAYS RENDERED FOR PRINT / TABLE EDIT / PRINT PREVIEW) */}
      <div
        className={`${
          viewMode === "cards" ? "hidden print:block" : "block"
        } bg-slate-100/80 p-3 sm:p-5 rounded-2xl border border-slate-300 overflow-x-auto print:bg-white print:p-0 print:border-none`}
      >
        <div
          id="print-area-master-menu"
          className="bg-white p-4 sm:p-6 mx-auto border border-slate-300 shadow-sm min-w-[980px] max-w-[1400px] text-black print:shadow-none print:border-none print:p-0 print:max-w-full"
          style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
        >
          {/* KOP SURAT TERPUSAT SPPG */}
          <KopSuratRenderHeader
            kopLine1={kopLine1 || "BADAN GIZI NASIONAL"}
            kopLine2={kopLine2 || "SATUAN PELAYANAN PEMENUHAN GIZI"}
            kopLine3={kopLine3 || profile?.namaLembaga || "SPPG"}
            kopLine4={kopLine4 || (profile?.alamat ? `Alamat : ${profile.alamat}` : "")}
            leftLogo={activeKopLeftSrc}
            rightLogo={rightLogo}
            leftLogoCrop={leftLogoCrop}
            rightLogoCrop={rightLogoCrop}
            kopConfig={kopConfig}
            docKey="menu"
          />

          {/* FULL-WIDTH GREEN TITLE BANNER (GAMBAR 1) */}
          <table className="w-full border-collapse border border-black text-[10.5pt]">
            <thead>
              <tr>
                <th
                  className="border border-black py-1.5 px-2 text-center font-extrabold uppercase text-black"
                  style={{ backgroundColor: "#00B050", fontSize: "14pt" }}
                >
                  {masterBannerHeadline}
                </th>
              </tr>
            </thead>
          </table>

          {/* WEEKLY BLOCKS MATCHING GAMBAR 1 */}
          {weekBlocks.map((week) => {
            const colWidthPct = `${84 / Math.max(week.days.length, 1)}%`;
            return (
              <table
                key={week.weekNumber}
                className="w-full border-collapse border border-black table-fixed -mt-[1px] break-inside-avoid"
                style={{ fontSize: "10pt" }}
              >
                <thead>
                  {/* Yellow Header Row: Kelompok Bahan Makanan + Calendar Dates */}
                  <tr style={{ backgroundColor: "#FFFF00" }}>
                    <th
                      className="border border-black py-1.5 px-2 text-center font-bold text-black align-middle"
                      style={{ backgroundColor: "#FFFF00", width: "16%" }}
                    >
                      {week.weekNumber === 1
                        ? "Kelompok Bahan Makanan"
                        : "Kelompok bahan makanan"}
                    </th>
                    {week.days.map((d) => (
                      <th
                        key={d.dayIndex}
                        className="border border-black py-1.5 px-1.5 text-center font-bold text-black align-middle"
                        style={{ backgroundColor: "#FFFF00", width: colWidthPct }}
                      >
                        {formatIndonesianDate(d.dateStr) || `Hari ${d.menuNumber}`}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {/* 6 Component Rows matching Gambar 1 */}
                  {MENU_COMPONENT_ROWS.map((comp) => (
                    <tr key={comp.key} className="bg-white">
                      <td className="border border-black py-1 px-2 text-left font-bold text-black align-middle bg-white">
                        {comp.label}
                      </td>
                      {week.days.map((d) => {
                        const dayItem = getDayMenuItem(d.dayIndex);
                        const val = dayItem[comp.key] || "";
                        return (
                          <td
                            key={d.dayIndex}
                            className="border border-black p-0.5 text-center align-middle bg-white"
                          >
                            {viewMode === "print" ? (
                              <div className="px-1 py-1 text-center font-medium text-black leading-tight min-h-[26px] flex items-center justify-center">
                                {val}
                              </div>
                            ) : (
                              <textarea
                                rows={val.length > 26 ? 2 : 1}
                                value={val}
                                onChange={(e) =>
                                  handleDayFieldChange(d.dayIndex, comp.key, e.target.value)
                                }
                                placeholder="-"
                                className="w-full text-center bg-transparent border-0 focus:ring-1 focus:ring-emerald-600 focus:bg-emerald-50/30 rounded-xs px-1 py-0.5 text-black font-medium leading-tight resize-none overflow-hidden"
                                style={{ fontSize: "10pt" }}
                              />
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  ))}

                  {/* Optional Inline Nama Menu row (hidden when printing so printed table matches Gambar 1) */}
                  {viewMode === "table" && (
                    <tr className="bg-slate-50 no-print">
                      <td className="border border-black py-1 px-2 text-left font-bold text-indigo-800 text-[9pt] align-middle">
                        Nama Menu (Food Cost)
                      </td>
                      {week.days.map((d) => {
                        const dayItem = getDayMenuItem(d.dayIndex);
                        const autoComposed = composeMenuName(dayItem);
                        return (
                          <td key={d.dayIndex} className="border border-black p-0.5 align-middle">
                            <input
                              type="text"
                              value={dayItem.namaMenu || ""}
                              onChange={(e) =>
                                handleDayFieldChange(d.dayIndex, "namaMenu", e.target.value)
                              }
                              placeholder={autoComposed || `Menu ${d.menuNumber}`}
                              title="Opsional: isi judul menu khusus atau biarkan kosong untuk otomatis dari komponen di atas"
                              className="w-full text-center bg-white/80 border-0 focus:ring-1 focus:ring-indigo-500 px-1 py-0.5 text-[9pt] font-semibold text-indigo-900 placeholder:text-slate-400"
                            />
                          </td>
                        );
                      })}
                    </tr>
                  )}

                  {/* Green Summary Footer Row: Minggu Ke- X | MENU 1..N */}
                  <tr style={{ backgroundColor: "#00B050" }}>
                    <td
                      className="border border-black py-1 px-2 text-center font-bold text-black align-middle"
                      style={{ backgroundColor: "#00B050" }}
                    >
                      Minggu Ke- {week.weekNumber}
                    </td>
                    {week.days.map((d) => (
                      <td
                        key={d.dayIndex}
                        className="border border-black py-1 px-1 text-center font-bold text-black align-middle font-mono"
                        style={{ backgroundColor: "#00B050" }}
                      >
                        MENU {d.menuNumber}
                      </td>
                    ))}
                  </tr>
                </tbody>
              </table>
            );
          })}
          <AutoSignatureBlock profile={profile} docKey="menu" />
        </div>
      </div>

      {/* Confirmation Modal (in-app, zero window.confirm) */}
      {confirmModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-50 p-4 no-print">
          <div className="bg-white rounded-2xl p-6 max-w-md w-full border border-slate-200 shadow-xl space-y-4">
            <h3 className="font-bold text-slate-900 text-base">{confirmModal.title}</h3>
            <p className="text-xs text-slate-600 leading-relaxed">{confirmModal.message}</p>
            <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setConfirmModal(null)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={confirmModal.onConfirm}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold rounded-xl cursor-pointer"
              >
                {confirmModal.confirmLabel}
              </button>
            </div>
          </div>
        </div>
      )}

      <PrintPreviewModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title={`Pratinjau Cetak: ${masterBannerHeadline}`}
        elementId="print-area-master-menu"
        filename={`Master_Menu_${totalDays}_Hari_${activePlan?.title || "SPPG"}`}
        defaultScale={100}
        defaultPaperSize={paperSize}
        defaultOrientation="auto"
      />
    </div>
  );
}
