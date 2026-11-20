/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useMemo, useState } from "react";
import {
  ClipboardCheck,
  TestTube2,
  ThermometerSnowflake,
  Printer,
  FileDown,
  Download,
  Eye,
  Plus,
  Trash2,
  RotateCcw,
  Users,
  Check,
  Calendar,
  X,
  Archive,
  ArrowRight,
  CheckCircle2,
  Layers,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  DEFAULT_UNIFIED_KOP_CONFIG,
  FoodCostDay,
  MasterMenu,
  MasterMenuPlan,
  SPPGProfile,
  SavedDocArchive,
  SavedDocArchiveWorksheet,
  SopDataState,
  SopOrganoleptikDayRecord,
  SopOrganoleptikDishRow,
  SopPenyimpananRowRecord,
  SopSampelOfficerTeam,
  SopSampelRowRecord,
  UnifiedKopConfig,
} from "../types";
import {
  composeMenuName,
  ensureMasterMenuPlans,
  formatIndonesianDate,
  isWorkDay,
  parseLocalDate,
  toYyyyMmDd,
} from "../utils/menuSchedule";
import { printElementSmart, exportToPDF } from "../lib/printUtils";
import { KopSuratRenderHeader, LogoCrop } from "./KopSuratConfigSection";
import AutoSignatureBlock, {
  IdSppgFooterStrip,
  resolveSignatureConfig,
} from "./AutoSignatureBlock";
import PrintPreviewModal from "./PrintPreviewModal";

type SopSubTabType = "organoleptik" | "sampel" | "penyimpanan";

interface FormSopTabProps {
  profile: SPPGProfile;
  masterMenu: MasterMenu;
  foodCostDays: FoodCostDay[];
  sopData: SopDataState;
  setSopData: React.Dispatch<React.SetStateAction<SopDataState>>;
  kopLine1: string;
  kopLine2: string;
  kopLine3: string;
  kopLine4: string;
  leftLogo: string;
  rightLogo: string;
  leftLogoCrop?: LogoCrop;
  rightLogoCrop?: LogoCrop;
  kopConfig?: UnifiedKopConfig;
  paperSize?: "A4" | "F4";
  onSaveToArchive?: (archive: SavedDocArchive) => void;
  onGoToArsip?: () => void;
}

const INDONESIAN_MONTHS = [
  "Januari",
  "Februari",
  "Maret",
  "April",
  "Mei",
  "Juni",
  "Juli",
  "Agustus",
  "September",
  "Oktober",
  "November",
  "Desember",
];

const ROMAN_MONTHS = [
  "I",
  "II",
  "III",
  "IV",
  "V",
  "VI",
  "VII",
  "VIII",
  "IX",
  "X",
  "XI",
  "XII",
];

const DEFAULT_OFFICER_TEAMS: SopSampelOfficerTeam[] = [
  {
    id: "team_1",
    label: "Tim Pengantaran 1",
    members: ["Petugas 1", "Petugas 2", "Petugas 3"],
  },
  {
    id: "team_2",
    label: "Tim Pengantaran 2",
    members: ["Petugas A", "Petugas B", "Petugas C"],
  },
];

const VALID_SOP_TABS: SopSubTabType[] = ["organoleptik", "sampel", "penyimpanan"];

const SOP_TAB_LABELS: Record<SopSubTabType, string> = {
  organoleptik: "Uji Organoleptik (Lamp. 3)",
  sampel: "Pengambilan Sampel (Lamp. 2)",
  penyimpanan: "Penyimpanan Sampel (Chiller)",
};

function formatShortDateId(dateStr: string): string {
  const dt = parseLocalDate(dateStr);
  if (!dt) return dateStr;
  const dd = String(dt.getDate()).padStart(2, "0");
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const yyyy = dt.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function formatCompactTabDate(dateStr: string): string {
  const dt = parseLocalDate(dateStr);
  if (!dt) return dateStr;
  const days = ["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"];
  const months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  return `${days[dt.getDay()]}, ${String(dt.getDate()).padStart(2, "0")} ${months[dt.getMonth()]}`;
}

function addDaysToYyyyMmDd(dateStr: string, daysToAdd: number): string {
  const dt = parseLocalDate(dateStr);
  if (!dt) return dateStr;
  const next = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate() + daysToAdd);
  return toYyyyMmDd(next);
}

function buildMonthWorkDates(ym: string, workDaysPerWeek: 5 | 6 | 7): string[] {
  const match = /^(\d{4})-(\d{2})$/.exec((ym || "").trim());
  if (!match) return [];
  const year = Number(match[1]);
  const monthIdx = Number(match[2]) - 1;
  const lastDay = new Date(year, monthIdx + 1, 0).getDate();
  const out: string[] = [];
  for (let d = 1; d <= lastDay; d++) {
    const dt = new Date(year, monthIdx, d);
    if (isWorkDay(dt, workDaysPerWeek)) {
      out.push(toYyyyMmDd(dt));
    }
  }
  return out;
}

function formatTeamMembers(members: string[]): string {
  return members
    .map((m) => m.trim())
    .filter(Boolean)
    .map((m, idx) => `${idx + 1}. ${m}`)
    .join("\n");
}

function generateSampleCode(dateStr: string, menuName: string): string {
  const dt = parseLocalDate(dateStr);
  const codeDate = dt
    ? `${String(dt.getDate()).padStart(2, "0")}${String(dt.getMonth() + 1).padStart(2, "0")}${String(dt.getFullYear()).slice(-2)}`
    : dateStr;
  return menuName ? `SM-${codeDate} · ${menuName}` : `SM-${codeDate}`;
}

function scopedSopKey(planId: string, dateStr: string): string {
  return `${planId}__${dateStr}`;
}

export default function FormSopTab({
  profile,
  masterMenu,
  foodCostDays,
  sopData,
  setSopData,
  kopLine1,
  kopLine2,
  kopLine3,
  kopLine4,
  leftLogo,
  rightLogo,
  leftLogoCrop,
  rightLogoCrop,
  kopConfig = DEFAULT_UNIFIED_KOP_CONFIG,
  paperSize = "A4",
  onSaveToArchive,
  onGoToArsip,
}: FormSopTabProps) {
  const [selectedDate, setSelectedDate] = useState<string>("");
  const [organoPrintAll, setOrganoPrintAll] = useState<boolean>(false);
  const [weekFilter, setWeekFilter] = useState<string>("all");
  const [showTeamEditor, setShowTeamEditor] = useState<boolean>(false);
  const [isDownloading, setIsDownloading] = useState<string | null>(null);
  const [showPrintModal, setShowPrintModal] = useState<boolean>(false);
  const [confirmResetModal, setConfirmResetModal] = useState<boolean>(false);
  const [savedArchiveNotice, setSavedArchiveNotice] = useState<string | null>(null);

  const normalizedMenu = useMemo(() => ensureMasterMenuPlans(masterMenu), [masterMenu]);
  const plans = normalizedMenu.plans || [];

  const defaultMonth = useMemo(() => {
    if (sopData.activeMonth && /^\d{4}-\d{2}$/.test(sopData.activeMonth)) {
      return sopData.activeMonth;
    }
    if (normalizedMenu.scheduleConfig?.calendarMonth) {
      return normalizedMenu.scheduleConfig.calendarMonth;
    }
    if (profile.periodeDates?.[0] && /^\d{4}-\d{2}/.test(profile.periodeDates[0])) {
      return profile.periodeDates[0].slice(0, 7);
    }
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  }, [sopData.activeMonth, normalizedMenu.scheduleConfig?.calendarMonth, profile.periodeDates]);

  const selectedSubTabs = useMemo<SopSubTabType[]>(() => {
    const raw = (sopData.selectedSubTabs || []).filter((t): t is SopSubTabType =>
      VALID_SOP_TABS.includes(t as SopSubTabType)
    );
    if (raw.length > 0) return Array.from(new Set(raw));
    const single = sopData.activeSubTab;
    return [VALID_SOP_TABS.includes(single as SopSubTabType) ? (single as SopSubTabType) : "organoleptik"];
  }, [sopData.selectedSubTabs, sopData.activeSubTab]);

  const activeSubTab = selectedSubTabs[0];

  const planMap = useMemo(() => {
    const m = new Map<string, MasterMenuPlan>();
    plans.forEach((p) => m.set(p.id, p));
    return m;
  }, [plans]);

  const activePlanId = sopData.selectedPlanId || normalizedMenu.activePlanId || plans[0]?.id || "";

  const selectedPlanIds = useMemo(() => {
    const raw = (sopData.selectedPlanIds || []).filter((id) => planMap.has(id));
    if (raw.length > 0) return Array.from(new Set(raw));
    if (activePlanId && planMap.has(activePlanId)) return [activePlanId];
    return plans[0] ? [plans[0].id] : [];
  }, [sopData.selectedPlanIds, activePlanId, planMap, plans]);

  const selectedPlans = useMemo(
    () => selectedPlanIds.map((id) => planMap.get(id)).filter((p): p is MasterMenuPlan => Boolean(p)),
    [selectedPlanIds, planMap]
  );

  const primaryPlanId = selectedPlans[0]?.id || activePlanId;

  const dateMode = sopData.dateMode || "bulan_aktif";
  const workDaysPerWeek: 5 | 6 | 7 =
    sopData.workDaysPerWeek || normalizedMenu.scheduleConfig?.workDaysPerWeek || 5;
  const showKop = (sopData.showKop ?? true) && kopConfig?.linkedDocs?.sop !== false;
  const showDigitalTtd = sopData.showDigitalTtd ?? true;
  const organoSigMode = sopData.organoleptikSignatureMode || "petugas";
  const penyimpananSigMode = sopData.penyimpananSignatureMode || "kotak_verifikasi";
  const retentionDays = sopData.penyimpananRetentionDays ?? 2;
  const defaultJumlahSampel = sopData.penyimpananDefaultJumlah ?? "1 Porsi";
  const defaultJamMasuk = sopData.penyimpananDefaultJamMasuk ?? "07:30";
  const defaultJamKeluar = sopData.penyimpananDefaultJamKeluar ?? "07:30";
  const officerTeams =
    sopData.sampelTeams && sopData.sampelTeams.length > 0
      ? sopData.sampelTeams
      : DEFAULT_OFFICER_TEAMS;

  const ttdConfig = resolveSignatureConfig(profile);

  const updateSop = (patch: Partial<SopDataState>) => {
    setSopData((prev) => ({ ...prev, ...patch }));
  };

  const toggleSubTab = (tab: SopSubTabType) => {
    const exists = selectedSubTabs.includes(tab);
    let next: SopSubTabType[];
    if (exists) {
      if (selectedSubTabs.length <= 1) return;
      next = selectedSubTabs.filter((t) => t !== tab);
    } else {
      next = VALID_SOP_TABS.filter((t) => t === tab || selectedSubTabs.includes(t));
    }
    updateSop({
      selectedSubTabs: next,
      activeSubTab: next[0],
    });
  };

  const selectAllSubTabs = () => {
    updateSop({
      selectedSubTabs: [...VALID_SOP_TABS],
      activeSubTab: "organoleptik",
    });
  };

  const selectSingleSubTab = (tab: SopSubTabType) => {
    updateSop({
      selectedSubTabs: [tab],
      activeSubTab: tab,
    });
  };

  const togglePlanSelection = (planId: string) => {
    const exists = selectedPlanIds.includes(planId);
    let next: string[];
    if (exists) {
      if (selectedPlanIds.length <= 1) return;
      next = selectedPlanIds.filter((id) => id !== planId);
    } else {
      next = plans.map((p) => p.id).filter((id) => id === planId || selectedPlanIds.includes(id));
    }
    updateSop({
      selectedPlanIds: next,
      selectedPlanId: next[0],
    });
  };

  const selectAllPlans = () => {
    const allIds = plans.map((p) => p.id);
    if (allIds.length === 0) return;
    updateSop({
      selectedPlanIds: allIds,
      selectedPlanId: allIds[0],
    });
  };

  const selectSinglePlan = (planId: string) => {
    updateSop({
      selectedPlanIds: [planId],
      selectedPlanId: planId,
    });
  };

  // Parse activeMonth year & Indonesian month label
  const { activeYear, activeMonthName, activeMonthRoman } = useMemo(() => {
    const match = /^(\d{4})-(\d{2})$/.exec(defaultMonth);
    const y = match ? match[1] : profile.tahunAnggaran || String(new Date().getFullYear());
    const idx = match ? Math.max(0, Math.min(11, Number(match[2]) - 1)) : new Date().getMonth();
    return {
      activeYear: y,
      activeMonthIndex: idx,
      activeMonthName: INDONESIAN_MONTHS[idx],
      activeMonthRoman: ROMAN_MONTHS[idx],
    };
  }, [defaultMonth, profile.tahunAnggaran]);

  // Scheduled cycle dates from Master Menu / Profil SPPG
  const cycleDates = useMemo(() => {
    const fromSchedule = normalizedMenu.scheduleConfig?.periodeDates || [];
    if (fromSchedule.length > 0) return fromSchedule;
    if (profile.periodeDates && profile.periodeDates.length > 0) return profile.periodeDates;
    return buildMonthWorkDates(defaultMonth, workDaysPerWeek);
  }, [normalizedMenu.scheduleConfig?.periodeDates, profile.periodeDates, defaultMonth, workDaysPerWeek]);

  // Final list of active dates depending on dateMode
  const activeDates = useMemo(() => {
    if (dateMode === "siklus") {
      return cycleDates;
    }
    const monthDates = buildMonthWorkDates(defaultMonth, workDaysPerWeek);
    return monthDates.length > 0 ? monthDates : cycleDates;
  }, [dateMode, defaultMonth, workDaysPerWeek, cycleDates]);

  // Active selected date for single-day Organoleptik view
  const effectiveOrganoDate = useMemo(() => {
    if (selectedDate && activeDates.includes(selectedDate)) return selectedDate;
    return activeDates[0] || toYyyyMmDd(new Date());
  }, [selectedDate, activeDates]);

  // Resolve auto menu dishes & full menu name per Plan & Date
  const resolvedMenuByPlanAndDate = useMemo(() => {
    const result: Record<string, Record<string, { dishes: string[]; fullName: string; hariKe: number }>> = {};

    selectedPlans.forEach((plan) => {
      const map: Record<string, { dishes: string[]; fullName: string; hariKe: number }> = {};
      const planDays = plan.days || [];
      const fallbackDays =
        plan.id === plans[0]?.id ? normalizedMenu.usiaSekolah || [] : [];
      const isAlergi = (plan.dietType || "").toLowerCase().includes("alergi");
      const isMpAsi =
        (plan.dietType || "").toLowerCase().includes("mp-asi") ||
        (plan.dietType || "").toLowerCase().includes("balita") ||
        (plan.title || "").toLowerCase().includes("balita");

      activeDates.forEach((dateStr, idx) => {
        const exactCycleIdx = cycleDates.indexOf(dateStr);
        const dayIdx = exactCycleIdx >= 0 ? exactCycleIdx : idx;
        const hariKe =
          planDays.length > 0 ? (dayIdx % planDays.length) + 1 : dayIdx + 1;

        const menuItem =
          planDays.length > 0
            ? planDays[dayIdx % planDays.length]
            : fallbackDays.length > 0
            ? fallbackDays[dayIdx % fallbackDays.length]
            : undefined;

        const dishes: string[] = [];
        if (menuItem) {
          const candidates = [
            menuItem.karbohidrat,
            menuItem.laukHewani,
            menuItem.laukNabati,
            menuItem.sayur,
            menuItem.buah,
            menuItem.susu,
          ]
            .map((s) => (s || "").trim())
            .filter((s) => s && s !== "-");

          if (candidates.length > 0) {
            dishes.push(...candidates);
          } else if (menuItem.buahSusu && menuItem.buahSusu.trim() && menuItem.buahSusu.trim() !== "-") {
            dishes.push(menuItem.buahSusu.trim());
          } else if (menuItem.namaMenu && menuItem.namaMenu.trim()) {
            const split = menuItem.namaMenu
              .split(/\+|,/)
              .map((p) => p.trim())
              .filter(Boolean);
            dishes.push(...split);
          }
        }

        let fullName = menuItem ? composeMenuName(menuItem).trim() : "";

        // Fallback to FoodCostDay for matching hariKe if Master Menu is empty
        if (dishes.length === 0 || !fullName) {
          const fcMatch =
            (isAlergi
              ? foodCostDays.find((fc) => fc.hariKe === hariKe && fc.jenisMenu === "Alergi")
              : undefined) ||
            (isMpAsi
              ? foodCostDays.find((fc) => fc.hariKe === hariKe && fc.jenisMenu === "MP-ASI")
              : undefined) ||
            foodCostDays.find((fc) => fc.hariKe === hariKe && fc.jenisMenu === "Basah") ||
            foodCostDays.find((fc) => fc.hariKe === hariKe);

          if (fcMatch) {
            if (!fullName && fcMatch.namaMenu?.trim()) {
              fullName = fcMatch.namaMenu.trim();
            }
            if (dishes.length === 0) {
              if (fcMatch.namaMenu?.trim()) {
                dishes.push(
                  ...fcMatch.namaMenu
                    .split(/\+|,/)
                    .map((p) => p.trim())
                    .filter(Boolean)
                );
              } else {
                const bahanNames = (fcMatch.porsiBesarBahan || fcMatch.porsiKecilBahan || [])
                  .map((b) => (b.nama || "").trim())
                  .filter((n) => n && n.toLowerCase() !== "ketik bahan...")
                  .slice(0, 6);
                if (bahanNames.length > 0) {
                  dishes.push(...bahanNames);
                  if (!fullName) fullName = bahanNames.join(", ");
                }
              }
            }
          }
        }

        map[dateStr] = {
          dishes,
          fullName,
          hariKe,
        };
      });

      result[plan.id] = map;
    });

    return result;
  }, [selectedPlans, plans, activeDates, cycleDates, normalizedMenu.usiaSekolah, foodCostDays]);

  // Week-filtered dates for multi-row tables (Sampel & Penyimpanan)
  const filteredMultiDates = useMemo(() => {
    if (weekFilter === "all") return activeDates;
    const wk = Number(weekFilter);
    if (!Number.isFinite(wk) || wk < 1) return activeDates;
    return activeDates.filter((dateStr) => {
      const dt = parseLocalDate(dateStr);
      if (!dt) return true;
      const weekOfMonth = Math.ceil(dt.getDate() / 7);
      return weekOfMonth === wk;
    });
  }, [activeDates, weekFilter]);

  // --- ORGANOLEPTIK HELPERS (PLAN-SCOPED) ---
  const getOrganoDayRecord = (planId: string, dateStr: string): SopOrganoleptikDayRecord => {
    const sKey = scopedSopKey(planId, dateStr);
    const saved =
      sopData.organoleptikByDate?.[sKey] ??
      (planId === primaryPlanId ? sopData.organoleptikByDate?.[dateStr] : undefined);
    const autoMenu = resolvedMenuByPlanAndDate[planId]?.[dateStr];
    const autoDishes = autoMenu?.dishes || [];
    const baseMeta = sopData.organoleptikMeta || {};

    let rows: SopOrganoleptikDishRow[];
    if (saved?.rows && saved.rows.length > 0) {
      rows = saved.rows;
    } else {
      const minRows = Math.max(5, autoDishes.length);
      rows = Array.from({ length: minRows }, (_, i) => ({
        id: `org_${planId}_${dateStr}_${i + 1}`,
        namaHidangan: autoDishes[i] || "",
        rasa: "",
        warna: "",
        aroma: "",
        tekstur: "",
        waktuUji: "",
        ket: "",
      }));
    }

    return {
      namaPemeriksa: saved?.namaPemeriksa ?? baseMeta.namaPemeriksa ?? profile.namaAhliGizi ?? "",
      tempatPemeriksaan:
        saved?.tempatPemeriksaan ?? baseMeta.tempatPemeriksaan ?? profile.namaLembaga ?? "",
      namaTempatPemeriksa:
        saved?.namaTempatPemeriksa ?? baseMeta.namaTempatPemeriksa ?? "Ruang Ahli Gizi",
      waktuJam: saved?.waktuJam ?? baseMeta.waktuDefault ?? "",
      zonaWaktu: saved?.zonaWaktu ?? baseMeta.zonaWaktu ?? "WITA",
      kesimpulan: saved?.kesimpulan ?? "belum",
      rows,
    };
  };

  const updateOrganoDay = (
    planId: string,
    dateStr: string,
    updater: (current: SopOrganoleptikDayRecord) => SopOrganoleptikDayRecord
  ) => {
    const current = getOrganoDayRecord(planId, dateStr);
    const nextRec = updater(current);
    const sKey = scopedSopKey(planId, dateStr);
    setSopData((prev) => ({
      ...prev,
      organoleptikByDate: {
        ...(prev.organoleptikByDate || {}),
        [sKey]: nextRec,
      },
    }));
  };

  const updateOrganoRow = (
    planId: string,
    dateStr: string,
    rowIdx: number,
    patch: Partial<SopOrganoleptikDishRow>
  ) => {
    updateOrganoDay(planId, dateStr, (rec) => {
      const nextRows = [...(rec.rows || [])];
      nextRows[rowIdx] = { ...nextRows[rowIdx], ...patch };
      return { ...rec, rows: nextRows };
    });
  };

  const addOrganoRow = (planId: string, dateStr: string) => {
    updateOrganoDay(planId, dateStr, (rec) => ({
      ...rec,
      rows: [
        ...(rec.rows || []),
        {
          id: `org_${planId}_${dateStr}_${Date.now()}`,
          namaHidangan: "",
          rasa: "",
          warna: "",
          aroma: "",
          tekstur: "",
          waktuUji: "",
          ket: "",
        },
      ],
    }));
  };

  const removeOrganoRow = (planId: string, dateStr: string, rowIdx: number) => {
    updateOrganoDay(planId, dateStr, (rec) => {
      const nextRows = (rec.rows || []).filter((_, i) => i !== rowIdx);
      return { ...rec, rows: nextRows.length > 0 ? nextRows : rec.rows };
    });
  };

  const applyQuickOrganoScore = (planId: string, dateStr: string, score: string) => {
    updateOrganoDay(planId, dateStr, (rec) => ({
      ...rec,
      kesimpulan: score === "" ? "belum" : Number(score) >= 4 ? "aman" : rec.kesimpulan,
      rows: (rec.rows || []).map((r) =>
        r.namaHidangan.trim()
          ? {
              ...r,
              rasa: score,
              warna: score,
              aroma: score,
              tekstur: score,
              waktuUji: score ? r.waktuUji || "Sebelum Pengantaran" : "",
              ket: score ? r.ket || (Number(score) >= 4 ? "Layak" : "") : "",
            }
          : { ...r, rasa: "", warna: "", aroma: "", tekstur: "", waktuUji: "", ket: "" }
      ),
    }));
  };

  // --- SAMPEL MAKANAN HELPERS (PLAN-SCOPED) ---
  const getSampelRow = (planId: string, dateStr: string, idx: number): SopSampelRowRecord => {
    const sKey = scopedSopKey(planId, dateStr);
    const saved =
      sopData.sampelByDate?.[sKey] ??
      (planId === primaryPlanId ? sopData.sampelByDate?.[dateStr] : undefined);
    const autoMenu = resolvedMenuByPlanAndDate[planId]?.[dateStr]?.fullName || "";
    const defaultTeam = officerTeams[idx % officerTeams.length];
    return {
      pengantaran: saved?.pengantaran ?? "",
      namaPetugas: saved?.namaPetugas ?? (defaultTeam ? formatTeamMembers(defaultTeam.members) : ""),
      namaMakanan: saved?.namaMakanan ?? autoMenu,
      parafChecked: saved?.parafChecked ?? false,
    };
  };

  const updateSampelRow = (
    planId: string,
    dateStr: string,
    idx: number,
    patch: Partial<SopSampelRowRecord>
  ) => {
    const current = getSampelRow(planId, dateStr, idx);
    const sKey = scopedSopKey(planId, dateStr);
    setSopData((prev) => ({
      ...prev,
      sampelByDate: {
        ...(prev.sampelByDate || {}),
        [sKey]: { ...current, ...patch },
      },
    }));
  };

  const applyTeamsAcrossMonth = () => {
    if (officerTeams.length === 0) return;
    const nextMap: Record<string, SopSampelRowRecord> = {
      ...(sopData.sampelByDate || {}),
    };
    selectedPlans.forEach((plan) => {
      activeDates.forEach((dateStr, idx) => {
        const sKey = scopedSopKey(plan.id, dateStr);
        const existing = nextMap[sKey] || (plan.id === primaryPlanId ? nextMap[dateStr] : undefined) || {};
        const team = officerTeams[idx % officerTeams.length];
        nextMap[sKey] = {
          ...existing,
          namaPetugas: formatTeamMembers(team.members),
        };
      });
    });
    updateSop({ sampelByDate: nextMap });
  };

  // --- PENYIMPANAN SAMPEL (CHILLER) HELPERS (PLAN-SCOPED) ---
  const getPenyimpananRow = (planId: string, dateStr: string): SopPenyimpananRowRecord => {
    const sKey = scopedSopKey(planId, dateStr);
    const saved =
      sopData.penyimpananByDate?.[sKey] ??
      (planId === primaryPlanId ? sopData.penyimpananByDate?.[dateStr] : undefined);
    const autoMenu = resolvedMenuByPlanAndDate[planId]?.[dateStr]?.fullName || "";
    const autoKeluarDate = addDaysToYyyyMmDd(dateStr, retentionDays);
    return {
      kodeIdentitas: saved?.kodeIdentitas ?? generateSampleCode(dateStr, autoMenu),
      jumlahSampel: saved?.jumlahSampel ?? defaultJumlahSampel,
      kondisi: saved?.kondisi ?? "",
      namaPenanggungJawab: saved?.namaPenanggungJawab ?? profile.namaAhliGizi ?? "",
      masukTanggal: saved?.masukTanggal ?? formatShortDateId(dateStr),
      masukJam: saved?.masukJam ?? defaultJamMasuk,
      keluarTanggal: saved?.keluarTanggal ?? formatShortDateId(autoKeluarDate),
      keluarJam: saved?.keluarJam ?? defaultJamKeluar,
    };
  };

  const updatePenyimpananRow = (
    planId: string,
    dateStr: string,
    patch: Partial<SopPenyimpananRowRecord>
  ) => {
    const current = getPenyimpananRow(planId, dateStr);
    const sKey = scopedSopKey(planId, dateStr);
    setSopData((prev) => ({
      ...prev,
      penyimpananByDate: {
        ...(prev.penyimpananByDate || {}),
        [sKey]: { ...current, ...patch },
      },
    }));
  };

  const applyAllPenyimpananKondisi = (kondisi: "OK" | "") => {
    const nextMap: Record<string, SopPenyimpananRowRecord> = {
      ...(sopData.penyimpananByDate || {}),
    };
    selectedPlans.forEach((plan) => {
      filteredMultiDates.forEach((dateStr) => {
        const sKey = scopedSopKey(plan.id, dateStr);
        const current = getPenyimpananRow(plan.id, dateStr);
        nextMap[sKey] = { ...current, kondisi };
      });
    });
    updateSop({ penyimpananByDate: nextMap });
  };

  // --- RESET / RESYNC AUTO DATA ---
  const handleConfirmResetCurrentTab = () => {
    const patch: Partial<SopDataState> = {};
    if (selectedSubTabs.includes("organoleptik")) patch.organoleptikByDate = {};
    if (selectedSubTabs.includes("sampel")) patch.sampelByDate = {};
    if (selectedSubTabs.includes("penyimpanan")) patch.penyimpananByDate = {};
    updateSop(patch);
    setConfirmResetModal(false);
  };

  // --- SIMPAN KE MENU ARSIP ---
  const handleSaveSopToArchive = () => {
    const worksheets: SavedDocArchiveWorksheet[] = [];

    selectedSubTabs.forEach((subTab) => {
      selectedPlans.forEach((plan) => {
        const sasaranLabel = `${plan.title} (${plan.dietType})`;

        if (subTab === "organoleptik") {
          const targetDates = organoPrintAll ? activeDates : [effectiveOrganoDate];
          const columns = [
            "Hari / Tanggal",
            "No",
            "Nama Hidangan",
            "Rasa",
            "Warna",
            "Aroma",
            "Tekstur",
            "Waktu Uji",
            "Ket",
            "Kesimpulan",
          ];
          const rows: string[][] = [];
          targetDates.forEach((dStr) => {
            const rec = getOrganoDayRecord(plan.id, dStr);
            const kesimpulanText =
              rec.kesimpulan === "aman"
                ? "Aman dikonsumsi"
                : rec.kesimpulan === "tidak_aman"
                ? "Tidak aman dikonsumsi"
                : "Aman/Tidak aman dikonsumsi";
            (rec.rows || []).forEach((r, idx) => {
              rows.push([
                formatIndonesianDate(dStr),
                String(idx + 1),
                r.namaHidangan || "-",
                r.rasa || "-",
                r.warna || "-",
                r.aroma || "-",
                r.tekstur || "-",
                r.waktuUji || "-",
                r.ket || "-",
                idx === 0 ? kesimpulanText : "",
              ]);
            });
          });
          worksheets.push({
            id: `sop_org_${plan.id}`,
            title: `Uji Organoleptik (Lamp. 3) — ${sasaranLabel}`,
            sasaranLabel,
            subLabel: `${targetDates.length} Tanggal · ${activeMonthName} ${activeYear}`,
            columns,
            rows,
          });
        } else if (subTab === "sampel") {
          const columns = [
            "Tanggal",
            "Pengantaran",
            "Nama Petugas",
            "Nama Makanan",
            "Paraf",
          ];
          const rows = filteredMultiDates.map((dStr, idx) => {
            const r = getSampelRow(plan.id, dStr, idx);
            return [
              formatIndonesianDate(dStr),
              r.pengantaran || "-",
              (r.namaPetugas || "-").replace(/\n/g, " | "),
              r.namaMakanan || "-",
              r.parafChecked ? "Terverifikasi" : "Belum",
            ];
          });
          worksheets.push({
            id: `sop_sampel_${plan.id}`,
            title: `Pengambilan Sampel (Lamp. 2) — ${sasaranLabel}`,
            sasaranLabel,
            subLabel: `${filteredMultiDates.length} Hari · ${activeMonthName} ${activeYear}`,
            columns,
            rows,
          });
        } else {
          const columns = [
            "Kode Identitas Makanan",
            "Jumlah Sampel",
            "OK",
            "NOK",
            "Penanggung Jawab",
            "Masuk Tgl",
            "Masuk Jam",
            "Keluar Tgl",
            "Keluar Jam",
          ];
          const rows = filteredMultiDates.map((dStr) => {
            const r = getPenyimpananRow(plan.id, dStr);
            return [
              r.kodeIdentitas || "-",
              r.jumlahSampel || "-",
              r.kondisi === "OK" ? "✓" : "",
              r.kondisi === "NOK" ? "✓" : "",
              r.namaPenanggungJawab || "-",
              r.masukTanggal || "-",
              r.masukJam || "-",
              r.keluarTanggal || "-",
              r.keluarJam || "-",
            ];
          });
          worksheets.push({
            id: `sop_chiller_${plan.id}`,
            title: `Penyimpanan Sampel (Chiller) — ${sasaranLabel}`,
            sasaranLabel,
            subLabel: `${filteredMultiDates.length} Hari · ${activeMonthName} ${activeYear}`,
            columns,
            rows,
          });
        }
      });
    });

    const formNames = selectedSubTabs.map((t) => SOP_TAB_LABELS[t]).join(" + ");
    const sasaranSummary = selectedPlans.map((p) => p.title).join(" + ") || "Umum";
    const archive: SavedDocArchive = {
      id: `sop_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      source: "sop",
      subType: selectedSubTabs.join("+"),
      title: `Form SOP (${formNames}) — ${sasaranSummary}`,
      periodeLabel: `${activeMonthName} ${activeYear}`,
      sppgName: profile.namaLembaga || "",
      idSppg: profile.idSPPG || "",
      sasaranSummary,
      savedAt: new Date().toISOString(),
      selectedPlanIds: [...selectedPlanIds],
      selectedSubTabs: [...selectedSubTabs],
      worksheets,
    };

    onSaveToArchive?.(archive);
    setSavedArchiveNotice(
      `Berhasil disimpan ke Menu Arsip: ${archive.title} (${worksheets.length} lembar kerja).`
    );
  };

  // --- PRINT / PDF / EXCEL ---
  const currentPrintElementId = "print-area-sop-active";

  const multiSheetCount = useMemo(() => {
    let total = 0;
    selectedSubTabs.forEach((tab) => {
      if (tab === "organoleptik") {
        total += selectedPlans.length * (organoPrintAll ? activeDates.length : 1);
      } else {
        total += selectedPlans.length;
      }
    });
    return total;
  }, [selectedSubTabs, selectedPlans.length, organoPrintAll, activeDates.length]);

  const currentOrientation: "portrait" | "landscape" | "auto" =
    selectedSubTabs.length === 1 && selectedSubTabs[0] === "penyimpanan"
      ? "landscape"
      : selectedSubTabs.includes("penyimpanan")
      ? "auto"
      : "portrait";

  const currentDocFilename =
    selectedSubTabs.length === 1 && selectedPlans.length === 1
      ? activeSubTab === "organoleptik"
        ? `SOP_Uji_Organoleptik_${organoPrintAll ? defaultMonth : effectiveOrganoDate}`
        : activeSubTab === "sampel"
        ? `SOP_Checklist_Sampel_Makanan_${defaultMonth}`
        : `SOP_Penyimpanan_Sampel_Chiller_${defaultMonth}`
      : `SOP_Gabungan_${selectedPlans.length}_Sasaran_${defaultMonth}`;

  const currentDocTitle =
    selectedSubTabs.length === 1 && selectedPlans.length === 1
      ? activeSubTab === "organoleptik"
        ? `Checklist Uji Organoleptik - ${activeMonthName} ${activeYear}`
        : activeSubTab === "sampel"
        ? `Checklist Sampel Makanan - ${activeMonthName} ${activeYear}`
        : `Penyimpanan Sampel Makanan (Chiller) - ${activeMonthName} ${activeYear}`
      : `Form SOP (${selectedSubTabs.length} Form × ${selectedPlans.length} Sasaran) - ${activeMonthName} ${activeYear}`;

  const handlePrint = () => {
    printElementSmart(currentPrintElementId, {
      title: currentDocTitle,
      paperSize,
      orientation: currentOrientation,
    });
  };

  const handleExportPdf = async () => {
    await exportToPDF(currentPrintElementId, currentDocFilename, 100, setIsDownloading, {
      paperSize,
      orientation: currentOrientation,
    });
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const usedNames = new Set<string>();
    const makeSheetName = (raw: string, idx: number) => {
      const cleaned = raw.replace(/[\\/?*[\]:]/g, " ").replace(/\s+/g, " ").trim();
      let candidate = (cleaned || `SOP ${idx + 1}`).slice(0, 31);
      if (usedNames.has(candidate)) {
        const tag = `_${idx + 1}`;
        candidate = `${candidate.slice(0, 31 - tag.length)}${tag}`;
      }
      usedNames.add(candidate);
      return candidate;
    };

    let sheetCounter = 0;

    selectedSubTabs.forEach((subTab) => {
      selectedPlans.forEach((plan) => {
        if (subTab === "organoleptik") {
          const targetDates = organoPrintAll ? activeDates : [effectiveOrganoDate];
          const rows: any[][] = [
            ["LAMPIRAN 3 - CHECKLIST UJI ORGANOLEPTIK"],
            ["Kelompok Sasaran", `${plan.title} (${plan.dietType})`],
            ["Nama SPPG", profile.namaLembaga || "-", "ID SPPG", profile.idSPPG || "-"],
            ["Bulan / Tahun", `${activeMonthName} ${activeYear}`],
            [],
            [
              "Hari / Tanggal",
              "No",
              "Nama Hidangan",
              "Rasa (1-5)",
              "Warna (1-5)",
              "Aroma (1-5)",
              "Tekstur (1-5)",
              "Waktu Uji",
              "Ket",
              "Kesimpulan Menu MBG",
            ],
          ];
          targetDates.forEach((dStr) => {
            const rec = getOrganoDayRecord(plan.id, dStr);
            const kesimpulanText =
              rec.kesimpulan === "aman"
                ? "Aman dikonsumsi"
                : rec.kesimpulan === "tidak_aman"
                ? "Tidak aman dikonsumsi"
                : "Aman/Tidak aman dikonsumsi";
            (rec.rows || []).forEach((r, idx) => {
              rows.push([
                formatIndonesianDate(dStr),
                idx + 1,
                r.namaHidangan,
                r.rasa,
                r.warna,
                r.aroma,
                r.tekstur,
                r.waktuUji,
                r.ket,
                idx === 0 ? kesimpulanText : "",
              ]);
            });
          });
          const ws = XLSX.utils.aoa_to_sheet(rows);
          ws["!cols"] = [
            { wch: 26 },
            { wch: 5 },
            { wch: 30 },
            { wch: 11 },
            { wch: 11 },
            { wch: 11 },
            { wch: 11 },
            { wch: 24 },
            { wch: 14 },
            { wch: 24 },
          ];
          const name = makeSheetName(`Organo ${plan.title}`, sheetCounter++);
          XLSX.utils.book_append_sheet(wb, ws, name);
        } else if (subTab === "sampel") {
          const rows: any[][] = [
            ["LAMPIRAN 2 - CHECKLIST SAMPEL MAKANAN"],
            ["Kelompok Sasaran", `${plan.title} (${plan.dietType})`],
            ["Nama SPPG", profile.namaLembaga || "-"],
            ["ID SPPG", profile.idSPPG || "-"],
            ["Periode (Bulan)", activeMonthName, "Tahun", activeYear],
            [],
            ["Tanggal", "Pengantaran", "Nama Petugas", "Nama Makanan", "Paraf (Pengawas Gizi/Ka SPPG)"],
          ];
          filteredMultiDates.forEach((dStr, idx) => {
            const row = getSampelRow(plan.id, dStr, idx);
            rows.push([
              formatIndonesianDate(dStr),
              row.pengantaran || "",
              row.namaPetugas || "",
              row.namaMakanan || "",
              row.parafChecked ? "Terverifikasi" : "",
            ]);
          });
          const ws = XLSX.utils.aoa_to_sheet(rows);
          ws["!cols"] = [{ wch: 26 }, { wch: 20 }, { wch: 28 }, { wch: 42 }, { wch: 24 }];
          const name = makeSheetName(`Sampel ${plan.title}`, sheetCounter++);
          XLSX.utils.book_append_sheet(wb, ws, name);
        } else {
          const rows: any[][] = [
            ["FORM PENYIMPANAN SAMPEL MAKANAN (CHILLER)"],
            ["Kelompok Sasaran", `${plan.title} (${plan.dietType})`],
            ["Bulan", activeMonthName, "Tahun", activeYear, "ID SPPG", profile.idSPPG || "-"],
            ["Nama SPPG", profile.namaLembaga || "-"],
            [],
            [
              "Kode Identitas Makanan",
              "Jumlah Sampel",
              "Kondisi OK",
              "Kondisi NOK",
              "Nama Penanggung Jawab",
              "Masuk Chiler - Tanggal",
              "Masuk Chiler - Jam",
              "Keluar Chiler - Tanggal",
              "Keluar Chiler - Jam",
            ],
          ];
          filteredMultiDates.forEach((dStr) => {
            const row = getPenyimpananRow(plan.id, dStr);
            rows.push([
              row.kodeIdentitas || "",
              row.jumlahSampel || "",
              row.kondisi === "OK" ? "✓" : "",
              row.kondisi === "NOK" ? "✓" : "",
              row.namaPenanggungJawab || "",
              row.masukTanggal || "",
              row.masukJam || "",
              row.keluarTanggal || "",
              row.keluarJam || "",
            ]);
          });
          const ws = XLSX.utils.aoa_to_sheet(rows);
          ws["!cols"] = [
            { wch: 44 },
            { wch: 15 },
            { wch: 11 },
            { wch: 12 },
            { wch: 30 },
            { wch: 18 },
            { wch: 14 },
            { wch: 18 },
            { wch: 14 },
          ];
          const name = makeSheetName(`Chiller ${plan.title}`, sheetCounter++);
          XLSX.utils.book_append_sheet(wb, ws, name);
        }
      });
    });

    XLSX.writeFile(wb, `${currentDocFilename}.xlsx`);
  };

  // --- RENDER SINGLE ORGANOLEPTIK PAGE (exact Gambar 1 layout, plan-scoped) ---
  const renderOrganoleptikSheet = (
    plan: MasterMenuPlan,
    dateStr: string,
    isPageBreak: boolean
  ) => {
    const rec = getOrganoDayRecord(plan.id, dateStr);
    const rows = rec.rows || [];
    const formattedDateFull = formatIndonesianDate(dateStr);
    const dateOnlyForSign = (() => {
      const dt = parseLocalDate(dateStr);
      if (!dt) return formattedDateFull;
      return `${String(dt.getDate()).padStart(2, "0")} ${INDONESIAN_MONTHS[dt.getMonth()]} ${dt.getFullYear()}`;
    })();
    const kotaText = ttdConfig.kotaTTD || profile.kabKota || "";

    return (
      <div
        key={`org_${plan.id}_${dateStr}`}
        className={`bg-white text-black mx-auto w-full max-w-[215mm] ${
          isPageBreak
            ? "print:break-before-page pt-8 border-t-2 border-dashed border-slate-300 print:border-t-0 print:pt-0 mt-8 print:mt-0"
            : ""
        }`}
        style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
      >
        {showKop && (
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
            docKey="sop"
          />
        )}

        <div className="flex items-center justify-between gap-2 mb-1">
          <p className="text-[11px] font-bold">Lampiran 3.</p>
          <span className="text-[10.5px] font-bold uppercase text-slate-700 print:text-black">
            Kelompok Sasaran: {plan.title} ({plan.dietType})
          </span>
        </div>
        <h2 className="text-center text-[13px] font-bold uppercase tracking-wide mb-4">
          CHECKLIST UJI ORGANOLEPTIK
        </h2>

        {/* Metadata Block matching Gambar 1 */}
        <div className="space-y-1 text-[11px] mb-3 max-w-2xl">
          <div className="flex items-center gap-2">
            <span className="w-44 shrink-0">Nama Pemeriksa</span>
            <span>:</span>
            <input
              type="text"
              value={rec.namaPemeriksa || ""}
              onChange={(e) =>
                updateOrganoDay(plan.id, dateStr, (r) => ({ ...r, namaPemeriksa: e.target.value }))
              }
              className="flex-1 bg-transparent border-b border-transparent hover:border-slate-300 focus:border-indigo-500 focus:outline-none py-0.5 text-[11px] text-black print:border-none"
              placeholder={profile.namaAhliGizi || "Nama Ahli Gizi / Pemeriksa"}
            />
          </div>
          <div className="flex items-center gap-2">
            <span className="w-44 shrink-0">Tempat Pemeriksaan</span>
            <span>:</span>
            <input
              type="text"
              value={rec.tempatPemeriksaan || ""}
              onChange={(e) =>
                updateOrganoDay(plan.id, dateStr, (r) => ({
                  ...r,
                  tempatPemeriksaan: e.target.value,
                }))
              }
              className="flex-1 bg-transparent border-b border-transparent hover:border-slate-300 focus:border-indigo-500 focus:outline-none py-0.5 text-[11px] text-black print:border-none"
              placeholder={profile.namaLembaga || "Nama SPPG"}
            />
          </div>
          <div className="flex items-center gap-2">
            <span className="w-44 shrink-0">Nama Tempat Pemeriksa</span>
            <span>:</span>
            <input
              type="text"
              value={rec.namaTempatPemeriksa || ""}
              onChange={(e) =>
                updateOrganoDay(plan.id, dateStr, (r) => ({
                  ...r,
                  namaTempatPemeriksa: e.target.value,
                }))
              }
              className="flex-1 bg-transparent border-b border-transparent hover:border-slate-300 focus:border-indigo-500 focus:outline-none py-0.5 text-[11px] text-black print:border-none"
              placeholder="Ruang Ahli Gizi"
            />
          </div>
          <div className="flex items-center gap-2">
            <span className="w-44 shrink-0">Hari/Tanggal Pemeriksa</span>
            <span>:</span>
            <span className="font-medium text-black">{formattedDateFull}</span>
          </div>
          <div className="flex items-center gap-2">
            <span className="w-44 shrink-0">Waktu Pemeriksa</span>
            <span>:</span>
            <div className="inline-flex items-center gap-1.5">
              <input
                type="text"
                value={rec.waktuJam || ""}
                onChange={(e) =>
                  updateOrganoDay(plan.id, dateStr, (r) => ({ ...r, waktuJam: e.target.value }))
                }
                placeholder="   :   "
                className="w-20 bg-transparent border-b border-slate-300 focus:border-indigo-500 focus:outline-none py-0.5 text-[11px] text-center tabular-nums print:border-none"
              />
              <select
                value={rec.zonaWaktu || "WITA"}
                onChange={(e) =>
                  updateOrganoDay(plan.id, dateStr, (r) => ({
                    ...r,
                    zonaWaktu: e.target.value as "WIB" | "WITA" | "WIT",
                  }))
                }
                className="bg-transparent text-[11px] font-medium focus:outline-none print:appearance-none"
              >
                <option value="WITA">WITA</option>
                <option value="WIB">WIB</option>
                <option value="WIT">WIT</option>
              </select>
            </div>
          </div>
        </div>

        {/* Main Organoleptik Table (exact Gambar 1 columns) */}
        <div className="overflow-x-auto">
          <table className="w-full border-collapse border border-black text-[11px] text-black">
            <thead>
              <tr>
                <th
                  rowSpan={2}
                  className="border border-black px-2 py-1.5 text-center font-bold w-9"
                >
                  No
                </th>
                <th
                  rowSpan={2}
                  className="border border-black px-2 py-1.5 text-center font-bold min-w-[180px]"
                >
                  Nama Hidangan
                </th>
                <th
                  colSpan={4}
                  className="border border-black px-2 py-1 text-center font-bold"
                >
                  Hasil Pemeriksaan
                  <div className="font-normal text-[10px]">(diberi skor 1-5)</div>
                </th>
                <th
                  rowSpan={2}
                  className="border border-black px-2 py-1.5 text-center font-bold w-40 leading-tight"
                >
                  Sebelum Pengantaran/ Tiba di Lokasi/ Sebelum Dikonsumsi
                </th>
                <th
                  rowSpan={2}
                  className="border border-black px-2 py-1.5 text-center font-bold w-24"
                >
                  Ket
                </th>
                <th className="border border-black px-1 py-1 text-center w-8 no-print bg-slate-50" rowSpan={2}>
                  ×
                </th>
              </tr>
              <tr>
                <th className="border border-black px-1.5 py-1 text-center font-normal text-[10px] w-12">
                  Rasa
                </th>
                <th className="border border-black px-1.5 py-1 text-center font-normal text-[10px] w-12">
                  Warna
                </th>
                <th className="border border-black px-1.5 py-1 text-center font-normal text-[10px] w-12">
                  Aroma
                </th>
                <th className="border border-black px-1.5 py-1 text-center font-normal text-[10px] w-12">
                  Tekstur
                </th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row, rIdx) => (
                <tr key={row.id || rIdx}>
                  <td className="border border-black px-1.5 py-1.5 text-center tabular-nums">
                    {rIdx + 1}
                  </td>
                  <td className="border border-black px-2 py-1">
                    <input
                      type="text"
                      value={row.namaHidangan}
                      onChange={(e) =>
                        updateOrganoRow(plan.id, dateStr, rIdx, { namaHidangan: e.target.value })
                      }
                      className="w-full bg-transparent text-[11px] text-black focus:outline-none focus:bg-indigo-50/40 px-1 py-0.5"
                      placeholder=""
                    />
                  </td>
                  {(["rasa", "warna", "aroma", "tekstur"] as const).map((field) => (
                    <td key={field} className="border border-black p-0.5 text-center">
                      <select
                        value={row[field]}
                        onChange={(e) =>
                          updateOrganoRow(plan.id, dateStr, rIdx, { [field]: e.target.value })
                        }
                        className="w-full text-center bg-transparent text-[11px] font-semibold tabular-nums focus:outline-none py-1 print:appearance-none cursor-pointer"
                      >
                        <option value=""> </option>
                        <option value="5">5</option>
                        <option value="4">4</option>
                        <option value="3">3</option>
                        <option value="2">2</option>
                        <option value="1">1</option>
                      </select>
                    </td>
                  ))}
                  <td className="border border-black px-1.5 py-1">
                    <input
                      type="text"
                      list={`waktu-uji-list-${plan.id}-${dateStr}`}
                      value={row.waktuUji}
                      onChange={(e) =>
                        updateOrganoRow(plan.id, dateStr, rIdx, { waktuUji: e.target.value })
                      }
                      className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 px-1 py-0.5"
                      placeholder=""
                    />
                  </td>
                  <td className="border border-black px-1.5 py-1">
                    <input
                      type="text"
                      value={row.ket}
                      onChange={(e) =>
                        updateOrganoRow(plan.id, dateStr, rIdx, { ket: e.target.value })
                      }
                      className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 px-1 py-0.5"
                      placeholder=""
                    />
                  </td>
                  <td className="border border-black p-0.5 text-center no-print">
                    <button
                      type="button"
                      onClick={() => removeOrganoRow(plan.id, dateStr, rIdx)}
                      disabled={rows.length <= 1}
                      className="p-1 text-slate-400 hover:text-rose-600 disabled:opacity-30 cursor-pointer"
                      title="Hapus baris"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </td>
                </tr>
              ))}

              {/* Kesimpulan Row matching Gambar 1 */}
              <tr>
                <td className="border border-black px-2 py-1.5" />
                <td colSpan={5} className="border border-black px-2 py-1.5 text-center font-bold">
                  Kesimpulan Menu MBG
                </td>
                <td colSpan={2} className="border border-black px-2 py-1.5 text-center">
                  <select
                    value={rec.kesimpulan || "belum"}
                    onChange={(e) =>
                      updateOrganoDay(plan.id, dateStr, (r) => ({
                        ...r,
                        kesimpulan: e.target.value as "aman" | "tidak_aman" | "belum",
                      }))
                    }
                    className="w-full text-center bg-transparent text-[11px] font-semibold focus:outline-none print:appearance-none cursor-pointer"
                  >
                    <option value="belum">Aman/Tidak aman dikonsumsi</option>
                    <option value="aman">Aman dikonsumsi</option>
                    <option value="tidak_aman">Tidak aman dikonsumsi</option>
                  </select>
                </td>
                <td className="border border-black no-print" />
              </tr>
            </tbody>
          </table>
          <datalist id={`waktu-uji-list-${plan.id}-${dateStr}`}>
            <option value="Sebelum Pengantaran" />
            <option value="Saat Tiba di Lokasi" />
            <option value="Sebelum Dikonsumsi" />
          </datalist>
        </div>

        <div className="mt-2 no-print flex flex-wrap items-center justify-between gap-2">
          <button
            type="button"
            onClick={() => addOrganoRow(plan.id, dateStr)}
            className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-semibold cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5" />
            Tambah Baris Hidangan
          </button>
          <div className="flex flex-wrap items-center gap-1.5 text-xs">
            <span className="text-slate-500">Isi cepat skor:</span>
            <button
              type="button"
              onClick={() => applyQuickOrganoScore(plan.id, dateStr, "5")}
              className="px-2.5 py-1 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 font-semibold cursor-pointer"
            >
              Skor 5 (Sangat Baik)
            </button>
            <button
              type="button"
              onClick={() => applyQuickOrganoScore(plan.id, dateStr, "4")}
              className="px-2.5 py-1 rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-800 border border-blue-200 font-semibold cursor-pointer"
            >
              Skor 4 (Baik)
            </button>
            <button
              type="button"
              onClick={() => applyQuickOrganoScore(plan.id, dateStr, "")}
              className="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold cursor-pointer"
            >
              Kosongkan (Blanko)
            </button>
          </div>
        </div>

        {/* Catatan & Keterangan Skor matching Gambar 1 verbatim */}
        <div className="mt-3 space-y-2.5 text-[10.5px] leading-snug text-black">
          <div>
            <p className="font-bold">Catatan :</p>
            <ol className="list-decimal pl-6 mt-0.5 space-y-0.5">
              <li>
                Pilih salah satu waktu untuk pengisian hasil pelaksanaan uji organoleptik (Sebelum Pengantaran/Saat Tiba di Lokasi/Sebelum Dikonsumsi)
              </li>
              <li>
                Form dibawa oleh supir untuk diberikan ke penanggung jawab program MBG dan diambil kembali bersamaan dengan pengambilan ompreng MBG dan dibawa kembali ke SPPG.
              </li>
            </ol>
          </div>

          <div>
            <p className="font-bold underline">Keterangan Skor :</p>
            <div className="mt-0.5 grid grid-cols-[95px_12px_auto] gap-y-0.5 max-w-[200px]">
              <span>Sangat Baik</span><span>:</span><span className="tabular-nums">5</span>
              <span>BAIK</span><span>:</span><span className="tabular-nums">4</span>
              <span>Cukup</span><span>:</span><span className="tabular-nums">3</span>
              <span>Kurang</span><span>:</span><span className="tabular-nums">2</span>
              <span>Tidak Baik</span><span>:</span><span className="tabular-nums">1</span>
            </div>
          </div>
        </div>

        {/* Signature + ID SPPG Footer */}
        {organoSigMode === "petugas" ? (
          <div className="mt-5 print:break-inside-avoid break-inside-avoid">
            <div className="flex justify-end pr-6">
              <div className="text-left text-[11px] leading-snug w-64">
                <p>
                  {kotaText ? `${kotaText}, ${dateOnlyForSign}` : dateOnlyForSign}
                </p>
                <p className="mt-0.5">Mengetahui</p>
                <p>Petugas Pemeriksaan</p>
                <p>{ttdConfig.jabatanAhliGizi || "Pengawas Gizi"}</p>
                <div className="h-16 flex items-center justify-start my-0.5">
                  {showDigitalTtd && ttdConfig.ttdImageAhliGizi ? (
                    <img
                      src={ttdConfig.ttdImageAhliGizi}
                      alt="TTD Petugas Pemeriksaan"
                      className="max-h-14 max-w-[150px] object-contain"
                      referrerPolicy="no-referrer"
                    />
                  ) : null}
                </div>
                <p className="font-bold underline">
                  {rec.namaPemeriksa || profile.namaAhliGizi || "........................................"}
                </p>
              </div>
            </div>
            <IdSppgFooterStrip profile={profile} />
          </div>
        ) : (
          <AutoSignatureBlock
            profile={profile}
            docKey="sop"
            dateText={dateOnlyForSign}
            showDigitalTtd={showDigitalTtd}
          />
        )}
      </div>
    );
  };

  // --- RENDER SAMPEL MAKANAN SHEET (exact Gambar 2 layout, plan-scoped) ---
  const renderSampelSheet = (plan: MasterMenuPlan, isPageBreak: boolean) => {
    return (
      <div
        key={`sampel_${plan.id}`}
        className={`bg-white text-black mx-auto w-full max-w-[215mm] ${
          isPageBreak
            ? "print:break-before-page pt-8 border-t-2 border-dashed border-slate-300 print:border-t-0 print:pt-0 mt-8 print:mt-0"
            : ""
        }`}
        style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
      >
        {showKop && (
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
            docKey="sop"
          />
        )}

        <div className="flex items-center justify-between gap-2 mb-1">
          <p className="text-[11px] font-bold">Lampiran 2.</p>
          <span className="text-[10.5px] font-bold uppercase text-slate-700 print:text-black">
            Kelompok Sasaran: {plan.title} ({plan.dietType})
          </span>
        </div>

        <div className="text-center mb-4">
          <h2 className="text-[13px] font-bold uppercase tracking-wide">
            CHECKLIST SAMPEL MAKANAN
          </h2>
          <div className="text-[11px] mt-0.5 inline-flex items-center justify-center gap-1">
            <span>No :</span>
            <input
              type="text"
              value={sopData.sampelNomorPrefix ?? ""}
              onChange={(e) => updateSop({ sampelNomorPrefix: e.target.value })}
              placeholder="....."
              className="w-12 text-center bg-transparent border-b border-slate-300 focus:border-indigo-500 focus:outline-none text-[11px] tabular-nums print:border-none"
            />
            <span>/ SOP / SM /</span>
            <span className="font-semibold">{activeMonthRoman}</span>
            <span>/ {activeYear}</span>
          </div>
        </div>

        {/* Metadata Block matching Gambar 2 */}
        <div className="space-y-1 text-[11px] mb-4 max-w-xl">
          <div className="flex gap-2">
            <span className="w-32 shrink-0">Nama SPPG</span>
            <span>:</span>
            <span className="font-semibold">{profile.namaLembaga || "-"}</span>
          </div>
          <div className="flex gap-2">
            <span className="w-32 shrink-0">ID SPPG</span>
            <span>:</span>
            <span className="font-bold tabular-nums">{profile.idSPPG || "-"}</span>
          </div>
          <div className="flex gap-2">
            <span className="w-32 shrink-0">Kelompok Sasaran</span>
            <span>:</span>
            <span className="font-semibold">{plan.title} ({plan.dietType})</span>
          </div>
          <div className="flex gap-2">
            <span className="w-32 shrink-0">Periode (Bulan)</span>
            <span>:</span>
            <span className="font-semibold">{activeMonthName}</span>
          </div>
          <div className="flex gap-2">
            <span className="w-32 shrink-0">Tahun</span>
            <span>:</span>
            <span className="font-semibold tabular-nums">{activeYear}</span>
          </div>
        </div>

        {/* Exact 5-column Table from Gambar 2 */}
        <div className="overflow-x-auto">
          <table className="w-full border-collapse border border-black text-[11px] text-black">
            <thead>
              <tr>
                <th className="border border-black px-2 py-2 text-center font-bold w-32">
                  Tanggal
                </th>
                <th className="border border-black px-2 py-2 text-center font-bold w-28">
                  Pengantaran
                </th>
                <th className="border border-black px-2 py-2 text-center font-bold w-48">
                  Nama Petugas
                </th>
                <th className="border border-black px-2 py-2 text-center font-bold">
                  Nama Makanan
                </th>
                <th className="border border-black px-2 py-2 text-center font-bold w-32 leading-tight">
                  Paraf (Pengawas Gizi/Ka SPPG)
                </th>
              </tr>
            </thead>
            <tbody>
              {filteredMultiDates.map((dStr, idx) => {
                const row = getSampelRow(plan.id, dStr, idx);
                return (
                  <tr key={`${plan.id}_${dStr}`} className="align-top break-inside-avoid">
                    <td className="border border-black px-2 py-2 text-[10.5px] leading-snug tabular-nums">
                      {formatIndonesianDate(dStr)}
                    </td>
                    <td className="border border-black p-1.5">
                      <textarea
                        rows={3}
                        value={row.pengantaran || ""}
                        onChange={(e) =>
                          updateSampelRow(plan.id, dStr, idx, { pengantaran: e.target.value })
                        }
                        placeholder="Kloter / Jam"
                        className="w-full bg-transparent text-[10.5px] text-black resize-none focus:outline-none focus:bg-indigo-50/40 p-0.5 leading-snug print:placeholder-transparent"
                      />
                    </td>
                    <td className="border border-black p-1.5">
                      <textarea
                        rows={3}
                        value={row.namaPetugas || ""}
                        onChange={(e) =>
                          updateSampelRow(plan.id, dStr, idx, { namaPetugas: e.target.value })
                        }
                        placeholder={"1. ...\n2. ...\n3. ..."}
                        className="w-full bg-transparent text-[10.5px] text-black resize-none focus:outline-none focus:bg-indigo-50/40 p-0.5 leading-snug print:placeholder-transparent"
                      />
                    </td>
                    <td className="border border-black p-1.5">
                      <textarea
                        rows={3}
                        value={row.namaMakanan || ""}
                        onChange={(e) =>
                          updateSampelRow(plan.id, dStr, idx, { namaMakanan: e.target.value })
                        }
                        className="w-full bg-transparent text-[10.5px] text-black resize-none focus:outline-none focus:bg-indigo-50/40 p-0.5 leading-snug"
                      />
                    </td>
                    <td
                      onClick={() =>
                        updateSampelRow(plan.id, dStr, idx, { parafChecked: !row.parafChecked })
                      }
                      className="border border-black p-1.5 text-center align-middle cursor-pointer hover:bg-slate-50 transition"
                      title="Klik untuk tampilkan/sembunyikan paraf digital"
                    >
                      {row.parafChecked ? (
                        showDigitalTtd && ttdConfig.ttdImageAhliGizi ? (
                          <img
                            src={ttdConfig.ttdImageAhliGizi}
                            alt="Paraf"
                            className="max-h-10 max-w-[90px] mx-auto object-contain"
                            referrerPolicy="no-referrer"
                          />
                        ) : (
                          <span className="inline-flex items-center justify-center text-emerald-700 font-bold text-xs">
                            <Check className="w-4 h-4" />
                          </span>
                        )
                      ) : (
                        <span className="text-[10px] text-slate-300 no-print">Klik paraf</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <AutoSignatureBlock
          profile={profile}
          docKey="sop"
          dateText={`${activeMonthName} ${activeYear}`}
          showDigitalTtd={showDigitalTtd}
        />
      </div>
    );
  };

  // --- RENDER PENYIMPANAN SAMPEL / CHILLER SHEET (exact Gambar 3 layout, plan-scoped) ---
  const renderPenyimpananSheet = (plan: MasterMenuPlan, isPageBreak: boolean) => {
    return (
      <div
        key={`chiller_${plan.id}`}
        className={`bg-white text-black mx-auto w-full max-w-[297mm] ${
          isPageBreak
            ? "print:break-before-page pt-8 border-t-2 border-dashed border-slate-300 print:border-t-0 print:pt-0 mt-8 print:mt-0"
            : ""
        }`}
        style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
      >
        {showKop && (
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
            docKey="sop"
          />
        )}

        <h2 className="text-center text-[13px] font-bold uppercase tracking-wide mb-1">
          FORM PENYIMPANAN SAMPEL MAKANAN (CHILLER)
        </h2>
        <p className="text-center text-[11px] font-bold text-slate-700 print:text-black mb-3">
          Kelompok Sasaran: {plan.title} ({plan.dietType})
        </p>

        {/* Exact Gambar 3 Bold Italic Header Row */}
        <div className="flex flex-wrap items-baseline justify-between gap-x-6 gap-y-1.5 mb-3 text-[12px]">
          <div className="flex flex-wrap items-center gap-10 font-bold italic">
            <span>Bulan: {activeMonthName}</span>
            <span>Tahun: {activeYear}</span>
          </div>
          <div className="text-[11px] font-semibold text-slate-800 print:text-black tabular-nums text-right">
            {profile.namaLembaga || "-"} · ID SPPG: {profile.idSPPG || "-"}
          </div>
        </div>

        {/* Exact Gambar 3 Table with #D9D9D9 Shaded Header */}
        <div className="overflow-x-auto">
          <table className="w-full border-collapse border border-black text-[11px] text-black">
            <thead>
              <tr>
                <th
                  rowSpan={2}
                  className="border border-black px-2 py-2 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "26%" }}
                >
                  Kode Identitas Makanan
                </th>
                <th
                  rowSpan={2}
                  className="border border-black px-2 py-2 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "10%" }}
                >
                  Jumlah Sampel
                </th>
                <th
                  colSpan={2}
                  className="border border-black px-2 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "10%" }}
                >
                  Kondisi Sampel
                </th>
                <th
                  rowSpan={2}
                  className="border border-black px-2 py-2 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "20%" }}
                >
                  Nama Penanggung Jawab
                </th>
                <th
                  colSpan={2}
                  className="border border-black px-2 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "17%" }}
                >
                  Masuk Chiler
                </th>
                <th
                  colSpan={2}
                  className="border border-black px-2 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "17%" }}
                >
                  Keluar Chiler
                </th>
              </tr>
              <tr>
                <th
                  className="border border-black px-1.5 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "5%" }}
                >
                  OK
                </th>
                <th
                  className="border border-black px-1.5 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "5%" }}
                >
                  NOK
                </th>
                <th
                  className="border border-black px-1.5 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "10%" }}
                >
                  Tanggal
                </th>
                <th
                  className="border border-black px-1.5 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "7%" }}
                >
                  Jam
                </th>
                <th
                  className="border border-black px-1.5 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "10%" }}
                >
                  Tanggal
                </th>
                <th
                  className="border border-black px-1.5 py-1.5 text-center align-middle font-bold"
                  style={{ backgroundColor: "#D9D9D9", width: "7%" }}
                >
                  Jam
                </th>
              </tr>
            </thead>
            <tbody>
              {filteredMultiDates.map((dStr) => {
                const row = getPenyimpananRow(plan.id, dStr);
                return (
                  <tr key={`${plan.id}_${dStr}`} className="break-inside-avoid">
                    <td className="border border-black px-2 py-1">
                      <input
                        type="text"
                        value={row.kodeIdentitas || ""}
                        onChange={(e) =>
                          updatePenyimpananRow(plan.id, dStr, { kodeIdentitas: e.target.value })
                        }
                        className="w-full bg-transparent text-[10.5px] text-black focus:outline-none focus:bg-indigo-50/40 py-0.5"
                      />
                    </td>
                    <td className="border border-black px-1.5 py-1">
                      <input
                        type="text"
                        value={row.jumlahSampel || ""}
                        onChange={(e) =>
                          updatePenyimpananRow(plan.id, dStr, { jumlahSampel: e.target.value })
                        }
                        className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                      />
                    </td>
                    <td
                      onClick={() =>
                        updatePenyimpananRow(plan.id, dStr, {
                          kondisi: row.kondisi === "OK" ? "" : "OK",
                        })
                      }
                      className="border border-black px-1 py-1 text-center font-bold cursor-pointer hover:bg-emerald-50 transition select-none"
                      title="Klik untuk menandai OK"
                    >
                      {row.kondisi === "OK" ? "✓" : ""}
                    </td>
                    <td
                      onClick={() =>
                        updatePenyimpananRow(plan.id, dStr, {
                          kondisi: row.kondisi === "NOK" ? "" : "NOK",
                        })
                      }
                      className="border border-black px-1 py-1 text-center font-bold text-rose-700 print:text-black cursor-pointer hover:bg-rose-50 transition select-none"
                      title="Klik untuk menandai NOK"
                    >
                      {row.kondisi === "NOK" ? "✓" : ""}
                    </td>
                    <td className="border border-black px-1.5 py-1">
                      <input
                        type="text"
                        value={row.namaPenanggungJawab || ""}
                        onChange={(e) =>
                          updatePenyimpananRow(plan.id, dStr, {
                            namaPenanggungJawab: e.target.value,
                          })
                        }
                        className="w-full bg-transparent text-[10.5px] text-black focus:outline-none focus:bg-indigo-50/40 py-0.5"
                      />
                    </td>
                    <td className="border border-black px-1 py-1">
                      <input
                        type="text"
                        value={row.masukTanggal || ""}
                        onChange={(e) =>
                          updatePenyimpananRow(plan.id, dStr, { masukTanggal: e.target.value })
                        }
                        className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                      />
                    </td>
                    <td className="border border-black px-1 py-1">
                      <input
                        type="text"
                        value={row.masukJam || ""}
                        onChange={(e) =>
                          updatePenyimpananRow(plan.id, dStr, { masukJam: e.target.value })
                        }
                        className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                      />
                    </td>
                    <td className="border border-black px-1 py-1">
                      <input
                        type="text"
                        value={row.keluarTanggal || ""}
                        onChange={(e) =>
                          updatePenyimpananRow(plan.id, dStr, { keluarTanggal: e.target.value })
                        }
                        className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                      />
                    </td>
                    <td className="border border-black px-1 py-1">
                      <input
                        type="text"
                        value={row.keluarJam || ""}
                        onChange={(e) =>
                          updatePenyimpananRow(plan.id, dStr, { keluarJam: e.target.value })
                        }
                        className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                      />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Verification Signature Box (exact Gambar 3 bordered box) OR Unified Profil SPPG TTD */}
        {penyimpananSigMode === "kotak_verifikasi" ? (
          <div className="mt-6 print:break-inside-avoid break-inside-avoid">
            <div className="flex justify-end pr-8">
              <table className="border-collapse border border-black text-[11px] text-center w-64">
                <tbody>
                  <tr>
                    <td className="border border-black px-3 py-1 font-normal">
                      Diverifikasi Oleh,
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-black px-3 py-2 h-16 align-middle">
                      {showDigitalTtd && ttdConfig.ttdImageAhliGizi ? (
                        <img
                          src={ttdConfig.ttdImageAhliGizi}
                          alt="TTD Verifikasi"
                          className="max-h-14 max-w-[160px] mx-auto object-contain"
                          referrerPolicy="no-referrer"
                        />
                      ) : null}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-black px-3 py-1 font-semibold">
                      {profile.namaAhliGizi || "........................................"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-black px-3 py-1">
                      {ttdConfig.jabatanAhliGizi || "Ahli Gizi"}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <IdSppgFooterStrip profile={profile} />
          </div>
        ) : (
          <AutoSignatureBlock
            profile={profile}
            docKey="sop"
            dateText={`${activeMonthName} ${activeYear}`}
            showDigitalTtd={showDigitalTtd}
          />
        )}
      </div>
    );
  };

  // Build ordered worksheets list: every selected SOP Form × every selected Kelompok Sasaran
  const orderedSheetElements = (() => {
    const elements: React.ReactNode[] = [];
    let sheetIndex = 0;

    selectedSubTabs.forEach((subTab) => {
      selectedPlans.forEach((plan) => {
        if (subTab === "organoleptik") {
          const dates = organoPrintAll ? activeDates : [effectiveOrganoDate];
          dates.forEach((dStr) => {
            elements.push(renderOrganoleptikSheet(plan, dStr, sheetIndex > 0));
            sheetIndex += 1;
          });
        } else if (subTab === "sampel") {
          elements.push(renderSampelSheet(plan, sheetIndex > 0));
          sheetIndex += 1;
        } else {
          elements.push(renderPenyimpananSheet(plan, sheetIndex > 0));
          sheetIndex += 1;
        }
      });
    });

    return elements;
  })();

  return (
    <div id="form-sop-tab" className="space-y-5 font-sans">
      {/* TOP TOOLBAR & FORM SWITCHER */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 space-y-4 no-print">
        <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-4 border-b border-slate-100 pb-4">
          <div>
            <h2 className="text-lg font-bold text-slate-900">
              Form SOP Pengisian Operasional SPPG
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Dapat menampilkan 1 atau beberapa jenis Form SOP &amp; Kelompok Sasaran sekaligus, lalu simpan langsung ke Menu Arsip.
            </p>
          </div>

          {/* Multi-Select 3-Form Toggle Pills */}
          <div className="flex flex-wrap items-center gap-1.5">
            {VALID_SOP_TABS.map((tab) => {
              const isSelected = selectedSubTabs.includes(tab);
              const Icon =
                tab === "organoleptik"
                  ? ClipboardCheck
                  : tab === "sampel"
                  ? TestTube2
                  : ThermometerSnowflake;
              const iconColor =
                tab === "organoleptik"
                  ? "text-teal-600"
                  : tab === "sampel"
                  ? "text-indigo-600"
                  : "text-sky-600";

              return (
                <div
                  key={tab}
                  className={`inline-flex items-center rounded-xl border transition ${
                    isSelected
                      ? "bg-indigo-50/90 border-indigo-300 text-slate-900 shadow-xs"
                      : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-white"
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => toggleSubTab(tab)}
                    title={
                      isSelected
                        ? "Centang/hapus centang untuk menampilkan beberapa form SOP sekaligus"
                        : "Tambahkan form ini ke tampilan"
                    }
                    className="pl-2.5 pr-1.5 py-2 inline-flex items-center gap-1.5 text-xs font-bold cursor-pointer"
                  >
                    <span
                      className={`w-3.5 h-3.5 rounded flex items-center justify-center border ${
                        isSelected
                          ? "bg-indigo-600 border-indigo-600 text-white"
                          : "bg-white border-slate-300"
                      }`}
                    >
                      {isSelected && <Check className="w-2.5 h-2.5" />}
                    </span>
                    <Icon className={`w-3.5 h-3.5 ${iconColor}`} />
                    <span className="whitespace-nowrap">{SOP_TAB_LABELS[tab]}</span>
                  </button>
                  {selectedSubTabs.length > 1 && isSelected && (
                    <button
                      type="button"
                      onClick={() => selectSingleSubTab(tab)}
                      title="Tampilkan form ini saja"
                      className="pr-2 py-2 text-[10px] font-bold text-indigo-700 hover:text-indigo-900 cursor-pointer"
                    >
                      Fokus
                    </button>
                  )}
                </div>
              );
            })}
            <button
              type="button"
              onClick={selectAllSubTabs}
              className="px-3 py-2 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
            >
              <Layers className="w-3.5 h-3.5 text-indigo-600" />
              Semua Form (3)
            </button>
          </div>
        </div>

        {/* Filter & Sync Controls Row */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
          {/* Active Month */}
          <div className="space-y-1">
            <label className="text-[11px] font-semibold text-slate-500 block">
              Bulan Aktif (Tanggal Otomatis)
            </label>
            <input
              type="month"
              value={defaultMonth}
              onChange={(e) => {
                updateSop({ activeMonth: e.target.value });
                setSelectedDate("");
              }}
              className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500 tabular-nums"
            />
          </div>

          {/* Date Mode + Workdays */}
          <div className="space-y-1">
            <label className="text-[11px] font-semibold text-slate-500 block">
              Pola Tanggal Harian
            </label>
            <div className="flex gap-1.5">
              <select
                value={dateMode}
                onChange={(e) =>
                  updateSop({ dateMode: e.target.value as "bulan_aktif" | "siklus" })
                }
                className="flex-1 px-2.5 py-2 rounded-xl border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500"
              >
                <option value="bulan_aktif">Bulan Aktif Penuh</option>
                <option value="siklus">Tanggal Siklus Menu</option>
              </select>
              {dateMode === "bulan_aktif" && (
                <select
                  value={workDaysPerWeek}
                  onChange={(e) =>
                    updateSop({ workDaysPerWeek: Number(e.target.value) as 5 | 6 | 7 })
                  }
                  className="w-24 px-2 py-2 rounded-xl border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500 tabular-nums"
                  title="Jumlah hari operasional per minggu"
                >
                  <option value={5}>5 Hr/Mgg</option>
                  <option value={6}>6 Hr/Mgg</option>
                  <option value={7}>7 Hr/Mgg</option>
                </select>
              )}
            </div>
          </div>

          {/* KOP & TTD Toggles */}
          <div className="space-y-1">
            <label className="text-[11px] font-semibold text-slate-500 block">
              KOP &amp; TTD dari Profil SPPG
            </label>
            <div className="flex flex-wrap items-center gap-3 pt-1.5 text-xs text-slate-700 font-medium">
              <label className="inline-flex items-center gap-1.5 cursor-pointer select-none whitespace-nowrap">
                <input
                  type="checkbox"
                  checked={sopData.showKop ?? true}
                  onChange={(e) => updateSop({ showKop: e.target.checked })}
                  className="w-3.5 h-3.5 accent-indigo-600 cursor-pointer"
                />
                <span>KOP Surat</span>
              </label>
              <label className="inline-flex items-center gap-1.5 cursor-pointer select-none whitespace-nowrap">
                <input
                  type="checkbox"
                  checked={showDigitalTtd}
                  onChange={(e) => updateSop({ showDigitalTtd: e.target.checked })}
                  className="w-3.5 h-3.5 accent-indigo-600 cursor-pointer"
                />
                <span>Scan TTD</span>
              </label>
            </div>
          </div>
        </div>

        {/* Multi-Select Kelompok Sasaran / Paket Menu */}
        <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-2">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
              <Layers className="w-3.5 h-3.5 text-indigo-600" />
              <span>
                Pilih Kelompok Sasaran / Paket Menu ({selectedPlans.length} dipilih · tampil{" "}
                {multiSheetCount} lembar kerja)
              </span>
            </div>
            <div className="flex items-center gap-1.5">
              {plans.length > 1 && (
                <button
                  type="button"
                  onClick={selectAllPlans}
                  className="px-2.5 py-1 rounded-lg bg-white hover:bg-indigo-50 border border-slate-200 text-indigo-700 text-[11px] font-bold cursor-pointer"
                >
                  Pilih Semua Kelompok Sasaran ({plans.length})
                </button>
              )}
              {selectedPlans.length > 1 && plans[0] && (
                <button
                  type="button"
                  onClick={() => selectSinglePlan(selectedPlans[0].id)}
                  className="px-2.5 py-1 rounded-lg bg-white hover:bg-slate-100 border border-slate-200 text-slate-600 text-[11px] font-semibold cursor-pointer"
                >
                  1 Sasaran Saja
                </button>
              )}
            </div>
          </div>

          <div className="flex flex-wrap gap-2">
            {plans.map((p) => {
              const checked = selectedPlanIds.includes(p.id);
              return (
                <button
                  key={p.id}
                  type="button"
                  onClick={() => togglePlanSelection(p.id)}
                  className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl border text-xs font-semibold transition cursor-pointer ${
                    checked
                      ? "bg-indigo-600 text-white border-indigo-600 shadow-xs"
                      : "bg-white text-slate-700 border-slate-200 hover:border-indigo-300"
                  }`}
                >
                  <span
                    className={`w-3.5 h-3.5 rounded flex items-center justify-center border ${
                      checked
                        ? "bg-white border-white text-indigo-600"
                        : "bg-slate-50 border-slate-300"
                    }`}
                  >
                    {checked && <Check className="w-2.5 h-2.5" />}
                  </span>
                  <span>
                    {p.title} ({p.dietType})
                  </span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Action Buttons Row */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-2 border-t border-slate-100">
          <div className="text-xs text-slate-500 tabular-nums">
            <span>{activeMonthName} {activeYear}</span>
            <span className="mx-1.5" aria-hidden="true">·</span>
            <span>{activeDates.length} hari aktif</span>
            <span className="mx-1.5" aria-hidden="true">·</span>
            <span>{selectedPlans.length} kelompok sasaran</span>
            <span className="mx-1.5" aria-hidden="true">·</span>
            <span>ID SPPG: {profile.idSPPG || "Belum diisi"}</span>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setConfirmResetModal(true)}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-semibold transition cursor-pointer whitespace-nowrap"
              title="Sinkronkan ulang isian form SOP aktif dari Master Menu & Profil SPPG"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Sinkron Ulang Otomatis
            </button>
            <button
              type="button"
              onClick={handleSaveSopToArchive}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold transition shadow-xs cursor-pointer whitespace-nowrap"
            >
              <Archive className="w-3.5 h-3.5" />
              Simpan Arsip
            </button>
            <button
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl border border-emerald-300 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-bold transition cursor-pointer whitespace-nowrap"
            >
              <Eye className="w-3.5 h-3.5" />
              Pratinjau
            </button>
            <button
              type="button"
              onClick={handleExportExcel}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition cursor-pointer whitespace-nowrap"
            >
              <Download className="w-3.5 h-3.5" />
              Excel (.XLSX)
            </button>
            <button
              type="button"
              disabled={!!isDownloading}
              onClick={handleExportPdf}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 disabled:opacity-50 text-white text-xs font-bold transition cursor-pointer whitespace-nowrap"
            >
              <FileDown className="w-3.5 h-3.5" />
              {isDownloading ? isDownloading : "Unduh PDF"}
            </button>
            <button
              type="button"
              onClick={handlePrint}
              className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition cursor-pointer whitespace-nowrap"
            >
              <Printer className="w-3.5 h-3.5" />
              Cetak ({paperSize})
            </button>
          </div>
        </div>

        {/* Archive Saved Confirmation Banner */}
        {savedArchiveNotice && (
          <div className="flex flex-wrap items-center justify-between gap-3 px-4 py-2.5 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-900 text-xs">
            <div className="flex items-center gap-2 font-semibold">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>{savedArchiveNotice}</span>
            </div>
            <div className="flex items-center gap-2">
              {onGoToArsip && (
                <button
                  type="button"
                  onClick={onGoToArsip}
                  className="inline-flex items-center gap-1 px-3 py-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-bold cursor-pointer"
                >
                  Buka Menu Arsip
                  <ArrowRight className="w-3 h-3" />
                </button>
              )}
              <button
                type="button"
                onClick={() => setSavedArchiveNotice(null)}
                className="p-1 text-emerald-700 hover:text-emerald-950 cursor-pointer"
                aria-label="Tutup notifikasi arsip"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        )}
      </div>

      {/* CONTEXTUAL SUB-CONTROLS FOR ACTIVE SOP FORMS */}
      {selectedSubTabs.includes("organoleptik") && (
        <div className="bg-white rounded-2xl border border-slate-200 p-4 space-y-3 no-print">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
              <Calendar className="w-4 h-4 text-teal-600" />
              <span>Kontrol Uji Organoleptik ({activeMonthName} {activeYear})</span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <select
                value={organoSigMode}
                onChange={(e) =>
                  updateSop({
                    organoleptikSignatureMode: e.target.value as "petugas" | "ttd_profil",
                  })
                }
                className="px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-700 bg-slate-50 focus:outline-none"
              >
                <option value="petugas">TTD: Petugas Pemeriksaan (Format Lamp. 3)</option>
                <option value="ttd_profil">TTD: Terpusat Profil SPPG</option>
              </select>

              <button
                type="button"
                onClick={() => setOrganoPrintAll((v) => !v)}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition cursor-pointer whitespace-nowrap ${
                  organoPrintAll
                    ? "bg-indigo-600 text-white border-indigo-600"
                    : "bg-white text-slate-700 border-slate-200 hover:bg-slate-50"
                }`}
              >
                {organoPrintAll
                  ? `Mode Semua Tanggal (${activeDates.length} Tgl × ${selectedPlans.length} Sasaran)`
                  : "Cetak Semua Tanggal Bulan Ini"}
              </button>
            </div>
          </div>

          {!organoPrintAll && (
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1">
              {activeDates.map((dStr) => {
                const isActive = dStr === effectiveOrganoDate;
                return (
                  <button
                    key={dStr}
                    type="button"
                    onClick={() => setSelectedDate(dStr)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold border transition shrink-0 whitespace-nowrap tabular-nums cursor-pointer ${
                      isActive
                        ? "bg-teal-600 text-white border-teal-600"
                        : "bg-slate-50 text-slate-700 border-slate-200 hover:bg-white"
                    }`}
                  >
                    {formatCompactTabDate(dStr)}
                  </button>
                );
              })}
            </div>
          )}
        </div>
      )}

      {selectedSubTabs.includes("sampel") && (
        <div className="bg-white rounded-2xl border border-slate-200 p-4 space-y-3 no-print">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs font-bold text-slate-700">
                Periode Sampel (Lamp. 2):
              </span>
              <div className="inline-flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
                {[
                  { id: "all", label: `Semua (${activeDates.length} Hari)` },
                  { id: "1", label: "Minggu 1" },
                  { id: "2", label: "Minggu 2" },
                  { id: "3", label: "Minggu 3" },
                  { id: "4", label: "Minggu 4" },
                  { id: "5", label: "Minggu 5" },
                ].map((w) => (
                  <button
                    key={w.id}
                    type="button"
                    onClick={() => setWeekFilter(w.id)}
                    className={`px-2.5 py-1 rounded-md text-xs font-semibold transition cursor-pointer whitespace-nowrap ${
                      weekFilter === w.id
                        ? "bg-white text-slate-900 shadow-xs"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    {w.label}
                  </button>
                ))}
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={() => setShowTeamEditor((v) => !v)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-semibold cursor-pointer whitespace-nowrap"
              >
                <Users className="w-3.5 h-3.5 text-indigo-600" />
                {showTeamEditor ? "Tutup Daftar Petugas" : "Atur Tim Petugas"}
              </button>
              <button
                type="button"
                onClick={applyTeamsAcrossMonth}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-bold cursor-pointer whitespace-nowrap"
              >
                Terapkan Tim ke Semua Tanggal
              </button>
            </div>
          </div>

          {showTeamEditor && (
            <div className="pt-3 border-t border-slate-100 space-y-3">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {officerTeams.map((team, tIdx) => (
                  <div
                    key={team.id}
                    className="p-3 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2"
                  >
                    <div className="flex items-center justify-between gap-2">
                      <input
                        type="text"
                        value={team.label}
                        onChange={(e) => {
                          const next = [...officerTeams];
                          next[tIdx] = { ...team, label: e.target.value };
                          updateSop({ sampelTeams: next });
                        }}
                        className="text-xs font-bold text-slate-800 bg-white border border-slate-200 rounded-lg px-2.5 py-1 flex-1 focus:outline-none focus:border-indigo-500"
                      />
                      {officerTeams.length > 1 && (
                        <button
                          type="button"
                          onClick={() => {
                            const next = officerTeams.filter((_, i) => i !== tIdx);
                            updateSop({ sampelTeams: next });
                          }}
                          className="p-1 text-slate-400 hover:text-rose-600 cursor-pointer"
                          title="Hapus tim"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </div>
                    <textarea
                      rows={3}
                      value={team.members.join("\n")}
                      onChange={(e) => {
                        const next = [...officerTeams];
                        next[tIdx] = {
                          ...team,
                          members: e.target.value.split("\n"),
                        };
                        updateSop({ sampelTeams: next });
                      }}
                      placeholder="Satu nama petugas per baris"
                      className="w-full text-xs text-slate-800 bg-white border border-slate-200 rounded-lg p-2 focus:outline-none focus:border-indigo-500"
                    />
                  </div>
                ))}
              </div>
              <button
                type="button"
                onClick={() =>
                  updateSop({
                    sampelTeams: [
                      ...officerTeams,
                      {
                        id: `team_${Date.now()}`,
                        label: `Tim Pengantaran ${officerTeams.length + 1}`,
                        members: ["Petugas 1", "Petugas 2", "Petugas 3"],
                      },
                    ],
                  })
                }
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-xs font-semibold text-slate-700 cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                Tambah Tim Petugas
              </button>
            </div>
          )}
        </div>
      )}

      {selectedSubTabs.includes("penyimpanan") && (
        <div className="bg-white rounded-2xl border border-slate-200 p-4 space-y-3 no-print">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-slate-500 block">
                Default Jumlah Sampel (Chiller)
              </label>
              <input
                type="text"
                value={defaultJumlahSampel}
                onChange={(e) => updateSop({ penyimpananDefaultJumlah: e.target.value })}
                className="w-full px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500"
                placeholder="1 Porsi"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-slate-500 block">
                Jam Masuk Chiler
              </label>
              <input
                type="text"
                value={defaultJamMasuk}
                onChange={(e) => updateSop({ penyimpananDefaultJamMasuk: e.target.value })}
                className="w-full px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500 tabular-nums"
                placeholder="07:30"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-slate-500 block">
                Retensi Chiler Otomatis
              </label>
              <select
                value={retentionDays}
                onChange={(e) =>
                  updateSop({ penyimpananRetentionDays: Number(e.target.value) })
                }
                className="w-full px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500"
              >
                <option value={1}>24 Jam (+1 Hari)</option>
                <option value={2}>48 Jam (+2 Hari — Standar BGN)</option>
                <option value={3}>72 Jam (+3 Hari)</option>
              </select>
            </div>
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-slate-500 block">
                Jam Keluar Chiler
              </label>
              <input
                type="text"
                value={defaultJamKeluar}
                onChange={(e) => updateSop({ penyimpananDefaultJamKeluar: e.target.value })}
                className="w-full px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500 tabular-nums"
                placeholder="07:30"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-slate-500 block">
                Format TTD Bawah (Chiller)
              </label>
              <select
                value={penyimpananSigMode}
                onChange={(e) =>
                  updateSop({
                    penyimpananSignatureMode: e.target.value as
                      | "kotak_verifikasi"
                      | "ttd_profil",
                  })
                }
                className="w-full px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500"
              >
                <option value="kotak_verifikasi">Kotak Diverifikasi (Format Gbr. 3)</option>
                <option value="ttd_profil">TTD Terpusat Profil SPPG</option>
              </select>
            </div>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <div className="inline-flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
              {[
                { id: "all", label: `Semua (${activeDates.length} Hari)` },
                { id: "1", label: "Minggu 1" },
                { id: "2", label: "Minggu 2" },
                { id: "3", label: "Minggu 3" },
                { id: "4", label: "Minggu 4" },
                { id: "5", label: "Minggu 5" },
              ].map((w) => (
                <button
                  key={w.id}
                  type="button"
                  onClick={() => setWeekFilter(w.id)}
                  className={`px-2.5 py-1 rounded-md text-xs font-semibold transition cursor-pointer whitespace-nowrap ${
                    weekFilter === w.id
                      ? "bg-white text-slate-900 shadow-xs"
                      : "text-slate-600 hover:text-slate-900"
                  }`}
                >
                  {w.label}
                </button>
              ))}
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => applyAllPenyimpananKondisi("OK")}
                className="px-3 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 text-xs font-bold cursor-pointer whitespace-nowrap"
              >
                Tandai Semua OK
              </button>
              <button
                type="button"
                onClick={() => applyAllPenyimpananKondisi("")}
                className="px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold cursor-pointer whitespace-nowrap"
              >
                Kosongkan Kondisi (Blanko)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* UNIFIED MULTI-WORKSHEET PRINTABLE CANVAS */}
      <div className="bg-slate-100/80 p-4 sm:p-6 rounded-2xl border border-slate-200 overflow-x-auto print:bg-transparent print:p-0 print:border-none">
        <div
          id="print-area-sop-active"
          className="bg-white mx-auto w-full max-w-[297mm] min-w-[680px] p-6 sm:p-8 border border-slate-300 shadow-xs print:border-none print:shadow-none print:p-0 text-black"
          style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
        >
          {orderedSheetElements}
        </div>
      </div>

      {/* Print Preview Modal */}
      {showPrintModal && (
        <PrintPreviewModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          elementId={currentPrintElementId}
          title={currentDocTitle}
          filename={currentDocFilename}
          defaultPaperSize={paperSize}
          defaultOrientation={currentOrientation}
        />
      )}

      {/* In-App Reset Confirm Modal */}
      {confirmResetModal && (
        <div className="fixed inset-0 z-[99999] bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 no-print">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h4 className="text-base font-bold text-slate-900">
                  Sinkron Ulang Isian Form SOP?
                </h4>
                <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                  Tindakan ini akan memuat ulang seluruh nama menu, tanggal otomatis, dan identitas dari Master Menu &amp; Profil SPPG pada form SOP yang sedang dipilih. Perubahan manual pada form tersebut akan direset.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setConfirmResetModal(false)}
                className="p-1 text-slate-400 hover:text-slate-600 rounded-lg cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setConfirmResetModal(false)}
                className="px-4 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmResetCurrentTab}
                className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold cursor-pointer"
              >
                Ya, Sinkron Ulang
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
