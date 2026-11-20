/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import {
  MasterMenu,
  MasterMenuDietType,
  MasterMenuDurationType,
  MasterMenuPlan,
  MasterMenuScheduleConfig,
  MenuItem,
} from "../types";

export interface SasaranKategoriDef {
  id: string;
  label: string;
  shortLabel: string;
  group: "sekolah" | "tigaB" | "mpAsi";
}

// 12 Kelompok Sasaran sesuai Gambar 2 & Juknis BGN
export const SASARAN_KATEGORI_LIST: SasaranKategoriDef[] = [
  { id: "tk_paud_lb", label: "Siswa TK/PAUD/LB", shortLabel: "TK/PAUD/LB", group: "sekolah" },
  { id: "sd_kelas_1_3", label: "Siswa SD/MI/SLB Kelas 1-3", shortLabel: "1-3 SD", group: "sekolah" },
  { id: "sd_kelas_4_6", label: "Siswa SD/MI/SLB Kelas 4-6", shortLabel: "4-6 SD", group: "sekolah" },
  { id: "smp_mts_smplb", label: "Siswa SMP/MTS/SMPLB", shortLabel: "SMP/MTS", group: "sekolah" },
  { id: "sma_smk_ma", label: "Siswa SMA/SMK/MK/MASMALB", shortLabel: "SMA/SMK/MA", group: "sekolah" },
  { id: "pendidik", label: "Pendidik", shortLabel: "Pendidik", group: "sekolah" },
  { id: "tenaga_kependidikan", label: "Tenaga Pendidikan", shortLabel: "Tenaga Pendidikan", group: "sekolah" },
  { id: "anak_balita", label: "Anak Balita", shortLabel: "Anak Balita", group: "tigaB" },
  { id: "anak_balita_13_59", label: "Anak Balita Usia 13-59 Bulan", shortLabel: "Balita 13-59 Bln", group: "tigaB" },
  { id: "balita_6_11", label: "Balita 6-11 Bulan", shortLabel: "Balita 6-11 Bln", group: "mpAsi" },
  { id: "ibu_hamil", label: "Ibu Hamil", shortLabel: "Ibu Hamil", group: "tigaB" },
  { id: "ibu_menyusui", label: "Ibu Menyusui", shortLabel: "Ibu Menyusui", group: "tigaB" },
];

export const SEKOLAH_SASARAN_IDS = SASARAN_KATEGORI_LIST.filter((s) => s.group === "sekolah").map((s) => s.id);
export const TIGAB_SASARAN_IDS = SASARAN_KATEGORI_LIST.filter((s) => s.group === "tigaB").map((s) => s.id);
export const MPASI_SASARAN_IDS = SASARAN_KATEGORI_LIST.filter((s) => s.group === "mpAsi").map((s) => s.id);

export const DEFAULT_SCHEDULE_CONFIG: MasterMenuScheduleConfig = {
  durationType: "2_minggu",
  startDate: "2026-07-06",
  calendarMonth: "2026-07",
  workDaysPerWeek: 5,
};

export function createEmptyMenuItem(): MenuItem {
  return {
    namaMenu: "",
    karbohidrat: "",
    laukHewani: "",
    laukNabati: "",
    sayur: "",
    buah: "",
    susu: "",
    buahSusu: "",
  };
}

export interface ParsedCommaMenuComponents {
  karbohidrat: string;
  laukHewani: string;
  laukNabati: string;
  sayur: string;
  buah: string;
  susu: string;
  buahSusu: string;
}

/**
 * Splits a comma-separated menu string ("Nasi Putih, Ikan Goreng, Tahu Kecap, Sayur Asem, Semangka, Susu UHT")
 * sequentially into the 6 Master Menu component fields:
 * 0: Makanan Pokok (karbohidrat)
 * 1: Lauk Hewani (laukHewani)
 * 2: Lauk Nabati (laukNabati)
 * 3: Sayuran (sayur)
 * 4: Buah (buah)
 * 5+: Susu (susu)
 */
export function parseCommaMenuComponents(rawNamaMenu: string): ParsedCommaMenuComponents {
  const parts = (rawNamaMenu || "").split(",").map((p) => p.trim());
  const karbohidrat = parts[0] || "";
  const laukHewani = parts[1] || "";
  const laukNabati = parts[2] || "";
  const sayur = parts[3] || "";
  const buah = parts[4] || "";
  const susu =
    parts.length > 6
      ? parts
          .slice(5)
          .filter(Boolean)
          .join(", ")
      : parts[5] || "";
  const buahSusu = [buah, susu].filter(Boolean).join(" + ");
  return {
    karbohidrat,
    laukHewani,
    laukNabati,
    sayur,
    buah,
    susu,
    buahSusu,
  };
}

export function normalizeMenuItem(item?: Partial<MenuItem>): MenuItem {
  const base = createEmptyMenuItem();
  if (!item) return base;

  let karbohidrat = item.karbohidrat || "";
  let laukHewani = item.laukHewani || "";
  let laukNabati = item.laukNabati || "";
  let sayur = item.sayur || "";
  let buah = item.buah ?? "";
  let susu = item.susu ?? "";
  let rawBuahSusu = (item.buahSusu || "").trim();

  // Backfill buah / susu from legacy combined buahSusu if needed
  if (!buah && !susu && rawBuahSusu) {
    const parts = rawBuahSusu.split("+").map((p) => p.trim()).filter(Boolean);
    const susuParts: string[] = [];
    const buahParts: string[] = [];
    parts.forEach((p) => {
      if (/susu|asi|formula|uht/i.test(p)) {
        susuParts.push(p);
      } else {
        buahParts.push(p);
      }
    });
    buah = buahParts.join(" + ");
    susu = susuParts.join(" + ");
  }

  // If namaMenu already contains comma-separated segments and all 6 component fields are empty, auto-populate them
  const allComponentsEmpty =
    !karbohidrat.trim() &&
    !laukHewani.trim() &&
    !laukNabati.trim() &&
    !sayur.trim() &&
    !buah.trim() &&
    !susu.trim();
  if (allComponentsEmpty && (item.namaMenu || "").includes(",")) {
    const parsed = parseCommaMenuComponents(item.namaMenu || "");
    karbohidrat = parsed.karbohidrat;
    laukHewani = parsed.laukHewani;
    laukNabati = parsed.laukNabati;
    sayur = parsed.sayur;
    buah = parsed.buah;
    susu = parsed.susu;
    rawBuahSusu = parsed.buahSusu;
  }

  const combinedBuahSusu =
    rawBuahSusu || [buah.trim(), susu.trim()].filter(Boolean).join(" + ");

  return {
    namaMenu: item.namaMenu || "",
    karbohidrat,
    laukHewani,
    laukNabati,
    sayur,
    buah,
    susu,
    buahSusu: combinedBuahSusu,
  };
}

export function composeMenuName(item?: Partial<MenuItem>): string {
  if (!item) return "";
  const explicit = (item.namaMenu || "").trim();
  if (explicit) return explicit;

  const parts = [
    item.karbohidrat,
    item.laukHewani,
    item.laukNabati,
    item.sayur,
    item.buah,
    item.susu,
  ]
    .map((p) => (p || "").trim())
    .filter((p) => p && p !== "-");

  if (parts.length === 0 && item.buahSusu) {
    return item.buahSusu.trim();
  }
  return parts.join(" + ");
}

// Parse YYYY-MM-DD strictly in local time to prevent timezone date shifting
export function parseLocalDate(dateStr: string): Date | null {
  if (!dateStr) return null;
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(dateStr.trim());
  if (match) {
    const y = Number(match[1]);
    const m = Number(match[2]) - 1;
    const d = Number(match[3]);
    const dt = new Date(y, m, d);
    if (!isNaN(dt.getTime())) return dt;
  }
  const fallback = new Date(dateStr);
  return isNaN(fallback.getTime()) ? null : fallback;
}

export function toYyyyMmDd(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

export function formatIndonesianDate(dateStr: string): string {
  if (!dateStr) return "";
  const d = parseLocalDate(dateStr);
  if (!d) return dateStr;
  const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Sabtu"];
  const months = [
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
  return `${days[d.getDay()]}, ${String(d.getDate()).padStart(2, "0")} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

export function formatIndonesianMonthYear(ym: string): string {
  const m = /^(\d{4})-(\d{2})/.exec((ym || "").trim());
  if (!m) return ym;
  const months = [
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
  const monthIdx = Math.max(0, Math.min(11, Number(m[2]) - 1));
  return `${months[monthIdx]} ${m[1]}`;
}

export function isWorkDay(date: Date, workDaysPerWeek: 5 | 6 | 7): boolean {
  const dow = date.getDay(); // 0=Sun, 6=Sat
  if (workDaysPerWeek === 5) return dow !== 0 && dow !== 6;
  if (workDaysPerWeek === 6) return dow !== 0;
  return true;
}

export function generateScheduleDates(config: MasterMenuScheduleConfig): string[] {
  const workDaysPerWeek = config.workDaysPerWeek || 5;

  if (config.durationType === "1_bulan") {
    const ym = config.calendarMonth || (config.startDate ? config.startDate.slice(0, 7) : "2026-11");
    const match = /^(\d{4})-(\d{2})/.exec(ym);
    if (!match) return [];
    const year = Number(match[1]);
    const monthIdx = Number(match[2]) - 1;
    const lastDay = new Date(year, monthIdx + 1, 0).getDate();
    const dates: string[] = [];
    for (let day = 1; day <= lastDay; day++) {
      const dt = new Date(year, monthIdx, day);
      if (isWorkDay(dt, workDaysPerWeek)) {
        dates.push(toYyyyMmDd(dt));
      }
    }
    return dates;
  }

  const startDt = parseLocalDate(config.startDate) || new Date();
  if (config.durationType === "1_hari") {
    return [toYyyyMmDd(startDt)];
  }

  const weekCountMap: Record<MasterMenuDurationType, number> = {
    "1_hari": 0,
    "1_minggu": 1,
    "2_minggu": 2,
    "3_minggu": 3,
    "4_minggu": 4,
    "1_bulan": 4,
  };
  const targetCount = weekCountMap[config.durationType] * workDaysPerWeek;
  const dates: string[] = [];
  const cursor = new Date(startDt.getFullYear(), startDt.getMonth(), startDt.getDate());
  let guard = 0;
  while (dates.length < targetCount && guard < 120) {
    if (isWorkDay(cursor, workDaysPerWeek)) {
      dates.push(toYyyyMmDd(cursor));
    }
    cursor.setDate(cursor.getDate() + 1);
    guard++;
  }
  return dates;
}

export function computeNextPeriodStart(dates: string[], workDaysPerWeek: 5 | 6 | 7 = 5): string {
  if (!dates || dates.length === 0) return "";
  const last = parseLocalDate(dates[dates.length - 1]);
  if (!last) return "";
  const next = new Date(last.getFullYear(), last.getMonth(), last.getDate() + 1);
  if (workDaysPerWeek < 7) {
    let guard = 0;
    while (next.getDay() !== 1 && guard < 8) {
      next.setDate(next.getDate() + 1);
      guard++;
    }
  }
  return toYyyyMmDd(next);
}

export interface MasterMenuWeekBlock {
  weekNumber: number;
  days: {
    dayIndex: number; // 0-based index across the whole period
    menuNumber: number; // 1-based MENU X
    dateStr: string;
  }[];
}

// Group planned dates into calendar week blocks (Monday-Sunday) so weeks align like Gambar 1
export function groupDatesIntoWeeks(dates: string[]): MasterMenuWeekBlock[] {
  if (!dates || dates.length === 0) return [];
  if (dates.length === 1) {
    return [{ weekNumber: 1, days: [{ dayIndex: 0, menuNumber: 1, dateStr: dates[0] }] }];
  }

  const blocks: MasterMenuWeekBlock[] = [];
  let currentBlock: MasterMenuWeekBlock | null = null;
  let lastWeekKey = "";

  dates.forEach((dateStr, idx) => {
    const dt = parseLocalDate(dateStr);
    if (!dt) {
      // Fallback chunking by 5
      const wIdx = Math.floor(idx / 5);
      if (!blocks[wIdx]) blocks[wIdx] = { weekNumber: wIdx + 1, days: [] };
      blocks[wIdx].days.push({ dayIndex: idx, menuNumber: idx + 1, dateStr });
      return;
    }

    // Compute Monday of the date's week as grouping key
    const dow = dt.getDay(); // 0=Sun..6=Sat
    const diffToMon = dow === 0 ? -6 : 1 - dow;
    const monday = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate() + diffToMon);
    const weekKey = toYyyyMmDd(monday);

    if (!currentBlock || weekKey !== lastWeekKey) {
      currentBlock = {
        weekNumber: blocks.length + 1,
        days: [],
      };
      blocks.push(currentBlock);
      lastWeekKey = weekKey;
    }

    currentBlock.days.push({
      dayIndex: idx,
      menuNumber: idx + 1,
      dateStr,
    });
  });

  return blocks;
}

export function buildCategoryLabelFromIds(sasaranIds: string[]): string {
  if (!sasaranIds || sasaranIds.length === 0) return "BELUM ADA KATEGORI";
  if (sasaranIds.length === SASARAN_KATEGORI_LIST.length) return "SEMUA KATEGORI SASARAN";

  const selectedDefs = SASARAN_KATEGORI_LIST.filter((s) => sasaranIds.includes(s.id));
  if (selectedDefs.length === 0) return "KATEGORI KUSTOM";
  if (selectedDefs.length <= 3) {
    return selectedDefs.map((s) => s.shortLabel).join(", ");
  }
  return `${selectedDefs.slice(0, 2).map((s) => s.shortLabel).join(", ")} +${selectedDefs.length - 2} Lainnya`;
}

export function resizeMenuDays(days: MenuItem[], targetCount: number): MenuItem[] {
  const safeCount = Math.max(1, targetCount);
  const result: MenuItem[] = [];
  for (let i = 0; i < safeCount; i++) {
    result.push(normalizeMenuItem(days[i]));
  }
  return result;
}

export function ensureMasterMenuPlans(menu: MasterMenu, defaultDates?: string[]): MasterMenu {
  const existingConfig = menu.scheduleConfig || {
    ...DEFAULT_SCHEDULE_CONFIG,
    startDate: defaultDates?.[0] || DEFAULT_SCHEDULE_CONFIG.startDate,
    calendarMonth: (defaultDates?.[0] || DEFAULT_SCHEDULE_CONFIG.startDate).slice(0, 7),
    durationType:
      defaultDates && defaultDates.length === 1
        ? "1_hari"
        : defaultDates && defaultDates.length <= 7
          ? "1_minggu"
          : defaultDates && defaultDates.length <= 14
            ? "2_minggu"
            : defaultDates && defaultDates.length <= 21
              ? "3_minggu"
              : "4_minggu",
    periodeDates: defaultDates && defaultDates.length > 0 ? [...defaultDates] : undefined,
  };

  const resolvedDates =
    existingConfig.periodeDates && existingConfig.periodeDates.length > 0
      ? existingConfig.periodeDates
      : generateScheduleDates(existingConfig);

  const dayCount = Math.max(resolvedDates.length, 10);

  if (menu.plans && menu.plans.length > 0) {
    const normalizedPlans = menu.plans.map((p) => ({
      ...p,
      days: (p.days || []).map((d) => normalizeMenuItem(d)),
    }));
    return {
      ...menu,
      scheduleConfig: { ...existingConfig, periodeDates: resolvedDates },
      plans: normalizedPlans,
      activePlanId: menu.activePlanId || normalizedPlans[0].id,
    };
  }

  const plans: MasterMenuPlan[] = [
    {
      id: "plan_sekolah_standar",
      title: "Usia Sekolah & Pendidik",
      sasaranIds: [...SEKOLAH_SASARAN_IDS],
      dietType: "Standar",
      days: resizeMenuDays(menu.usiaSekolah || [], dayCount),
    },
    {
      id: "plan_tigab_standar",
      title: "3B (Balita, Bumil & Busui)",
      sasaranIds: ["anak_balita", "anak_balita_13_59", "ibu_hamil", "ibu_menyusui"],
      dietType: "Standar",
      days: resizeMenuDays(menu.tigaB || [], dayCount),
    },
    {
      id: "plan_mpasi",
      title: "Balita 6-11 Bulan (MP-ASI)",
      sasaranIds: ["balita_6_11"],
      dietType: "MP-ASI",
      days: resizeMenuDays(menu.mpAsi || [], dayCount),
    },
    {
      id: "plan_sekolah_alergi",
      title: "Usia Sekolah (Diet Alergi)",
      sasaranIds: [...SEKOLAH_SASARAN_IDS],
      dietType: "Alergi",
      days: resizeMenuDays(menu.usiaSekolahAlergi || [], dayCount),
    },
  ];

  return {
    ...menu,
    scheduleConfig: { ...existingConfig, periodeDates: resolvedDates },
    plans,
    activePlanId: plans[0].id,
  };
}

// Keep legacy arrays synchronized from the modern multi-category plans
export function syncMasterMenuToLegacy(menu: MasterMenu): MasterMenu {
  const plans = menu.plans || [];
  if (plans.length === 0) return menu;

  const findPlanFor = (sasaranSubset: string[], diet: MasterMenuDietType): MasterMenuPlan | undefined => {
    return (
      plans.find(
        (p) => p.dietType === diet && p.sasaranIds.some((id) => sasaranSubset.includes(id))
      ) ||
      plans.find((p) => p.sasaranIds.some((id) => sasaranSubset.includes(id)))
    );
  };

  const sekolahStd = findPlanFor(SEKOLAH_SASARAN_IDS, "Standar");
  const tigaBStd = findPlanFor(["anak_balita", "anak_balita_13_59", "ibu_hamil", "ibu_menyusui"], "Standar");
  const mpAsiStd = findPlanFor(MPASI_SASARAN_IDS, "MP-ASI") || findPlanFor(MPASI_SASARAN_IDS, "Standar");
  const sekolahAlg = findPlanFor(SEKOLAH_SASARAN_IDS, "Alergi");
  const tigaBAlg = findPlanFor(["anak_balita", "anak_balita_13_59", "ibu_hamil", "ibu_menyusui"], "Alergi");
  const mpAsiAlg = findPlanFor(MPASI_SASARAN_IDS, "Alergi");

  const normalizeDays = (arr: MenuItem[] | undefined, fallback: MenuItem[]): MenuItem[] => {
    const src = arr && arr.length > 0 ? arr : fallback;
    return src.map((d) => {
      const norm = normalizeMenuItem(d);
      return {
        ...norm,
        namaMenu: composeMenuName(norm),
      };
    });
  };

  return {
    ...menu,
    usiaSekolah: normalizeDays(sekolahStd?.days, menu.usiaSekolah || []),
    tigaB: normalizeDays(tigaBStd?.days, menu.tigaB || []),
    mpAsi: normalizeDays(mpAsiStd?.days, menu.mpAsi || []),
    usiaSekolahAlergi: normalizeDays(sekolahAlg?.days, menu.usiaSekolahAlergi || []),
    tigaBAlergi: normalizeDays(tigaBAlg?.days, menu.tigaBAlergi || []),
    mpAsiAlergi: normalizeDays(mpAsiAlg?.days, menu.mpAsiAlergi || []),
  };
}

export function getDurationBadgeLabel(durationType: MasterMenuDurationType, totalDays: number): string {
  switch (durationType) {
    case "1_hari":
      return "1 HARI";
    case "1_minggu":
      return `${totalDays} HARI (1 MINGGU)`;
    case "2_minggu":
      return `${totalDays} HARI (2 MINGGU)`;
    case "3_minggu":
      return `${totalDays} HARI (3 MINGGU)`;
    case "4_minggu":
      return `${totalDays} HARI`;
    case "1_bulan":
      return `${totalDays} HARI (1 BULAN PENUH)`;
    default:
      return `${totalDays} HARI`;
  }
}
