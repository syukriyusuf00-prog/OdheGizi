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
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  DEFAULT_UNIFIED_KOP_CONFIG,
  FoodCostDay,
  MasterMenu,
  SPPGProfile,
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
}: FormSopTabProps) {
  const [selectedDate, setSelectedDate] = useState<string>("");
  const [organoPrintAll, setOrganoPrintAll] = useState<boolean>(false);
  const [weekFilter, setWeekFilter] = useState<string>("all");
  const [showTeamEditor, setShowTeamEditor] = useState<boolean>(false);
  const [isDownloading, setIsDownloading] = useState<string | null>(null);
  const [showPrintModal, setShowPrintModal] = useState<boolean>(false);
  const [confirmResetModal, setConfirmResetModal] = useState<boolean>(false);

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

  const activeSubTab = sopData.activeSubTab || "organoleptik";
  const dateMode = sopData.dateMode || "bulan_aktif";
  const workDaysPerWeek: 5 | 6 | 7 =
    sopData.workDaysPerWeek || normalizedMenu.scheduleConfig?.workDaysPerWeek || 5;
  const activePlanId = sopData.selectedPlanId || normalizedMenu.activePlanId || plans[0]?.id || "";
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

  // Parse activeMonth year & Indonesian month label
  const { activeYear, activeMonthIndex, activeMonthName, activeMonthRoman } = useMemo(() => {
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

  // Active Master Menu plan
  const activePlan = useMemo(() => {
    return plans.find((p) => p.id === activePlanId) || plans[0];
  }, [plans, activePlanId]);

  // Resolve auto menu dishes & full menu name for any dateStr
  const resolvedMenuByDate = useMemo(() => {
    const map: Record<string, { dishes: string[]; fullName: string; hariKe: number }> = {};
    const planDays = activePlan?.days || [];
    const fallbackDays = normalizedMenu.usiaSekolah || [];

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

    return map;
  }, [activeDates, cycleDates, activePlan, normalizedMenu.usiaSekolah, foodCostDays]);

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

  // --- ORGANOLEPTIK HELPERS ---
  const getOrganoDayRecord = (dateStr: string): SopOrganoleptikDayRecord => {
    const saved = sopData.organoleptikByDate?.[dateStr];
    const autoMenu = resolvedMenuByDate[dateStr];
    const autoDishes = autoMenu?.dishes || [];
    const baseMeta = sopData.organoleptikMeta || {};

    let rows: SopOrganoleptikDishRow[];
    if (saved?.rows && saved.rows.length > 0) {
      rows = saved.rows;
    } else {
      const minRows = Math.max(5, autoDishes.length);
      rows = Array.from({ length: minRows }, (_, i) => ({
        id: `org_${dateStr}_${i + 1}`,
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
    dateStr: string,
    updater: (current: SopOrganoleptikDayRecord) => SopOrganoleptikDayRecord
  ) => {
    const current = getOrganoDayRecord(dateStr);
    const nextRec = updater(current);
    setSopData((prev) => ({
      ...prev,
      organoleptikByDate: {
        ...(prev.organoleptikByDate || {}),
        [dateStr]: nextRec,
      },
    }));
  };

  const updateOrganoRow = (
    dateStr: string,
    rowIdx: number,
    patch: Partial<SopOrganoleptikDishRow>
  ) => {
    updateOrganoDay(dateStr, (rec) => {
      const nextRows = [...(rec.rows || [])];
      nextRows[rowIdx] = { ...nextRows[rowIdx], ...patch };
      return { ...rec, rows: nextRows };
    });
  };

  const addOrganoRow = (dateStr: string) => {
    updateOrganoDay(dateStr, (rec) => ({
      ...rec,
      rows: [
        ...(rec.rows || []),
        {
          id: `org_${dateStr}_${Date.now()}`,
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

  const removeOrganoRow = (dateStr: string, rowIdx: number) => {
    updateOrganoDay(dateStr, (rec) => {
      const nextRows = (rec.rows || []).filter((_, i) => i !== rowIdx);
      return { ...rec, rows: nextRows.length > 0 ? nextRows : rec.rows };
    });
  };

  const applyQuickOrganoScore = (dateStr: string, score: string) => {
    updateOrganoDay(dateStr, (rec) => ({
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

  // --- SAMPEL MAKANAN HELPERS ---
  const getSampelRow = (dateStr: string, idx: number): SopSampelRowRecord => {
    const saved = sopData.sampelByDate?.[dateStr];
    const autoMenu = resolvedMenuByDate[dateStr]?.fullName || "";
    const defaultTeam = officerTeams[idx % officerTeams.length];
    return {
      pengantaran: saved?.pengantaran ?? "",
      namaPetugas: saved?.namaPetugas ?? (defaultTeam ? formatTeamMembers(defaultTeam.members) : ""),
      namaMakanan: saved?.namaMakanan ?? autoMenu,
      parafChecked: saved?.parafChecked ?? false,
    };
  };

  const updateSampelRow = (dateStr: string, idx: number, patch: Partial<SopSampelRowRecord>) => {
    const current = getSampelRow(dateStr, idx);
    setSopData((prev) => ({
      ...prev,
      sampelByDate: {
        ...(prev.sampelByDate || {}),
        [dateStr]: { ...current, ...patch },
      },
    }));
  };

  const applyTeamsAcrossMonth = () => {
    if (officerTeams.length === 0) return;
    const nextMap: Record<string, SopSampelRowRecord> = {
      ...(sopData.sampelByDate || {}),
    };
    activeDates.forEach((dateStr, idx) => {
      const existing = nextMap[dateStr] || {};
      const team = officerTeams[idx % officerTeams.length];
      nextMap[dateStr] = {
        ...existing,
        namaPetugas: formatTeamMembers(team.members),
      };
    });
    updateSop({ sampelByDate: nextMap });
  };

  // --- PENYIMPANAN SAMPEL (CHILLER) HELPERS ---
  const getPenyimpananRow = (dateStr: string): SopPenyimpananRowRecord => {
    const saved = sopData.penyimpananByDate?.[dateStr];
    const autoMenu = resolvedMenuByDate[dateStr]?.fullName || "";
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
    dateStr: string,
    patch: Partial<SopPenyimpananRowRecord>
  ) => {
    const current = getPenyimpananRow(dateStr);
    setSopData((prev) => ({
      ...prev,
      penyimpananByDate: {
        ...(prev.penyimpananByDate || {}),
        [dateStr]: { ...current, ...patch },
      },
    }));
  };

  const applyAllPenyimpananKondisi = (kondisi: "OK" | "") => {
    const nextMap: Record<string, SopPenyimpananRowRecord> = {
      ...(sopData.penyimpananByDate || {}),
    };
    filteredMultiDates.forEach((dateStr) => {
      const current = getPenyimpananRow(dateStr);
      nextMap[dateStr] = { ...current, kondisi };
    });
    updateSop({ penyimpananByDate: nextMap });
  };

  // --- RESET / RESYNC AUTO DATA ---
  const handleConfirmResetCurrentTab = () => {
    if (activeSubTab === "organoleptik") {
      updateSop({ organoleptikByDate: {} });
    } else if (activeSubTab === "sampel") {
      updateSop({ sampelByDate: {} });
    } else {
      updateSop({ penyimpananByDate: {} });
    }
    setConfirmResetModal(false);
  };

  // --- PRINT / PDF / EXCEL ---
  const currentPrintElementId =
    activeSubTab === "organoleptik"
      ? "print-area-sop-organoleptik"
      : activeSubTab === "sampel"
      ? "print-area-sop-sampel"
      : "print-area-sop-penyimpanan";

  const currentDocFilename =
    activeSubTab === "organoleptik"
      ? `SOP_Uji_Organoleptik_${organoPrintAll ? defaultMonth : effectiveOrganoDate}`
      : activeSubTab === "sampel"
      ? `SOP_Checklist_Sampel_Makanan_${defaultMonth}`
      : `SOP_Penyimpanan_Sampel_Chiller_${defaultMonth}`;

  const currentDocTitle =
    activeSubTab === "organoleptik"
      ? `Checklist Uji Organoleptik - ${activeMonthName} ${activeYear}`
      : activeSubTab === "sampel"
      ? `Checklist Sampel Makanan - ${activeMonthName} ${activeYear}`
      : `Penyimpanan Sampel Makanan (Chiller) - ${activeMonthName} ${activeYear}`;

  const handlePrint = () => {
    printElementSmart(currentPrintElementId, {
      title: currentDocTitle,
      paperSize,
      orientation: activeSubTab === "penyimpanan" ? "landscape" : "portrait",
    });
  };

  const handleExportPdf = async () => {
    await exportToPDF(currentPrintElementId, currentDocFilename, 100, setIsDownloading, {
      paperSize,
      orientation: activeSubTab === "penyimpanan" ? "landscape" : "portrait",
    });
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    if (activeSubTab === "organoleptik") {
      const targetDates = organoPrintAll ? activeDates : [effectiveOrganoDate];
      const rows: any[][] = [
        ["LAMPIRAN 3 - CHECKLIST UJI ORGANOLEPTIK"],
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
        const rec = getOrganoDayRecord(dStr);
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
      XLSX.utils.book_append_sheet(wb, ws, "Uji Organoleptik");
    } else if (activeSubTab === "sampel") {
      const rows: any[][] = [
        ["LAMPIRAN 2 - CHECKLIST SAMPEL MAKANAN"],
        ["Nama SPPG", profile.namaLembaga || "-"],
        ["ID SPPG", profile.idSPPG || "-"],
        ["Periode (Bulan)", activeMonthName, "Tahun", activeYear],
        [],
        ["Tanggal", "Pengantaran", "Nama Petugas", "Nama Makanan", "Paraf (Pengawas Gizi/Ka SPPG)"],
      ];
      filteredMultiDates.forEach((dStr, idx) => {
        const row = getSampelRow(dStr, idx);
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
      XLSX.utils.book_append_sheet(wb, ws, "Checklist Sampel");
    } else {
      const rows: any[][] = [
        ["FORM PENYIMPANAN SAMPEL MAKANAN (CHILLER)"],
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
        const row = getPenyimpananRow(dStr);
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
      XLSX.utils.book_append_sheet(wb, ws, "Penyimpanan Chiller");
    }

    XLSX.writeFile(wb, `${currentDocFilename}.xlsx`);
  };

  // --- RENDER SINGLE ORGANOLEPTIK PAGE (exact Gambar 1 layout) ---
  const renderOrganoleptikSheet = (dateStr: string, isPageBreak: boolean) => {
    const rec = getOrganoDayRecord(dateStr);
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
        key={dateStr}
        className={`bg-white text-black ${
          isPageBreak ? "print:break-before-page pt-8 border-t-2 border-dashed border-slate-300 print:border-t-0 print:pt-0 mt-8 print:mt-0" : ""
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

        <p className="text-[11px] font-bold mb-1">Lampiran 3.</p>
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
                updateOrganoDay(dateStr, (r) => ({ ...r, namaPemeriksa: e.target.value }))
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
                updateOrganoDay(dateStr, (r) => ({ ...r, tempatPemeriksaan: e.target.value }))
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
                updateOrganoDay(dateStr, (r) => ({ ...r, namaTempatPemeriksa: e.target.value }))
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
                  updateOrganoDay(dateStr, (r) => ({ ...r, waktuJam: e.target.value }))
                }
                placeholder="   :   "
                className="w-20 bg-transparent border-b border-slate-300 focus:border-indigo-500 focus:outline-none py-0.5 text-[11px] text-center tabular-nums print:border-none"
              />
              <select
                value={rec.zonaWaktu || "WITA"}
                onChange={(e) =>
                  updateOrganoDay(dateStr, (r) => ({
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
                        updateOrganoRow(dateStr, rIdx, { namaHidangan: e.target.value })
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
                          updateOrganoRow(dateStr, rIdx, { [field]: e.target.value })
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
                      list={`waktu-uji-list-${dateStr}`}
                      value={row.waktuUji}
                      onChange={(e) =>
                        updateOrganoRow(dateStr, rIdx, { waktuUji: e.target.value })
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
                        updateOrganoRow(dateStr, rIdx, { ket: e.target.value })
                      }
                      className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 px-1 py-0.5"
                      placeholder=""
                    />
                  </td>
                  <td className="border border-black p-0.5 text-center no-print">
                    <button
                      type="button"
                      onClick={() => removeOrganoRow(dateStr, rIdx)}
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
                      updateOrganoDay(dateStr, (r) => ({
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
          <datalist id={`waktu-uji-list-${dateStr}`}>
            <option value="Sebelum Pengantaran" />
            <option value="Saat Tiba di Lokasi" />
            <option value="Sebelum Dikonsumsi" />
          </datalist>
        </div>

        <div className="mt-2 no-print flex flex-wrap items-center justify-between gap-2">
          <button
            type="button"
            onClick={() => addOrganoRow(dateStr)}
            className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-semibold cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5" />
            Tambah Baris Hidangan
          </button>
          <div className="flex flex-wrap items-center gap-1.5 text-xs">
            <span className="text-slate-500">Isi cepat skor:</span>
            <button
              type="button"
              onClick={() => applyQuickOrganoScore(dateStr, "5")}
              className="px-2.5 py-1 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 font-semibold cursor-pointer"
            >
              Skor 5 (Sangat Baik)
            </button>
            <button
              type="button"
              onClick={() => applyQuickOrganoScore(dateStr, "4")}
              className="px-2.5 py-1 rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-800 border border-blue-200 font-semibold cursor-pointer"
            >
              Skor 4 (Baik)
            </button>
            <button
              type="button"
              onClick={() => applyQuickOrganoScore(dateStr, "")}
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

  return (
    <div id="form-sop-tab" className="space-y-5 font-sans">
      {/* TOP TOOLBAR & FORM SWITCHER */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 space-y-4 no-print">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-slate-100 pb-4">
          <div>
            <h2 className="text-lg font-bold text-slate-900">
              Form SOP Pengisian Operasional SPPG
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Terisi otomatis menyesuaikan nama menu harian, bulan aktif, identitas &amp; ID SPPG, KOP Surat, dan TTD dari Profil SPPG.
            </p>
          </div>

          {/* 3-Tab Segmented Switcher */}
          <div className="inline-flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-xl self-start lg:self-auto">
            <button
              type="button"
              onClick={() => updateSop({ activeSubTab: "organoleptik" })}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition whitespace-nowrap inline-flex items-center gap-1.5 cursor-pointer ${
                activeSubTab === "organoleptik"
                  ? "bg-white text-slate-900 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <ClipboardCheck className="w-3.5 h-3.5 text-teal-600" />
              Uji Organoleptik (Lamp. 3)
            </button>
            <button
              type="button"
              onClick={() => updateSop({ activeSubTab: "sampel" })}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition whitespace-nowrap inline-flex items-center gap-1.5 cursor-pointer ${
                activeSubTab === "sampel"
                  ? "bg-white text-slate-900 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <TestTube2 className="w-3.5 h-3.5 text-indigo-600" />
              Pengambilan Sampel (Lamp. 2)
            </button>
            <button
              type="button"
              onClick={() => updateSop({ activeSubTab: "penyimpanan" })}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition whitespace-nowrap inline-flex items-center gap-1.5 cursor-pointer ${
                activeSubTab === "penyimpanan"
                  ? "bg-white text-slate-900 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <ThermometerSnowflake className="w-3.5 h-3.5 text-sky-600" />
              Penyimpanan Sampel (Chiller)
            </button>
          </div>
        </div>

        {/* Filter & Sync Controls Row */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
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

          {/* Menu Source Plan */}
          <div className="space-y-1">
            <label className="text-[11px] font-semibold text-slate-500 block">
              Sumber Nama Menu (Master Menu)
            </label>
            <select
              value={activePlanId}
              onChange={(e) => updateSop({ selectedPlanId: e.target.value })}
              className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs font-semibold text-slate-800 bg-slate-50/50 focus:outline-none focus:border-indigo-500"
            >
              {plans.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.title} ({p.dietType})
                </option>
              ))}
            </select>
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

        {/* Action Buttons Row */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-2 border-t border-slate-100">
          <div className="text-xs text-slate-500 tabular-nums">
            <span>{activeMonthName} {activeYear}</span>
            <span className="mx-1.5" aria-hidden="true">·</span>
            <span>{activeDates.length} hari aktif</span>
            <span className="mx-1.5" aria-hidden="true">·</span>
            <span>ID SPPG: {profile.idSPPG || "Belum diisi"}</span>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setConfirmResetModal(true)}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-semibold transition cursor-pointer whitespace-nowrap"
              title="Sinkronkan ulang isian tab ini dari Master Menu & Profil SPPG"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Sinkron Ulang Otomatis
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
      </div>

      {/* =================================================================== */}
      {/* SUB-TAB 1: UJI ORGANOLEPTIK (LAMPIRAN 3 — GAMBAR 1)                 */}
      {/* =================================================================== */}
      {activeSubTab === "organoleptik" && (
        <div className="space-y-4">
          {/* Date Selector Strip & Signature Mode */}
          <div className="bg-white rounded-2xl border border-slate-200 p-4 space-y-3 no-print">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <Calendar className="w-4 h-4 text-teal-600" />
                <span>Pilih Tanggal Uji Organoleptik ({activeMonthName} {activeYear})</span>
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
                    ? `Mode Semua Tanggal (${activeDates.length} Lembar)`
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

          {/* Printable Organoleptik Canvas */}
          <div className="bg-slate-100/80 p-4 sm:p-6 rounded-2xl border border-slate-200 overflow-x-auto print:bg-transparent print:p-0 print:border-none">
            <div
              id="print-area-sop-organoleptik"
              className="bg-white mx-auto w-full max-w-[215mm] min-w-[680px] p-6 sm:p-8 border border-slate-300 shadow-xs print:border-none print:shadow-none print:p-0"
            >
              {(organoPrintAll ? activeDates : [effectiveOrganoDate]).map((dStr, idx) =>
                renderOrganoleptikSheet(dStr, idx > 0)
              )}
            </div>
          </div>
        </div>
      )}

      {/* =================================================================== */}
      {/* SUB-TAB 2: PENGAMBILAN SAMPEL / CHECKLIST SAMPEL MAKANAN (GAMBAR 2) */}
      {/* =================================================================== */}
      {activeSubTab === "sampel" && (
        <div className="space-y-4">
          {/* Sub-Controls: Week Filter & Officer Teams */}
          <div className="bg-white rounded-2xl border border-slate-200 p-4 space-y-3 no-print">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs font-bold text-slate-700">Tampilkan Periode:</span>
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

          {/* Printable Checklist Sampel Makanan Canvas (exact Gambar 2 layout) */}
          <div className="bg-slate-100/80 p-4 sm:p-6 rounded-2xl border border-slate-200 overflow-x-auto print:bg-transparent print:p-0 print:border-none">
            <div
              id="print-area-sop-sampel"
              className="bg-white mx-auto w-full max-w-[215mm] min-w-[680px] p-6 sm:p-8 border border-slate-300 shadow-xs print:border-none print:shadow-none print:p-0 text-black"
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

              <p className="text-[11px] font-bold mb-1">Lampiran 2.</p>

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
                      const row = getSampelRow(dStr, idx);
                      return (
                        <tr key={dStr} className="align-top break-inside-avoid">
                          <td className="border border-black px-2 py-2 text-[10.5px] leading-snug tabular-nums">
                            {formatIndonesianDate(dStr)}
                          </td>
                          <td className="border border-black p-1.5">
                            <textarea
                              rows={3}
                              value={row.pengantaran || ""}
                              onChange={(e) =>
                                updateSampelRow(dStr, idx, { pengantaran: e.target.value })
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
                                updateSampelRow(dStr, idx, { namaPetugas: e.target.value })
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
                                updateSampelRow(dStr, idx, { namaMakanan: e.target.value })
                              }
                              className="w-full bg-transparent text-[10.5px] text-black resize-none focus:outline-none focus:bg-indigo-50/40 p-0.5 leading-snug"
                            />
                          </td>
                          <td
                            onClick={() =>
                              updateSampelRow(dStr, idx, { parafChecked: !row.parafChecked })
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
          </div>
        </div>
      )}

      {/* =================================================================== */}
      {/* SUB-TAB 3: PENYIMPANAN SAMPEL / CHILLER (GAMBAR 3)                  */}
      {/* =================================================================== */}
      {activeSubTab === "penyimpanan" && (
        <div className="space-y-4">
          {/* Defaults & Quick Actions Toolbar */}
          <div className="bg-white rounded-2xl border border-slate-200 p-4 space-y-3 no-print">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
              <div className="space-y-1">
                <label className="text-[11px] font-semibold text-slate-500 block">
                  Default Jumlah Sampel
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
                  Format TTD Bawah
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

          {/* Printable Penyimpanan Sampel Canvas (exact Gambar 3 layout + KOP + TTD + ID SPPG) */}
          <div className="bg-slate-100/80 p-4 sm:p-6 rounded-2xl border border-slate-200 overflow-x-auto print:bg-transparent print:p-0 print:border-none">
            <div
              id="print-area-sop-penyimpanan"
              className="bg-white mx-auto w-full max-w-[297mm] min-w-[820px] p-6 sm:p-8 border border-slate-300 shadow-xs print:border-none print:shadow-none print:p-0 text-black"
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

              <h2 className="text-center text-[13px] font-bold uppercase tracking-wide mb-3">
                FORM PENYIMPANAN SAMPEL MAKANAN (CHILLER)
              </h2>

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

              {/* Exact Gambar 3 Table with #D9D9D9 Shaded Header (bg placed strictly on th, never tr, to protect rowSpan cells) */}
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
                      const row = getPenyimpananRow(dStr);
                      return (
                        <tr key={dStr} className="break-inside-avoid">
                          <td className="border border-black px-2 py-1">
                            <input
                              type="text"
                              value={row.kodeIdentitas || ""}
                              onChange={(e) =>
                                updatePenyimpananRow(dStr, { kodeIdentitas: e.target.value })
                              }
                              className="w-full bg-transparent text-[10.5px] text-black focus:outline-none focus:bg-indigo-50/40 py-0.5"
                            />
                          </td>
                          <td className="border border-black px-1.5 py-1">
                            <input
                              type="text"
                              value={row.jumlahSampel || ""}
                              onChange={(e) =>
                                updatePenyimpananRow(dStr, { jumlahSampel: e.target.value })
                              }
                              className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                            />
                          </td>
                          <td
                            onClick={() =>
                              updatePenyimpananRow(dStr, {
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
                              updatePenyimpananRow(dStr, {
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
                                updatePenyimpananRow(dStr, {
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
                                updatePenyimpananRow(dStr, { masukTanggal: e.target.value })
                              }
                              className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                            />
                          </td>
                          <td className="border border-black px-1 py-1">
                            <input
                              type="text"
                              value={row.masukJam || ""}
                              onChange={(e) =>
                                updatePenyimpananRow(dStr, { masukJam: e.target.value })
                              }
                              className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                            />
                          </td>
                          <td className="border border-black px-1 py-1">
                            <input
                              type="text"
                              value={row.keluarTanggal || ""}
                              onChange={(e) =>
                                updatePenyimpananRow(dStr, { keluarTanggal: e.target.value })
                              }
                              className="w-full bg-transparent text-[10.5px] text-center text-black focus:outline-none focus:bg-indigo-50/40 py-0.5 tabular-nums"
                            />
                          </td>
                          <td className="border border-black px-1 py-1">
                            <input
                              type="text"
                              value={row.keluarJam || ""}
                              onChange={(e) =>
                                updatePenyimpananRow(dStr, { keluarJam: e.target.value })
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
          </div>
        </div>
      )}

      {/* Print Preview Modal */}
      {showPrintModal && (
        <PrintPreviewModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          elementId={currentPrintElementId}
          title={currentDocTitle}
          filename={currentDocFilename}
          defaultPaperSize={paperSize}
          defaultOrientation={activeSubTab === "penyimpanan" ? "landscape" : "portrait"}
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
                  Tindakan ini akan memuat ulang seluruh nama menu, tanggal otomatis, dan identitas dari Master Menu &amp; Profil SPPG pada tab SOP yang sedang aktif. Perubahan manual pada tab ini akan direset.
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
