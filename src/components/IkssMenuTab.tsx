/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useMemo, useState } from "react";
import {
  Camera,
  Image as ImageIcon,
  Printer,
  FileDown,
  Download,
  Eye,
  Trash2,
  RotateCcw,
  ClipboardList,
  Check,
  Layers,
  AlertCircle,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  DEFAULT_UNIFIED_KOP_CONFIG,
  FoodCostDay,
  HariPM,
  IkssDataState,
  MasterMenu,
  SPPGProfile,
  TKPIItem,
  UnifiedKopConfig,
} from "../types";
import { calculateDay, createTkpiMap } from "../utils/calc";
import {
  SASARAN_KATEGORI_LIST,
  composeMenuName,
  ensureMasterMenuPlans,
} from "../utils/menuSchedule";
import { TARGET_AKG_LIMITS } from "../tkpiData";
import { printElementSmart, exportToPDF } from "../lib/printUtils";
import { KopSuratRenderHeader, LogoCrop } from "./KopSuratConfigSection";
import AutoSignatureBlock from "./AutoSignatureBlock";
import PrintPreviewModal from "./PrintPreviewModal";

interface IkssMenuTabProps {
  profile: SPPGProfile;
  foodCostDays: FoodCostDay[];
  tkpiList: TKPIItem[];
  harianPM: HariPM[];
  masterMenu: MasterMenu;
  ikssData: IkssDataState;
  setIkssData: React.Dispatch<React.SetStateAction<IkssDataState>>;
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
  bumbuPct?: number;
}

const STANDARD_FOOD_COST_TYPES: Array<{
  id: FoodCostDay["jenisMenu"];
  label: string;
}> = [
  { id: "Basah", label: "Menu Basah" },
  { id: "Alergi", label: "Menu Alergi" },
  { id: "Kering", label: "Menu Kering" },
  { id: "MP-ASI", label: "Menu MP-ASI" },
  { id: "3B-Basah", label: "3B Basah" },
  { id: "3B-Alergi", label: "3B Alergi" },
  { id: "3B-Kering", label: "3B Kering" },
];

const PORSI_KECIL_IDS = new Set([
  "tk_paud_lb",
  "sd_kelas_1_3",
  "anak_balita",
  "anak_balita_13_59",
  "balita_6_11",
]);

const PORSI_BESAR_IDS = new Set([
  "sd_kelas_4_6",
  "smp_mts_smplb",
  "sma_smk_ma",
  "pendidik",
  "tenaga_kependidikan",
  "ibu_hamil",
  "ibu_menyusui",
]);

interface AkgPresetTarget {
  label: string;
  energi: number;
  protein: number;
  lemak: number;
  kh: number;
  serat: number;
}

// Reference AKG 20-25% (Porsi Kecil) & 30-35% (Porsi Besar) per sasaran ID
const SASARAN_AKG_REF: Record<string, AkgPresetTarget> = {
  balita_6_11: {
    label: "Rujukan AKG 20-25%",
    energi: 225,
    protein: 6.5,
    lemak: 8.5,
    kh: 30,
    serat: 2.5,
  },
  anak_balita: {
    label: "Rujukan AKG 20-25%",
    energi: TARGET_AKG_LIMITS.balita.energiMin,
    protein: TARGET_AKG_LIMITS.balita.proteinMin,
    lemak: TARGET_AKG_LIMITS.balita.lemakMin,
    kh: TARGET_AKG_LIMITS.balita.khMin,
    serat: 4.5,
  },
  anak_balita_13_59: {
    label: "Rujukan AKG 20-25%",
    energi: TARGET_AKG_LIMITS.balita.energiMax,
    protein: TARGET_AKG_LIMITS.balita.proteinMin,
    lemak: TARGET_AKG_LIMITS.balita.lemakMin,
    kh: TARGET_AKG_LIMITS.balita.khMin,
    serat: 4.8,
  },
  tk_paud_lb: {
    label: "Rujukan AKG 20-25%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.tk_paud.energiMin + TARGET_AKG_LIMITS.tk_paud.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.tk_paud.proteinMin,
    lemak: TARGET_AKG_LIMITS.tk_paud.lemakMin,
    kh: TARGET_AKG_LIMITS.tk_paud.khMin,
    serat: 5.0,
  },
  sd_kelas_1_3: {
    label: "Rujukan AKG 20-25%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.sd_kecil.energiMin + TARGET_AKG_LIMITS.sd_kecil.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.sd_kecil.proteinMin,
    lemak: TARGET_AKG_LIMITS.sd_kecil.lemakMin,
    kh: TARGET_AKG_LIMITS.sd_kecil.khMin,
    serat: 5.8,
  },
  sd_kelas_4_6: {
    label: "Rujukan AKG 30-35%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.sd_besar.energiMin + TARGET_AKG_LIMITS.sd_besar.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.sd_besar.proteinMin,
    lemak: TARGET_AKG_LIMITS.sd_besar.lemakMin,
    kh: TARGET_AKG_LIMITS.sd_besar.khMin,
    serat: 8.5,
  },
  smp_mts_smplb: {
    label: "Rujukan AKG 30-35%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.smp.energiMin + TARGET_AKG_LIMITS.smp.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.smp.proteinMin,
    lemak: TARGET_AKG_LIMITS.smp.lemakMin,
    kh: TARGET_AKG_LIMITS.smp.khMin,
    serat: 10.0,
  },
  sma_smk_ma: {
    label: "Rujukan AKG 30-35%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.sma_pendidik.energiMin + TARGET_AKG_LIMITS.sma_pendidik.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.sma_pendidik.proteinMin,
    lemak: TARGET_AKG_LIMITS.sma_pendidik.lemakMin,
    kh: TARGET_AKG_LIMITS.sma_pendidik.khMin,
    serat: 10.5,
  },
  pendidik: {
    label: "Rujukan AKG 30-35%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.sma_pendidik.energiMin + TARGET_AKG_LIMITS.sma_pendidik.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.sma_pendidik.proteinMin,
    lemak: TARGET_AKG_LIMITS.sma_pendidik.lemakMin,
    kh: TARGET_AKG_LIMITS.sma_pendidik.khMin,
    serat: 10.5,
  },
  tenaga_kependidikan: {
    label: "Rujukan AKG 30-35%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.sma_pendidik.energiMin + TARGET_AKG_LIMITS.sma_pendidik.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.sma_pendidik.proteinMin,
    lemak: TARGET_AKG_LIMITS.sma_pendidik.lemakMin,
    kh: TARGET_AKG_LIMITS.sma_pendidik.khMin,
    serat: 10.5,
  },
  ibu_hamil: {
    label: "Rujukan AKG 30-35%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.bumil.energiMin + TARGET_AKG_LIMITS.bumil.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.bumil.proteinMin,
    lemak: TARGET_AKG_LIMITS.bumil.lemakMin,
    kh: TARGET_AKG_LIMITS.bumil.khMin,
    serat: 11.0,
  },
  ibu_menyusui: {
    label: "Rujukan AKG 30-35%",
    energi: Math.round(
      (TARGET_AKG_LIMITS.busui.energiMin + TARGET_AKG_LIMITS.busui.energiMax) / 2
    ),
    protein: TARGET_AKG_LIMITS.busui.proteinMin,
    lemak: TARGET_AKG_LIMITS.busui.lemakMin,
    kh: TARGET_AKG_LIMITS.busui.khMin,
    serat: 11.5,
  },
};

const QUICK_PRESETS: Array<{
  id: string;
  label: string;
  title: string;
  sasaranIds: string[];
  foodCostTypes: string[];
}> = [
  {
    id: "balita",
    label: "Balita 12-23 Bln / Balita",
    title: "BALITA 12-23 BULAN/BALITA",
    sasaranIds: ["anak_balita", "anak_balita_13_59"],
    foodCostTypes: ["Basah", "3B-Basah"],
  },
  {
    id: "porsi_kecil",
    label: "Porsi Kecil Gabungan",
    title: "PAUD/TK, SD 1-3 & BALITA (PORSI KECIL)",
    sasaranIds: ["tk_paud_lb", "sd_kelas_1_3", "anak_balita", "anak_balita_13_59"],
    foodCostTypes: ["Basah", "3B-Basah"],
  },
  {
    id: "porsi_besar",
    label: "Porsi Besar Gabungan",
    title: "SD 4-6, SMP, SMA/SMK & PENDIDIK (PORSI BESAR)",
    sasaranIds: [
      "sd_kelas_4_6",
      "smp_mts_smplb",
      "sma_smk_ma",
      "pendidik",
      "tenaga_kependidikan",
    ],
    foodCostTypes: ["Basah"],
  },
  {
    id: "3b",
    label: "Ibu Hamil & Menyusui (3B)",
    title: "IBU HAMIL & IBU MENYUSUI (3B)",
    sasaranIds: ["ibu_hamil", "ibu_menyusui"],
    foodCostTypes: ["3B-Basah", "Basah"],
  },
  {
    id: "mpasi",
    label: "MP-ASI (6-11 Bln)",
    title: "BALITA 6-11 BULAN (MP-ASI)",
    sasaranIds: ["balita_6_11"],
    foodCostTypes: ["MP-ASI", "3B-Basah"],
  },
  {
    id: "all",
    label: "Semua Kategori",
    title: "SEMUA KATEGORI PENERIMA MANFAAT",
    sasaranIds: SASARAN_KATEGORI_LIST.map((s) => s.id),
    foodCostTypes: STANDARD_FOOD_COST_TYPES.map((t) => t.id),
  },
];

function formatIndoFullDate(dateStr?: string): string {
  if (!dateStr) return "-";
  try {
    const clean = dateStr.trim();
    const parts = clean.split("-");
    const d =
      parts.length === 3
        ? new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]))
        : new Date(clean);
    if (isNaN(d.getTime())) return dateStr;
    const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
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
    return `${days[d.getDay()]}, ${String(d.getDate()).padStart(2, "0")} ${
      months[d.getMonth()]
    } ${d.getFullYear()}`;
  } catch {
    return dateStr;
  }
}

function derivePeriodeBulanLabel(dates: string[]): string {
  const valid = (dates || []).filter((d) => /^\d{4}-\d{2}-\d{2}/.test(d || ""));
  if (valid.length === 0) {
    const now = new Date();
    return now
      .toLocaleDateString("id-ID", { month: "long", year: "numeric" })
      .toUpperCase();
  }
  const months = [
    "JANUARI",
    "FEBRUARI",
    "MARET",
    "APRIL",
    "MEI",
    "JUNI",
    "JULI",
    "AGUSTUS",
    "SEPTEMBER",
    "OKTOBER",
    "NOVEMBER",
    "DESEMBER",
  ];
  const parsed = valid.map((d) => {
    const [y, m] = d.split("-").map(Number);
    return { y, m: m - 1 };
  });
  const first = parsed[0];
  const last = parsed[parsed.length - 1];
  if (first.y === last.y && first.m === last.m) {
    return `${months[first.m]} ${first.y}`;
  }
  if (first.y === last.y) {
    return `${months[first.m]} SD ${months[last.m]} ${first.y}`;
  }
  return `${months[first.m]} ${first.y} SD ${months[last.m]} ${last.y}`;
}

function compressPhotoFile(file: File, maxWidth = 380, quality = 0.72): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const scale = img.width > maxWidth ? maxWidth / img.width : 1;
        const w = Math.max(1, Math.round(img.width * scale));
        const h = Math.max(1, Math.round(img.height * scale));
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          reject(new Error("Gagal memproses gambar."));
          return;
        }
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => reject(new Error("File gambar tidak valid."));
      img.src = String(reader.result || "");
    };
    reader.onerror = () => reject(new Error("Gagal membaca file gambar."));
    reader.readAsDataURL(file);
  });
}

function formatDec1(val: number): string {
  if (!Number.isFinite(val)) return "0.0";
  return (Math.round(val * 10) / 10).toFixed(1);
}

export default function IkssMenuTab({
  profile,
  foodCostDays,
  tkpiList,
  harianPM,
  masterMenu,
  ikssData,
  setIkssData,
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
  bumbuPct = 10,
}: IkssMenuTabProps) {
  const [subView, setSubView] = useState<"nilai_gizi" | "foto_menu">("nilai_gizi");
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [photoError, setPhotoError] = useState<string | null>(null);
  const [showCategoryDrawer, setShowCategoryDrawer] = useState(true);

  const tkpiMap = useMemo(() => createTkpiMap(tkpiList), [tkpiList]);

  // Read any custom Food Cost tables stored by FoodCostTab
  const customFcTables = useMemo(() => {
    try {
      const raw = localStorage.getItem("sppg_custom_fc_tables_v1");
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [] as any[];
  }, [foodCostDays]);

  const selectedFoodCostTypes =
    ikssData.selectedFoodCostTypes && ikssData.selectedFoodCostTypes.length > 0
      ? ikssData.selectedFoodCostTypes
      : ["Basah", "3B-Basah"];

  const selectedSasaranIds =
    ikssData.selectedSasaranIds && ikssData.selectedSasaranIds.length > 0
      ? ikssData.selectedSasaranIds
      : ["anak_balita", "anak_balita_13_59"];

  // Compute default category label dynamically from selected sasaran
  const autoCategoryTitle = useMemo(() => {
    const matchedPreset = QUICK_PRESETS.find(
      (p) =>
        p.sasaranIds.length === selectedSasaranIds.length &&
        p.sasaranIds.every((id) => selectedSasaranIds.includes(id))
    );
    if (matchedPreset) return matchedPreset.title;
    const labels = SASARAN_KATEGORI_LIST.filter((s) =>
      selectedSasaranIds.includes(s.id)
    ).map((s) => s.shortLabel.toUpperCase());
    if (labels.length === 0) return "SEMUA KATEGORI";
    if (labels.length <= 3) return labels.join(" / ");
    return `${labels.slice(0, 3).join(", ")} +${labels.length - 3} LAINNYA`;
  }, [selectedSasaranIds]);

  const displayCategoryTitle = (
    ikssData.customCategoryTitle?.trim() || autoCategoryTitle
  ).toUpperCase();

  // Period/Month automatically derived from profile.periodeDates
  const autoPeriodeBulan = useMemo(
    () => derivePeriodeBulanLabel(profile.periodeDates || []),
    [profile.periodeDates]
  );
  const displayPeriodeBulan = (
    ikssData.customPeriodeLabel?.trim() || autoPeriodeBulan
  ).toUpperCase();

  // Group key for storing photos per selected category combination (shared across same category selection)
  const photoGroupKey = useMemo(() => {
    const sorted = [...selectedSasaranIds].sort().join("_");
    return sorted || "default";
  }, [selectedSasaranIds]);

  const totalDays = Math.max(
    profile.periodeDates?.length || 0,
    harianPM?.length || 0,
    10
  );

  // Normalized master menu plans for rich menu name fallback
  const normalizedMenu = useMemo(
    () => ensureMasterMenuPlans(masterMenu),
    [masterMenu]
  );

  // Check which Food Cost types actually have ingredients in the current planning
  const populatedFcTypes = useMemo(() => {
    const set = new Set<string>();
    foodCostDays.forEach((d) => {
      const hasData =
        (d.porsiBesarBahan || []).some((b) => (b.beratBB || 0) > 0 && b.nama?.trim()) ||
        (d.porsiKecilBahan || []).some((b) => (b.beratBB || 0) > 0 && b.nama?.trim());
      if (hasData) set.add(d.jenisMenu);
    });
    customFcTables.forEach((t: any) => {
      if (Array.isArray(t.bahanList) && t.bahanList.some((b: any) => (b.beratBB || 0) > 0)) {
        set.add(`custom:${t.id}`);
      }
    });
    return set;
  }, [foodCostDays, customFcTables]);

  // Core Daily IKSS Rows computed automatically from Food Cost + Penerima Manfaat + Profil SPPG Dates
  const dailyRows = useMemo(() => {
    const hasKecilSelected = selectedSasaranIds.some((id) => PORSI_KECIL_IDS.has(id));
    const hasBesarSelected = selectedSasaranIds.some((id) => PORSI_BESAR_IDS.has(id));
    const isAlergiOnly =
      selectedFoodCostTypes.length > 0 &&
      selectedFoodCostTypes.every((t) => t.toLowerCase().includes("alergi"));

    const rows: Array<{
      hariKe: number;
      dateStr: string;
      formattedDate: string;
      namaMenu: string;
      energi: number;
      protein: number;
      lemak: number;
      kh: number;
      serat: number;
      jumlahManfaat: number;
      hasNutrition: boolean;
    }> = [];

    for (let i = 0; i < totalDays; i++) {
      const hariKe = i + 1;
      const dateStr = profile.periodeDates?.[i] || "";
      const formattedDate = dateStr
        ? formatIndoFullDate(dateStr)
        : `Hari Ke-${hariKe}`;

      // 1. Beneficiaries for selected Sasaran IDs on hariKe
      const dayPM = harianPM.find((h) => h.hariKe === hariKe) || harianPM[0];
      let pmKecilCount = 0;
      let pmBesarCount = 0;
      selectedSasaranIds.forEach((sId) => {
        const sas = dayPM?.sasaran?.find((s) => s.id === sId);
        if (!sas) return;
        const count = isAlergiOnly
          ? (Number(sas.alergiKecil) || 0) + (Number(sas.alergiBesar) || 0)
          : (Number(sas.porsiKecil) || 0) + (Number(sas.porsiBesar) || 0);
        if (PORSI_KECIL_IDS.has(sId)) {
          pmKecilCount += count;
        } else {
          pmBesarCount += count;
        }
      });

      const jumlahManfaat = pmKecilCount + pmBesarCount;

      // 2. Match FoodCostDays for hariKe & selectedFoodCostTypes
      const matchedDays = foodCostDays.filter(
        (fc) => fc.hariKe === hariKe && selectedFoodCostTypes.includes(fc.jenisMenu)
      );

      // Compute nutrition entries from matched standard FoodCostDays
      const nutriSamples: Array<{
        energi: number;
        protein: number;
        lemak: number;
        kh: number;
        serat: number;
        weight: number;
      }> = [];
      const menuNameParts: string[] = [];

      matchedDays.forEach((fc) => {
        const hasBesarBahan = (fc.porsiBesarBahan || []).some(
          (b) => (Number(b.beratBB) || 0) > 0 && (b.nama?.trim() || b.tkpiId)
        );
        const hasKecilBahan = (fc.porsiKecilBahan || []).some(
          (b) => (Number(b.beratBB) || 0) > 0 && (b.nama?.trim() || b.tkpiId)
        );
        if (!hasBesarBahan && !hasKecilBahan) return;

        const calc = calculateDay(
          fc.porsiBesarBahan || [],
          fc.porsiKecilBahan || [],
          Math.max(pmBesarCount, 1),
          Math.max(pmKecilCount, 1),
          fc.bufferPct ?? 5,
          tkpiMap,
          fc.bumbuPct ?? bumbuPct
        );

        const toSample = (n: Record<string, number>, weight: number) => ({
          energi: Number(n.energi) || 0,
          protein: Number(n.protein) || 0,
          lemak: Number(n.lemak) || 0,
          kh: Number(n.kh) || 0,
          serat: Number(n.serat) || 0,
          weight,
        });

        const nKecil = calc.nutrisiPorsiKecil;
        const nBesar = calc.nutrisiPorsiBesar;

        if (hasKecilSelected && hasBesarSelected) {
          if (hasKecilBahan && hasBesarBahan) {
            const wK = pmKecilCount > 0 ? pmKecilCount : 1;
            const wB = pmBesarCount > 0 ? pmBesarCount : 1;
            const totalW = wK + wB;
            nutriSamples.push({
              energi: ((Number(nKecil.energi) || 0) * wK + (Number(nBesar.energi) || 0) * wB) / totalW,
              protein: ((Number(nKecil.protein) || 0) * wK + (Number(nBesar.protein) || 0) * wB) / totalW,
              lemak: ((Number(nKecil.lemak) || 0) * wK + (Number(nBesar.lemak) || 0) * wB) / totalW,
              kh: ((Number(nKecil.kh) || 0) * wK + (Number(nBesar.kh) || 0) * wB) / totalW,
              serat: ((Number(nKecil.serat) || 0) * wK + (Number(nBesar.serat) || 0) * wB) / totalW,
              weight: totalW,
            });
          } else if (hasKecilBahan) {
            nutriSamples.push(toSample(nKecil, Math.max(pmKecilCount, 1)));
          } else if (hasBesarBahan) {
            nutriSamples.push(toSample(nBesar, Math.max(pmBesarCount, 1)));
          }
        } else if (hasKecilSelected) {
          const chosen = hasKecilBahan ? nKecil : nBesar;
          nutriSamples.push(toSample(chosen, Math.max(pmKecilCount, 1)));
        } else {
          const chosen = hasBesarBahan ? nBesar : nKecil;
          nutriSamples.push(toSample(chosen, Math.max(pmBesarCount, 1)));
        }

        if (fc.namaMenu && fc.namaMenu.trim()) {
          menuNameParts.push(fc.namaMenu.trim());
        } else {
          const sourceList =
            hasKecilSelected && !hasBesarSelected && hasKecilBahan
              ? fc.porsiKecilBahan
              : hasBesarBahan
              ? fc.porsiBesarBahan
              : fc.porsiKecilBahan;
          const names = (sourceList || [])
            .map((b) => (b.nama || "").trim())
            .filter((n) => n && n.toLowerCase() !== "ketik bahan...");
          if (names.length > 0) {
            menuNameParts.push(names.slice(0, 6).join(", "));
          }
        }
      });

      // Also include any selected custom Food Cost tables
      customFcTables.forEach((tbl: any) => {
        if (!selectedFoodCostTypes.includes(`custom:${tbl.id}`)) return;
        if (!Array.isArray(tbl.bahanList) || tbl.bahanList.length === 0) return;
        const isDayMatch = tbl.hariKe === undefined || Number(tbl.hariKe) === hariKe;
        if (!isDayMatch) return;
        const calc =
          tbl.porsi === "besar"
            ? calculateDay(tbl.bahanList, [], Math.max(tbl.pmCount || 1, 1), 0, tbl.bufferPct || 5, tkpiMap, bumbuPct)
            : calculateDay([], tbl.bahanList, 0, Math.max(tbl.pmCount || 1, 1), tbl.bufferPct || 5, tkpiMap, bumbuPct);
        const nut = tbl.porsi === "besar" ? calc.nutrisiPorsiBesar : calc.nutrisiPorsiKecil;
        const eVal = Number(nut.energi) || 0;
        const pVal = Number(nut.protein) || 0;
        if (eVal > 0 || pVal > 0) {
          nutriSamples.push({
            energi: eVal,
            protein: pVal,
            lemak: Number(nut.lemak) || 0,
            kh: Number(nut.kh) || 0,
            serat: Number(nut.serat) || 0,
            weight: Math.max(tbl.pmCount || 1, 1),
          });
        }
        if (tbl.namaMenu) menuNameParts.push(String(tbl.namaMenu));
      });

      // Resolve rich menu description from Master Menu plans matching selected Sasaran if available
      const masterMenuNames: string[] = [];
      (normalizedMenu.plans || []).forEach((plan) => {
        const overlaps =
          (plan.sasaranIds || []).some((id) => selectedSasaranIds.includes(id)) ||
          selectedSasaranIds.length === 0;
        if (!overlaps) return;
        const dayItem = plan.days?.[i];
        if (!dayItem) return;
        const composed = composeMenuName(dayItem).trim();
        if (composed && !masterMenuNames.includes(composed)) {
          masterMenuNames.push(composed);
        }
      });

      // Prefer rich Master Menu composition or Food Cost namaMenu
      let namaMenu = "";
      if (masterMenuNames.length > 0) {
        namaMenu = masterMenuNames.join(" | ");
      } else if (menuNameParts.length > 0) {
        const uniqueParts = Array.from(new Set(menuNameParts));
        namaMenu = uniqueParts.join(" | ");
      }

      let energi = 0;
      let protein = 0;
      let lemak = 0;
      let kh = 0;
      let serat = 0;
      const hasNutrition = nutriSamples.length > 0;

      if (hasNutrition) {
        const sumW = nutriSamples.reduce((a, b) => a + b.weight, 0) || nutriSamples.length;
        energi = nutriSamples.reduce((a, b) => a + b.energi * b.weight, 0) / sumW;
        protein = nutriSamples.reduce((a, b) => a + b.protein * b.weight, 0) / sumW;
        lemak = nutriSamples.reduce((a, b) => a + b.lemak * b.weight, 0) / sumW;
        kh = nutriSamples.reduce((a, b) => a + b.kh * b.weight, 0) / sumW;
        serat = nutriSamples.reduce((a, b) => a + b.serat * b.weight, 0) / sumW;
      }

      rows.push({
        hariKe,
        dateStr,
        formattedDate,
        namaMenu,
        energi,
        protein,
        lemak,
        kh,
        serat,
        jumlahManfaat,
        hasNutrition,
      });
    }

    return rows;
  }, [
    totalDays,
    profile.periodeDates,
    harianPM,
    foodCostDays,
    customFcTables,
    selectedFoodCostTypes,
    selectedSasaranIds,
    tkpiMap,
    bumbuPct,
    normalizedMenu,
  ]);

  // Compute Auto Rujukan AKG averaged across selected Sasaran IDs
  const autoAkg = useMemo<AkgPresetTarget>(() => {
    const matched = selectedSasaranIds
      .map((id) => SASARAN_AKG_REF[id])
      .filter(Boolean);
    if (matched.length === 0) {
      return {
        label: "Rujukan AKG 20-25%",
        energi: 350,
        protein: 10,
        lemak: 12,
        kh: 50,
        serat: 4.5,
      };
    }
    const allKecil = selectedSasaranIds.every((id) => PORSI_KECIL_IDS.has(id));
    const allBesar = selectedSasaranIds.every((id) => PORSI_BESAR_IDS.has(id));
    const label = allKecil
      ? "Rujukan AKG 20-25%"
      : allBesar
      ? "Rujukan AKG 30-35%"
      : "Rujukan AKG Gabungan";

    const n = matched.length;
    return {
      label,
      energi: Math.round((matched.reduce((a, b) => a + b.energi, 0) / n) * 10) / 10,
      protein: Math.round((matched.reduce((a, b) => a + b.protein, 0) / n) * 10) / 10,
      lemak: Math.round((matched.reduce((a, b) => a + b.lemak, 0) / n) * 10) / 10,
      kh: Math.round((matched.reduce((a, b) => a + b.kh, 0) / n) * 10) / 10,
      serat: Math.round((matched.reduce((a, b) => a + b.serat, 0) / n) * 10) / 10,
    };
  }, [selectedSasaranIds]);

  const effectiveAkg = {
    label: ikssData.akgOverride?.label || autoAkg.label,
    energi:
      ikssData.akgOverride?.energi !== undefined
        ? ikssData.akgOverride.energi
        : autoAkg.energi,
    protein:
      ikssData.akgOverride?.protein !== undefined
        ? ikssData.akgOverride.protein
        : autoAkg.protein,
    lemak:
      ikssData.akgOverride?.lemak !== undefined
        ? ikssData.akgOverride.lemak
        : autoAkg.lemak,
    kh:
      ikssData.akgOverride?.kh !== undefined
        ? ikssData.akgOverride.kh
        : autoAkg.kh,
    serat:
      ikssData.akgOverride?.serat !== undefined
        ? ikssData.akgOverride.serat
        : autoAkg.serat,
  };

  // Summary Aggregates (Exact Gambar 2 bottom block)
  const summary = useMemo(() => {
    const totalEnergi = dailyRows.reduce((s, r) => s + r.energi, 0);
    const totalProtein = dailyRows.reduce((s, r) => s + r.protein, 0);
    const totalLemak = dailyRows.reduce((s, r) => s + r.lemak, 0);
    const totalKh = dailyRows.reduce((s, r) => s + r.kh, 0);
    const totalSerat = dailyRows.reduce((s, r) => s + r.serat, 0);
    const totalManfaat = dailyRows.reduce((s, r) => s + r.jumlahManfaat, 0);

    const activeNutriDays = dailyRows.filter((r) => r.hasNutrition && r.energi > 0).length;
    const activePmDays = dailyRows.filter((r) => r.jumlahManfaat > 0).length;

    const avgEnergi = activeNutriDays > 0 ? totalEnergi / activeNutriDays : 0;
    const avgProtein = activeNutriDays > 0 ? totalProtein / activeNutriDays : 0;
    const avgLemak = activeNutriDays > 0 ? totalLemak / activeNutriDays : 0;
    const avgKh = activeNutriDays > 0 ? totalKh / activeNutriDays : 0;
    const avgSerat = activeNutriDays > 0 ? totalSerat / activeNutriDays : 0;
    const avgManfaat = activePmDays > 0 ? Math.round(totalManfaat / activePmDays) : 0;

    const pct = (avg: number, ref: number) =>
      ref > 0 ? Math.round((avg / ref) * 100) : 0;

    return {
      totalEnergi,
      totalProtein,
      totalLemak,
      totalKh,
      totalSerat,
      totalManfaat,
      activeNutriDays,
      activePmDays,
      avgEnergi,
      avgProtein,
      avgLemak,
      avgKh,
      avgSerat,
      avgManfaat,
      pctEnergi: pct(avgEnergi, effectiveAkg.energi),
      pctProtein: pct(avgProtein, effectiveAkg.protein),
      pctLemak: pct(avgLemak, effectiveAkg.lemak),
      pctKh: pct(avgKh, effectiveAkg.kh),
      pctSerat: pct(avgSerat, effectiveAkg.serat),
    };
  }, [dailyRows, effectiveAkg]);

  // Handlers
  const toggleFoodCostType = (typeId: string) => {
    setIkssData((prev) => {
      const current = prev.selectedFoodCostTypes || [];
      const exists = current.includes(typeId);
      const next = exists
        ? current.filter((x) => x !== typeId)
        : [...current, typeId];
      return {
        ...prev,
        selectedFoodCostTypes: next.length > 0 ? next : [typeId],
      };
    });
  };

  const toggleSasaranId = (sasId: string) => {
    setIkssData((prev) => {
      const current = prev.selectedSasaranIds || [];
      const exists = current.includes(sasId);
      const next = exists
        ? current.filter((x) => x !== sasId)
        : [...current, sasId];
      return {
        ...prev,
        selectedSasaranIds: next.length > 0 ? next : [sasId],
        customCategoryTitle: "",
        akgOverride: undefined,
      };
    });
  };

  const applyPreset = (preset: (typeof QUICK_PRESETS)[number]) => {
    setIkssData((prev) => ({
      ...prev,
      selectedSasaranIds: [...preset.sasaranIds],
      selectedFoodCostTypes: [...preset.foodCostTypes],
      customCategoryTitle: preset.title,
      akgOverride: undefined,
    }));
  };

  const handlePhotoUpload = async (
    hariKe: number,
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setPhotoError(null);
    try {
      const compressed = await compressPhotoFile(file, 380, 0.72);
      const key = `${photoGroupKey}__day_${hariKe}`;
      setIkssData((prev) => {
        const nextPhotos = { ...(prev.photos || {}), [key]: compressed };
        const approxSize = JSON.stringify(nextPhotos).length;
        if (approxSize > 860000) {
          setPhotoError(
            "Kapasitas penyimpanan foto cloud hampir penuh. Hapus foto periode lama yang tidak diperlukan agar sinkronisasi tetap cepat."
          );
          return prev;
        }
        return { ...prev, photos: nextPhotos };
      });
    } catch (err: any) {
      setPhotoError(err?.message || "Gagal mengunggah foto menu.");
    }
  };

  const handleDeletePhoto = (hariKe: number) => {
    const key = `${photoGroupKey}__day_${hariKe}`;
    setIkssData((prev) => {
      const next = { ...(prev.photos || {}) };
      delete next[key];
      return { ...prev, photos: next };
    });
  };

  const getPhotoForDay = (hariKe: number): string => {
    const exact = ikssData.photos?.[`${photoGroupKey}__day_${hariKe}`];
    if (exact) return exact;
    // Also fallback to generic day photo if uploaded under a single day key
    return ikssData.photos?.[`day_${hariKe}`] || "";
  };

  const activePrintId =
    subView === "nilai_gizi" ? "print-area-ikss-gizi" : "print-area-ikss-foto";
  const activePrintTitle =
    subView === "nilai_gizi"
      ? `IKSS Rekapan Nilai Gizi - ${displayCategoryTitle}`
      : `IKSS Rekapan Foto Menu - ${displayCategoryTitle}`;
  const activeFilename =
    subView === "nilai_gizi"
      ? `IKSS_Nilai_Gizi_${displayCategoryTitle.replace(/[^A-Za-z0-9]+/g, "_")}`
      : `IKSS_Foto_Menu_${displayCategoryTitle.replace(/[^A-Za-z0-9]+/g, "_")}`;

  const handlePrint = () => {
    printElementSmart(activePrintId, {
      title: activePrintTitle,
      paperSize,
      orientation: "auto",
    });
  };

  const handleExportPdf = async () => {
    await exportToPDF(activePrintId, activeFilename, 100, undefined, {
      paperSize,
      orientation: "auto",
    });
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const identityRows = [
      [
        subView === "nilai_gizi"
          ? `REKAPAN NILAI GIZI ${displayCategoryTitle}`
          : `REKAPAN MENU ${displayCategoryTitle}`,
      ],
      ["Nama SPPG", ":", (profile.namaLembaga || "").toUpperCase()],
      ["ID SPPG", ":", (profile.idSPPG || "").toUpperCase()],
      ["Alamat", ":", profile.alamat || ""],
      ["Kab/Kota", ":", (profile.kabKota || "").toUpperCase()],
      ["Provinsi", ":", (profile.provinsi || "").toUpperCase()],
      ["Nama SPPI/AG", ":", profile.namaAhliGizi || ""],
      ["Nomor Handphone", ":", profile.noHpSPPI || ""],
      ["Periode/Bulan", ":", displayPeriodeBulan],
      [],
    ];

    if (subView === "nilai_gizi") {
      const sheetData: any[][] = [
        ...identityRows,
        [
          "Hari / Tanggal",
          "Nama Menu",
          "Energi (kkal)",
          "Protein (gram)",
          "Lemak (gram)",
          "Karbohidrat (gram)",
          "Serat (gram)",
          "Jumlah Manfaat",
        ],
        ...dailyRows.map((r) => [
          r.formattedDate,
          r.namaMenu || "-",
          Number(formatDec1(r.energi)),
          Number(formatDec1(r.protein)),
          Number(formatDec1(r.lemak)),
          Number(formatDec1(r.kh)),
          Number(formatDec1(r.serat)),
          r.jumlahManfaat,
        ]),
        [
          "Total Nilai Gizi",
          "",
          Number(formatDec1(summary.totalEnergi)),
          Number(formatDec1(summary.totalProtein)),
          Number(formatDec1(summary.totalLemak)),
          Number(formatDec1(summary.totalKh)),
          Number(formatDec1(summary.totalSerat)),
          summary.totalManfaat,
        ],
        [
          "Jumlah Hari",
          "",
          `${summary.activeNutriDays} Hari`,
          `${summary.activeNutriDays} Hari`,
          `${summary.activeNutriDays} Hari`,
          `${summary.activeNutriDays} Hari`,
          `${summary.activeNutriDays} Hari`,
          `${summary.activePmDays} Hari`,
        ],
        [
          "Nilai Rata-Rata",
          "",
          Number(formatDec1(summary.avgEnergi)),
          Number(formatDec1(summary.avgProtein)),
          Number(formatDec1(summary.avgLemak)),
          Number(formatDec1(summary.avgKh)),
          Number(formatDec1(summary.avgSerat)),
          `${summary.avgManfaat} Orang`,
        ],
        [
          effectiveAkg.label,
          "",
          effectiveAkg.energi,
          effectiveAkg.protein,
          effectiveAkg.lemak,
          effectiveAkg.kh,
          effectiveAkg.serat,
          "",
        ],
        [
          "% Kebutuhan 90-110%",
          "",
          `${summary.pctEnergi}%`,
          `${summary.pctProtein}%`,
          `${summary.pctLemak}%`,
          `${summary.pctKh}%`,
          `${summary.pctSerat}%`,
          "",
        ],
      ];
      const ws = XLSX.utils.aoa_to_sheet(sheetData);
      ws["!cols"] = [
        { wch: 24 },
        { wch: 44 },
        { wch: 14 },
        { wch: 14 },
        { wch: 14 },
        { wch: 16 },
        { wch: 13 },
        { wch: 15 },
      ];
      XLSX.utils.book_append_sheet(wb, ws, "IKSS Nilai Gizi");
    } else {
      const sheetData: any[][] = [
        ...identityRows,
        [
          "Hari Ke",
          "Nama SPPG",
          "ID SPPG",
          "Hari / Tanggal",
          "Nama Menu",
          "Status Foto Menu",
        ],
        ...dailyRows.map((r) => [
          r.hariKe,
          (profile.namaLembaga || "").toUpperCase(),
          (profile.idSPPG || "").toUpperCase(),
          r.formattedDate,
          r.namaMenu || "-",
          getPhotoForDay(r.hariKe) ? "Foto Terlampir (Cetak PDF/Print)" : "Belum Diunggah",
        ]),
      ];
      const ws = XLSX.utils.aoa_to_sheet(sheetData);
      ws["!cols"] = [
        { wch: 10 },
        { wch: 32 },
        { wch: 14 },
        { wch: 24 },
        { wch: 44 },
        { wch: 24 },
      ];
      XLSX.utils.book_append_sheet(wb, ws, "IKSS Foto Menu");
    }

    XLSX.writeFile(wb, `${activeFilename}.xlsx`);
  };

  const renderIdentityBlock = () => (
    <table
      className="w-full text-[11px] leading-snug border-x border-black bg-white text-black"
      style={{ fontFamily: "Arial, Helvetica, sans-serif", borderCollapse: "collapse" }}
    >
      <tbody>
        {[
          { label: "Nama SPPG", val: (profile.namaLembaga || "-").toUpperCase() },
          { label: "ID SPPG", val: (profile.idSPPG || "-").toUpperCase() },
          { label: "Alamat", val: profile.alamat || "-" },
          { label: "Kab/Kota", val: (profile.kabKota || "-").toUpperCase() },
          { label: "Provinsi", val: (profile.provinsi || "-").toUpperCase() },
          { label: "Nama SPPI/AG", val: profile.namaAhliGizi || "-" },
          { label: "Nomor Handphone", val: profile.noHpSPPI || "-" },
          { label: "Periode/Bulan", val: displayPeriodeBulan },
        ].map((item) => (
          <tr key={item.label}>
            <td className="px-2 py-0.5 font-bold w-[165px] align-top">{item.label}</td>
            <td className="px-1 py-0.5 font-bold w-[14px] text-center align-top">:</td>
            <td className="px-1.5 py-0.5 font-bold align-top">{item.val}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderPctCell = (pctVal: number) => {
    const isGood = pctVal >= 90 && pctVal <= 110;
    return (
      <td
        className={`border border-black px-1.5 py-1.5 text-center font-bold tabular-nums text-[11px] ${
          isGood ? "bg-[#00B050] text-black" : "bg-[#FF0000] text-black"
        }`}
      >
        {pctVal > 0 ? `${pctVal}%` : "0"}
      </td>
    );
  };

  return (
    <div id="ikss-menu-tab" className="space-y-5 font-sans">
      {/* TOP CONTROL BAR */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4 no-print">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <h2 className="text-lg font-extrabold text-slate-900 flex items-center gap-2">
              <ClipboardList className="w-5 h-5 text-emerald-600" />
              Menu IKSS — Rekapan Nilai Gizi &amp; Foto Menu Makanan
            </h2>
            <p className="text-xs text-slate-500">
              Terhubung otomatis dengan <strong>Profil SPPG</strong> (identitas, tanggal periode &amp; TTD) dan seluruh perencanaan kategori pada <strong>Food Cost</strong>. Pilih satu atau gabungkan bebas beberapa kategori.
            </p>
          </div>

          {/* Sub-view Segmented Control */}
          <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 self-start lg:self-center shrink-0">
            <button
              type="button"
              onClick={() => setSubView("nilai_gizi")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition cursor-pointer flex items-center gap-1.5 ${
                subView === "nilai_gizi"
                  ? "bg-white text-emerald-700 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <ClipboardList className="w-3.5 h-3.5" />
              IKSS Nilai Gizi
            </button>
            <button
              type="button"
              onClick={() => setSubView("foto_menu")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition cursor-pointer flex items-center gap-1.5 ${
                subView === "foto_menu"
                  ? "bg-white text-emerald-700 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <Camera className="w-3.5 h-3.5" />
              IKSS Foto Menu Makanan
            </button>
          </div>
        </div>

        {/* Action Toolbar */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-slate-100">
          <div className="flex flex-wrap items-center gap-3">
            <button
              type="button"
              onClick={() => setShowCategoryDrawer((v) => !v)}
              className="px-3 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Layers className="w-3.5 h-3.5 text-emerald-600" />
              <span>
                {showCategoryDrawer
                  ? "Sembunyikan Pilihan Kategori"
                  : "Pilih / Gabung Kategori"}
              </span>
            </button>

            <label className="inline-flex items-center gap-2 text-xs font-semibold text-slate-700 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={!!ikssData.showKopOnIkss}
                onChange={(e) =>
                  setIkssData((prev) => ({
                    ...prev,
                    showKopOnIkss: e.target.checked,
                  }))
                }
                className="w-4 h-4 accent-emerald-600 rounded cursor-pointer"
              />
              <span>Sertakan KOP Surat SPPG</span>
            </label>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="px-3 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Eye className="w-3.5 h-3.5" />
              Pratinjau Cetak
            </button>
            <button
              type="button"
              onClick={handleExportPdf}
              className="px-3 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <FileDown className="w-3.5 h-3.5" />
              Unduh PDF
            </button>
            <button
              type="button"
              onClick={handleExportExcel}
              className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Unduh Excel
            </button>
            <button
              type="button"
              onClick={handlePrint}
              className="px-3 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Printer className="w-3.5 h-3.5" />
              Cetak Langsung
            </button>
          </div>
        </div>

        {/* FLEXIBLE CATEGORY & MERGE SELECTOR */}
        {showCategoryDrawer && (
          <div className="pt-4 border-t border-slate-100 space-y-4">
            {/* Quick Presets */}
            <div className="space-y-1.5">
              <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                Preset Cepat Kategori IKSS
              </span>
              <div className="flex flex-wrap gap-1.5">
                {QUICK_PRESETS.map((preset) => {
                  const isActive =
                    preset.sasaranIds.length === selectedSasaranIds.length &&
                    preset.sasaranIds.every((id) => selectedSasaranIds.includes(id));
                  return (
                    <button
                      key={preset.id}
                      type="button"
                      onClick={() => applyPreset(preset)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition cursor-pointer ${
                        isActive
                          ? "bg-emerald-600 border-emerald-600 text-white"
                          : "bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100"
                      }`}
                    >
                      {preset.label}
                    </button>
                  );
                })}
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {/* Source Food Cost Categories (Multi-select) */}
              <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 space-y-2.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[11px] font-bold text-slate-700 uppercase tracking-wider">
                    Sumber Perencanaan Food Cost (Bisa Gabung Bebas)
                  </span>
                  <button
                    type="button"
                    onClick={() =>
                      setIkssData((prev) => ({
                        ...prev,
                        selectedFoodCostTypes: [
                          ...STANDARD_FOOD_COST_TYPES.map((t) => t.id),
                          ...customFcTables.map((c: any) => `custom:${c.id}`),
                        ],
                      }))
                    }
                    className="text-[11px] font-bold text-emerald-700 hover:underline cursor-pointer"
                  >
                    Pilih Semua
                  </button>
                </div>

                <div className="flex flex-wrap gap-1.5">
                  {STANDARD_FOOD_COST_TYPES.map((fcType) => {
                    const checked = selectedFoodCostTypes.includes(fcType.id);
                    const hasData = populatedFcTypes.has(fcType.id);
                    return (
                      <button
                        key={fcType.id}
                        type="button"
                        onClick={() => toggleFoodCostType(fcType.id)}
                        className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 transition cursor-pointer ${
                          checked
                            ? "bg-white border-emerald-500 text-emerald-900"
                            : "bg-white/60 border-slate-200 text-slate-500 hover:bg-white"
                        }`}
                      >
                        <span
                          className={`w-3.5 h-3.5 rounded flex items-center justify-center border ${
                            checked
                              ? "bg-emerald-600 border-emerald-600 text-white"
                              : "border-slate-300 bg-white"
                          }`}
                        >
                          {checked && <Check className="w-2.5 h-2.5" />}
                        </span>
                        <span>{fcType.label}</span>
                        {hasData && (
                          <span className="text-[10px] font-bold text-emerald-600">
                            •
                          </span>
                        )}
                      </button>
                    );
                  })}

                  {customFcTables.map((tbl: any) => {
                    const customKey = `custom:${tbl.id}`;
                    const checked = selectedFoodCostTypes.includes(customKey);
                    return (
                      <button
                        key={customKey}
                        type="button"
                        onClick={() => toggleFoodCostType(customKey)}
                        className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 transition cursor-pointer ${
                          checked
                            ? "bg-white border-indigo-500 text-indigo-900"
                            : "bg-white/60 border-slate-200 text-slate-500 hover:bg-white"
                        }`}
                      >
                        <span
                          className={`w-3.5 h-3.5 rounded flex items-center justify-center border ${
                            checked
                              ? "bg-indigo-600 border-indigo-600 text-white"
                              : "border-slate-300 bg-white"
                          }`}
                        >
                          {checked && <Check className="w-2.5 h-2.5" />}
                        </span>
                        <span>{tbl.title || "Tabel Kustom"}</span>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Kelompok Sasaran Multi-select */}
              <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 space-y-2.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[11px] font-bold text-slate-700 uppercase tracking-wider">
                    Kelompok Sasaran Penerima Manfaat (Gabung Bebas)
                  </span>
                  <button
                    type="button"
                    onClick={() =>
                      setIkssData((prev) => ({
                        ...prev,
                        selectedSasaranIds: SASARAN_KATEGORI_LIST.map((s) => s.id),
                        customCategoryTitle: "SEMUA KATEGORI PENERIMA MANFAAT",
                        akgOverride: undefined,
                      }))
                    }
                    className="text-[11px] font-bold text-emerald-700 hover:underline cursor-pointer"
                  >
                    Pilih Semua
                  </button>
                </div>

                <div className="flex flex-wrap gap-1.5">
                  {SASARAN_KATEGORI_LIST.map((sas) => {
                    const checked = selectedSasaranIds.includes(sas.id);
                    return (
                      <button
                        key={sas.id}
                        type="button"
                        onClick={() => toggleSasaranId(sas.id)}
                        className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 transition cursor-pointer ${
                          checked
                            ? "bg-white border-emerald-500 text-emerald-900"
                            : "bg-white/60 border-slate-200 text-slate-500 hover:bg-white"
                        }`}
                      >
                        <span
                          className={`w-3.5 h-3.5 rounded flex items-center justify-center border ${
                            checked
                              ? "bg-emerald-600 border-emerald-600 text-white"
                              : "border-slate-300 bg-white"
                          }`}
                        >
                          {checked && <Check className="w-2.5 h-2.5" />}
                        </span>
                        <span>{sas.shortLabel}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Custom Banner Title & Periode Override */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              <div className="space-y-1">
                <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                  Judul Kategori pada Header Rekapan (Otomatis / Edit Bebas)
                </label>
                <input
                  type="text"
                  value={ikssData.customCategoryTitle ?? ""}
                  onChange={(e) =>
                    setIkssData((prev) => ({
                      ...prev,
                      customCategoryTitle: e.target.value,
                    }))
                  }
                  placeholder={autoCategoryTitle}
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                />
              </div>

              <div className="space-y-1">
                <div className="flex items-center justify-between">
                  <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                    Periode / Bulan (Otomatis dari Profil SPPG)
                  </label>
                  {ikssData.customPeriodeLabel && (
                    <button
                      type="button"
                      onClick={() =>
                        setIkssData((prev) => ({ ...prev, customPeriodeLabel: "" }))
                      }
                      className="text-[10px] font-bold text-emerald-700 hover:underline inline-flex items-center gap-1 cursor-pointer"
                    >
                      <RotateCcw className="w-3 h-3" /> Ikuti Profil
                    </button>
                  )}
                </div>
                <input
                  type="text"
                  value={ikssData.customPeriodeLabel ?? ""}
                  onChange={(e) =>
                    setIkssData((prev) => ({
                      ...prev,
                      customPeriodeLabel: e.target.value,
                    }))
                  }
                  placeholder={autoPeriodeBulan}
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                />
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Photo Error Alert */}
      {photoError && (
        <div className="p-3.5 rounded-xl bg-amber-50 border border-amber-200 text-amber-900 text-xs font-semibold flex items-center justify-between gap-3 no-print">
          <div className="flex items-center gap-2">
            <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
            <span>{photoError}</span>
          </div>
          <button
            type="button"
            onClick={() => setPhotoError(null)}
            className="p-1 hover:bg-amber-100 rounded-lg cursor-pointer"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* ================================================================= */}
      {/* SUB-VIEW A: IKSS NILAI GIZI (FORMAT GAMBAR 2)                     */}
      {/* ================================================================= */}
      {subView === "nilai_gizi" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-x-auto p-4 sm:p-6">
          <div
            id="print-area-ikss-gizi"
            className="w-full min-w-[780px] bg-white text-black"
            style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
          >
            {ikssData.showKopOnIkss && (
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
                docKey="ikss"
              />
            )}

            {/* Top Bright Green Title Banner */}
            <div className="bg-[#00B050] border border-black py-2 px-3 text-center">
              <h3 className="text-sm font-bold uppercase text-black tracking-wide">
                REKAPAN NILAI GIZI {displayCategoryTitle}
              </h3>
            </div>

            {/* Auto-Populated SPPG Identity Header */}
            {renderIdentityBlock()}

            {/* Main IKSS Nutrition Table */}
            <table
              className="w-full border-collapse border border-black text-[11px]"
              style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
            >
              <thead>
                <tr>
                  <th
                    rowSpan={2}
                    className="border border-black bg-[#8EA9DB] text-black font-bold px-2 py-2 text-center w-[155px]"
                  >
                    Hari / Tanggal
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black bg-[#8EA9DB] text-black font-bold px-2 py-2 text-center"
                  >
                    Nama Menu
                  </th>
                  <th
                    colSpan={5}
                    className="border border-black bg-[#8EA9DB] text-black font-bold px-2 py-1.5 text-center"
                  >
                    Data Nilai Gizi {displayCategoryTitle}
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black bg-[#00B050] text-black font-bold px-2 py-2 text-center w-[88px]"
                  >
                    Jumlah
                    <br />
                    Manfaat
                  </th>
                </tr>
                <tr>
                  <th className="border border-black bg-[#8EA9DB] text-black font-bold px-1.5 py-1.5 text-center w-[72px]">
                    Energi
                    <br />
                    (kkal)
                  </th>
                  <th className="border border-black bg-[#8EA9DB] text-black font-bold px-1.5 py-1.5 text-center w-[72px]">
                    Protein
                    <br />
                    (gram)
                  </th>
                  <th className="border border-black bg-[#8EA9DB] text-black font-bold px-1.5 py-1.5 text-center w-[72px]">
                    Lemak
                    <br />
                    (gram)
                  </th>
                  <th className="border border-black bg-[#8EA9DB] text-black font-bold px-1.5 py-1.5 text-center w-[84px]">
                    Karbohidrat
                    <br />
                    (gram)
                  </th>
                  <th className="border border-black bg-[#8EA9DB] text-black font-bold px-1.5 py-1.5 text-center w-[68px]">
                    Serat
                    <br />
                    (gram)
                  </th>
                </tr>
              </thead>

              <tbody>
                {dailyRows.map((row) => (
                  <tr key={row.hariKe}>
                    <td className="border border-black px-2 py-2.5 text-left align-middle font-normal whitespace-nowrap">
                      {row.formattedDate}
                    </td>
                    <td className="border border-black px-2 py-2.5 text-left align-middle leading-snug">
                      {row.namaMenu || "-"}
                    </td>
                    <td className="border border-black px-1.5 py-2.5 text-center align-middle tabular-nums">
                      {formatDec1(row.energi)}
                    </td>
                    <td className="border border-black px-1.5 py-2.5 text-center align-middle tabular-nums">
                      {formatDec1(row.protein)}
                    </td>
                    <td className="border border-black px-1.5 py-2.5 text-center align-middle tabular-nums">
                      {formatDec1(row.lemak)}
                    </td>
                    <td className="border border-black px-1.5 py-2.5 text-center align-middle tabular-nums">
                      {formatDec1(row.kh)}
                    </td>
                    <td className="border border-black px-1.5 py-2.5 text-center align-middle tabular-nums">
                      {formatDec1(row.serat)}
                    </td>
                    <td className="border border-black px-1.5 py-2.5 text-center align-middle tabular-nums">
                      {row.jumlahManfaat}
                    </td>
                  </tr>
                ))}

                {/* FOOTER ROW 1: Total Nilai Gizi */}
                <tr>
                  <td
                    colSpan={2}
                    className="border border-black bg-[#00B050] text-black font-bold px-2 py-1.5 text-left"
                  >
                    Total Nilai Gizi
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.totalEnergi)}
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.totalProtein)}
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.totalLemak)}
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.totalKh)}
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.totalSerat)}
                  </td>
                  <td className="border border-black bg-white text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {summary.totalManfaat}
                  </td>
                </tr>

                {/* FOOTER ROW 2: Jumlah Hari */}
                <tr>
                  <td
                    colSpan={2}
                    className="border border-black bg-[#00B050] text-black font-bold px-2 py-1.5 text-left"
                  >
                    Jumlah Hari
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {summary.activeNutriDays} Hari
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {summary.activeNutriDays} Hari
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {summary.activeNutriDays} Hari
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {summary.activeNutriDays} Hari
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {summary.activeNutriDays} Hari
                  </td>
                  <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {summary.activePmDays} Hari
                  </td>
                </tr>

                {/* FOOTER ROW 3: Nilai Rata-Rata */}
                <tr>
                  <td
                    colSpan={2}
                    className="border border-black bg-[#FFFF00] text-black font-bold px-2 py-1.5 text-left"
                  >
                    Nilai Rata-Rata
                  </td>
                  <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.avgEnergi)}
                  </td>
                  <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.avgProtein)}
                  </td>
                  <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.avgLemak)}
                  </td>
                  <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.avgKh)}
                  </td>
                  <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                    {formatDec1(summary.avgSerat)}
                  </td>
                  <td
                    rowSpan={3}
                    className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center align-middle tabular-nums"
                  >
                    {summary.avgManfaat} Orang
                  </td>
                </tr>

                {/* FOOTER ROW 4: Rujukan AKG (editable inline) */}
                <tr>
                  <td
                    colSpan={2}
                    className="border border-black bg-[#00B050] text-black font-bold px-2 py-1.5 text-left"
                  >
                    {effectiveAkg.label}
                  </td>
                  {(["energi", "protein", "lemak", "kh", "serat"] as const).map(
                    (field) => (
                      <td
                        key={field}
                        className="border border-black bg-[#00B050] text-black font-bold px-1 py-1 text-center tabular-nums"
                      >
                        <input
                          type="number"
                          step="0.1"
                          value={effectiveAkg[field]}
                          onChange={(e) => {
                            const val = parseFloat(e.target.value);
                            setIkssData((prev) => ({
                              ...prev,
                              akgOverride: {
                                ...(prev.akgOverride || {}),
                                [field]: Number.isFinite(val) ? val : 0,
                              },
                            }));
                          }}
                          className="w-full bg-transparent text-center font-bold text-black focus:outline-none focus:bg-white/30 rounded tabular-nums"
                        />
                      </td>
                    )
                  )}
                </tr>

                {/* FOOTER ROW 5: % Kebutuhan 90-110% */}
                <tr>
                  <td
                    colSpan={2}
                    className="border border-black bg-[#00B050] text-black font-bold px-2 py-1.5 text-left"
                  >
                    % Kebutuhan 90-110%
                  </td>
                  {renderPctCell(summary.pctEnergi)}
                  {renderPctCell(summary.pctProtein)}
                  {renderPctCell(summary.pctLemak)}
                  {renderPctCell(summary.pctKh)}
                  {renderPctCell(summary.pctSerat)}
                </tr>
              </tbody>
            </table>

            {/* Centralized Auto Signature Block from Profil SPPG */}
            <AutoSignatureBlock profile={profile} docKey="ikss" />
          </div>
        </div>
      )}

      {/* ================================================================= */}
      {/* SUB-VIEW B: IKSS FOTO MENU MAKANAN (FORMAT GAMBAR 3)              */}
      {/* ================================================================= */}
      {subView === "foto_menu" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-x-auto p-4 sm:p-6">
          <div
            id="print-area-ikss-foto"
            className="w-full min-w-[780px] bg-white text-black"
            style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
          >
            {ikssData.showKopOnIkss && (
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
                docKey="ikss"
              />
            )}

            {/* Top Bright Green Title Banner */}
            <div className="bg-[#00B050] border border-black py-2 px-3 text-center">
              <h3 className="text-sm font-bold uppercase text-black tracking-wide">
                REKAPAN MENU {displayCategoryTitle}
              </h3>
            </div>

            {/* Auto-Populated SPPG Identity Header */}
            {renderIdentityBlock()}

            {/* Main IKSS Menu Photo Table */}
            <table
              className="w-full border-collapse border border-black text-[11px]"
              style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
            >
              <thead>
                <tr>
                  <th
                    rowSpan={2}
                    className="border border-black bg-[#E54D3D] text-black font-bold px-1.5 py-2 text-center w-[46px]"
                  >
                    Hari
                    <br />
                    Ke
                  </th>
                  <th
                    colSpan={2}
                    className="border border-black bg-[#E54D3D] text-black font-bold px-2 py-1 text-center"
                  >
                    IDENTITAS SPPG
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black bg-[#8EA9DB] text-black font-bold px-2 py-2 text-center w-[155px]"
                  >
                    Hari / Tanggal
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black bg-[#8EA9DB] text-black font-bold px-2 py-2 text-center"
                  >
                    Nama Menu
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black bg-[#8EA9DB] text-black font-bold px-2 py-2 text-center w-[180px]"
                  >
                    FOTO MENU
                  </th>
                </tr>
                <tr>
                  <th className="border border-black bg-[#E54D3D] text-black font-bold px-2 py-1 text-center w-[195px]">
                    Nama SPPG
                  </th>
                  <th className="border border-black bg-[#E54D3D] text-black font-bold px-2 py-1 text-center w-[88px]">
                    ID SPPG
                  </th>
                </tr>
              </thead>

              <tbody>
                {dailyRows.map((row) => {
                  const photoSrc = getPhotoForDay(row.hariKe);
                  return (
                    <tr key={row.hariKe}>
                      <td className="border border-black px-1.5 py-2 text-center align-middle font-normal tabular-nums">
                        {row.hariKe}
                      </td>
                      <td className="border border-black px-2 py-2 text-left align-middle uppercase">
                        {profile.namaLembaga || "-"}
                      </td>
                      <td className="border border-black px-1.5 py-2 text-center align-middle uppercase tabular-nums">
                        {profile.idSPPG || "-"}
                      </td>
                      <td className="border border-black px-2 py-2 text-left align-middle whitespace-nowrap">
                        {row.formattedDate}
                      </td>
                      <td className="border border-black px-2 py-2 text-left align-middle leading-snug">
                        {row.namaMenu || "-"}
                      </td>
                      <td className="border border-black p-1.5 text-center align-middle">
                        <div className="flex flex-col items-center justify-center gap-1.5">
                          {photoSrc ? (
                            <div className="relative group w-[158px] h-[98px] overflow-hidden border border-slate-300 bg-slate-50">
                              <img
                                src={photoSrc}
                                alt={`Foto Menu Hari Ke-${row.hariKe}`}
                                className="w-full h-full object-cover"
                              />
                              <button
                                type="button"
                                onClick={() => handleDeletePhoto(row.hariKe)}
                                title="Hapus Foto"
                                className="no-print absolute top-1 right-1 p-1 rounded bg-rose-600/90 hover:bg-rose-700 text-white opacity-90 hover:opacity-100 transition cursor-pointer"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </div>
                          ) : (
                            <div className="w-[158px] h-[64px] border border-dashed border-slate-300 bg-slate-50/70 flex items-center justify-center text-[10px] text-slate-400 print:border-transparent print:bg-transparent">
                              <span className="no-print">Belum ada foto</span>
                            </div>
                          )}

                          {/* Direct Camera & Gallery Controls (Hidden on Print/PDF) */}
                          <div className="no-print flex items-center justify-center gap-1.5 w-full">
                            <label className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 text-[10px] font-bold cursor-pointer transition">
                              <Camera className="w-3 h-3" />
                              <span>Kamera</span>
                              <input
                                type="file"
                                accept="image/*"
                                capture="environment"
                                onChange={(e) => handlePhotoUpload(row.hariKe, e)}
                                className="hidden"
                              />
                            </label>
                            <label className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-800 text-[10px] font-bold cursor-pointer transition">
                              <ImageIcon className="w-3 h-3" />
                              <span>Galeri</span>
                              <input
                                type="file"
                                accept="image/*"
                                onChange={(e) => handlePhotoUpload(row.hariKe, e)}
                                className="hidden"
                              />
                            </label>
                          </div>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>

            {/* Centralized Auto Signature Block from Profil SPPG */}
            <AutoSignatureBlock profile={profile} docKey="ikss" />
          </div>
        </div>
      )}

      {/* Print Preview Modal */}
      <PrintPreviewModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title={`Pratinjau Cetak: ${activePrintTitle}`}
        elementId={activePrintId}
        filename={activeFilename}
        defaultScale={100}
        defaultPaperSize={paperSize}
        defaultOrientation="auto"
      />
    </div>
  );
}
