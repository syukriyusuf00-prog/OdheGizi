/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Camera,
  Image as ImageIcon,
  Printer,
  FileDown,
  Download,
  Eye,
  Trash2,
  RotateCcw,
  RotateCw,
  ClipboardList,
  Check,
  CheckCircle2,
  Layers,
  AlertCircle,
  X,
  Upload,
  Send,
  Cloud,
  CloudOff,
  Loader2,
  LayoutGrid,
  Table2,
  Pencil,
  Plus,
  Minus,
  ChevronLeft,
  ChevronRight,
  ZoomIn,
  Archive,
  ArrowRight,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  DEFAULT_UNIFIED_KOP_CONFIG,
  FoodCostDay,
  HariPM,
  IkssDailyPostMeta,
  IkssDataState,
  IkssDayOverride,
  IkssIdentityOverride,
  IkssWorksheetConfig,
  MasterMenu,
  SPPGProfile,
  SavedDocArchive,
  SavedDocArchiveWorksheet,
  TKPIItem,
  UnifiedKopConfig,
} from "../types";
import { calculateDay, createTkpiMap } from "../utils/calc";
import {
  SASARAN_KATEGORI_LIST,
  composeMenuName,
  ensureMasterMenuPlans,
  resolveSasaranPmFromTable2,
  stripAlergiPrefix,
} from "../utils/menuSchedule";
import { TARGET_AKG_LIMITS } from "../tkpiData";
import { printElementSmart, exportToPDF } from "../lib/printUtils";
import {
  compressMenuPhoto,
  loadIkssPhotosFromCloud,
  loadIkssPhotosLocal,
  rotatePhotoDataUrl90,
  saveIkssPhotosLocal,
  saveIkssPhotosToCloud,
} from "../utils/ikssPhotoStore";
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
  cloudUserId?: string;
  isCloudActive?: boolean;
  onSaveToArchive?: (archive: SavedDocArchive) => void;
  onGoToArsip?: () => void;
}

const DEFAULT_FC_TYPES: string[] = ["Basah", "3B-Basah"];
const DEFAULT_SASARAN_IDS: string[] = ["anak_balita", "anak_balita_13_59"];

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

function formatShortTimestamp(iso?: string): string {
  if (!iso) return "";
  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return "";
    const months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return `${d.getDate()} ${months[d.getMonth()]} ${hh}.${mm}`;
  } catch {
    return "";
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

function formatDec1(val: number): string {
  if (!Number.isFinite(val)) return "0.0";
  return (Math.round(val * 10) / 10).toFixed(1);
}

function parseNonNegativeNumber(
  raw: string,
  opts: { integer?: boolean; label: string }
): { valid: boolean; value: number; error?: string } {
  const trimmed = raw.trim().replace(",", ".");
  if (trimmed === "") return { valid: true, value: 0 };
  if (trimmed.startsWith("-")) {
    return { valid: false, value: 0, error: `${opts.label} tidak boleh negatif.` };
  }
  const num = Number(trimmed);
  if (!Number.isFinite(num)) {
    return { valid: false, value: 0, error: `${opts.label} harus berupa angka valid.` };
  }
  if (num < 0) {
    return { valid: false, value: 0, error: `${opts.label} tidak boleh negatif.` };
  }
  if (opts.integer && !Number.isInteger(num)) {
    return {
      valid: false,
      value: Math.floor(num),
      error: `${opts.label} harus berupa bilangan bulat.`,
    };
  }
  return {
    valid: true,
    value: opts.integer ? Math.round(num) : Math.round(num * 100) / 100,
  };
}

type IkssNumericField = "energi" | "protein" | "lemak" | "kh" | "serat" | "jumlahManfaat";
type IkssAkgNumericField = "energi" | "protein" | "lemak" | "kh" | "serat";

interface GiziRowDraft {
  tanggal: string;
  namaMenu: string;
  energi: string;
  protein: string;
  lemak: string;
  kh: string;
  serat: string;
  jumlahManfaat: string;
}

const NUTRIENT_FIELD_LABELS: Record<IkssNumericField, string> = {
  energi: "Energi (kkal)",
  protein: "Protein (gram)",
  lemak: "Lemak (gram)",
  kh: "Karbohidrat (gram)",
  serat: "Serat (gram)",
  jumlahManfaat: "Jumlah PM",
};

interface DailyNutritionValues {
  energi: number;
  protein: number;
  lemak: number;
  kh: number;
  serat: number;
  hasNutrition: boolean;
}

const EMPTY_NUTRITION: DailyNutritionValues = {
  energi: 0,
  protein: 0,
  lemak: 0,
  kh: 0,
  serat: 0,
  hasNutrition: false,
};

function inferFoodCostTypesForSasaranIds(sasaranIds: string[]): string[] {
  const types = new Set<string>();
  const hasMpAsi = sasaranIds.includes("balita_6_11");
  const hasTigaB = sasaranIds.some((id) =>
    ["anak_balita", "anak_balita_13_59", "ibu_hamil", "ibu_menyusui"].includes(id)
  );
  const hasSekolah = sasaranIds.some((id) =>
    [
      "tk_paud_lb",
      "sd_kelas_1_3",
      "sd_kelas_4_6",
      "smp_mts_smplb",
      "sma_smk_ma",
      "pendidik",
      "tenaga_kependidikan",
    ].includes(id)
  );
  if (hasSekolah) types.add("Basah");
  if (hasTigaB) {
    types.add("3B-Basah");
    types.add("Basah");
  }
  if (hasMpAsi) {
    types.add("MP-ASI");
    types.add("3B-Basah");
  }
  if (types.size === 0) types.add("Basah");
  return Array.from(types);
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
  cloudUserId = "local",
  isCloudActive = false,
  onSaveToArchive,
  onGoToArsip,
}: IkssMenuTabProps) {
  const [subView, setSubView] = useState<"nilai_gizi" | "foto_menu">(
    ikssData.activeSubView === "foto_menu" ? "foto_menu" : "nilai_gizi"
  );
  const [photoViewMode, setPhotoViewMode] = useState<"tabel" | "kartu">("tabel");
  const [photoFilter, setPhotoFilter] = useState<"semua" | "diposting" | "draft" | "belum">("semua");
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCategoryDrawer, setShowCategoryDrawer] = useState(false);

  // Localized async photo map (IndexedDB + chunked Cloud) so uploading never re-renders App.tsx
  const [photoMap, setPhotoMap] = useState<Record<string, string>>({});
  const [photosHydrated, setPhotosHydrated] = useState(false);
  const [processingDays, setProcessingDays] = useState<Record<number, boolean>>({});
  const [dragOverDay, setDragOverDay] = useState<number | null>(null);
  const [bulkProgress, setBulkProgress] = useState<{ current: number; total: number } | null>(null);
  const [syncState, setSyncState] = useState<"idle" | "saving" | "synced" | "local_only">("idle");
  const [photoBanner, setPhotoBanner] = useState<{ type: "success" | "error" | "info"; text: string } | null>(null);

  // Multi-worksheet active focus (for Foto Menu card editor & category drawer)
  const [activeWorksheetId, setActiveWorksheetId] = useState<string>("default");
  const [savedArchiveNotice, setSavedArchiveNotice] = useState<string | null>(null);

  // Lightbox & inline menu / full table editor state
  const [lightboxTarget, setLightboxTarget] = useState<{ wsId: string; hariKe: number } | null>(null);
  const [isTableEditMode, setIsTableEditMode] = useState(false);
  const [isEditingIdentity, setIsEditingIdentity] = useState(false);
  const [editingAkgWsId, setEditingAkgWsId] = useState<string | null>(null);
  const [confirmResetAllEdits, setConfirmResetAllEdits] = useState(false);

  // Per-row Pencil edit for IKSS Nilai Gizi (scoped `${wsId}__${hariKe}`)
  const [editingGiziKey, setEditingGiziKey] = useState<string | null>(null);
  const [giziRowDraft, setGiziRowDraft] = useState<GiziRowDraft>({
    tanggal: "",
    namaMenu: "",
    energi: "",
    protein: "",
    lemak: "",
    kh: "",
    serat: "",
    jumlahManfaat: "",
  });
  const [giziRowErrors, setGiziRowErrors] = useState<Record<string, string>>({});

  // Smooth string draft & validation error map for global table edit mode & AKG inputs
  const [bulkCellDrafts, setBulkCellDrafts] = useState<Record<string, string>>({});
  const [bulkCellErrors, setBulkCellErrors] = useState<Record<string, string>>({});

  // Per-row Pencil edit for IKSS Foto Menu (scoped `${wsId}__${hariKe}`)
  const [editingMenuKey, setEditingMenuKey] = useState<string | null>(null);
  const [editDateDraft, setEditDateDraft] = useState("");
  const [editSppgDraft, setEditSppgDraft] = useState("");
  const [editIdSppgDraft, setEditIdSppgDraft] = useState("");
  const [editMenuDraft, setEditMenuDraft] = useState("");
  const [editCaptionDraft, setEditCaptionDraft] = useState("");

  // Export guard to prevent duplicate overlapping html2canvas calls
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const [exportStatusMsg, setExportStatusMsg] = useState<string | null>(null);

  const photoMapRef = useRef<Record<string, string>>(photoMap);
  photoMapRef.current = photoMap;
  const cloudSyncTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const bulkInputRef = useRef<HTMLInputElement | null>(null);

  const showNotice = useCallback((type: "success" | "error" | "info", text: string) => {
    setPhotoBanner({ type, text });
  }, []);

  // Hydrate photos asynchronously from IndexedDB + Cloud + one-time migration of legacy ikssData.photos
  useEffect(() => {
    let cancelled = false;
    const scope = cloudUserId || "local";

    const hydrate = async () => {
      const localPhotos = await loadIkssPhotosLocal(scope);
      const legacyPhotos =
        ikssData.photos && typeof ikssData.photos === "object" ? ikssData.photos : {};
      const hasLegacy = Object.keys(legacyPhotos).length > 0;

      let merged: Record<string, string> = { ...legacyPhotos, ...localPhotos };

      if (!cancelled) {
        setPhotoMap(merged);
        setPhotosHydrated(true);
      }

      if (hasLegacy) {
        await saveIkssPhotosLocal(scope, merged);
        // Strip heavy base64 blobs from App.tsx ikssData state once safely persisted in IndexedDB
        if (!cancelled) {
          setIkssData((prev) => {
            if (!prev.photos || Object.keys(prev.photos).length === 0) return prev;
            const migratedPosts: Record<string, IkssDailyPostMeta> = { ...(prev.dailyPosts || {}) };
            Object.keys(prev.photos).forEach((k) => {
              if (!migratedPosts[k]) {
                migratedPosts[k] = { posted: true, postedAt: new Date().toISOString() };
              }
            });
            return { ...prev, photos: {}, dailyPosts: migratedPosts };
          });
        }
      }

      if (isCloudActive && cloudUserId && cloudUserId !== "local") {
        const cloudPhotos = await loadIkssPhotosFromCloud(cloudUserId);
        if (!cancelled && cloudPhotos && Object.keys(cloudPhotos).length > 0) {
          merged = { ...cloudPhotos, ...merged };
          setPhotoMap(merged);
          await saveIkssPhotosLocal(scope, merged);
        } else if (!cancelled && hasLegacy && Object.keys(merged).length > 0) {
          // Push migrated legacy photos to chunked cloud store
          saveIkssPhotosToCloud(cloudUserId, merged).catch(() => {});
        }
      }
    };

    hydrate();
    return () => {
      cancelled = true;
    };
  }, [cloudUserId, isCloudActive]);

  // Also migrate any legacy inline photos that arrive later after cloud hydration
  const legacyPhotoCount = ikssData.photos ? Object.keys(ikssData.photos).length : 0;
  useEffect(() => {
    if (legacyPhotoCount === 0) return;
    const legacyPhotos = ikssData.photos || {};
    const scope = cloudUserId || "local";
    const merged = { ...legacyPhotos, ...photoMapRef.current };
    setPhotoMap(merged);
    saveIkssPhotosLocal(scope, merged).then(() => {
      if (isCloudActive && cloudUserId && cloudUserId !== "local") {
        saveIkssPhotosToCloud(cloudUserId, merged).catch(() => {});
      }
    });
    setIkssData((prev) => {
      if (!prev.photos || Object.keys(prev.photos).length === 0) return prev;
      const migratedPosts: Record<string, IkssDailyPostMeta> = { ...(prev.dailyPosts || {}) };
      Object.keys(prev.photos).forEach((k) => {
        if (!migratedPosts[k]) {
          migratedPosts[k] = { posted: true, postedAt: new Date().toISOString() };
        }
      });
      return { ...prev, photos: {}, dailyPosts: migratedPosts };
    });
  }, [legacyPhotoCount, cloudUserId, isCloudActive]);

  useEffect(() => {
    return () => {
      if (cloudSyncTimerRef.current) {
        clearTimeout(cloudSyncTimerRef.current);
      }
    };
  }, []);

  const persistPhotos = useCallback(
    async (nextPhotos: Record<string, string>, immediateCloud = false): Promise<{ cloudOk: boolean; error?: string }> => {
      const scope = cloudUserId || "local";
      setSyncState("saving");
      await saveIkssPhotosLocal(scope, nextPhotos);

      if (!isCloudActive || !cloudUserId || cloudUserId === "local") {
        setSyncState("local_only");
        return { cloudOk: false };
      }

      if (cloudSyncTimerRef.current) {
        clearTimeout(cloudSyncTimerRef.current);
        cloudSyncTimerRef.current = null;
      }

      if (immediateCloud) {
        const res = await saveIkssPhotosToCloud(cloudUserId, nextPhotos);
        setSyncState(res.ok ? "synced" : "local_only");
        return { cloudOk: res.ok, error: res.error };
      }

      cloudSyncTimerRef.current = setTimeout(async () => {
        const res = await saveIkssPhotosToCloud(cloudUserId, photoMapRef.current);
        setSyncState(res.ok ? "synced" : "local_only");
      }, 900);
      return { cloudOk: true };
    },
    [cloudUserId, isCloudActive]
  );

  // Base fallback selections for default worksheet
  const selectedFoodCostTypes = useMemo(
    () =>
      ikssData.selectedFoodCostTypes && ikssData.selectedFoodCostTypes.length > 0
        ? ikssData.selectedFoodCostTypes
        : DEFAULT_FC_TYPES,
    [ikssData.selectedFoodCostTypes]
  );

  const selectedSasaranIds = useMemo(
    () =>
      ikssData.selectedSasaranIds && ikssData.selectedSasaranIds.length > 0
        ? ikssData.selectedSasaranIds
        : DEFAULT_SASARAN_IDS,
    [ikssData.selectedSasaranIds]
  );

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

  const autoPeriodeBulan = useMemo(
    () => derivePeriodeBulanLabel(profile.periodeDates || []),
    [profile.periodeDates]
  );
  const displayPeriodeBulan = (
    ikssData.customPeriodeLabel?.trim() || autoPeriodeBulan
  ).toUpperCase();

  // Safely cap baseAutoDays so corrupted arrays can never spin thousands of DOM rows
  const baseAutoDays = Math.min(
    62,
    Math.max(profile.periodeDates?.length || 0, harianPM?.length || 0, 10)
  );

  // Effective SPPG identity (defaults to Profil SPPG, editable directly inside IKSS)
  const effectiveIdentity = useMemo(() => {
    const ov = ikssData.identityOverride || {};
    return {
      namaLembaga: ov.namaLembaga !== undefined ? ov.namaLembaga : profile.namaLembaga || "",
      idSPPG: ov.idSPPG !== undefined ? ov.idSPPG : profile.idSPPG || "",
      alamat: ov.alamat !== undefined ? ov.alamat : profile.alamat || "",
      kabKota: ov.kabKota !== undefined ? ov.kabKota : profile.kabKota || "",
      provinsi: ov.provinsi !== undefined ? ov.provinsi : profile.provinsi || "",
      namaAhliGizi: ov.namaAhliGizi !== undefined ? ov.namaAhliGizi : profile.namaAhliGizi || "",
      noHpSPPI: ov.noHpSPPI !== undefined ? ov.noHpSPPI : profile.noHpSPPI || "",
    };
  }, [ikssData.identityOverride, profile]);

  const hasIdentityOverride = useMemo(() => {
    const ov = ikssData.identityOverride;
    return Boolean(ov && Object.keys(ov).length > 0);
  }, [ikssData.identityOverride]);

  const normalizedMenu = useMemo(() => ensureMasterMenuPlans(masterMenu), [masterMenu]);

  // Pre-index FoodCostDays & HarianPM by hariKe for O(1) lookup
  const fcByDay = useMemo(() => {
    const map = new Map<number, FoodCostDay[]>();
    for (let i = 0; i < foodCostDays.length; i++) {
      const fc = foodCostDays[i];
      if (!fc) continue;
      const list = map.get(fc.hariKe);
      if (list) list.push(fc);
      else map.set(fc.hariKe, [fc]);
    }
    return map;
  }, [foodCostDays]);

  const pmByDay = useMemo(() => {
    const map = new Map<number, HariPM>();
    for (let i = 0; i < harianPM.length; i++) {
      const h = harianPM[i];
      if (h) map.set(h.hariKe, h);
    }
    return map;
  }, [harianPM]);

  const populatedFcTypes = useMemo(() => {
    const set = new Set<string>();
    foodCostDays.forEach((d) => {
      const hasKategoriData =
        d.activeMode === "kategori" &&
        (d.categoryEntries || []).some((ce) =>
          (ce.bahanList || []).some((b) => (b.beratBB || 0) > 0 && b.nama?.trim())
        );
      const hasData =
        hasKategoriData ||
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

  const worksheetConfigs: IkssWorksheetConfig[] = useMemo(() => {
    if (ikssData.worksheetConfigs && ikssData.worksheetConfigs.length > 0) {
      return ikssData.worksheetConfigs.map((ws) => ({
        ...ws,
        sasaranIds: ws.sasaranIds && ws.sasaranIds.length > 0 ? ws.sasaranIds : DEFAULT_SASARAN_IDS,
        foodCostTypes:
          ws.foodCostTypes && ws.foodCostTypes.length > 0
            ? ws.foodCostTypes
            : inferFoodCostTypesForSasaranIds(ws.sasaranIds || DEFAULT_SASARAN_IDS),
      }));
    }
    return [
      {
        id: "default",
        title: ikssData.customCategoryTitle || "",
        sasaranIds: selectedSasaranIds,
        foodCostTypes: selectedFoodCostTypes,
        customDayCount: ikssData.customDayCount,
        akgOverride: ikssData.akgOverride,
      },
    ];
  }, [
    ikssData.worksheetConfigs,
    ikssData.customCategoryTitle,
    ikssData.customDayCount,
    ikssData.akgOverride,
    selectedSasaranIds,
    selectedFoodCostTypes,
  ]);

  useEffect(() => {
    if (!worksheetConfigs.some((w) => w.id === activeWorksheetId)) {
      setActiveWorksheetId(worksheetConfigs[0]?.id || "default");
    }
  }, [worksheetConfigs, activeWorksheetId]);

  const dailyPosts = ikssData.dailyPosts || {};
  const dayOverrides = ikssData.dayOverrides || {};

  const worksheetsComputed = useMemo(() => {
    return worksheetConfigs.map((wsConfig) => {
      const wsSasaranIds = wsConfig.sasaranIds;
      const wsFcTypes = wsConfig.foodCostTypes;
      const wsSasaranSet = new Set(wsSasaranIds);
      const wsFcTypeSet = new Set(wsFcTypes);

      const matchedPreset = QUICK_PRESETS.find(
        (p) =>
          p.sasaranIds.length === wsSasaranIds.length &&
          p.sasaranIds.every((id) => wsSasaranSet.has(id))
      );
      const autoCategoryTitle = matchedPreset
        ? matchedPreset.title
        : (() => {
            const labels = SASARAN_KATEGORI_LIST.filter((s) => wsSasaranSet.has(s.id)).map((s) =>
              s.shortLabel.toUpperCase()
            );
            if (labels.length === 0) return "SEMUA KATEGORI";
            if (labels.length <= 3) return labels.join(" / ");
            return `${labels.slice(0, 3).join(", ")} +${labels.length - 3} LAINNYA`;
          })();

      const displayCategoryTitle = (wsConfig.title?.trim() || autoCategoryTitle).toUpperCase();
      const sorted = [...wsSasaranIds].sort().join("_");
      const photoGroupKey = sorted || "default";
      const totalDays = Math.min(62, Math.max(1, wsConfig.customDayCount ?? baseAutoDays));

      const hasNonAlergiFc = wsFcTypes.some((t) => !t.toLowerCase().includes("alergi"));
      const hasAlergiFc = wsFcTypes.some((t) => t.toLowerCase().includes("alergi"));
      const hasMpAsiFc =
        wsFcTypes.some((t) => t === "MP-ASI") || wsSasaranSet.has("balita_6_11");
      const hasStandarFc = wsFcTypes.some(
        (t) =>
          t === "Basah" ||
          t === "Kering" ||
          t === "3B-Basah" ||
          t === "3B-Kering" ||
          t.startsWith("custom:")
      );

      const matchingPlans = (normalizedMenu.plans || []).filter((plan) => {
        const sasaranMatches =
          wsSasaranIds.length === 0 ||
          (plan.sasaranIds || []).some((id) => wsSasaranSet.has(id));
        if (!sasaranMatches) return false;
        if (plan.dietType === "Alergi") return hasAlergiFc;
        if (plan.dietType === "MP-ASI") return hasMpAsiFc;
        return hasStandarFc || (!hasAlergiFc && !hasMpAsiFc);
      });

      const baseDailyRows: Array<{
        hariKe: number;
        dateStr: string;
        formattedDate: string;
        autoNamaMenu: string;
        hasMasterMenu: boolean;
        pmKecilCount: number;
        pmBesarCount: number;
        jumlahManfaat: number;
        photoKey: string;
      }> = [];

      for (let i = 0; i < totalDays; i++) {
        const hariKe = i + 1;
        const dateStr = profile.periodeDates?.[i] || "";
        const formattedDate = dateStr ? formatIndoFullDate(dateStr) : `Hari Ke-${hariKe}`;

        const dayPM = pmByDay.get(hariKe);
        let pmKecilCount = 0;
        let pmBesarCount = 0;
        if (dayPM?.sasaran) {
          for (let sIdx = 0; sIdx < dayPM.sasaran.length; sIdx++) {
            const sas = dayPM.sasaran[sIdx];
            if (!wsSasaranSet.has(sas.id)) continue;
            const pK = Number(sas.porsiKecil) || 0;
            const pB = Number(sas.porsiBesar) || 0;
            const aK = Number(sas.alergiKecil) || 0;
            const aB = Number(sas.alergiBesar) || 0;
            if (hasAlergiFc && !hasNonAlergiFc) {
              pmKecilCount += aK;
              pmBesarCount += aB;
            } else if (hasNonAlergiFc && !hasAlergiFc) {
              pmKecilCount += pK;
              pmBesarCount += pB;
            } else {
              pmKecilCount += pK + aK;
              pmBesarCount += pB + aB;
            }
          }
        }

        const masterMenuNames: string[] = [];
        for (let pIdx = 0; pIdx < matchingPlans.length; pIdx++) {
          const dayItem = matchingPlans[pIdx].days?.[i];
          if (!dayItem) continue;
          const composed = composeMenuName(dayItem).trim();
          if (composed && !masterMenuNames.includes(composed)) {
            masterMenuNames.push(composed);
          }
        }

        const hasMasterMenu = masterMenuNames.length > 0;
        const autoNamaMenu = hasMasterMenu ? masterMenuNames.join(" | ") : "";

        baseDailyRows.push({
          hariKe,
          dateStr,
          formattedDate,
          autoNamaMenu,
          hasMasterMenu,
          pmKecilCount,
          pmBesarCount,
          jumlahManfaat: pmKecilCount + pmBesarCount,
          photoKey: `${photoGroupKey}__day_${hariKe}`,
        });
      }

      const nutritionByDay = new Map<number, DailyNutritionValues>();
      if (subView === "nilai_gizi") {
        const hasKecilSelected = wsSasaranIds.some((id) => PORSI_KECIL_IDS.has(id));
        const hasBesarSelected = wsSasaranIds.some((id) => PORSI_BESAR_IDS.has(id));

        for (let i = 0; i < baseDailyRows.length; i++) {
          const base = baseDailyRows[i];
          const { hariKe, pmKecilCount, pmBesarCount } = base;
          const matchedDays = (fcByDay.get(hariKe) || []).filter((fc) =>
            wsFcTypeSet.has(fc.jenisMenu)
          );

          const nutriSamples: Array<{
            energi: number;
            protein: number;
            lemak: number;
            kh: number;
            serat: number;
            weight: number;
          }> = [];

          for (let mIdx = 0; mIdx < matchedDays.length; mIdx++) {
            const fc = matchedDays[mIdx];

            if (
              fc.activeMode === "kategori" &&
              Array.isArray(fc.categoryEntries) &&
              fc.categoryEntries.length > 0
            ) {
              const isAlergi = fc.jenisMenu.toLowerCase().includes("alergi");
              const matchingEntries = fc.categoryEntries.filter((ce) =>
                (ce.sasaranIds || []).some((id) =>
                  isAlergi
                    ? wsSasaranSet.has(id) || wsSasaranSet.has(stripAlergiPrefix(id))
                    : wsSasaranSet.has(id)
                )
              );
              const targetEntries = matchingEntries.length > 0 ? matchingEntries : fc.categoryEntries;
              const dayPM = pmByDay.get(hariKe);

              for (let eIdx = 0; eIdx < targetEntries.length; eIdx++) {
                const ce = targetEntries[eIdx];
                const hasCeBahan = (ce.bahanList || []).some(
                  (b) => (Number(b.beratBB) || 0) > 0 && (b.nama?.trim() || b.tkpiId)
                );
                if (!hasCeBahan) continue;

                let cePm = 0;
                if (ce.customPmCount !== undefined) {
                  cePm = Math.max(0, Number(ce.customPmCount) || 0);
                } else if (dayPM?.sasaran) {
                  cePm = (ce.sasaranIds || []).reduce(
                    (sum, sId) =>
                      sum + resolveSasaranPmFromTable2(dayPM.sasaran, sId, ce.porsi, isAlergi, ["sampel"]),
                    0
                  );
                }

                const calcCat =
                  ce.porsi === "besar"
                    ? calculateDay(
                        ce.bahanList || [],
                        [],
                        Math.max(cePm, 1),
                        0,
                        ce.bufferPct ?? fc.bufferPct ?? 5,
                        tkpiMap,
                        fc.bumbuPct ?? bumbuPct
                      )
                    : calculateDay(
                        [],
                        ce.bahanList || [],
                        0,
                        Math.max(cePm, 1),
                        ce.bufferPct ?? fc.bufferPct ?? 5,
                        tkpiMap,
                        fc.bumbuPct ?? bumbuPct
                      );
                const nutCat = ce.porsi === "besar" ? calcCat.nutrisiPorsiBesar : calcCat.nutrisiPorsiKecil;
                nutriSamples.push({
                  energi: Number(nutCat.energi) || 0,
                  protein: Number(nutCat.protein) || 0,
                  lemak: Number(nutCat.lemak) || 0,
                  kh: Number(nutCat.kh) || 0,
                  serat: Number(nutCat.serat) || 0,
                  weight: Math.max(cePm, 1),
                });
              }
              continue;
            }

            const hasBesarBahan = (fc.porsiBesarBahan || []).some(
              (b) => (Number(b.beratBB) || 0) > 0 && (b.nama?.trim() || b.tkpiId)
            );
            const hasKecilBahan = (fc.porsiKecilBahan || []).some(
              (b) => (Number(b.beratBB) || 0) > 0 && (b.nama?.trim() || b.tkpiId)
            );
            if (!hasBesarBahan && !hasKecilBahan) continue;

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
                  energi:
                    ((Number(nKecil.energi) || 0) * wK + (Number(nBesar.energi) || 0) * wB) /
                    totalW,
                  protein:
                    ((Number(nKecil.protein) || 0) * wK + (Number(nBesar.protein) || 0) * wB) /
                    totalW,
                  lemak:
                    ((Number(nKecil.lemak) || 0) * wK + (Number(nBesar.lemak) || 0) * wB) /
                    totalW,
                  kh: ((Number(nKecil.kh) || 0) * wK + (Number(nBesar.kh) || 0) * wB) / totalW,
                  serat:
                    ((Number(nKecil.serat) || 0) * wK + (Number(nBesar.serat) || 0) * wB) /
                    totalW,
                  weight: totalW,
                });
              } else if (hasKecilBahan) {
                nutriSamples.push(toSample(nKecil, Math.max(pmKecilCount, 1)));
              } else if (hasBesarBahan) {
                nutriSamples.push(toSample(nBesar, Math.max(pmBesarCount, 1)));
              }
            } else if (hasKecilSelected) {
              if (hasKecilBahan) {
                nutriSamples.push(toSample(nKecil, Math.max(pmKecilCount, 1)));
              }
            } else if (hasBesarSelected) {
              if (hasBesarBahan) {
                nutriSamples.push(toSample(nBesar, Math.max(pmBesarCount, 1)));
              }
            }
          }

          for (let cIdx = 0; cIdx < customFcTables.length; cIdx++) {
            const tbl = customFcTables[cIdx];
            if (!wsFcTypeSet.has(`custom:${tbl.id}`)) continue;
            if (!Array.isArray(tbl.bahanList) || tbl.bahanList.length === 0) continue;
            const hasCustomBahan = tbl.bahanList.some(
              (b: any) => (Number(b.beratBB) || 0) > 0 && (b.nama?.trim() || b.tkpiId)
            );
            if (!hasCustomBahan) continue;
            if (tbl.hariKe !== undefined && Number(tbl.hariKe) !== hariKe) continue;
            const calc =
              tbl.porsi === "besar"
                ? calculateDay(
                    tbl.bahanList,
                    [],
                    Math.max(tbl.pmCount || 1, 1),
                    0,
                    tbl.bufferPct || 5,
                    tkpiMap,
                    bumbuPct
                  )
                : calculateDay(
                    [],
                    tbl.bahanList,
                    0,
                    Math.max(tbl.pmCount || 1, 1),
                    tbl.bufferPct || 5,
                    tkpiMap,
                    bumbuPct
                  );
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
          }

          if (nutriSamples.length > 0) {
            const sumW = nutriSamples.reduce((a, b) => a + b.weight, 0) || nutriSamples.length;
            nutritionByDay.set(hariKe, {
              energi: nutriSamples.reduce((a, b) => a + b.energi * b.weight, 0) / sumW,
              protein: nutriSamples.reduce((a, b) => a + b.protein * b.weight, 0) / sumW,
              lemak: nutriSamples.reduce((a, b) => a + b.lemak * b.weight, 0) / sumW,
              kh: nutriSamples.reduce((a, b) => a + b.kh * b.weight, 0) / sumW,
              serat: nutriSamples.reduce((a, b) => a + b.serat * b.weight, 0) / sumW,
              hasNutrition: true,
            });
          } else {
            nutritionByDay.set(hariKe, EMPTY_NUTRITION);
          }
        }
      }

      const dailyRows = baseDailyRows.map((row) => {
        const nut = nutritionByDay.get(row.hariKe) || EMPTY_NUTRITION;
        const meta = dailyPosts[row.photoKey] || dailyPosts[`day_${row.hariKe}`];
        const ov = dayOverrides[row.photoKey];

        const autoFormattedDate = row.formattedDate;
        const formattedDate = ov?.tanggal !== undefined ? ov.tanggal : autoFormattedDate;
        const namaMenu =
          ov?.namaMenu !== undefined
            ? ov.namaMenu
            : (meta?.customMenuName || "").trim() || row.autoNamaMenu;
        const namaSppg = ov?.namaSppg !== undefined ? ov.namaSppg : effectiveIdentity.namaLembaga;
        const idSppg = ov?.idSppg !== undefined ? ov.idSppg : effectiveIdentity.idSPPG;

        const energi = ov?.energi !== undefined ? ov.energi : nut.energi;
        const protein = ov?.protein !== undefined ? ov.protein : nut.protein;
        const lemak = ov?.lemak !== undefined ? ov.lemak : nut.lemak;
        const kh = ov?.kh !== undefined ? ov.kh : nut.kh;
        const serat = ov?.serat !== undefined ? ov.serat : nut.serat;
        const jumlahManfaat = ov?.jumlahManfaat !== undefined ? ov.jumlahManfaat : row.jumlahManfaat;

        const hasManualNutri =
          ov?.energi !== undefined ||
          ov?.protein !== undefined ||
          ov?.lemak !== undefined ||
          ov?.kh !== undefined ||
          ov?.serat !== undefined;

        const hasNutrition =
          nut.hasNutrition ||
          (hasManualNutri && (energi > 0 || protein > 0 || lemak > 0 || kh > 0 || serat > 0));

        const editedFields = {
          tanggal: ov?.tanggal !== undefined,
          namaMenu: ov?.namaMenu !== undefined || Boolean((meta?.customMenuName || "").trim()),
          namaSppg: ov?.namaSppg !== undefined,
          idSppg: ov?.idSppg !== undefined,
          energi: ov?.energi !== undefined,
          protein: ov?.protein !== undefined,
          lemak: ov?.lemak !== undefined,
          kh: ov?.kh !== undefined,
          serat: ov?.serat !== undefined,
          jumlahManfaat: ov?.jumlahManfaat !== undefined,
        };

        const isEdited = Object.values(editedFields).some(Boolean);

        return {
          ...row,
          autoFormattedDate,
          formattedDate,
          namaMenu,
          namaSppg,
          idSppg,
          caption: meta?.caption || "",
          autoEnergi: nut.energi,
          autoProtein: nut.protein,
          autoLemak: nut.lemak,
          autoKh: nut.kh,
          autoSerat: nut.serat,
          autoJumlahManfaat: row.jumlahManfaat,
          energi,
          protein,
          lemak,
          kh,
          serat,
          jumlahManfaat,
          hasNutrition,
          isEdited,
          editedFields,
        };
      });

      const autoAkg: AkgPresetTarget = (() => {
        const matched = wsSasaranIds.map((id) => SASARAN_AKG_REF[id]).filter(Boolean);
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
        const allKecil = wsSasaranIds.every((id) => PORSI_KECIL_IDS.has(id));
        const allBesar = wsSasaranIds.every((id) => PORSI_BESAR_IDS.has(id));
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
      })();

      const wsAkgOv = wsConfig.akgOverride;
      const effectiveAkg = {
        label: wsAkgOv?.label || autoAkg.label,
        energi: wsAkgOv?.energi !== undefined ? wsAkgOv.energi : autoAkg.energi,
        protein: wsAkgOv?.protein !== undefined ? wsAkgOv.protein : autoAkg.protein,
        lemak: wsAkgOv?.lemak !== undefined ? wsAkgOv.lemak : autoAkg.lemak,
        kh: wsAkgOv?.kh !== undefined ? wsAkgOv.kh : autoAkg.kh,
        serat: wsAkgOv?.serat !== undefined ? wsAkgOv.serat : autoAkg.serat,
      };

      const totalEnergi = dailyRows.reduce((s, r) => s + (r.hasNutrition ? r.energi : 0), 0);
      const totalProtein = dailyRows.reduce((s, r) => s + (r.hasNutrition ? r.protein : 0), 0);
      const totalLemak = dailyRows.reduce((s, r) => s + (r.hasNutrition ? r.lemak : 0), 0);
      const totalKh = dailyRows.reduce((s, r) => s + (r.hasNutrition ? r.kh : 0), 0);
      const totalSerat = dailyRows.reduce((s, r) => s + (r.hasNutrition ? r.serat : 0), 0);
      const totalManfaat = dailyRows.reduce((s, r) => s + r.jumlahManfaat, 0);

      const activeNutriDays = dailyRows.filter(
        (r) =>
          r.hasNutrition &&
          (r.energi > 0 || r.protein > 0 || r.lemak > 0 || r.kh > 0 || r.serat > 0)
      ).length;
      const activePmDays = dailyRows.filter((r) => r.jumlahManfaat > 0).length;

      const avgEnergi = activeNutriDays > 0 ? totalEnergi / activeNutriDays : 0;
      const avgProtein = activeNutriDays > 0 ? totalProtein / activeNutriDays : 0;
      const avgLemak = activeNutriDays > 0 ? totalLemak / activeNutriDays : 0;
      const avgKh = activeNutriDays > 0 ? totalKh / activeNutriDays : 0;
      const avgSerat = activeNutriDays > 0 ? totalSerat / activeNutriDays : 0;
      const avgManfaat = activePmDays > 0 ? Math.round(totalManfaat / activePmDays) : 0;
      const pct = (avg: number, ref: number) => (ref > 0 ? Math.round((avg / ref) * 100) : 0);

      const summary = {
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

      let uploaded = 0;
      let posted = 0;
      let draft = 0;
      for (let i = 0; i < baseDailyRows.length; i++) {
        const r = baseDailyRows[i];
        const hasPhoto = Boolean(photoMap[r.photoKey] || photoMap[`day_${r.hariKe}`]);
        if (hasPhoto) {
          uploaded += 1;
          const meta = dailyPosts[r.photoKey] || dailyPosts[`day_${r.hariKe}`];
          if (meta?.posted) posted += 1;
          else draft += 1;
        }
      }
      const empty = Math.max(0, baseDailyRows.length - uploaded);
      const photoPct =
        baseDailyRows.length > 0 ? Math.round((posted / baseDailyRows.length) * 100) : 0;

      const hasAkg = Boolean(wsAkgOv && Object.keys(wsAkgOv).length > 0);
      const hasCustomCount =
        wsConfig.customDayCount !== undefined && wsConfig.customDayCount !== baseAutoDays;
      const editedDays = dailyRows.filter((r) => r.isEdited).length;

      return {
        config: wsConfig,
        wsSasaranIds,
        wsFcTypes,
        wsSasaranSet,
        wsFcTypeSet,
        autoCategoryTitle,
        displayCategoryTitle,
        photoGroupKey,
        totalDays,
        baseDailyRows,
        dailyRows,
        autoAkg,
        effectiveAkg,
        summary,
        photoStats: {
          total: baseDailyRows.length,
          uploaded,
          posted,
          draft,
          empty,
          pct: photoPct,
        },
        manualEditStats: {
          editedDays,
          hasIdentity: hasIdentityOverride,
          hasAkg,
          hasCustomCount,
          totalCount:
            editedDays + (hasIdentityOverride ? 1 : 0) + (hasAkg ? 1 : 0) + (hasCustomCount ? 1 : 0),
        },
      };
    });
  }, [
    worksheetConfigs,
    baseAutoDays,
    profile.periodeDates,
    pmByDay,
    normalizedMenu,
    subView,
    fcByDay,
    tkpiMap,
    bumbuPct,
    customFcTables,
    dailyPosts,
    dayOverrides,
    effectiveIdentity,
    photoMap,
    hasIdentityOverride,
  ]);

  const activeWs =
    worksheetsComputed.find((w) => w.config.id === activeWorksheetId) || worksheetsComputed[0];

  const selectedFoodCostTypesForFocus = activeWs.wsFcTypes;
  const selectedSasaranIdsForFocus = activeWs.wsSasaranIds;
  const selectedFcTypeSet = activeWs.wsFcTypeSet;
  const selectedSasaranSet = activeWs.wsSasaranSet;
  const displayCategoryTitle = activeWs.displayCategoryTitle;
  const photoGroupKey = activeWs.photoGroupKey;
  const totalDays = activeWs.totalDays;
  const baseDailyRows = activeWs.baseDailyRows;
  const dailyRows = activeWs.dailyRows;
  const photoStats = activeWs.photoStats;
  const manualEditStats = activeWs.manualEditStats;

  const totalPhotoAggregate = useMemo(() => {
    return worksheetsComputed.reduce(
      (acc, ws) => ({
        posted: acc.posted + ws.photoStats.posted,
        total: acc.total + ws.photoStats.total,
      }),
      { posted: 0, total: 0 }
    );
  }, [worksheetsComputed]);

  const getPhotoForWsDay = useCallback(
    (wsPhotoGroupKey: string, hariKe: number): string => {
      return (
        photoMap[`${wsPhotoGroupKey}__day_${hariKe}`] ||
        photoMap[`day_${hariKe}`] ||
        ""
      );
    },
    [photoMap]
  );

  const getPostMetaForWsDay = useCallback(
    (wsPhotoGroupKey: string, hariKe: number): IkssDailyPostMeta | undefined => {
      return dailyPosts[`${wsPhotoGroupKey}__day_${hariKe}`] || dailyPosts[`day_${hariKe}`];
    },
    [dailyPosts]
  );

  const getPhotoForDay = useCallback(
    (hariKe: number): string => getPhotoForWsDay(photoGroupKey, hariKe),
    [getPhotoForWsDay, photoGroupKey]
  );

  const getPostMetaForDay = useCallback(
    (hariKe: number): IkssDailyPostMeta | undefined => getPostMetaForWsDay(photoGroupKey, hariKe),
    [getPostMetaForWsDay, photoGroupKey]
  );

  const filteredPhotoRows = useMemo(() => {
    if (photoFilter === "semua") return dailyRows;
    return dailyRows.filter((row) => {
      const hasPhoto = Boolean(getPhotoForDay(row.hariKe));
      const meta = getPostMetaForDay(row.hariKe);
      if (photoFilter === "diposting") return hasPhoto && Boolean(meta?.posted);
      if (photoFilter === "draft") return hasPhoto && !meta?.posted;
      return !hasPhoto;
    });
  }, [dailyRows, photoFilter, getPhotoForDay, getPostMetaForDay]);

  // Sub-view switch handler (remember user choice)
  const handleSwitchSubView = (nextView: "nilai_gizi" | "foto_menu") => {
    setSubView(nextView);
    setIkssData((prev) =>
      prev.activeSubView === nextView ? prev : { ...prev, activeSubView: nextView }
    );
  };

  // Multi-worksheet update helper (mirrors [0] to legacy top-level fields)
  const updateWorksheetConfig = useCallback(
    (wsId: string, updater: (ws: IkssWorksheetConfig) => IkssWorksheetConfig) => {
      setIkssData((prev) => {
        const baseList: IkssWorksheetConfig[] =
          prev.worksheetConfigs && prev.worksheetConfigs.length > 0
            ? prev.worksheetConfigs
            : [
                {
                  id: "default",
                  title: prev.customCategoryTitle || "",
                  sasaranIds:
                    prev.selectedSasaranIds && prev.selectedSasaranIds.length > 0
                      ? prev.selectedSasaranIds
                      : DEFAULT_SASARAN_IDS,
                  foodCostTypes:
                    prev.selectedFoodCostTypes && prev.selectedFoodCostTypes.length > 0
                      ? prev.selectedFoodCostTypes
                      : DEFAULT_FC_TYPES,
                  customDayCount: prev.customDayCount,
                  akgOverride: prev.akgOverride,
                },
              ];

        const nextList = baseList.map((item) => (item.id === wsId ? updater(item) : item));
        const primary = nextList[0];
        const nextState: IkssDataState = {
          ...prev,
          selectedSasaranIds: primary.sasaranIds,
          selectedFoodCostTypes: primary.foodCostTypes,
          customCategoryTitle: primary.title,
        };
        if (primary.customDayCount !== undefined) nextState.customDayCount = primary.customDayCount;
        else delete nextState.customDayCount;
        if (primary.akgOverride !== undefined) nextState.akgOverride = primary.akgOverride;
        else delete nextState.akgOverride;

        if (prev.worksheetConfigs && prev.worksheetConfigs.length > 0) {
          nextState.worksheetConfigs = nextList;
        }
        return nextState;
      });
    },
    [setIkssData]
  );

  const setAllWorksheetConfigs = useCallback(
    (nextList: IkssWorksheetConfig[], nextFocusId?: string) => {
      const sanitized = nextList.length > 0 ? nextList : worksheetConfigs.slice(0, 1);
      const primary = sanitized[0];
      setIkssData((prev) => {
        const nextState: IkssDataState = {
          ...prev,
          worksheetConfigs: sanitized,
          selectedSasaranIds: primary.sasaranIds,
          selectedFoodCostTypes: primary.foodCostTypes,
          customCategoryTitle: primary.title,
        };
        if (primary.customDayCount !== undefined) nextState.customDayCount = primary.customDayCount;
        else delete nextState.customDayCount;
        if (primary.akgOverride !== undefined) nextState.akgOverride = primary.akgOverride;
        else delete nextState.akgOverride;
        return nextState;
      });
      if (nextFocusId) setActiveWorksheetId(nextFocusId);
      else if (sanitized[0]) setActiveWorksheetId(sanitized[0].id);
    },
    [setIkssData, worksheetConfigs]
  );

  const toggleIndividualSasaranWorksheet = (sasId: string, label: string) => {
    setEditingGiziKey(null);
    setEditingMenuKey(null);
    setBulkCellDrafts({});
    setBulkCellErrors({});
    const existingIdx = worksheetConfigs.findIndex(
      (w) => w.sasaranIds.length === 1 && w.sasaranIds[0] === sasId
    );
    if (existingIdx >= 0) {
      if (worksheetConfigs.length <= 1) {
        showNotice("info", "Minimal 1 lembar kerja IKSS harus tetap ditampilkan.");
        return;
      }
      const next = worksheetConfigs.filter((_, idx) => idx !== existingIdx);
      setAllWorksheetConfigs(next);
      return;
    }
    const newWs: IkssWorksheetConfig = {
      id: `ws_${sasId}_${Date.now().toString(36)}`,
      title: label.toUpperCase(),
      sasaranIds: [sasId],
      foodCostTypes: inferFoodCostTypesForSasaranIds([sasId]),
    };
    setAllWorksheetConfigs([...worksheetConfigs, newWs], newWs.id);
  };

  const togglePresetWorksheet = (preset: (typeof QUICK_PRESETS)[number]) => {
    setEditingGiziKey(null);
    setEditingMenuKey(null);
    setBulkCellDrafts({});
    setBulkCellErrors({});
    const presetSet = new Set(preset.sasaranIds);
    const existingIdx = worksheetConfigs.findIndex(
      (w) =>
        w.sasaranIds.length === preset.sasaranIds.length &&
        w.sasaranIds.every((id) => presetSet.has(id))
    );
    if (existingIdx >= 0) {
      if (worksheetConfigs.length <= 1) {
        showNotice("info", "Minimal 1 lembar kerja IKSS harus tetap ditampilkan.");
        return;
      }
      const next = worksheetConfigs.filter((_, idx) => idx !== existingIdx);
      setAllWorksheetConfigs(next);
      return;
    }
    const newWs: IkssWorksheetConfig = {
      id: `ws_preset_${preset.id}_${Date.now().toString(36)}`,
      title: preset.title,
      sasaranIds: [...preset.sasaranIds],
      foodCostTypes: [...preset.foodCostTypes],
    };
    setAllWorksheetConfigs([...worksheetConfigs, newWs], newWs.id);
  };

  const handleSplitAllSasaranWorksheets = () => {
    setEditingGiziKey(null);
    setEditingMenuKey(null);
    setBulkCellDrafts({});
    setBulkCellErrors({});
    const allSplit: IkssWorksheetConfig[] = SASARAN_KATEGORI_LIST.map((s) => ({
      id: `ws_${s.id}`,
      title: s.label.toUpperCase(),
      sasaranIds: [s.id],
      foodCostTypes: inferFoodCostTypesForSasaranIds([s.id]),
    }));
    setAllWorksheetConfigs(allSplit, allSplit[0].id);
    showNotice("success", `Menampilkan ${allSplit.length} lembar kerja terpisah untuk seluruh kelompok sasaran.`);
  };

  const handleAddCustomWorksheet = () => {
    const newId = `ws_custom_${Date.now().toString(36)}`;
    const base = activeWs.config;
    const newWs: IkssWorksheetConfig = {
      id: newId,
      title: `${activeWs.displayCategoryTitle} (LEMBAR ${worksheetConfigs.length + 1})`,
      sasaranIds: [...base.sasaranIds],
      foodCostTypes: [...base.foodCostTypes],
    };
    setAllWorksheetConfigs([...worksheetConfigs, newWs], newId);
  };

  const handleRemoveWorksheet = (wsId: string) => {
    if (worksheetConfigs.length <= 1) return;
    const next = worksheetConfigs.filter((w) => w.id !== wsId);
    setAllWorksheetConfigs(next);
  };

  // Category & Preset Handlers (for currently focused worksheet)
  const toggleFoodCostType = (typeId: string) => {
    updateWorksheetConfig(activeWs.config.id, (ws) => {
      const current = ws.foodCostTypes || [];
      const exists = current.includes(typeId);
      const next = exists ? current.filter((x) => x !== typeId) : [...current, typeId];
      return {
        ...ws,
        foodCostTypes: next.length > 0 ? next : [typeId],
      };
    });
  };

  const toggleSasaranId = (sasId: string) => {
    setEditingGiziKey(null);
    setEditingMenuKey(null);
    setBulkCellDrafts({});
    setBulkCellErrors({});
    updateWorksheetConfig(activeWs.config.id, (ws) => {
      const current = ws.sasaranIds || [];
      const exists = current.includes(sasId);
      const next = exists ? current.filter((x) => x !== sasId) : [...current, sasId];
      const copy: IkssWorksheetConfig = {
        ...ws,
        sasaranIds: next.length > 0 ? next : [sasId],
        title: "",
      };
      delete copy.akgOverride;
      return copy;
    });
  };

  const applyPreset = (preset: (typeof QUICK_PRESETS)[number]) => {
    setEditingGiziKey(null);
    setEditingMenuKey(null);
    setBulkCellDrafts({});
    setBulkCellErrors({});
    updateWorksheetConfig(activeWs.config.id, (ws) => {
      const copy: IkssWorksheetConfig = {
        ...ws,
        sasaranIds: [...preset.sasaranIds],
        foodCostTypes: [...preset.foodCostTypes],
        title: preset.title,
      };
      delete copy.akgOverride;
      return copy;
    });
  };

  // =========================================================================
  // PHOTO UPLOAD, ROTATE, POSTING & BULK HANDLERS (100% NON-BLOCKING)
  // =========================================================================

  const processSinglePhotoFile = useCallback(
    async (wsPhotoGroupKey: string, hariKe: number, file: File) => {
      setPhotoBanner(null);
      setProcessingDays((prev) => ({ ...prev, [hariKe]: true }));
      try {
        const compressed = await compressMenuPhoto(file, {
          maxWidth: 640,
          maxHeight: 480,
          targetMaxChars: 72_000,
        });
        const key = `${wsPhotoGroupKey}__day_${hariKe}`;
        const nextPhotos = { ...photoMapRef.current, [key]: compressed };
        setPhotoMap(nextPhotos);

        const autoPost = Boolean(ikssData.autoPostOnUpload);
        const nowIso = new Date().toISOString();
        setIkssData((prev) => ({
          ...prev,
          dailyPosts: {
            ...(prev.dailyPosts || {}),
            [key]: {
              ...(prev.dailyPosts?.[key] || {}),
              posted: autoPost,
              postedAt: autoPost ? nowIso : prev.dailyPosts?.[key]?.postedAt,
              updatedAt: nowIso,
            },
          },
        }));

        const res = await persistPhotos(nextPhotos, autoPost);
        if (autoPost) {
          showNotice(
            "success",
            res.cloudOk
              ? `Foto menu Hari Ke-${hariKe} berhasil diunggah, diposting, dan tersinkron ke cloud.`
              : `Foto menu Hari Ke-${hariKe} berhasil diunggah & diposting (tersimpan di perangkat ini).`
          );
        } else {
          showNotice(
            "info",
            `Foto menu Hari Ke-${hariKe} berhasil diunggah sebagai Draft. Klik "Posting Foto" untuk mempublikasikan ke rekapan.`
          );
        }
      } catch (err: any) {
        showNotice("error", err?.message || `Gagal memproses foto Hari Ke-${hariKe}.`);
      } finally {
        setProcessingDays((prev) => {
          const copy = { ...prev };
          delete copy[hariKe];
          return copy;
        });
      }
    },
    [ikssData.autoPostOnUpload, persistPhotos, setIkssData, showNotice]
  );

  const handlePhotoInputChange = (
    wsPhotoGroupKey: string,
    hariKe: number,
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    void processSinglePhotoFile(wsPhotoGroupKey, hariKe, file);
  };

  const handleDropOnDay = (
    wsPhotoGroupKey: string,
    hariKe: number,
    e: React.DragEvent
  ) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverDay(null);
    const file = e.dataTransfer.files?.[0];
    if (file && file.type.startsWith("image/")) {
      void processSinglePhotoFile(wsPhotoGroupKey, hariKe, file);
    } else if (file) {
      showNotice("error", "Harap lepaskan file gambar (JPG, PNG, WEBP).");
    }
  };

  const handleBulkUploadChange = async (
    ws: (typeof worksheetsComputed)[number],
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const fileList = e.target.files;
    const files: File[] = fileList
      ? Array.from<File>(fileList).filter((f) => f.type.startsWith("image/"))
      : [];
    e.target.value = "";
    if (files.length === 0) return;

    const unfilledDays = ws.baseDailyRows
      .filter((r) => !photoMapRef.current[r.photoKey] && !photoMapRef.current[`day_${r.hariKe}`])
      .map((r) => r.hariKe);
    const targetDays =
      unfilledDays.length > 0
        ? unfilledDays
        : ws.baseDailyRows.map((r) => r.hariKe);

    const pairs = files.slice(0, targetDays.length).map((file, idx) => ({
      file,
      hariKe: targetDays[idx],
    }));

    setPhotoBanner(null);
    setBulkProgress({ current: 0, total: pairs.length });

    const updatedPhotos: Record<string, string> = { ...photoMapRef.current };
    const postedMetaPatch: Record<string, IkssDailyPostMeta> = {};
    const autoPost = Boolean(ikssData.autoPostOnUpload);
    const nowIso = new Date().toISOString();
    let successCount = 0;

    for (let i = 0; i < pairs.length; i++) {
      const { file, hariKe } = pairs[i];
      setBulkProgress({ current: i + 1, total: pairs.length });
      setProcessingDays((prev) => ({ ...prev, [hariKe]: true }));
      try {
        const compressed = await compressMenuPhoto(file, {
          maxWidth: 640,
          maxHeight: 480,
          targetMaxChars: 72_000,
        });
        const key = `${ws.photoGroupKey}__day_${hariKe}`;
        updatedPhotos[key] = compressed;
        postedMetaPatch[key] = {
          ...(ikssData.dailyPosts?.[key] || {}),
          posted: autoPost,
          postedAt: autoPost ? nowIso : ikssData.dailyPosts?.[key]?.postedAt,
          updatedAt: nowIso,
        };
        successCount += 1;
        setPhotoMap({ ...updatedPhotos });
      } catch {
        // Continue with remaining files
      } finally {
        setProcessingDays((prev) => {
          const copy = { ...prev };
          delete copy[hariKe];
          return copy;
        });
      }
      await new Promise<void>((r) => requestAnimationFrame(() => r()));
    }

    setBulkProgress(null);

    if (successCount > 0) {
      setIkssData((prev) => ({
        ...prev,
        dailyPosts: {
          ...(prev.dailyPosts || {}),
          ...postedMetaPatch,
        },
      }));
      const res = await persistPhotos(updatedPhotos, autoPost);
      showNotice(
        "success",
        autoPost
          ? `${successCount} foto menu harian (${ws.displayCategoryTitle}) berhasil diunggah & diposting${res.cloudOk ? " ke cloud" : " (tersimpan lokal)"}.`
          : `${successCount} foto berhasil diunggah ke tabel harian. Klik "Posting Semua Foto" untuk mempublikasikan.`
      );
    } else {
      showNotice("error", "Tidak ada foto valid yang berhasil diproses.");
    }
  };

  const handlePostDayPhoto = async (wsPhotoGroupKey: string, hariKe: number) => {
    const photoSrc = getPhotoForWsDay(wsPhotoGroupKey, hariKe);
    if (!photoSrc) {
      showNotice("error", `Unggah foto menu Hari Ke-${hariKe} terlebih dahulu sebelum memposting.`);
      return;
    }
    const key = `${wsPhotoGroupKey}__day_${hariKe}`;
    const nowIso = new Date().toISOString();

    const nextPhotos = { ...photoMapRef.current, [key]: photoSrc };
    setPhotoMap(nextPhotos);

    setIkssData((prev) => ({
      ...prev,
      dailyPosts: {
        ...(prev.dailyPosts || {}),
        [key]: {
          ...(prev.dailyPosts?.[key] || {}),
          posted: true,
          postedAt: nowIso,
          updatedAt: nowIso,
        },
      },
    }));

    const res = await persistPhotos(nextPhotos, true);
    showNotice(
      "success",
      res.cloudOk
        ? `Foto menu Hari Ke-${hariKe} resmi diposting & tersinkron ke cloud.`
        : `Foto menu Hari Ke-${hariKe} resmi diposting & tersimpan aman di perangkat.`
    );
  };

  const handleUnpostDayPhoto = (wsPhotoGroupKey: string, hariKe: number) => {
    const key = `${wsPhotoGroupKey}__day_${hariKe}`;
    setIkssData((prev) => ({
      ...prev,
      dailyPosts: {
        ...(prev.dailyPosts || {}),
        [key]: {
          ...(prev.dailyPosts?.[key] || {}),
          posted: false,
          updatedAt: new Date().toISOString(),
        },
      },
    }));
  };

  const handlePostAllPhotos = async (ws: (typeof worksheetsComputed)[number]) => {
    const uploadedKeys: string[] = [];
    for (let i = 0; i < ws.baseDailyRows.length; i++) {
      const r = ws.baseDailyRows[i];
      const src = photoMapRef.current[r.photoKey] || photoMapRef.current[`day_${r.hariKe}`];
      if (src) {
        uploadedKeys.push(r.photoKey);
      }
    }

    if (uploadedKeys.length === 0) {
      showNotice("error", "Belum ada foto menu harian yang diunggah untuk diposting.");
      return;
    }

    const nowIso = new Date().toISOString();
    const nextPhotos = { ...photoMapRef.current };
    ws.baseDailyRows.forEach((r) => {
      const fallbackSrc = photoMapRef.current[`day_${r.hariKe}`];
      if (!nextPhotos[r.photoKey] && fallbackSrc) {
        nextPhotos[r.photoKey] = fallbackSrc;
      }
    });
    setPhotoMap(nextPhotos);

    setIkssData((prev) => {
      const nextPosts = { ...(prev.dailyPosts || {}) };
      uploadedKeys.forEach((k) => {
        nextPosts[k] = {
          ...(nextPosts[k] || {}),
          posted: true,
          postedAt: nowIso,
          updatedAt: nowIso,
        };
      });
      return { ...prev, dailyPosts: nextPosts };
    });

    const res = await persistPhotos(nextPhotos, true);
    showNotice(
      "success",
      res.cloudOk
        ? `${uploadedKeys.length} foto menu harian berhasil diposting & disinkronkan ke cloud (${formatShortTimestamp(nowIso)}).`
        : `${uploadedKeys.length} foto menu harian berhasil diposting & disimpan lokal${res.error ? ` — ${res.error}` : "."}`
    );
  };

  const handleRotateDayPhoto = async (wsPhotoGroupKey: string, hariKe: number) => {
    const photoSrc = getPhotoForWsDay(wsPhotoGroupKey, hariKe);
    if (!photoSrc) return;
    setProcessingDays((prev) => ({ ...prev, [hariKe]: true }));
    try {
      const rotated = await rotatePhotoDataUrl90(photoSrc);
      const key = `${wsPhotoGroupKey}__day_${hariKe}`;
      const nextPhotos = { ...photoMapRef.current, [key]: rotated };
      setPhotoMap(nextPhotos);
      await persistPhotos(nextPhotos, false);
    } catch (err: any) {
      showNotice("error", err?.message || "Gagal memutar foto.");
    } finally {
      setProcessingDays((prev) => {
        const copy = { ...prev };
        delete copy[hariKe];
        return copy;
      });
    }
  };

  const handleDeletePhoto = async (wsPhotoGroupKey: string, hariKe: number) => {
    const key = `${wsPhotoGroupKey}__day_${hariKe}`;
    const fallbackKey = `day_${hariKe}`;
    const nextPhotos = { ...photoMapRef.current };
    delete nextPhotos[key];
    delete nextPhotos[fallbackKey];
    setPhotoMap(nextPhotos);

    setIkssData((prev) => {
      const nextPosts = { ...(prev.dailyPosts || {}) };
      if (nextPosts[key]) {
        nextPosts[key] = { ...nextPosts[key], posted: false, postedAt: undefined };
      }
      if (nextPosts[fallbackKey]) {
        nextPosts[fallbackKey] = { ...nextPosts[fallbackKey], posted: false, postedAt: undefined };
      }
      return { ...prev, dailyPosts: nextPosts };
    });

    if (lightboxTarget?.hariKe === hariKe) setLightboxTarget(null);
    await persistPhotos(nextPhotos, false);
  };

  // =========================================================================
  // FULL TABLE, ROW, IDENTITY & AKG EDITING HANDLERS (PENCIL EDIT)
  // =========================================================================

  const updateIdentityField = (field: keyof IkssIdentityOverride, value: string) => {
    const profileVal = String((profile as any)[field] || "");
    setIkssData((prev) => {
      const nextOv: IkssIdentityOverride = { ...(prev.identityOverride || {}) };
      if (value === profileVal) {
        delete nextOv[field];
      } else {
        nextOv[field] = value;
      }
      const nextState: IkssDataState = { ...prev };
      if (Object.keys(nextOv).length === 0) {
        delete nextState.identityOverride;
      } else {
        nextState.identityOverride = nextOv;
      }
      return nextState;
    });
  };

  const resetIdentityOverrides = () => {
    setIkssData((prev) => {
      const nextState: IkssDataState = { ...prev, customPeriodeLabel: "" };
      delete nextState.identityOverride;
      return nextState;
    });
    showNotice("info", "Identitas SPPG & Periode dikembalikan mengikuti Profil SPPG.");
  };

  const patchDayOverride = useCallback(
    (
      wsPhotoGroupKey: string,
      hariKe: number,
      patch: Partial<{
        tanggal: string | null;
        namaMenu: string | null;
        namaSppg: string | null;
        idSppg: string | null;
        energi: number | null;
        protein: number | null;
        lemak: number | null;
        kh: number | null;
        serat: number | null;
        jumlahManfaat: number | null;
      }>,
      captionUpdate?: string
    ) => {
      const key = `${wsPhotoGroupKey}__day_${hariKe}`;
      const nowIso = new Date().toISOString();

      setIkssData((prev) => {
        const nextOverrides: Record<string, IkssDayOverride> = {
          ...(prev.dayOverrides || {}),
        };
        const currentDay: IkssDayOverride = { ...(nextOverrides[key] || {}) };

        (Object.keys(patch) as Array<keyof typeof patch>).forEach((k) => {
          const val = patch[k];
          if (val === null || val === undefined) {
            delete (currentDay as any)[k];
          } else {
            (currentDay as any)[k] = val;
          }
        });

        const remainingKeys = Object.keys(currentDay).filter((k) => k !== "updatedAt");
        if (remainingKeys.length === 0) {
          delete nextOverrides[key];
        } else {
          currentDay.updatedAt = nowIso;
          nextOverrides[key] = currentDay;
        }

        const nextState: IkssDataState = { ...prev };
        if (Object.keys(nextOverrides).length === 0) {
          delete nextState.dayOverrides;
        } else {
          nextState.dayOverrides = nextOverrides;
        }

        if (patch.namaMenu !== undefined || captionUpdate !== undefined) {
          const nextPosts: Record<string, IkssDailyPostMeta> = {
            ...(prev.dailyPosts || {}),
          };
          const existingPost: IkssDailyPostMeta = {
            ...(nextPosts[key] || { posted: false }),
          };
          if (patch.namaMenu !== undefined) {
            existingPost.customMenuName = patch.namaMenu || "";
          }
          if (captionUpdate !== undefined) {
            existingPost.caption = captionUpdate;
          }
          existingPost.updatedAt = nowIso;
          nextPosts[key] = existingPost;
          nextState.dailyPosts = nextPosts;
        }

        return nextState;
      });
    },
    [setIkssData]
  );

  const resetDayOverride = (ws: (typeof worksheetsComputed)[number], hariKe: number) => {
    const key = `${ws.photoGroupKey}__day_${hariKe}`;
    const cellPrefix = `${ws.config.id}:${hariKe}:`;
    setBulkCellDrafts((prev) => {
      const copy = { ...prev };
      Object.keys(copy).forEach((k) => {
        if (k.startsWith(cellPrefix)) delete copy[k];
      });
      return copy;
    });
    setBulkCellErrors((prev) => {
      const copy = { ...prev };
      Object.keys(copy).forEach((k) => {
        if (k.startsWith(cellPrefix)) delete copy[k];
      });
      return copy;
    });
    const rowKey = `${ws.config.id}__${hariKe}`;
    if (editingGiziKey === rowKey) {
      setEditingGiziKey(null);
      setGiziRowErrors({});
    }
    if (editingMenuKey === rowKey) {
      setEditingMenuKey(null);
    }

    setIkssData((prev) => {
      const nextOverrides: Record<string, IkssDayOverride> = { ...(prev.dayOverrides || {}) };
      delete nextOverrides[key];
      const nextPosts: Record<string, IkssDailyPostMeta> = { ...(prev.dailyPosts || {}) };
      if (nextPosts[key]) {
        nextPosts[key] = { ...nextPosts[key], customMenuName: "" };
      }
      const nextState: IkssDataState = { ...prev, dailyPosts: nextPosts };
      if (Object.keys(nextOverrides).length === 0) {
        delete nextState.dayOverrides;
      } else {
        nextState.dayOverrides = nextOverrides;
      }
      return nextState;
    });
    showNotice("info", `Baris Hari Ke-${hariKe} dikembalikan ke nilai otomatis dari Food Cost & PM.`);
  };

  const resetAllCurrentManualEdits = (ws: (typeof worksheetsComputed)[number]) => {
    const prefix = `${ws.photoGroupKey}__day_`;
    setEditingGiziKey(null);
    setEditingMenuKey(null);
    setGiziRowErrors({});
    setBulkCellDrafts({});
    setBulkCellErrors({});
    setConfirmResetAllEdits(false);

    setIkssData((prev) => {
      const nextOverrides: Record<string, IkssDayOverride> = { ...(prev.dayOverrides || {}) };
      Object.keys(nextOverrides).forEach((k) => {
        if (k.startsWith(prefix)) delete nextOverrides[k];
      });

      const nextPosts: Record<string, IkssDailyPostMeta> = { ...(prev.dailyPosts || {}) };
      Object.keys(nextPosts).forEach((k) => {
        if (k.startsWith(prefix) && nextPosts[k]) {
          nextPosts[k] = { ...nextPosts[k], customMenuName: "" };
        }
      });

      const nextState: IkssDataState = {
        ...prev,
        dailyPosts: nextPosts,
        customPeriodeLabel: "",
      };
      if (Object.keys(nextOverrides).length === 0) {
        delete nextState.dayOverrides;
      } else {
        nextState.dayOverrides = nextOverrides;
      }
      delete nextState.identityOverride;
      delete nextState.akgOverride;
      delete nextState.customDayCount;
      if (prev.worksheetConfigs && prev.worksheetConfigs.length > 0) {
        nextState.worksheetConfigs = prev.worksheetConfigs.map((item) => {
          if (item.id !== ws.config.id) return item;
          const copy = { ...item };
          delete copy.akgOverride;
          delete copy.customDayCount;
          return copy;
        });
      }
      return nextState;
    });
    showNotice("success", `Seluruh perubahan manual pada lembar "${ws.displayCategoryTitle}" dikembalikan ke perhitungan otomatis.`);
  };

  const handleAddDayRow = (ws: (typeof worksheetsComputed)[number]) => {
    if (ws.totalDays >= 62) {
      showNotice("info", "Maksimal 62 baris hari dapat ditampilkan pada tabel IKSS.");
      return;
    }
    const nextCount = ws.totalDays + 1;
    updateWorksheetConfig(ws.config.id, (cfg) => ({
      ...cfg,
      customDayCount: nextCount,
    }));
    showNotice("success", `Baris Hari Ke-${nextCount} ditambahkan ke lembar "${ws.displayCategoryTitle}".`);
  };

  const handleRemoveLastExtraDay = (ws: (typeof worksheetsComputed)[number]) => {
    if (ws.totalDays <= baseAutoDays) return;
    const removedDay = ws.totalDays;
    const nextCount = ws.totalDays - 1;
    const key = `${ws.photoGroupKey}__day_${removedDay}`;
    updateWorksheetConfig(ws.config.id, (cfg) => {
      const copy = { ...cfg };
      if (nextCount <= baseAutoDays) {
        delete copy.customDayCount;
      } else {
        copy.customDayCount = nextCount;
      }
      return copy;
    });
    setIkssData((prev) => {
      const nextOverrides: Record<string, IkssDayOverride> = { ...(prev.dayOverrides || {}) };
      delete nextOverrides[key];
      const nextState: IkssDataState = { ...prev };
      if (Object.keys(nextOverrides).length === 0) {
        delete nextState.dayOverrides;
      } else {
        nextState.dayOverrides = nextOverrides;
      }
      return nextState;
    });
    showNotice("info", `Baris tambahan Hari Ke-${removedDay} dihapus.`);
  };

  const handleToggleTableEditMode = () => {
    if (isTableEditMode) {
      setIsTableEditMode(false);
      setIsEditingIdentity(false);
      setEditingAkgWsId(null);
      setBulkCellDrafts({});
      setBulkCellErrors({});
      setConfirmResetAllEdits(false);
    } else {
      setEditingGiziKey(null);
      setEditingMenuKey(null);
      setGiziRowErrors({});
      setIsTableEditMode(true);
    }
  };

  // Per-row Pencil edit for IKSS Nilai Gizi
  const startEditGiziRow = (ws: (typeof worksheetsComputed)[number], hariKe: number) => {
    const row = ws.dailyRows.find((r) => r.hariKe === hariKe);
    if (!row) return;
    setEditingGiziKey(`${ws.config.id}__${hariKe}`);
    setGiziRowErrors({});
    setGiziRowDraft({
      tanggal: row.formattedDate,
      namaMenu: row.namaMenu || "",
      energi: row.hasNutrition || row.editedFields.energi ? formatDec1(row.energi) : "",
      protein: row.hasNutrition || row.editedFields.protein ? formatDec1(row.protein) : "",
      lemak: row.hasNutrition || row.editedFields.lemak ? formatDec1(row.lemak) : "",
      kh: row.hasNutrition || row.editedFields.kh ? formatDec1(row.kh) : "",
      serat: row.hasNutrition || row.editedFields.serat ? formatDec1(row.serat) : "",
      jumlahManfaat: String(row.jumlahManfaat),
    });
  };

  const updateGiziRowDraftField = (field: keyof GiziRowDraft, raw: string) => {
    setGiziRowDraft((prev) => ({ ...prev, [field]: raw }));
    if (
      field === "energi" ||
      field === "protein" ||
      field === "lemak" ||
      field === "kh" ||
      field === "serat" ||
      field === "jumlahManfaat"
    ) {
      const parsed = parseNonNegativeNumber(raw, {
        integer: field === "jumlahManfaat",
        label: NUTRIENT_FIELD_LABELS[field],
      });
      setGiziRowErrors((prev) => {
        const next = { ...prev };
        if (!parsed.valid && parsed.error) next[field] = parsed.error;
        else delete next[field];
        return next;
      });
    }
  };

  const saveEditGiziRow = (ws: (typeof worksheetsComputed)[number], hariKe: number) => {
    const row = ws.dailyRows.find((r) => r.hariKe === hariKe);
    if (!row) return;

    const numFields: IkssNumericField[] = [
      "energi",
      "protein",
      "lemak",
      "kh",
      "serat",
      "jumlahManfaat",
    ];
    const errors: Record<string, string> = {};
    const parsedValues: Record<IkssNumericField, number> = {
      energi: 0,
      protein: 0,
      lemak: 0,
      kh: 0,
      serat: 0,
      jumlahManfaat: 0,
    };

    numFields.forEach((f) => {
      const res = parseNonNegativeNumber(giziRowDraft[f], {
        integer: f === "jumlahManfaat",
        label: NUTRIENT_FIELD_LABELS[f],
      });
      if (!res.valid && res.error) {
        errors[f] = res.error;
      } else {
        parsedValues[f] = res.value;
      }
    });

    if (Object.keys(errors).length > 0) {
      setGiziRowErrors(errors);
      showNotice("error", Object.values(errors)[0]);
      return;
    }

    const trimmedTanggal = giziRowDraft.tanggal.trim();
    const trimmedMenu = giziRowDraft.namaMenu.trim();
    const autoMap: Record<IkssNumericField, number> = {
      energi: Math.round(row.autoEnergi * 10) / 10,
      protein: Math.round(row.autoProtein * 10) / 10,
      lemak: Math.round(row.autoLemak * 10) / 10,
      kh: Math.round(row.autoKh * 10) / 10,
      serat: Math.round(row.autoSerat * 10) / 10,
      jumlahManfaat: row.autoJumlahManfaat,
    };

    const patch: Parameters<typeof patchDayOverride>[2] = {
      tanggal:
        trimmedTanggal === "" || trimmedTanggal === row.autoFormattedDate.trim()
          ? null
          : trimmedTanggal,
      namaMenu:
        trimmedMenu === "" || trimmedMenu === (row.autoNamaMenu || "").trim()
          ? null
          : trimmedMenu,
    };

    numFields.forEach((f) => {
      if (f !== "jumlahManfaat" && giziRowDraft[f].trim() === "") {
        patch[f] = null;
        return;
      }
      const isUnchangedFromAuto =
        Math.abs(parsedValues[f] - autoMap[f]) < 0.001 && !row.editedFields[f];
      patch[f] = isUnchangedFromAuto ? null : parsedValues[f];
    });

    patchDayOverride(ws.photoGroupKey, hariKe, patch);
    setEditingGiziKey(null);
    setGiziRowErrors({});
    showNotice("success", `Nilai gizi & jumlah PM Hari Ke-${hariKe} berhasil diperbarui.`);
  };

  // Global Edit Mode live cell handlers
  const handleBulkNumericChange = (
    ws: (typeof worksheetsComputed)[number],
    hariKe: number,
    field: IkssNumericField,
    raw: string
  ) => {
    const cellKey = `${ws.config.id}:${hariKe}:${field}`;
    setBulkCellDrafts((prev) => ({ ...prev, [cellKey]: raw }));

    if (field !== "jumlahManfaat" && raw.trim() === "") {
      setBulkCellErrors((prev) => {
        if (!prev[cellKey]) return prev;
        const copy = { ...prev };
        delete copy[cellKey];
        return copy;
      });
      patchDayOverride(ws.photoGroupKey, hariKe, { [field]: null });
      return;
    }

    const res = parseNonNegativeNumber(raw, {
      integer: field === "jumlahManfaat",
      label: NUTRIENT_FIELD_LABELS[field],
    });

    if (!res.valid && res.error) {
      setBulkCellErrors((prev) => ({ ...prev, [cellKey]: res.error! }));
      return;
    }

    setBulkCellErrors((prev) => {
      if (!prev[cellKey]) return prev;
      const copy = { ...prev };
      delete copy[cellKey];
      return copy;
    });

    patchDayOverride(ws.photoGroupKey, hariKe, { [field]: res.value });
  };

  const handleBulkNumericBlur = (wsId: string, hariKe: number, field: IkssNumericField) => {
    const cellKey = `${wsId}:${hariKe}:${field}`;
    setBulkCellDrafts((prev) => {
      if (!(cellKey in prev)) return prev;
      const copy = { ...prev };
      delete copy[cellKey];
      return copy;
    });
    setBulkCellErrors((prev) => {
      if (!(cellKey in prev)) return prev;
      const copy = { ...prev };
      delete copy[cellKey];
      return copy;
    });
  };

  const handleBulkTextChange = (
    ws: (typeof worksheetsComputed)[number],
    hariKe: number,
    field: "tanggal" | "namaMenu" | "namaSppg" | "idSppg",
    value: string
  ) => {
    const row = ws.dailyRows.find((r) => r.hariKe === hariKe);
    if (!row) return;
    if (field === "tanggal") {
      patchDayOverride(ws.photoGroupKey, hariKe, {
        tanggal: value.trim() === row.autoFormattedDate.trim() ? null : value,
      });
    } else if (field === "namaMenu") {
      patchDayOverride(ws.photoGroupKey, hariKe, {
        namaMenu: value.trim() === (row.autoNamaMenu || "").trim() ? null : value,
      });
    } else if (field === "namaSppg") {
      patchDayOverride(ws.photoGroupKey, hariKe, {
        namaSppg: value.trim() === effectiveIdentity.namaLembaga.trim() ? null : value,
      });
    } else if (field === "idSppg") {
      patchDayOverride(ws.photoGroupKey, hariKe, {
        idSppg: value.trim() === effectiveIdentity.idSPPG.trim() ? null : value,
      });
    }
  };

  const handleAkgNumericChange = (wsId: string, field: IkssAkgNumericField, raw: string) => {
    const cellKey = `${wsId}:akg:${field}`;
    setBulkCellDrafts((prev) => ({ ...prev, [cellKey]: raw }));
    const res = parseNonNegativeNumber(raw, {
      integer: false,
      label: `Rujukan AKG ${NUTRIENT_FIELD_LABELS[field]}`,
    });
    if (!res.valid && res.error) {
      setBulkCellErrors((prev) => ({ ...prev, [cellKey]: res.error! }));
      return;
    }
    setBulkCellErrors((prev) => {
      if (!prev[cellKey]) return prev;
      const copy = { ...prev };
      delete copy[cellKey];
      return copy;
    });
    updateWorksheetConfig(wsId, (cfg) => ({
      ...cfg,
      akgOverride: {
        ...(cfg.akgOverride || {}),
        [field]: res.value,
      },
    }));
  };

  const handleAkgNumericBlur = (wsId: string, field: IkssAkgNumericField) => {
    const cellKey = `${wsId}:akg:${field}`;
    setBulkCellDrafts((prev) => {
      if (!(cellKey in prev)) return prev;
      const copy = { ...prev };
      delete copy[cellKey];
      return copy;
    });
    setBulkCellErrors((prev) => {
      if (!(cellKey in prev)) return prev;
      const copy = { ...prev };
      delete copy[cellKey];
      return copy;
    });
  };

  const resetAkgOverride = (wsId: string) => {
    updateWorksheetConfig(wsId, (cfg) => {
      const copy = { ...cfg };
      delete copy.akgOverride;
      return copy;
    });
    setEditingAkgWsId(null);
    showNotice("info", "Rujukan AKG dikembalikan ke standar otomatis kategori terpilih.");
  };

  // Per-row Pencil edit for IKSS Foto Menu (edits Tanggal, Nama SPPG, ID SPPG, Nama Menu, Catatan)
  const startEditDayMenu = (ws: (typeof worksheetsComputed)[number], hariKe: number) => {
    const row = ws.dailyRows.find((r) => r.hariKe === hariKe);
    const meta = getPostMetaForWsDay(ws.photoGroupKey, hariKe);
    setEditingMenuKey(`${ws.config.id}__${hariKe}`);
    setEditDateDraft(row?.formattedDate ?? "");
    setEditSppgDraft(row?.namaSppg ?? effectiveIdentity.namaLembaga);
    setEditIdSppgDraft(row?.idSppg ?? effectiveIdentity.idSPPG);
    setEditMenuDraft(row?.namaMenu ?? "");
    setEditCaptionDraft(meta?.caption ?? "");
  };

  const saveEditDayMenu = (ws: (typeof worksheetsComputed)[number], hariKe: number) => {
    const row = ws.dailyRows.find((r) => r.hariKe === hariKe);
    if (!row) return;
    const trimmedDate = editDateDraft.trim();
    const trimmedSppg = editSppgDraft.trim();
    const trimmedIdSppg = editIdSppgDraft.trim();
    const trimmedMenu = editMenuDraft.trim();

    patchDayOverride(
      ws.photoGroupKey,
      hariKe,
      {
        tanggal:
          trimmedDate === "" || trimmedDate === row.autoFormattedDate.trim()
            ? null
            : trimmedDate,
        namaSppg:
          trimmedSppg === "" || trimmedSppg === effectiveIdentity.namaLembaga.trim()
            ? null
            : trimmedSppg,
        idSppg:
          trimmedIdSppg === "" || trimmedIdSppg === effectiveIdentity.idSPPG.trim()
            ? null
            : trimmedIdSppg,
        namaMenu:
          trimmedMenu === "" || trimmedMenu === (row.autoNamaMenu || "").trim()
            ? null
            : trimmedMenu,
      },
      editCaptionDraft.trim()
    );
    setEditingMenuKey(null);
    showNotice("success", `Data tabel Hari Ke-${hariKe} berhasil diperbarui.`);
  };

  // =========================================================================
  // SIMPAN KE MENU ARSIP
  // =========================================================================

  const handleSaveIkssToArchive = () => {
    const isNilaiGizi = subView === "nilai_gizi";
    const worksheets: SavedDocArchiveWorksheet[] = worksheetsComputed.map((ws) => {
      if (isNilaiGizi) {
        return {
          title: `REKAPAN NILAI GIZI ${ws.displayCategoryTitle}`,
          sasaranLabel: ws.displayCategoryTitle,
          headers: [
            "Hari / Tanggal",
            "Nama Menu",
            "Energi (kkal)",
            "Protein (g)",
            "Lemak (g)",
            "Karbohidrat (g)",
            "Serat (g)",
            "Jumlah PM",
          ],
          rows: [
            ...ws.dailyRows.map((r) => [
              r.formattedDate,
              r.namaMenu || "-",
              r.hasNutrition ? formatDec1(r.energi) : "-",
              r.hasNutrition ? formatDec1(r.protein) : "-",
              r.hasNutrition ? formatDec1(r.lemak) : "-",
              r.hasNutrition ? formatDec1(r.kh) : "-",
              r.hasNutrition ? formatDec1(r.serat) : "-",
              String(r.jumlahManfaat),
            ]),
            [
              "Nilai Rata-Rata",
              "",
              ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgEnergi) : "-",
              ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgProtein) : "-",
              ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgLemak) : "-",
              ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgKh) : "-",
              ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgSerat) : "-",
              `${ws.summary.avgManfaat} Orang`,
            ],
            [
              ws.effectiveAkg.label,
              "",
              String(ws.effectiveAkg.energi),
              String(ws.effectiveAkg.protein),
              String(ws.effectiveAkg.lemak),
              String(ws.effectiveAkg.kh),
              String(ws.effectiveAkg.serat),
              "",
            ],
            [
              "% Kebutuhan 90-110%",
              "",
              ws.summary.activeNutriDays > 0 ? `${ws.summary.pctEnergi}%` : "-",
              ws.summary.activeNutriDays > 0 ? `${ws.summary.pctProtein}%` : "-",
              ws.summary.activeNutriDays > 0 ? `${ws.summary.pctLemak}%` : "-",
              ws.summary.activeNutriDays > 0 ? `${ws.summary.pctKh}%` : "-",
              ws.summary.activeNutriDays > 0 ? `${ws.summary.pctSerat}%` : "-",
              "",
            ],
          ],
        };
      }

      return {
        title: `REKAPAN FOTO MENU ${ws.displayCategoryTitle}`,
        sasaranLabel: ws.displayCategoryTitle,
        headers: [
          "Hari Ke",
          "Nama SPPG",
          "ID SPPG",
          "Hari / Tanggal",
          "Nama Menu",
          "Catatan",
          "Status Foto",
        ],
        rows: ws.dailyRows.map((r) => {
          const hasPhoto = Boolean(getPhotoForWsDay(ws.photoGroupKey, r.hariKe));
          const meta = getPostMetaForWsDay(ws.photoGroupKey, r.hariKe);
          const status = !hasPhoto
            ? "Belum Foto"
            : meta?.posted
            ? "Diposting"
            : "Draft";
          return [
            String(r.hariKe),
            (r.namaSppg || effectiveIdentity.namaLembaga || "-").toUpperCase(),
            (r.idSppg || effectiveIdentity.idSPPG || "-").toUpperCase(),
            r.formattedDate,
            r.namaMenu || "-",
            r.caption || "-",
            status,
          ];
        }),
      };
    });

    const titlePrefix = isNilaiGizi ? "IKSS Nilai Gizi" : "IKSS Foto Menu";
    const sasaranSummary =
      worksheetsComputed.length > 1
        ? `${worksheetsComputed.length} Kelompok Sasaran (${worksheetsComputed
            .map((w) => w.displayCategoryTitle)
            .slice(0, 2)
            .join(", ")}${worksheetsComputed.length > 2 ? ", ..." : ""})`
        : displayCategoryTitle;

    const archive: SavedDocArchive = {
      id: `arsip_ikss_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      source: "ikss",
      subType: subView,
      title: `${titlePrefix} — ${sasaranSummary}`,
      periodeLabel: displayPeriodeBulan,
      sppgName: effectiveIdentity.namaLembaga || profile.namaLembaga || "",
      idSppg: effectiveIdentity.idSPPG || profile.idSPPG || "",
      sasaranSummary,
      savedAt: new Date().toISOString(),
      worksheetConfigs: worksheetConfigs.map((c) => ({ ...c })),
      worksheets,
    };

    onSaveToArchive?.(archive);
    setSavedArchiveNotice(
      `Berhasil disimpan ke Menu Arsip: ${archive.title} (${worksheets.length} lembar kerja).`
    );
  };

  // =========================================================================
  // PRINT / PDF / EXCEL EXPORTS
  // =========================================================================

  const activePrintId =
    subView === "nilai_gizi" ? "print-area-ikss-gizi" : "print-area-ikss-foto";
  const multiSuffix =
    worksheetsComputed.length > 1
      ? `${worksheetsComputed.length}_Kelompok_Sasaran`
      : displayCategoryTitle;
  const activePrintTitle =
    subView === "nilai_gizi"
      ? `IKSS Rekapan Nilai Gizi - ${worksheetsComputed.length > 1 ? `${worksheetsComputed.length} Lembar Sasaran` : displayCategoryTitle}`
      : `IKSS Rekapan Foto Menu - ${worksheetsComputed.length > 1 ? `${worksheetsComputed.length} Lembar Sasaran` : displayCategoryTitle}`;
  const activeFilename =
    subView === "nilai_gizi"
      ? `IKSS_Nilai_Gizi_${multiSuffix.replace(/[^A-Za-z0-9]+/g, "_")}`
      : `IKSS_Foto_Menu_${multiSuffix.replace(/[^A-Za-z0-9]+/g, "_")}`;

  const handlePrint = () => {
    const needsUiReset =
      photoFilter !== "semua" ||
      isTableEditMode ||
      isEditingIdentity ||
      editingAkgWsId !== null ||
      editingGiziKey !== null ||
      editingMenuKey !== null;

    if (needsUiReset) {
      if (photoFilter !== "semua") setPhotoFilter("semua");
      setIsTableEditMode(false);
      setIsEditingIdentity(false);
      setEditingAkgWsId(null);
      setEditingGiziKey(null);
      setEditingMenuKey(null);
      setTimeout(() => {
        printElementSmart(activePrintId, {
          title: activePrintTitle,
          paperSize,
          orientation: "auto",
        });
      }, 80);
      return;
    }
    printElementSmart(activePrintId, {
      title: activePrintTitle,
      paperSize,
      orientation: "auto",
    });
  };

  const handleExportPdf = async () => {
    if (isExportingPdf) return;
    const needsUiReset =
      photoFilter !== "semua" ||
      editingGiziKey !== null ||
      editingMenuKey !== null ||
      editingAkgWsId !== null;
    if (needsUiReset) {
      if (photoFilter !== "semua") setPhotoFilter("semua");
      setEditingGiziKey(null);
      setEditingMenuKey(null);
      setEditingAkgWsId(null);
      await new Promise<void>((r) => setTimeout(r, 70));
    }
    setIsExportingPdf(true);
    setExportStatusMsg("Menyiapkan dokumen PDF presisi...");
    await new Promise<void>((r) => setTimeout(r, 40));
    try {
      await exportToPDF(
        activePrintId,
        activeFilename,
        100,
        (msg) => setExportStatusMsg(msg),
        {
          paperSize,
          orientation: "auto",
        }
      );
    } finally {
      setIsExportingPdf(false);
      setExportStatusMsg(null);
    }
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const usedSheetNames = new Set<string>();
    const makeSheetName = (rawTitle: string, idx: number) => {
      const cleaned = rawTitle
        .replace(/[\\/?*[\]:]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
      let candidate = (cleaned || `Lembar ${idx + 1}`).slice(0, 31);
      if (usedSheetNames.has(candidate)) {
        const tag = `_${idx + 1}`;
        candidate = `${candidate.slice(0, 31 - tag.length)}${tag}`;
      }
      usedSheetNames.add(candidate);
      return candidate;
    };

    worksheetsComputed.forEach((ws, wsIdx) => {
      const identityRows = [
        [
          subView === "nilai_gizi"
            ? `REKAPAN NILAI GIZI ${ws.displayCategoryTitle}`
            : `REKAPAN MENU ${ws.displayCategoryTitle}`,
        ],
        ["Nama SPPG", ":", (effectiveIdentity.namaLembaga || "").toUpperCase()],
        ["ID SPPG", ":", (effectiveIdentity.idSPPG || "").toUpperCase()],
        ["Alamat", ":", effectiveIdentity.alamat || ""],
        ["Kab/Kota", ":", (effectiveIdentity.kabKota || "").toUpperCase()],
        ["Provinsi", ":", (effectiveIdentity.provinsi || "").toUpperCase()],
        ["Nama SPPI/AG", ":", effectiveIdentity.namaAhliGizi || ""],
        ["Nomor Handphone", ":", effectiveIdentity.noHpSPPI || ""],
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
          ...ws.dailyRows.map((r) => [
            r.formattedDate,
            r.namaMenu || "-",
            r.hasNutrition ? Number(formatDec1(r.energi)) : "-",
            r.hasNutrition ? Number(formatDec1(r.protein)) : "-",
            r.hasNutrition ? Number(formatDec1(r.lemak)) : "-",
            r.hasNutrition ? Number(formatDec1(r.kh)) : "-",
            r.hasNutrition ? Number(formatDec1(r.serat)) : "-",
            r.jumlahManfaat,
          ]),
          [
            "Total Nilai Gizi",
            "",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.totalEnergi)) : "-",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.totalProtein)) : "-",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.totalLemak)) : "-",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.totalKh)) : "-",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.totalSerat)) : "-",
            ws.summary.totalManfaat,
          ],
          [
            "Jumlah Hari",
            "",
            `${ws.summary.activeNutriDays} Hari`,
            `${ws.summary.activeNutriDays} Hari`,
            `${ws.summary.activeNutriDays} Hari`,
            `${ws.summary.activeNutriDays} Hari`,
            `${ws.summary.activeNutriDays} Hari`,
            `${ws.summary.activePmDays} Hari`,
          ],
          [
            "Nilai Rata-Rata",
            "",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.avgEnergi)) : "-",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.avgProtein)) : "-",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.avgLemak)) : "-",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.avgKh)) : "-",
            ws.summary.activeNutriDays > 0 ? Number(formatDec1(ws.summary.avgSerat)) : "-",
            `${ws.summary.avgManfaat} Orang`,
          ],
          [
            ws.effectiveAkg.label,
            "",
            ws.effectiveAkg.energi,
            ws.effectiveAkg.protein,
            ws.effectiveAkg.lemak,
            ws.effectiveAkg.kh,
            ws.effectiveAkg.serat,
            "",
          ],
          [
            "% Kebutuhan 90-110%",
            "",
            ws.summary.activeNutriDays > 0 ? `${ws.summary.pctEnergi}%` : "-",
            ws.summary.activeNutriDays > 0 ? `${ws.summary.pctProtein}%` : "-",
            ws.summary.activeNutriDays > 0 ? `${ws.summary.pctLemak}%` : "-",
            ws.summary.activeNutriDays > 0 ? `${ws.summary.pctKh}%` : "-",
            ws.summary.activeNutriDays > 0 ? `${ws.summary.pctSerat}%` : "-",
            "",
          ],
        ];
        const sheet = XLSX.utils.aoa_to_sheet(sheetData);
        sheet["!cols"] = [
          { wch: 24 },
          { wch: 44 },
          { wch: 14 },
          { wch: 14 },
          { wch: 14 },
          { wch: 16 },
          { wch: 13 },
          { wch: 15 },
        ];
        const name =
          worksheetsComputed.length === 1
            ? "IKSS Nilai Gizi"
            : makeSheetName(ws.displayCategoryTitle, wsIdx);
        XLSX.utils.book_append_sheet(wb, sheet, name);
      } else {
        const sheetData: any[][] = [
          ...identityRows,
          [
            "Hari Ke",
            "Nama SPPG",
            "ID SPPG",
            "Hari / Tanggal",
            "Nama Menu",
            "Catatan",
            "Status Foto Menu",
          ],
          ...ws.dailyRows.map((r) => {
            const hasPhoto = Boolean(getPhotoForWsDay(ws.photoGroupKey, r.hariKe));
            const meta = getPostMetaForWsDay(ws.photoGroupKey, r.hariKe);
            const status = !hasPhoto
              ? "Belum Diunggah"
              : meta?.posted
              ? `Diposting (${formatShortTimestamp(meta.postedAt)})`
              : "Draft (Belum Diposting)";
            return [
              r.hariKe,
              (r.namaSppg || effectiveIdentity.namaLembaga || "").toUpperCase(),
              (r.idSppg || effectiveIdentity.idSPPG || "").toUpperCase(),
              r.formattedDate,
              r.namaMenu || "-",
              r.caption || "-",
              status,
            ];
          }),
        ];
        const sheet = XLSX.utils.aoa_to_sheet(sheetData);
        sheet["!cols"] = [
          { wch: 10 },
          { wch: 32 },
          { wch: 14 },
          { wch: 24 },
          { wch: 44 },
          { wch: 24 },
          { wch: 28 },
        ];
        const name =
          worksheetsComputed.length === 1
            ? "IKSS Foto Menu"
            : makeSheetName(ws.displayCategoryTitle, wsIdx);
        XLSX.utils.book_append_sheet(wb, sheet, name);
      }
    });

    XLSX.writeFile(wb, `${activeFilename}.xlsx`);
  };

  // Visibility flags for interactive inputs (automatically stripped during PDF/Print Preview)
  const showInlineTableInputs = isTableEditMode && !isExportingPdf && !showPrintModal;
  const showIdentityInputs = (isTableEditMode || isEditingIdentity) && !isExportingPdf && !showPrintModal;

  const activeValidationMessages = useMemo(() => {
    const msgs = [
      ...Object.values(giziRowErrors),
      ...Object.values(bulkCellErrors),
    ];
    return Array.from(new Set(msgs));
  }, [giziRowErrors, bulkCellErrors]);

  const renderIdentityBlock = () => {
    const identityFields: Array<{
      key: keyof IkssIdentityOverride | "periode";
      label: string;
      val: string;
      displayVal: string;
      placeholder: string;
      uppercase?: boolean;
    }> = [
      {
        key: "namaLembaga",
        label: "Nama SPPG",
        val: effectiveIdentity.namaLembaga,
        displayVal: (effectiveIdentity.namaLembaga || "-").toUpperCase(),
        placeholder: profile.namaLembaga || "Nama SPPG",
        uppercase: true,
      },
      {
        key: "idSPPG",
        label: "ID SPPG",
        val: effectiveIdentity.idSPPG,
        displayVal: (effectiveIdentity.idSPPG || "-").toUpperCase(),
        placeholder: profile.idSPPG || "ID SPPG",
        uppercase: true,
      },
      {
        key: "alamat",
        label: "Alamat",
        val: effectiveIdentity.alamat,
        displayVal: effectiveIdentity.alamat || "-",
        placeholder: profile.alamat || "Alamat SPPG",
      },
      {
        key: "kabKota",
        label: "Kab/Kota",
        val: effectiveIdentity.kabKota,
        displayVal: (effectiveIdentity.kabKota || "-").toUpperCase(),
        placeholder: profile.kabKota || "Kabupaten / Kota",
        uppercase: true,
      },
      {
        key: "provinsi",
        label: "Provinsi",
        val: effectiveIdentity.provinsi,
        displayVal: (effectiveIdentity.provinsi || "-").toUpperCase(),
        placeholder: profile.provinsi || "Provinsi",
        uppercase: true,
      },
      {
        key: "namaAhliGizi",
        label: "Nama SPPI/AG",
        val: effectiveIdentity.namaAhliGizi,
        displayVal: effectiveIdentity.namaAhliGizi || "-",
        placeholder: profile.namaAhliGizi || "Nama SPPI / Ahli Gizi",
      },
      {
        key: "noHpSPPI",
        label: "Nomor Handphone",
        val: effectiveIdentity.noHpSPPI,
        displayVal: effectiveIdentity.noHpSPPI || "-",
        placeholder: profile.noHpSPPI || "Nomor Handphone",
      },
      {
        key: "periode",
        label: "Periode/Bulan",
        val: ikssData.customPeriodeLabel ?? "",
        displayVal: displayPeriodeBulan,
        placeholder: autoPeriodeBulan,
        uppercase: true,
      },
    ];

    return (
      <div className="relative border-x border-black bg-white text-black">
        <div className="no-print absolute top-1 right-1.5 z-10 flex items-center gap-1 font-sans">
          {(hasIdentityOverride || Boolean(ikssData.customPeriodeLabel)) && (
            <button
              type="button"
              onClick={resetIdentityOverrides}
              title="Kembalikan Identitas ke Profil SPPG"
              className="px-1.5 py-0.5 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 text-[10px] font-semibold inline-flex items-center gap-1 cursor-pointer"
            >
              <RotateCcw className="w-3 h-3" />
              <span>Profil</span>
            </button>
          )}
          <button
            type="button"
            onClick={() => setIsEditingIdentity((v) => !v)}
            title={showIdentityInputs ? "Selesai Edit Identitas" : "Edit Identitas & Judul Tabel"}
            className={`px-2 py-0.5 rounded text-[10px] font-bold inline-flex items-center gap-1 cursor-pointer transition ${
              showIdentityInputs
                ? "bg-emerald-600 text-white hover:bg-emerald-700"
                : "bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-800"
            }`}
          >
            {showIdentityInputs ? (
              <>
                <Check className="w-3 h-3" />
                <span>Selesai</span>
              </>
            ) : (
              <>
                <Pencil className="w-3 h-3" />
                <span>Edit Identitas</span>
              </>
            )}
          </button>
        </div>

        <table
          className="w-full text-[11px] leading-snug bg-white text-black"
          style={{ fontFamily: "Arial, Helvetica, sans-serif", borderCollapse: "collapse" }}
        >
          <tbody>
            {identityFields.map((item) => (
              <tr key={item.key}>
                <td className="px-2 py-0.5 font-bold w-[165px] align-middle">{item.label}</td>
                <td className="px-1 py-0.5 font-bold w-[14px] text-center align-middle">:</td>
                <td className="px-1.5 py-0.5 font-bold align-middle pr-28 print:pr-1.5">
                  {showIdentityInputs ? (
                    <>
                      <input
                        type="text"
                        value={item.val}
                        placeholder={item.placeholder}
                        onChange={(e) => {
                          if (item.key === "periode") {
                            setIkssData((prev) => ({
                              ...prev,
                              customPeriodeLabel: e.target.value,
                            }));
                          } else {
                            updateIdentityField(item.key, e.target.value);
                          }
                        }}
                        className={`print:hidden w-full max-w-[460px] px-1.5 py-0.5 text-[11px] font-bold text-slate-900 bg-amber-50/70 border border-amber-300 rounded focus:outline-none focus:bg-white focus:border-emerald-600 ${
                          item.uppercase ? "uppercase" : ""
                        }`}
                      />
                      <span className="hidden print:inline">{item.displayVal}</span>
                    </>
                  ) : (
                    item.displayVal
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  const renderPctCell = (pctVal: number, hasData: boolean) => {
    if (!hasData) {
      return (
        <td className="border border-black px-1.5 py-1.5 text-center font-bold tabular-nums text-[11px] bg-white text-slate-400 print:text-black">
          -
        </td>
      );
    }
    const isGood = pctVal >= 90 && pctVal <= 110;
    return (
      <td
        className={`border border-black px-1.5 py-1.5 text-center font-bold tabular-nums text-[11px] ${
          isGood ? "bg-[#00B050] text-black" : "bg-[#FF0000] text-black"
        }`}
      >
        {pctVal > 0 ? `${pctVal}%` : "0%"}
      </td>
    );
  };

  // Active lightbox row info
  const activeLightboxWs = lightboxTarget
    ? worksheetsComputed.find((w) => w.config.id === lightboxTarget.wsId) || activeWs
    : null;
  const activeLightboxRow =
    lightboxTarget && activeLightboxWs
      ? activeLightboxWs.dailyRows.find((r) => r.hariKe === lightboxTarget.hariKe) || null
      : null;
  const activeLightboxPhoto =
    lightboxTarget && activeLightboxWs
      ? getPhotoForWsDay(activeLightboxWs.photoGroupKey, lightboxTarget.hariKe)
      : "";
  const activeLightboxMeta =
    lightboxTarget && activeLightboxWs
      ? getPostMetaForWsDay(activeLightboxWs.photoGroupKey, lightboxTarget.hariKe)
      : undefined;

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
              Terhubung otomatis dengan <strong>Profil SPPG</strong>, <strong>Master Menu</strong>, <strong>Penerima Manfaat</strong>, dan <strong>Food Cost</strong>. Dapat menampilkan beberapa lembar kerja sekaligus sesuai Kelompok Sasaran yang dipilih dan disimpan ke Menu Arsip.
            </p>
          </div>

          {/* Sub-view Segmented Control */}
          <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 self-start lg:self-center shrink-0">
            <button
              type="button"
              onClick={() => handleSwitchSubView("nilai_gizi")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition cursor-pointer flex items-center gap-1.5 whitespace-nowrap ${
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
              onClick={() => handleSwitchSubView("foto_menu")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition cursor-pointer flex items-center gap-1.5 whitespace-nowrap ${
                subView === "foto_menu"
                  ? "bg-white text-emerald-700 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <Camera className="w-3.5 h-3.5" />
              IKSS Foto Menu Makanan
              <span className="ml-0.5 text-[11px] font-extrabold tabular-nums text-emerald-700">
                ({totalPhotoAggregate.posted}/{totalPhotoAggregate.total})
              </span>
            </button>
          </div>
        </div>

        {/* Action Toolbar */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-slate-100">
          <div className="flex flex-wrap items-center gap-2.5">
            <button
              type="button"
              onClick={handleToggleTableEditMode}
              className={`px-3.5 py-1.5 rounded-xl border text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap transition shadow-xs ${
                isTableEditMode
                  ? "bg-amber-500 hover:bg-amber-600 border-amber-500 text-white"
                  : "bg-amber-50 hover:bg-amber-100 border-amber-300 text-amber-900"
              }`}
            >
              {isTableEditMode ? (
                <>
                  <Check className="w-3.5 h-3.5" />
                  <span>Selesai Edit Tabel</span>
                </>
              ) : (
                <>
                  <Pencil className="w-3.5 h-3.5" />
                  <span>Edit Semua Tabel</span>
                </>
              )}
            </button>

            <button
              type="button"
              onClick={() => setShowCategoryDrawer((v) => !v)}
              className={`px-3 py-1.5 rounded-xl border text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap transition ${
                showCategoryDrawer
                  ? "bg-emerald-50 border-emerald-300 text-emerald-900"
                  : "border-slate-200 hover:bg-slate-50 text-slate-700"
              }`}
            >
              <Layers className="w-3.5 h-3.5 text-emerald-600" />
              <span>
                {showCategoryDrawer
                  ? "Sembunyikan Pilihan Sasaran"
                  : `Kelompok Sasaran & Lembar Kerja (${worksheetsComputed.length})`}
              </span>
            </button>

            <label className="inline-flex items-center gap-2 text-xs font-semibold text-slate-700 cursor-pointer select-none whitespace-nowrap">
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
              onClick={handleSaveIkssToArchive}
              className="px-3 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap shadow-xs"
            >
              <Archive className="w-3.5 h-3.5" />
              Simpan Arsip
            </button>
            <button
              type="button"
              onClick={() => {
                setEditingGiziKey(null);
                setEditingMenuKey(null);
                setEditingAkgWsId(null);
                setShowPrintModal(true);
              }}
              className="px-3 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap"
            >
              <Eye className="w-3.5 h-3.5" />
              Pratinjau Cetak
            </button>
            <button
              type="button"
              onClick={handleExportPdf}
              disabled={isExportingPdf}
              className="px-3 py-2 bg-rose-600 hover:bg-rose-700 disabled:opacity-60 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap"
            >
              {isExportingPdf ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <FileDown className="w-3.5 h-3.5" />
              )}
              {isExportingPdf ? "Memproses PDF..." : "Unduh PDF"}
            </button>
            <button
              type="button"
              onClick={handleExportExcel}
              className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap"
            >
              <Download className="w-3.5 h-3.5" />
              Unduh Excel
            </button>
            <button
              type="button"
              onClick={handlePrint}
              className="px-3 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap"
            >
              <Printer className="w-3.5 h-3.5" />
              Cetak Langsung
            </button>
          </div>
        </div>

        {/* Saved Archive Confirmation Banner */}
        {savedArchiveNotice && (
          <div className="px-3.5 py-2.5 rounded-xl bg-teal-50 border border-teal-200 text-teal-950 text-xs font-semibold flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-teal-600 shrink-0" />
              <span>{savedArchiveNotice}</span>
            </div>
            <div className="flex items-center gap-2">
              {onGoToArsip && (
                <button
                  type="button"
                  onClick={onGoToArsip}
                  className="px-2.5 py-1 rounded-lg bg-teal-600 hover:bg-teal-700 text-white text-[11px] font-bold inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                >
                  <span>Buka Menu Arsip</span>
                  <ArrowRight className="w-3 h-3" />
                </button>
              )}
              <button
                type="button"
                onClick={() => setSavedArchiveNotice(null)}
                className="p-1 rounded-lg hover:bg-teal-100 text-teal-700 cursor-pointer"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        )}

        {/* Multi-Worksheet Focus Strip (always visible when > 1 worksheet active) */}
        {worksheetsComputed.length > 1 && (
          <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center gap-1.5">
            <span className="text-[11px] font-bold text-slate-500 mr-1">
              Lembar Kerja Ditampilkan ({worksheetsComputed.length}):
            </span>
            {worksheetsComputed.map((ws, idx) => {
              const isFocused = ws.config.id === activeWs.config.id;
              return (
                <div
                  key={ws.config.id}
                  className={`inline-flex items-center gap-1 rounded-lg border text-xs font-bold transition ${
                    isFocused
                      ? "bg-emerald-600 border-emerald-600 text-white"
                      : "bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100"
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => setActiveWorksheetId(ws.config.id)}
                    className="px-2.5 py-1 cursor-pointer"
                  >
                    {idx + 1}. {ws.displayCategoryTitle}
                  </button>
                  <button
                    type="button"
                    onClick={() => handleRemoveWorksheet(ws.config.id)}
                    title={`Tutup lembar ${ws.displayCategoryTitle}`}
                    className={`pr-1.5 py-1 cursor-pointer ${
                      isFocused ? "text-emerald-100 hover:text-white" : "text-slate-400 hover:text-rose-600"
                    }`}
                  >
                    <X className="w-3 h-3" />
                  </button>
                </div>
              );
            })}
          </div>
        )}

        {/* Active Full-Table Edit Helper Bar */}
        {isTableEditMode && (
          <div className="p-3.5 rounded-xl bg-amber-50/90 border border-amber-200 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
            <div className="flex items-start gap-2 text-xs text-amber-950">
              <Pencil className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
              <div>
                <div className="font-bold">
                  Mode Input &amp; Edit Semua Tabel Aktif — Fokus Lembar: {displayCategoryTitle}
                </div>
                <p className="text-[11px] text-amber-900/90 mt-0.5">
                  Anda dapat mengubah langsung <strong>Identitas SPPG</strong>, <strong>Hari/Tanggal</strong>, <strong>Nama Menu</strong>, seluruh <strong>Nilai Gizi</strong>, <strong>Jumlah PM</strong>, dan <strong>Rujukan AKG</strong> pada semua lembar kerja yang tampil.
                </p>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2 shrink-0">
              <button
                type="button"
                onClick={() => handleAddDayRow(activeWs)}
                className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-emerald-50 border border-emerald-300 text-emerald-800 text-[11px] font-bold inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
              >
                <Plus className="w-3.5 h-3.5" />
                Tambah Baris Hari
              </button>

              {totalDays > baseAutoDays && (
                <button
                  type="button"
                  onClick={() => handleRemoveLastExtraDay(activeWs)}
                  className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-rose-50 border border-rose-200 text-rose-700 text-[11px] font-bold inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                >
                  <Minus className="w-3.5 h-3.5" />
                  Hapus Baris Ke-{totalDays}
                </button>
              )}

              {manualEditStats.totalCount > 0 && (
                confirmResetAllEdits ? (
                  <div className="inline-flex items-center gap-1.5 bg-white px-2 py-1 rounded-lg border border-rose-300">
                    <span className="text-[11px] font-bold text-rose-700">Reset semua edit?</span>
                    <button
                      type="button"
                      onClick={() => resetAllCurrentManualEdits(activeWs)}
                      className="px-2 py-0.5 rounded bg-rose-600 hover:bg-rose-700 text-white text-[11px] font-bold cursor-pointer"
                    >
                      Ya
                    </button>
                    <button
                      type="button"
                      onClick={() => setConfirmResetAllEdits(false)}
                      className="px-2 py-0.5 rounded bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-semibold cursor-pointer"
                    >
                      Batal
                    </button>
                  </div>
                ) : (
                  <button
                    type="button"
                    onClick={() => setConfirmResetAllEdits(true)}
                    className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-rose-50 border border-rose-200 text-rose-700 text-[11px] font-bold inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    Reset ke Otomatis ({manualEditStats.editedDays} Hari)
                  </button>
                )
              )}
            </div>
          </div>
        )}

        {/* Inline non-negative validation alert */}
        {activeValidationMessages.length > 0 && (
          <div
            role="alert"
            className="px-3.5 py-2.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-900 text-xs font-semibold flex items-center gap-2"
          >
            <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
            <span>{activeValidationMessages[0]}</span>
          </div>
        )}

        {/* PDF export status bar */}
        {exportStatusMsg && (
          <div className="px-3.5 py-2 rounded-xl bg-indigo-50 border border-indigo-200 text-indigo-800 text-xs font-semibold flex items-center gap-2">
            <Loader2 className="w-3.5 h-3.5 animate-spin text-indigo-600 shrink-0" />
            <span>{exportStatusMsg}</span>
          </div>
        )}

        {/* FLEXIBLE MULTI-WORKSHEET & CATEGORY SELECTOR */}
        {showCategoryDrawer && (
          <div className="pt-4 border-t border-slate-100 space-y-4">
            {/* MULTI-WORKSHEET PANEL BY KELOMPOK SASARAN */}
            <div className="p-4 rounded-xl bg-emerald-50/70 border border-emerald-200 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div>
                  <span className="text-xs font-extrabold text-emerald-950 block">
                    Tampilkan Multi-Lembar Kerja Sekaligus Berdasarkan Kelompok Sasaran
                  </span>
                  <span className="text-[11px] text-emerald-800/80">
                    Centang preset atau kelompok sasaran di bawah untuk menambah/mengurangi lembar kerja terpisah di halaman IKSS. Seluruh lembar kerja yang tampil dapat diunduh, dicetak, atau disimpan ke Menu Arsip sekaligus.
                  </span>
                </div>
                <div className="flex flex-wrap items-center gap-1.5">
                  <button
                    type="button"
                    onClick={handleSplitAllSasaranWorksheets}
                    className="px-2.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-bold cursor-pointer whitespace-nowrap"
                  >
                    Pisah Semua 12 Kelompok
                  </button>
                  <button
                    type="button"
                    onClick={handleAddCustomWorksheet}
                    className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-emerald-100 border border-emerald-300 text-emerald-800 text-[11px] font-bold inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                  >
                    <Plus className="w-3 h-3" />
                    Tambah Lembar Kosong
                  </button>
                </div>
              </div>

              <div className="space-y-1.5">
                <span className="text-[10px] font-bold text-emerald-900 uppercase tracking-wider block">
                  Tambah / Hapus Lembar dari Preset Kategori:
                </span>
                <div className="flex flex-wrap gap-1.5">
                  {QUICK_PRESETS.map((preset) => {
                    const pSet = new Set(preset.sasaranIds);
                    const hasSheet = worksheetConfigs.some(
                      (w) =>
                        w.sasaranIds.length === preset.sasaranIds.length &&
                        w.sasaranIds.every((id) => pSet.has(id))
                    );
                    return (
                      <button
                        key={`multi_${preset.id}`}
                        type="button"
                        onClick={() => togglePresetWorksheet(preset)}
                        className={`px-2.5 py-1.5 rounded-lg text-xs font-bold border inline-flex items-center gap-1.5 transition cursor-pointer ${
                          hasSheet
                            ? "bg-emerald-600 border-emerald-600 text-white"
                            : "bg-white border-emerald-200 text-emerald-900 hover:bg-emerald-100/60"
                        }`}
                      >
                        <span
                          className={`w-3.5 h-3.5 rounded flex items-center justify-center border ${
                            hasSheet
                              ? "bg-white border-white text-emerald-700"
                              : "border-emerald-300 bg-white"
                          }`}
                        >
                          {hasSheet && <Check className="w-2.5 h-2.5" />}
                        </span>
                        <span>{preset.label}</span>
                      </button>
                    );
                  })}
                </div>
              </div>

              <div className="space-y-1.5">
                <span className="text-[10px] font-bold text-emerald-900 uppercase tracking-wider block">
                  Tambah / Hapus Lembar Terpisah Per Kelompok Sasaran:
                </span>
                <div className="flex flex-wrap gap-1.5">
                  {SASARAN_KATEGORI_LIST.map((sas) => {
                    const hasIndividual = worksheetConfigs.some(
                      (w) => w.sasaranIds.length === 1 && w.sasaranIds[0] === sas.id
                    );
                    return (
                      <button
                        key={`ind_${sas.id}`}
                        type="button"
                        onClick={() => toggleIndividualSasaranWorksheet(sas.id, sas.label)}
                        className={`px-2.5 py-1 rounded-lg border text-[11px] font-semibold inline-flex items-center gap-1.5 transition cursor-pointer ${
                          hasIndividual
                            ? "bg-teal-700 border-teal-700 text-white"
                            : "bg-white border-slate-200 text-slate-700 hover:bg-white"
                        }`}
                      >
                        <span
                          className={`w-3 h-3 rounded flex items-center justify-center border ${
                            hasIndividual
                              ? "bg-white border-white text-teal-700"
                              : "border-slate-300 bg-white"
                          }`}
                        >
                          {hasIndividual && <Check className="w-2 h-2" />}
                        </span>
                        <span>{sas.shortLabel}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* EDITOR FOR CURRENTLY FOCUSED WORKSHEET */}
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-[11px] font-extrabold text-slate-700 uppercase tracking-wider">
                  Pengaturan Gabungan Isi untuk Lembar Fokus: {displayCategoryTitle}
                </span>
                {worksheetsComputed.length > 1 && (
                  <div className="flex flex-wrap gap-1">
                    {worksheetsComputed.map((ws, idx) => (
                      <button
                        key={`focus_${ws.config.id}`}
                        type="button"
                        onClick={() => setActiveWorksheetId(ws.config.id)}
                        className={`px-2 py-0.5 rounded text-[11px] font-bold cursor-pointer ${
                          ws.config.id === activeWs.config.id
                            ? "bg-slate-800 text-white"
                            : "bg-white border border-slate-200 text-slate-600 hover:bg-slate-100"
                        }`}
                      >
                        Lembar {idx + 1}: {ws.displayCategoryTitle}
                      </button>
                    ))}
                  </div>
                )}
              </div>

              <div className="space-y-1.5">
                <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                  Ganti Lembar Fokus dengan Preset Gabungan
                </span>
                <div className="flex flex-wrap gap-1.5">
                  {QUICK_PRESETS.map((preset) => {
                    const isActive =
                      preset.sasaranIds.length === selectedSasaranIdsForFocus.length &&
                      preset.sasaranIds.every((id) => selectedSasaranSet.has(id));
                    return (
                      <button
                        key={preset.id}
                        type="button"
                        onClick={() => applyPreset(preset)}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition cursor-pointer ${
                          isActive
                            ? "bg-emerald-600 border-emerald-600 text-white"
                            : "bg-white border-slate-200 text-slate-700 hover:bg-slate-100"
                        }`}
                      >
                        {preset.label}
                      </button>
                    );
                  })}
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                <div className="p-3.5 rounded-xl bg-white border border-slate-200/80 space-y-2.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-[11px] font-bold text-slate-700 uppercase tracking-wider">
                      Sumber Perencanaan Food Cost (Bisa Gabung Bebas)
                    </span>
                    <button
                      type="button"
                      onClick={() =>
                        updateWorksheetConfig(activeWs.config.id, (cfg) => ({
                          ...cfg,
                          foodCostTypes: [
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
                      const checked = selectedFcTypeSet.has(fcType.id);
                      const hasData = populatedFcTypes.has(fcType.id);
                      return (
                        <button
                          key={fcType.id}
                          type="button"
                          onClick={() => toggleFoodCostType(fcType.id)}
                          className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 transition cursor-pointer ${
                            checked
                              ? "bg-emerald-50 border-emerald-500 text-emerald-900"
                              : "bg-white border-slate-200 text-slate-500 hover:bg-slate-50"
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
                            <span className="text-[10px] font-bold text-emerald-600">•</span>
                          )}
                        </button>
                      );
                    })}

                    {customFcTables.map((tbl: any) => {
                      const customKey = `custom:${tbl.id}`;
                      const checked = selectedFcTypeSet.has(customKey);
                      return (
                        <button
                          key={customKey}
                          type="button"
                          onClick={() => toggleFoodCostType(customKey)}
                          className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 transition cursor-pointer ${
                            checked
                              ? "bg-indigo-50 border-indigo-500 text-indigo-900"
                              : "bg-white border-slate-200 text-slate-500 hover:bg-slate-50"
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

                <div className="p-3.5 rounded-xl bg-white border border-slate-200/80 space-y-2.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-[11px] font-bold text-slate-700 uppercase tracking-wider">
                      Kelompok Sasaran Penerima Manfaat (Gabung Bebas)
                    </span>
                    <button
                      type="button"
                      onClick={() => {
                        setEditingGiziKey(null);
                        setEditingMenuKey(null);
                        setBulkCellDrafts({});
                        setBulkCellErrors({});
                        updateWorksheetConfig(activeWs.config.id, (cfg) => {
                          const copy: IkssWorksheetConfig = {
                            ...cfg,
                            sasaranIds: SASARAN_KATEGORI_LIST.map((s) => s.id),
                            title: "SEMUA KATEGORI PENERIMA MANFAAT",
                          };
                          delete copy.akgOverride;
                          return copy;
                        });
                      }}
                      className="text-[11px] font-bold text-emerald-700 hover:underline cursor-pointer"
                    >
                      Pilih Semua
                    </button>
                  </div>

                  <div className="flex flex-wrap gap-1.5">
                    {SASARAN_KATEGORI_LIST.map((sas) => {
                      const checked = selectedSasaranSet.has(sas.id);
                      return (
                        <button
                          key={sas.id}
                          type="button"
                          onClick={() => toggleSasaranId(sas.id)}
                          className={`px-2.5 py-1.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 transition cursor-pointer ${
                            checked
                              ? "bg-emerald-50 border-emerald-500 text-emerald-900"
                              : "bg-white border-slate-200 text-slate-500 hover:bg-slate-50"
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

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div className="space-y-1">
                  <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                    Judul Kategori pada Header Rekapan (Otomatis / Edit Bebas)
                  </label>
                  <input
                    type="text"
                    value={activeWs.config.title ?? ""}
                    onChange={(e) =>
                      updateWorksheetConfig(activeWs.config.id, (cfg) => ({
                        ...cfg,
                        title: e.target.value,
                      }))
                    }
                    placeholder={activeWs.autoCategoryTitle}
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
          </div>
        )}
      </div>

      {/* Inline Feedback Banner */}
      {photoBanner && (
        <div
          className={`p-3.5 rounded-xl border text-xs font-semibold flex items-center justify-between gap-3 no-print ${
            photoBanner.type === "error"
              ? "bg-rose-50 border-rose-200 text-rose-900"
              : photoBanner.type === "success"
              ? "bg-emerald-50 border-emerald-200 text-emerald-900"
              : "bg-sky-50 border-sky-200 text-sky-900"
          }`}
        >
          <div className="flex items-center gap-2">
            {photoBanner.type === "error" ? (
              <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
            ) : (
              <CheckCircle2
                className={`w-4 h-4 shrink-0 ${
                  photoBanner.type === "success" ? "text-emerald-600" : "text-sky-600"
                }`}
              />
            )}
            <span>{photoBanner.text}</span>
          </div>
          <button
            type="button"
            onClick={() => setPhotoBanner(null)}
            className="p-1 hover:bg-black/5 rounded-lg cursor-pointer"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* ================================================================= */}
      {/* SUB-VIEW A: IKSS NILAI GIZI                                       */}
      {/* ================================================================= */}
      {subView === "nilai_gizi" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-x-auto p-4 sm:p-6">
          <div
            id="print-area-ikss-gizi"
            className="w-full min-w-[780px] bg-white text-black"
            style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
          >
            {worksheetsComputed.map((ws, wsIdx) => {
              const showAkgInputsWs =
                (isTableEditMode || editingAkgWsId === ws.config.id) &&
                !isExportingPdf &&
                !showPrintModal;

              return (
                <div
                  key={ws.config.id}
                  className={
                    wsIdx > 0
                      ? "mt-10 pt-8 border-t-4 border-dashed border-slate-300 print:break-before-page print:border-t-0 print:pt-0 print:mt-0"
                      : ""
                  }
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

                  <div className="relative bg-[#00B050] border border-black py-2 px-3 text-center">
                    {showIdentityInputs ? (
                      <div className="print:hidden flex items-center justify-center gap-2 font-sans">
                        <span className="text-xs font-extrabold uppercase text-black shrink-0">
                          REKAPAN NILAI GIZI
                        </span>
                        <input
                          type="text"
                          value={ws.config.title ?? ""}
                          onChange={(e) =>
                            updateWorksheetConfig(ws.config.id, (cfg) => ({
                              ...cfg,
                              title: e.target.value,
                            }))
                          }
                          placeholder={ws.autoCategoryTitle}
                          className="w-full max-w-[380px] px-2 py-0.5 text-xs font-bold uppercase text-slate-900 bg-white/95 border border-black/40 rounded focus:outline-none focus:ring-2 focus:ring-white"
                        />
                      </div>
                    ) : null}
                    <h3
                      className={`text-sm font-bold uppercase text-black tracking-wide ${
                        showIdentityInputs ? "hidden print:block" : ""
                      }`}
                    >
                      REKAPAN NILAI GIZI {ws.displayCategoryTitle}
                    </h3>
                    <button
                      type="button"
                      onClick={() => setIsEditingIdentity((v) => !v)}
                      title="Edit Judul & Identitas SPPG"
                      className="no-print absolute right-2 top-1/2 -translate-y-1/2 p-1 rounded bg-white/85 hover:bg-white text-slate-800 border border-black/20 cursor-pointer"
                    >
                      <Pencil className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {renderIdentityBlock()}

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
                          Data Nilai Gizi {ws.displayCategoryTitle}
                        </th>
                        <th
                          rowSpan={2}
                          className="border border-black bg-[#00B050] text-black font-bold px-2 py-2 text-center w-[88px]"
                        >
                          Jumlah
                          <br />
                          Manfaat
                        </th>
                        <th
                          rowSpan={2}
                          className="no-print border border-black bg-[#8EA9DB] text-black font-bold px-1.5 py-2 text-center w-[68px] font-sans"
                        >
                          Edit
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
                      {ws.dailyRows.map((row) => {
                        const rowKey = `${ws.config.id}__${row.hariKe}`;
                        const isRowEditing = editingGiziKey === rowKey && !showInlineTableInputs;
                        const onRowKeyDown = (e: React.KeyboardEvent) => {
                          if (e.key === "Enter" && (e.target as HTMLElement).tagName !== "TEXTAREA") {
                            e.preventDefault();
                            saveEditGiziRow(ws, row.hariKe);
                          } else if (e.key === "Escape") {
                            setEditingGiziKey(null);
                            setGiziRowErrors({});
                          }
                        };

                        return (
                          <tr
                            key={row.hariKe}
                            className={
                              isRowEditing
                                ? "bg-amber-50/70 print:bg-transparent"
                                : row.isEdited
                                ? "bg-amber-50/25 print:bg-transparent"
                                : ""
                            }
                          >
                            {/* Hari / Tanggal */}
                            <td className="border border-black px-1.5 py-2 text-left align-middle font-normal whitespace-nowrap">
                              {showInlineTableInputs ? (
                                <>
                                  <input
                                    type="text"
                                    value={row.formattedDate}
                                    onChange={(e) =>
                                      handleBulkTextChange(ws, row.hariKe, "tanggal", e.target.value)
                                    }
                                    aria-label={`Hari / Tanggal Hari Ke-${row.hariKe}`}
                                    className="print:hidden w-full min-w-[138px] px-1.5 py-1 text-[11px] bg-white border border-amber-300 rounded focus:outline-none focus:border-emerald-600 font-sans"
                                  />
                                  <span className="hidden print:inline">{row.formattedDate}</span>
                                </>
                              ) : isRowEditing ? (
                                <input
                                  type="text"
                                  value={giziRowDraft.tanggal}
                                  onChange={(e) => updateGiziRowDraftField("tanggal", e.target.value)}
                                  onKeyDown={onRowKeyDown}
                                  className="w-full min-w-[138px] px-1.5 py-1 text-[11px] bg-white border border-emerald-500 rounded focus:outline-none focus:ring-2 focus:ring-emerald-500/20 font-sans"
                                />
                              ) : (
                                <span className="px-0.5">{row.formattedDate}</span>
                              )}
                            </td>

                            {/* Nama Menu */}
                            <td className="border border-black px-1.5 py-2 text-left align-middle leading-snug">
                              {showInlineTableInputs ? (
                                <>
                                  <textarea
                                    rows={2}
                                    value={row.namaMenu}
                                    onChange={(e) =>
                                      handleBulkTextChange(ws, row.hariKe, "namaMenu", e.target.value)
                                    }
                                    placeholder="Isi nama menu..."
                                    aria-label={`Nama Menu Hari Ke-${row.hariKe}`}
                                    className="print:hidden w-full min-w-[160px] px-1.5 py-1 text-[11px] bg-white border border-amber-300 rounded focus:outline-none focus:border-emerald-600 leading-snug resize-y font-sans"
                                  />
                                  <span className="hidden print:inline">{row.namaMenu || "-"}</span>
                                </>
                              ) : isRowEditing ? (
                                <textarea
                                  rows={2}
                                  value={giziRowDraft.namaMenu}
                                  onChange={(e) => updateGiziRowDraftField("namaMenu", e.target.value)}
                                  onKeyDown={onRowKeyDown}
                                  placeholder="Isi nama menu..."
                                  className="w-full min-w-[160px] px-1.5 py-1 text-[11px] bg-white border border-emerald-500 rounded focus:outline-none focus:ring-2 focus:ring-emerald-500/20 leading-snug resize-y font-sans"
                                />
                              ) : (
                                <span className="px-0.5">{row.namaMenu || "-"}</span>
                              )}
                            </td>

                            {/* 5 Nutrition Columns: Energi, Protein, Lemak, Karbohidrat, Serat */}
                            {(["energi", "protein", "lemak", "kh", "serat"] as const).map((field) => {
                              const cellKey = `${ws.config.id}:${row.hariKe}:${field}`;
                              const isPopulated = row.hasNutrition || row.editedFields[field];
                              const bulkVal =
                                cellKey in bulkCellDrafts
                                  ? bulkCellDrafts[cellKey]
                                  : isPopulated
                                  ? formatDec1(row[field])
                                  : "";
                              const displayVal = isPopulated ? formatDec1(row[field]) : "-";
                              const bulkErr = bulkCellErrors[cellKey];
                              const rowErr = isRowEditing ? giziRowErrors[field] : undefined;

                              return (
                                <td
                                  key={field}
                                  className={`border border-black px-1 py-2 text-center align-middle tabular-nums ${
                                    row.editedFields[field] && !showInlineTableInputs && !isRowEditing
                                      ? "bg-amber-50/70 print:bg-transparent font-semibold"
                                      : ""
                                  }`}
                                >
                                  {showInlineTableInputs ? (
                                    <>
                                      <input
                                        type="text"
                                        inputMode="decimal"
                                        value={bulkVal}
                                        placeholder="-"
                                        onChange={(e) =>
                                          handleBulkNumericChange(ws, row.hariKe, field, e.target.value)
                                        }
                                        onBlur={() => handleBulkNumericBlur(ws.config.id, row.hariKe, field)}
                                        aria-label={`${NUTRIENT_FIELD_LABELS[field]} Hari Ke-${row.hariKe}`}
                                        title={bulkErr || NUTRIENT_FIELD_LABELS[field]}
                                        className={`print:hidden w-full min-w-[54px] px-1 py-1 text-center text-[11px] tabular-nums rounded border focus:outline-none font-sans ${
                                          bulkErr
                                            ? "bg-rose-50 border-rose-500 text-rose-900 ring-1 ring-rose-500"
                                            : row.editedFields[field]
                                            ? "bg-amber-50 border-amber-400 text-slate-900 font-bold focus:bg-white focus:border-emerald-600"
                                            : "bg-white border-amber-300 text-slate-900 focus:border-emerald-600"
                                        }`}
                                      />
                                      <span className="hidden print:inline">{displayVal}</span>
                                    </>
                                  ) : isRowEditing ? (
                                    <input
                                      type="text"
                                      inputMode="decimal"
                                      value={giziRowDraft[field]}
                                      placeholder="-"
                                      onChange={(e) => updateGiziRowDraftField(field, e.target.value)}
                                      onKeyDown={onRowKeyDown}
                                      title={rowErr || NUTRIENT_FIELD_LABELS[field]}
                                      className={`w-full min-w-[54px] px-1 py-1 text-center text-[11px] tabular-nums rounded border focus:outline-none font-sans ${
                                        rowErr
                                          ? "bg-rose-50 border-rose-500 text-rose-900 ring-1 ring-rose-500"
                                          : "bg-white border-emerald-500 text-slate-900 focus:ring-2 focus:ring-emerald-500/20"
                                      }`}
                                    />
                                  ) : (
                                    <span className={isPopulated ? "" : "text-slate-400 print:text-black"}>
                                      {displayVal}
                                    </span>
                                  )}
                                </td>
                              );
                            })}

                            {/* Jumlah Manfaat (PM) */}
                            {(() => {
                              const pmKey = `${ws.config.id}:${row.hariKe}:jumlahManfaat`;
                              const bulkPmVal =
                                pmKey in bulkCellDrafts
                                  ? bulkCellDrafts[pmKey]
                                  : String(row.jumlahManfaat);
                              const bulkPmErr = bulkCellErrors[pmKey];
                              const rowPmErr = isRowEditing ? giziRowErrors.jumlahManfaat : undefined;

                              return (
                                <td
                                  className={`border border-black px-1 py-2 text-center align-middle tabular-nums ${
                                    row.editedFields.jumlahManfaat && !showInlineTableInputs && !isRowEditing
                                      ? "bg-amber-50/70 print:bg-transparent font-bold"
                                      : ""
                                  }`}
                                >
                                  {showInlineTableInputs ? (
                                    <>
                                      <input
                                        type="text"
                                        inputMode="numeric"
                                        value={bulkPmVal}
                                        onChange={(e) =>
                                          handleBulkNumericChange(
                                            ws,
                                            row.hariKe,
                                            "jumlahManfaat",
                                            e.target.value
                                          )
                                        }
                                        onBlur={() =>
                                          handleBulkNumericBlur(ws.config.id, row.hariKe, "jumlahManfaat")
                                        }
                                        aria-label={`Jumlah PM Hari Ke-${row.hariKe}`}
                                        title={bulkPmErr || "Jumlah Penerima Manfaat (PM)"}
                                        className={`print:hidden w-full min-w-[60px] px-1 py-1 text-center text-[11px] tabular-nums rounded border focus:outline-none font-sans ${
                                          bulkPmErr
                                            ? "bg-rose-50 border-rose-500 text-rose-900 ring-1 ring-rose-500"
                                            : row.editedFields.jumlahManfaat
                                            ? "bg-amber-50 border-amber-400 text-slate-900 font-bold focus:bg-white focus:border-emerald-600"
                                            : "bg-white border-amber-300 text-slate-900 focus:border-emerald-600"
                                        }`}
                                      />
                                      <span className="hidden print:inline">{row.jumlahManfaat}</span>
                                    </>
                                  ) : isRowEditing ? (
                                    <input
                                      type="text"
                                      inputMode="numeric"
                                      value={giziRowDraft.jumlahManfaat}
                                      onChange={(e) =>
                                        updateGiziRowDraftField("jumlahManfaat", e.target.value)
                                      }
                                      onKeyDown={onRowKeyDown}
                                      title={rowPmErr || "Jumlah Penerima Manfaat (PM)"}
                                      className={`w-full min-w-[60px] px-1 py-1 text-center text-[11px] tabular-nums rounded border focus:outline-none font-sans ${
                                        rowPmErr
                                          ? "bg-rose-50 border-rose-500 text-rose-900 ring-1 ring-rose-500"
                                          : "bg-white border-emerald-500 text-slate-900 focus:ring-2 focus:ring-emerald-500/20"
                                      }`}
                                    />
                                  ) : (
                                    row.jumlahManfaat
                                  )}
                                </td>
                              );
                            })()}

                            {/* Per-row Pencil Edit Action Column (Hidden on Print / PDF) */}
                            <td className="no-print border border-black px-1 py-1.5 text-center align-middle font-sans bg-slate-50/60">
                              {showInlineTableInputs ? (
                                <div className="flex items-center justify-center gap-1">
                                  {row.isEdited ? (
                                    <button
                                      type="button"
                                      onClick={() => resetDayOverride(ws, row.hariKe)}
                                      title="Reset baris ini ke perhitungan otomatis"
                                      className="p-1 rounded bg-white hover:bg-rose-50 border border-slate-200 hover:border-rose-300 text-slate-600 hover:text-rose-700 cursor-pointer"
                                    >
                                      <RotateCcw className="w-3.5 h-3.5" />
                                    </button>
                                  ) : (
                                    <span className="text-[10px] font-semibold text-slate-400">Auto</span>
                                  )}
                                </div>
                              ) : isRowEditing ? (
                                <div className="flex items-center justify-center gap-1">
                                  <button
                                    type="button"
                                    onClick={() => saveEditGiziRow(ws, row.hariKe)}
                                    disabled={Object.keys(giziRowErrors).length > 0}
                                    title="Simpan Perubahan Baris"
                                    className="p-1 rounded bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white cursor-pointer"
                                  >
                                    <Check className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setEditingGiziKey(null);
                                      setGiziRowErrors({});
                                    }}
                                    title="Batal Edit"
                                    className="p-1 rounded bg-slate-200 hover:bg-slate-300 text-slate-700 cursor-pointer"
                                  >
                                    <X className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              ) : (
                                <div className="flex items-center justify-center gap-1">
                                  <button
                                    type="button"
                                    onClick={() => startEditGiziRow(ws, row.hariKe)}
                                    title="Edit Nilai Gizi, Menu, Tanggal & Jumlah PM"
                                    className="p-1 rounded bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-800 cursor-pointer transition"
                                  >
                                    <Pencil className="w-3.5 h-3.5" />
                                  </button>
                                  {row.isEdited && (
                                    <button
                                      type="button"
                                      onClick={() => resetDayOverride(ws, row.hariKe)}
                                      title="Kembalikan baris ini ke nilai otomatis"
                                      className="p-1 rounded bg-white hover:bg-rose-50 border border-slate-200 hover:border-rose-300 text-slate-500 hover:text-rose-700 cursor-pointer transition"
                                    >
                                      <RotateCcw className="w-3 h-3" />
                                    </button>
                                  )}
                                </div>
                              )}
                            </td>
                          </tr>
                        );
                      })}

                      <tr>
                        <td
                          colSpan={2}
                          className="border border-black bg-[#00B050] text-black font-bold px-2 py-1.5 text-left"
                        >
                          Total Nilai Gizi
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.totalEnergi) : "-"}
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.totalProtein) : "-"}
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.totalLemak) : "-"}
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.totalKh) : "-"}
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.totalSerat) : "-"}
                        </td>
                        <td className="border border-black bg-white text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.totalManfaat}
                        </td>
                        <td className="no-print border border-black bg-[#00B050]/40 px-1 py-1 text-center text-[10px] font-semibold text-slate-700 font-sans">
                          Auto
                        </td>
                      </tr>

                      <tr>
                        <td
                          colSpan={2}
                          className="border border-black bg-[#00B050] text-black font-bold px-2 py-1.5 text-left"
                        >
                          Jumlah Hari
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays} Hari
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays} Hari
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays} Hari
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays} Hari
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays} Hari
                        </td>
                        <td className="border border-black bg-[#00B050] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activePmDays} Hari
                        </td>
                        <td className="no-print border border-black bg-[#00B050]/40 px-1 py-1 text-center text-[10px] font-semibold text-slate-700 font-sans">
                          Auto
                        </td>
                      </tr>

                      <tr>
                        <td
                          colSpan={2}
                          className="border border-black bg-[#FFFF00] text-black font-bold px-2 py-1.5 text-left"
                        >
                          Nilai Rata-Rata
                        </td>
                        <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgEnergi) : "-"}
                        </td>
                        <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgProtein) : "-"}
                        </td>
                        <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgLemak) : "-"}
                        </td>
                        <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgKh) : "-"}
                        </td>
                        <td className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center tabular-nums">
                          {ws.summary.activeNutriDays > 0 ? formatDec1(ws.summary.avgSerat) : "-"}
                        </td>
                        <td
                          rowSpan={3}
                          className="border border-black bg-[#FFFF00] text-black font-bold px-1.5 py-1.5 text-center align-middle tabular-nums"
                        >
                          {ws.summary.avgManfaat} Orang
                        </td>
                        <td className="no-print border border-black bg-[#FFFF00]/60 px-1 py-1 text-center text-[10px] font-semibold text-slate-700 font-sans">
                          Auto
                        </td>
                      </tr>

                      <tr>
                        <td
                          colSpan={2}
                          className="border border-black bg-[#00B050] text-black font-bold px-2 py-1.5 text-left"
                        >
                          {showAkgInputsWs ? (
                            <>
                              <input
                                type="text"
                                value={ws.effectiveAkg.label}
                                onChange={(e) =>
                                  updateWorksheetConfig(ws.config.id, (cfg) => ({
                                    ...cfg,
                                    akgOverride: {
                                      ...(cfg.akgOverride || {}),
                                      label: e.target.value,
                                    },
                                  }))
                                }
                                aria-label="Label Rujukan AKG"
                                className="print:hidden w-full max-w-[280px] px-1.5 py-0.5 text-[11px] font-bold text-slate-900 bg-white/95 border border-black/40 rounded focus:outline-none font-sans"
                              />
                              <span className="hidden print:inline">{ws.effectiveAkg.label}</span>
                            </>
                          ) : (
                            ws.effectiveAkg.label
                          )}
                        </td>
                        {(["energi", "protein", "lemak", "kh", "serat"] as const).map(
                          (field) => {
                            const akgKey = `${ws.config.id}:akg:${field}`;
                            const akgDraftVal =
                              akgKey in bulkCellDrafts
                                ? bulkCellDrafts[akgKey]
                                : String(ws.effectiveAkg[field]);
                            const akgErr = bulkCellErrors[akgKey];

                            return (
                              <td
                                key={field}
                                className="border border-black bg-[#00B050] text-black font-bold px-1 py-1 text-center tabular-nums"
                              >
                                <input
                                  type="text"
                                  inputMode="decimal"
                                  value={akgDraftVal}
                                  onChange={(e) =>
                                    handleAkgNumericChange(ws.config.id, field, e.target.value)
                                  }
                                  onBlur={() => handleAkgNumericBlur(ws.config.id, field)}
                                  aria-label={`Rujukan AKG ${NUTRIENT_FIELD_LABELS[field]}`}
                                  title={akgErr || `Edit Rujukan AKG ${NUTRIENT_FIELD_LABELS[field]}`}
                                  className={`print:hidden w-full text-center font-bold rounded tabular-nums px-1 py-0.5 focus:outline-none ${
                                    akgErr
                                      ? "bg-rose-100 text-rose-900 ring-1 ring-rose-600"
                                      : showAkgInputsWs
                                      ? "bg-white/95 text-slate-900 border border-black/30"
                                      : "bg-transparent text-black hover:bg-white/25 focus:bg-white/90"
                                  }`}
                                />
                                <span className="hidden print:inline">{ws.effectiveAkg[field]}</span>
                              </td>
                            );
                          }
                        )}
                        <td className="no-print border border-black bg-[#00B050]/40 px-1 py-1 text-center align-middle font-sans">
                          <div className="flex items-center justify-center gap-1">
                            <button
                              type="button"
                              onClick={() =>
                                setEditingAkgWsId((prev) =>
                                  prev === ws.config.id ? null : ws.config.id
                                )
                              }
                              title={
                                showAkgInputsWs
                                  ? "Selesai Edit AKG"
                                  : "Edit Label & Nilai Rujukan AKG"
                              }
                              className={`p-1 rounded cursor-pointer transition ${
                                showAkgInputsWs
                                  ? "bg-emerald-700 text-white"
                                  : "bg-white/90 hover:bg-white border border-black/25 text-slate-800"
                              }`}
                            >
                              {showAkgInputsWs ? (
                                <Check className="w-3 h-3" />
                              ) : (
                                <Pencil className="w-3 h-3" />
                              )}
                            </button>
                            {ws.manualEditStats.hasAkg && (
                              <button
                                type="button"
                                onClick={() => resetAkgOverride(ws.config.id)}
                                title="Reset Rujukan AKG ke otomatis"
                                className="p-1 rounded bg-white/90 hover:bg-rose-50 border border-black/20 text-slate-700 hover:text-rose-700 cursor-pointer"
                              >
                                <RotateCcw className="w-3 h-3" />
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>

                      <tr>
                        <td
                          colSpan={2}
                          className="border border-black bg-[#00B050] text-black font-bold px-2 py-1.5 text-left"
                        >
                          % Kebutuhan 90-110%
                        </td>
                        {renderPctCell(ws.summary.pctEnergi, ws.summary.activeNutriDays > 0)}
                        {renderPctCell(ws.summary.pctProtein, ws.summary.activeNutriDays > 0)}
                        {renderPctCell(ws.summary.pctLemak, ws.summary.activeNutriDays > 0)}
                        {renderPctCell(ws.summary.pctKh, ws.summary.activeNutriDays > 0)}
                        {renderPctCell(ws.summary.pctSerat, ws.summary.activeNutriDays > 0)}
                        <td className="no-print border border-black bg-[#00B050]/40 px-1 py-1 text-center text-[10px] font-semibold text-slate-700 font-sans">
                          Auto
                        </td>
                      </tr>
                    </tbody>
                  </table>

                  <AutoSignatureBlock profile={profile} docKey="ikss" />
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ================================================================= */}
      {/* SUB-VIEW B: IKSS FOTO MENU MAKANAN (UPLOAD & POSTING HARIAN)      */}
      {/* ================================================================= */}
      {subView === "foto_menu" && (
        <div className="space-y-4">
          {/* DAILY PHOTO UPLOAD & POSTING CONTROL PANEL */}
          <div className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-5 shadow-xs space-y-4 no-print">
            <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-4">
              {/* Progress & Stats */}
              <div className="space-y-2 flex-1 min-w-[240px]">
                <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs">
                  <span className="font-extrabold text-slate-900 text-sm">
                    Posting Foto Menu Harian
                    {worksheetsComputed.length > 1 && (
                      <span className="ml-2 text-xs font-bold text-indigo-700 bg-indigo-50 border border-indigo-200 px-2 py-0.5 rounded-lg">
                        Fokus: {activeWs.displayCategoryTitle}
                      </span>
                    )}
                  </span>
                  <span className="text-slate-600 font-semibold tabular-nums">
                    Terunggah: <strong className="text-slate-900">{photoStats.uploaded}</strong>/{photoStats.total} Hari
                  </span>
                  <span className="text-emerald-700 font-semibold tabular-nums">
                    Diposting: <strong>{photoStats.posted}</strong> Hari
                  </span>
                  {photoStats.draft > 0 && (
                    <span className="text-amber-700 font-semibold tabular-nums">
                      Draft: <strong>{photoStats.draft}</strong>
                    </span>
                  )}
                  <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-slate-500">
                    {syncState === "saving" ? (
                      <>
                        <Loader2 className="w-3 h-3 animate-spin text-indigo-600" />
                        Menyimpan...
                      </>
                    ) : syncState === "synced" || (isCloudActive && photosHydrated && syncState === "idle") ? (
                      <>
                        <Cloud className="w-3.5 h-3.5 text-emerald-600" />
                        Sinkron Cloud Aktif
                      </>
                    ) : (
                      <>
                        <CloudOff className="w-3.5 h-3.5 text-slate-400" />
                        Tersimpan Lokal
                      </>
                    )}
                  </span>
                </div>

                {/* Progress Bar */}
                <div className="w-full max-w-md h-2 bg-slate-100 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-emerald-600 transition-all duration-300"
                    style={{ width: `${photoStats.pct}%` }}
                  />
                </div>
              </div>

              {/* Bulk Upload + Post All Primary Controls */}
              <div className="flex flex-wrap items-center gap-2 shrink-0">
                <input
                  ref={bulkInputRef}
                  type="file"
                  accept="image/*"
                  multiple
                  onChange={(e) => void handleBulkUploadChange(activeWs, e)}
                  className="hidden"
                />

                <button
                  type="button"
                  onClick={() => bulkInputRef.current?.click()}
                  disabled={Boolean(bulkProgress)}
                  className="px-3.5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-800 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer disabled:opacity-50 whitespace-nowrap"
                >
                  {bulkProgress ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-600" />
                  ) : (
                    <Upload className="w-3.5 h-3.5 text-emerald-600" />
                  )}
                  {bulkProgress
                    ? `Memproses ${bulkProgress.current}/${bulkProgress.total}...`
                    : "Unggah Banyak Foto"}
                </button>

                <button
                  type="button"
                  onClick={() => void handlePostAllPhotos(activeWs)}
                  disabled={photoStats.uploaded === 0 || syncState === "saving"}
                  className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-xs font-bold inline-flex items-center gap-1.5 shadow-xs cursor-pointer whitespace-nowrap"
                >
                  {syncState === "saving" ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Send className="w-3.5 h-3.5" />
                  )}
                  Posting Semua Foto ({photoStats.uploaded})
                </button>
              </div>
            </div>

            {/* Secondary Row: Filter, Auto-Post Toggle, Layout Mode */}
            <div className="flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-slate-100">
              <div className="flex flex-wrap items-center gap-1.5">
                {(
                  [
                    { id: "semua", label: `Semua (${photoStats.total})` },
                    { id: "diposting", label: `Diposting (${photoStats.posted})` },
                    { id: "draft", label: `Draft (${photoStats.draft})` },
                    { id: "belum", label: `Belum Foto (${photoStats.empty})` },
                  ] as const
                ).map((item) => (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => setPhotoFilter(item.id)}
                    className={`px-2.5 py-1 rounded-lg text-xs font-semibold transition cursor-pointer ${
                      photoFilter === item.id
                        ? "bg-slate-900 text-white"
                        : "bg-slate-100 text-slate-600 hover:bg-slate-200/70"
                    }`}
                  >
                    {item.label}
                  </button>
                ))}
              </div>

              <div className="flex flex-wrap items-center gap-3">
                <label className="inline-flex items-center gap-2 text-xs font-semibold text-slate-700 cursor-pointer select-none whitespace-nowrap">
                  <input
                    type="checkbox"
                    checked={Boolean(ikssData.autoPostOnUpload)}
                    onChange={(e) =>
                      setIkssData((prev) => ({
                        ...prev,
                        autoPostOnUpload: e.target.checked,
                      }))
                    }
                    className="w-4 h-4 accent-emerald-600 rounded cursor-pointer"
                  />
                  <span>Auto-Posting saat Unggah</span>
                </label>

                <div className="flex bg-slate-100 p-0.5 rounded-lg border border-slate-200">
                  <button
                    type="button"
                    onClick={() => setPhotoViewMode("tabel")}
                    className={`px-2.5 py-1 rounded-md text-xs font-bold inline-flex items-center gap-1 cursor-pointer transition ${
                      photoViewMode === "tabel"
                        ? "bg-white text-slate-900 shadow-xs"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    <Table2 className="w-3.5 h-3.5" />
                    Tabel Resmi
                  </button>
                  <button
                    type="button"
                    onClick={() => setPhotoViewMode("kartu")}
                    className={`px-2.5 py-1 rounded-md text-xs font-bold inline-flex items-center gap-1 cursor-pointer transition ${
                      photoViewMode === "kartu"
                        ? "bg-white text-slate-900 shadow-xs"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    <LayoutGrid className="w-3.5 h-3.5" />
                    Kartu Posting
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* INTERACTIVE CARD POSTING VIEW (OPTIMAL FOR MOBILE CAMERA & RAPID DAILY POSTING) */}
          {photoViewMode === "kartu" && (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4 no-print">
              {filteredPhotoRows.map((row) => {
                const photoSrc = getPhotoForWsDay(activeWs.photoGroupKey, row.hariKe);
                const postMeta = getPostMetaForWsDay(activeWs.photoGroupKey, row.hariKe);
                const isPosted = Boolean(photoSrc && postMeta?.posted);
                const isBusy = Boolean(processingDays[row.hariKe]);
                const isDragging = dragOverDay === row.hariKe;
                const rowKey = `${activeWs.config.id}__${row.hariKe}`;
                const isEditing = editingMenuKey === rowKey;

                return (
                  <div
                    key={rowKey}
                    onDragOver={(e) => {
                      e.preventDefault();
                      setDragOverDay(row.hariKe);
                    }}
                    onDragLeave={() => setDragOverDay(null)}
                    onDrop={(e) => handleDropOnDay(activeWs.photoGroupKey, row.hariKe, e)}
                    className={`bg-white rounded-2xl border transition flex flex-col overflow-hidden ${
                      isDragging
                        ? "border-emerald-500 ring-2 ring-emerald-500/30"
                        : isPosted
                        ? "border-emerald-200"
                        : "border-slate-200"
                    }`}
                  >
                    {/* Card Header */}
                    <div className="px-3.5 py-2.5 border-b border-slate-100 flex items-center justify-between gap-2 bg-slate-50/70">
                      <div className="min-w-0">
                        <div className="text-xs font-extrabold text-slate-900 tabular-nums">
                          Hari Ke-{row.hariKe}
                        </div>
                        <div className="text-[11px] text-slate-500 truncate">
                          {row.formattedDate}
                        </div>
                      </div>
                      <span
                        className={`text-[11px] font-bold shrink-0 ${
                          isPosted
                            ? "text-emerald-700"
                            : photoSrc
                            ? "text-amber-700"
                            : "text-slate-400"
                        }`}
                      >
                        {isPosted ? "Diposting" : photoSrc ? "Draft" : "Belum Foto"}
                      </span>
                    </div>

                    {/* Photo Stage */}
                    <div className="p-3 space-y-2.5 flex-1 flex flex-col">
                      <div className="relative w-full aspect-[4/3] rounded-xl overflow-hidden border border-slate-200 bg-slate-50 flex items-center justify-center">
                        {isBusy ? (
                          <div className="flex flex-col items-center gap-1.5 text-emerald-700">
                            <Loader2 className="w-6 h-6 animate-spin" />
                            <span className="text-[11px] font-bold">Memproses foto...</span>
                          </div>
                        ) : photoSrc ? (
                          <>
                            <img
                              src={photoSrc}
                              alt={`Foto Menu Hari Ke-${row.hariKe}`}
                              onClick={() =>
                                setLightboxTarget({
                                  wsId: activeWs.config.id,
                                  hariKe: row.hariKe,
                                })
                              }
                              className="w-full h-full object-cover cursor-zoom-in"
                              loading="lazy"
                              decoding="async"
                            />
                            <div className="absolute top-1.5 right-1.5 flex items-center gap-1">
                              <button
                                type="button"
                                onClick={() =>
                                  setLightboxTarget({
                                    wsId: activeWs.config.id,
                                    hariKe: row.hariKe,
                                  })
                                }
                                title="Perbesar Foto"
                                className="p-1.5 rounded-lg bg-slate-900/75 hover:bg-slate-900 text-white cursor-pointer"
                              >
                                <ZoomIn className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() =>
                                  void handleRotateDayPhoto(activeWs.photoGroupKey, row.hariKe)
                                }
                                title="Putar 90°"
                                className="p-1.5 rounded-lg bg-slate-900/75 hover:bg-slate-900 text-white cursor-pointer"
                              >
                                <RotateCw className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() =>
                                  void handleDeletePhoto(activeWs.photoGroupKey, row.hariKe)
                                }
                                title="Hapus Foto"
                                className="p-1.5 rounded-lg bg-rose-600/90 hover:bg-rose-700 text-white cursor-pointer"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </>
                        ) : (
                          <div className="text-center px-3 space-y-1">
                            <Camera className="w-6 h-6 text-slate-300 mx-auto" />
                            <p className="text-[11px] text-slate-400 font-medium">
                              Ambil kamera, pilih galeri, atau seret foto ke sini
                            </p>
                          </div>
                        )}
                      </div>

                      {/* Date, Menu Name & Caption */}
                      {isEditing ? (
                        <div className="space-y-1.5">
                          <input
                            type="text"
                            value={editDateDraft}
                            onChange={(e) => setEditDateDraft(e.target.value)}
                            placeholder="Hari / Tanggal..."
                            className="w-full px-2.5 py-1 text-xs border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500"
                          />
                          <textarea
                            rows={2}
                            value={editMenuDraft}
                            onChange={(e) => setEditMenuDraft(e.target.value)}
                            placeholder="Nama Menu Harian..."
                            className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500"
                          />
                          <input
                            type="text"
                            value={editCaptionDraft}
                            onChange={(e) => setEditCaptionDraft(e.target.value)}
                            placeholder="Catatan penyajian (opsional)"
                            className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500"
                          />
                          <div className="flex items-center justify-between gap-1.5">
                            {row.isEdited ? (
                              <button
                                type="button"
                                onClick={() => resetDayOverride(activeWs, row.hariKe)}
                                className="px-2 py-1 text-[11px] font-semibold text-rose-600 hover:bg-rose-50 rounded-md cursor-pointer inline-flex items-center gap-1"
                              >
                                <RotateCcw className="w-3 h-3" /> Reset
                              </button>
                            ) : (
                              <span />
                            )}
                            <div className="flex items-center gap-1.5">
                              <button
                                type="button"
                                onClick={() => setEditingMenuKey(null)}
                                className="px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100 rounded-md cursor-pointer"
                              >
                                Batal
                              </button>
                              <button
                                type="button"
                                onClick={() => saveEditDayMenu(activeWs, row.hariKe)}
                                className="px-2.5 py-1 text-[11px] font-bold bg-emerald-600 text-white rounded-md cursor-pointer"
                              >
                                Simpan
                              </button>
                            </div>
                          </div>
                        </div>
                      ) : (
                        <div className="flex items-start justify-between gap-1.5 min-h-[38px]">
                          <div className="min-w-0">
                            <p className="text-xs font-bold text-slate-800 leading-snug line-clamp-2">
                              {row.namaMenu || "Menu belum diisi"}
                            </p>
                            {row.caption && (
                              <p className="text-[11px] text-slate-500 mt-0.5 line-clamp-1">
                                {row.caption}
                              </p>
                            )}
                          </div>
                          <button
                            type="button"
                            onClick={() => startEditDayMenu(activeWs, row.hariKe)}
                            title="Edit Tanggal, Nama Menu & Catatan"
                            className="p-1.5 bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-800 rounded-lg cursor-pointer shrink-0 transition"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      )}

                      {/* Camera + Gallery + Post Action */}
                      <div className="pt-2 border-t border-slate-100 mt-auto space-y-2">
                        <div className="grid grid-cols-2 gap-1.5">
                          <label className="inline-flex items-center justify-center gap-1 px-2.5 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 text-xs font-bold cursor-pointer transition">
                            <Camera className="w-3.5 h-3.5" />
                            <span>Kamera</span>
                            <input
                              type="file"
                              accept="image/*"
                              capture="environment"
                              onChange={(e) =>
                                handlePhotoInputChange(activeWs.photoGroupKey, row.hariKe, e)
                              }
                              className="hidden"
                            />
                          </label>
                          <label className="inline-flex items-center justify-center gap-1 px-2.5 py-1.5 rounded-lg bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-800 text-xs font-bold cursor-pointer transition">
                            <ImageIcon className="w-3.5 h-3.5" />
                            <span>Galeri</span>
                            <input
                              type="file"
                              accept="image/*"
                              onChange={(e) =>
                                handlePhotoInputChange(activeWs.photoGroupKey, row.hariKe, e)
                              }
                              className="hidden"
                            />
                          </label>
                        </div>

                        {photoSrc && !isPosted && (
                          <button
                            type="button"
                            onClick={() =>
                              void handlePostDayPhoto(activeWs.photoGroupKey, row.hariKe)
                            }
                            className="w-full py-1.5 px-3 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold inline-flex items-center justify-center gap-1.5 cursor-pointer transition"
                          >
                            <Send className="w-3.5 h-3.5" />
                            Posting Foto Hari Ke-{row.hariKe}
                          </button>
                        )}
                        {isPosted && (
                          <div className="flex items-center justify-between text-[11px] text-emerald-700 font-semibold px-1">
                            <span className="inline-flex items-center gap-1">
                              <CheckCircle2 className="w-3.5 h-3.5" />
                              Diposting {formatShortTimestamp(postMeta?.postedAt)}
                            </span>
                            <button
                              type="button"
                              onClick={() =>
                                handleUnpostDayPhoto(activeWs.photoGroupKey, row.hariKe)
                              }
                              className="text-slate-400 hover:text-slate-700 underline cursor-pointer"
                            >
                              Draft
                            </button>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                );
              })}

              {filteredPhotoRows.length === 0 && (
                <div className="col-span-full bg-white rounded-2xl border border-slate-200 p-8 text-center text-xs text-slate-500">
                  Tidak ada hari yang cocok dengan filter status foto yang dipilih.
                </div>
              )}
            </div>
          )}

          {/* OFFICIAL PRINTABLE IKSS FOTO MENU TABLES (MULTI-WORKSHEET BY KELOMPOK SASARAN) */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-x-auto p-4 sm:p-6">
            {photoViewMode === "kartu" && (
              <div className="mb-3 pb-2 border-b border-slate-100 flex items-center justify-between no-print">
                <span className="text-xs font-bold text-slate-700">
                  Pratinjau Lembar Resmi IKSS Foto Menu ({worksheetsComputed.length} Kelompok Sasaran — Siap Cetak / PDF)
                </span>
                <span className="text-[11px] text-slate-500">
                  Seluruh foto yang diunggah &amp; diposting tampil otomatis di tabel ini
                </span>
              </div>
            )}

            <div
              id="print-area-ikss-foto"
              className="w-full min-w-[780px] bg-white text-black"
              style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
            >
              {worksheetsComputed.map((ws, wsIdx) => {
                const wsFilteredRows =
                  photoViewMode === "tabel" &&
                  !isExportingPdf &&
                  !showPrintModal &&
                  photoFilter !== "semua"
                    ? ws.dailyRows.filter((r) => {
                        const hasPhoto = Boolean(getPhotoForWsDay(ws.photoGroupKey, r.hariKe));
                        const meta = getPostMetaForWsDay(ws.photoGroupKey, r.hariKe);
                        if (photoFilter === "diposting") return hasPhoto && Boolean(meta?.posted);
                        if (photoFilter === "draft") return hasPhoto && !meta?.posted;
                        return !hasPhoto;
                      })
                    : ws.dailyRows;

                return (
                  <div
                    key={ws.config.id}
                    className={
                      wsIdx > 0
                        ? "mt-10 pt-8 border-t-4 border-dashed border-slate-300 print:break-before-page print:border-t-0 print:pt-0 print:mt-0"
                        : ""
                    }
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

                    <div className="relative bg-[#00B050] border border-black py-2 px-3 text-center">
                      {showIdentityInputs ? (
                        <div className="print:hidden flex items-center justify-center gap-2 font-sans">
                          <span className="text-xs font-extrabold uppercase text-black shrink-0">
                            REKAPAN MENU
                          </span>
                          <input
                            type="text"
                            value={ws.config.title ?? ""}
                            onChange={(e) =>
                              updateWorksheetConfig(ws.config.id, (cfg) => ({
                                ...cfg,
                                title: e.target.value,
                              }))
                            }
                            placeholder={ws.autoCategoryTitle}
                            className="w-full max-w-[380px] px-2 py-0.5 text-xs font-bold uppercase text-slate-900 bg-white/95 border border-black/40 rounded focus:outline-none focus:ring-2 focus:ring-white"
                          />
                        </div>
                      ) : null}
                      <h3
                        className={`text-sm font-bold uppercase text-black tracking-wide ${
                          showIdentityInputs ? "hidden print:block" : ""
                        }`}
                      >
                        REKAPAN MENU {ws.displayCategoryTitle}
                      </h3>
                      <button
                        type="button"
                        onClick={() => setIsEditingIdentity((v) => !v)}
                        title="Edit Judul & Identitas SPPG"
                        className="no-print absolute right-2 top-1/2 -translate-y-1/2 p-1 rounded bg-white/85 hover:bg-white text-slate-800 border border-black/20 cursor-pointer"
                      >
                        <Pencil className="w-3.5 h-3.5" />
                      </button>
                    </div>

                    {renderIdentityBlock()}

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
                            className="border border-black bg-[#8EA9DB] text-black font-bold px-2 py-2 text-center w-[210px]"
                          >
                            FOTO MENU
                          </th>
                          <th
                            rowSpan={2}
                            className="no-print border border-black bg-[#8EA9DB] text-black font-bold px-1.5 py-2 text-center w-[68px] font-sans"
                          >
                            Edit
                          </th>
                        </tr>
                        <tr>
                          <th className="border border-black bg-[#E54D3D] text-black font-bold px-2 py-1 text-center w-[190px]">
                            Nama SPPG
                          </th>
                          <th className="border border-black bg-[#E54D3D] text-black font-bold px-2 py-1 text-center w-[88px]">
                            ID SPPG
                          </th>
                        </tr>
                      </thead>

                      <tbody>
                        {wsFilteredRows.map((row) => {
                          const photoSrc = getPhotoForWsDay(ws.photoGroupKey, row.hariKe);
                          const postMeta = getPostMetaForWsDay(ws.photoGroupKey, row.hariKe);
                          const isPosted = Boolean(photoSrc && postMeta?.posted);
                          const isBusy = Boolean(processingDays[row.hariKe]);
                          const isDragging = dragOverDay === row.hariKe;
                          const rowKey = `${ws.config.id}__${row.hariKe}`;
                          const isEditing = editingMenuKey === rowKey && !showInlineTableInputs;

                          return (
                            <tr
                              key={rowKey}
                              className={
                                isEditing
                                  ? "bg-amber-50/70 print:bg-transparent"
                                  : row.isEdited
                                  ? "bg-amber-50/20 print:bg-transparent"
                                  : ""
                              }
                            >
                              <td className="border border-black px-1.5 py-2 text-center align-middle font-normal tabular-nums">
                                {row.hariKe}
                              </td>
                              <td className="border border-black px-1.5 py-2 text-left align-middle uppercase">
                                {showInlineTableInputs ? (
                                  <>
                                    <input
                                      type="text"
                                      value={row.namaSppg}
                                      onChange={(e) =>
                                        handleBulkTextChange(ws, row.hariKe, "namaSppg", e.target.value)
                                      }
                                      aria-label={`Nama SPPG Hari Ke-${row.hariKe}`}
                                      className="print:hidden w-full min-w-[140px] px-1.5 py-1 text-[11px] uppercase bg-white border border-amber-300 rounded focus:outline-none focus:border-emerald-600 font-sans"
                                    />
                                    <span className="hidden print:inline">{row.namaSppg || "-"}</span>
                                  </>
                                ) : isEditing ? (
                                  <input
                                    type="text"
                                    value={editSppgDraft}
                                    onChange={(e) => setEditSppgDraft(e.target.value)}
                                    className="w-full min-w-[140px] px-1.5 py-1 text-[11px] uppercase bg-white border border-emerald-500 rounded focus:outline-none font-sans"
                                  />
                                ) : (
                                  <span className="px-0.5">{row.namaSppg || "-"}</span>
                                )}
                              </td>
                              <td className="border border-black px-1.5 py-2 text-center align-middle uppercase tabular-nums">
                                {showInlineTableInputs ? (
                                  <>
                                    <input
                                      type="text"
                                      value={row.idSppg}
                                      onChange={(e) =>
                                        handleBulkTextChange(ws, row.hariKe, "idSppg", e.target.value)
                                      }
                                      aria-label={`ID SPPG Hari Ke-${row.hariKe}`}
                                      className="print:hidden w-full min-w-[74px] px-1 py-1 text-center text-[11px] uppercase bg-white border border-amber-300 rounded focus:outline-none focus:border-emerald-600 font-sans"
                                    />
                                    <span className="hidden print:inline">{row.idSppg || "-"}</span>
                                  </>
                                ) : isEditing ? (
                                  <input
                                    type="text"
                                    value={editIdSppgDraft}
                                    onChange={(e) => setEditIdSppgDraft(e.target.value)}
                                    className="w-full min-w-[74px] px-1 py-1 text-center text-[11px] uppercase bg-white border border-emerald-500 rounded focus:outline-none font-sans"
                                  />
                                ) : (
                                  row.idSppg || "-"
                                )}
                              </td>
                              <td className="border border-black px-1.5 py-2 text-left align-middle whitespace-nowrap">
                                {showInlineTableInputs ? (
                                  <>
                                    <input
                                      type="text"
                                      value={row.formattedDate}
                                      onChange={(e) =>
                                        handleBulkTextChange(ws, row.hariKe, "tanggal", e.target.value)
                                      }
                                      aria-label={`Hari / Tanggal Hari Ke-${row.hariKe}`}
                                      className="print:hidden w-full min-w-[138px] px-1.5 py-1 text-[11px] bg-white border border-amber-300 rounded focus:outline-none focus:border-emerald-600 font-sans"
                                    />
                                    <span className="hidden print:inline">{row.formattedDate}</span>
                                  </>
                                ) : isEditing ? (
                                  <input
                                    type="text"
                                    value={editDateDraft}
                                    onChange={(e) => setEditDateDraft(e.target.value)}
                                    className="w-full min-w-[138px] px-1.5 py-1 text-[11px] bg-white border border-emerald-500 rounded focus:outline-none font-sans"
                                  />
                                ) : (
                                  <span className="px-0.5">{row.formattedDate}</span>
                                )}
                              </td>
                              <td className="border border-black px-1.5 py-2 text-left align-middle leading-snug">
                                {showInlineTableInputs ? (
                                  <>
                                    <div className="print:hidden space-y-1 font-sans">
                                      <textarea
                                        rows={2}
                                        value={row.namaMenu}
                                        onChange={(e) =>
                                          handleBulkTextChange(ws, row.hariKe, "namaMenu", e.target.value)
                                        }
                                        placeholder="Nama Menu Harian..."
                                        className="w-full min-w-[160px] px-1.5 py-1 text-[11px] bg-white border border-amber-300 rounded focus:outline-none focus:border-emerald-600 resize-y"
                                      />
                                      <input
                                        type="text"
                                        value={row.caption}
                                        onChange={(e) =>
                                          patchDayOverride(ws.photoGroupKey, row.hariKe, {}, e.target.value)
                                        }
                                        placeholder="Catatan singkat (opsional)"
                                        className="w-full px-1.5 py-0.5 text-[10px] bg-white border border-slate-200 rounded focus:outline-none focus:border-emerald-600"
                                      />
                                    </div>
                                    <div className="hidden print:block">
                                      <div>{row.namaMenu || "-"}</div>
                                      {row.caption && (
                                        <div className="text-[10px] text-slate-600 mt-0.5 italic">
                                          {row.caption}
                                        </div>
                                      )}
                                    </div>
                                  </>
                                ) : isEditing ? (
                                  <div className="space-y-1.5 no-print font-sans">
                                    <textarea
                                      rows={2}
                                      value={editMenuDraft}
                                      onChange={(e) => setEditMenuDraft(e.target.value)}
                                      placeholder="Nama Menu Harian..."
                                      className="w-full px-2 py-1 text-xs border border-emerald-500 rounded focus:outline-none"
                                    />
                                    <input
                                      type="text"
                                      value={editCaptionDraft}
                                      onChange={(e) => setEditCaptionDraft(e.target.value)}
                                      placeholder="Catatan singkat (opsional)"
                                      className="w-full px-2 py-1 text-xs border border-slate-300 rounded focus:outline-none focus:border-emerald-500"
                                    />
                                  </div>
                                ) : (
                                  <div className="px-0.5">
                                    <div>{row.namaMenu || "-"}</div>
                                    {row.caption && (
                                      <div className="text-[10px] text-slate-600 mt-0.5 italic">
                                        {row.caption}
                                      </div>
                                    )}
                                  </div>
                                )}
                              </td>
                              <td
                                className="border border-black p-1.5 text-center align-middle"
                                onDragOver={(e) => {
                                  e.preventDefault();
                                  setDragOverDay(row.hariKe);
                                }}
                                onDragLeave={() => setDragOverDay(null)}
                                onDrop={(e) => handleDropOnDay(ws.photoGroupKey, row.hariKe, e)}
                              >
                                <div className="flex flex-col items-center justify-center gap-1.5">
                                  {isBusy ? (
                                    <div className="w-[176px] h-[106px] border border-emerald-300 bg-emerald-50/70 flex flex-col items-center justify-center gap-1 text-[10px] font-bold text-emerald-700 font-sans">
                                      <Loader2 className="w-4 h-4 animate-spin" />
                                      <span>Memproses foto...</span>
                                    </div>
                                  ) : photoSrc ? (
                                    <div className="relative group w-[176px] h-[106px] overflow-hidden border border-slate-300 bg-slate-50">
                                      <img
                                        src={photoSrc}
                                        alt={`Foto Menu Hari Ke-${row.hariKe}`}
                                        onClick={() =>
                                          setLightboxTarget({
                                            wsId: ws.config.id,
                                            hariKe: row.hariKe,
                                          })
                                        }
                                        className="w-full h-full object-cover cursor-zoom-in"
                                        loading="lazy"
                                        decoding="async"
                                      />
                                      <div className="no-print absolute top-1 right-1 flex items-center gap-1 opacity-95">
                                        <button
                                          type="button"
                                          onClick={() =>
                                            setLightboxTarget({
                                              wsId: ws.config.id,
                                              hariKe: row.hariKe,
                                            })
                                          }
                                          title="Perbesar Foto"
                                          className="p-1 rounded bg-slate-900/80 hover:bg-slate-900 text-white cursor-pointer"
                                        >
                                          <ZoomIn className="w-3 h-3" />
                                        </button>
                                        <button
                                          type="button"
                                          onClick={() =>
                                            void handleRotateDayPhoto(ws.photoGroupKey, row.hariKe)
                                          }
                                          title="Putar 90°"
                                          className="p-1 rounded bg-slate-900/80 hover:bg-slate-900 text-white cursor-pointer"
                                        >
                                          <RotateCw className="w-3 h-3" />
                                        </button>
                                        <button
                                          type="button"
                                          onClick={() =>
                                            void handleDeletePhoto(ws.photoGroupKey, row.hariKe)
                                          }
                                          title="Hapus Foto"
                                          className="p-1 rounded bg-rose-600/90 hover:bg-rose-700 text-white cursor-pointer"
                                        >
                                          <Trash2 className="w-3 h-3" />
                                        </button>
                                      </div>
                                    </div>
                                  ) : (
                                    <div
                                      className={`w-[176px] h-[72px] border border-dashed flex flex-col items-center justify-center gap-0.5 text-[10px] transition print:border-transparent print:bg-transparent font-sans ${
                                        isDragging
                                          ? "border-emerald-600 bg-emerald-50 text-emerald-700 font-bold"
                                          : "border-slate-300 bg-slate-50/70 text-slate-400"
                                      }`}
                                    >
                                      <span className="no-print">
                                        {isDragging ? "Lepaskan foto di sini" : "Belum ada foto / seret ke sini"}
                                      </span>
                                    </div>
                                  )}

                                  {/* Upload & Posting Controls (Hidden on Print/PDF) */}
                                  <div className="no-print flex flex-col items-center gap-1 w-full font-sans">
                                    <div className="flex items-center justify-center gap-1.5 w-full">
                                      <label className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 text-[10px] font-bold cursor-pointer transition whitespace-nowrap">
                                        <Camera className="w-3 h-3" />
                                        <span>Kamera</span>
                                        <input
                                          type="file"
                                          accept="image/*"
                                          capture="environment"
                                          onChange={(e) =>
                                            handlePhotoInputChange(ws.photoGroupKey, row.hariKe, e)
                                          }
                                          className="hidden"
                                        />
                                      </label>
                                      <label className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-800 text-[10px] font-bold cursor-pointer transition whitespace-nowrap">
                                        <ImageIcon className="w-3 h-3" />
                                        <span>Galeri</span>
                                        <input
                                          type="file"
                                          accept="image/*"
                                          onChange={(e) =>
                                            handlePhotoInputChange(ws.photoGroupKey, row.hariKe, e)
                                          }
                                          className="hidden"
                                        />
                                      </label>
                                    </div>

                                    {photoSrc && !isPosted && (
                                      <button
                                        type="button"
                                        onClick={() =>
                                          void handlePostDayPhoto(ws.photoGroupKey, row.hariKe)
                                        }
                                        className="w-[176px] py-1 px-2 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-bold inline-flex items-center justify-center gap-1 cursor-pointer transition"
                                      >
                                        <Send className="w-3 h-3" />
                                        Posting Foto
                                      </button>
                                    )}

                                    {isPosted && (
                                      <div className="w-[176px] flex items-center justify-between text-[10px] text-emerald-700 font-bold px-0.5">
                                        <span className="inline-flex items-center gap-1">
                                          <CheckCircle2 className="w-3 h-3 shrink-0" />
                                          Diposting{postMeta?.postedAt ? ` • ${formatShortTimestamp(postMeta.postedAt)}` : ""}
                                        </span>
                                        <button
                                          type="button"
                                          onClick={() =>
                                            handleUnpostDayPhoto(ws.photoGroupKey, row.hariKe)
                                          }
                                          className="text-slate-400 hover:text-slate-700 underline font-semibold cursor-pointer"
                                        >
                                          Draft
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              </td>
                              <td className="no-print border border-black px-1 py-1.5 text-center align-middle font-sans bg-slate-50/60">
                                {showInlineTableInputs ? (
                                  <div className="flex items-center justify-center gap-1">
                                    {row.isEdited ? (
                                      <button
                                        type="button"
                                        onClick={() => resetDayOverride(ws, row.hariKe)}
                                        title="Reset baris ini ke otomatis"
                                        className="p-1 rounded bg-white hover:bg-rose-50 border border-slate-200 hover:border-rose-300 text-slate-600 hover:text-rose-700 cursor-pointer"
                                      >
                                        <RotateCcw className="w-3.5 h-3.5" />
                                      </button>
                                    ) : (
                                      <span className="text-[10px] font-semibold text-slate-400">Auto</span>
                                    )}
                                  </div>
                                ) : isEditing ? (
                                  <div className="flex items-center justify-center gap-1">
                                    <button
                                      type="button"
                                      onClick={() => saveEditDayMenu(ws, row.hariKe)}
                                      title="Simpan Perubahan Baris"
                                      className="p-1 rounded bg-emerald-600 hover:bg-emerald-700 text-white cursor-pointer"
                                    >
                                      <Check className="w-3.5 h-3.5" />
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => setEditingMenuKey(null)}
                                      title="Batal Edit"
                                      className="p-1 rounded bg-slate-200 hover:bg-slate-300 text-slate-700 cursor-pointer"
                                    >
                                      <X className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                ) : (
                                  <div className="flex items-center justify-center gap-1">
                                    <button
                                      type="button"
                                      onClick={() => startEditDayMenu(ws, row.hariKe)}
                                      title="Edit Nama SPPG, ID SPPG, Tanggal, Menu & Catatan"
                                      className="p-1 rounded bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-800 cursor-pointer transition"
                                    >
                                      <Pencil className="w-3.5 h-3.5" />
                                    </button>
                                    {row.isEdited && (
                                      <button
                                        type="button"
                                        onClick={() => resetDayOverride(ws, row.hariKe)}
                                        title="Reset baris ini ke otomatis"
                                        className="p-1 rounded bg-white hover:bg-rose-50 border border-slate-200 hover:border-rose-300 text-slate-500 hover:text-rose-700 cursor-pointer transition"
                                      >
                                        <RotateCcw className="w-3 h-3" />
                                      </button>
                                    )}
                                  </div>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>

                    <AutoSignatureBlock profile={profile} docKey="ikss" />
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* PHOTO LIGHTBOX MODAL */}
      {lightboxTarget !== null && activeLightboxWs && activeLightboxRow && (
        <div
          className="fixed inset-0 z-[9998] bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-4 no-print"
          onClick={() => setLightboxTarget(null)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full overflow-hidden flex flex-col max-h-[90vh]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-3">
              <div className="min-w-0">
                <div className="text-sm font-extrabold text-slate-900">
                  Hari Ke-{activeLightboxRow.hariKe} — {activeLightboxRow.formattedDate}
                  {worksheetsComputed.length > 1 && (
                    <span className="ml-2 text-xs font-bold text-indigo-700">
                      ({activeLightboxWs.displayCategoryTitle})
                    </span>
                  )}
                </div>
                <div className="text-xs text-slate-600 truncate">
                  {activeLightboxRow.namaMenu || "Menu belum diisi"}
                </div>
              </div>
              <button
                type="button"
                onClick={() => setLightboxTarget(null)}
                className="p-1.5 text-slate-400 hover:text-slate-800 hover:bg-slate-100 rounded-lg cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="bg-slate-900 flex items-center justify-center p-3 min-h-[260px] max-h-[58vh] overflow-hidden relative">
              {activeLightboxPhoto ? (
                <img
                  src={activeLightboxPhoto}
                  alt={`Foto Menu Hari Ke-${activeLightboxRow.hariKe}`}
                  className="max-h-[54vh] max-w-full object-contain rounded"
                />
              ) : (
                <div className="text-slate-400 text-xs">Foto belum diunggah</div>
              )}

              {lightboxTarget.hariKe > 1 && (
                <button
                  type="button"
                  onClick={() =>
                    setLightboxTarget((prev) =>
                      prev && prev.hariKe > 1 ? { ...prev, hariKe: prev.hariKe - 1 } : prev
                    )
                  }
                  className="absolute left-3 top-1/2 -translate-y-1/2 p-2 rounded-full bg-white/15 hover:bg-white/30 text-white cursor-pointer"
                  title="Hari Sebelumnya"
                >
                  <ChevronLeft className="w-5 h-5" />
                </button>
              )}
              {lightboxTarget.hariKe < activeLightboxWs.totalDays && (
                <button
                  type="button"
                  onClick={() =>
                    setLightboxTarget((prev) =>
                      prev && prev.hariKe < activeLightboxWs.totalDays
                        ? { ...prev, hariKe: prev.hariKe + 1 }
                        : prev
                    )
                  }
                  className="absolute right-3 top-1/2 -translate-y-1/2 p-2 rounded-full bg-white/15 hover:bg-white/30 text-white cursor-pointer"
                  title="Hari Berikutnya"
                >
                  <ChevronRight className="w-5 h-5" />
                </button>
              )}
            </div>

            <div className="px-4 py-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
              <div className="flex flex-wrap items-center gap-2">
                <label className="px-3 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer">
                  <Camera className="w-3.5 h-3.5" />
                  <span>Ganti Kamera</span>
                  <input
                    type="file"
                    accept="image/*"
                    capture="environment"
                    onChange={(e) =>
                      handlePhotoInputChange(
                        activeLightboxWs.photoGroupKey,
                        activeLightboxRow.hariKe,
                        e
                      )
                    }
                    className="hidden"
                  />
                </label>
                <label className="px-3 py-1.5 rounded-lg bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-800 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer">
                  <ImageIcon className="w-3.5 h-3.5" />
                  <span>Ganti Galeri</span>
                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e) =>
                      handlePhotoInputChange(
                        activeLightboxWs.photoGroupKey,
                        activeLightboxRow.hariKe,
                        e
                      )
                    }
                    className="hidden"
                  />
                </label>
                {activeLightboxPhoto && (
                  <>
                    <button
                      type="button"
                      onClick={() =>
                        void handleRotateDayPhoto(
                          activeLightboxWs.photoGroupKey,
                          activeLightboxRow.hariKe
                        )
                      }
                      className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
                    >
                      <RotateCw className="w-3.5 h-3.5" />
                      Putar 90°
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        void handleDeletePhoto(
                          activeLightboxWs.photoGroupKey,
                          activeLightboxRow.hariKe
                        )
                      }
                      className="px-3 py-1.5 rounded-lg border border-rose-200 bg-rose-50 hover:bg-rose-100 text-rose-700 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      Hapus
                    </button>
                  </>
                )}
              </div>

              {activeLightboxPhoto && (
                activeLightboxMeta?.posted ? (
                  <span className="inline-flex items-center gap-1.5 text-xs font-bold text-emerald-700">
                    <CheckCircle2 className="w-4 h-4" />
                    Sudah Diposting ({formatShortTimestamp(activeLightboxMeta.postedAt)})
                  </span>
                ) : (
                  <button
                    type="button"
                    onClick={() =>
                      void handlePostDayPhoto(
                        activeLightboxWs.photoGroupKey,
                        activeLightboxRow.hariKe
                      )
                    }
                    className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
                  >
                    <Send className="w-3.5 h-3.5" />
                    Posting Foto Ini
                  </button>
                )
              )}
            </div>
          </div>
        </div>
      )}

      {/* Print Preview Modal (mounted strictly when open) */}
      {showPrintModal && (
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
      )}
    </div>
  );
}
