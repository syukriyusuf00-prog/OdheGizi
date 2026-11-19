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
  Edit3,
  ChevronLeft,
  ChevronRight,
  ZoomIn,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  DEFAULT_UNIFIED_KOP_CONFIG,
  FoodCostDay,
  HariPM,
  IkssDailyPostMeta,
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

  // Lightbox & inline menu editor state
  const [lightboxDay, setLightboxDay] = useState<number | null>(null);
  const [editingMenuDay, setEditingMenuDay] = useState<number | null>(null);
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

  // Stable references so useMemo caches never bust spuriously
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

  const selectedFcTypeSet = useMemo(() => new Set(selectedFoodCostTypes), [selectedFoodCostTypes]);
  const selectedSasaranSet = useMemo(() => new Set(selectedSasaranIds), [selectedSasaranIds]);

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

  const autoCategoryTitle = useMemo(() => {
    const matchedPreset = QUICK_PRESETS.find(
      (p) =>
        p.sasaranIds.length === selectedSasaranIds.length &&
        p.sasaranIds.every((id) => selectedSasaranSet.has(id))
    );
    if (matchedPreset) return matchedPreset.title;
    const labels = SASARAN_KATEGORI_LIST.filter((s) => selectedSasaranSet.has(s.id)).map((s) =>
      s.shortLabel.toUpperCase()
    );
    if (labels.length === 0) return "SEMUA KATEGORI";
    if (labels.length <= 3) return labels.join(" / ");
    return `${labels.slice(0, 3).join(", ")} +${labels.length - 3} LAINNYA`;
  }, [selectedSasaranIds, selectedSasaranSet]);

  const displayCategoryTitle = (
    ikssData.customCategoryTitle?.trim() || autoCategoryTitle
  ).toUpperCase();

  const autoPeriodeBulan = useMemo(
    () => derivePeriodeBulanLabel(profile.periodeDates || []),
    [profile.periodeDates]
  );
  const displayPeriodeBulan = (
    ikssData.customPeriodeLabel?.trim() || autoPeriodeBulan
  ).toUpperCase();

  const photoGroupKey = useMemo(() => {
    const sorted = [...selectedSasaranIds].sort().join("_");
    return sorted || "default";
  }, [selectedSasaranIds]);

  // Safely cap totalDays so corrupted arrays can never spin thousands of DOM rows
  const totalDays = Math.min(
    62,
    Math.max(profile.periodeDates?.length || 0, harianPM?.length || 0, 10)
  );

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

  // Fast base daily rows (dates, auto menu names, beneficiary counts) — zero heavy nutrition calc
  const baseDailyRows = useMemo(() => {
    const hasKecilSelected = selectedSasaranIds.some((id) => PORSI_KECIL_IDS.has(id));
    const hasBesarSelected = selectedSasaranIds.some((id) => PORSI_BESAR_IDS.has(id));
    const isAlergiOnly =
      selectedFoodCostTypes.length > 0 &&
      selectedFoodCostTypes.every((t) => t.toLowerCase().includes("alergi"));

    const matchingPlans = (normalizedMenu.plans || []).filter(
      (plan) =>
        selectedSasaranIds.length === 0 ||
        (plan.sasaranIds || []).some((id) => selectedSasaranSet.has(id))
    );

    const rows: Array<{
      hariKe: number;
      dateStr: string;
      formattedDate: string;
      autoNamaMenu: string;
      pmKecilCount: number;
      pmBesarCount: number;
      jumlahManfaat: number;
      photoKey: string;
    }> = [];

    for (let i = 0; i < totalDays; i++) {
      const hariKe = i + 1;
      const dateStr = profile.periodeDates?.[i] || "";
      const formattedDate = dateStr ? formatIndoFullDate(dateStr) : `Hari Ke-${hariKe}`;

      const dayPM = pmByDay.get(hariKe) || harianPM[0];
      let pmKecilCount = 0;
      let pmBesarCount = 0;
      if (dayPM?.sasaran) {
        for (let sIdx = 0; sIdx < dayPM.sasaran.length; sIdx++) {
          const sas = dayPM.sasaran[sIdx];
          if (!selectedSasaranSet.has(sas.id)) continue;
          const count = isAlergiOnly
            ? (Number(sas.alergiKecil) || 0) + (Number(sas.alergiBesar) || 0)
            : (Number(sas.porsiKecil) || 0) + (Number(sas.porsiBesar) || 0);
          if (PORSI_KECIL_IDS.has(sas.id)) pmKecilCount += count;
          else pmBesarCount += count;
        }
      }

      // Resolve menu name cleanly from Master Menu plans first, then Food Cost
      const masterMenuNames: string[] = [];
      for (let pIdx = 0; pIdx < matchingPlans.length; pIdx++) {
        const dayItem = matchingPlans[pIdx].days?.[i];
        if (!dayItem) continue;
        const composed = composeMenuName(dayItem).trim();
        if (composed && !masterMenuNames.includes(composed)) {
          masterMenuNames.push(composed);
        }
      }

      let autoNamaMenu = "";
      if (masterMenuNames.length > 0) {
        autoNamaMenu = masterMenuNames.join(" | ");
      } else {
        const matchedDays = (fcByDay.get(hariKe) || []).filter((fc) =>
          selectedFcTypeSet.has(fc.jenisMenu)
        );
        const menuNameParts: string[] = [];
        for (let mIdx = 0; mIdx < matchedDays.length; mIdx++) {
          const fc = matchedDays[mIdx];
          if (fc.namaMenu && fc.namaMenu.trim()) {
            menuNameParts.push(fc.namaMenu.trim());
          } else {
            const hasKecilBahan = (fc.porsiKecilBahan || []).some(
              (b) => (Number(b.beratBB) || 0) > 0 && (b.nama?.trim() || b.tkpiId)
            );
            const hasBesarBahan = (fc.porsiBesarBahan || []).some(
              (b) => (Number(b.beratBB) || 0) > 0 && (b.nama?.trim() || b.tkpiId)
            );
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
        }
        for (let cIdx = 0; cIdx < customFcTables.length; cIdx++) {
          const tbl = customFcTables[cIdx];
          if (!selectedFcTypeSet.has(`custom:${tbl.id}`)) continue;
          if (tbl.hariKe !== undefined && Number(tbl.hariKe) !== hariKe) continue;
          if (tbl.namaMenu) menuNameParts.push(String(tbl.namaMenu));
        }
        if (menuNameParts.length > 0) {
          autoNamaMenu = Array.from(new Set(menuNameParts)).join(" | ");
        }
      }

      rows.push({
        hariKe,
        dateStr,
        formattedDate,
        autoNamaMenu,
        pmKecilCount,
        pmBesarCount,
        jumlahManfaat: pmKecilCount + pmBesarCount,
        photoKey: `${photoGroupKey}__day_${hariKe}`,
      });
    }

    return rows;
  }, [
    totalDays,
    profile.periodeDates,
    pmByDay,
    harianPM,
    fcByDay,
    customFcTables,
    selectedFoodCostTypes,
    selectedFcTypeSet,
    selectedSasaranIds,
    selectedSasaranSet,
    normalizedMenu,
    photoGroupKey,
  ]);

  // Compute heavy nutrition only when on IKSS Nilai Gizi sub-view so Foto Menu tab stays 100% instant
  const nutritionByDay = useMemo(() => {
    const map = new Map<number, DailyNutritionValues>();
    if (subView !== "nilai_gizi") return map;

    const hasKecilSelected = selectedSasaranIds.some((id) => PORSI_KECIL_IDS.has(id));
    const hasBesarSelected = selectedSasaranIds.some((id) => PORSI_BESAR_IDS.has(id));

    for (let i = 0; i < baseDailyRows.length; i++) {
      const base = baseDailyRows[i];
      const { hariKe, pmKecilCount, pmBesarCount } = base;
      const matchedDays = (fcByDay.get(hariKe) || []).filter((fc) =>
        selectedFcTypeSet.has(fc.jenisMenu)
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
      }

      for (let cIdx = 0; cIdx < customFcTables.length; cIdx++) {
        const tbl = customFcTables[cIdx];
        if (!selectedFcTypeSet.has(`custom:${tbl.id}`)) continue;
        if (!Array.isArray(tbl.bahanList) || tbl.bahanList.length === 0) continue;
        if (tbl.hariKe !== undefined && Number(tbl.hariKe) !== hariKe) continue;
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
      }

      if (nutriSamples.length > 0) {
        const sumW = nutriSamples.reduce((a, b) => a + b.weight, 0) || nutriSamples.length;
        map.set(hariKe, {
          energi: nutriSamples.reduce((a, b) => a + b.energi * b.weight, 0) / sumW,
          protein: nutriSamples.reduce((a, b) => a + b.protein * b.weight, 0) / sumW,
          lemak: nutriSamples.reduce((a, b) => a + b.lemak * b.weight, 0) / sumW,
          kh: nutriSamples.reduce((a, b) => a + b.kh * b.weight, 0) / sumW,
          serat: nutriSamples.reduce((a, b) => a + b.serat * b.weight, 0) / sumW,
          hasNutrition: true,
        });
      } else {
        map.set(hariKe, EMPTY_NUTRITION);
      }
    }

    return map;
  }, [
    subView,
    baseDailyRows,
    fcByDay,
    customFcTables,
    selectedFcTypeSet,
    selectedSasaranIds,
    tkpiMap,
    bumbuPct,
  ]);

  const dailyPosts = ikssData.dailyPosts || {};

  const getPhotoForDay = useCallback(
    (hariKe: number): string => {
      return (
        photoMap[`${photoGroupKey}__day_${hariKe}`] ||
        photoMap[`day_${hariKe}`] ||
        ""
      );
    },
    [photoMap, photoGroupKey]
  );

  const getPostMetaForDay = useCallback(
    (hariKe: number): IkssDailyPostMeta | undefined => {
      return dailyPosts[`${photoGroupKey}__day_${hariKe}`] || dailyPosts[`day_${hariKe}`];
    },
    [dailyPosts, photoGroupKey]
  );

  const dailyRows = useMemo(() => {
    return baseDailyRows.map((row) => {
      const nut = nutritionByDay.get(row.hariKe) || EMPTY_NUTRITION;
      const meta = dailyPosts[row.photoKey] || dailyPosts[`day_${row.hariKe}`];
      const namaMenu = (meta?.customMenuName || "").trim() || row.autoNamaMenu;
      return {
        ...row,
        namaMenu,
        caption: meta?.caption || "",
        energi: nut.energi,
        protein: nut.protein,
        lemak: nut.lemak,
        kh: nut.kh,
        serat: nut.serat,
        hasNutrition: nut.hasNutrition,
      };
    });
  }, [baseDailyRows, nutritionByDay, dailyPosts]);

  // Photo & Posting KPI stats
  const photoStats = useMemo(() => {
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
    const pct = baseDailyRows.length > 0 ? Math.round((posted / baseDailyRows.length) * 100) : 0;
    return { total: baseDailyRows.length, uploaded, posted, draft, empty, pct };
  }, [baseDailyRows, photoMap, dailyPosts]);

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

  const autoAkg = useMemo<AkgPresetTarget>(() => {
    const matched = selectedSasaranIds.map((id) => SASARAN_AKG_REF[id]).filter(Boolean);
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

  const effectiveAkg = useMemo(
    () => ({
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
        ikssData.akgOverride?.kh !== undefined ? ikssData.akgOverride.kh : autoAkg.kh,
      serat:
        ikssData.akgOverride?.serat !== undefined
          ? ikssData.akgOverride.serat
          : autoAkg.serat,
    }),
    [ikssData.akgOverride, autoAkg]
  );

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

  // Sub-view switch handler (remember user choice)
  const handleSwitchSubView = (nextView: "nilai_gizi" | "foto_menu") => {
    setSubView(nextView);
    setIkssData((prev) =>
      prev.activeSubView === nextView ? prev : { ...prev, activeSubView: nextView }
    );
  };

  // Category & Preset Handlers
  const toggleFoodCostType = (typeId: string) => {
    setIkssData((prev) => {
      const current = prev.selectedFoodCostTypes || [];
      const exists = current.includes(typeId);
      const next = exists ? current.filter((x) => x !== typeId) : [...current, typeId];
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
      const next = exists ? current.filter((x) => x !== sasId) : [...current, sasId];
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

  // =========================================================================
  // PHOTO UPLOAD, ROTATE, POSTING & BULK HANDLERS (100% NON-BLOCKING)
  // =========================================================================

  const processSinglePhotoFile = useCallback(
    async (hariKe: number, file: File) => {
      setPhotoBanner(null);
      setProcessingDays((prev) => ({ ...prev, [hariKe]: true }));
      try {
        const compressed = await compressMenuPhoto(file, {
          maxWidth: 640,
          maxHeight: 480,
          targetMaxChars: 72_000,
        });
        const key = `${photoGroupKey}__day_${hariKe}`;
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
    [photoGroupKey, ikssData.autoPostOnUpload, persistPhotos, setIkssData, showNotice]
  );

  const handlePhotoInputChange = (
    hariKe: number,
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    void processSinglePhotoFile(hariKe, file);
  };

  const handleDropOnDay = (hariKe: number, e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverDay(null);
    const file = e.dataTransfer.files?.[0];
    if (file && file.type.startsWith("image/")) {
      void processSinglePhotoFile(hariKe, file);
    } else if (file) {
      showNotice("error", "Harap lepaskan file gambar (JPG, PNG, WEBP).");
    }
  };

  const handleBulkUploadChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    const files: File[] = fileList
      ? Array.from<File>(fileList).filter((f) => f.type.startsWith("image/"))
      : [];
    e.target.value = "";
    if (files.length === 0) return;

    // Assign sequentially to unfilled days first, then remaining days from day 1
    const unfilledDays = baseDailyRows
      .filter((r) => !photoMapRef.current[r.photoKey] && !photoMapRef.current[`day_${r.hariKe}`])
      .map((r) => r.hariKe);
    const targetDays =
      unfilledDays.length > 0
        ? unfilledDays
        : baseDailyRows.map((r) => r.hariKe);

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
        const key = `${photoGroupKey}__day_${hariKe}`;
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
      // Yield back to browser paint loop between images
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
          ? `${successCount} foto menu harian berhasil diunggah & diposting${res.cloudOk ? " ke cloud" : " (tersimpan lokal)"}.`
          : `${successCount} foto berhasil diunggah ke tabel harian. Klik "Posting Semua Foto" untuk mempublikasikan.`
      );
    } else {
      showNotice("error", "Tidak ada foto valid yang berhasil diproses.");
    }
  };

  const handlePostDayPhoto = async (hariKe: number) => {
    const photoSrc = getPhotoForDay(hariKe);
    if (!photoSrc) {
      showNotice("error", `Unggah foto menu Hari Ke-${hariKe} terlebih dahulu sebelum memposting.`);
      return;
    }
    const key = `${photoGroupKey}__day_${hariKe}`;
    const nowIso = new Date().toISOString();

    // Ensure photo is stored under current group key
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

  const handleUnpostDayPhoto = (hariKe: number) => {
    const key = `${photoGroupKey}__day_${hariKe}`;
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

  const handlePostAllPhotos = async () => {
    const uploadedKeys: string[] = [];
    for (let i = 0; i < baseDailyRows.length; i++) {
      const r = baseDailyRows[i];
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
    baseDailyRows.forEach((r) => {
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

  const handleRotateDayPhoto = async (hariKe: number) => {
    const photoSrc = getPhotoForDay(hariKe);
    if (!photoSrc) return;
    setProcessingDays((prev) => ({ ...prev, [hariKe]: true }));
    try {
      const rotated = await rotatePhotoDataUrl90(photoSrc);
      const key = `${photoGroupKey}__day_${hariKe}`;
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

  const handleDeletePhoto = async (hariKe: number) => {
    const key = `${photoGroupKey}__day_${hariKe}`;
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

    if (lightboxDay === hariKe) setLightboxDay(null);
    await persistPhotos(nextPhotos, false);
  };

  const startEditDayMenu = (hariKe: number) => {
    const row = dailyRows.find((r) => r.hariKe === hariKe);
    const meta = getPostMetaForDay(hariKe);
    setEditingMenuDay(hariKe);
    setEditMenuDraft(meta?.customMenuName ?? row?.namaMenu ?? "");
    setEditCaptionDraft(meta?.caption ?? "");
  };

  const saveEditDayMenu = (hariKe: number) => {
    const key = `${photoGroupKey}__day_${hariKe}`;
    const row = baseDailyRows.find((r) => r.hariKe === hariKe);
    const trimmedMenu = editMenuDraft.trim();
    const isSameAsAuto = trimmedMenu === (row?.autoNamaMenu || "").trim();

    setIkssData((prev) => ({
      ...prev,
      dailyPosts: {
        ...(prev.dailyPosts || {}),
        [key]: {
          ...(prev.dailyPosts?.[key] || { posted: false }),
          customMenuName: isSameAsAuto ? "" : trimmedMenu,
          caption: editCaptionDraft.trim(),
          updatedAt: new Date().toISOString(),
        },
      },
    }));
    setEditingMenuDay(null);
  };

  // =========================================================================
  // PRINT / PDF / EXCEL EXPORTS
  // =========================================================================

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
    if (photoFilter !== "semua") {
      setPhotoFilter("semua");
      setTimeout(() => {
        printElementSmart(activePrintId, {
          title: activePrintTitle,
          paperSize,
          orientation: "auto",
        });
      }, 60);
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
    if (photoFilter !== "semua") {
      setPhotoFilter("semua");
      await new Promise<void>((r) => setTimeout(r, 60));
    }
    setIsExportingPdf(true);
    setExportStatusMsg("Menyiapkan dokumen PDF presisi...");
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
          "Catatan",
          "Status Foto Menu",
        ],
        ...dailyRows.map((r) => {
          const hasPhoto = Boolean(getPhotoForDay(r.hariKe));
          const meta = getPostMetaForDay(r.hariKe);
          const status = !hasPhoto
            ? "Belum Diunggah"
            : meta?.posted
            ? `Diposting (${formatShortTimestamp(meta.postedAt)})`
            : "Draft (Belum Diposting)";
          return [
            r.hariKe,
            (profile.namaLembaga || "").toUpperCase(),
            (profile.idSPPG || "").toUpperCase(),
            r.formattedDate,
            r.namaMenu || "-",
            r.caption || "-",
            status,
          ];
        }),
      ];
      const ws = XLSX.utils.aoa_to_sheet(sheetData);
      ws["!cols"] = [
        { wch: 10 },
        { wch: 32 },
        { wch: 14 },
        { wch: 24 },
        { wch: 44 },
        { wch: 24 },
        { wch: 28 },
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

  // Active lightbox row info
  const activeLightboxRow =
    lightboxDay !== null ? dailyRows.find((r) => r.hariKe === lightboxDay) || null : null;
  const activeLightboxPhoto = lightboxDay !== null ? getPhotoForDay(lightboxDay) : "";
  const activeLightboxMeta = lightboxDay !== null ? getPostMetaForDay(lightboxDay) : undefined;

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
              Terhubung otomatis dengan <strong>Profil SPPG</strong> (identitas, tanggal periode &amp; TTD) dan perencanaan <strong>Food Cost</strong>. Unggah &amp; posting foto menu harian langsung dari kamera atau galeri tanpa hambatan.
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
                ({photoStats.posted}/{photoStats.total})
              </span>
            </button>
          </div>
        </div>

        {/* Action Toolbar */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-slate-100">
          <div className="flex flex-wrap items-center gap-3">
            <button
              type="button"
              onClick={() => setShowCategoryDrawer((v) => !v)}
              className="px-3 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap"
            >
              <Layers className="w-3.5 h-3.5 text-emerald-600" />
              <span>
                {showCategoryDrawer
                  ? "Sembunyikan Pilihan Kategori"
                  : `Pilih / Gabung Kategori (${selectedSasaranIds.length})`}
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
              onClick={() => setShowPrintModal(true)}
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

        {/* PDF export status bar */}
        {exportStatusMsg && (
          <div className="px-3.5 py-2 rounded-xl bg-indigo-50 border border-indigo-200 text-indigo-800 text-xs font-semibold flex items-center gap-2">
            <Loader2 className="w-3.5 h-3.5 animate-spin text-indigo-600 shrink-0" />
            <span>{exportStatusMsg}</span>
          </div>
        )}

        {/* FLEXIBLE CATEGORY & MERGE SELECTOR */}
        {showCategoryDrawer && (
          <div className="pt-4 border-t border-slate-100 space-y-4">
            <div className="space-y-1.5">
              <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
                Preset Cepat Kategori IKSS
              </span>
              <div className="flex flex-wrap gap-1.5">
                {QUICK_PRESETS.map((preset) => {
                  const isActive =
                    preset.sasaranIds.length === selectedSasaranIds.length &&
                    preset.sasaranIds.every((id) => selectedSasaranSet.has(id));
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
                    const checked = selectedFcTypeSet.has(fcType.id);
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
                    const checked = selectedSasaranSet.has(sas.id);
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

            <div className="bg-[#00B050] border border-black py-2 px-3 text-center">
              <h3 className="text-sm font-bold uppercase text-black tracking-wide">
                REKAPAN NILAI GIZI {displayCategoryTitle}
              </h3>
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

            <AutoSignatureBlock profile={profile} docKey="ikss" />
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
                  onChange={handleBulkUploadChange}
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
                  onClick={handlePostAllPhotos}
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
                const photoSrc = getPhotoForDay(row.hariKe);
                const postMeta = getPostMetaForDay(row.hariKe);
                const isPosted = Boolean(photoSrc && postMeta?.posted);
                const isBusy = Boolean(processingDays[row.hariKe]);
                const isDragging = dragOverDay === row.hariKe;
                const isEditing = editingMenuDay === row.hariKe;

                return (
                  <div
                    key={row.hariKe}
                    onDragOver={(e) => {
                      e.preventDefault();
                      setDragOverDay(row.hariKe);
                    }}
                    onDragLeave={() => setDragOverDay(null)}
                    onDrop={(e) => handleDropOnDay(row.hariKe, e)}
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
                              onClick={() => setLightboxDay(row.hariKe)}
                              className="w-full h-full object-cover cursor-zoom-in"
                              loading="lazy"
                              decoding="async"
                            />
                            <div className="absolute top-1.5 right-1.5 flex items-center gap-1">
                              <button
                                type="button"
                                onClick={() => setLightboxDay(row.hariKe)}
                                title="Perbesar Foto"
                                className="p-1.5 rounded-lg bg-slate-900/75 hover:bg-slate-900 text-white cursor-pointer"
                              >
                                <ZoomIn className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => handleRotateDayPhoto(row.hariKe)}
                                title="Putar 90°"
                                className="p-1.5 rounded-lg bg-slate-900/75 hover:bg-slate-900 text-white cursor-pointer"
                              >
                                <RotateCw className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => handleDeletePhoto(row.hariKe)}
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

                      {/* Menu Name & Caption */}
                      {isEditing ? (
                        <div className="space-y-1.5">
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
                          <div className="flex items-center justify-end gap-1.5">
                            <button
                              type="button"
                              onClick={() => setEditingMenuDay(null)}
                              className="px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100 rounded-md cursor-pointer"
                            >
                              Batal
                            </button>
                            <button
                              type="button"
                              onClick={() => saveEditDayMenu(row.hariKe)}
                              className="px-2.5 py-1 text-[11px] font-bold bg-emerald-600 text-white rounded-md cursor-pointer"
                            >
                              Simpan
                            </button>
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
                            onClick={() => startEditDayMenu(row.hariKe)}
                            title="Edit Nama Menu / Catatan"
                            className="p-1 text-slate-400 hover:text-slate-700 rounded cursor-pointer shrink-0"
                          >
                            <Edit3 className="w-3.5 h-3.5" />
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
                              onChange={(e) => handlePhotoInputChange(row.hariKe, e)}
                              className="hidden"
                            />
                          </label>
                          <label className="inline-flex items-center justify-center gap-1 px-2.5 py-1.5 rounded-lg bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-800 text-xs font-bold cursor-pointer transition">
                            <ImageIcon className="w-3.5 h-3.5" />
                            <span>Galeri</span>
                            <input
                              type="file"
                              accept="image/*"
                              onChange={(e) => handlePhotoInputChange(row.hariKe, e)}
                              className="hidden"
                            />
                          </label>
                        </div>

                        {photoSrc && !isPosted && (
                          <button
                            type="button"
                            onClick={() => handlePostDayPhoto(row.hariKe)}
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
                              onClick={() => handleUnpostDayPhoto(row.hariKe)}
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

          {/* OFFICIAL PRINTABLE IKSS FOTO MENU TABLE (ALWAYS MOUNTED SO PRINT & PDF WORK FROM ANY MODE) */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-x-auto p-4 sm:p-6">
            {photoViewMode === "kartu" && (
              <div className="mb-3 pb-2 border-b border-slate-100 flex items-center justify-between no-print">
                <span className="text-xs font-bold text-slate-700">
                  Pratinjau Lembar Resmi IKSS Foto Menu (Siap Cetak / PDF)
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

              <div className="bg-[#00B050] border border-black py-2 px-3 text-center">
                <h3 className="text-sm font-bold uppercase text-black tracking-wide">
                  REKAPAN MENU {displayCategoryTitle}
                </h3>
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
                  {(photoViewMode === "tabel" && !isExportingPdf && !showPrintModal
                    ? filteredPhotoRows
                    : dailyRows
                  ).map((row) => {
                    const photoSrc = getPhotoForDay(row.hariKe);
                    const postMeta = getPostMetaForDay(row.hariKe);
                    const isPosted = Boolean(photoSrc && postMeta?.posted);
                    const isBusy = Boolean(processingDays[row.hariKe]);
                    const isDragging = dragOverDay === row.hariKe;
                    const isEditing = editingMenuDay === row.hariKe;

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
                          {isEditing ? (
                            <div className="space-y-1.5 no-print font-sans">
                              <textarea
                                rows={2}
                                value={editMenuDraft}
                                onChange={(e) => setEditMenuDraft(e.target.value)}
                                placeholder="Nama Menu Harian..."
                                className="w-full px-2 py-1 text-xs border border-slate-300 rounded focus:outline-none focus:border-emerald-500"
                              />
                              <input
                                type="text"
                                value={editCaptionDraft}
                                onChange={(e) => setEditCaptionDraft(e.target.value)}
                                placeholder="Catatan singkat (opsional)"
                                className="w-full px-2 py-1 text-xs border border-slate-300 rounded focus:outline-none focus:border-emerald-500"
                              />
                              <div className="flex items-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => saveEditDayMenu(row.hariKe)}
                                  className="px-2.5 py-1 bg-emerald-600 text-white text-[11px] font-bold rounded cursor-pointer"
                                >
                                  Simpan
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setEditingMenuDay(null)}
                                  className="px-2 py-1 bg-slate-100 text-slate-700 text-[11px] font-semibold rounded cursor-pointer"
                                >
                                  Batal
                                </button>
                              </div>
                            </div>
                          ) : (
                            <div className="flex items-start justify-between gap-1.5 group">
                              <div>
                                <div>{row.namaMenu || "-"}</div>
                                {row.caption && (
                                  <div className="text-[10px] text-slate-600 mt-0.5 italic">
                                    {row.caption}
                                  </div>
                                )}
                              </div>
                              <button
                                type="button"
                                onClick={() => startEditDayMenu(row.hariKe)}
                                title="Edit Nama Menu / Catatan"
                                className="no-print p-1 rounded text-slate-400 hover:text-slate-800 hover:bg-slate-100 opacity-70 group-hover:opacity-100 cursor-pointer shrink-0"
                              >
                                <Edit3 className="w-3 h-3" />
                              </button>
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
                          onDrop={(e) => handleDropOnDay(row.hariKe, e)}
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
                                  onClick={() => setLightboxDay(row.hariKe)}
                                  className="w-full h-full object-cover cursor-zoom-in"
                                  loading="lazy"
                                  decoding="async"
                                />
                                <div className="no-print absolute top-1 right-1 flex items-center gap-1 opacity-95">
                                  <button
                                    type="button"
                                    onClick={() => setLightboxDay(row.hariKe)}
                                    title="Perbesar Foto"
                                    className="p-1 rounded bg-slate-900/80 hover:bg-slate-900 text-white cursor-pointer"
                                  >
                                    <ZoomIn className="w-3 h-3" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => handleRotateDayPhoto(row.hariKe)}
                                    title="Putar 90°"
                                    className="p-1 rounded bg-slate-900/80 hover:bg-slate-900 text-white cursor-pointer"
                                  >
                                    <RotateCw className="w-3 h-3" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => handleDeletePhoto(row.hariKe)}
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
                                    onChange={(e) => handlePhotoInputChange(row.hariKe, e)}
                                    className="hidden"
                                  />
                                </label>
                                <label className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-800 text-[10px] font-bold cursor-pointer transition whitespace-nowrap">
                                  <ImageIcon className="w-3 h-3" />
                                  <span>Galeri</span>
                                  <input
                                    type="file"
                                    accept="image/*"
                                    onChange={(e) => handlePhotoInputChange(row.hariKe, e)}
                                    className="hidden"
                                  />
                                </label>
                              </div>

                              {photoSrc && !isPosted && (
                                <button
                                  type="button"
                                  onClick={() => handlePostDayPhoto(row.hariKe)}
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
                                    onClick={() => handleUnpostDayPhoto(row.hariKe)}
                                    className="text-slate-400 hover:text-slate-700 underline font-semibold cursor-pointer"
                                  >
                                    Draft
                                  </button>
                                </div>
                              )}
                            </div>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>

              <AutoSignatureBlock profile={profile} docKey="ikss" />
            </div>
          </div>
        </div>
      )}

      {/* PHOTO LIGHTBOX MODAL */}
      {lightboxDay !== null && activeLightboxRow && (
        <div
          className="fixed inset-0 z-[9998] bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-4 no-print"
          onClick={() => setLightboxDay(null)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full overflow-hidden flex flex-col max-h-[90vh]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-3">
              <div className="min-w-0">
                <div className="text-sm font-extrabold text-slate-900">
                  Hari Ke-{activeLightboxRow.hariKe} — {activeLightboxRow.formattedDate}
                </div>
                <div className="text-xs text-slate-600 truncate">
                  {activeLightboxRow.namaMenu || "Menu belum diisi"}
                </div>
              </div>
              <button
                type="button"
                onClick={() => setLightboxDay(null)}
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

              {lightboxDay > 1 && (
                <button
                  type="button"
                  onClick={() => setLightboxDay((d) => (d && d > 1 ? d - 1 : d))}
                  className="absolute left-3 top-1/2 -translate-y-1/2 p-2 rounded-full bg-white/15 hover:bg-white/30 text-white cursor-pointer"
                  title="Hari Sebelumnya"
                >
                  <ChevronLeft className="w-5 h-5" />
                </button>
              )}
              {lightboxDay < totalDays && (
                <button
                  type="button"
                  onClick={() => setLightboxDay((d) => (d && d < totalDays ? d + 1 : d))}
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
                    onChange={(e) => handlePhotoInputChange(activeLightboxRow.hariKe, e)}
                    className="hidden"
                  />
                </label>
                <label className="px-3 py-1.5 rounded-lg bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-800 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer">
                  <ImageIcon className="w-3.5 h-3.5" />
                  <span>Ganti Galeri</span>
                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e) => handlePhotoInputChange(activeLightboxRow.hariKe, e)}
                    className="hidden"
                  />
                </label>
                {activeLightboxPhoto && (
                  <>
                    <button
                      type="button"
                      onClick={() => handleRotateDayPhoto(activeLightboxRow.hariKe)}
                      className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
                    >
                      <RotateCw className="w-3.5 h-3.5" />
                      Putar 90°
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDeletePhoto(activeLightboxRow.hariKe)}
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
                    onClick={() => handlePostDayPhoto(activeLightboxRow.hariKe)}
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
