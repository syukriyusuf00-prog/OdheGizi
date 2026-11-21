/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from "react";
import {
  TKPIItem,
  BahanMakananInput,
  FoodCostDay,
  FoodCostCategoryGroup,
  FoodCostCategoryEntry,
  MasterMenu,
  HariPM,
  UnifiedKopConfig,
  DEFAULT_UNIFIED_KOP_CONFIG,
} from "../types";
import { calculateDay, formatRupiah, getCountsForDay, CalcIngredientResult } from "../utils/calc";
import SearchableTkpiDropdown from "./SearchableTkpiDropdown";
import { PriceCalculatorPopover } from "./PriceCalculatorPopover";
import {
  Plus,
  Trash2,
  HelpCircle,
  Sparkles,
  AlertTriangle,
  CheckCircle2,
  Calculator,
  Info,
  Sliders,
  Printer,
  Download,
  FileSpreadsheet,
  Edit3,
  RefreshCw,
  Image as ImageIcon,
  Settings,
  Eye,
  Layers,
  Tag,
  ChevronDown,
  ChevronUp,
  X,
  Check,
  Send,
  Copy,
  Clipboard,
  ClipboardCheck,
  Combine,
  Split,
  Calendar,
  ChevronLeft,
  ChevronRight,
} from "lucide-react";
import { downloadElementAsImage, exportToPDF, printElementSmart } from "../lib/printUtils";
import { TARGET_AKG_LIMITS } from "../tkpiData";
import * as XLSX from "xlsx";

import { KopSuratRenderHeader, LogoCrop } from "./KopSuratConfigSection";
import AutoSignatureBlock from "./AutoSignatureBlock";
import PrintPreviewModal from "./PrintPreviewModal";
import {
  ALERGI_KATEGORI_LIST,
  ALL_KATEGORI_WITH_ALERGI,
  composeMenuName,
  ensureMasterMenuPlans,
  formatIndonesianDate as formatScheduleIndonesianDate,
  isAlergiSasaranId,
  MPASI_SASARAN_IDS,
  normalizeMenuItem,
  resolveCategoryPorsiDefault,
  resolveSasaranPmFromTable2,
  SASARAN_KATEGORI_LIST,
  SEKOLAH_SASARAN_IDS,
  stripAlergiPrefix,
  TIGAB_SASARAN_IDS,
} from "../utils/menuSchedule";

interface FoodCostTabProps {
  foodCostDays: FoodCostDay[];
  tkpiList: TKPIItem[];
  masterMenu: MasterMenu;
  harianPM: HariPM[];
  onFoodCostDaysChange: (updated: FoodCostDay[]) => void;
  profile?: any;
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
  bumbuPct?: number;
  setBumbuPct?: (val: number) => void;
  pmSettings?: any;
  setPmSettings?: (val: any) => void;
  onNavigateToMasterMenu?: () => void;
  onNavigateToPM?: () => void;
  onSendToNotaPesanan?: () => void;
  isAdmin?: boolean;
}

const formatThousandSeparator = (value: number | string | undefined | null): string => {
  if (value === undefined || value === null || value === "") return "";
  let num: number;
  if (typeof value === "number") {
    num = value;
  } else {
    const str = String(value).trim();
    if (!str) return "";
    const cleanStr = str.replace(/\./g, "").replace(/,/g, ".");
    num = parseFloat(cleanStr);
  }
  if (isNaN(num) || !isFinite(num)) return "";
  return new Intl.NumberFormat("id-ID").format(Math.round(num));
};

const parseThousandSeparator = (value: string): number => {
  const clean = value.replace(/[^0-9]/g, "");
  return clean ? parseInt(clean, 10) : 0;
};

const SASARAN_AKG_PRESET_MAP: Record<string, string> = {
  tk_paud_lb: "tk_paud",
  sd_kelas_1_3: "sd_kecil",
  sd_kelas_4_6: "sd_besar",
  smp_mts_smplb: "smp",
  sma_smk_ma: "sma_pendidik",
  pendidik: "sma_pendidik",
  tenaga_kependidikan: "sma_pendidik",
  anak_balita: "balita",
  anak_balita_13_59: "balita",
  balita_6_11: "balita",
  ibu_hamil: "bumil",
  ibu_menyusui: "busui",
  sampel: "sd_besar",
};

const getDefaultAkgTypeForGroup = (sasaranIds: string[], porsi: "besar" | "kecil"): string => {
  for (const id of sasaranIds) {
    const cleanId = stripAlergiPrefix(id);
    if (SASARAN_AKG_PRESET_MAP[cleanId]) return SASARAN_AKG_PRESET_MAP[cleanId];
  }
  return porsi === "besar" ? "sd_besar" : "sd_kecil";
};

const cloneBahanListFresh = (list: BahanMakananInput[], prefix = "cp"): BahanMakananInput[] =>
  (list || []).map((b, idx) => ({
    id: `${prefix}_${Date.now()}_${idx}_${Math.random().toString(36).slice(2, 6)}`,
    tkpiId: b.tkpiId || "",
    nama: b.nama,
    beratBB: Number(b.beratBB) || 0,
    urt: b.urt || "",
    hargaSatuan: Number(b.hargaSatuan) || 0,
    bdd: b.bdd,
    potong: b.potong,
    ekor: b.ekor,
    buah: b.buah,
    butir: b.butir,
    formula: b.formula,
    bufferBase: b.bufferBase,
    bufferCustomVal: b.bufferCustomVal,
    jumlahBufferChoice: b.jumlahBufferChoice,
    jumlahBufferCustomVal: b.jumlahBufferCustomVal,
  }));

interface CopiedFoodCostPayload {
  sourceLabel: string;
  sourcePorsi: "besar" | "kecil";
  copiedAt: string;
  bufferPct: number;
  bahanList: BahanMakananInput[];
}

interface EditableBddCellProps {
  value: number; // resolved BDD
  isCustom: boolean; // whether currently customized
  onChange: (newVal: number | undefined) => void;
}

function EditableBddCell({ value, isCustom, onChange }: EditableBddCellProps) {
  const [localVal, setLocalVal] = React.useState<string>(String(value));

  // Sync when the resolved value changes from outside (e.g. changing ingredient)
  React.useEffect(() => {
    setLocalVal(String(value));
  }, [value]);

  const handleBlur = () => {
    if (localVal === "") {
      onChange(undefined);
    } else {
      const parsed = Number(localVal);
      if (!isNaN(parsed) && parsed >= 0 && parsed <= 100) {
        onChange(parsed);
      } else {
        // Revert to current prop value
        setLocalVal(String(value));
      }
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "Enter") {
      handleBlur();
      e.currentTarget.blur();
    }
  };

  return (
    <div className="flex items-center justify-center">
      <input
        type="number"
        min="1"
        max="100"
        value={localVal}
        onChange={(e) => setLocalVal(e.target.value)}
        onBlur={handleBlur}
        onKeyDown={handleKeyDown}
        className={`w-full text-center bg-transparent border-0 font-mono font-bold p-0 text-xs focus:ring-1 focus:ring-indigo-500 rounded ${isCustom ? "text-indigo-600 bg-indigo-50/20" : "text-slate-950"}`}
        placeholder="bdd"
        title={isCustom ? "BDD Dikustomisasi (Hapus untuk menggunakan bawaan TKPI)" : "BDD Bawaan TKPI (Klik untuk kustomisasi)"}
      />
      <span className="text-slate-500 text-[9px] font-bold pr-1">%</span>
    </div>
  );
}

export default function FoodCostTab({
  foodCostDays,
  tkpiList,
  masterMenu,
  harianPM,
  onFoodCostDaysChange,
  profile,
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
  bumbuPct: propBumbuPct,
  setBumbuPct: propSetBumbuPct,
  pmSettings,
  onNavigateToMasterMenu,
  onNavigateToPM,
  onSendToNotaPesanan
}: FoodCostTabProps) {
  const normalizedMasterMenu = React.useMemo(
    () => ensureMasterMenuPlans(masterMenu, profile?.periodeDates),
    [masterMenu, profile?.periodeDates]
  );

  const scheduleDates = React.useMemo(
    () =>
      profile?.periodeDates && profile.periodeDates.length > 0
        ? profile.periodeDates
        : normalizedMasterMenu.scheduleConfig?.periodeDates || [],
    [profile?.periodeDates, normalizedMasterMenu]
  );

  const totalDays = Math.max(scheduleDates.length > 0 ? scheduleDates.length : 10, 1);

  const [selectedDay, setSelectedDay] = useState<number>(1);
  const [selectedGroup, setSelectedGroup] = useState<"sekolah" | "tigaB" | "custom">("sekolah");
  const [selectedType, setSelectedType] = useState<"Basah" | "Alergi" | "Kering" | "MP-ASI">("Basah");
  const [selectedPlanOverrideId, setSelectedPlanOverrideId] = useState<string>("");
  const [viewMode, setViewMode] = useState<"edit" | "print">("edit");
  const [isDownloading, setIsDownloading] = useState<string | null>(null);
  const [showPrintModal, setShowPrintModal] = useState<boolean>(false);

  useEffect(() => {
    if (selectedDay > totalDays) {
      setSelectedDay(totalDays);
    }
  }, [totalDays, selectedDay]);

  const [planningMode, setPlanningMode] = useState<"with_bumbu_10" | "all_ingredients">(() => {
    return (localStorage.getItem("sisper_planning_mode") as any) || "with_bumbu_10";
  });

  const handlePlanningModeChange = (mode: "with_bumbu_10" | "all_ingredients") => {
    setPlanningMode(mode);
    localStorage.setItem("sisper_planning_mode", mode);
  };

  const [localBumbuPct, setLocalBumbuPct] = useState<number>(() => {
    const saved = Number(localStorage.getItem("sisper_bumbu_pct"));
    return !isNaN(saved) && saved >= 0 && saved <= 100 ? saved : 10;
  });

  const activeBumbuPct = propBumbuPct !== undefined ? propBumbuPct : localBumbuPct;

  const handleBumbuPctChange = (val: number) => {
    const clamped = Math.max(0, Math.min(100, isNaN(val) ? 0 : val));
    setLocalBumbuPct(clamped);
    localStorage.setItem("sisper_bumbu_pct", String(clamped));
    if (propSetBumbuPct) propSetBumbuPct(clamped);
  };

  // State for Calculator Modal
  const [confirmModal, setConfirmModal] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    onConfirm: () => void;
    confirmText?: string;
    cancelText?: string;
    variant?: "danger" | "warning" | "info";
  }>({
    isOpen: false,
    title: "",
    message: "",
    onConfirm: () => {},
  });

  const triggerConfirm = ({
    title,
    message,
    onConfirm,
    confirmText = "Hapus",
    cancelText = "Batal",
    variant = "danger"
  }: {
    title: string;
    message: string;
    onConfirm: () => void;
    confirmText?: string;
    cancelText?: string;
    variant?: "danger" | "warning" | "info";
  }) => {
    setConfirmModal({
      isOpen: true,
      title,
      message,
      onConfirm: () => {
        onConfirm();
        setConfirmModal(prev => ({ ...prev, isOpen: false }));
      },
      confirmText,
      cancelText,
      variant
    });
  };

  const [activeCalcField, setActiveCalcField] = useState<{
    porsi: "besar" | "kecil" | string; // If custom table, it is the tableId
    rowIndex: number;
    field: "potong" | "ekor" | "buah" | "butir";
    initialValue: string;
    ingredientName?: string;
  } | null>(null);

  const [calcExpr, setCalcExpr] = useState<string>("");
  const [calcResult, setCalcResult] = useState<string>("");

  // State for Buffer & Jumlah Config Modal
  const [activeBufferConfig, setActiveBufferConfig] = useState<{
    porsi: "besar" | "kecil" | string; // If custom table, it is the tableId
    rowIndex: number;
    namaBahan: string;
    bufferBase: string;
    bufferCustomVal?: string;
    jumlahBufferChoice: string;
    jumlahBufferCustomVal?: string;
  } | null>(null);

  const openBufferConfig = (porsi: "besar" | "kecil" | string, rowIndex: number, b: any) => {
    setActiveBufferConfig({
      porsi,
      rowIndex,
      namaBahan: b.nama || b.namaBahan || "Bahan Makanan",
      bufferBase: b.bufferBase || "auto",
      bufferCustomVal: b.bufferCustomVal || "",
      jumlahBufferChoice: b.jumlahBufferChoice || "auto",
      jumlahBufferCustomVal: b.jumlahBufferCustomVal || ""
    });
  };

  const saveBufferConfig = (config: typeof activeBufferConfig) => {
    if (!config) return;
    const { porsi, rowIndex, bufferBase, bufferCustomVal, jumlahBufferChoice, jumlahBufferCustomVal } = config;

    const updates = {
      bufferBase: bufferBase as any,
      bufferCustomVal,
      jumlahBufferChoice,
      jumlahBufferCustomVal
    };

    if (porsi === "besar" || porsi === "kecil") {
      handleRowBatchChange(porsi, rowIndex, updates);
    } else if (porsi.startsWith("cat:")) {
      editIngredientInCategoryEntryBatch(porsi.slice(4), rowIndex, updates);
    } else {
      editIngredientInCustomTableBatch(porsi, rowIndex, updates);
    }
    setActiveBufferConfig(null);
  };

  const safeEvalMath = (expr: string): string => {
    try {
      // Sanitasi input: hanya boleh angka, operator (+ - * / . ( )), spasi
      const sanitized = expr.replace(/[^0-9+\-*/().\s]/g, "");
      if (!sanitized.trim()) return "";
      
      // Gunakan Function constructor untuk kalkulasi ekspresi matematika ter-sanitasi secara aman
      const evaluated = new Function(`return (${sanitized})`)();
      if (typeof evaluated === "number" && !isNaN(evaluated) && isFinite(evaluated)) {
        // Bulatkan ke desimal yang wajar jika panjang desimalnya terlalu banyak
        return Number(evaluated.toFixed(3)).toString();
      }
      return "";
    } catch (error) {
      return "";
    }
  };

  const openCalculator = (
    porsi: "besar" | "kecil" | string,
    rowIndex: number,
    field: "potong" | "ekor" | "buah" | "butir",
    currentVal: string,
    ingredientName?: string
  ) => {
    setActiveCalcField({ porsi, rowIndex, field, initialValue: currentVal, ingredientName });
    setCalcExpr(currentVal || "");
    const initialRes = safeEvalMath(currentVal);
    setCalcResult(initialRes || currentVal || "");
  };

  const handleCalcKeyPress = (key: string) => {
    if (key === "C") {
      setCalcExpr("");
      setCalcResult("");
    } else if (key === "Del") {
      const newVal = calcExpr.slice(0, -1);
      setCalcExpr(newVal);
      const res = safeEvalMath(newVal);
      setCalcResult(res || newVal || "");
    } else if (key === "=") {
      const res = safeEvalMath(calcExpr);
      if (res) {
        setCalcExpr(res);
        setCalcResult(res);
      }
    } else {
      const newVal = calcExpr + key;
      setCalcExpr(newVal);
      const res = safeEvalMath(newVal);
      setCalcResult(res || newVal || "");
    }
  };

  const saveCalculatorResult = (valToSave: string) => {
    if (!activeCalcField) return;
    const { porsi, rowIndex, field } = activeCalcField;
    const finalVal = valToSave === "" ? undefined : (isNaN(Number(valToSave)) ? valToSave : Number(valToSave));

    if (porsi === "besar" || porsi === "kecil") {
      handleRowChange(porsi, rowIndex, field, finalVal);
    } else if (porsi.startsWith("cat:")) {
      editIngredientInCategoryEntry(porsi.slice(4), rowIndex, field, finalVal);
    } else {
      editIngredientInCustomTable(porsi, rowIndex, field, finalVal);
    }
    setActiveCalcField(null);
  };

  // Keyboard support for Food Cost calculator modal (laptop/tablet keyboard typing)
  useEffect(() => {
    if (!activeCalcField) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      const targetTag = (e.target as HTMLElement)?.tagName?.toLowerCase();
      if (targetTag === "input" || targetTag === "textarea") {
        if (e.key === "Escape") {
          setActiveCalcField(null);
          e.preventDefault();
        } else if (e.key === "Enter") {
          saveCalculatorResult(calcResult || calcExpr);
          e.preventDefault();
        }
        return;
      }

      const key = e.key;
      if (key >= "0" && key <= "9") {
        handleCalcKeyPress(key);
        e.preventDefault();
      } else if (key === "+" || key === "-" || key === "*" || key === "/" || key === "(" || key === ")" || key === ".") {
        handleCalcKeyPress(key);
        e.preventDefault();
      } else if (key === "Enter") {
        saveCalculatorResult(calcResult || calcExpr);
        e.preventDefault();
      } else if (key === "=") {
        handleCalcKeyPress("=");
        e.preventDefault();
      } else if (key === "Backspace") {
        handleCalcKeyPress("Del");
        e.preventDefault();
      } else if (key === "Escape") {
        setActiveCalcField(null);
        e.preventDefault();
      } else if (key.toLowerCase() === "c") {
        handleCalcKeyPress("C");
        e.preventDefault();
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => {
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [activeCalcField, calcExpr, calcResult]);

  // PM & Harga per porsi settings synced directly from menu Penerima Manfaat
  const settings = pmSettings || {
    porsiKecilHarga: 8000,
    porsiBesarHarga: 10000,
    porsiKecilSasaranIds: ["tk_paud_lb", "sd_kelas_1_3", "anak_balita", "anak_balita_13_59", "balita_6_11", "sampel"],
    porsiBesarSasaranIds: ["sd_kelas_4_6", "smp_mts_smplb", "sma_smk_ma", "pendidik", "tenaga_kependidikan", "ibu_hamil", "ibu_menyusui", "sampel"]
  };

  const defaultKecilIds: string[] = Array.isArray(settings.porsiKecilSasaranIds)
    ? settings.porsiKecilSasaranIds
    : ["tk_paud_lb", "sd_kelas_1_3", "anak_balita", "anak_balita_13_59", "balita_6_11"];
  const defaultBesarIds: string[] = Array.isArray(settings.porsiBesarSasaranIds)
    ? settings.porsiBesarSasaranIds
    : ["sd_kelas_4_6", "smp_mts_smplb", "sma_smk_ma", "pendidik", "tenaga_kependidikan", "ibu_hamil", "ibu_menyusui"];

  // Helper to resolve default porsi & label for any sasaran ID (Standar & Alergi)
  const resolveDefaultPorsiForSasaran = (sId: string): "besar" | "kecil" =>
    resolveCategoryPorsiDefault(sId, defaultKecilIds, defaultBesarIds);

  const resolveStaticSasaranLabel = (sId: string): string => {
    if (sId === "sampel") return "Sampel";
    if (sId === "alergi_sampel") return "Alergi · Sampel";
    return ALL_KATEGORI_WITH_ALERGI.find((s) => s.id === sId)?.label || sId;
  };

  const buildDefaultAlergiGroupsForTabGroup = (
    tabGroup: "sekolah" | "tigaB" | "custom"
  ): FoodCostCategoryGroup[] => {
    const baseIds =
      tabGroup === "sekolah"
        ? SEKOLAH_SASARAN_IDS
        : tabGroup === "tigaB"
          ? [...TIGAB_SASARAN_IDS, ...MPASI_SASARAN_IDS]
          : SASARAN_KATEGORI_LIST.map((s) => s.id);
    const kecilAlergiIds = baseIds
      .filter((id) => resolveDefaultPorsiForSasaran(id) === "kecil")
      .map((id) => `alergi_${id}`);
    const besarAlergiIds = baseIds
      .filter((id) => resolveDefaultPorsiForSasaran(id) === "besar")
      .map((id) => `alergi_${id}`);
    const groups: FoodCostCategoryGroup[] = [];
    if (kecilAlergiIds.length > 0) {
      groups.push({
        id: `cat_alergi_kecil_${tabGroup}`,
        title:
          tabGroup === "sekolah"
            ? "Alergi · Sekolah (Porsi Kecil)"
            : tabGroup === "tigaB"
              ? "Alergi · Balita / 3B (Porsi Kecil)"
              : "Alergi · Porsi Kecil",
        porsi: "kecil",
        sasaranIds: kecilAlergiIds,
      });
    }
    if (besarAlergiIds.length > 0) {
      groups.push({
        id: `cat_alergi_besar_${tabGroup}`,
        title:
          tabGroup === "sekolah"
            ? "Alergi · Sekolah (Porsi Besar)"
            : tabGroup === "tigaB"
              ? "Alergi · Bumil & Busui (Porsi Besar)"
              : "Alergi · Porsi Besar",
        porsi: "besar",
        sasaranIds: besarAlergiIds,
      });
    }
    return groups;
  };

  const buildDefaultGroupsForTabGroup = (tabGroup: "sekolah" | "tigaB" | "custom"): FoodCostCategoryGroup[] => {
    const ids =
      tabGroup === "sekolah"
        ? SEKOLAH_SASARAN_IDS
        : tabGroup === "tigaB"
          ? [...TIGAB_SASARAN_IDS, ...MPASI_SASARAN_IDS]
          : SASARAN_KATEGORI_LIST.map((s) => s.id);
    const standardGroups: FoodCostCategoryGroup[] = ids.map((sId) => ({
      id: `cat_${sId}`,
      title: resolveStaticSasaranLabel(sId),
      porsi: resolveDefaultPorsiForSasaran(sId),
      sasaranIds: [sId],
    }));
    return [...standardGroups, ...buildDefaultAlergiGroupsForTabGroup(tabGroup)];
  };

  const ensureAlergiInSavedGroups = (
    groups: FoodCostCategoryGroup[],
    tabGroup: "sekolah" | "tigaB" | "custom"
  ): FoodCostCategoryGroup[] => {
    const hasAlergi = groups.some((g) => (g.sasaranIds || []).some((id) => isAlergiSasaranId(id)));
    if (hasAlergi) return groups;
    return [...groups, ...buildDefaultAlergiGroupsForTabGroup(tabGroup)];
  };

  // Food Cost display scope, per-category / merged-category grouping & confirmation workflow
  const [appliedConfig, setAppliedConfig] = useState<{
    confirmed: boolean;
    layoutMode: "per_kategori" | "gabung_porsi";
    scope: "kecil" | "besar" | "all";
    kecilIds: string[];
    besarIds: string[];
    categoryGroupsByGroup: Record<"sekolah" | "tigaB" | "custom", FoodCostCategoryGroup[]>;
  }>(() => {
    const defaultGroupsByGroup: Record<"sekolah" | "tigaB" | "custom", FoodCostCategoryGroup[]> = {
      sekolah: buildDefaultGroupsForTabGroup("sekolah"),
      tigaB: buildDefaultGroupsForTabGroup("tigaB"),
      custom: buildDefaultGroupsForTabGroup("custom"),
    };
    try {
      const savedV2 = localStorage.getItem("sppg_fc_applied_config_v2");
      if (savedV2) {
        const parsed = JSON.parse(savedV2);
        const needAlergiMigration = !localStorage.getItem("sppg_fc_alergi_groups_added_v1");
        if (needAlergiMigration) {
          localStorage.setItem("sppg_fc_alergi_groups_added_v1", "1");
        }
        const hydrateGroups = (
          raw: FoodCostCategoryGroup[] | undefined,
          key: "sekolah" | "tigaB" | "custom"
        ) => {
          const base =
            Array.isArray(raw) && raw.length > 0 ? raw : defaultGroupsByGroup[key];
          return needAlergiMigration ? ensureAlergiInSavedGroups(base, key) : base;
        };
        return {
          confirmed: parsed.confirmed !== undefined ? Boolean(parsed.confirmed) : true,
          layoutMode: parsed.layoutMode === "gabung_porsi" ? "gabung_porsi" : "per_kategori",
          scope: parsed.scope === "kecil" || parsed.scope === "besar" || parsed.scope === "all" ? parsed.scope : "all",
          kecilIds: Array.isArray(parsed.kecilIds) ? parsed.kecilIds : defaultKecilIds,
          besarIds: Array.isArray(parsed.besarIds) ? parsed.besarIds : defaultBesarIds,
          categoryGroupsByGroup: {
            sekolah: hydrateGroups(parsed.categoryGroupsByGroup?.sekolah, "sekolah"),
            tigaB: hydrateGroups(parsed.categoryGroupsByGroup?.tigaB, "tigaB"),
            custom: hydrateGroups(parsed.categoryGroupsByGroup?.custom, "custom"),
          },
        };
      }
      const savedV1 = localStorage.getItem("sppg_fc_applied_config_v1");
      if (savedV1) {
        const parsed = JSON.parse(savedV1);
        return {
          confirmed: true,
          layoutMode: "per_kategori",
          scope: parsed.scope === "kecil" || parsed.scope === "besar" || parsed.scope === "all" ? parsed.scope : "all",
          kecilIds: Array.isArray(parsed.kecilIds) ? parsed.kecilIds : defaultKecilIds,
          besarIds: Array.isArray(parsed.besarIds) ? parsed.besarIds : defaultBesarIds,
          categoryGroupsByGroup: defaultGroupsByGroup,
        };
      }
    } catch {
      // ignore corrupted storage
    }
    return {
      confirmed: true,
      layoutMode: "per_kategori",
      scope: "all",
      kecilIds: defaultKecilIds,
      besarIds: defaultBesarIds,
      categoryGroupsByGroup: defaultGroupsByGroup,
    };
  });

  const [draftLayoutMode, setDraftLayoutMode] = useState<"per_kategori" | "gabung_porsi">(appliedConfig.layoutMode);
  const [draftScope, setDraftScope] = useState<"kecil" | "besar" | "all">(appliedConfig.scope);
  const [draftKecilIds, setDraftKecilIds] = useState<string[]>(appliedConfig.kecilIds);
  const [draftBesarIds, setDraftBesarIds] = useState<string[]>(appliedConfig.besarIds);
  const [draftCategoryGroupsByGroup, setDraftCategoryGroupsByGroup] = useState<
    Record<"sekolah" | "tigaB" | "custom", FoodCostCategoryGroup[]>
  >(appliedConfig.categoryGroupsByGroup);
  const [mergeSelectedGroupIds, setMergeSelectedGroupIds] = useState<string[]>([]);
  const [mergeCustomTitle, setMergeCustomTitle] = useState<string>("");
  const [mergeCustomPorsi, setMergeCustomPorsi] = useState<"auto" | "kecil" | "besar">("auto");
  const [showCrossGroupPicker, setShowCrossGroupPicker] = useState<boolean>(false);
  const [isSelectorExpanded, setIsSelectorExpanded] = useState<boolean>(true);
  const [showCategoryConfirmModal, setShowCategoryConfirmModal] = useState<boolean>(false);
  const [autoFillOnConfirm, setAutoFillOnConfirm] = useState<boolean>(true);
  const [fcNotice, setFcNotice] = useState<{ type: "info" | "success" | "warning"; text: string } | null>(null);

  // Copy & Paste Food Cost clipboard (to quickly reuse all ingredients & descriptions across categories/days and just change BB)
  const [copiedFoodCost, setCopiedFoodCost] = useState<CopiedFoodCostPayload | null>(() => {
    try {
      const saved = localStorage.getItem("sppg_fc_clipboard_v1");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && Array.isArray(parsed.bahanList) && parsed.bahanList.length > 0) {
          return parsed;
        }
      }
    } catch {
      // ignore corrupted clipboard
    }
    return null;
  });

  useEffect(() => {
    if (copiedFoodCost) {
      localStorage.setItem("sppg_fc_clipboard_v1", JSON.stringify(copiedFoodCost));
    } else {
      localStorage.removeItem("sppg_fc_clipboard_v1");
    }
  }, [copiedFoodCost]);

  const [pastedHighlightMap, setPastedHighlightMap] = useState<
    Record<string, { from: string; count: number }>
  >({});

  const [pasteConfirmModal, setPasteConfirmModal] = useState<{
    isOpen: boolean;
    targetKey: string;
    targetLabel: string;
    existingCount: number;
    onReplace: () => void;
    onAppend: () => void;
  } | null>(null);

  useEffect(() => {
    localStorage.setItem("sppg_fc_applied_config_v2", JSON.stringify(appliedConfig));
  }, [appliedConfig]);

  // Keep selected category IDs synchronized with PM classification changes
  useEffect(() => {
    setDraftKecilIds((prev) => {
      const valid = prev.filter((id) => defaultKecilIds.includes(id));
      return valid.length > 0 ? valid : defaultKecilIds;
    });
    setDraftBesarIds((prev) => {
      const valid = prev.filter((id) => defaultBesarIds.includes(id));
      return valid.length > 0 ? valid : defaultBesarIds;
    });
    setAppliedConfig((prev) => {
      const nextKecil = prev.kecilIds.filter((id) => defaultKecilIds.includes(id));
      const nextBesar = prev.besarIds.filter((id) => defaultBesarIds.includes(id));
      const resolvedKecil = nextKecil.length > 0 ? nextKecil : defaultKecilIds;
      const resolvedBesar = nextBesar.length > 0 ? nextBesar : defaultBesarIds;
      if (
        resolvedKecil.length === prev.kecilIds.length &&
        resolvedBesar.length === prev.besarIds.length &&
        resolvedKecil.every((v, i) => v === prev.kecilIds[i]) &&
        resolvedBesar.every((v, i) => v === prev.besarIds[i])
      ) {
        return prev;
      }
      return { ...prev, kecilIds: resolvedKecil, besarIds: resolvedBesar };
    });
  }, [settings.porsiKecilSasaranIds, settings.porsiBesarSasaranIds]);

  const handleSelectDraftScope = (nextScope: "kecil" | "besar" | "all") => {
    setDraftScope(nextScope);
    if (nextScope === "kecil") {
      setDraftKecilIds(defaultKecilIds);
    } else if (nextScope === "besar") {
      setDraftBesarIds(defaultBesarIds);
    }
  };

  const toggleDraftCategory = (porsi: "kecil" | "besar", sasaranId: string) => {
    if (porsi === "kecil") {
      setDraftKecilIds((prev) =>
        prev.includes(sasaranId) ? prev.filter((id) => id !== sasaranId) : [...prev, sasaranId]
      );
    } else {
      setDraftBesarIds((prev) =>
        prev.includes(sasaranId) ? prev.filter((id) => id !== sasaranId) : [...prev, sasaranId]
      );
    }
  };

  // Over-budget decision tracking ("Tetap Lanjutkan" vs "Ubah Menu")
  const [overBudgetDecisions, setOverBudgetDecisions] = useState<
    Record<string, { cost: number; action: "continue" }>
  >({});
  const [changeMenuModal, setChangeMenuModal] = useState<{
    isOpen: boolean;
    scope: string;
    label: string;
    actualCost: number;
    targetPrice: number;
    items: { id: string; nama: string; costPerPortion: number }[];
  } | null>(null);
  const [changeMenuNameInput, setChangeMenuNameInput] = useState<string>("");

  // Custom tables state
  const [customTables, setCustomTables] = useState<any[]>(() => {
    const saved = localStorage.getItem("sppg_custom_fc_tables_v1");
    return saved ? JSON.parse(saved) : [];
  });

  React.useEffect(() => {
    localStorage.setItem("sppg_custom_fc_tables_v1", JSON.stringify(customTables));
  }, [customTables]);

  // Kriteria PM labels state
  const [customKriteriaLabels, setCustomKriteriaLabels] = useState<Record<string, string>>(() => {
    const saved = localStorage.getItem("sppg_custom_kriteria_labels");
    return saved ? JSON.parse(saved) : {};
  });

  const saveCustomKriteriaLabel = (key: string, value: string) => {
    const updated = { ...customKriteriaLabels, [key]: value };
    setCustomKriteriaLabels(updated);
    localStorage.setItem("sppg_custom_kriteria_labels", JSON.stringify(updated));
  };

  // Custom tables action handlers
  const addIngredientToCustomTable = (tableId: string) => {
    const newBahan: BahanMakananInput = {
      id: `custom_row_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
      tkpiId: "",
      beratBB: 0,
      urt: "",
      hargaSatuan: 0
    };
    setCustomTables(prev => prev.map(t => {
      if (t.id === tableId) {
        return { ...t, bahanList: [...t.bahanList, newBahan] };
      }
      return t;
    }));
  };

  const removeIngredientFromCustomTable = (tableId: string, rowId: string, index: number) => {
    setCustomTables(prev => {
      const updated = prev.map(t => {
        if (t.id === tableId) {
          const updatedList = t.bahanList.filter((b: any, idx: number) => {
            if (rowId && b.id && b.id === rowId) return false;
            if (idx === index) return false;
            return true;
          });
          return { ...t, bahanList: updatedList };
        }
        return t;
      });
      localStorage.setItem("sppg_custom_fc_tables_v1", JSON.stringify(updated));
      return updated;
    });
  };

  const editIngredientInCustomTable = (tableId: string, rowIndex: number, field: keyof BahanMakananInput, value: any) => {
    setCustomTables(prev => prev.map(t => {
      if (t.id === tableId) {
        const updatedList = [...t.bahanList];
        updatedList[rowIndex] = { ...updatedList[rowIndex], [field]: value };
        return { ...t, bahanList: updatedList };
      }
      return t;
    }));
  };

  const editIngredientInCustomTableBatch = (tableId: string, rowIndex: number, updates: Partial<BahanMakananInput>) => {
    setCustomTables(prev => prev.map(t => {
      if (t.id === tableId) {
        const updatedList = [...t.bahanList];
        updatedList[rowIndex] = { ...updatedList[rowIndex], ...updates };
        return { ...t, bahanList: updatedList };
      }
      return t;
    }));
  };

  const editCustomTableMeta = (tableId: string, field: string, value: any) => {
    setCustomTables(prev => prev.map(t => {
      if (t.id === tableId) {
        return { ...t, [field]: value };
      }
      return t;
    }));
  };

  const deleteCustomTable = (tableId: string) => {
    triggerConfirm({
      title: "Hapus Tabel Komponen Kustom",
      message: "Apakah Anda yakin ingin menghapus tabel komponen kustom ini? Seluruh isi data bahan di dalamnya akan hilang secara permanen.",
      onConfirm: () => {
        setCustomTables(prev => {
          const updated = prev.filter(t => t.id !== tableId);
          localStorage.setItem("sppg_custom_fc_tables_v1", JSON.stringify(updated));
          return updated;
        });
      }
    });
  };

  // Customizable column headers
  const [headerPotong, setHeaderPotong] = useState(() => localStorage.getItem("fc_header_potong") || "Potong");
  const [headerEkor, setHeaderEkor] = useState(() => localStorage.getItem("fc_header_ekor") || "Ekor");
  const [headerBuah, setHeaderBuah] = useState(() => localStorage.getItem("fc_header_buah") || "Buah");
  const [headerButir, setHeaderButir] = useState(() => localStorage.getItem("fc_header_butir") || "Butir");

  // Selected AKG reference templates
  const [besarAkgType, setBesarAkgType] = useState<string>("sd_besar");
  const [kecilAkgType, setKecilAkgType] = useState<string>("sd_kecil");

  // Selected AKG reference modes ("min" | "max" | "avg" | "custom" | "")
  const [besarAkgMode, setBesarAkgMode] = useState<string>(() => {
    const saved = localStorage.getItem("sppg_besar_akg_mode");
    return saved !== null ? saved : "min";
  });
  const [kecilAkgMode, setKecilAkgMode] = useState<string>(() => {
    const saved = localStorage.getItem("sppg_kecil_akg_mode");
    return saved !== null ? saved : "min";
  });

  // Per-Nutrient AKG reference modes (energi, protein, lemak, kh, serat)
  const [besarNutrientModes, setBesarNutrientModes] = useState<Record<string, string>>(() => {
    const saved = localStorage.getItem("sppg_besar_nutrient_modes");
    return saved ? JSON.parse(saved) : { energi: "min", protein: "min", lemak: "min", kh: "min", serat: "min" };
  });
  const [kecilNutrientModes, setKecilNutrientModes] = useState<Record<string, string>>(() => {
    const saved = localStorage.getItem("sppg_kecil_nutrient_modes");
    return saved ? JSON.parse(saved) : { energi: "min", protein: "min", lemak: "min", kh: "min", serat: "min" };
  });

  React.useEffect(() => {
    localStorage.setItem("sppg_besar_akg_mode", besarAkgMode);
  }, [besarAkgMode]);

  React.useEffect(() => {
    localStorage.setItem("sppg_kecil_akg_mode", kecilAkgMode);
  }, [kecilAkgMode]);

  React.useEffect(() => {
    localStorage.setItem("sppg_besar_nutrient_modes", JSON.stringify(besarNutrientModes));
  }, [besarNutrientModes]);

  React.useEffect(() => {
    localStorage.setItem("sppg_kecil_nutrient_modes", JSON.stringify(kecilNutrientModes));
  }, [kecilNutrientModes]);

  const updateIndividualNutrientMode = (porsi: "besar" | "kecil", key: string, mode: string) => {
    if (porsi === "besar") {
      setBesarNutrientModes(prev => ({ ...prev, [key]: mode }));
    } else {
      setKecilNutrientModes(prev => ({ ...prev, [key]: mode }));
    }
  };

  // Custom values for "Custom" option
  const [customBesar, setCustomBesar] = useState<Record<string, number>>(() => {
    const saved = localStorage.getItem("sppg_custom_besar");
    return saved ? JSON.parse(saved) : { energi: 634, protein: 15, lemak: 18, kh: 85, serat: 8 };
  });
  const [customKecil, setCustomKecil] = useState<Record<string, number>>(() => {
    const saved = localStorage.getItem("sppg_custom_kecil");
    return saved ? JSON.parse(saved) : { energi: 371, protein: 8, lemak: 11, kh: 55, serat: 5 };
  });

  React.useEffect(() => {
    localStorage.setItem("sppg_custom_besar", JSON.stringify(customBesar));
  }, [customBesar]);

  React.useEffect(() => {
    localStorage.setItem("sppg_custom_kecil", JSON.stringify(customKecil));
  }, [customKecil]);

  // Synchronize target templates based on selectedGroup
  React.useEffect(() => {
    if (selectedGroup === "sekolah") {
      setBesarAkgType("sd_besar");
      setKecilAkgType("sd_kecil");
    } else {
      setBesarAkgType("bumil");
      setKecilAkgType("balita");
    }
  }, [selectedGroup]);

  const getAkgTarget = (porsi: "besar" | "kecil", key: string): number => {
    const isBesar = porsi === "besar";
    const type = isBesar ? besarAkgType : kecilAkgType;
    const modesMap = isBesar ? besarNutrientModes : kecilNutrientModes;
    const mode = modesMap[key] || (isBesar ? besarAkgMode : kecilAkgMode) || "min";

    if (!mode || mode === "") {
      return 0;
    }

    if (type === "custom" || mode === "custom") {
      const customMap = isBesar ? customBesar : customKecil;
      return customMap[key] || 0;
    }

    const config = TARGET_AKG_LIMITS[type];
    if (!config) return 0;

    const minVal = key === "energi" ? config.energiMin : (config as any)[`${key}Min`] || 0;
    const ratio = config.energiMax / config.energiMin;

    if (mode === "min") {
      return minVal;
    }
    if (mode === "max") {
      if (key === "energi") return config.energiMax;
      return Math.round(minVal * ratio * 10) / 10;
    }
    if (mode === "avg") {
      if (key === "energi") {
        return (config.energiMin + config.energiMax) / 2;
      }
      const maxVal = minVal * ratio;
      return Math.round(((minVal + maxVal) / 2) * 10) / 10;
    }
    return 0;
  };

  const handleAkgValChange = (
    porsi: "besar" | "kecil",
    key: string,
    rawVal: number,
    currentMode: string,
    akgTypeFromTable?: string,
    tableId?: string
  ) => {
    const val = isNaN(rawVal) ? 0 : rawVal;
    const isBesar = porsi === "besar";
    const setCustom = isBesar ? setCustomBesar : setCustomKecil;

    if (tableId) {
      editCustomTableMeta(tableId, "akgMode", "custom");
    } else {
      if (isBesar) {
        setBesarAkgMode("custom");
      } else {
        setKecilAkgMode("custom");
      }
    }

    if (currentMode === "custom") {
      setCustom(prev => ({ ...prev, [key]: val }));
    } else {
      const keys = ["energi", "protein", "lemak", "kh", "serat"];
      const newCustom: Record<string, number> = {};
      const currentCustomMap = isBesar ? customBesar : customKecil;

      keys.forEach(k => {
        if (k === key) {
          newCustom[k] = val;
        } else {
          let displayedVal = 0;
          const type = akgTypeFromTable || (isBesar ? besarAkgType : kecilAkgType);
          const config = TARGET_AKG_LIMITS[type];
          if (config && currentMode !== "") {
            const minVal = k === "energi" ? config.energiMin : (config as any)[`${k}Min`] || 0;
            const ratio = config.energiMax / config.energiMin;

            if (currentMode === "min") {
              displayedVal = minVal;
            } else if (currentMode === "max") {
              if (k === "energi") displayedVal = config.energiMax;
              else displayedVal = Math.round(minVal * ratio * 10) / 10;
            } else if (currentMode === "avg") {
              if (k === "energi") {
                displayedVal = (config.energiMin + config.energiMax) / 2;
              } else {
                const maxVal = minVal * ratio;
                displayedVal = Math.round(((minVal + maxVal) / 2) * 10) / 10;
              }
            }
          } else {
            displayedVal = currentCustomMap[k] || 0;
          }
          newCustom[k] = displayedVal;
        }
      });
      setCustom(newCustom);
    }
  };

  // Force selectedType to MP-ASI if threeB group is active and user wants it, or vice versa
  React.useEffect(() => {
    if (selectedGroup === "sekolah" && selectedType === "MP-ASI") {
      setSelectedType("Basah");
    }
  }, [selectedGroup, selectedType]);

  // Compute database category representation
  const getDbType = (group: "sekolah" | "tigaB" | "custom", type: "Basah" | "Alergi" | "Kering" | "MP-ASI") => {
    if (group === "sekolah") {
      return type === "MP-ASI" ? "Basah" : type;
    } else {
      if (type === "MP-ASI") return "MP-ASI";
      return `3B-${type}` as FoodCostDay["jenisMenu"];
    }
  };

  const activeDbType = getDbType(selectedGroup, selectedType);

  // Get active Day's data
  const dayIndex = selectedDay - 1;
  const currentDayData = (foodCostDays.find((d) => d.hariKe === selectedDay && d.jenisMenu === activeDbType) || {
    hariKe: selectedDay,
    jenisMenu: activeDbType,
    porsiBesarBahan: [],
    porsiKecilBahan: [],
    bufferPct: 5
  }) as FoodCostDay;

  // Get recipient counts based on selected group and type from harianPM
  const dayCounts = getCountsForDay(harianPM, selectedDay);

  const dayPMForLabels = harianPM.find(h => h.hariKe === selectedDay) || harianPM[0] || { sasaran: [] };

  const getSasaranLabel = (sId: string): string => {
    if (sId === "sampel") return "Sampel";
    if (sId === "alergi_sampel") return "Alergi · Sampel";
    const cleanId = stripAlergiPrefix(sId);
    const fromDay = dayPMForLabels.sasaran.find((s) => s.id === cleanId)?.label;
    if (fromDay) return isAlergiSasaranId(sId) ? `Alergi · ${fromDay}` : fromDay;
    const fromConst = ALL_KATEGORI_WITH_ALERGI.find((s) => s.id === sId)?.label;
    return fromConst || sId;
  };

  const sharedPortionIds = React.useMemo(
    () =>
      Array.from(
        new Set([
          "sampel",
          ...defaultKecilIds.filter((id) => defaultBesarIds.includes(id)),
        ])
      ),
    [defaultKecilIds, defaultBesarIds]
  );

  const getSasaranPmCount = (sId: string, porsi: "besar" | "kecil"): number =>
    resolveSasaranPmFromTable2(
      dayPMForLabels.sasaran,
      sId,
      porsi,
      selectedType === "Alergi",
      sharedPortionIds
    );

  const activeKecilIds = (appliedConfig.confirmed ? appliedConfig.kecilIds : draftKecilIds).filter((id) =>
    defaultKecilIds.includes(id)
  );
  const activeBesarIds = (appliedConfig.confirmed ? appliedConfig.besarIds : draftBesarIds).filter((id) =>
    defaultBesarIds.includes(id)
  );

  const isPerKategoriMode = appliedConfig.layoutMode === "per_kategori";
  const activeGroupKey: "sekolah" | "tigaB" | "custom" = selectedGroup;

  const getEntryAutoPmCount = (sasaranIds: string[], porsi: "besar" | "kecil"): number => {
    return (sasaranIds || []).reduce((sum, sId) => sum + getSasaranPmCount(sId, porsi), 0);
  };

  const getEntryEffectivePmCount = (entry: FoodCostCategoryEntry): number => {
    if (entry.customPmCount !== undefined && entry.customPmCount !== null) {
      return Math.max(0, Number(entry.customPmCount) || 0);
    }
    return getEntryAutoPmCount(entry.sasaranIds, entry.porsi);
  };

  const activeCategoryGroupsRaw =
    appliedConfig.categoryGroupsByGroup?.[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey);

  const activeCategoryGroups = activeCategoryGroupsRaw.filter((g) => {
    if (appliedConfig.scope !== "all" && g.porsi !== appliedConfig.scope) return false;
    if (selectedGroup === "tigaB" && selectedType === "MP-ASI") {
      const hasMpAsiGroup = activeCategoryGroupsRaw.some((x) =>
        x.sasaranIds.some((id) => MPASI_SASARAN_IDS.includes(id))
      );
      if (hasMpAsiGroup) {
        return g.sasaranIds.some((id) => MPASI_SASARAN_IDS.includes(id));
      }
    }
    return true;
  });

  const resolvedCategoryEntries: FoodCostCategoryEntry[] = activeCategoryGroups.map((g) => {
    const existingList = currentDayData.categoryEntries || [];
    const existing =
      existingList.find((e) => e.id === g.id) ||
      (g.sasaranIds.length === 1
        ? existingList.find((e) => e.sasaranIds.length === 1 && e.sasaranIds[0] === g.sasaranIds[0])
        : undefined);

    if (existing) {
      return {
        ...existing,
        id: g.id,
        title: g.title,
        porsi: g.porsi,
        sasaranIds: g.sasaranIds,
        bahanList: existing.bahanList || [],
        bufferPct: existing.bufferPct !== undefined ? existing.bufferPct : (currentDayData.bufferPct ?? 5),
        akgType: existing.akgType || getDefaultAkgTypeForGroup(g.sasaranIds, g.porsi),
        akgMode: existing.akgMode || "min",
        nutrientModes: existing.nutrientModes || {
          energi: "min",
          protein: "min",
          lemak: "min",
          kh: "min",
          serat: "min",
        },
        kriteriaPm: existing.kriteriaPm !== undefined ? existing.kriteriaPm : g.title,
      };
    }

    return {
      id: g.id,
      title: g.title,
      porsi: g.porsi,
      sasaranIds: g.sasaranIds,
      bahanList: [],
      bufferPct: currentDayData.bufferPct ?? 5,
      akgType: getDefaultAkgTypeForGroup(g.sasaranIds, g.porsi),
      akgMode: "min",
      nutrientModes: {
        energi: "min",
        protein: "min",
        lemak: "min",
        kh: "min",
        serat: "min",
      },
      kriteriaPm: g.title,
    };
  });

  // Helper to compute dynamic combined/merged count and target price based on chosen categories
  const getPmCountAndTargetPrice = (porsi: "besar" | "kecil") => {
    const targetPrice = porsi === "besar" ? settings.porsiBesarHarga : settings.porsiKecilHarga;

    if (isPerKategoriMode) {
      const matchingEntries = resolvedCategoryEntries.filter((e) => e.porsi === porsi);
      const count = matchingEntries.reduce((sum, e) => sum + getEntryEffectivePmCount(e), 0);
      return { count, targetPrice };
    }

    const targetIds = porsi === "besar" ? activeBesarIds : activeKecilIds;
    let count = 0;
    targetIds.forEach((sId: string) => {
      count += getSasaranPmCount(sId, porsi);
    });

    // Check if there is an active custom count for today
    const currentCustom = porsi === "besar" ? currentDayData.customPmBesarCount : currentDayData.customPmKecilCount;
    if (currentCustom !== undefined) {
      count = currentCustom;
    }

    return { count, targetPrice };
  };

  const pmBesarInfo = getPmCountAndTargetPrice("besar");
  const pmKecilInfo = getPmCountAndTargetPrice("kecil");

  const pmBesar = pmBesarInfo.count;
  const pmKecil = pmKecilInfo.count;

  const activeTargetPriceBesar = pmBesarInfo.targetPrice;
  const activeTargetPriceKecil = pmKecilInfo.targetPrice;

  const selectedKecilCategories = isPerKategoriMode
    ? resolvedCategoryEntries
        .filter((e) => e.porsi === "kecil")
        .map((e) => ({ id: e.id, label: e.title, count: getEntryEffectivePmCount(e) }))
    : activeKecilIds.map((id) => ({
        id,
        label: getSasaranLabel(id),
        count: getSasaranPmCount(id, "kecil"),
      }));

  const selectedBesarCategories = isPerKategoriMode
    ? resolvedCategoryEntries
        .filter((e) => e.porsi === "besar")
        .map((e) => ({ id: e.id, label: e.title, count: getEntryEffectivePmCount(e) }))
    : activeBesarIds.map((id) => ({
        id,
        label: getSasaranLabel(id),
        count: getSasaranPmCount(id, "besar"),
      }));

  const porsiKecilLabels = selectedKecilCategories.map((c) => c.label).join(", ");
  const porsiBesarLabels = selectedBesarCategories.map((c) => c.label).join(", ");

  const showBesarTable =
    !isPerKategoriMode &&
    appliedConfig.confirmed &&
    (appliedConfig.scope === "all" || appliedConfig.scope === "besar");
  const showKecilTable =
    !isPerKategoriMode &&
    appliedConfig.confirmed &&
    (appliedConfig.scope === "all" || appliedConfig.scope === "kecil");

  const rabPorsiKecil = pmKecil * activeTargetPriceKecil;
  const rabPorsiBesar = pmBesar * activeTargetPriceBesar;

  // Match Master Menu plans to active Food Cost group & diet type
  const allPlans = normalizedMasterMenu.plans || [];
  const targetDiet =
    selectedType === "Basah"
      ? "Standar"
      : selectedType;

  const targetSubset =
    selectedGroup === "sekolah"
      ? SEKOLAH_SASARAN_IDS
      : selectedType === "MP-ASI"
        ? MPASI_SASARAN_IDS
        : ["anak_balita", "anak_balita_13_59", "ibu_hamil", "ibu_menyusui"];

  const matchingPlans =
    selectedGroup === "custom"
      ? allPlans
      : allPlans.filter((p) => p.sasaranIds.some((id) => targetSubset.includes(id)));

  const activeMatchedPlan =
    matchingPlans.find((p) => p.id === selectedPlanOverrideId) ||
    matchingPlans.find((p) => p.dietType === targetDiet) ||
    matchingPlans[0] ||
    allPlans.find((p) => p.id === currentDayData.masterMenuPlanId);

  const matchedPlanDayItem = activeMatchedPlan?.days?.[dayIndex]
    ? normalizeMenuItem(activeMatchedPlan.days[dayIndex])
    : undefined;

  let legacyMenuItem: any | undefined;
  if (selectedGroup === "sekolah") {
    const list = selectedType === "Alergi" ? masterMenu.usiaSekolahAlergi : masterMenu.usiaSekolah;
    legacyMenuItem = list?.[dayIndex];
  } else if (selectedGroup === "tigaB") {
    if (selectedType === "MP-ASI") {
      legacyMenuItem = masterMenu.mpAsi?.[dayIndex];
    } else {
      const list = selectedType === "Alergi" ? masterMenu.tigaBAlergi : masterMenu.tigaB;
      legacyMenuItem = list?.[dayIndex];
    }
  }

  const effectiveMenuItem =
    matchedPlanDayItem && composeMenuName(matchedPlanDayItem)
      ? matchedPlanDayItem
      : legacyMenuItem
        ? normalizeMenuItem(legacyMenuItem)
        : matchedPlanDayItem;

  const namaMenuHariIni =
    composeMenuName(matchedPlanDayItem) ||
    (currentDayData.namaMenu || "").trim() ||
    composeMenuName(legacyMenuItem) ||
    (selectedGroup === "custom"
      ? `Menu Kustom ${selectedType}`
      : selectedGroup === "tigaB"
        ? selectedType === "MP-ASI"
          ? "Menu MP-ASI (6-11 Bln)"
          : `Menu 3B ${selectedType}`
        : `Menu Sekolah ${selectedType}`);

  const rawDateStr = currentDayData.tanggal || scheduleDates[dayIndex] || "";
  const currentCalendarDate = rawDateStr
    ? formatScheduleIndonesianDate(rawDateStr)
    : `Hari Ke-${selectedDay}`;

  // Calculate results for the day (Gabung per Porsi mode)
  const results = calculateDay(
    currentDayData.porsiBesarBahan,
    currentDayData.porsiKecilBahan,
    pmBesar,
    pmKecil,
    currentDayData.bufferPct,
    tkpiList,
    activeBumbuPct
  );

  // Per-Category calculation results
  const categoryCalcResultsMap = React.useMemo(() => {
    const map: Record<
      string,
      {
        pmCount: number;
        targetPrice: number;
        items: CalcIngredientResult[];
        totalBahanCost: number;
        bufferCost: number;
        bumbuCost: number;
        subtotalCost: number;
        costPerPorsi: number;
        calculatedCostPerPorsi: number;
        actualTotalCost: number;
      }
    > = {};
    resolvedCategoryEntries.forEach((entry) => {
      const entryPm = getEntryEffectivePmCount(entry);
      const entryTargetPrice = entry.porsi === "besar" ? activeTargetPriceBesar : activeTargetPriceKecil;
      const entryBuffer = entry.bufferPct !== undefined ? entry.bufferPct : (currentDayData.bufferPct ?? 5);
      const calc =
        entry.porsi === "besar"
          ? calculateDay(entry.bahanList || [], [], entryPm, 0, entryBuffer, tkpiList, activeBumbuPct)
          : calculateDay([], entry.bahanList || [], 0, entryPm, entryBuffer, tkpiList, activeBumbuPct);
      const items = entry.porsi === "besar" ? calc.porsiBesarItems : calc.porsiKecilItems;
      const totalBahanCost = entry.porsi === "besar" ? calc.totalBesarBahanCost : calc.totalKecilBahanCost;
      const bufferCost = entry.porsi === "besar" ? calc.bufferBesarCost : calc.bufferKecilCost;
      const bumbuCost = entry.porsi === "besar" ? calc.bumbuBesarCost : calc.bumbuKecilCost;
      const subtotalCost = entry.porsi === "besar" ? calc.subtotalBesarCost : calc.subtotalKecilCost;
      const costPerPorsi = entry.porsi === "besar" ? calc.costPerPorsiBesar : calc.costPerPorsiKecil;
      const calculatedCostPerPorsi =
        planningMode === "all_ingredients" ? (entryPm > 0 ? totalBahanCost / entryPm : 0) : costPerPorsi;
      const actualTotalCost = planningMode === "with_bumbu_10" ? subtotalCost : totalBahanCost;

      map[entry.id] = {
        pmCount: entryPm,
        targetPrice: entryTargetPrice,
        items,
        totalBahanCost,
        bufferCost,
        bumbuCost,
        subtotalCost,
        costPerPorsi,
        calculatedCostPerPorsi,
        actualTotalCost,
      };
    });
    return map;
  }, [
    resolvedCategoryEntries,
    activeTargetPriceBesar,
    activeTargetPriceKecil,
    currentDayData.bufferPct,
    tkpiList,
    activeBumbuPct,
    planningMode,
    dayPMForLabels,
    selectedType,
  ]);

  // Custom tables RAB calculation
  const customTablesRABTotal = (customTables || []).reduce((acc, t) => {
    const pCount = Number(t.pmCount) || 0;
    const pPrice = Number(t.plafon) || (t.porsi === "besar" ? activeTargetPriceBesar : activeTargetPriceKecil);
    return acc + (pCount * pPrice);
  }, 0);

  const includeKecilInDashboard = isPerKategoriMode
    ? resolvedCategoryEntries.some((e) => e.porsi === "kecil")
    : !appliedConfig.confirmed || showKecilTable;
  const includeBesarInDashboard = isPerKategoriMode
    ? resolvedCategoryEntries.some((e) => e.porsi === "besar")
    : !appliedConfig.confirmed || showBesarTable;

  const totalRAB =
    (includeKecilInDashboard ? rabPorsiKecil : 0) +
    (includeBesarInDashboard ? rabPorsiBesar : 0) +
    customTablesRABTotal;

  // Realisasi Belanja (Actual) calculation
  const porsiKecilActualCost = isPerKategoriMode
    ? resolvedCategoryEntries
        .filter((e) => e.porsi === "kecil")
        .reduce((s, e) => s + (categoryCalcResultsMap[e.id]?.actualTotalCost || 0), 0)
    : planningMode === "with_bumbu_10"
      ? results.subtotalKecilCost
      : results.totalKecilBahanCost;

  const porsiBesarActualCost = isPerKategoriMode
    ? resolvedCategoryEntries
        .filter((e) => e.porsi === "besar")
        .reduce((s, e) => s + (categoryCalcResultsMap[e.id]?.actualTotalCost || 0), 0)
    : planningMode === "with_bumbu_10"
      ? results.subtotalBesarCost
      : results.totalBesarBahanCost;

  const porsiKecilPerPorsi = pmKecil > 0 ? Math.round(porsiKecilActualCost / pmKecil) : 0;
  const porsiBesarPerPorsi = pmBesar > 0 ? Math.round(porsiBesarActualCost / pmBesar) : 0;

  const bumbuKecilNominal = isPerKategoriMode
    ? resolvedCategoryEntries
        .filter((e) => e.porsi === "kecil")
        .reduce((s, e) => s + (categoryCalcResultsMap[e.id]?.bumbuCost || 0), 0)
    : results.bumbuKecilCost || 0;
  const bumbuBesarNominal = isPerKategoriMode
    ? resolvedCategoryEntries
        .filter((e) => e.porsi === "besar")
        .reduce((s, e) => s + (categoryCalcResultsMap[e.id]?.bumbuCost || 0), 0)
    : results.bumbuBesarCost || 0;

  // Custom tables actual cost & bumbu nominal
  let customTablesBumbuTotal = 0;
  const customTablesActualTotal = (customTables || []).reduce((acc, t) => {
    const calcCustom = t.porsi === "besar"
      ? calculateDay(t.bahanList || [], [], Number(t.pmCount) || 0, 0, t.bufferPct || 5, tkpiList, activeBumbuPct)
      : calculateDay([], t.bahanList || [], 0, Number(t.pmCount) || 0, t.bufferPct || 5, tkpiList, activeBumbuPct);
    
    const bumbuCustom = t.porsi === "besar" ? calcCustom.bumbuBesarCost : calcCustom.bumbuKecilCost;
    customTablesBumbuTotal += bumbuCustom;

    if (t.porsi === "besar") {
      return acc + (planningMode === "with_bumbu_10" ? calcCustom.subtotalBesarCost : calcCustom.totalBesarBahanCost);
    } else {
      return acc + (planningMode === "with_bumbu_10" ? calcCustom.subtotalKecilCost : calcCustom.totalKecilBahanCost);
    }
  }, 0);

  const totalBumbu10Nominal =
    (includeKecilInDashboard ? bumbuKecilNominal : 0) +
    (includeBesarInDashboard ? bumbuBesarNominal : 0) +
    customTablesBumbuTotal;

  const totalActualCost =
    (includeKecilInDashboard ? porsiKecilActualCost : 0) +
    (includeBesarInDashboard ? porsiBesarActualCost : 0) +
    customTablesActualTotal;
  const sisaEfisiensiAnggaran = totalRAB - totalActualCost;

  const updateDayData = (updatedDay: FoodCostDay) => {
    const updatedDays = foodCostDays.map((d) => {
      if (d.hariKe === updatedDay.hariKe && d.jenisMenu === updatedDay.jenisMenu) {
        return updatedDay;
      }
      return d;
    });

    if (!foodCostDays.some((d) => d.hariKe === updatedDay.hariKe && d.jenisMenu === updatedDay.jenisMenu)) {
      updatedDays.push(updatedDay);
    }
    onFoodCostDaysChange(updatedDays);
  };

  // Category Entry CRUD helpers (for "per_kategori" & selective merged categories)
  const commitCategoryEntries = (nextVisibleEntries: FoodCostCategoryEntry[], extraDayPatch?: Partial<FoodCostDay>) => {
    const visibleMap = new Map(nextVisibleEntries.map((e) => [e.id, e]));
    const existingList = currentDayData.categoryEntries || [];
    const mergedList: FoodCostCategoryEntry[] = [];

    activeCategoryGroupsRaw.forEach((g) => {
      if (visibleMap.has(g.id)) {
        mergedList.push(visibleMap.get(g.id)!);
      } else {
        const prevHidden = existingList.find((e) => e.id === g.id);
        if (prevHidden) mergedList.push(prevHidden);
      }
    });

    // Also include any custom entry in nextVisibleEntries not already in activeCategoryGroupsRaw
    nextVisibleEntries.forEach((e) => {
      if (!mergedList.some((m) => m.id === e.id)) {
        mergedList.push(e);
      }
    });

    updateDayData({
      ...currentDayData,
      ...extraDayPatch,
      activeMode: "kategori",
      categoryEntries: mergedList,
    });
  };

  const updateCategoryEntry = (entryId: string, updater: (entry: FoodCostCategoryEntry) => FoodCostCategoryEntry) => {
    const nextVisible = resolvedCategoryEntries.map((e) => (e.id === entryId ? updater(e) : e));
    commitCategoryEntries(nextVisible);
  };

  const addIngredientToCategoryEntry = (entryId: string) => {
    const newBahan: BahanMakananInput = {
      id: `cat_${entryId}_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
      tkpiId: "",
      beratBB: 0,
      urt: "",
      hargaSatuan: 0,
    };
    updateCategoryEntry(entryId, (e) => ({
      ...e,
      bahanList: [...(e.bahanList || []), newBahan],
    }));
  };

  const removeIngredientFromCategoryEntry = (entryId: string, rowId: string, index: number) => {
    updateCategoryEntry(entryId, (e) => ({
      ...e,
      bahanList: (e.bahanList || []).filter((b, idx) => {
        if (rowId && b.id && b.id === rowId) return false;
        if (idx === index) return false;
        return true;
      }),
    }));
  };

  const editIngredientInCategoryEntry = (
    entryId: string,
    rowIndex: number,
    field: keyof BahanMakananInput,
    value: any
  ) => {
    updateCategoryEntry(entryId, (e) => {
      const nextList = [...(e.bahanList || [])];
      if (!nextList[rowIndex]) return e;
      nextList[rowIndex] = { ...nextList[rowIndex], [field]: value };
      return { ...e, bahanList: nextList };
    });
  };

  const editIngredientInCategoryEntryBatch = (
    entryId: string,
    rowIndex: number,
    updates: Partial<BahanMakananInput>
  ) => {
    updateCategoryEntry(entryId, (e) => {
      const nextList = [...(e.bahanList || [])];
      if (!nextList[rowIndex]) return e;
      nextList[rowIndex] = { ...nextList[rowIndex], ...updates };
      return { ...e, bahanList: nextList };
    });
  };

  const getCategoryEntryAkgTarget = (entry: FoodCostCategoryEntry, key: string): number => {
    const type = entry.akgType || getDefaultAkgTypeForGroup(entry.sasaranIds, entry.porsi);
    const mode = entry.nutrientModes?.[key] || entry.akgMode || "min";
    if (!mode || mode === "") return 0;

    if (type === "custom" || mode === "custom") {
      if (entry.customAkgValues && entry.customAkgValues[key] !== undefined) {
        return entry.customAkgValues[key];
      }
    }

    const config = TARGET_AKG_LIMITS[type];
    if (!config) {
      return entry.customAkgValues?.[key] ?? (entry.porsi === "besar" ? customBesar[key] : customKecil[key]) ?? 0;
    }

    const minVal = key === "energi" ? config.energiMin : (config as any)[`${key}Min`] || 0;
    const ratio = config.energiMax / config.energiMin;

    if (mode === "min") return minVal;
    if (mode === "max") {
      if (key === "energi") return config.energiMax;
      return Math.round(minVal * ratio * 10) / 10;
    }
    if (mode === "avg") {
      if (key === "energi") return (config.energiMin + config.energiMax) / 2;
      const maxVal = minVal * ratio;
      return Math.round(((minVal + maxVal) / 2) * 10) / 10;
    }
    return entry.customAkgValues?.[key] || minVal || 0;
  };

  const handleCategoryEntryAkgValChange = (entry: FoodCostCategoryEntry, key: string, rawVal: number) => {
    const val = isNaN(rawVal) ? 0 : rawVal;
    const keys = ["energi", "protein", "lemak", "kh", "serat"];
    const nextCustom: Record<string, number> = { ...(entry.customAkgValues || {}) };
    keys.forEach((k) => {
      if (k === key) {
        nextCustom[k] = val;
      } else if (nextCustom[k] === undefined) {
        nextCustom[k] = getCategoryEntryAkgTarget(entry, k);
      }
    });
    updateCategoryEntry(entry.id, (e) => ({
      ...e,
      akgMode: "custom",
      nutrientModes: { ...(e.nutrientModes || {}), [key]: "custom" },
      customAkgValues: nextCustom,
    }));
  };

  const updateCategoryEntryNutrientMode = (entryId: string, key: string, mode: string) => {
    updateCategoryEntry(entryId, (e) => ({
      ...e,
      nutrientModes: {
        energi: e.nutrientModes?.energi || e.akgMode || "min",
        protein: e.nutrientModes?.protein || e.akgMode || "min",
        lemak: e.nutrientModes?.lemak || e.akgMode || "min",
        kh: e.nutrientModes?.kh || e.akgMode || "min",
        serat: e.nutrientModes?.serat || e.akgMode || "min",
        [key]: mode,
      },
    }));
  };

  // Copy & Paste Food Cost workflow (copy all ingredients & descriptions, paste to next Food Cost and edit BB quickly)
  const focusAndHighlightBb = (targetKey: string, sourceLabel: string, count: number) => {
    setPastedHighlightMap((prev) => ({
      ...prev,
      [targetKey]: { from: sourceLabel, count },
    }));
    setTimeout(() => {
      const el = document.getElementById(`fc-input-bb-${targetKey}-0`) as HTMLInputElement | null;
      if (el) {
        el.focus();
        el.select();
      }
    }, 90);
  };

  const handleCopyFoodCost = (
    sourceLabel: string,
    sourcePorsi: "besar" | "kecil",
    bahanList: BahanMakananInput[],
    bufferPct?: number
  ) => {
    if (!bahanList || bahanList.length === 0) {
      setFcNotice({
        type: "warning",
        text: `Isian "${sourceLabel}" masih kosong, belum ada bahan yang bisa disalin.`,
      });
      return;
    }
    const payload: CopiedFoodCostPayload = {
      sourceLabel,
      sourcePorsi,
      copiedAt: new Date().toLocaleTimeString("id-ID", { hour: "2-digit", minute: "2-digit" }),
      bufferPct: bufferPct !== undefined ? bufferPct : (currentDayData.bufferPct ?? 5),
      bahanList: cloneBahanListFresh(bahanList, "clip"),
    };
    setCopiedFoodCost(payload);
    setFcNotice({
      type: "success",
      text: `Berhasil menyalin ${payload.bahanList.length} bahan & uraian dari "${sourceLabel}". Silakan klik "Tempel Isian" pada kategori/hari berikutnya lalu ubah kolom BERAT (BB).`,
    });
  };

  const executePastePayload = (
    targetType: "category" | "porsi" | "custom",
    targetId: string,
    targetLabel: string,
    payload: CopiedFoodCostPayload,
    mode: "replace" | "append"
  ) => {
    const cloned = cloneBahanListFresh(payload.bahanList, `p_${targetId}`);
    const highlightKey =
      targetType === "category"
        ? `cat_${targetId}`
        : targetType === "custom"
          ? `custom_${targetId}`
          : targetId;

    if (targetType === "category") {
      updateCategoryEntry(targetId, (e) => ({
        ...e,
        bahanList: mode === "append" ? [...(e.bahanList || []), ...cloned] : cloned,
        bufferPct: mode === "replace" ? payload.bufferPct : e.bufferPct,
      }));
    } else if (targetType === "porsi") {
      const listKey = targetId === "besar" ? "porsiBesarBahan" : "porsiKecilBahan";
      const existing = currentDayData[listKey] || [];
      updateDayData({
        ...currentDayData,
        activeMode: "porsi",
        [listKey]: mode === "append" ? [...existing, ...cloned] : cloned,
      });
    } else if (targetType === "custom") {
      setCustomTables((prev) =>
        prev.map((t) => {
          if (t.id !== targetId) return t;
          return {
            ...t,
            bahanList: mode === "append" ? [...(t.bahanList || []), ...cloned] : cloned,
            bufferPct: mode === "replace" ? payload.bufferPct : t.bufferPct,
          };
        })
      );
    }

    focusAndHighlightBb(highlightKey, payload.sourceLabel, cloned.length);
    setFcNotice({
      type: "success",
      text: `Isian "${payload.sourceLabel}" (${cloned.length} bahan) ditempel ke "${targetLabel}". Silakan langsung ubah kolom BERAT (BB) dan tekan Enter untuk lanjut ke baris berikutnya.`,
    });
  };

  const handlePasteToTarget = (
    targetType: "category" | "porsi" | "custom",
    targetId: string,
    targetLabel: string,
    existingCount: number,
    explicitPayload?: CopiedFoodCostPayload
  ) => {
    const payload = explicitPayload || copiedFoodCost;
    if (!payload || !payload.bahanList || payload.bahanList.length === 0) {
      setFcNotice({
        type: "warning",
        text: 'Belum ada data Food Cost yang disalin. Klik tombol "Salin Semua Isian" pada salah satu Food Cost terlebih dahulu.',
      });
      return;
    }

    if (existingCount === 0) {
      executePastePayload(targetType, targetId, targetLabel, payload, "replace");
      return;
    }

    const highlightKey =
      targetType === "category"
        ? `cat_${targetId}`
        : targetType === "custom"
          ? `custom_${targetId}`
          : targetId;

    setPasteConfirmModal({
      isOpen: true,
      targetKey: highlightKey,
      targetLabel,
      existingCount,
      onReplace: () => {
        executePastePayload(targetType, targetId, targetLabel, payload, "replace");
        setPasteConfirmModal(null);
      },
      onAppend: () => {
        executePastePayload(targetType, targetId, targetLabel, payload, "append");
        setPasteConfirmModal(null);
      },
    });
  };

  const handlePasteToAllEmptyCategoryEntries = () => {
    if (!copiedFoodCost || copiedFoodCost.bahanList.length === 0) return;
    const emptyTargets = resolvedCategoryEntries.filter((e) => (e.bahanList || []).length === 0);
    if (emptyTargets.length === 0) {
      setFcNotice({
        type: "info",
        text: "Semua tabel kategori yang tampil sudah memiliki isi bahan baku.",
      });
      return;
    }

    const nextEntries = resolvedCategoryEntries.map((e) => {
      if ((e.bahanList || []).length > 0) return e;
      return {
        ...e,
        bahanList: cloneBahanListFresh(copiedFoodCost.bahanList, `bulk_${e.id}`),
        bufferPct: copiedFoodCost.bufferPct,
      };
    });
    commitCategoryEntries(nextEntries);

    const newHighlights: Record<string, { from: string; count: number }> = { ...pastedHighlightMap };
    emptyTargets.forEach((e) => {
      newHighlights[`cat_${e.id}`] = {
        from: copiedFoodCost.sourceLabel,
        count: copiedFoodCost.bahanList.length,
      };
    });
    setPastedHighlightMap(newHighlights);

    setFcNotice({
      type: "success",
      text: `Berhasil menempel isian "${copiedFoodCost.sourceLabel}" ke ${emptyTargets.length} tabel kategori kosong. Anda tinggal menyesuaikan BERAT (BB) di masing-masing kategori.`,
    });
  };

  // Quick-pull sources available right from dropdown on each table
  const copySourceOptions = React.useMemo(() => {
    const opts: { key: string; label: string; porsi: "besar" | "kecil"; bufferPct: number; bahanList: BahanMakananInput[] }[] = [];
    if (isPerKategoriMode) {
      resolvedCategoryEntries.forEach((e) => {
        if ((e.bahanList || []).length > 0) {
          opts.push({
            key: `cat:${e.id}`,
            label: `${e.title} (${e.bahanList.length} bahan)`,
            porsi: e.porsi,
            bufferPct: e.bufferPct ?? currentDayData.bufferPct ?? 5,
            bahanList: e.bahanList,
          });
        }
      });
    }
    if ((currentDayData.porsiBesarBahan || []).length > 0) {
      opts.push({
        key: "porsi:besar",
        label: `Porsi Besar Hari ${selectedDay} (${currentDayData.porsiBesarBahan.length} bahan)`,
        porsi: "besar",
        bufferPct: currentDayData.bufferPct ?? 5,
        bahanList: currentDayData.porsiBesarBahan,
      });
    }
    if ((currentDayData.porsiKecilBahan || []).length > 0) {
      opts.push({
        key: "porsi:kecil",
        label: `Porsi Kecil Hari ${selectedDay} (${currentDayData.porsiKecilBahan.length} bahan)`,
        porsi: "kecil",
        bufferPct: currentDayData.bufferPct ?? 5,
        bahanList: currentDayData.porsiKecilBahan,
      });
    }
    // Also collect populated entries from other days in foodCostDays
    foodCostDays.forEach((d) => {
      if (d.hariKe === selectedDay && d.jenisMenu === activeDbType) return;
      (d.categoryEntries || []).forEach((ce) => {
        if ((ce.bahanList || []).length > 0) {
          opts.push({
            key: `day_${d.hariKe}_${d.jenisMenu}_cat_${ce.id}`,
            label: `H${d.hariKe} (${d.jenisMenu}) · ${ce.title} (${ce.bahanList.length} bahan)`,
            porsi: ce.porsi,
            bufferPct: ce.bufferPct ?? d.bufferPct ?? 5,
            bahanList: ce.bahanList,
          });
        }
      });
      if ((d.porsiBesarBahan || []).length > 0) {
        opts.push({
          key: `day_${d.hariKe}_${d.jenisMenu}_besar`,
          label: `H${d.hariKe} (${d.jenisMenu}) · Porsi Besar (${d.porsiBesarBahan.length} bahan)`,
          porsi: "besar",
          bufferPct: d.bufferPct ?? 5,
          bahanList: d.porsiBesarBahan,
        });
      }
      if ((d.porsiKecilBahan || []).length > 0) {
        opts.push({
          key: `day_${d.hariKe}_${d.jenisMenu}_kecil`,
          label: `H${d.hariKe} (${d.jenisMenu}) · Porsi Kecil (${d.porsiKecilBahan.length} bahan)`,
          porsi: "kecil",
          bufferPct: d.bufferPct ?? 5,
          bahanList: d.porsiKecilBahan,
        });
      }
    });
    return opts;
  }, [
    isPerKategoriMode,
    resolvedCategoryEntries,
    currentDayData,
    foodCostDays,
    selectedDay,
    activeDbType,
  ]);

  // Selective category grouping helpers (1 category each vs merge selected categories)
  const syncGroupsState = (nextGroups: FoodCostCategoryGroup[], nextMode?: "per_kategori" | "gabung_porsi") => {
    setDraftCategoryGroupsByGroup((prev) => ({
      ...prev,
      [activeGroupKey]: nextGroups,
    }));
    setAppliedConfig((prev) => ({
      ...prev,
      confirmed: true,
      layoutMode: nextMode || prev.layoutMode,
      categoryGroupsByGroup: {
        ...prev.categoryGroupsByGroup,
        [activeGroupKey]: nextGroups,
      },
    }));
  };

  const handleSwitchLayoutMode = (mode: "per_kategori" | "gabung_porsi") => {
    setDraftLayoutMode(mode);
    setMergeSelectedGroupIds([]);
    setAppliedConfig((prev) => ({
      ...prev,
      confirmed: true,
      layoutMode: mode,
    }));
    updateDayData({
      ...currentDayData,
      activeMode: mode === "per_kategori" ? "kategori" : "porsi",
    });
  };

  const toggleMergeSelection = (groupId: string) => {
    setMergeSelectedGroupIds((prev) =>
      prev.includes(groupId) ? prev.filter((id) => id !== groupId) : [...prev, groupId]
    );
  };

  const handleMergeSelectedDraftGroups = () => {
    const currentGroups = draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey);
    const selectedSet = new Set(mergeSelectedGroupIds);
    const toMerge = currentGroups.filter((g) => selectedSet.has(g.id));
    if (toMerge.length < 2) {
      setFcNotice({
        type: "warning",
        text: "Centang minimal 2 kategori terlebih dahulu untuk disatukan menjadi 1 Food Cost.",
      });
      return;
    }

    const mergedSasaranIds: string[] = Array.from(
      new Set<string>(toMerge.flatMap((g) => g.sasaranIds))
    );
    const resolvedPorsi: "besar" | "kecil" =
      mergeCustomPorsi !== "auto"
        ? mergeCustomPorsi
        : toMerge.every((g) => g.porsi === "kecil")
          ? "kecil"
          : "besar";
    const resolvedTitle =
      mergeCustomTitle.trim() || toMerge.map((g) => g.title).join(" + ");
    const mergedId = `merged_${Date.now()}_${Math.random().toString(36).slice(2, 5)}`;

    const mergedGroup: FoodCostCategoryGroup = {
      id: mergedId,
      title: resolvedTitle,
      porsi: resolvedPorsi,
      sasaranIds: mergedSasaranIds,
    };

    const nextGroups: FoodCostCategoryGroup[] = [];
    let inserted = false;
    currentGroups.forEach((g) => {
      if (selectedSet.has(g.id)) {
        if (!inserted) {
          nextGroups.push(mergedGroup);
          inserted = true;
        }
      } else {
        nextGroups.push(g);
      }
    });

    // Preserve first populated source's ingredients into the newly created merged entry if available
    const existingEntries = currentDayData.categoryEntries || [];
    const firstSourceWithData = toMerge
      .map((g) => existingEntries.find((e) => e.id === g.id))
      .find((e) => e && (e.bahanList || []).length > 0);

    if (firstSourceWithData) {
      const mergedEntry: FoodCostCategoryEntry = {
        id: mergedId,
        title: resolvedTitle,
        porsi: resolvedPorsi,
        sasaranIds: mergedSasaranIds,
        bahanList: cloneBahanListFresh(firstSourceWithData.bahanList, "mrg"),
        bufferPct: firstSourceWithData.bufferPct ?? currentDayData.bufferPct ?? 5,
        akgType: getDefaultAkgTypeForGroup(mergedSasaranIds, resolvedPorsi),
        akgMode: firstSourceWithData.akgMode || "min",
        nutrientModes: firstSourceWithData.nutrientModes,
        kriteriaPm: resolvedTitle,
      };
      updateDayData({
        ...currentDayData,
        activeMode: "kategori",
        categoryEntries: [...existingEntries, mergedEntry],
      });
    }

    syncGroupsState(nextGroups, "per_kategori");
    setDraftLayoutMode("per_kategori");
    setMergeSelectedGroupIds([]);
    setMergeCustomTitle("");
    setMergeCustomPorsi("auto");
    setFcNotice({
      type: "success",
      text: `Berhasil menyatukan ${toMerge.length} kategori menjadi 1 Food Cost: "${resolvedTitle}".`,
    });
  };

  const handleSplitMergedGroup = (groupId: string) => {
    const currentGroups = draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey);
    const target = currentGroups.find((g) => g.id === groupId);
    if (!target || target.sasaranIds.length <= 1) return;

    const splitItems: FoodCostCategoryGroup[] = target.sasaranIds.map((sId) => ({
      id: `cat_${sId}`,
      title: getSasaranLabel(sId),
      porsi: resolveDefaultPorsiForSasaran(sId),
      sasaranIds: [sId],
    }));

    const nextGroups: FoodCostCategoryGroup[] = [];
    const seenIds = new Set<string>();
    currentGroups.forEach((g) => {
      if (g.id === groupId) {
        splitItems.forEach((si) => {
          if (!seenIds.has(si.id)) {
            seenIds.add(si.id);
            nextGroups.push(si);
          }
        });
      } else if (!seenIds.has(g.id)) {
        seenIds.add(g.id);
        nextGroups.push(g);
      }
    });

    syncGroupsState(nextGroups, "per_kategori");
    setMergeSelectedGroupIds((prev) => prev.filter((id) => id !== groupId));
    setFcNotice({
      type: "info",
      text: `Grup "${target.title}" dipisahkan kembali menjadi ${splitItems.length} kategori Food Cost masing-masing.`,
    });
  };

  const handleRemoveGroupFromActive = (groupId: string) => {
    const currentGroups = draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey);
    if (currentGroups.length <= 1) {
      setFcNotice({
        type: "warning",
        text: "Minimal harus ada 1 kategori Food Cost yang aktif.",
      });
      return;
    }
    const nextGroups = currentGroups.filter((g) => g.id !== groupId);
    syncGroupsState(nextGroups, "per_kategori");
    setMergeSelectedGroupIds((prev) => prev.filter((id) => id !== groupId));
  };

  const handleAddSingleCategoryToActive = (sId: string) => {
    const currentGroups = draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey);
    if (currentGroups.some((g) => g.sasaranIds.includes(sId))) return;
    const newGroup: FoodCostCategoryGroup = {
      id: `cat_${sId}`,
      title: getSasaranLabel(sId),
      porsi: resolveDefaultPorsiForSasaran(sId),
      sasaranIds: [sId],
    };
    syncGroupsState([...currentGroups, newGroup], "per_kategori");
  };

  const handleResetToSingleCategories = (allWithAlergi = false) => {
    const nextGroups = allWithAlergi
      ? ALL_KATEGORI_WITH_ALERGI.map((s) => ({
          id: `cat_${s.id}`,
          title: getSasaranLabel(s.id),
          porsi: resolveDefaultPorsiForSasaran(s.id),
          sasaranIds: [s.id],
        }))
      : buildDefaultGroupsForTabGroup(activeGroupKey);
    syncGroupsState(nextGroups, "per_kategori");
    setDraftLayoutMode("per_kategori");
    setMergeSelectedGroupIds([]);
    setFcNotice({
      type: "info",
      text: allWithAlergi
        ? `Menampilkan seluruh ${nextGroups.length} kategori sasaran (Standar & Diet Alergi, masing-masing 1 Food Cost).`
        : `Masing-masing 1 kategori Food Cost aktif + grup Kategori Alergi sesuai kelompok terpilih.`,
    });
  };

  const handleAddAllAlergiCategories = () => {
    const currentGroups =
      draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey);
    const existingAssigned = new Set(currentGroups.flatMap((g) => g.sasaranIds));
    const missingAlergi = ALERGI_KATEGORI_LIST.filter(
      (a) => !existingAssigned.has(a.id) && getSasaranPmCount(a.id, resolveDefaultPorsiForSasaran(a.id)) > 0
    );
    if (missingAlergi.length === 0) {
      const alergiGroupsToAdd = buildDefaultAlergiGroupsForTabGroup(activeGroupKey).filter(
        (ag) => !currentGroups.some((cg) => cg.id === ag.id)
      );
      if (alergiGroupsToAdd.length === 0) {
        setFcNotice({
          type: "info",
          text: "Semua Kategori Alergi yang aktif sudah ada di panel Food Cost.",
        });
        return;
      }
      syncGroupsState([...currentGroups, ...alergiGroupsToAdd], "per_kategori");
      setFcNotice({
        type: "success",
        text: `Berhasil menambahkan ${alergiGroupsToAdd.length} panel Kategori Alergi ke Food Cost.`,
      });
      return;
    }

    const newAlergiGroups: FoodCostCategoryGroup[] = missingAlergi.map((a) => ({
      id: `cat_${a.id}`,
      title: getSasaranLabel(a.id),
      porsi: resolveDefaultPorsiForSasaran(a.id),
      sasaranIds: [a.id],
    }));
    syncGroupsState([...currentGroups, ...newAlergiGroups], "per_kategori");
    setFcNotice({
      type: "success",
      text: `Berhasil menambahkan ${newAlergiGroups.length} Kategori Alergi dari Tabel 2 ke panel Food Cost.`,
    });
  };

  const handleKeepOnlyPmPositiveCategories = () => {
    const allSingleGroups: FoodCostCategoryGroup[] = ALL_KATEGORI_WITH_ALERGI.map((s) => ({
      id: `cat_${s.id}`,
      title: getSasaranLabel(s.id),
      porsi: resolveDefaultPorsiForSasaran(s.id),
      sasaranIds: [s.id],
    }));
    const positiveGroups = allSingleGroups.filter(
      (g) => getEntryAutoPmCount(g.sasaranIds, g.porsi) > 0
    );
    if (positiveGroups.length === 0) {
      const baseGroups = buildDefaultGroupsForTabGroup(activeGroupKey);
      setFcNotice({
        type: "warning",
        text: `Belum ada kategori dengan jumlah PM > 0 pada Hari Ke-${selectedDay}. Menampilkan semua kategori default kelompok.`,
      });
      syncGroupsState(baseGroups, "per_kategori");
      return;
    }
    syncGroupsState(positiveGroups, "per_kategori");
    setDraftLayoutMode("per_kategori");
    setMergeSelectedGroupIds([]);
    setFcNotice({
      type: "success",
      text: `Menampilkan ${positiveGroups.length} kategori (Standar & Diet Alergi) yang memiliki Penerima Manfaat (PM > 0) di Tabel 2.`,
    });
  };

  const handleBufferChange = (val: number) => {
    if (isPerKategoriMode && (currentDayData.categoryEntries || []).length > 0) {
      const updatedEntries = resolvedCategoryEntries.map((e) => ({ ...e, bufferPct: val }));
      commitCategoryEntries(updatedEntries, { bufferPct: val });
      return;
    }
    updateDayData({ ...currentDayData, bufferPct: val });
  };

  const autoMatchFromMasterMenu = (options?: { silent?: boolean; onlyIfEmpty?: boolean; scope?: "kecil" | "besar" | "all" }) => {
    const menuItem = effectiveMenuItem;

    if (!menuItem || !composeMenuName(menuItem)) {
      if (!options?.silent) {
        setFcNotice({
          type: "warning",
          text: "Menu Master belum diisi untuk hari dan kategori ini. Silakan isi di tab Master Menu terlebih dahulu.",
        });
      }
      return;
    }

    const matchTkpi = (text: string, defaultId: string): string => {
      if (!text) return defaultId;
      const t = text.toLowerCase();
      
      const found = tkpiList.find(item => {
        const name = item.nama.toLowerCase();
        return name.includes(t) || t.includes(name);
      });
      if (found) return found.id;

      if (t.includes("beras") || t.includes("nasi") || t.includes("bubur")) return "beras_giling";
      if (t.includes("kentang")) return "kentang_segar";
      if (t.includes("singkong") || t.includes("ubi")) return "singkong_segar";
      
      if (t.includes("ayam")) return "daging_ayam_tanpa_kulit";
      if (t.includes("sapi") || t.includes("daging")) return "daging_sapi";
      if (t.includes("telur")) return "telur_ayam_ras";
      if (t.includes("ikan")) {
        if (t.includes("kembung")) return "ikan_kembung_segar";
        if (t.includes("mujair")) return "ikan_mujair_segar";
        if (t.includes("teri")) return "ikan_teri_kering";
        return "ikan_kembung_segar";
      }
      if (t.includes("udang")) return "udang_segar";
      
      if (t.includes("tempe")) return "tempe_kedelai";
      if (t.includes("tahu")) return "tahu_mentah";
      if (t.includes("kacang hijau")) return "kacang_hijau_kering";
      if (t.includes("kacang merah")) return "kacang_merah_segar";
      if (t.includes("kacang tanah")) return "kacang_tanah_kupas";
      
      if (t.includes("wortel")) return "wortel_segar";
      if (t.includes("bayam")) return "bayam_segar";
      if (t.includes("kangkung")) return "kangkung_segar";
      if (t.includes("labu siam")) return "labu_siam_segar";
      if (t.includes("labu kuning")) return "labu_kuning";
      if (t.includes("sawi")) return "sawi_segar";
      if (t.includes("kol") || t.includes("kubis")) return "kubis_kol_segar";
      if (t.includes("tomat")) return "tomat_masak_segar";
      
      if (t.includes("pisang")) {
        if (t.includes("ambon")) return "pisang_ambon";
        if (t.includes("raja")) return "pisang_raja";
        return "pisang_ambon";
      }
      if (t.includes("jeruk")) return "jeruk_manis";
      if (t.includes("pepaya")) return "pepaya_segar";
      if (t.includes("apel")) return "apel_segar";
      if (t.includes("semangka")) return "semangka";
      if (t.includes("susu")) return "susu_bubuk_fullcream";

      return defaultId;
    };

    const pbBahan: BahanMakananInput[] = [];
    const pkBahan: BahanMakananInput[] = [];

    const components = [
      { text: menuItem.karbohidrat, defaultTkpi: "beras_giling", defaultBeratPb: 80, defaultBeratPk: 50, defaultUrtPb: "1 piring", defaultUrtPk: "1/2 piring" },
      { text: menuItem.laukHewani, defaultTkpi: "daging_ayam_tanpa_kulit", defaultBeratPb: 65, defaultBeratPk: 40, defaultUrtPb: "1 potong md", defaultUrtPk: "1 potong kcl" },
      { text: menuItem.laukNabati, defaultTkpi: "tempe_kedelai", defaultBeratPb: 30, defaultBeratPk: 20, defaultUrtPb: "1 potong", defaultUrtPk: "1/2 potong" },
      { text: menuItem.sayur, defaultTkpi: "wortel_segar", defaultBeratPb: 40, defaultBeratPk: 25, defaultUrtPb: "1/2 gelas", defaultUrtPk: "1/3 gelas" },
      { text: menuItem.buah || menuItem.buahSusu, defaultTkpi: "pisang_ambon", defaultBeratPb: 100, defaultBeratPk: 80, defaultUrtPb: "1 buah", defaultUrtPk: "1 buah kcl" },
      ...(menuItem.susu ? [{ text: menuItem.susu, defaultTkpi: "susu_bubuk_fullcream", defaultBeratPb: 25, defaultBeratPk: 20, defaultUrtPb: "1 kotak/gls", defaultUrtPk: "1 kotak kcl" }] : [])
    ];

    const standardPrices: Record<string, number> = {
      beras_giling: 14000,
      kentang_segar: 20000,
      singkong_segar: 10000,
      daging_ayam_tanpa_kulit: 40000,
      daging_sapi: 120000,
      telur_ayam_ras: 28000,
      ikan_kembung_segar: 35000,
      ikan_mujair_segar: 30000,
      ikan_teri_kering: 80000,
      udang_segar: 70000,
      tempe_kedelai: 15000,
      tahu_mentah: 10000,
      kacang_hijau_kering: 25000,
      kacang_merah_segar: 24000,
      kacang_tanah_kupas: 28000,
      wortel_segar: 18000,
      bayam_segar: 15000,
      kangkung_segar: 12000,
      labu_siam_segar: 12000,
      labu_kuning: 15000,
      sawi_segar: 14000,
      kubis_kol_segar: 12000,
      tomat_masak_segar: 16000,
      pisang_ambon: 12000,
      pisang_raja: 15000,
      jeruk_manis: 20000,
      pepaya_segar: 10000,
      apel_segar: 30000,
      semangka: 10000,
      susu_bubuk_fullcream: 95000
    };

    components.forEach((comp, idx) => {
      if (!comp.text || comp.text.trim() === "" || comp.text.toLowerCase().includes("tidak ada") || comp.text.toLowerCase() === "-") {
        return;
      }

      const tkpiId = matchTkpi(comp.text, comp.defaultTkpi);
      const hargaDb = standardPrices[tkpiId] || 20000;

      pbBahan.push({
        id: `pb_auto_${idx}_${Date.now()}`,
        tkpiId,
        beratBB: comp.defaultBeratPb,
        urt: comp.defaultUrtPb,
        hargaSatuan: hargaDb
      });

      pkBahan.push({
        id: `pk_auto_${idx}_${Date.now()}`,
        tkpiId,
        beratBB: comp.defaultBeratPk,
        urt: comp.defaultUrtPk,
        hargaSatuan: hargaDb
      });
    });

    if (pbBahan.length === 0) {
      if (!options?.silent) {
        setFcNotice({
          type: "warning",
          text: "Gagal mengidentifikasi bahan makanan yang valid dari Master Menu.",
        });
      }
      return;
    }

    const targetScope = options?.scope || "all";

    if (isPerKategoriMode) {
      let updatedCount = 0;
      const nextEntries = resolvedCategoryEntries.map((entry) => {
        if (targetScope !== "all" && entry.porsi !== targetScope) return entry;
        if (options?.onlyIfEmpty && (entry.bahanList || []).length > 0) return entry;
        const sourceBahan = entry.porsi === "besar" ? pbBahan : pkBahan;
        updatedCount++;
        return {
          ...entry,
          bahanList: cloneBahanListFresh(sourceBahan, `mm_${entry.id}`),
        };
      });
      if (updatedCount > 0) {
        commitCategoryEntries(nextEntries);
        if (!options?.silent) {
          setFcNotice({
            type: "success",
            text: `Berhasil memetakan ${pbBahan.length} komponen menu "${namaMenuHariIni}" ke ${updatedCount} tabel kategori Food Cost.`,
          });
        }
      }
      return;
    }

    const fillBesar =
      (targetScope === "all" || targetScope === "besar") &&
      (!options?.onlyIfEmpty || currentDayData.porsiBesarBahan.length === 0);
    const fillKecil =
      (targetScope === "all" || targetScope === "kecil") &&
      (!options?.onlyIfEmpty || currentDayData.porsiKecilBahan.length === 0);

    if (!fillBesar && !fillKecil) return;

    updateDayData({
      ...currentDayData,
      activeMode: "porsi",
      porsiBesarBahan: fillBesar ? pbBahan : currentDayData.porsiBesarBahan,
      porsiKecilBahan: fillKecil ? pkBahan : currentDayData.porsiKecilBahan,
    });

    if (!options?.silent) {
      setFcNotice({
        type: "success",
        text: `Berhasil menyalin & memetakan ${pbBahan.length} komponen menu "${namaMenuHariIni}" ke tabel Food Cost.`,
      });
    }
  };

  const handleOpenCategoryConfirmModal = () => {
    if (draftLayoutMode === "per_kategori") {
      const currentGroups = draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey);
      const filteredGroups = currentGroups.filter((g) => draftScope === "all" || g.porsi === draftScope);
      if (filteredGroups.length === 0) {
        setFcNotice({
          type: "warning",
          text: "Pilih minimal 1 kategori Food Cost sebelum melakukan konfirmasi.",
        });
        return;
      }
      setFcNotice(null);
      setShowCategoryConfirmModal(true);
      return;
    }

    const missingKecil = (draftScope === "kecil" || draftScope === "all") && draftKecilIds.length === 0;
    const missingBesar = (draftScope === "besar" || draftScope === "all") && draftBesarIds.length === 0;

    if (draftScope === "kecil" && missingKecil) {
      setFcNotice({
        type: "warning",
        text: "Pilih minimal 1 kategori sasaran Porsi Kecil sebelum melakukan konfirmasi.",
      });
      return;
    }
    if (draftScope === "besar" && missingBesar) {
      setFcNotice({
        type: "warning",
        text: "Pilih minimal 1 kategori sasaran Porsi Besar sebelum melakukan konfirmasi.",
      });
      return;
    }
    if (draftScope === "all" && missingKecil && missingBesar) {
      setFcNotice({
        type: "warning",
        text: "Pilih minimal 1 kategori Porsi Kecil atau Porsi Besar terlebih dahulu sebelum konfirmasi.",
      });
      return;
    }
    setFcNotice(null);
    setShowCategoryConfirmModal(true);
  };

  const handleConfirmApplyFoodCost = () => {
    const resolvedKecilIds = draftScope === "besar" ? draftKecilIds : draftKecilIds;
    const resolvedBesarIds = draftScope === "kecil" ? draftBesarIds : draftBesarIds;

    setAppliedConfig((prev) => ({
      ...prev,
      confirmed: true,
      layoutMode: draftLayoutMode,
      scope: draftScope,
      kecilIds: resolvedKecilIds,
      besarIds: resolvedBesarIds,
      categoryGroupsByGroup: draftCategoryGroupsByGroup,
    }));
    updateDayData({
      ...currentDayData,
      activeMode: draftLayoutMode === "per_kategori" ? "kategori" : "porsi",
    });
    setShowCategoryConfirmModal(false);
    setIsSelectorExpanded(false);

    if (autoFillOnConfirm && selectedGroup !== "custom") {
      autoMatchFromMasterMenu({ silent: true, onlyIfEmpty: true, scope: draftScope });
    }

    const modeLabel =
      draftLayoutMode === "per_kategori"
        ? "Masing-Masing 1 Kategori"
        : draftScope === "kecil"
          ? "Gabung Porsi Kecil"
          : draftScope === "besar"
            ? "Gabung Porsi Besar"
            : "Gabung Semua Porsi (Kecil & Besar)";

    setFcNotice({
      type: "success",
      text: `Konfirmasi berhasil! Menampilkan Food Cost dengan mode: ${modeLabel}.`,
    });

    setTimeout(() => {
      document.getElementById("fc-tables-container")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 80);
  };

  const renderCategorySignBanner = (porsi: "besar" | "kecil", isPrint: boolean) => {
    const categories = porsi === "besar" ? selectedBesarCategories : selectedKecilCategories;
    const totalPm = porsi === "besar" ? pmBesar : pmKecil;
    const plafon = porsi === "besar" ? targetPriceBesar : targetPriceKecil;
    const isKecil = porsi === "kecil";

    if (isPrint) {
      return (
        <div
          className={`rounded-lg border px-3 py-2 flex flex-wrap items-center justify-between gap-2 text-[10px] ${
            isKecil
              ? "bg-indigo-50/70 border-indigo-400 text-indigo-950"
              : "bg-cyan-50/70 border-cyan-500 text-cyan-950"
          }`}
        >
          <div className="flex flex-wrap items-center gap-1.5">
            <span className="font-black uppercase tracking-wide px-1.5 py-0.5 rounded bg-white border border-black/20">
              TANDA KATEGORI: {isKecil ? "PORSI KECIL" : "PORSI BESAR"} ({selectedType})
            </span>
            <span className="font-bold">
              {categories.length > 0
                ? categories.map((c) => `${c.label} (${c.count} PM)`).join(" · ")
                : "Belum ada kategori sasaran terpilih"}
            </span>
          </div>
          <span className="font-mono font-black">
            Total: {totalPm} PM · Plafon {formatRupiah(plafon)}/porsi
          </span>
        </div>
      );
    }

    return (
      <div
        className={`rounded-xl border px-4 py-3 flex flex-col gap-2 ${
          isKecil
            ? "bg-indigo-50/80 border-indigo-200"
            : "bg-cyan-50/80 border-cyan-200"
        }`}
      >
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex flex-wrap items-center gap-2">
            <span
              className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider text-white shadow-xs ${
                isKecil ? "bg-indigo-600" : "bg-cyan-700"
              }`}
            >
              <Tag className="w-3 h-3" />
              Tanda Kategori Dipilih: {isKecil ? "Porsi Kecil" : "Porsi Besar"}
            </span>
            <span className="px-2 py-0.5 rounded-md bg-white border border-slate-200 text-[10px] font-extrabold text-slate-700 uppercase">
              Diet: {selectedType}
            </span>
            <span className="text-[11px] font-bold text-slate-700">
              Menu: <span className="text-slate-950 font-extrabold">{namaMenuHariIni || "-"}</span>
            </span>
          </div>
          <span
            className={`text-[11px] font-black font-mono px-2.5 py-1 rounded-lg bg-white border ${
              isKecil ? "text-indigo-800 border-indigo-200" : "text-cyan-800 border-cyan-200"
            }`}
          >
            {totalPm} PM · Plafon {formatRupiah(plafon)}/porsi
          </span>
        </div>
        <div className="flex flex-wrap items-center gap-1.5">
          <span className="text-[10px] font-bold text-slate-500 uppercase mr-1">
            Sasaran ({categories.length}):
          </span>
          {categories.length > 0 ? (
            categories.map((c) => (
              <span
                key={c.id}
                className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[10px] font-bold border bg-white ${
                  isKecil
                    ? "text-indigo-800 border-indigo-200"
                    : "text-cyan-800 border-cyan-200"
                }`}
              >
                <Check className="w-3 h-3 shrink-0" />
                {c.label}
                <span className="text-slate-400 font-mono">({c.count})</span>
              </span>
            ))
          ) : (
            <span className="text-[10px] font-bold text-rose-600">
              Belum ada kategori sasaran dipilih untuk porsi ini.
            </span>
          )}
        </div>
      </div>
    );
  };

  const addIngredientRow = (porsi: "besar" | "kecil") => {
    const newBahan: BahanMakananInput = {
      id: `${porsi}_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
      tkpiId: "",
      beratBB: 0,
      urt: "",
      hargaSatuan: 0
    };

    if (porsi === "besar") {
      updateDayData({
        ...currentDayData,
        porsiBesarBahan: [...currentDayData.porsiBesarBahan, newBahan]
      });
    } else {
      updateDayData({
        ...currentDayData,
        porsiKecilBahan: [...currentDayData.porsiKecilBahan, newBahan]
      });
    }
  };

  const removeIngredientRow = (porsi: "besar" | "kecil", id: string, index: number) => {
    const listKey = porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan";
    const currentList = currentDayData[listKey] || [];
    const updatedList = currentList.filter((b, idx) => {
      if (id && b.id && b.id === id) return false;
      if (idx === index) return false;
      return true;
    });

    updateDayData({
      ...currentDayData,
      [listKey]: updatedList
    });
  };

  const handleRowChange = (
    porsi: "besar" | "kecil",
    index: number,
    field: keyof BahanMakananInput,
    value: any
  ) => {
    const listKey = porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan";
    const list = [...currentDayData[listKey]];
    list[index] = { ...list[index], [field]: value };
    updateDayData({ ...currentDayData, [listKey]: list });
  };

  const handleRowBatchChange = (
    porsi: "besar" | "kecil",
    index: number,
    updates: Partial<BahanMakananInput>
  ) => {
    const listKey = porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan";
    const list = [...currentDayData[listKey]];
    list[index] = { ...list[index], ...updates };
    updateDayData({ ...currentDayData, [listKey]: list });
  };

  const targetPriceKecil = activeTargetPriceKecil;
  const targetPriceBesar = activeTargetPriceBesar;

  const activeCostPerPorsiBesar =
    planningMode === "all_ingredients"
      ? pmBesar > 0
        ? results.totalBesarBahanCost / pmBesar
        : 0
      : results.costPerPorsiBesar;
  const activeCostPerPorsiKecil =
    planningMode === "all_ingredients"
      ? pmKecil > 0
        ? results.totalKecilBahanCost / pmKecil
        : 0
      : results.costPerPorsiKecil;
  const isOverBudgetBesar = pmBesar > 0 && activeCostPerPorsiBesar > targetPriceBesar;
  const isOverBudgetKecil = pmKecil > 0 && activeCostPerPorsiKecil > targetPriceKecil;

  const getOverBudgetKey = (scope: string) => `${selectedDay}_${activeDbType}_${scope}`;

  const isOverBudgetContinued = (scope: string, actualCost: number) => {
    const decision = overBudgetDecisions[getOverBudgetKey(scope)];
    if (!decision || decision.action !== "continue") return false;
    return Math.round(actualCost) <= Math.round(decision.cost) + 1;
  };

  const handleContinueOverBudget = (scope: string, actualCost: number) => {
    setOverBudgetDecisions((prev) => ({
      ...prev,
      [getOverBudgetKey(scope)]: { cost: actualCost, action: "continue" },
    }));
  };

  const openChangeMenuModal = (
    scope: string,
    label: string,
    actualCost: number,
    targetPrice: number,
    items: CalcIngredientResult[],
    pmCount: number
  ) => {
    const multiplier = planningMode === "with_bumbu_10" ? (1 + activeBumbuPct / 100) : 1;
    const sortedDrivers = items
      .map((item) => ({
        id: item.id,
        nama: item.nama || "Bahan belum dipilih",
        costPerPortion: pmCount > 0 ? (item.hargaTotal * multiplier) / pmCount : 0,
      }))
      .filter((x) => x.costPerPortion > 0)
      .sort((a, b) => b.costPerPortion - a.costPerPortion)
      .slice(0, 5);

    setChangeMenuNameInput(namaMenuHariIni);
    setChangeMenuModal({
      isOpen: true,
      scope,
      label,
      actualCost,
      targetPrice,
      items: sortedDrivers,
    });
  };

  const renderOverBudgetNotice = (
    scope: string,
    label: string,
    actualCost: number,
    targetPrice: number,
    recipientCount: number,
    items: CalcIngredientResult[]
  ) => {
    if (recipientCount <= 0 || actualCost <= 0) return null;
    const diff = actualCost - targetPrice;
    if (diff <= 0) {
      return (
        <div className="mt-2 px-3.5 py-2 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-between gap-2 text-[11px] font-bold text-emerald-800 no-print">
          <span className="flex items-center gap-1.5">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            Harga per porsi {label} (Rp {formatThousandSeparator(Math.round(actualCost))}) sesuai plafon Data PM (Rp {formatThousandSeparator(targetPrice)}) — efisiensi Rp {formatThousandSeparator(Math.round(Math.abs(diff)))}/porsi.
          </span>
        </div>
      );
    }

    const continued = isOverBudgetContinued(scope, actualCost);
    if (continued) {
      return (
        <div className="mt-2 px-4 py-2.5 rounded-xl bg-amber-50 border border-amber-300 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs no-print">
          <div className="flex items-center gap-2 text-amber-900 font-bold">
            <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
            <span>
              Over Budget (+Rp {formatThousandSeparator(Math.round(diff))}/porsi dari plafon Rp {formatThousandSeparator(targetPrice)}) —{" "}
              <span className="underline">Tetap Dilanjutkan</span>.
            </span>
          </div>
          <button
            type="button"
            onClick={() => openChangeMenuModal(scope, label, actualCost, targetPrice, items, recipientCount)}
            className="px-3 py-1.5 bg-white hover:bg-amber-100 text-amber-900 border border-amber-300 rounded-lg text-[11px] font-extrabold transition cursor-pointer shrink-0 self-start sm:self-auto"
          >
            Ubah Menu
          </button>
        </div>
      );
    }

    return (
      <div className="mt-2.5 p-4 rounded-xl bg-rose-50 border-2 border-rose-400 flex flex-col lg:flex-row lg:items-center justify-between gap-3 shadow-xs no-print">
        <div className="flex items-start gap-2.5">
          <div className="p-2 bg-rose-100 text-rose-700 rounded-lg shrink-0 mt-0.5">
            <AlertTriangle className="w-5 h-5" />
          </div>
          <div className="space-y-0.5">
            <h5 className="text-xs font-black text-rose-900 uppercase tracking-wide">
              Peringatan Over Budget Harga Per Porsi ({label})
            </h5>
            <p className="text-xs text-rose-800 font-semibold leading-relaxed">
              Harga per porsi saat ini{" "}
              <span className="font-mono font-black text-rose-950">
                Rp {formatThousandSeparator(Math.round(actualCost))}
              </span>{" "}
              melebihi plafon harga per porsi dari Data PM (
              <span className="font-mono font-black">
                Rp {formatThousandSeparator(targetPrice)}
              </span>
              ) sebesar{" "}
              <span className="font-mono font-black text-rose-700">
                +Rp {formatThousandSeparator(Math.round(diff))} / porsi
              </span>
              . Pilih apakah ingin tetap melanjutkan atau mengubah menu.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0 pl-9 lg:pl-0">
          <button
            type="button"
            onClick={() => handleContinueOverBudget(scope, actualCost)}
            className="px-3.5 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-xl text-xs font-extrabold transition cursor-pointer shadow-xs"
          >
            Tetap Lanjutkan
          </button>
          <button
            type="button"
            onClick={() => openChangeMenuModal(scope, label, actualCost, targetPrice, items, recipientCount)}
            className="px-3.5 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-extrabold transition cursor-pointer shadow-xs"
          >
            Ubah Menu
          </button>
        </div>
      </div>
    );
  };

  const logoSrc = leftLogo || customLogo || localStorage.getItem("kop_left_logo") || "";

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    const generateSheetData = (porsi: "besar" | "kecil") => {
      const activeItems = porsi === "besar" ? results.porsiBesarItems : results.porsiKecilItems;
      const totalBahan = porsi === "besar" ? results.totalBesarBahanCost : results.totalKecilBahanCost;
      const buffer = porsi === "besar" ? results.bufferBesarCost : results.bufferKecilCost;
      const bumbu = porsi === "besar" ? results.bumbuBesarCost : results.bumbuKecilCost;
      const subtotal = porsi === "besar" ? results.subtotalBesarCost : results.subtotalKecilCost;
      const perPorsi = porsi === "besar" ? results.costPerPorsiBesar : results.costPerPorsiKecil;
      const pmCount = porsi === "besar" ? pmBesar : pmKecil;

      const sheetData: any[][] = [];

      sheetData.push([kopLine1]);
      sheetData.push([kopLine2]);
      sheetData.push([kopLine3]);
      sheetData.push([kopLine4]);
      sheetData.push([]);
      
      sheetData.push(["LAPORAN FOOD COST (ANALISIS BIAYA & GIZI) - " + (porsi === "besar" ? "PORSI BESAR" : "PORSI KECIL")]);
      sheetData.push(["Siklus Hari Ke:", selectedDay, "Kategori:", selectedType, "Menu:", namaMenuHariIni]);
      sheetData.push(["Jumlah Penerima Manfaat (PM):", pmCount, "Penerima Manfaat"]);
      sheetData.push([]);

      sheetData.push([
        "KRITERIA PM", "MENU", "BAHAN MAKANAN", "BERAT (BB) (g)", "URT",
        "ENERGI (Kal)", "PROTEIN (g)", "LEMAK (g)", "KH (g)", "SERAT (g)",
        "BDD (%)", "BK (gr)", "Jumlah Manfaat", "Gram", "Kg",
        headerPotong, headerEkor, `Buffer (${currentDayData.bufferPct}%)`, "Jumlah+Buffer", "Harga Satuan (Rp)", "Harga Total (Rp)"
      ]);

      const defaultLabel = porsi === "besar"
        ? (porsiBesarLabels || (selectedGroup === "sekolah" ? "Siswa SD 4-6 / SMP / SMA" : selectedGroup === "tigaB" ? "Ibu Hamil / Ibu Menyusui" : "PM Porsi Besar (Custom)"))
        : (porsiKecilLabels || (selectedGroup === "sekolah" ? "Siswa TK/PAUD / SD 1-3" : selectedGroup === "tigaB" ? "Balita (6-59 Bln)" : "PM Porsi Kecil (Custom)"));

      const labelKey = `${selectedDay}_${selectedGroup}_${selectedType}_${porsi}`;
      const kriteriaPmText = customKriteriaLabels[labelKey] ?? defaultLabel;

      activeItems.forEach((b) => {
        sheetData.push([
          kriteriaPmText,
          namaMenuHariIni || "-",
          b.nama,
          b.beratBB,
          b.urt,
          b.energi,
          b.protein,
          b.lemak,
          b.kh,
          b.serat,
          b.bdd,
          b.beratKotor,
          pmCount,
          b.totalKebutuhanGram,
          b.totalKebutuhanKg,
          b.potong || 0,
          b.ekor || 0,
          b.buah || 0,
          b.butir || 0,
          b.hargaSatuan,
          b.hargaTotal
        ]);
      });

      sheetData.push([]);

      const sumEnergi = activeItems.reduce((acc, b) => acc + b.energi, 0);
      const sumProtein = activeItems.reduce((acc, b) => acc + b.protein, 0);
      const sumLemak = activeItems.reduce((acc, b) => acc + b.lemak, 0);
      const sumKH = activeItems.reduce((acc, b) => acc + b.kh, 0);
      const sumSerat = activeItems.reduce((acc, b) => acc + b.serat, 0);

      const targetEnergi = getAkgTarget(porsi, "energi");
      const targetProtein = getAkgTarget(porsi, "protein");
      const targetLemak = getAkgTarget(porsi, "lemak");
      const targetKH = getAkgTarget(porsi, "kh");
      const targetSerat = getAkgTarget(porsi, "serat");

      const pctEnergi = targetEnergi > 0 ? (sumEnergi / targetEnergi) * 100 : 0;
      const pctProtein = targetProtein > 0 ? (sumProtein / targetProtein) * 100 : 0;
      const pctLemak = targetLemak > 0 ? (sumLemak / targetLemak) * 100 : 0;
      const pctKH = targetKH > 0 ? (sumKH / targetKH) * 100 : 0;
      const pctSerat = targetSerat > 0 ? (sumSerat / targetSerat) * 100 : 0;

      sheetData.push(["KANDUNGAN GIZI", "", "", "", "", sumEnergi, sumProtein, sumLemak, sumKH, sumSerat, "", "", "JUMLAH (Rp)", totalBahan, "", "", "JUMLAH (Rp)", totalBahan]);
      sheetData.push([`KEBUTUHAN RUJUKAN AKG`, "", "", "", "", targetEnergi, targetProtein, targetLemak, targetKH, targetSerat, "", "", `BUMBU ${activeBumbuPct}% (Rp)`, totalBahan * (activeBumbuPct / 100), "", "", `BUMBU ${activeBumbuPct}% (Rp)`, totalBahan * (activeBumbuPct / 100)]);
      sheetData.push(["KEBUTUHAN 90-110%", "", "", "", "", `${pctEnergi.toFixed(0)}%`, `${pctProtein.toFixed(0)}%`, `${pctLemak.toFixed(0)}%`, `${pctKH.toFixed(0)}%`, `${pctSerat.toFixed(0)}%`, "", "", `JUMLAH+BUMBU ${activeBumbuPct}% (Rp)`, totalBahan * (1 + activeBumbuPct / 100), "", "", `JUMLAH+BUMBU ${activeBumbuPct}% (Rp)`, totalBahan * (1 + activeBumbuPct / 100)]);
      sheetData.push([]);
      sheetData.push(["", "", "", "", "", "", "", "", "", "", "", "", "HARGA PER PORSI (Rp)", perPorsi]);

      return sheetData;
    };

    if (isPerKategoriMode) {
      resolvedCategoryEntries.forEach((entry, idx) => {
        const calcInfo = categoryCalcResultsMap[entry.id];
        if (!calcInfo) return;
        const { items: activeItems, totalBahanCost: totalBahan, perPorsiCost, pmCount } = {
          items: calcInfo.items,
          totalBahanCost: calcInfo.totalBahanCost,
          perPorsiCost: calcInfo.calculatedCostPerPorsi,
          pmCount: calcInfo.pmCount,
        };
        const entryBuffer = entry.bufferPct !== undefined ? entry.bufferPct : (currentDayData.bufferPct ?? 5);

        const sheetData: any[][] = [];
        sheetData.push([kopLine1]);
        sheetData.push([kopLine2]);
        sheetData.push([kopLine3]);
        sheetData.push([kopLine4]);
        sheetData.push([]);
        sheetData.push([`LAPORAN FOOD COST (ANALISIS BIAYA & GIZI) - ${entry.title.toUpperCase()} (${entry.porsi === "besar" ? "PORSI BESAR" : "PORSI KECIL"})`]);
        sheetData.push(["Siklus Hari Ke:", selectedDay, "Kategori:", selectedType, "Menu:", namaMenuHariIni]);
        sheetData.push(["Jumlah Penerima Manfaat (PM):", pmCount, "Penerima Manfaat"]);
        sheetData.push([]);
        sheetData.push([
          "KRITERIA PM", "MENU", "BAHAN MAKANAN", "BERAT (BB) (g)", "URT",
          "ENERGI (Kal)", "PROTEIN (g)", "LEMAK (g)", "KH (g)", "SERAT (g)",
          "BDD (%)", "BK (gr)", "Jumlah Manfaat", "Gram", "Kg",
          headerPotong, headerEkor, `Buffer (${entryBuffer}%)`, "Jumlah+Buffer", "Harga Satuan (Rp)", "Harga Total (Rp)"
        ]);

        const kriteriaPmText = entry.kriteriaPm || entry.title;
        activeItems.forEach((b) => {
          sheetData.push([
            kriteriaPmText,
            namaMenuHariIni || "-",
            b.nama,
            b.beratBB,
            b.urt,
            b.energi,
            b.protein,
            b.lemak,
            b.kh,
            b.serat,
            b.bdd,
            b.beratKotor,
            pmCount,
            b.totalKebutuhanGram,
            b.totalKebutuhanKg,
            b.potong || 0,
            b.ekor || 0,
            b.buah || 0,
            b.butir || 0,
            b.hargaSatuan,
            b.hargaTotal
          ]);
        });

        sheetData.push([]);
        const sumEnergi = activeItems.reduce((acc, b) => acc + b.energi, 0);
        const sumProtein = activeItems.reduce((acc, b) => acc + b.protein, 0);
        const sumLemak = activeItems.reduce((acc, b) => acc + b.lemak, 0);
        const sumKH = activeItems.reduce((acc, b) => acc + b.kh, 0);
        const sumSerat = activeItems.reduce((acc, b) => acc + b.serat, 0);

        const targetEnergi = getCategoryEntryAkgTarget(entry, "energi");
        const targetProtein = getCategoryEntryAkgTarget(entry, "protein");
        const targetLemak = getCategoryEntryAkgTarget(entry, "lemak");
        const targetKH = getCategoryEntryAkgTarget(entry, "kh");
        const targetSerat = getCategoryEntryAkgTarget(entry, "serat");

        const pctEnergi = targetEnergi > 0 ? (sumEnergi / targetEnergi) * 100 : 0;
        const pctProtein = targetProtein > 0 ? (sumProtein / targetProtein) * 100 : 0;
        const pctLemak = targetLemak > 0 ? (sumLemak / targetLemak) * 100 : 0;
        const pctKH = targetKH > 0 ? (sumKH / targetKH) * 100 : 0;
        const pctSerat = targetSerat > 0 ? (sumSerat / targetSerat) * 100 : 0;

        sheetData.push(["KANDUNGAN GIZI", "", "", "", "", sumEnergi, sumProtein, sumLemak, sumKH, sumSerat, "", "", "JUMLAH (Rp)", totalBahan, "", "", "JUMLAH (Rp)", totalBahan]);
        sheetData.push([`KEBUTUHAN RUJUKAN AKG`, "", "", "", "", targetEnergi, targetProtein, targetLemak, targetKH, targetSerat, "", "", `BUMBU ${activeBumbuPct}% (Rp)`, totalBahan * (activeBumbuPct / 100), "", "", `BUMBU ${activeBumbuPct}% (Rp)`, totalBahan * (activeBumbuPct / 100)]);
        sheetData.push(["KEBUTUHAN 90-110%", "", "", "", "", `${pctEnergi.toFixed(0)}%`, `${pctProtein.toFixed(0)}%`, `${pctLemak.toFixed(0)}%`, `${pctKH.toFixed(0)}%`, `${pctSerat.toFixed(0)}%`, "", "", `JUMLAH+BUMBU ${activeBumbuPct}% (Rp)`, totalBahan * (1 + activeBumbuPct / 100), "", "", `JUMLAH+BUMBU ${activeBumbuPct}% (Rp)`, totalBahan * (1 + activeBumbuPct / 100)]);
        sheetData.push([]);
        sheetData.push(["", "", "", "", "", "", "", "", "", "", "", "", "HARGA PER PORSI (Rp)", perPorsiCost]);

        const wsCat = XLSX.utils.aoa_to_sheet(sheetData);
        const rawSheetName = (entry.title || `Kat ${idx + 1}`).replace(/[\\\/\?\*\[\]:]/g, "").trim();
        const sheetName = rawSheetName.substring(0, 28) || `Kategori ${idx + 1}`;
        XLSX.utils.book_append_sheet(wb, wsCat, sheetName);
      });
    } else {
      if (!appliedConfig.confirmed || appliedConfig.scope === "all" || appliedConfig.scope === "besar") {
        const besarData = generateSheetData("besar");
        const wsBesar = XLSX.utils.aoa_to_sheet(besarData);
        XLSX.utils.book_append_sheet(wb, wsBesar, "Porsi Besar");
      }

      if (!appliedConfig.confirmed || appliedConfig.scope === "all" || appliedConfig.scope === "kecil") {
        const kecilData = generateSheetData("kecil");
        const wsKecil = XLSX.utils.aoa_to_sheet(kecilData);
        XLSX.utils.book_append_sheet(wb, wsKecil, "Porsi Kecil");
      }
    }

    // Append custom tables to Excel
    customTables.forEach((table, index) => {
      const resultsCustom = table.porsi === "besar"
        ? calculateDay(table.bahanList, [], table.pmCount, 0, table.bufferPct || 5, tkpiList, activeBumbuPct)
        : calculateDay([], table.bahanList, 0, table.pmCount, table.bufferPct || 5, tkpiList, activeBumbuPct);

      const itemsCustom = table.porsi === "besar" ? resultsCustom.porsiBesarItems : resultsCustom.porsiKecilItems;
      const totalBahanCustom = table.porsi === "besar" ? resultsCustom.totalBesarBahanCost : resultsCustom.totalKecilBahanCost;
      const bufferCustom = table.porsi === "besar" ? resultsCustom.bufferBesarCost : resultsCustom.bufferKecilCost;
      const subtotalCustom = table.porsi === "besar" ? resultsCustom.subtotalBesarCost : resultsCustom.subtotalKecilCost;
      const perPorsiCustom = table.porsi === "besar" ? resultsCustom.costPerPorsiBesar : resultsCustom.costPerPorsiKecil;
      const pmCountCustom = table.pmCount;

      const sheetDataCustom: any[][] = [];

      sheetDataCustom.push([kopLine1]);
      sheetDataCustom.push([kopLine2]);
      sheetDataCustom.push([kopLine3]);
      sheetDataCustom.push([kopLine4]);
      sheetDataCustom.push([]);
      
      sheetDataCustom.push(["LAPORAN FOOD COST (ANALISIS BIAYA & GIZI) - " + table.namaTabel.toUpperCase()]);
      sheetDataCustom.push(["Siklus Hari Ke:", selectedDay, "Kategori:", selectedType, "Porsi:", table.porsi.toUpperCase()]);
      sheetDataCustom.push(["Jumlah Penerima Manfaat (PM):", pmCountCustom, "Penerima Manfaat"]);
      sheetDataCustom.push([]);

      sheetDataCustom.push([
        "KRITERIA PM", "MENU", "BAHAN MAKANAN", "BERAT (BB) (g)", "URT",
        "ENERGI (Kal)", "PROTEIN (g)", "LEMAK (g)", "KH (g)", "SERAT (g)",
        "BDD (%)", "BK (gr)", "Jumlah Manfaat", "Gram", "Kg",
        "Potong", "Ekor", "Buah", "Butir", "Harga Satuan (Rp)", "Harga Total (Rp)"
      ]);

      const tableKriteria = table.kriteriaPm || (table.porsi === "besar" ? "Kriteria PM Besar" : "Kriteria PM Kecil");

      itemsCustom.forEach((b) => {
        sheetDataCustom.push([
          tableKriteria,
          namaMenuHariIni || "-",
          b.nama,
          b.beratBB,
          b.urt,
          b.energi,
          b.protein,
          b.lemak,
          b.kh,
          b.serat,
          b.bdd,
          b.beratKotor,
          pmCountCustom,
          b.totalKebutuhanGram,
          b.totalKebutuhanKg,
          b.potong || 0,
          b.ekor || 0,
          b.buah || 0,
          b.butir || 0,
          b.hargaSatuan,
          b.hargaTotal
        ]);
      });

      sheetDataCustom.push([]);

      const sumEnergiCustom = itemsCustom.reduce((acc, b) => acc + b.energi, 0);
      const sumProteinCustom = itemsCustom.reduce((acc, b) => acc + b.protein, 0);
      const sumLemakCustom = itemsCustom.reduce((acc, b) => acc + b.lemak, 0);
      const sumKHCustom = itemsCustom.reduce((acc, b) => acc + b.kh, 0);
      const sumSeratCustom = itemsCustom.reduce((acc, b) => acc + b.serat, 0);

      const tableAkgType = table.akgType || (table.porsi === "besar" ? "sd_besar" : "sd_kecil");
      const tableAkgMode = table.akgMode || "avg";
      const getCustomTableAkgTarget = (key: string): number => {
        if (tableAkgType === "custom" || tableAkgMode === "custom") {
          const customMap = table.porsi === "besar" ? customBesar : customKecil;
          return customMap[key] || 0;
        }
        const config = TARGET_AKG_LIMITS[tableAkgType];
        if (!config) return 0;

        const minVal = key === "energi" ? config.energiMin : (config as any)[`${key}Min`] || 0;
        const ratio = config.energiMax / config.energiMin;

        if (tableAkgMode === "min") {
          return minVal;
        }
        if (tableAkgMode === "max") {
          if (key === "energi") return config.energiMax;
          return Math.round(minVal * ratio * 10) / 10;
        }
        // "avg" mode
        if (key === "energi") {
          return (config.energiMin + config.energiMax) / 2;
        }
        const maxVal = minVal * ratio;
        return Math.round(((minVal + maxVal) / 2) * 10) / 10;
      };

      const targetEnergiCustom = getCustomTableAkgTarget("energi");
      const targetProteinCustom = getCustomTableAkgTarget("protein");
      const targetLemakCustom = getCustomTableAkgTarget("lemak");
      const targetKHCustom = getCustomTableAkgTarget("kh");
      const targetSeratCustom = getCustomTableAkgTarget("serat");

      const pctEnergiCustom = targetEnergiCustom > 0 ? (sumEnergiCustom / targetEnergiCustom) * 100 : 0;
      const pctProteinCustom = targetProteinCustom > 0 ? (sumProteinCustom / targetProteinCustom) * 100 : 0;
      const pctLemakCustom = targetLemakCustom > 0 ? (sumLemakCustom / targetLemakCustom) * 100 : 0;
      const pctKHCustom = targetKHCustom > 0 ? (sumKHCustom / targetKHCustom) * 100 : 0;
      const pctSeratCustom = targetSeratCustom > 0 ? (sumSeratCustom / targetSeratCustom) * 100 : 0;

      sheetDataCustom.push(["KANDUNGAN GIZI", "", "", "", "", sumEnergiCustom, sumProteinCustom, sumLemakCustom, sumKHCustom, sumSeratCustom, "", "", "JUMLAH (Rp)", totalBahanCustom, "", "", "JUMLAH (Rp)", totalBahanCustom]);
      sheetDataCustom.push([`KEBUTUHAN RUJUKAN AKG`, "", "", "", "", targetEnergiCustom, targetProteinCustom, targetLemakCustom, targetKHCustom, targetSeratCustom, "", "", `BUMBU ${activeBumbuPct}% (Rp)`, totalBahanCustom * (activeBumbuPct / 100), "", "", `BUMBU ${activeBumbuPct}% (Rp)`, totalBahanCustom * (activeBumbuPct / 100)]);
      sheetDataCustom.push(["KEBUTUHAN 90-110%", "", "", "", "", `${pctEnergiCustom.toFixed(0)}%`, `${pctProteinCustom.toFixed(0)}%`, `${pctLemakCustom.toFixed(0)}%`, `${pctKHCustom.toFixed(0)}%`, `${pctSeratCustom.toFixed(0)}%`, "", "", `JUMLAH+BUMBU ${activeBumbuPct}% (Rp)`, totalBahanCustom * (1 + activeBumbuPct / 100), "", "", `JUMLAH+BUMBU ${activeBumbuPct}% (Rp)`, totalBahanCustom * (1 + activeBumbuPct / 100)]);
      sheetDataCustom.push([]);
      sheetDataCustom.push(["", "", "", "", "", "", "", "", "", "", "", "", "HARGA PER PORSI (Rp)", perPorsiCustom]);

      const wsCustom = XLSX.utils.aoa_to_sheet(sheetDataCustom);
      const cleanSheetName = table.namaTabel.substring(0, 25).replace(/[\\\/\?\*\[\]]/g, "") || `Tabel ${index + 1}`;
      XLSX.utils.book_append_sheet(wb, wsCustom, cleanSheetName);
    });

    XLSX.writeFile(wb, `Food_Cost_Hari_${selectedDay}_${selectedGroup}_${selectedType}.xlsx`);
  };

  const renderFoodCostTable = (porsi: "besar" | "kecil", isPrint: boolean) => {
    const items = porsi === "besar" ? results.porsiBesarItems : results.porsiKecilItems;
    const totalBahanCost = porsi === "besar" ? results.totalBesarBahanCost : results.totalKecilBahanCost;
    const bufferCost = porsi === "besar" ? results.bufferBesarCost : results.bufferKecilCost;
    const bumbuCost = porsi === "besar" ? results.bumbuBesarCost : results.bumbuKecilCost;
    const subtotalCost = porsi === "besar" ? results.subtotalBesarCost : results.subtotalKecilCost;
    const costPerPorsi = porsi === "besar" ? results.costPerPorsiBesar : results.costPerPorsiKecil;
    const recipientCount = porsi === "besar" ? pmBesar : pmKecil;
    const targetPrice = porsi === "besar" ? targetPriceBesar : targetPriceKecil;

    const calculatedCostPerPorsi = planningMode === "all_ingredients"
      ? (recipientCount > 0 ? totalBahanCost / recipientCount : 0)
      : costPerPorsi;

    const akgType = porsi === "besar" ? besarAkgType : kecilAkgType;
    const setAkgType = porsi === "besar" ? setBesarAkgType : setKecilAkgType;
    const customMap = porsi === "besar" ? customBesar : customKecil;
    const setCustomMap = porsi === "besar" ? setCustomBesar : setCustomKecil;

    const sumEnergi = items.reduce((acc, b) => acc + b.energi, 0);
    const sumProtein = items.reduce((acc, b) => acc + b.protein, 0);
    const sumLemak = items.reduce((acc, b) => acc + b.lemak, 0);
    const sumKH = items.reduce((acc, b) => acc + b.kh, 0);
    const sumSerat = items.reduce((acc, b) => acc + b.serat, 0);

    const targetEnergi = getAkgTarget(porsi, "energi");
    const targetProtein = getAkgTarget(porsi, "protein");
    const targetLemak = getAkgTarget(porsi, "lemak");
    const targetKH = getAkgTarget(porsi, "kh");
    const targetSerat = getAkgTarget(porsi, "serat");

    const pctEnergi = targetEnergi > 0 ? (sumEnergi / targetEnergi) * 100 : 0;
    const pctProtein = targetProtein > 0 ? (sumProtein / targetProtein) * 100 : 0;
    const pctLemak = targetLemak > 0 ? (sumLemak / targetLemak) * 100 : 0;
    const pctKH = targetKH > 0 ? (sumKH / targetKH) * 100 : 0;
    const pctSerat = targetSerat > 0 ? (sumSerat / targetSerat) * 100 : 0;

    const diffPrice = calculatedCostPerPorsi - targetPrice;
    const isOverBudget = diffPrice > 0;

    const defaultLabel = porsi === "besar"
      ? (porsiBesarLabels || (selectedGroup === "sekolah" ? "Siswa SD 4-6 / SMP / SMA" : selectedGroup === "tigaB" ? "Ibu Hamil / Ibu Menyusui" : "PM Porsi Besar (Custom)"))
      : (porsiKecilLabels || (selectedGroup === "sekolah" ? "Siswa TK/PAUD / SD 1-3" : selectedGroup === "tigaB" ? "Balita (6-59 Bln)" : "PM Porsi Kecil (Custom)"));

    const labelKey = `${selectedDay}_${selectedGroup}_${selectedType}_${porsi}`;
    const kriteriaPmText = customKriteriaLabels[labelKey] ?? defaultLabel;

    if (items.length === 0) {
      if (isPrint) {
        return (
          <div className="p-4 border border-black rounded-lg text-center text-xs font-bold text-slate-500 italic uppercase bg-slate-50">
            Tidak ada komponen bahan baku untuk porsi {porsi} hari ini.
          </div>
        );
      }
      return (
        <div className="flex flex-col items-center justify-center p-8 bg-slate-50 border border-dashed border-slate-300 rounded-2xl space-y-3 text-center my-2">
          <div className="p-3 bg-rose-50 text-rose-500 rounded-full animate-pulse">
            <Trash2 className="w-6 h-6" />
          </div>
          <div>
            <h5 className="text-xs font-black text-slate-700 uppercase tracking-wide">
              Komponen Bahan Baku Kosong
            </h5>
            <p className="text-[11px] text-slate-500 mt-1 max-w-sm">
              Seluruh bahan baku telah dihapus. Silakan klik tombol <strong className="text-indigo-600 font-extrabold">+ Tambah Bahan Baku</strong> di atas untuk menginput ulang.
            </p>
          </div>
        </div>
      );
    }

    const bbHighlight = !isPrint && Boolean(pastedHighlightMap[porsi]);

    return (
      <>
      <div className={isPrint ? "w-full border border-black rounded-lg shadow-xs overflow-hidden" : "overflow-x-auto border border-black rounded-lg shadow-xs"}>
        <table className={`w-full border-collapse border border-black font-sans text-center ${isPrint ? "table-fixed text-[9px] leading-tight" : "text-xs"}`} style={isPrint ? { tableLayout: 'fixed', width: '100%' } : { fontSize: '12px' }}>
          <thead className="bg-[#92D050] text-black font-bold uppercase border border-black text-center">
            <tr>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[110px] min-w-[110px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>KRITERIA PM</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[110px] min-w-[110px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>MENU</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[160px] min-w-[160px]"}`} style={isPrint ? { width: '10%' } : undefined}>BAHAN MAKANAN</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[58px] min-w-[58px]"} ${bbHighlight ? "bg-amber-300 ring-2 ring-inset ring-amber-600" : ""}`} style={isPrint ? { width: '3.5%' } : undefined}>
                BERAT (BB){bbHighlight && <span className="block text-[8px] font-black text-amber-950 normal-case">Ubah ↵</span>}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[58px] min-w-[58px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>URT</th>
              <th colSpan={5} className="p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap bg-[#76933C] print:text-black">KOMPOSISI ZAT GIZI MAKANAN</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[58px] min-w-[58px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>BDD (%)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>BK (g)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black leading-tight ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>Jumlah<br/>Manfaat</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[76px] min-w-[76px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>Gram (g)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[76px] min-w-[76px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>Kg (kg)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap bg-[#92D050] ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>
                {isPrint ? (
                  headerPotong
                ) : (
                  <input
                    type="text"
                    value={headerPotong}
                    onChange={(e) => {
                      setHeaderPotong(e.target.value);
                      localStorage.setItem("fc_header_potong", e.target.value);
                    }}
                    className="w-full text-center bg-[#92D050] border-0 hover:bg-[#86c145] focus:bg-white font-black p-0.5 rounded text-slate-950 uppercase text-[11px] focus:ring-1 focus:ring-slate-600"
                    title="Ubah judul kolom Potong"
                  />
                )}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap bg-[#92D050] ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>
                {isPrint ? (
                  headerEkor
                ) : (
                  <input
                    type="text"
                    value={headerEkor}
                    onChange={(e) => {
                      setHeaderEkor(e.target.value);
                      localStorage.setItem("fc_header_ekor", e.target.value);
                    }}
                    className="w-full text-center bg-[#92D050] border-0 hover:bg-[#86c145] focus:bg-white font-black p-0.5 rounded text-slate-950 uppercase text-[11px] focus:ring-1 focus:ring-slate-600"
                    title="Ubah judul kolom Ekor"
                  />
                )}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black bg-[#92D050] text-[11px] leading-tight select-none ${isPrint ? "" : "w-[80px] min-w-[80px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>
                Buffer<br/><span className="text-[10px] font-bold">({currentDayData.bufferPct}%)</span>
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black bg-[#92D050] text-[11px] leading-tight select-none ${isPrint ? "" : "w-[90px] min-w-[90px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>
                Jumlah +<br/>Buffer
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[130px] min-w-[130px]"}`} style={isPrint ? { width: '7.5%' } : undefined}>Harga Satuan</th>
              <th rowSpan={2} className={`p-1 border border-black border-r-2 border-r-slate-950 text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[136px] min-w-[136px]"}`} style={isPrint ? { width: '8.5%' } : undefined}>Harga Total</th>
              {!isPrint && <th rowSpan={2} className="p-1 border border-black border-l-2 border-l-slate-950 text-center text-rose-800 w-[44px] min-w-[44px] no-print">Aksi</th>}
            </tr>
            <tr>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[54px] min-w-[54px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>ENERGI (Kal)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>Protein (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>LEMAK (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>KH (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>SERAT (g)</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-black text-[11px] text-slate-900 bg-white">
            {items.length === 0 ? (
              <tr>
                <td colSpan={isPrint ? 21 : 22} className="p-6 text-center text-slate-400 italic bg-slate-50 border border-black">
                  Belum ada bahan makanan yang diinput untuk porsi {porsi}.
                </td>
              </tr>
            ) : (
              items.map((b, idx) => (
                <tr key={b.id} className="hover:bg-slate-50/50">
                  {idx === 0 && (
                    <td rowSpan={items.length} className="p-1 border border-black text-center font-extrabold bg-[#E2EFDA] align-middle text-[10px] uppercase leading-tight w-[110px] min-w-[110px]">
                      {isPrint ? (
                        <span className="block p-1 text-[10px] font-extrabold text-slate-950 uppercase">{kriteriaPmText}</span>
                      ) : (
                        <textarea
                          value={kriteriaPmText}
                          onChange={(e) => saveCustomKriteriaLabel(labelKey, e.target.value)}
                          className="w-full text-center bg-transparent border-0 font-extrabold text-[10px] uppercase focus:ring-1 focus:ring-indigo-500 p-1 resize-none overflow-hidden"
                          rows={3}
                          placeholder="Edit Kriteria PM..."
                        />
                      )}
                    </td>
                  )}
                  {idx === 0 && (
                    <td rowSpan={items.length} className="p-1 border border-black text-center font-bold bg-[#F2F2F2] align-middle text-[10px] w-[110px] min-w-[110px] break-words text-slate-950">
                      {namaMenuHariIni || "-"}
                    </td>
                  )}
                  <td className="p-1 border border-black font-semibold w-[160px] min-w-[160px] text-center">
                    {isPrint ? (
                      <span className="p-0.5 block text-xs font-bold text-slate-950 text-center">{b.nama}</span>
                    ) : (
                      <div className="flex items-center gap-1.5 justify-center">
                        <div className="flex-1 min-w-0">
                          <SearchableTkpiDropdown
                            id={`fc-select-tkpi-${porsi}-${idx}`}
                            tkpiList={tkpiList}
                            selectedValue={currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.tkpiId || ""}
                            onChange={(val) => handleRowChange(porsi, idx, "tkpiId", val)}
                            minimal
                          />
                        </div>
                        <button
                          type="button"
                          onClick={() => removeIngredientRow(porsi, b.id, idx)}
                          className="p-1 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg transition hover:scale-105 active:scale-95 shrink-0 no-print"
                          title="Hapus Bahan Makanan Ini"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    )}
                  </td>
                  <td className={`p-1 border border-black text-center w-[56px] min-w-[56px] ${bbHighlight ? "bg-amber-50 ring-1 ring-inset ring-amber-400" : ""}`}>
                    {isPrint ? (
                      <span className="font-mono font-bold text-slate-950 text-xs text-center">{b.beratBB}</span>
                    ) : (
                      <input
                        id={`fc-input-bb-${porsi}-${idx}`}
                        type="number"
                        value={currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.beratBB || ""}
                        onChange={(e) => handleRowChange(porsi, idx, "beratBB", Number(e.target.value))}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") {
                            e.preventDefault();
                            const nextEl = document.getElementById(`fc-input-bb-${porsi}-${idx + 1}`) as HTMLInputElement | null;
                            if (nextEl) {
                              nextEl.focus();
                              nextEl.select();
                            }
                          }
                        }}
                        className={`w-full text-center border-0 font-mono font-bold p-0 text-slate-950 text-xs rounded ${bbHighlight ? "bg-amber-100/80 focus:bg-white focus:ring-2 focus:ring-amber-500" : "bg-transparent"}`}
                      />
                    )}
                  </td>
                  <td className="p-1 border border-black text-center w-[56px] min-w-[56px]">
                    {isPrint ? (
                      <span className="text-slate-950 text-xs text-center">{b.urt}</span>
                    ) : (
                      <input
                        id={`fc-input-urt-${porsi}-${idx}`}
                        type="text"
                        value={currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.urt || ""}
                        onChange={(e) => handleRowChange(porsi, idx, "urt", e.target.value)}
                        className="w-full text-center bg-transparent border-0 p-0 text-slate-950 text-xs text-center"
                      />
                    )}
                  </td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[54px] min-w-[54px] text-xs">{b.energi.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.protein.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.lemak.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.kh.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.serat.toFixed(1)}</td>
                  
                  <td className="p-1 border border-black text-center w-[56px] min-w-[56px]">
                    {isPrint ? (
                      <span className="font-mono font-bold text-slate-950 text-xs">{b.bdd}%</span>
                    ) : (
                      <EditableBddCell
                        value={b.bdd}
                        isCustom={currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.bdd !== undefined}
                        onChange={(newVal) => handleRowChange(porsi, idx, "bdd", newVal)}
                      />
                    )}
                  </td>
                  <td className="p-1 border border-black text-center font-mono text-indigo-955 font-bold bg-[#E6F0FA]/20 w-[64px] min-w-[64px] text-slate-950 text-xs">{b.beratKotor.toFixed(1)}</td>
                  
                  {idx === 0 && (
                    <td rowSpan={items.length} className="p-1 border border-black text-center font-extrabold bg-[#FFF2CC] align-middle font-mono w-[64px] min-w-[64px] text-slate-950 text-xs">
                      {isPrint ? (
                        <span className="font-mono font-extrabold text-slate-950 text-xs">{recipientCount}</span>
                      ) : (
                        <input
                          type="number"
                          min="0"
                          value={
                            porsi === "besar"
                              ? (currentDayData.customPmBesarCount !== undefined ? currentDayData.customPmBesarCount : recipientCount)
                              : (currentDayData.customPmKecilCount !== undefined ? currentDayData.customPmKecilCount : recipientCount)
                          }
                          onChange={(e) => {
                            const raw = e.target.value;
                            const val = raw === "" ? undefined : Math.max(0, Number(raw));
                            if (porsi === "besar") {
                              updateDayData({ ...currentDayData, customPmBesarCount: val });
                            } else {
                              updateDayData({ ...currentDayData, customPmKecilCount: val });
                            }
                          }}
                          className="w-full text-center bg-transparent border-0 font-mono font-extrabold p-0.5 text-slate-950 text-xs focus:bg-white focus:ring-1 focus:ring-amber-500 rounded cursor-pointer hover:bg-amber-100/60"
                          title="Klik untuk mengedit atau menghapus/mengubah Jumlah Penerima Manfaat"
                        />
                      )}
                    </td>
                  )}
                  
                  <td className="p-1 border border-black text-center font-mono text-slate-950 w-[72px] min-w-[72px] font-semibold text-xs">{b.totalKebutuhanGram.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono font-bold text-slate-950 w-[72px] min-w-[72px] text-xs">{b.totalKebutuhanKg.toFixed(1)}</td>
                  
                   <td className="p-1 border border-black text-center w-[68px] min-w-[68px]">
                    {isPrint ? (
                      <span className="font-mono text-slate-950">{b.potong || "-"}</span>
                    ) : (
                      <div className="relative group/calc flex items-center justify-center h-full w-full">
                        <input
                          type="text"
                          value={currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.potong ?? ""}
                          onChange={(e) => handleRowChange(porsi, idx, "potong", e.target.value)}
                          className="w-full text-center bg-white border border-slate-900 rounded-lg font-mono font-bold text-xs p-1 focus:ring-2 focus:ring-indigo-600 focus:outline-none shadow-2xs"
                        />
                        <button
                          type="button"
                          onClick={() => openCalculator(porsi, idx, "potong", String(currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.potong ?? ""), b.nama)}
                          className="absolute right-0.5 top-1/2 -translate-y-1/2 opacity-0 group-hover/calc:opacity-100 focus:opacity-100 bg-white hover:bg-slate-100 border border-slate-300 shadow-xs px-1 py-0.5 rounded text-[10px] text-slate-600 z-10 transition-opacity"
                          title={`Hitung Potong - ${b.nama}`}
                        >
                          🧮
                        </button>
                      </div>
                    )}
                  </td>
                  <td className="p-1 border border-black text-center w-[68px] min-w-[68px]">
                    {isPrint ? (
                      <span className="font-mono text-slate-950">{b.ekor || "-"}</span>
                    ) : (
                      <div className="relative group/calc flex items-center justify-center h-full w-full">
                        <input
                          type="text"
                          value={currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.ekor ?? ""}
                          onChange={(e) => handleRowChange(porsi, idx, "ekor", e.target.value)}
                          className="w-full text-center bg-white border border-slate-900 rounded-lg font-mono font-bold text-xs p-1 focus:ring-2 focus:ring-indigo-600 focus:outline-none shadow-2xs"
                        />
                        <button
                          type="button"
                          onClick={() => openCalculator(porsi, idx, "ekor", String(currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.ekor ?? ""), b.nama)}
                          className="absolute right-0.5 top-1/2 -translate-y-1/2 opacity-0 group-hover/calc:opacity-100 focus:opacity-100 bg-white hover:bg-slate-100 border border-slate-300 shadow-xs px-1 py-0.5 rounded text-[10px] text-slate-600 z-10 transition-opacity"
                          title={`Hitung Ekor - ${b.nama}`}
                        >
                          🧮
                        </button>
                      </div>
                    )}
                  </td>
                  <td
                    onClick={() => !isPrint && openBufferConfig(porsi, idx, b)}
                    className={`p-1 border border-black text-center w-[80px] min-w-[80px] bg-slate-50/50 ${!isPrint ? "cursor-pointer hover:bg-indigo-50/70 transition-all group/buffer" : ""}`}
                    title={!isPrint ? "Klik untuk memilih basis Buffer (Kg, Potong, Ekor)" : undefined}
                  >
                    <div className="flex flex-col items-center justify-center h-full min-h-[36px]">
                      <span className="font-mono text-slate-950 text-xs font-semibold">
                        {typeof b.buah === "number" && b.buah >= 0 ? b.buah.toFixed(1) : "-"}
                      </span>
                      {!isPrint && (
                        <span className="text-[8px] text-indigo-600 font-extrabold uppercase tracking-wider block leading-none mt-0.5 opacity-70 group-hover/buffer:opacity-100 transition-opacity">
                          {b.bufferBase === "kg" ? "KG" : b.bufferBase === "potong" ? "Ptg" : b.bufferBase === "ekor" ? "Ekr" : b.bufferBase === "custom" ? "Cust" : "Auto"} ⚙️
                        </span>
                      )}
                    </div>
                  </td>
                  <td
                    onClick={() => !isPrint && openBufferConfig(porsi, idx, b)}
                    className={`p-1 border border-black text-center w-[90px] min-w-[90px] bg-[#E6F0FA]/30 ${!isPrint ? "cursor-pointer hover:bg-indigo-50/70 transition-all group/jumlah" : ""}`}
                    title={!isPrint ? "Klik untuk mengatur rumus Jumlah + Buffer" : undefined}
                  >
                    <div className="flex flex-col items-center justify-center h-full min-h-[36px]">
                      <span className="font-mono text-slate-950 text-xs font-bold text-indigo-950">
                        {typeof b.butir === "number" && b.butir >= 0 ? b.butir.toFixed(1) : "-"}
                      </span>
                      {!isPrint && (
                        <span className="text-[8px] text-indigo-700 font-extrabold uppercase tracking-wider block leading-none mt-0.5 opacity-70 group-hover/jumlah:opacity-100 transition-opacity">
                          {(() => {
                            const val = b.jumlahBufferChoice || "auto";
                            if (val === "auto") return "Auto";
                            if (val === "kg_with") return "KG + Buf";
                            if (val === "kg_without") return "KG Saja";
                            if (val === "potong_with") return "Ptg + Buf";
                            if (val === "potong_without") return "Ptg Saja";
                            if (val === "ekor_with") return "Ekr + Buf";
                            if (val === "ekor_without") return "Ekr Saja";
                            if (val === "custom_with") return "Cust + Buf";
                            if (val === "custom_without") return "Cust Saja";
                            return val;
                          })()} ⚙️
                        </span>
                      )}
                    </div>
                  </td>
                  
                  <td className="p-1 border border-black text-center w-[130px] min-w-[130px]">
                    {isPrint ? (
                      <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-bold text-slate-950">
                        <span className="text-slate-500">Rp</span>
                        <span>{formatRupiah(b.hargaSatuan).replace("Rp", "").trim()}</span>
                      </div>
                    ) : (
                      <div className="flex items-center justify-center gap-1 w-full">
                        <div className="flex items-center gap-1 bg-white/50 px-1 rounded border border-slate-200 w-full max-w-[84px]">
                          <span className="text-[10px] font-bold text-slate-400">Rp</span>
                          <input
                            id={`fc-input-price-${porsi}-${idx}`}
                            type="text"
                            inputMode="numeric"
                            value={formatThousandSeparator(currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.hargaSatuan)}
                            onChange={(e) => handleRowChange(porsi, idx, "hargaSatuan", parseThousandSeparator(e.target.value))}
                            className="w-full text-right bg-transparent border-0 font-mono font-bold p-0 focus:ring-0 text-slate-950 text-xs"
                          />
                        </div>
                        <PriceCalculatorPopover
                          initialValue={currentDayData[porsi === "besar" ? "porsiBesarBahan" : "porsiKecilBahan"][idx]?.hargaSatuan || 0}
                          onApply={(val) => handleRowChange(porsi, idx, "hargaSatuan", val)}
                          placeholder="Hitung Harga Satuan"
                          ingredientName={b.nama}
                        />
                      </div>
                    )}
                  </td>
                  <td className="p-1 border border-black border-r-2 border-r-slate-950 text-center w-[136px] min-w-[136px]">
                    <div className="flex items-center justify-between px-1 w-full h-full">
                      <span className="text-slate-500 text-[10px] font-bold">Rp</span>
                      <div className="flex items-center gap-1 font-mono text-xs font-bold text-slate-950 whitespace-nowrap">
                        <span>{formatRupiah(b.hargaTotal).replace("Rp", "").trim()}</span>
                      </div>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className="p-1 border border-black border-l-2 border-l-slate-950 text-center w-[44px] min-w-[44px] no-print">
                      <button
                        id={`fc-btn-del-${porsi}-${idx}`}
                        type="button"
                        onClick={() => removeIngredientRow(porsi, b.id, idx)}
                        className="p-0.5 text-slate-400 hover:text-rose-600 rounded"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </td>
                  )}
                </tr>
              ))
            )}
          </tbody>
          
          <tfoot className="border border-black font-semibold text-slate-950 text-xs text-center">
            {/* ROW 1: KANDUNGAN GIZI + JUMLAH */}
            <tr className="border border-black">
              {/* Left Part: Kandungan Gizi (cols 1-10) */}
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                KANDUNGAN GIZI
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[54px] min-w-[54px] bg-[#92D050]">
                {sumEnergi.toFixed(1)}
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {sumProtein.toFixed(1)}
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {sumLemak.toFixed(1)}
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {sumKH.toFixed(1)}
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {sumSerat.toFixed(1)}
              </td>

              {/* Right Part: JUMLAH (cols 11-22) */}
              <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#BFBFBF]">
                JUMLAH
              </td>
              <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#BFBFBF]">
                <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                  <span className="text-slate-600 font-black">Rp</span>
                  <span className="font-black">{formatRupiah(totalBahanCost).replace("Rp", "").trim()}</span>
                </div>
              </td>
              {!isPrint && (
                <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#BFBFBF] no-print"></td>
              )}
            </tr>

            {/* ROW 2: KEBUTUHAN RUJUKAN AKG + (BUMBU 10% OR HARGA PER PORSI) */}
            <tr className="border border-black">
              {/* Left Part: Rujukan AKG (cols 1-10) */}
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                <div className="flex items-center justify-center gap-1.5 flex-wrap">
                  <span>KEBUTUHAN RUJUKAN AKG PORSI {porsi === "besar" ? "BESAR" : "KECIL"}</span>
                  {!isPrint && (
                    <select
                      value={porsi === "besar" ? besarAkgMode : kecilAkgMode}
                      onChange={(e) => {
                        const newMode = e.target.value;
                        if (porsi === "besar") {
                          setBesarAkgMode(newMode);
                          setBesarNutrientModes({ energi: newMode, protein: newMode, lemak: newMode, kh: newMode, serat: newMode });
                        } else {
                          setKecilAkgMode(newMode);
                          setKecilNutrientModes({ energi: newMode, protein: newMode, lemak: newMode, kh: newMode, serat: newMode });
                        }
                      }}
                      className="bg-white hover:bg-slate-50 text-slate-900 text-[11px] font-black rounded-lg px-2 py-0.5 border border-indigo-400 shadow-xs focus:outline-none focus:ring-2 focus:ring-indigo-500 uppercase cursor-pointer"
                    >
                      <option value="">-- SEMUA RUJUKAN --</option>
                      <option value="min">SEMUA MIN</option>
                      <option value="max">SEMUA MAX</option>
                      <option value="avg">SEMUA RATA-RATA</option>
                      <option value="custom">SEMUA KUSTOM ✍️</option>
                    </select>
                  )}
                </div>
              </td>
              
              {/* ENERGI */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[58px] min-w-[58px] bg-[#92D050]">
                {!isPrint && (
                  <select
                    value={porsi === "besar" ? besarNutrientModes.energi : kecilNutrientModes.energi}
                    onChange={(e) => updateIndividualNutrientMode(porsi, "energi", e.target.value)}
                    className="w-full text-[8.5px] font-black uppercase bg-white border border-emerald-700/50 rounded px-0.5 py-0.5 mb-1 text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-600 shadow-2xs cursor-pointer"
                    title="Pilih rujukan Energi"
                  >
                    <option value="min">MIN</option>
                    <option value="max">MAX</option>
                    <option value="avg">RATA2</option>
                    <option value="custom">KUST</option>
                  </select>
                )}
                {isPrint ? (
                  <span>{targetEnergi > 0 ? targetEnergi.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={targetEnergi > 0 ? Number(targetEnergi.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      updateIndividualNutrientMode(porsi, "energi", "custom");
                      handleAkgValChange(porsi, "energi", val, "custom");
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-0.5 rounded-md text-slate-900 text-[11px] shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {/* PROTEIN */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[58px] min-w-[58px] bg-[#92D050]">
                {!isPrint && (
                  <select
                    value={porsi === "besar" ? besarNutrientModes.protein : kecilNutrientModes.protein}
                    onChange={(e) => updateIndividualNutrientMode(porsi, "protein", e.target.value)}
                    className="w-full text-[8.5px] font-black uppercase bg-white border border-emerald-700/50 rounded px-0.5 py-0.5 mb-1 text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-600 shadow-2xs cursor-pointer"
                    title="Pilih rujukan Protein"
                  >
                    <option value="min">MIN</option>
                    <option value="max">MAX</option>
                    <option value="avg">RATA2</option>
                    <option value="custom">KUST</option>
                  </select>
                )}
                {isPrint ? (
                  <span>{targetProtein > 0 ? targetProtein.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={targetProtein > 0 ? Number(targetProtein.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      updateIndividualNutrientMode(porsi, "protein", "custom");
                      handleAkgValChange(porsi, "protein", val, "custom");
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-0.5 rounded-md text-slate-900 text-[11px] shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {/* LEMAK */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[58px] min-w-[58px] bg-[#92D050]">
                {!isPrint && (
                  <select
                    value={porsi === "besar" ? besarNutrientModes.lemak : kecilNutrientModes.lemak}
                    onChange={(e) => updateIndividualNutrientMode(porsi, "lemak", e.target.value)}
                    className="w-full text-[8.5px] font-black uppercase bg-white border border-emerald-700/50 rounded px-0.5 py-0.5 mb-1 text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-600 shadow-2xs cursor-pointer"
                    title="Pilih rujukan Lemak"
                  >
                    <option value="min">MIN</option>
                    <option value="max">MAX</option>
                    <option value="avg">RATA2</option>
                    <option value="custom">KUST</option>
                  </select>
                )}
                {isPrint ? (
                  <span>{targetLemak > 0 ? targetLemak.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={targetLemak > 0 ? Number(targetLemak.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      updateIndividualNutrientMode(porsi, "lemak", "custom");
                      handleAkgValChange(porsi, "lemak", val, "custom");
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-0.5 rounded-md text-slate-900 text-[11px] shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {/* KH */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[58px] min-w-[58px] bg-[#92D050]">
                {!isPrint && (
                  <select
                    value={porsi === "besar" ? besarNutrientModes.kh : kecilNutrientModes.kh}
                    onChange={(e) => updateIndividualNutrientMode(porsi, "kh", e.target.value)}
                    className="w-full text-[8.5px] font-black uppercase bg-white border border-emerald-700/50 rounded px-0.5 py-0.5 mb-1 text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-600 shadow-2xs cursor-pointer"
                    title="Pilih rujukan Karbohidrat (KH)"
                  >
                    <option value="min">MIN</option>
                    <option value="max">MAX</option>
                    <option value="avg">RATA2</option>
                    <option value="custom">KUST</option>
                  </select>
                )}
                {isPrint ? (
                  <span>{targetKH > 0 ? targetKH.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={targetKH > 0 ? Number(targetKH.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      updateIndividualNutrientMode(porsi, "kh", "custom");
                      handleAkgValChange(porsi, "kh", val, "custom");
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-0.5 rounded-md text-slate-900 text-[11px] shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {/* SERAT */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[58px] min-w-[58px] bg-[#92D050]">
                {!isPrint && (
                  <select
                    value={porsi === "besar" ? besarNutrientModes.serat : kecilNutrientModes.serat}
                    onChange={(e) => updateIndividualNutrientMode(porsi, "serat", e.target.value)}
                    className="w-full text-[8.5px] font-black uppercase bg-white border border-emerald-700/50 rounded px-0.5 py-0.5 mb-1 text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-600 shadow-2xs cursor-pointer"
                    title="Pilih rujukan Serat"
                  >
                    <option value="min">MIN</option>
                    <option value="max">MAX</option>
                    <option value="avg">RATA2</option>
                    <option value="custom">KUST</option>
                  </select>
                )}
                {isPrint ? (
                  <span>{targetSerat > 0 ? targetSerat.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={targetSerat > 0 ? Number(targetSerat.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      updateIndividualNutrientMode(porsi, "serat", "custom");
                      handleAkgValChange(porsi, "serat", val, "custom");
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-0.5 rounded-md text-slate-900 text-[11px] shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {planningMode === "all_ingredients" ? (
                <>
                  {/* Right Part: HARGA PER PORSI (cols 11-22) */}
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#A9D08E]">
                    HARGA PER PORSI
                  </td>
                  <td className={`p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}>
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950 print:text-black">
                      <span className="text-slate-950 print:text-black font-black">Rp</span>
                      <span className="font-black text-slate-950 print:text-black">{formatRupiah(calculatedCostPerPorsi).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className={`p-1 border border-black border-l-2 border-l-slate-950 no-print ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}></td>
                  )}
                </>
              ) : (
                <>
                  {/* Right Part: BUMBU CUSTOM % (cols 11-22) */}
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#FFFF00]">
                    BUMBU {activeBumbuPct}%
                  </td>
                  <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#FFFF00]">
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                      <span className="text-slate-600 font-black">Rp</span>
                      <span className="font-black">{formatRupiah(bumbuCost).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#FFFF00] no-print"></td>
                  )}
                </>
              )}
            </tr>

            {/* ROW 3: KEBUTUHAN 90-110% + (JUMLAH+BUMBU % OR BLANK) */}
            <tr className="border border-black">
              {/* Left Part: Kebutuhan 90-110% (cols 1-10) */}
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                KEBUTUHAN 90-110%
              </td>
              
              {[
                { val: pctEnergi, active: targetEnergi > 0 && !isNaN(pctEnergi) && isFinite(pctEnergi) },
                { val: pctProtein, active: targetProtein > 0 && !isNaN(pctProtein) && isFinite(pctProtein) },
                { val: pctLemak, active: targetLemak > 0 && !isNaN(pctLemak) && isFinite(pctLemak) },
                { val: pctKH, active: targetKH > 0 && !isNaN(pctKH) && isFinite(pctKH) },
                { val: pctSerat, active: targetSerat > 0 && !isNaN(pctSerat) && isFinite(pctSerat) },
              ].map((pctInfo, idx) => {
                const isOk = pctInfo.active && pctInfo.val >= 90 && pctInfo.val <= 110;
                const cellBg = pctInfo.active ? (isOk ? '#00B050' : '#FF0000') : '#FF0000';
                const cellText = 'black';
                return (
                  <td 
                    key={idx} 
                    className="p-1.5 border border-black text-center font-mono font-black text-black" 
                    style={{ backgroundColor: cellBg, color: cellText }}
                  >
                    {pctInfo.active ? `${pctInfo.val.toFixed(0)}%` : ""}
                  </td>
                );
              })}

              {planningMode === "all_ingredients" ? (
                <>
                  <td colSpan={10} className="p-1.5 border border-black bg-slate-50/10"></td>
                  <td className="p-1 border border-black bg-slate-50/10"></td>
                  {!isPrint && (
                    <td className="p-1 border border-black bg-slate-50/10 no-print"></td>
                  )}
                </>
              ) : (
                <>
                  {/* Right Part: JUMLAH+BUMBU % (cols 11-22) */}
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#F8CBAD]">
                    JUMLAH+BUMBU {activeBumbuPct}%
                  </td>
                  <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#F8CBAD]">
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                      <span className="text-slate-600 font-black">Rp</span>
                      <span className="font-black">{formatRupiah(subtotalCost).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#F8CBAD] no-print"></td>
                  )}
                </>
              )}
            </tr>

            {/* ROW 4: HARGA PER PORSI */}
            {planningMode === "with_bumbu_10" && (
              <tr className="border border-black font-extrabold text-black">
                {/* Left Part: Empty (cols 1-10) */}
                <td colSpan={10} className="p-1.5 border border-black bg-slate-50/30"></td>

                {/* Right Part: HARGA PER PORSI (cols 11-22) */}
                <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#A9D08E]">
                  HARGA PER PORSI
                </td>
                <td className={`p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}>
                  <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950 print:text-black">
                    <span className="text-slate-950 print:text-black font-black">Rp</span>
                    <span className="font-black text-slate-950 print:text-black">{formatRupiah(calculatedCostPerPorsi).replace("Rp", "").trim()}</span>
                  </div>
                </td>
                {!isPrint && (
                  <td className={`p-1 border border-black border-l-2 border-l-slate-950 no-print ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}></td>
                )}
              </tr>
            )}
          </tfoot>
        </table>
      </div>
      {!isPrint &&
        renderOverBudgetNotice(
          porsi,
          porsi === "besar" ? "Porsi Besar" : "Porsi Kecil",
          calculatedCostPerPorsi,
          targetPrice,
          recipientCount,
          items
        )}
      </>
    );
  };

  const renderCategoryFoodCostTable = (entry: FoodCostCategoryEntry, isPrint: boolean) => {
    const calcInfo = categoryCalcResultsMap[entry.id];
    const items = calcInfo ? calcInfo.items : [];
    const totalBahanCost = calcInfo ? calcInfo.totalBahanCost : 0;
    const bumbuCost = calcInfo ? calcInfo.bumbuCost : 0;
    const subtotalCost = calcInfo ? calcInfo.subtotalCost : 0;
    const calculatedCostPerPorsi = calcInfo ? calcInfo.calculatedCostPerPorsi : 0;
    const recipientCount = calcInfo ? calcInfo.pmCount : getEntryEffectivePmCount(entry);
    const targetPrice = calcInfo ? calcInfo.targetPrice : (entry.porsi === "besar" ? activeTargetPriceBesar : activeTargetPriceKecil);
    const entryBuffer = entry.bufferPct !== undefined ? entry.bufferPct : (currentDayData.bufferPct ?? 5);

    const sumEnergi = items.reduce((acc, b) => acc + b.energi, 0);
    const sumProtein = items.reduce((acc, b) => acc + b.protein, 0);
    const sumLemak = items.reduce((acc, b) => acc + b.lemak, 0);
    const sumKH = items.reduce((acc, b) => acc + b.kh, 0);
    const sumSerat = items.reduce((acc, b) => acc + b.serat, 0);

    const targetEnergi = getCategoryEntryAkgTarget(entry, "energi");
    const targetProtein = getCategoryEntryAkgTarget(entry, "protein");
    const targetLemak = getCategoryEntryAkgTarget(entry, "lemak");
    const targetKH = getCategoryEntryAkgTarget(entry, "kh");
    const targetSerat = getCategoryEntryAkgTarget(entry, "serat");

    const pctEnergi = targetEnergi > 0 ? (sumEnergi / targetEnergi) * 100 : 0;
    const pctProtein = targetProtein > 0 ? (sumProtein / targetProtein) * 100 : 0;
    const pctLemak = targetLemak > 0 ? (sumLemak / targetLemak) * 100 : 0;
    const pctKH = targetKH > 0 ? (sumKH / targetKH) * 100 : 0;
    const pctSerat = targetSerat > 0 ? (sumSerat / targetSerat) * 100 : 0;

    const tableKey = `cat_${entry.id}`;
    const bbHighlight = !isPrint && Boolean(pastedHighlightMap[tableKey]);
    const kriteriaPmText = entry.kriteriaPm !== undefined ? entry.kriteriaPm : entry.title;

    if (items.length === 0) {
      if (isPrint) {
        return (
          <div className="p-4 border border-black rounded-lg text-center text-xs font-bold text-slate-500 italic uppercase bg-slate-50">
            Tidak ada komponen bahan baku untuk kategori {entry.title} hari ini.
          </div>
        );
      }
      return (
        <div className="flex flex-col items-center justify-center p-7 bg-slate-50 border border-dashed border-slate-300 rounded-2xl space-y-3 text-center my-2">
          <div className="p-2.5 bg-indigo-50 text-indigo-500 rounded-full">
            <Layers className="w-5 h-5" />
          </div>
          <div>
            <h5 className="text-xs font-black text-slate-700 uppercase tracking-wide">
              Komponen Bahan Baku "{entry.title}" Masih Kosong
            </h5>
            <p className="text-[11px] text-slate-500 mt-1 max-w-md">
              Klik <strong>+ Tambah Bahan Baku</strong> untuk input manual, atau klik <strong>Tempel Isian</strong> / pilih di dropdown <strong>Salin dari...</strong> jika bahan &amp; uraiannya sama dengan Food Cost lain (Anda tinggal ubah BB).
            </p>
          </div>
        </div>
      );
    }

    return (
      <>
      <div className={isPrint ? "w-full border border-black rounded-lg shadow-xs overflow-hidden" : "overflow-x-auto border border-black rounded-lg shadow-xs"}>
        <table className={`w-full border-collapse border border-black font-sans text-center ${isPrint ? "table-fixed text-[9px] leading-tight" : "text-xs"}`} style={isPrint ? { tableLayout: 'fixed', width: '100%' } : { fontSize: '12px' }}>
          <thead className="bg-[#92D050] text-black font-bold uppercase border border-black text-center">
            <tr>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[110px] min-w-[110px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>KRITERIA PM</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[110px] min-w-[110px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>MENU</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[160px] min-w-[160px]"}`} style={isPrint ? { width: '10%' } : undefined}>BAHAN MAKANAN</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[58px] min-w-[58px]"} ${bbHighlight ? "bg-amber-300 ring-2 ring-inset ring-amber-600" : ""}`} style={isPrint ? { width: '3.5%' } : undefined}>
                BERAT (BB){bbHighlight && <span className="block text-[8px] font-black text-amber-950 normal-case">Ubah ↵</span>}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[58px] min-w-[58px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>URT</th>
              <th colSpan={5} className="p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap bg-[#76933C] print:text-black">KOMPOSISI ZAT GIZI MAKANAN</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[58px] min-w-[58px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>BDD (%)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>BK (g)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black leading-tight ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>Jumlah<br/>Manfaat</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[76px] min-w-[76px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>Gram (g)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap ${isPrint ? "" : "w-[76px] min-w-[76px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>Kg (kg)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap bg-[#92D050] ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>
                {isPrint ? headerPotong : (
                  <input
                    type="text"
                    value={headerPotong}
                    onChange={(e) => {
                      setHeaderPotong(e.target.value);
                      localStorage.setItem("fc_header_potong", e.target.value);
                    }}
                    className="w-full text-center bg-[#92D050] border-0 hover:bg-[#86c145] focus:bg-white font-black p-0.5 rounded text-slate-950 uppercase text-[11px] focus:ring-1 focus:ring-slate-600"
                    title="Ubah judul kolom Potong"
                  />
                )}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black whitespace-nowrap bg-[#92D050] ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>
                {isPrint ? headerEkor : (
                  <input
                    type="text"
                    value={headerEkor}
                    onChange={(e) => {
                      setHeaderEkor(e.target.value);
                      localStorage.setItem("fc_header_ekor", e.target.value);
                    }}
                    className="w-full text-center bg-[#92D050] border-0 hover:bg-[#86c145] focus:bg-white font-black p-0.5 rounded text-slate-950 uppercase text-[11px] focus:ring-1 focus:ring-slate-600"
                    title="Ubah judul kolom Ekor"
                  />
                )}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black bg-[#92D050] text-[11px] leading-tight select-none ${isPrint ? "" : "w-[80px] min-w-[80px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>
                Buffer<br/><span className="text-[10px] font-bold">({entryBuffer}%)</span>
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-black bg-[#92D050] text-[11px] leading-tight select-none ${isPrint ? "" : "w-[90px] min-w-[90px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>
                Jumlah +<br/>Buffer
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[130px] min-w-[130px]"}`} style={isPrint ? { width: '7.5%' } : undefined}>Harga Satuan</th>
              <th rowSpan={2} className={`p-1 border border-black border-r-2 border-r-slate-950 text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[136px] min-w-[136px]"}`} style={isPrint ? { width: '8.5%' } : undefined}>Harga Total</th>
              {!isPrint && <th rowSpan={2} className="p-1 border border-black border-l-2 border-l-slate-950 text-center text-rose-800 w-[44px] min-w-[44px] no-print">Aksi</th>}
            </tr>
            <tr>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[54px] min-w-[54px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>ENERGI (Kal)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>Protein (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>LEMAK (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>KH (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>SERAT (g)</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-black text-[11px] text-slate-900 bg-white">
            {items.map((b, idx) => (
              <tr key={b.id} className="hover:bg-slate-50/50">
                {idx === 0 && (
                  <td rowSpan={items.length} className="p-1 border border-black text-center font-extrabold bg-[#E2EFDA] align-middle text-[10px] uppercase leading-tight w-[110px] min-w-[110px]">
                    {isPrint ? (
                      <span className="block p-1 text-[10px] font-extrabold text-slate-950 uppercase">{kriteriaPmText}</span>
                    ) : (
                      <textarea
                        value={kriteriaPmText}
                        onChange={(e) => updateCategoryEntry(entry.id, (cur) => ({ ...cur, kriteriaPm: e.target.value }))}
                        className="w-full text-center bg-transparent border-0 font-extrabold text-[10px] uppercase focus:ring-1 focus:ring-indigo-500 p-1 resize-none overflow-hidden"
                        rows={3}
                        placeholder="Edit Kriteria PM..."
                      />
                    )}
                  </td>
                )}
                {idx === 0 && (
                  <td rowSpan={items.length} className="p-1 border border-black text-center font-bold bg-[#F2F2F2] align-middle text-[10px] w-[110px] min-w-[110px] break-words text-slate-950">
                    {namaMenuHariIni || "-"}
                  </td>
                )}
                <td className="p-1 border border-black font-semibold w-[160px] min-w-[160px] text-center">
                  {isPrint ? (
                    <span className="p-0.5 block text-xs font-bold text-slate-950 text-center">{b.nama}</span>
                  ) : (
                    <div className="flex items-center gap-1.5 justify-center">
                      <div className="flex-1 min-w-0">
                        <SearchableTkpiDropdown
                          id={`fc-select-tkpi-${tableKey}-${idx}`}
                          tkpiList={tkpiList}
                          selectedValue={entry.bahanList[idx]?.tkpiId || ""}
                          onChange={(val) => editIngredientInCategoryEntry(entry.id, idx, "tkpiId", val)}
                          minimal
                        />
                      </div>
                      <button
                        type="button"
                        onClick={() => removeIngredientFromCategoryEntry(entry.id, b.id, idx)}
                        className="p-1 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg transition hover:scale-105 active:scale-95 shrink-0 no-print"
                        title="Hapus Bahan Makanan Ini"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  )}
                </td>
                <td className={`p-1 border border-black text-center w-[56px] min-w-[56px] ${bbHighlight ? "bg-amber-50 ring-1 ring-inset ring-amber-400" : ""}`}>
                  {isPrint ? (
                    <span className="font-mono font-bold text-slate-950 text-xs text-center">{b.beratBB}</span>
                  ) : (
                    <input
                      id={`fc-input-bb-${tableKey}-${idx}`}
                      type="number"
                      value={entry.bahanList[idx]?.beratBB || ""}
                      onChange={(e) => editIngredientInCategoryEntry(entry.id, idx, "beratBB", Number(e.target.value))}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          e.preventDefault();
                          const nextEl = document.getElementById(`fc-input-bb-${tableKey}-${idx + 1}`) as HTMLInputElement | null;
                          if (nextEl) {
                            nextEl.focus();
                            nextEl.select();
                          }
                        }
                      }}
                      className={`w-full text-center border-0 font-mono font-bold p-0 text-slate-950 text-xs rounded ${bbHighlight ? "bg-amber-100/80 focus:bg-white focus:ring-2 focus:ring-amber-500" : "bg-transparent"}`}
                    />
                  )}
                </td>
                <td className="p-1 border border-black text-center w-[56px] min-w-[56px]">
                  {isPrint ? (
                    <span className="text-slate-950 text-xs text-center">{b.urt}</span>
                  ) : (
                    <input
                      id={`fc-input-urt-${tableKey}-${idx}`}
                      type="text"
                      value={entry.bahanList[idx]?.urt || ""}
                      onChange={(e) => editIngredientInCategoryEntry(entry.id, idx, "urt", e.target.value)}
                      className="w-full text-center bg-transparent border-0 p-0 text-slate-950 text-xs"
                    />
                  )}
                </td>
                <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[54px] min-w-[54px] text-xs">{b.energi.toFixed(1)}</td>
                <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.protein.toFixed(1)}</td>
                <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.lemak.toFixed(1)}</td>
                <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.kh.toFixed(1)}</td>
                <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.serat.toFixed(1)}</td>

                <td className="p-1 border border-black text-center w-[56px] min-w-[56px]">
                  {isPrint ? (
                    <span className="font-mono font-bold text-slate-950 text-xs">{b.bdd}%</span>
                  ) : (
                    <EditableBddCell
                      value={b.bdd}
                      isCustom={entry.bahanList[idx]?.bdd !== undefined}
                      onChange={(newVal) => editIngredientInCategoryEntry(entry.id, idx, "bdd", newVal)}
                    />
                  )}
                </td>
                <td className="p-1 border border-black text-center font-mono font-bold bg-[#E6F0FA]/20 w-[64px] min-w-[64px] text-slate-950 text-xs">{b.beratKotor.toFixed(1)}</td>

                {idx === 0 && (
                  <td rowSpan={items.length} className="p-1 border border-black text-center font-extrabold bg-[#FFF2CC] align-middle font-mono w-[64px] min-w-[64px] text-slate-950 text-xs">
                    {isPrint ? (
                      <span className="font-mono font-extrabold text-slate-950 text-xs">{recipientCount}</span>
                    ) : (
                      <input
                        type="number"
                        min="0"
                        value={entry.customPmCount !== undefined ? entry.customPmCount : recipientCount}
                        onChange={(e) => {
                          const raw = e.target.value;
                          const val = raw === "" ? undefined : Math.max(0, Number(raw));
                          updateCategoryEntry(entry.id, (cur) => ({ ...cur, customPmCount: val }));
                        }}
                        className="w-full text-center bg-transparent border-0 font-mono font-extrabold p-0.5 text-slate-950 text-xs focus:bg-white focus:ring-1 focus:ring-amber-500 rounded cursor-pointer hover:bg-amber-100/60"
                        title="Klik untuk mengedit Jumlah Penerima Manfaat kategori ini"
                      />
                    )}
                  </td>
                )}

                <td className="p-1 border border-black text-center font-mono text-slate-950 w-[72px] min-w-[72px] font-semibold text-xs">{b.totalKebutuhanGram.toFixed(1)}</td>
                <td className="p-1 border border-black text-center font-mono font-bold text-slate-950 w-[72px] min-w-[72px] text-xs">{b.totalKebutuhanKg.toFixed(1)}</td>

                <td className="p-1 border border-black text-center w-[68px] min-w-[68px]">
                  {isPrint ? (
                    <span className="font-mono text-slate-950">{b.potong || "-"}</span>
                  ) : (
                    <div className="relative group/calc flex items-center justify-center h-full w-full">
                      <input
                        type="text"
                        value={entry.bahanList[idx]?.potong ?? ""}
                        onChange={(e) => editIngredientInCategoryEntry(entry.id, idx, "potong", e.target.value)}
                        className="w-full text-center bg-white border border-slate-900 rounded-lg font-mono font-bold text-xs p-1 focus:ring-2 focus:ring-indigo-600 focus:outline-none shadow-2xs"
                      />
                      <button
                        type="button"
                        onClick={() => openCalculator(`cat:${entry.id}`, idx, "potong", String(entry.bahanList[idx]?.potong ?? ""), b.nama)}
                        className="absolute right-0.5 top-1/2 -translate-y-1/2 opacity-0 group-hover/calc:opacity-100 focus:opacity-100 bg-white hover:bg-slate-100 border border-slate-300 shadow-xs px-1 py-0.5 rounded text-[10px] text-slate-600 z-10 transition-opacity"
                        title={`Hitung Potong - ${b.nama}`}
                      >
                        🧮
                      </button>
                    </div>
                  )}
                </td>
                <td className="p-1 border border-black text-center w-[68px] min-w-[68px]">
                  {isPrint ? (
                    <span className="font-mono text-slate-950">{b.ekor || "-"}</span>
                  ) : (
                    <div className="relative group/calc flex items-center justify-center h-full w-full">
                      <input
                        type="text"
                        value={entry.bahanList[idx]?.ekor ?? ""}
                        onChange={(e) => editIngredientInCategoryEntry(entry.id, idx, "ekor", e.target.value)}
                        className="w-full text-center bg-white border border-slate-900 rounded-lg font-mono font-bold text-xs p-1 focus:ring-2 focus:ring-indigo-600 focus:outline-none shadow-2xs"
                      />
                      <button
                        type="button"
                        onClick={() => openCalculator(`cat:${entry.id}`, idx, "ekor", String(entry.bahanList[idx]?.ekor ?? ""), b.nama)}
                        className="absolute right-0.5 top-1/2 -translate-y-1/2 opacity-0 group-hover/calc:opacity-100 focus:opacity-100 bg-white hover:bg-slate-100 border border-slate-300 shadow-xs px-1 py-0.5 rounded text-[10px] text-slate-600 z-10 transition-opacity"
                        title={`Hitung Ekor - ${b.nama}`}
                      >
                        🧮
                      </button>
                    </div>
                  )}
                </td>
                <td
                  onClick={() => !isPrint && openBufferConfig(`cat:${entry.id}`, idx, b)}
                  className={`p-1 border border-black text-center w-[80px] min-w-[80px] bg-slate-50/50 ${!isPrint ? "cursor-pointer hover:bg-indigo-50/70 transition-all group/buffer" : ""}`}
                  title={!isPrint ? "Klik untuk memilih basis Buffer (Kg, Potong, Ekor)" : undefined}
                >
                  <div className="flex flex-col items-center justify-center h-full min-h-[36px]">
                    <span className="font-mono text-slate-950 text-xs font-semibold">
                      {typeof b.buah === "number" && b.buah >= 0 ? b.buah.toFixed(1) : "-"}
                    </span>
                    {!isPrint && (
                      <span className="text-[8px] text-indigo-600 font-extrabold uppercase tracking-wider block leading-none mt-0.5 opacity-70 group-hover/buffer:opacity-100 transition-opacity">
                        {b.bufferBase === "kg" ? "KG" : b.bufferBase === "potong" ? "Ptg" : b.bufferBase === "ekor" ? "Ekr" : b.bufferBase === "custom" ? "Cust" : "Auto"} ⚙️
                      </span>
                    )}
                  </div>
                </td>
                <td
                  onClick={() => !isPrint && openBufferConfig(`cat:${entry.id}`, idx, b)}
                  className={`p-1 border border-black text-center w-[90px] min-w-[90px] bg-[#E6F0FA]/30 ${!isPrint ? "cursor-pointer hover:bg-indigo-50/70 transition-all group/jumlah" : ""}`}
                  title={!isPrint ? "Klik untuk mengatur rumus Jumlah + Buffer" : undefined}
                >
                  <div className="flex flex-col items-center justify-center h-full min-h-[36px]">
                    <span className="font-mono text-slate-950 text-xs font-bold text-indigo-950">
                      {typeof b.butir === "number" && b.butir >= 0 ? b.butir.toFixed(1) : "-"}
                    </span>
                    {!isPrint && (
                      <span className="text-[8px] text-indigo-700 font-extrabold uppercase tracking-wider block leading-none mt-0.5 opacity-70 group-hover/jumlah:opacity-100 transition-opacity">
                        {(() => {
                          const val = b.jumlahBufferChoice || "auto";
                          if (val === "auto") return "Auto";
                          if (val === "kg_with") return "KG + Buf";
                          if (val === "kg_without") return "KG Saja";
                          if (val === "potong_with") return "Ptg + Buf";
                          if (val === "potong_without") return "Ptg Saja";
                          if (val === "ekor_with") return "Ekr + Buf";
                          if (val === "ekor_without") return "Ekr Saja";
                          if (val === "custom_with") return "Cust + Buf";
                          if (val === "custom_without") return "Cust Saja";
                          return val;
                        })()} ⚙️
                      </span>
                    )}
                  </div>
                </td>
                <td className="p-1 border border-black text-center w-[130px] min-w-[130px]">
                  {isPrint ? (
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-bold text-slate-950">
                      <span className="text-slate-500">Rp</span>
                      <span>{formatRupiah(b.hargaSatuan).replace("Rp", "").trim()}</span>
                    </div>
                  ) : (
                    <div className="flex items-center justify-center gap-1 w-full">
                      <div className="flex items-center gap-1 bg-white/50 px-1 rounded border border-slate-200 w-full max-w-[84px]">
                        <span className="text-[10px] font-bold text-slate-400">Rp</span>
                        <input
                          id={`fc-input-price-${tableKey}-${idx}`}
                          type="text"
                          inputMode="numeric"
                          value={formatThousandSeparator(entry.bahanList[idx]?.hargaSatuan)}
                          onChange={(e) => editIngredientInCategoryEntry(entry.id, idx, "hargaSatuan", parseThousandSeparator(e.target.value))}
                          className="w-full text-right bg-transparent border-0 font-mono font-bold p-0 focus:ring-0 text-slate-950 text-xs"
                        />
                      </div>
                      <PriceCalculatorPopover
                        initialValue={entry.bahanList[idx]?.hargaSatuan || 0}
                        onApply={(val) => editIngredientInCategoryEntry(entry.id, idx, "hargaSatuan", val)}
                        placeholder="Hitung Harga Satuan"
                        ingredientName={b.nama}
                      />
                    </div>
                  )}
                </td>
                <td className="p-1 border border-black border-r-2 border-r-slate-950 text-center w-[136px] min-w-[136px]">
                  <div className="flex items-center justify-between px-1 w-full h-full">
                    <span className="text-slate-500 text-[10px] font-bold">Rp</span>
                    <div className="flex items-center gap-1 font-mono text-xs font-bold text-slate-950 whitespace-nowrap">
                      <span>{formatRupiah(b.hargaTotal).replace("Rp", "").trim()}</span>
                    </div>
                  </div>
                </td>
                {!isPrint && (
                  <td className="p-1 border border-black border-l-2 border-l-slate-950 text-center w-[44px] min-w-[44px] no-print">
                    <button
                      type="button"
                      onClick={() => removeIngredientFromCategoryEntry(entry.id, b.id, idx)}
                      className="p-0.5 text-slate-400 hover:text-rose-600 rounded"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </td>
                )}
              </tr>
            ))}
          </tbody>

          <tfoot className="border border-black font-semibold text-slate-950 text-xs text-center">
            <tr className="border border-black">
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                KANDUNGAN GIZI
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[54px] min-w-[54px] bg-[#92D050]">{sumEnergi.toFixed(1)}</td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">{sumProtein.toFixed(1)}</td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">{sumLemak.toFixed(1)}</td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">{sumKH.toFixed(1)}</td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">{sumSerat.toFixed(1)}</td>
              <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#BFBFBF]">JUMLAH</td>
              <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#BFBFBF]">
                <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                  <span className="text-slate-600 font-black">Rp</span>
                  <span className="font-black">{formatRupiah(totalBahanCost).replace("Rp", "").trim()}</span>
                </div>
              </td>
              {!isPrint && <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#BFBFBF] no-print"></td>}
            </tr>

            <tr className="border border-black">
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                <div className="flex items-center justify-center gap-1.5 flex-wrap">
                  <span>AKG: {entry.title.toUpperCase()}</span>
                  {!isPrint && (
                    <select
                      value={entry.akgMode || "min"}
                      onChange={(e) => {
                        const newMode = e.target.value;
                        updateCategoryEntry(entry.id, (cur) => ({
                          ...cur,
                          akgMode: newMode,
                          nutrientModes: { energi: newMode, protein: newMode, lemak: newMode, kh: newMode, serat: newMode },
                        }));
                      }}
                      className="bg-white hover:bg-slate-50 text-slate-900 text-[11px] font-black rounded-lg px-2 py-0.5 border border-indigo-400 shadow-xs focus:outline-none uppercase cursor-pointer"
                    >
                      <option value="min">SEMUA MIN</option>
                      <option value="max">SEMUA MAX</option>
                      <option value="avg">SEMUA RATA-RATA</option>
                      <option value="custom">SEMUA KUSTOM ✍️</option>
                    </select>
                  )}
                </div>
              </td>

              {[
                { key: "energi", target: targetEnergi },
                { key: "protein", target: targetProtein },
                { key: "lemak", target: targetLemak },
                { key: "kh", target: targetKH },
                { key: "serat", target: targetSerat },
              ].map((nut) => (
                <td key={nut.key} className="p-1 border border-black text-center font-mono font-bold text-black w-[58px] min-w-[58px] bg-[#92D050]">
                  {!isPrint && (
                    <select
                      value={entry.nutrientModes?.[nut.key] || entry.akgMode || "min"}
                      onChange={(e) => updateCategoryEntryNutrientMode(entry.id, nut.key, e.target.value)}
                      className="w-full text-[8.5px] font-black uppercase bg-white border border-emerald-700/50 rounded px-0.5 py-0.5 mb-1 text-slate-900 focus:outline-none cursor-pointer"
                    >
                      <option value="min">MIN</option>
                      <option value="max">MAX</option>
                      <option value="avg">RATA2</option>
                      <option value="custom">KUST</option>
                    </select>
                  )}
                  {isPrint ? (
                    <span>{nut.target > 0 ? nut.target.toFixed(1) : "-"}</span>
                  ) : (
                    <input
                      type="number"
                      step="any"
                      value={nut.target > 0 ? Number(nut.target.toFixed(1)) : ""}
                      onChange={(e) => handleCategoryEntryAkgValChange(entry, nut.key, Number(e.target.value))}
                      className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-0.5 rounded-md text-slate-900 text-[11px] shadow-2xs"
                    />
                  )}
                </td>
              ))}

              {planningMode === "all_ingredients" ? (
                <>
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#A9D08E]">
                    HARGA PER PORSI
                  </td>
                  <td className={`p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}>
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950 print:text-black">
                      <span className="font-black">Rp</span>
                      <span className="font-black">{formatRupiah(calculatedCostPerPorsi).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className={`p-1 border border-black border-l-2 border-l-slate-950 no-print ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}></td>
                  )}
                </>
              ) : (
                <>
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#FFFF00]">
                    BUMBU {activeBumbuPct}%
                  </td>
                  <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#FFFF00]">
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                      <span className="text-slate-600 font-black">Rp</span>
                      <span className="font-black">{formatRupiah(bumbuCost).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#FFFF00] no-print"></td>}
                </>
              )}
            </tr>

            <tr className="border border-black">
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                KEBUTUHAN 90-110%
              </td>
              {[
                { val: pctEnergi, active: targetEnergi > 0 && !isNaN(pctEnergi) && isFinite(pctEnergi) },
                { val: pctProtein, active: targetProtein > 0 && !isNaN(pctProtein) && isFinite(pctProtein) },
                { val: pctLemak, active: targetLemak > 0 && !isNaN(pctLemak) && isFinite(pctLemak) },
                { val: pctKH, active: targetKH > 0 && !isNaN(pctKH) && isFinite(pctKH) },
                { val: pctSerat, active: targetSerat > 0 && !isNaN(pctSerat) && isFinite(pctSerat) },
              ].map((pctInfo, idx) => {
                const isOk = pctInfo.active && pctInfo.val >= 90 && pctInfo.val <= 110;
                const cellBg = pctInfo.active ? (isOk ? '#00B050' : '#FF0000') : '#FF0000';
                return (
                  <td key={idx} className="p-1.5 border border-black text-center font-mono font-black text-black" style={{ backgroundColor: cellBg, color: 'black' }}>
                    {pctInfo.active ? `${pctInfo.val.toFixed(0)}%` : ""}
                  </td>
                );
              })}

              {planningMode === "all_ingredients" ? (
                <>
                  <td colSpan={10} className="p-1.5 border border-black bg-slate-50/10"></td>
                  <td className="p-1 border border-black bg-slate-50/10"></td>
                  {!isPrint && <td className="p-1 border border-black bg-slate-50/10 no-print"></td>}
                </>
              ) : (
                <>
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#F8CBAD]">
                    JUMLAH+BUMBU {activeBumbuPct}%
                  </td>
                  <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#F8CBAD]">
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                      <span className="text-slate-600 font-black">Rp</span>
                      <span className="font-black">{formatRupiah(subtotalCost).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#F8CBAD] no-print"></td>}
                </>
              )}
            </tr>

            {planningMode === "with_bumbu_10" && (
              <tr className="border border-black font-extrabold text-black">
                <td colSpan={10} className="p-1.5 border border-black bg-slate-50/30"></td>
                <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#A9D08E]">
                  HARGA PER PORSI
                </td>
                <td className={`p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}>
                  <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950 print:text-black">
                    <span className="font-black">Rp</span>
                    <span className="font-black">{formatRupiah(calculatedCostPerPorsi).replace("Rp", "").trim()}</span>
                  </div>
                </td>
                {!isPrint && (
                  <td className={`p-1 border border-black border-l-2 border-l-slate-950 no-print ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}></td>
                )}
              </tr>
            )}
          </tfoot>
        </table>
      </div>
      {!isPrint &&
        renderOverBudgetNotice(
          `cat_${entry.id}`,
          entry.title,
          calculatedCostPerPorsi,
          targetPrice,
          recipientCount,
          items
        )}
      </>
    );
  };

  const renderCustomFoodCostTable = (table: any, isPrint: boolean) => {
    const results = table.porsi === "besar"
      ? calculateDay(table.bahanList, [], table.pmCount, 0, table.bufferPct || 5, tkpiList, activeBumbuPct)
      : calculateDay([], table.bahanList, 0, table.pmCount, table.bufferPct || 5, tkpiList, activeBumbuPct);

    const items = table.porsi === "besar" ? results.porsiBesarItems : results.porsiKecilItems;
    const totalBahanCost = table.porsi === "besar" ? results.totalBesarBahanCost : results.totalKecilBahanCost;
    const bufferCost = table.porsi === "besar" ? results.bufferBesarCost : results.bufferKecilCost;
    const bumbuCost = table.porsi === "besar" ? results.bumbuBesarCost : results.bumbuKecilCost;
    const subtotalCost = table.porsi === "besar" ? results.subtotalBesarCost : results.subtotalKecilCost;
    const costPerPorsi = table.porsi === "besar" ? results.costPerPorsiBesar : results.costPerPorsiKecil;
    const recipientCount = table.pmCount;
    const targetPrice = table.plafon;

    const calculatedCostPerPorsi = planningMode === "all_ingredients"
      ? (recipientCount > 0 ? totalBahanCost / recipientCount : 0)
      : costPerPorsi;

    const akgType = table.akgType || (table.porsi === "besar" ? "sd_besar" : "sd_kecil");
    const akgMode = table.akgMode || "avg";

    const getCustomTableAkgTarget = (key: string): number => {
      if (akgType === "custom" || akgMode === "custom") {
        const customMap = table.porsi === "besar" ? customBesar : customKecil;
        return customMap[key] || 0;
      }
      const config = TARGET_AKG_LIMITS[akgType];
      if (!config) return 0;

      const minVal = key === "energi" ? config.energiMin : (config as any)[`${key}Min`] || 0;
      const ratio = config.energiMax / config.energiMin;

      if (akgMode === "min") {
        return minVal;
      }
      if (akgMode === "max") {
        if (key === "energi") return config.energiMax;
        return Math.round(minVal * ratio * 10) / 10;
      }
      // "avg" mode
      if (key === "energi") {
        return (config.energiMin + config.energiMax) / 2;
      }
      const maxVal = minVal * ratio;
      return Math.round(((minVal + maxVal) / 2) * 10) / 10;
    };

    const sumEnergi = items.reduce((acc, b) => acc + b.energi, 0);
    const sumProtein = items.reduce((acc, b) => acc + b.protein, 0);
    const sumLemak = items.reduce((acc, b) => acc + b.lemak, 0);
    const sumKH = items.reduce((acc, b) => acc + b.kh, 0);
    const sumSerat = items.reduce((acc, b) => acc + b.serat, 0);

    const targetEnergi = getCustomTableAkgTarget("energi");
    const targetProtein = getCustomTableAkgTarget("protein");
    const targetLemak = getCustomTableAkgTarget("lemak");
    const targetKH = getCustomTableAkgTarget("kh");
    const targetSerat = getCustomTableAkgTarget("serat");

    const pctEnergi = targetEnergi > 0 ? (sumEnergi / targetEnergi) * 100 : 0;
    const pctProtein = targetProtein > 0 ? (sumProtein / targetProtein) * 100 : 0;
    const pctLemak = targetLemak > 0 ? (sumLemak / targetLemak) * 100 : 0;
    const pctKH = targetKH > 0 ? (sumKH / targetKH) * 100 : 0;
    const pctSerat = targetSerat > 0 ? (sumSerat / targetSerat) * 100 : 0;

    const tableKriteria = table.kriteriaPm || (table.porsi === "besar" ? "Kriteria PM Besar" : "Kriteria PM Kecil");

    if (items.length === 0) {
      if (isPrint) {
        return (
          <div className="p-4 border border-black rounded-lg text-center text-xs font-bold text-slate-500 italic uppercase bg-slate-50">
            Tidak ada komponen bahan baku untuk porsi {table.porsi} di tabel kustom ini.
          </div>
        );
      }
      return (
        <div className="flex flex-col items-center justify-center p-8 bg-slate-50 border border-dashed border-slate-300 rounded-2xl space-y-3 text-center my-2">
          <div className="p-3 bg-rose-50 text-rose-500 rounded-full animate-pulse">
            <Trash2 className="w-6 h-6" />
          </div>
          <div>
            <h5 className="text-xs font-black text-slate-700 uppercase tracking-wide">
              Komponen Bahan Baku Kosong
            </h5>
            <p className="text-[11px] text-slate-500 mt-1 max-w-sm">
              Seluruh bahan baku telah dihapus. Silakan klik tombol <strong className="text-indigo-600 font-extrabold">+ Tambah Bahan Baku</strong> di atas untuk menginput ulang.
            </p>
          </div>
        </div>
      );
    }

    const customBufferPct = table.bufferPct || 5;
    const customTableKey = `custom_${table.id}`;
    const customBbHighlight = !isPrint && Boolean(pastedHighlightMap[customTableKey]);

    return (
      <>
      <div className={isPrint ? "w-full border border-black rounded-lg shadow-xs overflow-hidden" : "overflow-x-auto border border-black rounded-lg shadow-xs"}>
        <table className={`w-full border-collapse border border-black font-sans text-center ${isPrint ? "table-fixed text-[9px] leading-tight" : "text-xs"}`} style={isPrint ? { tableLayout: 'fixed', width: '100%' } : { fontSize: '12px' }}>
          <thead className="bg-[#92D050] text-black font-bold uppercase border border-black text-center">
            <tr>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[110px] min-w-[110px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>KRITERIA PM</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[110px] min-w-[110px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>MENU</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[160px] min-w-[160px]"}`} style={isPrint ? { width: '10%' } : undefined}>BAHAN MAKANAN</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[56px] min-w-[56px]"} ${customBbHighlight ? "bg-amber-300 ring-2 ring-inset ring-amber-600" : ""}`} style={isPrint ? { width: '3.5%' } : undefined}>
                BERAT (BB){customBbHighlight && <span className="block text-[8px] font-black text-amber-950 normal-case">Ubah ↵</span>}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>URT</th>
              <th colSpan={5} className="p-1 border border-black text-center text-slate-950 font-extrabold bg-[#76933C] print:text-black">KOMPOSISI ZAT GIZI MAKANAN</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>BDD (%)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[64px] min-w-[64px]"}`} style={isPrint ? { width: '4%' } : undefined}>BK (g)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold leading-tight ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>Jumlah<br/>Manfaat</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[72px] min-w-[72px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>Gram (g)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[72px] min-w-[72px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>Kg (kg)</th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold bg-[#92D050] ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>
                {isPrint ? (
                  headerPotong
                ) : (
                  <input
                    type="text"
                    value={headerPotong}
                    onChange={(e) => {
                      setHeaderPotong(e.target.value);
                      localStorage.setItem("fc_header_potong", e.target.value);
                    }}
                    className="w-full text-center bg-[#92D050] border-0 hover:bg-[#86c145] focus:bg-white font-extrabold p-0.5 rounded text-slate-950 uppercase text-[11px] focus:ring-1 focus:ring-slate-600"
                    title="Ubah judul kolom Potong"
                  />
                )}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold bg-[#92D050] ${isPrint ? "" : "w-[68px] min-w-[68px]"}`} style={isPrint ? { width: '4%' } : undefined}>
                {isPrint ? (
                  headerEkor
                ) : (
                  <input
                    type="text"
                    value={headerEkor}
                    onChange={(e) => {
                      setHeaderEkor(e.target.value);
                      localStorage.setItem("fc_header_ekor", e.target.value);
                    }}
                    className="w-full text-center bg-[#92D050] border-0 hover:bg-[#86c145] focus:bg-white font-extrabold p-0.5 rounded text-slate-950 uppercase text-[11px] focus:ring-1 focus:ring-slate-600"
                    title="Ubah judul kolom Ekor"
                  />
                )}
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold bg-[#92D050] text-[10px] leading-tight select-none ${isPrint ? "" : "w-[80px] min-w-[80px]"}`} style={isPrint ? { width: '4.5%' } : undefined}>
                Buffer<br/><span className="text-[10px] font-bold">({customBufferPct}%)</span>
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold bg-[#92D050] text-[10px] leading-tight select-none ${isPrint ? "" : "w-[90px] min-w-[90px]"}`} style={isPrint ? { width: '5.5%' } : undefined}>
                Jumlah +<br/>Buffer
              </th>
              <th rowSpan={2} className={`p-1 border border-black text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[130px] min-w-[130px]"}`} style={isPrint ? { width: '7.5%' } : undefined}>Harga Satuan</th>
              <th rowSpan={2} className={`p-1 border border-black border-r-2 border-r-slate-950 text-center text-slate-950 font-extrabold ${isPrint ? "" : "w-[136px] min-w-[136px]"}`} style={isPrint ? { width: '8.5%' } : undefined}>Harga Total</th>
              {!isPrint && <th rowSpan={2} className="p-1 border border-black border-l-2 border-l-slate-950 text-center text-rose-800 w-[44px] min-w-[44px] no-print">Aksi</th>}
            </tr>
            <tr>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[54px] min-w-[54px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>ENERGI (Kal)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>Protein (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>LEMAK (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>KH (g)</th>
              <th className={`p-1 border border-black text-center bg-[#D8E4BC] text-slate-950 font-bold ${isPrint ? "" : "w-[56px] min-w-[56px]"}`} style={isPrint ? { width: '3.5%' } : undefined}>SERAT (g)</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-black text-[11px] text-slate-900 bg-white">
            {items.length === 0 ? (
              <tr>
                <td colSpan={isPrint ? 21 : 22} className="p-6 text-center text-slate-400 italic bg-slate-50 border border-black">
                  Belum ada bahan makanan yang diinput untuk tabel ini.
                </td>
              </tr>
            ) : (
              items.map((b, idx) => (
                <tr key={b.id} className="hover:bg-slate-50/50">
                  {idx === 0 && (
                    <td rowSpan={items.length} className="p-1 border border-black text-center font-extrabold bg-[#E2EFDA] align-middle text-[10px] uppercase leading-tight w-[110px] min-w-[110px]">
                      {isPrint ? (
                        <span className="block p-1 text-[10px] font-extrabold text-slate-950 uppercase">{tableKriteria}</span>
                      ) : (
                        <textarea
                          value={tableKriteria}
                          onChange={(e) => editCustomTableMeta(table.id, "kriteriaPm", e.target.value)}
                          className="w-full text-center bg-transparent border-0 font-extrabold text-[10px] uppercase focus:ring-1 focus:ring-indigo-500 p-1 resize-none overflow-hidden"
                          rows={3}
                          placeholder="Edit Kriteria PM..."
                        />
                      )}
                    </td>
                  )}
                  {idx === 0 && (
                    <td rowSpan={items.length} className="p-1 border border-black text-center font-bold bg-[#F2F2F2] align-middle text-[10px] w-[110px] min-w-[110px] break-words text-slate-950">
                      {namaMenuHariIni || "-"}
                    </td>
                  )}
                  <td className="p-1 border border-black font-semibold w-[160px] min-w-[160px] text-center">
                    {isPrint ? (
                      <span className="p-0.5 block text-xs font-bold text-slate-950 text-center">{b.nama}</span>
                    ) : (
                      <SearchableTkpiDropdown
                        id={`fc-custom-select-tkpi-${table.id}-${idx}`}
                        tkpiList={tkpiList}
                        selectedValue={table.bahanList[idx]?.tkpiId || ""}
                        onChange={(tkpiId) => editIngredientInCustomTable(table.id, idx, "tkpiId", tkpiId)}
                        minimal
                      />
                    )}
                  </td>
                  <td className={`p-1 border border-black text-center w-[56px] min-w-[56px] ${customBbHighlight ? "bg-amber-50 ring-1 ring-inset ring-amber-400" : ""}`}>
                    {isPrint ? (
                      <span className="font-mono font-bold text-slate-950 text-xs text-center">{b.beratBB}</span>
                    ) : (
                      <input
                        id={`fc-input-bb-${customTableKey}-${idx}`}
                        type="number"
                        value={table.bahanList[idx]?.beratBB || ""}
                        onChange={(e) => editIngredientInCustomTable(table.id, idx, "beratBB", Number(e.target.value))}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") {
                            e.preventDefault();
                            const nextEl = document.getElementById(`fc-input-bb-${customTableKey}-${idx + 1}`) as HTMLInputElement | null;
                            if (nextEl) {
                              nextEl.focus();
                              nextEl.select();
                            }
                          }
                        }}
                        className={`w-full text-center border-0 font-mono font-bold p-0 text-slate-950 text-xs rounded ${customBbHighlight ? "bg-amber-100/80 focus:bg-white focus:ring-2 focus:ring-amber-500" : "bg-transparent"}`}
                      />
                    )}
                  </td>
                  <td className="p-1 border border-black text-center w-[56px] min-w-[56px]">
                    {isPrint ? (
                      <span className="text-slate-950 text-xs text-center">{b.urt}</span>
                    ) : (
                      <input
                        type="text"
                        value={table.bahanList[idx]?.urt || ""}
                        onChange={(e) => editIngredientInCustomTable(table.id, idx, "urt", e.target.value)}
                        className="w-full text-center bg-transparent border-0 p-0 text-slate-950 text-xs text-center"
                      />
                    )}
                  </td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[54px] min-w-[54px] text-xs">{b.energi.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.protein.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.lemak.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.kh.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono text-slate-950 font-semibold bg-slate-50/30 w-[56px] min-w-[56px] text-xs">{b.serat.toFixed(1)}</td>
                  
                  <td className="p-1 border border-black text-center w-[56px] min-w-[56px]">
                    {isPrint ? (
                      <span className="font-mono font-bold text-slate-950 text-xs">{b.bdd}%</span>
                    ) : (
                      <EditableBddCell
                        value={b.bdd}
                        isCustom={table.bahanList[idx]?.bdd !== undefined}
                        onChange={(newVal) => editIngredientInCustomTable(table.id, idx, "bdd", newVal)}
                      />
                    )}
                  </td>
                  <td className="p-1 border border-black text-center font-mono text-indigo-955 font-bold bg-[#E6F0FA]/20 w-[64px] min-w-[64px] text-slate-950 text-xs">{b.beratKotor.toFixed(1)}</td>
                  
                  {idx === 0 && (
                    <td rowSpan={items.length} className="p-1 border border-black text-center font-extrabold bg-[#FFF2CC] align-middle font-mono w-[64px] min-w-[64px] text-slate-950 text-xs">
                      {isPrint ? (
                        <span className="font-mono font-extrabold text-slate-950 text-xs">{recipientCount}</span>
                      ) : (
                        <input
                          type="number"
                          min="0"
                          value={table.pmCount !== undefined ? table.pmCount : ""}
                          onChange={(e) => {
                            const raw = e.target.value;
                            const val = raw === "" ? 0 : Math.max(0, Number(raw));
                            editCustomTableMeta(table.id, "pmCount", val);
                          }}
                          className="w-full text-center bg-transparent border-0 font-mono font-extrabold p-0.5 text-slate-950 text-xs focus:bg-white focus:ring-1 focus:ring-amber-500 rounded cursor-pointer hover:bg-amber-100/60"
                          title="Klik untuk mengedit atau menghapus/mengubah Jumlah Penerima Manfaat"
                        />
                      )}
                    </td>
                  )}
                  
                  <td className="p-1 border border-black text-center font-mono text-slate-950 w-[72px] min-w-[72px] font-semibold text-xs">{b.totalKebutuhanGram.toFixed(1)}</td>
                  <td className="p-1 border border-black text-center font-mono font-bold text-slate-950 w-[72px] min-w-[72px] text-xs">{b.totalKebutuhanKg.toFixed(1)}</td>
                  
                  <td className="p-1 border border-black text-center w-[68px] min-w-[68px]">
                    {isPrint ? (
                      <span className="font-mono text-slate-950">{b.potong || "-"}</span>
                    ) : (
                      <div className="relative group/calc flex items-center justify-center h-full w-full">
                        <input
                          type="text"
                          value={b.potong ?? ""}
                          onChange={(e) => editIngredientInCustomTable(table.id, idx, "potong", e.target.value)}
                          className="w-full text-center bg-white border border-slate-900 rounded-lg font-mono font-bold text-xs p-1 focus:ring-2 focus:ring-indigo-600 focus:outline-none shadow-2xs"
                        />
                        <button
                          type="button"
                          onClick={() => openCalculator(table.id, idx, "potong", String(b.potong ?? ""), b.nama)}
                          className="absolute right-0.5 top-1/2 -translate-y-1/2 opacity-0 group-hover/calc:opacity-100 focus:opacity-100 bg-white hover:bg-slate-100 border border-slate-300 shadow-xs px-1 py-0.5 rounded text-[10px] text-slate-600 z-10 transition-opacity"
                          title={`Hitung Potong - ${b.nama}`}
                        >
                          🧮
                        </button>
                      </div>
                    )}
                  </td>
                  <td className="p-1 border border-black text-center w-[68px] min-w-[68px]">
                    {isPrint ? (
                      <span className="font-mono text-slate-950">{b.ekor || "-"}</span>
                    ) : (
                      <div className="relative group/calc flex items-center justify-center h-full w-full">
                        <input
                          type="text"
                          value={b.ekor ?? ""}
                          onChange={(e) => editIngredientInCustomTable(table.id, idx, "ekor", e.target.value)}
                          className="w-full text-center bg-white border border-slate-900 rounded-lg font-mono font-bold text-xs p-1 focus:ring-2 focus:ring-indigo-600 focus:outline-none shadow-2xs"
                        />
                        <button
                          type="button"
                          onClick={() => openCalculator(table.id, idx, "ekor", String(b.ekor ?? ""), b.nama)}
                          className="absolute right-0.5 top-1/2 -translate-y-1/2 opacity-0 group-hover/calc:opacity-100 focus:opacity-100 bg-white hover:bg-slate-100 border border-slate-300 shadow-xs px-1 py-0.5 rounded text-[10px] text-slate-600 z-10 transition-opacity"
                          title={`Hitung Ekor - ${b.nama}`}
                        >
                          🧮
                        </button>
                      </div>
                    )}
                  </td>
                  <td
                    onClick={() => !isPrint && openBufferConfig(table.id, idx, b)}
                    className={`p-1 border border-black text-center w-[80px] min-w-[80px] bg-slate-50/50 ${!isPrint ? "cursor-pointer hover:bg-indigo-50/70 transition-all group/buffer" : ""}`}
                    title={!isPrint ? "Klik untuk memilih basis Buffer (Kg, Potong, Ekor)" : undefined}
                  >
                    <div className="flex flex-col items-center justify-center h-full min-h-[36px]">
                      <span className="font-mono text-slate-950 text-xs font-semibold">
                        {typeof b.buah === "number" && b.buah >= 0 ? b.buah.toFixed(1) : "-"}
                      </span>
                      {!isPrint && (
                        <span className="text-[8px] text-indigo-600 font-extrabold uppercase tracking-wider block leading-none mt-0.5 opacity-70 group-hover/buffer:opacity-100 transition-opacity">
                          {b.bufferBase === "kg" ? "KG" : b.bufferBase === "potong" ? "Ptg" : b.bufferBase === "ekor" ? "Ekr" : b.bufferBase === "custom" ? "Cust" : "Auto"} ⚙️
                        </span>
                      )}
                    </div>
                  </td>
                  <td
                    onClick={() => !isPrint && openBufferConfig(table.id, idx, b)}
                    className={`p-1 border border-black text-center w-[90px] min-w-[90px] bg-[#E6F0FA]/30 ${!isPrint ? "cursor-pointer hover:bg-indigo-50/70 transition-all group/jumlah" : ""}`}
                    title={!isPrint ? "Klik untuk mengatur rumus Jumlah + Buffer" : undefined}
                  >
                    <div className="flex flex-col items-center justify-center h-full min-h-[36px]">
                      <span className="font-mono text-slate-950 text-xs font-bold text-indigo-950">
                        {typeof b.butir === "number" && b.butir >= 0 ? b.butir.toFixed(1) : "-"}
                      </span>
                      {!isPrint && (
                        <span className="text-[8px] text-indigo-700 font-extrabold uppercase tracking-wider block leading-none mt-0.5 opacity-70 group-hover/jumlah:opacity-100 transition-opacity">
                          {(() => {
                            const val = b.jumlahBufferChoice || "auto";
                            if (val === "auto") return "Auto";
                            if (val === "kg_with") return "KG + Buf";
                            if (val === "kg_without") return "KG Saja";
                            if (val === "potong_with") return "Ptg + Buf";
                            if (val === "potong_without") return "Ptg Saja";
                            if (val === "ekor_with") return "Ekr + Buf";
                            if (val === "ekor_without") return "Ekr Saja";
                            if (val === "custom_with") return "Cust + Buf";
                            if (val === "custom_without") return "Cust Saja";
                            return val;
                          })()} ⚙️
                        </span>
                      )}
                    </div>
                  </td>
                  
                  <td className="p-1 border border-black text-center w-[130px] min-w-[130px]">
                    {isPrint ? (
                      <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-bold text-slate-950">
                        <span className="text-slate-500">Rp</span>
                        <span>{formatRupiah(b.hargaSatuan).replace("Rp", "").trim()}</span>
                      </div>
                    ) : (
                      <div className="flex items-center justify-center gap-1 w-full">
                        <div className="flex items-center gap-1 bg-white/50 px-1 rounded border border-slate-200 w-full max-w-[84px]">
                          <span className="text-[10px] font-bold text-slate-400">Rp</span>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formatThousandSeparator(b.hargaSatuan)}
                            onChange={(e) => editIngredientInCustomTable(table.id, idx, "hargaSatuan", parseThousandSeparator(e.target.value))}
                            className="w-full text-right bg-transparent border-0 font-mono font-bold p-0 focus:ring-0 text-slate-950 text-xs"
                          />
                        </div>
                        <PriceCalculatorPopover
                          initialValue={b.hargaSatuan || 0}
                          onApply={(val) => editIngredientInCustomTable(table.id, idx, "hargaSatuan", val)}
                          placeholder="Hitung Harga Satuan"
                          ingredientName={b.nama}
                        />
                      </div>
                    )}
                  </td>
                  <td className="p-1 border border-black border-r-2 border-r-slate-950 text-center w-[136px] min-w-[136px]">
                    <div className="flex items-center justify-between px-1 w-full h-full">
                      <span className="text-slate-500 text-[10px] font-bold">Rp</span>
                      <div className="flex items-center gap-1 font-mono text-xs font-bold text-slate-950 whitespace-nowrap">
                        <span>{formatRupiah(b.hargaTotal).replace("Rp", "").trim()}</span>
                      </div>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className="p-1 border border-black border-l-2 border-l-slate-950 text-center w-[44px] min-w-[44px] no-print">
                      <button
                        type="button"
                        onClick={() => removeIngredientFromCustomTable(table.id, b.id, idx)}
                        className="p-0.5 text-slate-400 hover:text-rose-600 rounded"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </td>
                  )}
                </tr>
              ))
            )}
          </tbody>
          
          <tfoot className="border border-black font-semibold text-slate-950 text-xs text-center">
            {/* ROW 1: KANDUNGAN GIZI + JUMLAH */}
            <tr className="border border-black">
              {/* Left Part: Kandungan Gizi (cols 1-10) */}
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                KANDUNGAN GIZI
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[54px] min-w-[54px] bg-[#92D050]">
                {sumEnergi.toFixed(1)}
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {sumProtein.toFixed(1)}
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {sumLemak.toFixed(1)}
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {sumKH.toFixed(1)}
              </td>
              <td className="p-1.5 border border-black text-center font-mono font-extrabold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {sumSerat.toFixed(1)}
              </td>

              {/* Right Part: JUMLAH (cols 11-22) */}
              <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#BFBFBF]">
                JUMLAH
              </td>
              <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#BFBFBF]">
                <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                  <span className="text-slate-600 font-black">Rp</span>
                  <span className="font-black">{formatRupiah(totalBahanCost).replace("Rp", "").trim()}</span>
                </div>
              </td>
              {!isPrint && (
                <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#BFBFBF] no-print"></td>
              )}
            </tr>

            {/* ROW 2: KEBUTUHAN RUJUKAN AKG + (BUMBU 10% OR HARGA PER PORSI) */}
            <tr className="border border-black">
              {/* Left Part: Rujukan AKG (cols 1-10) */}
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                <div className="flex items-center justify-center gap-1.5 flex-wrap">
                  <span>KEBUTUHAN RUJUKAN AKG PORSI {table.porsi === "besar" ? "BESAR" : "KECIL"}</span>
                  {!isPrint && (
                    <select
                      value={akgMode || ""}
                      onChange={(e) => editCustomTableMeta(table.id, "akgMode", e.target.value)}
                      className="bg-white hover:bg-slate-50 text-slate-900 text-[11px] font-black rounded-lg px-2 py-0.5 border border-indigo-400 shadow-xs focus:outline-none focus:ring-2 focus:ring-indigo-500 uppercase cursor-pointer"
                    >
                      <option value="">-- PILIH RUJUKAN --</option>
                      <option value="min">RUJUKAN MIN</option>
                      <option value="max">RUJUKAN MAX</option>
                      <option value="avg">RUJUKAN RATA-RATA</option>
                      <option value="custom">KUSTOM (EDIT BEBAS) ✍️</option>
                    </select>
                  )}
                </div>
              </td>
              
              {/* ENERGI */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[54px] min-w-[54px] bg-[#92D050]">
                {isPrint ? (
                  <span>{(akgMode && targetEnergi > 0) ? targetEnergi.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={(akgMode && targetEnergi > 0) ? Number(targetEnergi.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      handleAkgValChange(table.porsi, "energi", val, akgMode, table.akgType, table.id);
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-1 rounded-md text-slate-900 text-xs shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {/* PROTEIN */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {isPrint ? (
                  <span>{(akgMode && targetProtein > 0) ? targetProtein.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={(akgMode && targetProtein > 0) ? Number(targetProtein.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      handleAkgValChange(table.porsi, "protein", val, akgMode, table.akgType, table.id);
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-1 rounded-md text-slate-900 text-xs shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {/* LEMAK */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {isPrint ? (
                  <span>{(akgMode && targetLemak > 0) ? targetLemak.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={(akgMode && targetLemak > 0) ? Number(targetLemak.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      handleAkgValChange(table.porsi, "lemak", val, akgMode, table.akgType, table.id);
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-1 rounded-md text-slate-900 text-xs shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {/* KH */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {isPrint ? (
                  <span>{(akgMode && targetKH > 0) ? targetKH.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={(akgMode && targetKH > 0) ? Number(targetKH.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      handleAkgValChange(table.porsi, "kh", val, akgMode, table.akgType, table.id);
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-1 rounded-md text-slate-900 text-xs shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {/* SERAT */}
              <td className="p-1 border border-black text-center font-mono font-bold text-black w-[56px] min-w-[56px] bg-[#92D050]">
                {isPrint ? (
                  <span>{(akgMode && targetSerat > 0) ? targetSerat.toFixed(1) : "-"}</span>
                ) : (
                  <input
                    type="number"
                    step="any"
                    placeholder=""
                    value={(akgMode && targetSerat > 0) ? Number(targetSerat.toFixed(1)) : ""}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      handleAkgValChange(table.porsi, "serat", val, akgMode, table.akgType, table.id);
                    }}
                    className="w-full text-center bg-[#C6E0B4] hover:bg-white focus:bg-white border border-emerald-700/40 focus:ring-2 focus:ring-indigo-600 font-mono font-extrabold px-0.5 py-1 rounded-md text-slate-900 text-xs shadow-2xs transition-all [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                )}
              </td>

              {planningMode === "all_ingredients" ? (
                <>
                  {/* Right Part: HARGA PER PORSI (cols 11-22) */}
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#A9D08E]">
                    HARGA PER PORSI
                  </td>
                  <td className={`p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}>
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950 print:text-black">
                      <span className="text-slate-950 print:text-black font-black">Rp</span>
                      <span className="font-black text-slate-950 print:text-black">{formatRupiah(calculatedCostPerPorsi).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className={`p-1 border border-black border-l-2 border-l-slate-950 no-print ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}></td>
                  )}
                </>
              ) : (
                <>
                  {/* Right Part: BUMBU CUSTOM % (cols 11-22) */}
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#FFFF00]">
                    BUMBU {activeBumbuPct}%
                  </td>
                  <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#FFFF00]">
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                      <span className="text-slate-600 font-black">Rp</span>
                      <span className="font-black">{formatRupiah(bumbuCost).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#FFFF00] no-print"></td>
                  )}
                </>
              )}
            </tr>

            {/* ROW 3: KEBUTUHAN 90-110% + (JUMLAH+BUMBU % OR BLANK) */}
            <tr className="border border-black">
              {/* Left Part: Kebutuhan 90-110% (cols 1-10) */}
              <td colSpan={5} className="p-1.5 border border-black font-extrabold text-black uppercase bg-[#92D050] text-center">
                KEBUTUHAN 90-110%
              </td>
              
              {[
                { val: pctEnergi, active: targetEnergi > 0 && !isNaN(pctEnergi) && isFinite(pctEnergi) },
                { val: pctProtein, active: targetProtein > 0 && !isNaN(pctProtein) && isFinite(pctProtein) },
                { val: pctLemak, active: targetLemak > 0 && !isNaN(pctLemak) && isFinite(pctLemak) },
                { val: pctKH, active: targetKH > 0 && !isNaN(pctKH) && isFinite(pctKH) },
                { val: pctSerat, active: targetSerat > 0 && !isNaN(pctSerat) && isFinite(pctSerat) },
              ].map((pctInfo, idx) => {
                const isOk = pctInfo.active && pctInfo.val >= 90 && pctInfo.val <= 110;
                const cellBg = pctInfo.active ? (isOk ? '#00B050' : '#FF0000') : '#FF0000';
                const cellText = 'black';
                return (
                  <td 
                    key={idx} 
                    className="p-1.5 border border-black text-center font-mono font-black text-black" 
                    style={{ backgroundColor: cellBg, color: cellText }}
                  >
                    {pctInfo.active ? `${pctInfo.val.toFixed(0)}%` : ""}
                  </td>
                );
              })}

              {planningMode === "all_ingredients" ? (
                <>
                  <td colSpan={10} className="p-1.5 border border-black bg-slate-50/10"></td>
                  <td className="p-1 border border-black bg-slate-50/10"></td>
                  {!isPrint && (
                    <td className="p-1 border border-black bg-slate-50/10 no-print"></td>
                  )}
                </>
              ) : (
                <>
                  {/* Right Part: JUMLAH+BUMBU % (cols 11-22) */}
                  <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#F8CBAD]">
                    JUMLAH+BUMBU {activeBumbuPct}%
                  </td>
                  <td className="p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 bg-[#F8CBAD]">
                    <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950">
                      <span className="text-slate-600 font-black">Rp</span>
                      <span className="font-black">{formatRupiah(subtotalCost).replace("Rp", "").trim()}</span>
                    </div>
                  </td>
                  {!isPrint && (
                    <td className="p-1 border border-black border-l-2 border-l-slate-950 bg-[#F8CBAD] no-print"></td>
                  )}
                </>
              )}
            </tr>

            {/* ROW 4: HARGA PER PORSI */}
            {planningMode === "with_bumbu_10" && (
              <tr className="border border-black font-extrabold text-black">
                {/* Left Part: Empty (cols 1-10) */}
                <td colSpan={10} className="p-1.5 border border-black bg-slate-50/30"></td>

                {/* Right Part: HARGA PER PORSI (cols 11-22) */}
                <td colSpan={10} className="p-1.5 border border-black border-r-2 border-r-slate-950 text-right pr-4 uppercase text-slate-950 font-black bg-[#A9D08E]">
                  HARGA PER PORSI
                </td>
                <td className={`p-1 border border-black border-l-2 border-l-slate-950 border-r-2 border-r-slate-950 ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}>
                  <div className="flex items-center justify-between px-1 w-full font-mono text-xs font-black text-slate-950 print:text-black">
                    <span className="text-slate-950 print:text-black font-black">Rp</span>
                    <span className="font-black text-slate-950 print:text-black">{formatRupiah(calculatedCostPerPorsi).replace("Rp", "").trim()}</span>
                  </div>
                </td>
                {!isPrint && (
                  <td className={`p-1 border border-black border-l-2 border-l-slate-950 no-print ${calculatedCostPerPorsi > targetPrice ? "bg-[#FF0000]" : "bg-[#00B050]"}`}></td>
                )}
              </tr>
            )}
          </tfoot>
        </table>
      </div>
      {!isPrint &&
        renderOverBudgetNotice(
          `custom_${table.id}`,
          table.namaTabel || (table.porsi === "besar" ? "Porsi Besar Kustom" : "Porsi Kecil Kustom"),
          calculatedCostPerPorsi,
          targetPrice,
          recipientCount,
          items
        )}
      </>
    );
  };

  return (
    <div id="food-cost-container" className="space-y-6">
      {/* 1. Injected Print Styles */}
      <style dangerouslySetInnerHTML={{ __html: `
        @media print {
          @page {
            size: A4 landscape;
            margin: 6mm 8mm;
          }
          html, body, #sisper-app-root, main, #food-cost-container {
            height: auto !important;
            overflow: visible !important;
            max-width: none !important;
            margin: 0 !important;
            padding: 0 !important;
            background-color: white !important;
            color: black !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          .no-print, header, nav, footer, .sidebar, button, .tab-bar, #select-fc-day, #select-fc-group, #select-fc-type {
            display: none !important;
          }
          #print-area-food-cost {
            display: block !important;
            width: 100% !important;
            position: relative !important;
            left: 0 !important;
            right: 0 !important;
            top: 0 !important;
            margin: 0 auto !important;
            padding: 0 !important;
          }
          /* Centered block layout */
          #print-area-food-cost > div {
            margin: 0 auto !important;
            width: 100% !important;
            max-width: 100% !important;
            display: block !important;
          }
          /* Enforce compact layout and prevent spill */
          #print-area-food-cost table {
            width: 100% !important;
            table-layout: fixed !important;
            border-collapse: collapse !important;
            margin: 0 auto !important;
          }
          #print-area-food-cost th, #print-area-food-cost td {
            font-size: 8px !important;
            padding: 2px 2px !important;
            line-height: 1.12 !important;
            word-wrap: break-word !important;
            white-space: normal !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
            border: 1px solid black !important;
          }

          input, select {
            border: none !important;
            background: transparent !important;
            box-shadow: none !important;
            outline: none !important;
            padding: 0 !important;
            font-size: inherit !important;
            color: black !important;
          }
        }
      `}} />

      {/* 2. Top Navigation header & Mode Switcher */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 pb-4 no-print">
        <div>
          <h3 className="text-base font-black text-slate-800 flex items-center gap-2">
            <Calculator className="w-5 h-5 text-indigo-600" />
            ANALISIS & KALKULATOR FOOD COST SPPG
          </h3>
          <p className="text-xs text-slate-500">
            Kalkulasi unit cost, kecukupan gizi rujukan AKG, penyusutan (buffer), dan ekspor A4 Landscape presisi.
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2 shrink-0 self-end sm:self-center">
          {onSendToNotaPesanan && (
            <button
              type="button"
              id="btn-send-fc-to-nota"
              onClick={() => {
                onFoodCostDaysChange([...foodCostDays]);
                onSendToNotaPesanan();
              }}
              className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-2 transition cursor-pointer whitespace-nowrap shadow-xs"
              title="Kirim perencanaan Food Cost dan langsung menuju ke panel Nota Pesanan"
            >
              <Send className="w-3.5 h-3.5" />
              Kirim Perencanaan
            </button>
          )}
          <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 shadow-xs items-center gap-1.5">
            <button
              type="button"
              id="btn-mode-fc-edit"
              onClick={() => setViewMode("edit")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 cursor-pointer ${
                viewMode === "edit" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500 hover:text-slate-800"
              }`}
            >
              <Edit3 className="w-3.5 h-3.5" />
              Mode Edit Kalkulator
            </button>
            <button
              type="button"
              id="btn-mode-fc-print"
              onClick={() => setViewMode("print")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 cursor-pointer ${
                viewMode === "print" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500 hover:text-slate-800"
              }`}
            >
              <Printer className="w-3.5 h-3.5" />
              Pratinjau Cetak {paperSize} / Ekspor
            </button>
          </div>
        </div>
      </div>

      {/* Inline Non-Blocking Feedback Banner */}
      {fcNotice && (
        <div
          className={`px-4 py-3 rounded-xl border flex items-center justify-between gap-3 text-xs font-bold no-print ${
            fcNotice.type === "warning"
              ? "bg-amber-50 border-amber-300 text-amber-900"
              : fcNotice.type === "success"
                ? "bg-emerald-50 border-emerald-300 text-emerald-900"
                : "bg-indigo-50 border-indigo-200 text-indigo-900"
          }`}
        >
          <span className="flex items-center gap-2">
            {fcNotice.type === "warning" ? (
              <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
            ) : fcNotice.type === "success" ? (
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            ) : (
              <Info className="w-4 h-4 text-indigo-600 shrink-0" />
            )}
            {fcNotice.text}
          </span>
          <button
            type="button"
            onClick={() => setFcNotice(null)}
            className="p-1 rounded-lg hover:bg-black/5 transition cursor-pointer shrink-0"
            title="Tutup pesan"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* --- VIEW MODE: EDIT INTERAKTIF --- */}
      {viewMode === "edit" && (
        <div className="space-y-6 no-print">

          {/* Pilah Kategori Porsi Kecil & Porsi Besar + Masing-Masing 1 Kategori / Gabung Pilihan */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
            <div className="p-4 sm:px-5 bg-slate-50/80 border-b border-slate-200/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-start sm:items-center gap-3">
                <div className="p-2 rounded-xl bg-indigo-100 text-indigo-700 shrink-0">
                  <Layers className="w-4 h-4" />
                </div>
                <div>
                  <div className="flex flex-wrap items-center gap-2">
                    <h4 className="text-xs font-black text-slate-900 uppercase tracking-wide">
                      Pilih & Pilah Kategori Food Cost
                    </h4>
                    {appliedConfig.confirmed ? (
                      isPerKategoriMode ? (
                        <span className="px-2 py-0.5 rounded-md bg-emerald-100 text-emerald-800 border border-emerald-200 text-[10px] font-black uppercase">
                          Aktif: Masing-Masing Kategori ({resolvedCategoryEntries.length} Panel Food Cost)
                        </span>
                      ) : (
                        <span className="px-2 py-0.5 rounded-md bg-indigo-100 text-indigo-800 border border-indigo-200 text-[10px] font-black uppercase">
                          Aktif: Gabung{" "}
                          {appliedConfig.scope === "kecil"
                            ? "Porsi Kecil"
                            : appliedConfig.scope === "besar"
                              ? "Porsi Besar"
                              : "Semua Porsi (Kecil & Besar)"}
                        </span>
                      )
                    ) : (
                      <span className="px-2 py-0.5 rounded-md bg-amber-100 text-amber-800 border border-amber-200 text-[10px] font-black uppercase">
                        Menunggu Pemilihan & Konfirmasi
                      </span>
                    )}
                  </div>
                  <p className="text-[11px] text-slate-500 mt-0.5">
                    {isPerKategoriMode
                      ? "Tampilan default: masing-masing 1 kategori sasaran memiliki panel Food Cost sendiri. Anda dapat mencentang & menyatukan beberapa kategori menjadi 1 Food Cost kapan saja."
                      : "Mode gabung per porsi: seluruh kategori Porsi Kecil dan Porsi Besar digabung menjadi maksimal 2 tabel Food Cost."}
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setIsSelectorExpanded((prev) => !prev)}
                className="px-3.5 py-2 bg-white hover:bg-slate-100 border border-slate-200 text-slate-700 rounded-xl text-xs font-bold flex items-center gap-1.5 self-start sm:self-center shrink-0 cursor-pointer transition"
              >
                {isSelectorExpanded ? (
                  <>
                    <ChevronUp className="w-3.5 h-3.5" />
                    Tutup Pengaturan Kategori
                  </>
                ) : (
                  <>
                    <ChevronDown className="w-3.5 h-3.5" />
                    Atur / Satukan Kategori
                  </>
                )}
              </button>
            </div>

            {isSelectorExpanded && (
              <div className="p-5 space-y-5">
                {/* Step 1: Layout Mode Switcher & Scope Filter */}
                <div className="space-y-3">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                      1. Pilih Mode Panel Food Cost
                    </span>
                    <div className="flex flex-wrap items-center gap-1.5">
                      <span className="text-[10px] font-bold text-slate-400 uppercase mr-1">Filter Porsi:</span>
                      {(
                        [
                          { id: "all", label: "Semua Porsi" },
                          { id: "kecil", label: "Hanya Porsi Kecil" },
                          { id: "besar", label: "Hanya Porsi Besar" },
                        ] as const
                      ).map((sc) => (
                        <button
                          key={sc.id}
                          type="button"
                          onClick={() => {
                            handleSelectDraftScope(sc.id);
                            if (draftLayoutMode === "per_kategori") {
                              setAppliedConfig((prev) => ({
                                ...prev,
                                confirmed: true,
                                layoutMode: "per_kategori",
                                scope: sc.id,
                              }));
                            }
                          }}
                          className={`px-2.5 py-1 rounded-lg text-[10px] font-extrabold border transition cursor-pointer ${
                            draftScope === sc.id
                              ? "bg-slate-900 text-white border-slate-900"
                              : "bg-white text-slate-600 border-slate-200 hover:border-slate-300"
                          }`}
                        >
                          {sc.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <button
                      type="button"
                      onClick={() => handleSwitchLayoutMode("per_kategori")}
                      className={`p-3.5 rounded-xl border text-left transition cursor-pointer ${
                        draftLayoutMode === "per_kategori"
                          ? "bg-emerald-50 border-emerald-500 ring-2 ring-emerald-500/20"
                          : "bg-slate-50/70 border-slate-200 hover:border-emerald-300"
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-black text-emerald-950 uppercase flex items-center gap-1.5">
                          <Layers className="w-3.5 h-3.5 text-emerald-600" />
                          Masing-Masing 1 Kategori
                        </span>
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">
                          Bisa Satukan Pilihan
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-600 mt-1">
                        Setiap kategori penerima manfaat tampil terpisah 1 Food Cost. Centang 2+ kategori di bawah jika ingin menyatukannya menjadi 1 Food Cost.
                      </p>
                    </button>

                    <button
                      type="button"
                      onClick={() => handleSwitchLayoutMode("gabung_porsi")}
                      className={`p-3.5 rounded-xl border text-left transition cursor-pointer ${
                        draftLayoutMode === "gabung_porsi"
                          ? "bg-indigo-50 border-indigo-500 ring-2 ring-indigo-500/20"
                          : "bg-slate-50/70 border-slate-200 hover:border-indigo-300"
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-black text-indigo-950 uppercase flex items-center gap-1.5">
                          <Combine className="w-3.5 h-3.5 text-indigo-600" />
                          Gabung Semua per Porsi (Besar & Kecil)
                        </span>
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-indigo-100 text-indigo-800">
                          2 Panel Gabungan
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-600 mt-1">
                        Menggabungkan seluruh kategori Porsi Kecil dan Porsi Besar ke dalam 2 tabel Food Cost gabungan (Besar & Kecil).
                      </p>
                    </button>
                  </div>
                </div>

                {/* Step 2A: Per-Kategori + Selective Merge UI */}
                {draftLayoutMode === "per_kategori" ? (
                  (() => {
                    const currentGroups =
                      draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey);
                    const assignedIds = new Set(currentGroups.flatMap((g) => g.sasaranIds));
                    const omittedCategories = ALL_KATEGORI_WITH_ALERGI.filter((s) => !assignedIds.has(s.id));

                    return (
                      <div className="space-y-3 pt-2 border-t border-slate-100">
                        <div className="flex flex-wrap items-center justify-between gap-2">
                          <div>
                            <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider block">
                              2. Kelola Panel Kategori & Satukan Kategori Terpilih
                            </span>
                            <p className="text-[11px] text-slate-500">
                              Centang minimal 2 kartu kategori (Standar maupun{" "}
                              <strong className="text-orange-700">Diet Alergi</strong>) lalu klik{" "}
                              <strong className="text-indigo-700">Satukan Kategori Terpilih</strong> untuk menggabungkan jadi 1 Food Cost.
                            </p>
                          </div>

                          <div className="flex flex-wrap items-center gap-1.5">
                            <button
                              type="button"
                              onClick={handleKeepOnlyPmPositiveCategories}
                              className="px-2.5 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 text-[10px] font-extrabold transition cursor-pointer"
                            >
                              Hanya Ada PM &gt; 0
                            </button>
                            <button
                              type="button"
                              onClick={() => handleResetToSingleCategories(false)}
                              className="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 text-[10px] font-extrabold transition cursor-pointer"
                            >
                              Sesuai Kelompok + Alergi
                            </button>
                            <button
                              type="button"
                              onClick={() => handleResetToSingleCategories(true)}
                              className="px-2.5 py-1.5 rounded-lg bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-700 text-[10px] font-extrabold transition cursor-pointer"
                            >
                              Semua 24 Kategori (Standar &amp; Alergi)
                            </button>
                            <button
                              type="button"
                              onClick={handleAddAllAlergiCategories}
                              className="px-2.5 py-1.5 rounded-lg bg-orange-50 hover:bg-orange-100 border border-orange-200 text-orange-800 text-[10px] font-extrabold transition cursor-pointer"
                            >
                              + Kategori Alergi
                            </button>
                            <button
                              type="button"
                              onClick={() => setShowCrossGroupPicker((prev) => !prev)}
                              className={`px-2.5 py-1.5 rounded-lg border text-[10px] font-extrabold transition cursor-pointer ${
                                showCrossGroupPicker
                                  ? "bg-amber-100 border-amber-300 text-amber-900"
                                  : "bg-white border-slate-200 text-slate-600 hover:bg-slate-50"
                              }`}
                            >
                              + Tambah Kategori Lainnya ({omittedCategories.length})
                            </button>
                          </div>
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5">
                          {currentGroups.map((g) => {
                            const isChecked = mergeSelectedGroupIds.includes(g.id);
                            const autoPm = getEntryAutoPmCount(g.sasaranIds, g.porsi);
                            const isMerged = g.sasaranIds.length > 1;
                            const isAlergiGroup = g.sasaranIds.some((id) => isAlergiSasaranId(id));
                            const isHiddenByScope = draftScope !== "all" && g.porsi !== draftScope;

                            return (
                              <div
                                key={g.id}
                                className={`rounded-xl border p-3 transition flex flex-col justify-between gap-2 ${
                                  isChecked
                                    ? "bg-indigo-50/80 border-indigo-400 ring-2 ring-indigo-500/20 shadow-xs"
                                    : isAlergiGroup
                                      ? "bg-orange-50/50 border-orange-300"
                                      : isMerged
                                        ? "bg-amber-50/50 border-amber-300"
                                        : "bg-slate-50/60 border-slate-200 hover:border-slate-300"
                                } ${isHiddenByScope ? "opacity-55" : ""}`}
                              >
                                <div className="flex items-start justify-between gap-2">
                                  <label className="flex items-start gap-2 cursor-pointer min-w-0 flex-1">
                                    <input
                                      type="checkbox"
                                      checked={isChecked}
                                      onChange={() => toggleMergeSelection(g.id)}
                                      className="mt-0.5 rounded border-indigo-300 text-indigo-600 focus:ring-indigo-500 shrink-0"
                                    />
                                    <div className="min-w-0">
                                      <span className="text-xs font-black text-slate-900 block break-words leading-snug">
                                        {g.title}
                                      </span>
                                      <div className="flex flex-wrap items-center gap-1 mt-1">
                                        <span
                                          className={`px-1.5 py-0.5 rounded text-[9px] font-black uppercase ${
                                            g.porsi === "besar"
                                              ? "bg-amber-100 text-amber-800"
                                              : "bg-cyan-100 text-cyan-800"
                                          }`}
                                        >
                                          {g.porsi === "besar" ? "Porsi Besar" : "Porsi Kecil"}
                                        </span>
                                        {isAlergiGroup && (
                                          <span className="px-1.5 py-0.5 rounded bg-orange-100 text-orange-800 text-[9px] font-black uppercase">
                                            Diet Alergi
                                          </span>
                                        )}
                                        <span className="px-1.5 py-0.5 rounded bg-white border border-slate-200 text-[9px] font-mono font-bold text-slate-700">
                                          {autoPm} PM
                                        </span>
                                        {isMerged && (
                                          <span className="px-1.5 py-0.5 rounded bg-indigo-100 text-indigo-800 text-[9px] font-black">
                                            Gabungan {g.sasaranIds.length}
                                          </span>
                                        )}
                                      </div>
                                    </div>
                                  </label>

                                  <div className="flex items-center gap-1 shrink-0">
                                    {isMerged && (
                                      <button
                                        type="button"
                                        onClick={() => handleSplitMergedGroup(g.id)}
                                        title="Pisahkan kembali menjadi kategori masing-masing"
                                        className="px-2 py-1 rounded-lg bg-white hover:bg-amber-100 border border-amber-300 text-amber-800 text-[10px] font-extrabold flex items-center gap-1 cursor-pointer transition"
                                      >
                                        <Split className="w-3 h-3" />
                                        Pisahkan
                                      </button>
                                    )}
                                    {currentGroups.length > 1 && (
                                      <button
                                        type="button"
                                        onClick={() => handleRemoveGroupFromActive(g.id)}
                                        title="Sembunyikan kategori ini dari tampilan"
                                        className="p-1 rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition cursor-pointer"
                                      >
                                        <X className="w-3.5 h-3.5" />
                                      </button>
                                    )}
                                  </div>
                                </div>

                                {isMerged && (
                                  <div className="text-[10px] text-slate-600 bg-white/80 rounded-lg px-2 py-1 border border-amber-200/70 font-semibold">
                                    Termasuk: {g.sasaranIds.map((id) => getSasaranLabel(id)).join(", ")}
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>

                        {/* Merge Selected Categories Action Bar */}
                        {mergeSelectedGroupIds.length >= 2 && (
                          <div className="p-3.5 rounded-xl bg-indigo-50 border-2 border-indigo-300 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
                            <div className="flex flex-col sm:flex-row sm:items-center gap-2 flex-1">
                              <span className="text-xs font-black text-indigo-950 shrink-0">
                                {mergeSelectedGroupIds.length} Kategori Dipilih:
                              </span>
                              <input
                                type="text"
                                value={mergeCustomTitle}
                                onChange={(e) => setMergeCustomTitle(e.target.value)}
                                placeholder="Nama Food Cost gabungan (opsional, kosongkan agar otomatis)"
                                className="flex-1 bg-white border border-indigo-200 rounded-lg px-2.5 py-1.5 text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500/30"
                              />
                              <select
                                value={mergeCustomPorsi}
                                onChange={(e) => setMergeCustomPorsi(e.target.value as "auto" | "kecil" | "besar")}
                                className="bg-white border border-indigo-200 rounded-lg px-2.5 py-1.5 text-xs font-bold text-slate-800 focus:outline-none"
                              >
                                <option value="auto">Plafon &amp; Porsi Otomatis</option>
                                <option value="besar">Tetapkan Porsi Besar</option>
                                <option value="kecil">Tetapkan Porsi Kecil</option>
                              </select>
                            </div>
                            <div className="flex items-center gap-2 shrink-0">
                              <button
                                type="button"
                                onClick={handleMergeSelectedDraftGroups}
                                className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-black flex items-center gap-1.5 shadow-xs transition cursor-pointer"
                              >
                                <Combine className="w-3.5 h-3.5" />
                                Satukan {mergeSelectedGroupIds.length} Kategori Terpilih menjadi 1 Food Cost
                              </button>
                              <button
                                type="button"
                                onClick={() => setMergeSelectedGroupIds([])}
                                className="px-3 py-2 bg-white hover:bg-slate-100 border border-slate-200 text-slate-600 rounded-xl text-xs font-bold transition cursor-pointer"
                              >
                                Batal Pilih
                              </button>
                            </div>
                          </div>
                        )}

                        {/* Omitted / Cross-Group & Alergi Categories Picker */}
                        {(showCrossGroupPicker || omittedCategories.length > 0) && omittedCategories.length > 0 && (
                          <div className="p-3 rounded-xl bg-slate-50 border border-slate-200/80 space-y-2">
                            <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider block">
                              Tambahkan Kategori Lain / Kategori Alergi ke Panel Food Cost:
                            </span>
                            <div className="flex flex-wrap gap-1.5">
                              {omittedCategories.map((s) => {
                                const defPorsi = resolveDefaultPorsiForSasaran(s.id);
                                const cnt = getSasaranPmCount(s.id, defPorsi);
                                const isAlg = !!s.isAlergi;
                                return (
                                  <button
                                    key={s.id}
                                    type="button"
                                    onClick={() => handleAddSingleCategoryToActive(s.id)}
                                    className={`px-2.5 py-1 rounded-lg border text-[11px] font-bold transition cursor-pointer inline-flex items-center gap-1 ${
                                      isAlg
                                        ? "bg-orange-50/70 hover:bg-orange-100 border-orange-200 text-orange-900"
                                        : "bg-white hover:bg-indigo-50 border-slate-200 hover:border-indigo-300 text-slate-700 hover:text-indigo-800"
                                    }`}
                                  >
                                    <Plus className={`w-3 h-3 ${isAlg ? "text-orange-600" : "text-indigo-600"}`} />
                                    {s.label}
                                    <span className={`text-[10px] font-mono ${isAlg ? "text-orange-700" : "text-slate-400"}`}>
                                      ({cnt})
                                    </span>
                                  </button>
                                );
                              })}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()
                ) : (
                  /* Step 2B: Separated Porsi Kecil & Porsi Besar Checklists (Gabung per Porsi) */
                  <div className="space-y-2 pt-2 border-t border-slate-100">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                        2. Pilah Kategori Penerima Manfaat yang Diterapkan (Mode Gabung per Porsi)
                      </span>
                      {onNavigateToPM && (
                        <button
                          type="button"
                          onClick={onNavigateToPM}
                          className="text-[11px] font-bold text-indigo-600 hover:text-indigo-800 cursor-pointer"
                        >
                          Ubah Klasifikasi Porsi di Data PM →
                        </button>
                      )}
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                      {/* Porsi Kecil Categories Card */}
                      <div
                        className={`rounded-xl border p-4 space-y-3 transition ${
                          draftScope === "besar"
                            ? "bg-slate-50 border-slate-200 opacity-55"
                            : "bg-indigo-50/40 border-indigo-200"
                        }`}
                      >
                        <div className="flex items-center justify-between gap-2 border-b border-indigo-200/60 pb-2">
                          <div>
                            <span className="text-xs font-black text-indigo-950 uppercase">
                              Kategori Porsi Kecil
                            </span>
                            <p className="text-[10px] text-indigo-700 font-semibold">
                              Plafon {formatRupiah(activeTargetPriceKecil)}/porsi · Terpilih {draftKecilIds.length}/{defaultKecilIds.length}
                            </p>
                          </div>
                          <div className="flex items-center gap-1.5">
                            <button
                              type="button"
                              disabled={draftScope === "besar"}
                              onClick={() => setDraftKecilIds(defaultKecilIds)}
                              className="px-2.5 py-1 rounded-lg bg-white border border-indigo-200 text-[10px] font-bold text-indigo-700 hover:bg-indigo-50 disabled:opacity-50 cursor-pointer"
                            >
                              Pilih Semua
                            </button>
                            <button
                              type="button"
                              disabled={draftScope === "besar"}
                              onClick={() => setDraftKecilIds([])}
                              className="px-2.5 py-1 rounded-lg bg-white border border-slate-200 text-[10px] font-bold text-slate-600 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
                            >
                              Kosongkan
                            </button>
                          </div>
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                          {defaultKecilIds.map((sId) => {
                            const checked = draftKecilIds.includes(sId);
                            const label = getSasaranLabel(sId);
                            const count = getSasaranPmCount(sId, "kecil");
                            return (
                              <label
                                key={sId}
                                className={`flex items-center justify-between gap-2 px-3 py-2 rounded-lg border text-xs cursor-pointer transition ${
                                  checked
                                    ? "bg-white border-indigo-400 text-indigo-950 font-bold shadow-2xs"
                                    : "bg-white/60 border-slate-200 text-slate-600 font-medium"
                                }`}
                              >
                                <span className="flex items-center gap-2 min-w-0">
                                  <input
                                    type="checkbox"
                                    disabled={draftScope === "besar"}
                                    checked={checked}
                                    onChange={() => toggleDraftCategory("kecil", sId)}
                                    className="rounded border-indigo-300 text-indigo-600 focus:ring-indigo-500"
                                  />
                                  <span className="truncate">{label}</span>
                                </span>
                                <span className="text-[10px] font-mono font-bold text-indigo-700 shrink-0">
                                  {count} PM
                                </span>
                              </label>
                            );
                          })}
                        </div>
                      </div>

                      {/* Porsi Besar Categories Card */}
                      <div
                        className={`rounded-xl border p-4 space-y-3 transition ${
                          draftScope === "kecil"
                            ? "bg-slate-50 border-slate-200 opacity-55"
                            : "bg-cyan-50/40 border-cyan-200"
                        }`}
                      >
                        <div className="flex items-center justify-between gap-2 border-b border-cyan-200/60 pb-2">
                          <div>
                            <span className="text-xs font-black text-cyan-950 uppercase">
                              Kategori Porsi Besar
                            </span>
                            <p className="text-[10px] text-cyan-800 font-semibold">
                              Plafon {formatRupiah(activeTargetPriceBesar)}/porsi · Terpilih {draftBesarIds.length}/{defaultBesarIds.length}
                            </p>
                          </div>
                          <div className="flex items-center gap-1.5">
                            <button
                              type="button"
                              disabled={draftScope === "kecil"}
                              onClick={() => setDraftBesarIds(defaultBesarIds)}
                              className="px-2.5 py-1 rounded-lg bg-white border border-cyan-200 text-[10px] font-bold text-cyan-800 hover:bg-cyan-50 disabled:opacity-50 cursor-pointer"
                            >
                              Pilih Semua
                            </button>
                            <button
                              type="button"
                              disabled={draftScope === "kecil"}
                              onClick={() => setDraftBesarIds([])}
                              className="px-2.5 py-1 rounded-lg bg-white border border-slate-200 text-[10px] font-bold text-slate-600 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
                            >
                              Kosongkan
                            </button>
                          </div>
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                          {defaultBesarIds.map((sId) => {
                            const checked = draftBesarIds.includes(sId);
                            const label = getSasaranLabel(sId);
                            const count = getSasaranPmCount(sId, "besar");
                            return (
                              <label
                                key={sId}
                                className={`flex items-center justify-between gap-2 px-3 py-2 rounded-lg border text-xs cursor-pointer transition ${
                                  checked
                                    ? "bg-white border-cyan-500 text-cyan-950 font-bold shadow-2xs"
                                    : "bg-white/60 border-slate-200 text-slate-600 font-medium"
                                }`}
                              >
                                <span className="flex items-center gap-2 min-w-0">
                                  <input
                                    type="checkbox"
                                    disabled={draftScope === "kecil"}
                                    checked={checked}
                                    onChange={() => toggleDraftCategory("besar", sId)}
                                    className="rounded border-cyan-300 text-cyan-600 focus:ring-cyan-500"
                                  />
                                  <span className="truncate">{label}</span>
                                </span>
                                <span className="text-[10px] font-mono font-bold text-cyan-800 shrink-0">
                                  {count} PM
                                </span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* Step 3: Confirmation CTA */}
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pt-3 border-t border-slate-100">
                  <div className="text-[11px] text-slate-600 font-semibold">
                    Mode aktif dipilih:{" "}
                    <span className="font-black text-slate-900">
                      {draftLayoutMode === "per_kategori"
                        ? `Masing-Masing 1 Kategori (${(draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey)).filter((g) => draftScope === "all" || g.porsi === draftScope).length} Panel)`
                        : draftScope === "kecil"
                          ? `Gabung Porsi Kecil (${draftKecilIds.length} kategori)`
                          : draftScope === "besar"
                            ? `Gabung Porsi Besar (${draftBesarIds.length} kategori)`
                            : `Gabung Semua Food Cost (${draftKecilIds.length} Kecil + ${draftBesarIds.length} Besar)`}
                    </span>
                  </div>
                  <button
                    type="button"
                    id="btn-confirm-fc-categories"
                    onClick={handleOpenCategoryConfirmModal}
                    className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-black flex items-center justify-center gap-2 shadow-sm transition cursor-pointer shrink-0"
                  >
                    <CheckCircle2 className="w-4 h-4" />
                    Konfirmasi &amp; Tampilkan Food Cost Otomatis
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* RAB & Realisasi Belanja Dashboard */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {/* 1. Rencana Anggaran Biaya (RAB) */}
            <div className="bg-gradient-to-br from-emerald-50 to-emerald-100/40 p-4 rounded-2xl border border-emerald-100 shadow-xs flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[10px] font-black text-emerald-800 uppercase tracking-wider block">
                    Rencana Anggaran Biaya (RAB Harian)
                  </span>
                  {onNavigateToPM && (
                    <button
                      type="button"
                      onClick={onNavigateToPM}
                      className="text-[9px] font-bold bg-emerald-600/10 hover:bg-emerald-600/20 text-emerald-800 px-2 py-0.5 rounded-md transition cursor-pointer shrink-0"
                      title="Atur harga per porsi & klasifikasi sasaran di menu Penerima Manfaat"
                    >
                      Atur di Data PM →
                    </button>
                  )}
                </div>
                <div className="mt-2 flex items-baseline gap-1">
                  <span className="text-xl font-black text-emerald-900 font-mono">
                    Rp {formatThousandSeparator(totalRAB)}
                  </span>
                </div>
              </div>
              <div className="text-[10px] text-emerald-700 font-semibold mt-3 space-y-0.5 border-t border-emerald-200/50 pt-2 font-mono">
                {includeKecilInDashboard && (
                  <div>Kecil: {pmKecil} PM x Rp {formatThousandSeparator(activeTargetPriceKecil)} = Rp {formatThousandSeparator(rabPorsiKecil)}</div>
                )}
                {includeBesarInDashboard && (
                  <div>Besar: {pmBesar} PM x Rp {formatThousandSeparator(activeTargetPriceBesar)} = Rp {formatThousandSeparator(rabPorsiBesar)}</div>
                )}
                {customTables && customTables.map((t) => {
                  const tRAB = (Number(t.pmCount) || 0) * (Number(t.plafon) || (t.porsi === "besar" ? activeTargetPriceBesar : activeTargetPriceKecil));
                  return (
                    <div key={t.id} className="text-purple-800 font-bold">
                      {t.namaTabel}: {t.pmCount || 0} PM x Rp {formatThousandSeparator(t.plafon || 0)} = Rp {formatThousandSeparator(tRAB)}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* 2. Realisasi Belanja Harian */}
            <div className="bg-gradient-to-br from-indigo-50 to-indigo-100/40 p-4 rounded-2xl border border-indigo-100 shadow-xs flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between">
                  <span className="text-[10px] font-black text-indigo-800 uppercase tracking-wider block">
                    Realisasi Belanja Harian (Actual)
                  </span>
                  <span className="text-[9px] font-black px-1.5 py-0.5 rounded bg-indigo-100 text-indigo-800 border border-indigo-200">
                    {planningMode === "with_bumbu_10" ? `+${activeBumbuPct}% Bumbu` : "Uraian Bumbu"}
                  </span>
                </div>
                <div className="mt-2 flex items-baseline justify-between gap-1">
                  <span className="text-xl font-black text-indigo-900 font-mono">
                    Rp {formatThousandSeparator(totalActualCost)}
                  </span>
                  {planningMode === "with_bumbu_10" && (
                    <div className="text-right">
                      <span className="text-[9px] font-extrabold text-indigo-600 block uppercase tracking-tight">Total Bumbu {activeBumbuPct}%:</span>
                      <span className="text-xs font-black text-indigo-900 font-mono">
                        Rp {formatThousandSeparator(totalBumbu10Nominal)}
                      </span>
                    </div>
                  )}
                </div>
              </div>
              <div className="text-[10px] text-indigo-700 font-semibold mt-3 space-y-1.5 border-t border-indigo-200/50 pt-2 font-mono">
                {includeKecilInDashboard && (
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1">
                    <span className="flex items-center gap-1.5 flex-wrap">
                      <span>Kecil: Rp {formatThousandSeparator(porsiKecilActualCost)} (Rp {formatThousandSeparator(porsiKecilPerPorsi)}/porsi)</span>
                      {isOverBudgetKecil && (
                        <span className="px-1.5 py-0.5 rounded bg-rose-600 text-white text-[9px] font-black">
                          OVER BUDGET
                        </span>
                      )}
                    </span>
                    {planningMode === "with_bumbu_10" && (
                      <span className="text-[9.5px] font-bold text-indigo-900 bg-indigo-100/90 px-1.5 py-0.5 rounded border border-indigo-200/70 shrink-0 w-fit">
                        Bumbu {activeBumbuPct}%: Rp {formatThousandSeparator(bumbuKecilNominal)}
                      </span>
                    )}
                  </div>
                )}

                {includeBesarInDashboard && (
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1">
                    <span className="flex items-center gap-1.5 flex-wrap">
                      <span>Besar: Rp {formatThousandSeparator(porsiBesarActualCost)} (Rp {formatThousandSeparator(porsiBesarPerPorsi)}/porsi)</span>
                      {isOverBudgetBesar && (
                        <span className="px-1.5 py-0.5 rounded bg-rose-600 text-white text-[9px] font-black">
                          OVER BUDGET
                        </span>
                      )}
                    </span>
                    {planningMode === "with_bumbu_10" && (
                      <span className="text-[9.5px] font-bold text-indigo-900 bg-indigo-100/90 px-1.5 py-0.5 rounded border border-indigo-200/70 shrink-0 w-fit">
                        Bumbu {activeBumbuPct}%: Rp {formatThousandSeparator(bumbuBesarNominal)}
                      </span>
                    )}
                  </div>
                )}

                {customTables && customTables.map((t) => {
                  const calcCustom = t.porsi === "besar"
                    ? calculateDay(t.bahanList || [], [], Number(t.pmCount) || 0, 0, t.bufferPct || 5, tkpiList, activeBumbuPct)
                    : calculateDay([], t.bahanList || [], 0, Number(t.pmCount) || 0, t.bufferPct || 5, tkpiList, activeBumbuPct);
                  const tActual = t.porsi === "besar"
                    ? (planningMode === "with_bumbu_10" ? calcCustom.subtotalBesarCost : calcCustom.totalBesarBahanCost)
                    : (planningMode === "with_bumbu_10" ? calcCustom.subtotalKecilCost : calcCustom.totalKecilBahanCost);
                  const tPerPorsi = t.pmCount > 0 ? Math.round(tActual / t.pmCount) : 0;
                  const tBumbuNominal = t.porsi === "besar" ? calcCustom.bumbuBesarCost : calcCustom.bumbuKecilCost;

                  return (
                    <div key={t.id} className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 text-purple-800 font-bold">
                      <span>{t.namaTabel}: Rp {formatThousandSeparator(tActual)} (Rp {formatThousandSeparator(tPerPorsi)}/porsi)</span>
                      {planningMode === "with_bumbu_10" && (
                        <span className="text-[9.5px] font-bold text-purple-900 bg-purple-100/90 px-1.5 py-0.5 rounded border border-purple-200/70 shrink-0 w-fit">
                          Bumbu {activeBumbuPct}%: Rp {formatThousandSeparator(tBumbuNominal)}
                        </span>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* 3. Sisa / Efisiensi Saldo */}
            <div className={`p-4 rounded-2xl border shadow-xs flex flex-col justify-between ${
              sisaEfisiensiAnggaran >= 0 
                ? "bg-gradient-to-br from-teal-50 to-teal-100/40 border-teal-100" 
                : "bg-gradient-to-br from-rose-50 to-rose-100/40 border-rose-100"
            }`}>
              <div>
                <span className={`text-[10px] font-black uppercase tracking-wider block ${
                  sisaEfisiensiAnggaran >= 0 ? "text-teal-800" : "text-rose-800"
                }`}>
                  {sisaEfisiensiAnggaran >= 0 ? "Sisa / Efisiensi Anggaran (Surplus)" : "Defisit / Melebihi Anggaran"}
                </span>
                <div className="mt-2 flex items-baseline gap-1">
                  <span className={`text-xl font-black font-mono ${
                    sisaEfisiensiAnggaran >= 0 ? "text-teal-900" : "text-rose-900"
                  }`}>
                    {sisaEfisiensiAnggaran >= 0 ? "+" : ""}{formatRupiah(sisaEfisiensiAnggaran)}
                  </span>
                </div>
              </div>
              <div className={`text-[10px] font-semibold mt-3 border-t pt-2 ${
                sisaEfisiensiAnggaran >= 0 
                  ? "text-teal-700 border-teal-200/50" 
                  : "text-rose-700 border-rose-200/50"
              }`}>
                {sisaEfisiensiAnggaran >= 0 
                  ? "✓ Komponen belanja aman di bawah plafon harian." 
                  : "⚠️ Perhatian: Anggaran belanja melebihi pagu dana yang diatur!"}
              </div>
            </div>
          </div>

          {/* Panel Tanggal Perencanaan (Tepat di Dekat Tabel Food Cost) */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-xs p-4 sm:p-5 space-y-3">
            {/* Pilih Tanggal Perencanaan + Navigasi Harian */}
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
              <div className="flex items-start sm:items-center gap-3">
                <div className="p-2.5 rounded-xl bg-indigo-100 text-indigo-700 shrink-0">
                  <Calendar className="w-4 h-4" />
                </div>
                <div className="space-y-0.5">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                      Tanggal Perencanaan Food Cost
                    </span>
                    <span className="px-2 py-0.5 rounded-md bg-indigo-50 border border-indigo-200 text-[10px] font-black text-indigo-800">
                      Hari Ke-{selectedDay} / {totalDays}
                    </span>
                  </div>
                  <p className="text-sm font-black text-slate-900">
                    {currentCalendarDate}
                  </p>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  disabled={selectedDay <= 1}
                  onClick={() => setSelectedDay((d) => Math.max(1, d - 1))}
                  className="p-2 rounded-xl bg-slate-50 hover:bg-slate-100 disabled:opacity-40 border border-slate-200 text-slate-700 transition cursor-pointer disabled:cursor-not-allowed"
                  title="Tanggal perencanaan sebelumnya"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>

                <select
                  id="select-fc-day"
                  value={selectedDay}
                  onChange={(e) => setSelectedDay(Number(e.target.value))}
                  className="bg-slate-50 border border-slate-200 rounded-xl text-xs px-3 py-2 text-slate-900 font-extrabold focus:outline-none focus:ring-2 focus:ring-indigo-500/20 min-w-[230px] h-9 cursor-pointer"
                  title="Pilih sesuai tanggal perencanaan"
                >
                  {Array.from({ length: totalDays }).map((_, i) => {
                    const dayNum = i + 1;
                    const dayOverrideDate = foodCostDays.find(
                      (d) => d.hariKe === dayNum && d.jenisMenu === activeDbType
                    )?.tanggal;
                    const optDate = dayOverrideDate || scheduleDates[i] || "";
                    return (
                      <option key={dayNum} value={dayNum}>
                        {optDate
                          ? `${formatScheduleIndonesianDate(optDate)} (Hari Ke-${dayNum})`
                          : `Hari Ke-${dayNum}`}
                      </option>
                    );
                  })}
                </select>

                <button
                  type="button"
                  disabled={selectedDay >= totalDays}
                  onClick={() => setSelectedDay((d) => Math.min(totalDays, d + 1))}
                  className="p-2 rounded-xl bg-slate-50 hover:bg-slate-100 disabled:opacity-40 border border-slate-200 text-slate-700 transition cursor-pointer disabled:cursor-not-allowed"
                  title="Tanggal perencanaan berikutnya"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>

                <div
                  className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-xl px-2.5 h-9"
                  title="Sesuaikan tanggal khusus untuk hari perencanaan ini"
                >
                  <span className="text-[10px] font-bold text-slate-500 uppercase whitespace-nowrap">
                    Tgl:
                  </span>
                  <input
                    type="date"
                    value={rawDateStr}
                    onChange={(e) => {
                      const nextDate = e.target.value;
                      updateDayData({
                        ...currentDayData,
                        tanggal: nextDate || undefined,
                      });
                    }}
                    className="bg-transparent border-0 text-xs font-bold text-slate-800 focus:outline-none focus:ring-0 p-0 cursor-pointer"
                  />
                </div>
              </div>
            </div>

            {/* Strip Cepat Tanggal Perencanaan */}
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1">
              {Array.from({ length: totalDays }).map((_, i) => {
                const dayNum = i + 1;
                const isCurrent = selectedDay === dayNum;
                const dayOverrideDate = foodCostDays.find(
                  (d) => d.hariKe === dayNum && d.jenisMenu === activeDbType
                )?.tanggal;
                const chipDateRaw = dayOverrideDate || scheduleDates[i] || "";
                const chipDateLabel = chipDateRaw
                  ? formatScheduleIndonesianDate(chipDateRaw)
                  : `Hari ${dayNum}`;
                return (
                  <button
                    key={dayNum}
                    type="button"
                    onClick={() => setSelectedDay(dayNum)}
                    className={`px-2.5 py-1.5 rounded-xl border text-[11px] font-bold whitespace-nowrap transition cursor-pointer shrink-0 ${
                      isCurrent
                        ? "bg-indigo-600 text-white border-indigo-600 shadow-xs"
                        : "bg-slate-50 hover:bg-indigo-50 text-slate-700 border-slate-200 hover:border-indigo-300"
                    }`}
                  >
                    {chipDateRaw ? `${chipDateLabel} · H-${dayNum}` : `Hari Ke-${dayNum}`}
                  </button>
                );
              })}
            </div>
          </div>

          <div id="fc-tables-container" className="space-y-6">
            {/* Persistent Clipboard Banner for Quick Copy & Paste */}
            {copiedFoodCost && (
              <div className="bg-amber-50/90 border border-amber-300 rounded-2xl p-3.5 sm:px-5 flex flex-col md:flex-row md:items-center justify-between gap-3 shadow-xs">
                <div className="flex items-start sm:items-center gap-3">
                  <div className="p-2 rounded-xl bg-amber-100 text-amber-800 shrink-0">
                    <Clipboard className="w-4 h-4" />
                  </div>
                  <div>
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-xs font-black text-amber-950 uppercase">
                        Clipboard Food Cost Aktif: {copiedFoodCost.sourceLabel}
                      </span>
                      <span className="px-2 py-0.5 rounded bg-white border border-amber-200 text-[10px] font-mono font-bold text-amber-800">
                        {copiedFoodCost.bahanList.length} Bahan · Buffer {copiedFoodCost.bufferPct}% · {copiedFoodCost.copiedAt}
                      </span>
                    </div>
                    <p className="text-[11px] text-amber-800 font-medium mt-0.5">
                      Klik tombol <strong>Tempel Isian</strong> pada Food Cost kategori atau hari berikutnya, lalu cukup ubah kolom <strong>BERAT (BB)</strong> &amp; tekan Enter untuk lanjut baris berikutnya.
                    </p>
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2 shrink-0">
                  {isPerKategoriMode && (
                    <button
                      type="button"
                      onClick={handlePasteToAllEmptyCategoryEntries}
                      className="px-3.5 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-xl text-xs font-black flex items-center gap-1.5 shadow-xs transition cursor-pointer"
                    >
                      <ClipboardCheck className="w-3.5 h-3.5" />
                      Tempel ke Semua Tabel Kosong
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={() => setCopiedFoodCost(null)}
                    className="px-3 py-2 bg-white hover:bg-amber-100 border border-amber-300 text-amber-900 rounded-xl text-xs font-bold transition cursor-pointer"
                  >
                    Hapus Clipboard
                  </button>
                </div>
              </div>
            )}

            {!appliedConfig.confirmed ? (
              <div className="bg-gradient-to-br from-slate-50 to-indigo-50/40 border-2 border-dashed border-indigo-200 rounded-2xl p-8 text-center space-y-3">
                <div className="w-11 h-11 rounded-2xl bg-indigo-100 text-indigo-700 flex items-center justify-center mx-auto">
                  <Layers className="w-5 h-5" />
                </div>
                <div className="space-y-1 max-w-xl mx-auto">
                  <h4 className="text-sm font-black text-slate-900 uppercase tracking-tight">
                    Konfirmasi Kategori Food Cost Terlebih Dahulu
                  </h4>
                  <p className="text-xs text-slate-600 leading-relaxed">
                    Pilih mode tampilan beserta kategori sasaran pada panel di atas, lalu klik{" "}
                    <strong>Konfirmasi &amp; Tampilkan Food Cost Otomatis</strong>.
                  </p>
                </div>
                <div className="pt-1 flex flex-wrap items-center justify-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setIsSelectorExpanded(true);
                      handleOpenCategoryConfirmModal();
                    }}
                    className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-black shadow-xs transition cursor-pointer inline-flex items-center gap-1.5"
                  >
                    <CheckCircle2 className="w-4 h-4" />
                    Konfirmasi Pilihan Saat Ini
                  </button>
                </div>
              </div>
            ) : isPerKategoriMode ? (
              resolvedCategoryEntries.length === 0 ? (
                <div className="bg-slate-50 border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center space-y-3">
                  <p className="text-xs font-bold text-slate-600">
                    Tidak ada kategori Food Cost yang cocok dengan filter porsi saat ini.
                  </p>
                  <button
                    type="button"
                    onClick={() => {
                      handleSelectDraftScope("all");
                      setAppliedConfig((prev) => ({ ...prev, scope: "all", confirmed: true }));
                      handleResetToSingleCategories(false);
                    }}
                    className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold cursor-pointer"
                  >
                    Tampilkan Semua Kategori Kelompok
                  </button>
                </div>
              ) : (
                resolvedCategoryEntries.map((entry) => {
                  const calcInfo = categoryCalcResultsMap[entry.id];
                  const entryPm = calcInfo?.pmCount ?? getEntryEffectivePmCount(entry);
                  const entryTargetPrice =
                    calcInfo?.targetPrice ?? (entry.porsi === "besar" ? activeTargetPriceBesar : activeTargetPriceKecil);
                  const entryActualPerPorsi = calcInfo?.calculatedCostPerPorsi ?? 0;
                  const entryActualTotal = calcInfo?.actualTotalCost ?? 0;
                  const isEntryOver = entryPm > 0 && entryActualPerPorsi > entryTargetPrice;
                  const highlightInfo = pastedHighlightMap[`cat_${entry.id}`];
                  const isMergedEntry = entry.sasaranIds.length > 1;
                  const quickSources = copySourceOptions.filter((opt) => opt.key !== `cat:${entry.id}`);

                  return (
                    <div
                      key={entry.id}
                      id={`fc-block-cat_${entry.id}`}
                      className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden space-y-3 p-4"
                    >
                      <div className="border-b border-slate-200 pb-3">
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
                          docKey="foodcost"
                        />
                      </div>

                      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3 pb-2 border-b border-slate-100">
                        <div className="space-y-1">
                          <div className="flex flex-wrap items-center gap-2">
                            <span
                              className={`w-2.5 h-2.5 rounded-full shrink-0 ${
                                entry.sasaranIds.some((id) => isAlergiSasaranId(id))
                                  ? "bg-orange-500"
                                  : entry.porsi === "besar"
                                    ? "bg-amber-500"
                                    : "bg-cyan-500"
                              }`}
                            ></span>
                            <h4 className="font-extrabold text-slate-900 uppercase text-xs tracking-wide">
                              FOOD COST — {entry.title}
                            </h4>
                            <span
                              className={`px-2 py-0.5 rounded-md text-[10px] font-black uppercase border ${
                                entry.porsi === "besar"
                                  ? "bg-amber-50 text-amber-800 border-amber-200"
                                  : "bg-cyan-50 text-cyan-800 border-cyan-200"
                              }`}
                            >
                              {entry.porsi === "besar" ? "Porsi Besar" : "Porsi Kecil"}
                            </span>
                            {entry.sasaranIds.some((id) => isAlergiSasaranId(id)) && (
                              <span className="px-2 py-0.5 rounded-md bg-orange-100 text-orange-800 border border-orange-200 text-[10px] font-black uppercase">
                                Diet Alergi
                              </span>
                            )}
                            {isMergedEntry && (
                              <>
                                <span className="px-2 py-0.5 rounded-md bg-indigo-50 text-indigo-800 border border-indigo-200 text-[10px] font-black">
                                  Gabungan {entry.sasaranIds.length} Kategori
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleSplitMergedGroup(entry.id)}
                                  className="px-2 py-0.5 rounded-md bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-300 text-[10px] font-extrabold inline-flex items-center gap-1 cursor-pointer transition"
                                  title="Pisahkan kembali kategori gabungan ini"
                                >
                                  <Split className="w-3 h-3" />
                                  Pisahkan
                                </button>
                              </>
                            )}
                            {isEntryOver && (
                              <span className="px-2 py-0.5 rounded-md bg-rose-100 text-rose-800 border border-rose-300 text-[10px] font-black">
                                OVER BUDGET
                              </span>
                            )}
                          </div>

                          <p className="text-[11px] text-slate-500 font-semibold">
                            Plafon PM:{" "}
                            <span className="text-slate-800 font-bold">{formatRupiah(entryTargetPrice)}/porsi</span>
                            {" · "}
                            Aktual:{" "}
                            <span className={isEntryOver ? "text-rose-600 font-black" : "text-emerald-700 font-bold"}>
                              {formatRupiah(Math.round(entryActualPerPorsi))}/porsi
                            </span>
                            {" · "}
                            Total Biaya:{" "}
                            <span className="text-slate-800 font-bold">{formatRupiah(Math.round(entryActualTotal))}</span>
                            {" · "}
                            Jumlah: <span className="text-slate-900 font-black">{entryPm} PM</span>
                          </p>

                          {isMergedEntry && (
                            <div className="flex flex-wrap items-center gap-1 pt-0.5">
                              {entry.sasaranIds.map((sId) => (
                                <span
                                  key={sId}
                                  className="px-1.5 py-0.5 rounded bg-slate-100 border border-slate-200 text-[10px] font-bold text-slate-700"
                                >
                                  {getSasaranLabel(sId)} ({getSasaranPmCount(sId, entry.porsi)})
                                </span>
                              ))}
                            </div>
                          )}
                        </div>

                        <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                          <button
                            type="button"
                            onClick={() =>
                              handleCopyFoodCost(entry.title, entry.porsi, entry.bahanList, entry.bufferPct)
                            }
                            className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-lg text-[11px] font-extrabold transition cursor-pointer"
                            title="Salin semua bahan, uraian, harga, dan buffer dari Food Cost ini"
                          >
                            <Copy className="w-3.5 h-3.5 text-indigo-600" />
                            Salin Semua Isian
                          </button>

                          <button
                            type="button"
                            onClick={() =>
                              handlePasteToTarget(
                                "category",
                                entry.id,
                                entry.title,
                                (entry.bahanList || []).length
                              )
                            }
                            className={`flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-[11px] font-extrabold transition cursor-pointer ${
                              copiedFoodCost
                                ? "bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300"
                                : "bg-slate-50 hover:bg-slate-100 text-slate-500 border border-slate-200"
                            }`}
                            title="Tempel semua isian Food Cost yang sudah disalin"
                          >
                            <ClipboardCheck className="w-3.5 h-3.5" />
                            Tempel Isian
                          </button>

                          {quickSources.length > 0 && (
                            <select
                              value=""
                              onChange={(e) => {
                                const picked = quickSources.find((o) => o.key === e.target.value);
                                if (!picked) return;
                                handlePasteToTarget(
                                  "category",
                                  entry.id,
                                  entry.title,
                                  (entry.bahanList || []).length,
                                  {
                                    sourceLabel: picked.label,
                                    sourcePorsi: picked.porsi,
                                    copiedAt: "Langsung",
                                    bufferPct: picked.bufferPct,
                                    bahanList: picked.bahanList,
                                  }
                                );
                              }}
                              className="bg-white border border-slate-200 hover:border-indigo-300 rounded-lg text-[11px] font-bold text-slate-700 px-2 py-1.5 max-w-[165px] cursor-pointer focus:outline-none"
                              title="Tarik langsung isian dari Food Cost kategori atau hari lain"
                            >
                              <option value="">Salin dari...</option>
                              {quickSources.map((opt) => (
                                <option key={opt.key} value={opt.key}>
                                  {opt.label}
                                </option>
                              ))}
                            </select>
                          )}

                          <button
                            type="button"
                            onClick={() => addIngredientToCategoryEntry(entry.id)}
                            className="flex items-center gap-1 px-2.5 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-lg text-[11px] font-bold transition shadow-xs cursor-pointer"
                          >
                            <Plus className="w-3.5 h-3.5" />
                            Tambah Bahan Baku
                          </button>

                          <button
                            type="button"
                            onClick={() => {
                              triggerConfirm({
                                title: `Hapus Semua Bahan (${entry.title})`,
                                message: `Apakah Anda yakin ingin menghapus seluruh komponen bahan baku pada kategori "${entry.title}"?`,
                                onConfirm: () => {
                                  updateCategoryEntry(entry.id, (cur) => ({ ...cur, bahanList: [] }));
                                },
                              });
                            }}
                            className="flex items-center gap-1 px-2.5 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg text-[11px] font-bold transition shadow-xs cursor-pointer"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            Hapus Semua
                          </button>
                        </div>
                      </div>

                      {highlightInfo && (
                        <div className="bg-amber-50 border border-amber-300 rounded-xl px-3.5 py-2 flex items-center justify-between gap-2 text-[11px] text-amber-900">
                          <span>
                            <strong>Isian ditempel dari &ldquo;{highlightInfo.from}&rdquo;.</strong> Kolom{" "}
                            <strong>BERAT (BB)</strong> disorot — silakan ubah angka BB sesuai standar porsi kategori ini &amp; tekan{" "}
                            <kbd className="px-1.5 py-0.5 bg-white border border-amber-300 rounded text-[10px] font-mono">
                              Enter
                            </kbd>{" "}
                            untuk lanjut ke baris berikutnya.
                          </span>
                          <button
                            type="button"
                            onClick={() =>
                              setPastedHighlightMap((prev) => {
                                const next = { ...prev };
                                delete next[`cat_${entry.id}`];
                                return next;
                              })
                            }
                            className="px-2 py-0.5 rounded bg-white border border-amber-300 text-[10px] font-bold hover:bg-amber-100 shrink-0 cursor-pointer"
                          >
                            Selesai
                          </button>
                        </div>
                      )}

                      {renderCategoryFoodCostTable(entry, false)}

                      <div className="bg-slate-50 p-3 rounded-xl border border-slate-200/60 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="font-bold text-slate-700 flex items-center gap-1.5">
                            <Sliders className="w-3.5 h-3.5 text-indigo-500" />
                            Rujukan AKG {entry.title}:
                          </span>
                          <select
                            value={entry.akgType || getDefaultAkgTypeForGroup(entry.sasaranIds, entry.porsi)}
                            onChange={(e) =>
                              updateCategoryEntry(entry.id, (cur) => ({ ...cur, akgType: e.target.value }))
                            }
                            className="bg-white border border-slate-200 rounded-lg text-xs px-2 py-1 text-slate-700 font-bold focus:outline-none"
                          >
                            {Object.entries(TARGET_AKG_LIMITS).map(([key, val]) => (
                              <option key={key} value={key}>
                                {val.label}
                              </option>
                            ))}
                            <option value="custom">✍️ Target Kustom (Input Manual)</option>
                          </select>
                        </div>

                        <div className="flex flex-wrap items-center gap-1.5">
                          <span className="text-[11px] font-bold text-slate-600">Buffer Kategori (%):</span>
                          {[2.5, 3, 5].map((pctVal) => (
                            <button
                              key={pctVal}
                              type="button"
                              onClick={() =>
                                updateCategoryEntry(entry.id, (cur) => ({ ...cur, bufferPct: pctVal }))
                              }
                              className={`px-2 py-1 rounded-lg text-[10px] font-mono font-bold border transition cursor-pointer ${
                                (entry.bufferPct ?? 5) === pctVal
                                  ? "bg-indigo-600 text-white border-indigo-600"
                                  : "bg-white text-slate-700 border-slate-200 hover:bg-slate-100"
                              }`}
                            >
                              {pctVal}%
                            </button>
                          ))}
                          <input
                            type="number"
                            step="0.5"
                            min="0"
                            max="50"
                            value={entry.bufferPct ?? 5}
                            onChange={(e) =>
                              updateCategoryEntry(entry.id, (cur) => ({
                                ...cur,
                                bufferPct: Math.max(0, Number(e.target.value) || 0),
                              }))
                            }
                            className="w-16 bg-white border border-slate-200 rounded-lg text-xs px-2 py-1 text-center font-mono font-bold text-slate-800"
                          />
                        </div>
                      </div>

                      {profile && (
                        <div className="border-t border-slate-200 pt-4 mt-2">
                          <AutoSignatureBlock
                            profile={profile}
                            docKey="foodcost"
                            dateText={currentCalendarDate || undefined}
                          />
                        </div>
                      )}
                    </div>
                  );
                })
              )
            ) : (
              <>
                {/* Table 1: Porsi Besar Block (Shown when scope is all or besar in gabung_porsi mode) */}
                {showBesarTable && (
                  <div id="fc-block-besar" className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden space-y-3 p-4">
                    <div className="border-b border-slate-200 pb-3">
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
                        docKey="foodcost"
                      />
                    </div>
                    <div className="flex items-center justify-between flex-wrap gap-2">
                      <div>
                        <h4 className="font-extrabold text-slate-800 flex items-center gap-2 uppercase text-xs">
                          <span className="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                          KOMPONEN BAHAN BAKU - PORSI BESAR
                          {isOverBudgetBesar && (
                            <span className="px-2 py-0.5 rounded-md bg-rose-100 text-rose-800 border border-rose-300 text-[10px] font-black">
                              OVER BUDGET
                            </span>
                          )}
                        </h4>
                        <p className="text-[11px] text-slate-500 mt-0.5 font-semibold">
                          Plafon PM: <span className="text-slate-800 font-bold">{formatRupiah(targetPriceBesar)}/porsi</span>
                          {" · "}
                          Aktual: <span className={isOverBudgetBesar ? "text-rose-600 font-black" : "text-emerald-700 font-bold"}>{formatRupiah(Math.round(activeCostPerPorsiBesar))}/porsi</span>
                          {" · "}
                          Jumlah: <span className="text-slate-800 font-bold">{pmBesar} PM</span>
                        </p>
                      </div>
                      <div className="flex flex-wrap items-center gap-1.5">
                        <button
                          type="button"
                          onClick={() =>
                            handleCopyFoodCost(
                              `Porsi Besar Hari ${selectedDay}`,
                              "besar",
                              currentDayData.porsiBesarBahan || [],
                              currentDayData.bufferPct
                            )
                          }
                          className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-lg text-xs font-extrabold transition cursor-pointer"
                        >
                          <Copy className="w-3.5 h-3.5 text-indigo-600" />
                          Salin Semua Isian
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            handlePasteToTarget(
                              "porsi",
                              "besar",
                              `Porsi Besar Hari ${selectedDay}`,
                              (currentDayData.porsiBesarBahan || []).length
                            )
                          }
                          className={`flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-extrabold transition cursor-pointer ${
                            copiedFoodCost
                              ? "bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300"
                              : "bg-slate-50 hover:bg-slate-100 text-slate-500 border border-slate-200"
                          }`}
                        >
                          <ClipboardCheck className="w-3.5 h-3.5" />
                          Tempel Isian
                        </button>
                        {copySourceOptions.filter((o) => o.key !== "porsi:besar").length > 0 && (
                          <select
                            value=""
                            onChange={(e) => {
                              const picked = copySourceOptions.find((o) => o.key === e.target.value);
                              if (!picked) return;
                              handlePasteToTarget(
                                "porsi",
                                "besar",
                                `Porsi Besar Hari ${selectedDay}`,
                                (currentDayData.porsiBesarBahan || []).length,
                                {
                                  sourceLabel: picked.label,
                                  sourcePorsi: picked.porsi,
                                  copiedAt: "Langsung",
                                  bufferPct: picked.bufferPct,
                                  bahanList: picked.bahanList,
                                }
                              );
                            }}
                            className="bg-white border border-slate-200 rounded-lg text-xs font-bold text-slate-700 px-2 py-1.5 max-w-[160px] cursor-pointer focus:outline-none"
                          >
                            <option value="">Salin dari...</option>
                            {copySourceOptions
                              .filter((o) => o.key !== "porsi:besar")
                              .map((opt) => (
                                <option key={opt.key} value={opt.key}>
                                  {opt.label}
                                </option>
                              ))}
                          </select>
                        )}
                        <button
                          id="btn-add-ing-besar"
                          type="button"
                          onClick={() => addIngredientRow("besar")}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-600 rounded-lg text-xs font-bold transition shadow-xs"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          Tambah Bahan Baku
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            triggerConfirm({
                              title: "Hapus Semua Bahan Porsi Besar",
                              message: "Apakah Anda yakin ingin menghapus seluruh komponen bahan baku porsi besar hari ini?",
                              onConfirm: () => {
                                updateDayData({
                                  ...currentDayData,
                                  porsiBesarBahan: []
                                });
                              }
                            });
                          }}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg text-xs font-bold transition shadow-xs"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          Hapus Semua
                        </button>
                      </div>
                    </div>

                    {pastedHighlightMap["besar"] && (
                      <div className="bg-amber-50 border border-amber-300 rounded-xl px-3.5 py-2 flex items-center justify-between gap-2 text-[11px] text-amber-900">
                        <span>
                          <strong>Isian ditempel dari &ldquo;{pastedHighlightMap["besar"].from}&rdquo;.</strong> Kolom{" "}
                          <strong>BERAT (BB)</strong> disorot — ubah angka BB &amp; tekan Enter untuk lanjut baris berikutnya.
                        </span>
                        <button
                          type="button"
                          onClick={() =>
                            setPastedHighlightMap((prev) => {
                              const next = { ...prev };
                              delete next["besar"];
                              return next;
                            })
                          }
                          className="px-2 py-0.5 rounded bg-white border border-amber-300 text-[10px] font-bold hover:bg-amber-100 shrink-0 cursor-pointer"
                        >
                          Selesai
                        </button>
                      </div>
                    )}

                    {/* Tanda Kategori Terpilih di atas Tabel Porsi Besar */}
                    {renderCategorySignBanner("besar", false)}

                    {renderFoodCostTable("besar", false)}

                    <div className="bg-slate-50 p-3 rounded-xl border border-slate-200/60 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                      <span className="font-bold text-slate-700 flex items-center gap-1.5">
                        <Sliders className="w-3.5 h-3.5 text-indigo-500" />
                        Rujukan AKG Porsi Besar Aktif:
                      </span>
                      <select
                        id="select-akg-besar"
                        value={besarAkgType}
                        onChange={(e) => setBesarAkgType(e.target.value)}
                        className="bg-white border border-slate-200 rounded-lg text-xs px-2 py-1 text-slate-700 font-bold focus:outline-none"
                      >
                        {Object.entries(TARGET_AKG_LIMITS).map(([key, val]) => (
                          <option key={key} value={key}>
                            {val.label}
                          </option>
                        ))}
                        <option value="custom">✍️ Target Kustom (Input Manual)</option>
                      </select>
                    </div>

                    {profile && (
                      <div className="border-t border-slate-200 pt-4 mt-2">
                        <AutoSignatureBlock
                          profile={profile}
                          docKey="foodcost"
                          dateText={currentCalendarDate || undefined}
                        />
                      </div>
                    )}
                  </div>
                )}

                {/* Table 2: Porsi Kecil Block (Shown when scope is all or kecil in gabung_porsi mode) */}
                {showKecilTable && (
                  <div id="fc-block-kecil" className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden space-y-3 p-4">
                    <div className="border-b border-slate-200 pb-3">
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
                        docKey="foodcost"
                      />
                    </div>
                    <div className="flex items-center justify-between flex-wrap gap-2">
                      <div>
                        <h4 className="font-extrabold text-slate-800 flex items-center gap-2 uppercase text-xs">
                          <span className="w-2.5 h-2.5 rounded-full bg-cyan-500"></span>
                          KOMPONEN BAHAN BAKU - PORSI KECIL
                          {isOverBudgetKecil && (
                            <span className="px-2 py-0.5 rounded-md bg-rose-100 text-rose-800 border border-rose-300 text-[10px] font-black">
                              OVER BUDGET
                            </span>
                          )}
                        </h4>
                        <p className="text-[11px] text-slate-500 mt-0.5 font-semibold">
                          Plafon PM: <span className="text-slate-800 font-bold">{formatRupiah(targetPriceKecil)}/porsi</span>
                          {" · "}
                          Aktual: <span className={isOverBudgetKecil ? "text-rose-600 font-black" : "text-emerald-700 font-bold"}>{formatRupiah(Math.round(activeCostPerPorsiKecil))}/porsi</span>
                          {" · "}
                          Jumlah: <span className="text-slate-800 font-bold">{pmKecil} PM</span>
                        </p>
                      </div>
                      <div className="flex flex-wrap items-center gap-1.5">
                        <button
                          type="button"
                          onClick={() =>
                            handleCopyFoodCost(
                              `Porsi Kecil Hari ${selectedDay}`,
                              "kecil",
                              currentDayData.porsiKecilBahan || [],
                              currentDayData.bufferPct
                            )
                          }
                          className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-lg text-xs font-extrabold transition cursor-pointer"
                        >
                          <Copy className="w-3.5 h-3.5 text-indigo-600" />
                          Salin Semua Isian
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            handlePasteToTarget(
                              "porsi",
                              "kecil",
                              `Porsi Kecil Hari ${selectedDay}`,
                              (currentDayData.porsiKecilBahan || []).length
                            )
                          }
                          className={`flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-extrabold transition cursor-pointer ${
                            copiedFoodCost
                              ? "bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300"
                              : "bg-slate-50 hover:bg-slate-100 text-slate-500 border border-slate-200"
                          }`}
                        >
                          <ClipboardCheck className="w-3.5 h-3.5" />
                          Tempel Isian
                        </button>
                        {copySourceOptions.filter((o) => o.key !== "porsi:kecil").length > 0 && (
                          <select
                            value=""
                            onChange={(e) => {
                              const picked = copySourceOptions.find((o) => o.key === e.target.value);
                              if (!picked) return;
                              handlePasteToTarget(
                                "porsi",
                                "kecil",
                                `Porsi Kecil Hari ${selectedDay}`,
                                (currentDayData.porsiKecilBahan || []).length,
                                {
                                  sourceLabel: picked.label,
                                  sourcePorsi: picked.porsi,
                                  copiedAt: "Langsung",
                                  bufferPct: picked.bufferPct,
                                  bahanList: picked.bahanList,
                                }
                              );
                            }}
                            className="bg-white border border-slate-200 rounded-lg text-xs font-bold text-slate-700 px-2 py-1.5 max-w-[160px] cursor-pointer focus:outline-none"
                          >
                            <option value="">Salin dari...</option>
                            {copySourceOptions
                              .filter((o) => o.key !== "porsi:kecil")
                              .map((opt) => (
                                <option key={opt.key} value={opt.key}>
                                  {opt.label}
                                </option>
                              ))}
                          </select>
                        )}
                        <button
                          id="btn-add-ing-kecil"
                          type="button"
                          onClick={() => addIngredientRow("kecil")}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-600 rounded-lg text-xs font-bold transition shadow-xs"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          Tambah Bahan Baku
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            triggerConfirm({
                              title: "Hapus Semua Bahan Porsi Kecil",
                              message: "Apakah Anda yakin ingin menghapus seluruh komponen bahan baku porsi kecil hari ini?",
                              onConfirm: () => {
                                updateDayData({
                                  ...currentDayData,
                                  porsiKecilBahan: []
                                });
                              }
                            });
                          }}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg text-xs font-bold transition shadow-xs"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          Hapus Semua
                        </button>
                      </div>
                    </div>

                    {pastedHighlightMap["kecil"] && (
                      <div className="bg-amber-50 border border-amber-300 rounded-xl px-3.5 py-2 flex items-center justify-between gap-2 text-[11px] text-amber-900">
                        <span>
                          <strong>Isian ditempel dari &ldquo;{pastedHighlightMap["kecil"].from}&rdquo;.</strong> Kolom{" "}
                          <strong>BERAT (BB)</strong> disorot — ubah angka BB &amp; tekan Enter untuk lanjut baris berikutnya.
                        </span>
                        <button
                          type="button"
                          onClick={() =>
                            setPastedHighlightMap((prev) => {
                              const next = { ...prev };
                              delete next["kecil"];
                              return next;
                            })
                          }
                          className="px-2 py-0.5 rounded bg-white border border-amber-300 text-[10px] font-bold hover:bg-amber-100 shrink-0 cursor-pointer"
                        >
                          Selesai
                        </button>
                      </div>
                    )}

                    {/* Tanda Kategori Terpilih di atas Tabel Porsi Kecil */}
                    {renderCategorySignBanner("kecil", false)}

                    {renderFoodCostTable("kecil", false)}

                    <div className="bg-slate-50 p-3 rounded-xl border border-slate-200/60 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                      <span className="font-bold text-slate-700 flex items-center gap-1.5">
                        <Sliders className="w-3.5 h-3.5 text-indigo-500" />
                        Rujukan AKG Porsi Kecil Aktif:
                      </span>
                      <select
                        id="select-akg-kecil"
                        value={kecilAkgType}
                        onChange={(e) => setKecilAkgType(e.target.value)}
                        className="bg-white border border-slate-200 rounded-lg text-xs px-2 py-1 text-slate-700 font-bold focus:outline-none"
                      >
                        {Object.entries(TARGET_AKG_LIMITS).map(([key, val]) => (
                          <option key={key} value={key}>
                            {val.label}
                          </option>
                        ))}
                        <option value="custom">✍️ Target Kustom (Input Manual)</option>
                      </select>
                    </div>

                    {profile && (
                      <div className="border-t border-slate-200 pt-4 mt-2">
                        <AutoSignatureBlock
                          profile={profile}
                          docKey="foodcost"
                          dateText={currentCalendarDate || undefined}
                        />
                      </div>
                    )}
                  </div>
                )}
              </>
            )}
          </div>

          {/* Existing Custom Tables (only rendered if legacy custom tables exist) */}
          {customTables.length > 0 && (
          <div className="space-y-6">
            {/* Custom Tables List */}
            {customTables.map((table) => {
              return (
                <div key={table.id} id={`fc-block-custom_${table.id}`} className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden space-y-3 p-4">
                  <div className="border-b border-slate-200 pb-3">
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
                      docKey="foodcost"
                    />
                  </div>
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-slate-100">
                    <div className="space-y-1 flex-grow">
                      <div className="flex items-center gap-2">
                        <span className={`w-2.5 h-2.5 rounded-full ${table.porsi === "besar" ? "bg-amber-500" : "bg-cyan-500"}`}></span>
                        <input
                          type="text"
                          value={table.namaTabel}
                          onChange={(e) => editCustomTableMeta(table.id, "namaTabel", e.target.value)}
                          className="font-extrabold text-slate-800 text-xs bg-slate-50 hover:bg-slate-100 border border-transparent hover:border-slate-200 focus:bg-white focus:border-indigo-500 rounded px-2 py-1 focus:outline-none"
                        />
                        <span className="text-[9px] font-black uppercase text-purple-600 bg-purple-50 px-1.5 py-0.5 rounded">Tabel Baru</span>
                      </div>
                      <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-[11px] text-slate-400 font-semibold pl-4">
                        <div className="flex items-center gap-1">
                          Plafon: 
                          <input
                            type="number"
                            value={table.plafon || ""}
                            onChange={(e) => editCustomTableMeta(table.id, "plafon", Number(e.target.value))}
                            className="w-16 font-mono text-xs font-bold text-slate-700 bg-slate-50 border border-slate-100 rounded px-1 text-right"
                          />
                          / porsi
                        </div>
                        <div className="flex items-center gap-1">
                          Jumlah PM: 
                          <input
                            type="number"
                            value={table.pmCount || ""}
                            onChange={(e) => editCustomTableMeta(table.id, "pmCount", Number(e.target.value))}
                            className="w-12 font-mono text-xs font-bold text-slate-700 bg-slate-50 border border-slate-100 rounded px-1 text-center"
                          />
                        </div>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-1.5 self-end sm:self-auto">
                      <button
                        type="button"
                        onClick={() =>
                          handleCopyFoodCost(table.namaTabel, table.porsi, table.bahanList || [], table.bufferPct)
                        }
                        className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-lg text-xs font-extrabold transition cursor-pointer"
                      >
                        <Copy className="w-3.5 h-3.5 text-indigo-600" />
                        Salin Semua Isian
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          handlePasteToTarget(
                            "custom",
                            table.id,
                            table.namaTabel,
                            (table.bahanList || []).length
                          )
                        }
                        className={`flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-extrabold transition cursor-pointer ${
                          copiedFoodCost
                            ? "bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300"
                            : "bg-slate-50 hover:bg-slate-100 text-slate-500 border border-slate-200"
                        }`}
                      >
                        <ClipboardCheck className="w-3.5 h-3.5" />
                        Tempel Isian
                      </button>
                      {copySourceOptions.length > 0 && (
                        <select
                          value=""
                          onChange={(e) => {
                            const picked = copySourceOptions.find((o) => o.key === e.target.value);
                            if (!picked) return;
                            handlePasteToTarget(
                              "custom",
                              table.id,
                              table.namaTabel,
                              (table.bahanList || []).length,
                              {
                                sourceLabel: picked.label,
                                sourcePorsi: picked.porsi,
                                copiedAt: "Langsung",
                                bufferPct: picked.bufferPct,
                                bahanList: picked.bahanList,
                              }
                            );
                          }}
                          className="bg-white border border-slate-200 rounded-lg text-xs font-bold text-slate-700 px-2 py-1.5 max-w-[150px] cursor-pointer focus:outline-none"
                        >
                          <option value="">Salin dari...</option>
                          {copySourceOptions.map((opt) => (
                            <option key={opt.key} value={opt.key}>
                              {opt.label}
                            </option>
                          ))}
                        </select>
                      )}
                      <button
                        type="button"
                        onClick={() => addIngredientToCustomTable(table.id)}
                        className="flex items-center gap-1 px-2.5 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-600 rounded-lg text-xs font-bold transition shadow-xs"
                      >
                        <Plus className="w-3.5 h-3.5" />
                        Tambah Bahan Baku
                      </button>
                      <button
                        type="button"
                        onClick={() => deleteCustomTable(table.id)}
                        className="flex items-center gap-1 px-2.5 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg text-xs font-bold transition shadow-xs"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                        Hapus Tabel
                      </button>
                    </div>
                  </div>

                  {pastedHighlightMap[`custom_${table.id}`] && (
                    <div className="bg-amber-50 border border-amber-300 rounded-xl px-3.5 py-2 flex items-center justify-between gap-2 text-[11px] text-amber-900">
                      <span>
                        <strong>Isian ditempel dari &ldquo;{pastedHighlightMap[`custom_${table.id}`].from}&rdquo;.</strong> Kolom{" "}
                        <strong>BERAT (BB)</strong> disorot — ubah angka BB &amp; tekan Enter untuk lanjut baris berikutnya.
                      </span>
                      <button
                        type="button"
                        onClick={() =>
                          setPastedHighlightMap((prev) => {
                            const next = { ...prev };
                            delete next[`custom_${table.id}`];
                            return next;
                          })
                        }
                        className="px-2 py-0.5 rounded bg-white border border-amber-300 text-[10px] font-bold hover:bg-amber-100 shrink-0 cursor-pointer"
                      >
                        Selesai
                      </button>
                    </div>
                  )}

                  {renderCustomFoodCostTable(table, false)}

                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-200/60 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                    <span className="font-bold text-slate-700 flex items-center gap-1.5">
                      <Sliders className="w-3.5 h-3.5 text-indigo-500" />
                      Rujukan AKG Aktif:
                    </span>
                    <select
                      value={table.akgType || (table.porsi === "besar" ? "sd_besar" : "sd_kecil")}
                      onChange={(e) => editCustomTableMeta(table.id, "akgType", e.target.value)}
                      className="bg-white border border-slate-200 rounded-lg text-xs px-2 py-1 text-slate-700 font-bold focus:outline-none"
                    >
                      {Object.entries(TARGET_AKG_LIMITS).map(([key, val]) => (
                        <option key={key} value={key}>
                          {val.label}
                        </option>
                      ))}
                      <option value="custom">✍️ Target Kustom (Input Manual)</option>
                    </select>
                  </div>

                  {profile && (
                    <div className="border-t border-slate-200 pt-4 mt-2">
                      <AutoSignatureBlock
                        profile={profile}
                        docKey="foodcost"
                        dateText={currentCalendarDate || undefined}
                      />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
          )}
        </div>
      )}

      {/* --- VIEW MODE: PRATINJAU CETAK A4 & EXPORT --- */}
      <div className={viewMode === "print" ? "space-y-6" : "hidden print:block space-y-6"}>
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4 no-print">
            <div className="flex flex-col sm:flex-row justify-between sm:items-center pb-2 border-b border-slate-100 gap-3">
              <h4 className="text-xs font-black text-slate-900 uppercase tracking-wider">Sesuaikan KOP Surat & Cetak / Ekspor</h4>
              <div className="flex items-center gap-2 self-start sm:self-auto">
                <button
                  id="btn-preview-fc-modal"
                  type="button"
                  onClick={() => setShowPrintModal(true)}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition cursor-pointer"
                >
                  <Eye className="w-4 h-4" />
                  Pratinjau Cetak
                </button>
                <button
                  id="btn-fc-export-pdf"
                  type="button"
                  disabled={!!isDownloading}
                  onClick={() => exportToPDF("print-area-food-cost", `Food_Cost_Hari_${selectedDay}_${selectedType}`, 100, setIsDownloading, { paperSize, orientation: "auto" })}
                  className="bg-red-600 hover:bg-red-700 disabled:opacity-50 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition"
                >
                  <Download className="w-4 h-4" />
                  {isDownloading === "Memproses dokumen PDF..." || isDownloading === "Membuat file PDF..." ? isDownloading : "Ekspor PDF"}
                </button>
                <button
                  id="btn-fc-export-excel"
                  type="button"
                  onClick={handleExportExcel}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition"
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  Unduh Excel (.XLSX)
                </button>
                <button
                  id="btn-fc-download-img"
                  type="button"
                  disabled={!!isDownloading}
                  onClick={() => downloadElementAsImage("print-area-food-cost", `Food_Cost_Hari_${selectedDay}_${selectedType}`, setIsDownloading)}
                  className="bg-amber-600 hover:bg-amber-700 disabled:opacity-50 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition"
                >
                  <ImageIcon className="w-4 h-4" />
                  {isDownloading === "Memproses gambar..." || isDownloading === "Mengunduh gambar..." ? isDownloading : "Unduh Gambar (PNG)"}
                </button>
                <button
                  id="btn-fc-print"
                  type="button"
                  onClick={() => printElementSmart("print-area-food-cost", { title: `Food Cost ${selectedType} — Hari Ke-${selectedDay}`, paperSize, orientation: "auto" })}
                  className="bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold shadow-xs flex items-center gap-1.5 transition"
                >
                  <Printer className="w-4 h-4" />
                  Cetak / Simpan PDF ({paperSize})
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
                <label className="text-[10px] font-bold text-slate-400 uppercase">KOP Baris 3 (Nama SPPG)</label>
                <input
                  type="text"
                  value={kopLine3}
                  onChange={(e) => setKopLine3(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-lg text-xs p-2 text-slate-800 font-bold"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-slate-400 uppercase">KOP Baris 4 (Alamat SPPG)</label>
                <input
                  type="text"
                  value={kopLine4}
                  onChange={(e) => setKopLine4(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-lg text-xs p-2 text-slate-800"
                />
              </div>
            </div>
            <p className="text-[10px] text-slate-400 italic">
              Tip: Pengaturan KOP, logo kiri/kanan, font Arial presisi, dan orientasi Potret/Lansekap otomatis disinkronkan dari Panel Profil SPPG.
            </p>
          </div>

          <div className="bg-slate-100 p-6 rounded-2xl border border-slate-300 shadow-inner flex justify-center overflow-x-auto print:bg-transparent print:p-0 print:border-none print:shadow-none">
            <div 
              id="print-area-food-cost" 
              className="bg-white p-8 border border-slate-400 shadow-md w-full max-w-[297mm] min-w-[210mm] text-slate-950 print:text-black print:border-none print:shadow-none print:p-0 print:m-0 space-y-10"
              style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
            >
              {/* SHEETS: PER-KATEGORI OR PORSI BESAR/KECIL */}
              {isPerKategoriMode ? (
                resolvedCategoryEntries.map((entry, entryIdx) => {
                  const entryPm = categoryCalcResultsMap[entry.id]?.pmCount ?? getEntryEffectivePmCount(entry);
                  return (
                    <div
                      key={entry.id}
                      className={`space-y-6 ${entryIdx > 0 ? "print:break-before-page pt-10 print:pt-0" : ""}`}
                    >
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
                        docKey="foodcost"
                      />

                      <div className="text-center space-y-1">
                        <h3 className="font-black text-black uppercase tracking-wide" style={{ fontSize: "13pt" }}>
                          LAPORAN KOMPONEN BAHAN BAKU &amp; FOOD COST - {entry.title.toUpperCase()}
                        </h3>
                        <p className="text-xs font-bold text-slate-700 uppercase">
                          SIKLUS HARI KE-{selectedDay} ({selectedType}) | NAMA MENU: {namaMenuHariIni || "-"} | PORSI:{" "}
                          {entry.porsi.toUpperCase()} | JUMLAH PM: {entryPm} ORANG
                        </p>
                      </div>

                      {renderCategoryFoodCostTable(entry, true)}

                      {profile && (
                        <div className="pt-4">
                          <AutoSignatureBlock
                            profile={profile}
                            docKey="foodcost"
                            dateText={currentCalendarDate || undefined}
                          />
                        </div>
                      )}
                    </div>
                  );
                })
              ) : (
                <>
                  {/* SHEET 1: PORSI BESAR PAGE */}
                  {(!appliedConfig.confirmed || showBesarTable) && (
                    <div className="space-y-6 print:break-after-page">
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
                        docKey="foodcost"
                      />

                      <div className="text-center space-y-1">
                        <h3 className="font-black text-black uppercase tracking-wide" style={{ fontSize: '13pt' }}>
                          LAPORAN KOMPONEN BAHAN BAKU &amp; FOOD COST - PORSI BESAR
                        </h3>
                        <p className="text-xs font-bold text-slate-700 uppercase">
                          SIKLUS HARI KE-{selectedDay} ({selectedType}) | NAMA MENU: {namaMenuHariIni || "-"} | JUMLAH PM: {pmBesar} ORANG
                        </p>
                      </div>

                      {renderCategorySignBanner("besar", true)}

                      {renderFoodCostTable("besar", true)}

                      {profile && (
                        <div className="pt-4">
                          <AutoSignatureBlock
                            profile={profile}
                            docKey="foodcost"
                            dateText={currentCalendarDate || undefined}
                          />
                        </div>
                      )}
                    </div>
                  )}

                  {/* SHEET 2: PORSI KECIL PAGE */}
                  {(!appliedConfig.confirmed || showKecilTable) && (
                    <div className={`space-y-6 ${(!appliedConfig.confirmed || showBesarTable) ? "print:break-before-page pt-10 print:pt-0" : ""}`}>
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
                        docKey="foodcost"
                      />

                      <div className="text-center space-y-1">
                        <h3 className="font-black text-black uppercase tracking-wide" style={{ fontSize: '13pt' }}>
                          LAPORAN KOMPONEN BAHAN BAKU &amp; FOOD COST - PORSI KECIL
                        </h3>
                        <p className="text-xs font-bold text-slate-700 uppercase">
                          SIKLUS HARI KE-{selectedDay} ({selectedType}) | NAMA MENU: {namaMenuHariIni || "-"} | JUMLAH PM: {pmKecil} ORANG
                        </p>
                      </div>

                      {renderCategorySignBanner("kecil", true)}

                      {renderFoodCostTable("kecil", true)}

                      {profile && (
                        <div className="pt-4">
                          <AutoSignatureBlock
                            profile={profile}
                            docKey="foodcost"
                            dateText={currentCalendarDate || undefined}
                          />
                        </div>
                      )}
                    </div>
                  )}
                </>
              )}

              {/* SHEETS FOR CUSTOM TABLES */}
              {customTables.map((table) => {
                return (
                  <div key={table.id} className="space-y-6 print:break-before-page pt-10 print:pt-0">
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
                      docKey="foodcost"
                    />

                    <div className="text-center space-y-1">
                      <h3 className="font-black text-black uppercase tracking-wide" style={{ fontSize: '13pt' }}>
                        LAPORAN KOMPONEN BAHAN BAKU & FOOD COST - {table.namaTabel.toUpperCase()}
                      </h3>
                      <p className="text-xs font-bold text-slate-700 uppercase">
                        SIKLUS HARI KE-{selectedDay} | PORSI: {table.porsi.toUpperCase()} | JUMLAH PM: {table.pmCount} ORANG
                      </p>
                    </div>

                    {renderCustomFoodCostTable(table, true)}

                    {profile && (
                      <div className="pt-4">
                        <AutoSignatureBlock
                          profile={profile}
                          docKey="foodcost"
                          dateText={currentCalendarDate || undefined}
                        />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        </div>

      {/* CALCULATOR POPUP MODAL */}
      {activeCalcField && (
        <div className="fixed inset-0 bg-black/55 backdrop-blur-xs flex items-center justify-center z-50 animate-fade-in no-print">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl p-6 shadow-2xl w-full max-w-sm overflow-hidden flex flex-col relative text-white">
            {/* Header */}
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-2 overflow-hidden">
                <Calculator className="w-5 h-5 text-indigo-400 shrink-0" />
                <span className="font-sans font-bold text-sm tracking-wide text-indigo-200 truncate">
                  Hitung {activeCalcField.field.toUpperCase()}: {activeCalcField.ingredientName || "Bahan Makanan"}
                </span>
              </div>
              <button
                type="button"
                onClick={() => setActiveCalcField(null)}
                className="text-slate-400 hover:text-white hover:bg-slate-800 p-1 rounded-full transition-colors shrink-0"
              >
                ✕
              </button>
            </div>

            {/* Display Screen */}
            <div className="bg-slate-950 p-3 rounded-2xl mb-4 border border-slate-800 flex flex-col justify-between shadow-inner">
              <label className="text-[10px] font-bold text-indigo-400 uppercase tracking-wider mb-1 block">
                Ketik / Keyboard:
              </label>
              <input
                type="text"
                autoFocus
                value={calcExpr}
                onChange={(e) => {
                  const val = e.target.value;
                  setCalcExpr(val);
                  const res = safeEvalMath(val);
                  setCalcResult(res || val || "");
                }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    saveCalculatorResult(calcResult || calcExpr);
                  } else if (e.key === "Escape") {
                    e.preventDefault();
                    setActiveCalcField(null);
                  }
                }}
                placeholder="0"
                className="w-full bg-slate-900 border border-slate-700 focus:border-indigo-500 rounded-xl px-3 py-1.5 text-right font-mono text-white text-base font-semibold focus:ring-2 focus:ring-indigo-500/50 focus:outline-none transition-all placeholder:text-slate-600 mb-2"
              />
              <div className="text-right font-mono">
                <span className="text-[10px] text-slate-500 block">Hasil:</span>
                <span className="text-emerald-400 text-2xl font-black truncate block">
                  {calcResult || "0"}
                </span>
              </div>
            </div>

            {/* Keypad Grid */}
            <div className="grid grid-cols-4 gap-2 mb-4">
              {/* Row 1 */}
              <button
                type="button"
                onClick={() => handleCalcKeyPress("C")}
                className="p-3 bg-rose-950 hover:bg-rose-900 border border-rose-900 text-rose-300 font-bold rounded-xl text-center text-sm transition-all"
              >
                C
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("(")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-indigo-300 font-bold rounded-xl text-center text-sm transition-all"
              >
                (
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress(")")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-indigo-300 font-bold rounded-xl text-center text-sm transition-all"
              >
                )
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("/")}
                className="p-3 bg-indigo-950 hover:bg-indigo-900 border border-indigo-900 text-indigo-300 font-bold rounded-xl text-center text-sm transition-all"
              >
                ÷
              </button>

              {/* Row 2 */}
              <button
                type="button"
                onClick={() => handleCalcKeyPress("7")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                7
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("8")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                8
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("9")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                9
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("*")}
                className="p-3 bg-indigo-950 hover:bg-indigo-900 border border-indigo-900 text-indigo-300 font-bold rounded-xl text-center text-sm transition-all"
              >
                ×
              </button>

              {/* Row 3 */}
              <button
                type="button"
                onClick={() => handleCalcKeyPress("4")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                4
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("5")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                5
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("6")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                6
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("-")}
                className="p-3 bg-indigo-950 hover:bg-indigo-900 border border-indigo-900 text-indigo-300 font-bold rounded-xl text-center text-sm transition-all"
              >
                -
              </button>

              {/* Row 4 */}
              <button
                type="button"
                onClick={() => handleCalcKeyPress("1")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                1
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("2")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                2
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("3")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                3
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("+")}
                className="p-3 bg-indigo-950 hover:bg-indigo-900 border border-indigo-900 text-indigo-300 font-bold rounded-xl text-center text-sm transition-all"
              >
                +
              </button>

              {/* Row 5 */}
              <button
                type="button"
                onClick={() => handleCalcKeyPress("0")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all col-span-2"
              >
                0
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress(".")}
                className="p-3 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-center text-sm transition-all"
              >
                .
              </button>
              <button
                type="button"
                onClick={() => handleCalcKeyPress("Del")}
                className="p-3 bg-rose-950 hover:bg-rose-900 border border-rose-900 text-rose-300 font-bold rounded-xl text-center text-sm transition-all"
              >
                ⌫
              </button>
            </div>

            {/* Actions Panel */}
            <div className="grid grid-cols-2 gap-2 mt-2">
              <button
                type="button"
                onClick={() => setActiveCalcField(null)}
                className="py-3 px-4 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold rounded-xl text-sm transition-all uppercase"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={() => saveCalculatorResult(calcResult || calcExpr)}
                className="py-3 px-4 bg-emerald-600 hover:bg-emerald-500 text-white font-extrabold rounded-xl text-sm transition-all shadow-md shadow-emerald-950/20 uppercase"
              >
                Simpan
              </button>
            </div>
          </div>
        </div>
      )}

      {/* BUFFER & JUMLAH CONFIG MODAL */}
      {activeBufferConfig && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center z-50 animate-fade-in no-print p-4">
          <div className="bg-white border border-slate-200 rounded-3xl p-6 shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col relative text-slate-800 animate-in zoom-in-95 duration-200">
            {/* Header */}
            <div className="flex items-center justify-between mb-5 pb-4 border-b border-slate-100">
              <div className="flex items-center gap-3">
                <div className="p-2.5 bg-indigo-50 text-indigo-600 rounded-2xl">
                  <Settings className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-base text-slate-900 leading-tight">
                    Konfigurasi Penyusutan/Kerusakan (Buffer) & Jumlah + Buffer
                  </h3>
                  <p className="text-xs text-slate-500 font-medium mt-0.5">
                    Bahan: <span className="text-indigo-600 font-bold font-mono">{activeBufferConfig.namaBahan}</span>
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setActiveBufferConfig(null)}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-100 p-1.5 rounded-full transition cursor-pointer"
              >
                ✕
              </button>
            </div>

            {/* Scrollable Content */}
            <div className="space-y-6 max-h-[60vh] overflow-y-auto pr-1">
              
              {/* Bagian 1: BUFFER */}
              <div className="space-y-3">
                <div className="flex items-center gap-2">
                  <span className="flex items-center justify-center w-5 h-5 bg-indigo-100 text-indigo-700 text-xs font-black rounded-full">
                    1
                  </span>
                  <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider">
                    Pilih Basis Perhitungan Nilai Penyusutan/Kerusakan (Buffer)
                  </h4>
                </div>
                <p className="text-xs text-slate-500 font-medium leading-relaxed">
                  Tentukan basis perhitungan untuk menentukan nilai penyusutan (Buffer):
                </p>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
                  {/* Option: Total Kilogram */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, bufferBase: "kg" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.bufferBase === "kg" || activeBufferConfig.bufferBase === "custom" || !activeBufferConfig.bufferBase
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="bufferBase"
                      checked={activeBufferConfig.bufferBase === "kg" || activeBufferConfig.bufferBase === "custom" || !activeBufferConfig.bufferBase}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Kilogram (KG)
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Otomatis dari total kebutuhan bersih (KG)
                      </span>
                    </div>
                  </label>

                  {/* Option: Total Potongan */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, bufferBase: "potong" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.bufferBase === "potong"
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="bufferBase"
                      checked={activeBufferConfig.bufferBase === "potong"}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Potongan
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Penyusutan dari jumlah potong porsi PM
                      </span>
                    </div>
                  </label>

                  {/* Option: Total Ekor */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, bufferBase: "ekor" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.bufferBase === "ekor"
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="bufferBase"
                      checked={activeBufferConfig.bufferBase === "ekor"}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Ekor
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Penyusutan dari total jumlah ekor porsi PM
                      </span>
                    </div>
                  </label>
                </div>
              </div>

              {/* Bagian 2: JUMLAH & BUFFER */}
              <div className="space-y-4 pt-4 border-t border-slate-100">
                <div className="flex items-center gap-2">
                  <span className="flex items-center justify-center w-5 h-5 bg-indigo-100 text-indigo-700 text-xs font-black rounded-full">
                    2
                  </span>
                  <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider">
                    Pilih Basis Perhitungan Nilai Jumlah + Buffer
                  </h4>
                </div>
                <p className="text-xs text-slate-500 font-medium leading-relaxed">
                  Pilih nilai dasar yang akan ditambahkan dengan buffer untuk masuk secara otomatis ke kolom <strong>Jumlah + Buffer</strong>:
                </p>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {/* Option: Total kilo + Buffer */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, jumlahBufferChoice: "kg_with" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.jumlahBufferChoice === "kg_with" || activeBufferConfig.jumlahBufferChoice === "custom_with" || !activeBufferConfig.jumlahBufferChoice
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="jumlahBufferChoice"
                      checked={activeBufferConfig.jumlahBufferChoice === "kg_with" || activeBufferConfig.jumlahBufferChoice === "custom_with" || !activeBufferConfig.jumlahBufferChoice}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Kilo + Buffer (KG)
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Kebutuhan (KG) ditambah nilai penyusutan buffer
                      </span>
                    </div>
                  </label>

                  {/* Option: Total kilo tanpa buffer */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, jumlahBufferChoice: "kg_without" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.jumlahBufferChoice === "kg_without" || activeBufferConfig.jumlahBufferChoice === "custom_without"
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="jumlahBufferChoice"
                      checked={activeBufferConfig.jumlahBufferChoice === "kg_without" || activeBufferConfig.jumlahBufferChoice === "custom_without"}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Kilo Saja
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Mengambil nilai berat total (KG) tanpa buffer
                      </span>
                    </div>
                  </label>

                  {/* Option: Total potong + buffer */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, jumlahBufferChoice: "potong_with" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.jumlahBufferChoice === "potong_with"
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="jumlahBufferChoice"
                      checked={activeBufferConfig.jumlahBufferChoice === "potong_with"}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Potong + Buffer
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Total potongan porsi ditambah nilai buffer
                      </span>
                    </div>
                  </label>

                  {/* Option: Total potong tanpa buffer */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, jumlahBufferChoice: "potong_without" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.jumlahBufferChoice === "potong_without"
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="jumlahBufferChoice"
                      checked={activeBufferConfig.jumlahBufferChoice === "potong_without"}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Potong Saja
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Mengambil total potongan porsi tanpa buffer
                      </span>
                    </div>
                  </label>

                  {/* Option: Total ekor + buffer */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, jumlahBufferChoice: "ekor_with" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.jumlahBufferChoice === "ekor_with"
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="jumlahBufferChoice"
                      checked={activeBufferConfig.jumlahBufferChoice === "ekor_with"}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Ekor + Buffer
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Total ekor porsi ditambah nilai buffer
                      </span>
                    </div>
                  </label>

                  {/* Option: Total ekor tanpa buffer */}
                  <label
                    onClick={() => setActiveBufferConfig({ ...activeBufferConfig, jumlahBufferChoice: "ekor_without" })}
                    className={`p-3 rounded-2xl border flex items-start gap-3 cursor-pointer transition-all ${
                      activeBufferConfig.jumlahBufferChoice === "ekor_without"
                        ? "bg-indigo-50/50 border-indigo-400 ring-2 ring-indigo-100 shadow-xs"
                        : "bg-white border-slate-200 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="jumlahBufferChoice"
                      checked={activeBufferConfig.jumlahBufferChoice === "ekor_without"}
                      onChange={() => {}}
                      className="mt-1 text-indigo-600 focus:ring-indigo-500 h-4 w-4 cursor-pointer shrink-0"
                    />
                    <div className="space-y-0.5 leading-snug">
                      <span className="text-xs font-extrabold text-slate-800 block">
                        Total Ekor Saja
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium block leading-tight">
                        Mengambil total ekor porsi tanpa buffer
                      </span>
                    </div>
                  </label>
                </div>
              </div>

            </div>

            {/* Footer Buttons */}
            <div className="flex items-center justify-end gap-3 mt-6 pt-4 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setActiveBufferConfig(null)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-xl transition cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={() => saveBufferConfig(activeBufferConfig)}
                className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-extrabold rounded-xl transition shadow-md shadow-indigo-100 cursor-pointer"
              >
                Simpan Konfigurasi
              </button>
            </div>
          </div>
        </div>
      )}

      {/* CUSTOM NON-BLOCKING CONFIRMATION MODAL */}
      {confirmModal.isOpen && (
        <div 
          className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-[9999] p-4 no-print animate-fade-in"
          onClick={() => setConfirmModal(prev => ({ ...prev, isOpen: false }))}
        >
          <div 
            className="bg-white border border-slate-100 rounded-3xl p-6 shadow-2xl w-full max-w-sm animate-in zoom-in-95 duration-150 space-y-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-start gap-3.5">
              <div className={`p-2.5 rounded-full shrink-0 ${
                confirmModal.variant === "warning" 
                  ? "bg-amber-50 text-amber-600" 
                  : confirmModal.variant === "info" 
                    ? "bg-indigo-50 text-indigo-600" 
                    : "bg-rose-50 text-rose-600"
              }`}>
                <AlertTriangle className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h4 className="font-black text-slate-900 text-sm tracking-tight">
                  {confirmModal.title}
                </h4>
                <p className="text-[11px] text-slate-500 font-semibold leading-relaxed">
                  {confirmModal.message}
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setConfirmModal(prev => ({ ...prev, isOpen: false }))}
                className="px-3.5 py-2 text-xs font-bold text-slate-600 hover:text-slate-800 bg-slate-100 hover:bg-slate-200 rounded-xl transition"
              >
                {confirmModal.cancelText || "Batal"}
              </button>
              <button
                type="button"
                onClick={confirmModal.onConfirm}
                className={`px-4 py-2 text-xs font-black text-white rounded-xl transition shadow-xs ${
                  confirmModal.variant === "warning"
                    ? "bg-amber-600 hover:bg-amber-700"
                    : confirmModal.variant === "info"
                      ? "bg-indigo-600 hover:bg-indigo-700"
                      : "bg-rose-600 hover:bg-rose-700"
                }`}
              >
                {confirmModal.confirmText || "Hapus"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* OVER BUDGET: UBAH MENU & PENYESUAIAN ANGGARAN MODAL */}
      {changeMenuModal?.isOpen && (
        <div
          className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-[9999] p-4 no-print"
          onClick={() => setChangeMenuModal(null)}
        >
          <div
            className="bg-white border border-slate-200 rounded-3xl p-6 shadow-2xl w-full max-w-lg space-y-5"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-start justify-between gap-3">
              <div className="flex items-start gap-3">
                <div className="p-2.5 rounded-xl bg-rose-50 text-rose-600 shrink-0">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="font-black text-slate-900 text-sm uppercase tracking-tight">
                    Ubah Menu / Penyesuaian Anggaran — {changeMenuModal.label}
                  </h4>
                  <p className="text-xs text-slate-500 mt-0.5">
                    Hari Ke-{selectedDay} ({currentCalendarDate}) · Kelebihan{" "}
                    <span className="font-mono font-bold text-rose-600">
                      +Rp {formatThousandSeparator(Math.round(changeMenuModal.actualCost - changeMenuModal.targetPrice))} / porsi
                    </span>
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setChangeMenuModal(null)}
                className="text-xs font-bold text-slate-400 hover:text-slate-700 cursor-pointer"
              >
                Tutup
              </button>
            </div>

            <div className="grid grid-cols-2 gap-3 bg-slate-50 p-3.5 rounded-2xl border border-slate-200/80">
              <div>
                <span className="text-[10px] font-bold text-slate-500 uppercase block">Harga Per Porsi Saat Ini</span>
                <span className="text-base font-black font-mono text-rose-700">
                  Rp {formatThousandSeparator(Math.round(changeMenuModal.actualCost))}
                </span>
              </div>
              <div>
                <span className="text-[10px] font-bold text-slate-500 uppercase block">Plafon Target (Data PM)</span>
                <span className="text-base font-black font-mono text-emerald-700">
                  Rp {formatThousandSeparator(changeMenuModal.targetPrice)}
                </span>
              </div>
            </div>

            <div className="space-y-1.5">
              <label className="text-[11px] font-bold text-slate-700 uppercase block">
                Ganti Nama Menu Hari Ini ({selectedType})
              </label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={changeMenuNameInput}
                  onChange={(e) => setChangeMenuNameInput(e.target.value)}
                  placeholder="Contoh: Nasi + Telur Balado + Tempe + Tumis Kol..."
                  className="flex-1 bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                />
                <button
                  type="button"
                  onClick={() => {
                    updateDayData({ ...currentDayData, namaMenu: changeMenuNameInput.trim() });
                    const targetEl = document.getElementById(`fc-block-${changeMenuModal.scope}`);
                    setChangeMenuModal(null);
                    if (targetEl) targetEl.scrollIntoView({ behavior: "smooth", block: "start" });
                  }}
                  className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition cursor-pointer shrink-0"
                >
                  Simpan & Edit Bahan
                </button>
              </div>
            </div>

            {changeMenuModal.items.length > 0 && (
              <div className="space-y-2">
                <span className="text-[11px] font-bold text-slate-700 uppercase block">
                  Komponen Biaya Tertinggi Per Porsi:
                </span>
                <div className="divide-y divide-slate-100 border border-slate-200 rounded-2xl overflow-hidden bg-white">
                  {changeMenuModal.items.map((item) => (
                    <div key={item.id} className="px-3 py-2 flex items-center justify-between text-xs">
                      <span className="font-semibold text-slate-800 truncate pr-2">{item.nama}</span>
                      <span className="font-mono font-bold text-rose-700 shrink-0">
                        Rp {formatThousandSeparator(Math.round(item.costPerPortion))}/porsi
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="flex flex-wrap items-center justify-between gap-2 pt-3 border-t border-slate-100">
              <div className="flex flex-wrap items-center gap-2">
                {onNavigateToMasterMenu && (
                  <button
                    type="button"
                    onClick={() => {
                      setChangeMenuModal(null);
                      onNavigateToMasterMenu();
                    }}
                    className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-extrabold transition cursor-pointer"
                  >
                    Buka Master Menu →
                  </button>
                )}
                {selectedGroup !== "custom" && (
                  <button
                    type="button"
                    onClick={() => {
                      autoMatchFromMasterMenu();
                      setChangeMenuModal(null);
                    }}
                    className="px-3.5 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl text-xs font-bold transition cursor-pointer"
                  >
                    Sinkron Ulang Master Menu
                  </button>
                )}
              </div>
              <button
                type="button"
                onClick={() => {
                  const targetEl = document.getElementById(`fc-block-${changeMenuModal.scope}`);
                  setChangeMenuModal(null);
                  if (targetEl) targetEl.scrollIntoView({ behavior: "smooth", block: "start" });
                }}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-xl text-xs font-bold transition cursor-pointer"
              >
                Sesuaikan Bahan di Tabel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* KONFIRMASI PILIHAN KATEGORI FOOD COST MODAL */}
      {showCategoryConfirmModal && (
        <div
          className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-[9999] p-4 no-print"
          onClick={() => setShowCategoryConfirmModal(false)}
        >
          <div
            className="bg-white border border-slate-200 rounded-3xl p-6 shadow-2xl w-full max-w-lg space-y-5"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-start justify-between gap-3">
              <div className="flex items-start gap-3">
                <div className="p-2.5 rounded-xl bg-emerald-50 text-emerald-700 shrink-0">
                  <Layers className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="font-black text-slate-900 text-sm uppercase tracking-tight">
                    Konfirmasi Penerapan Kategori Food Cost
                  </h4>
                  <p className="text-xs text-slate-500 mt-0.5">
                    Hari Ke-{selectedDay} ({currentCalendarDate}) · Diet {selectedType}
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowCategoryConfirmModal(false)}
                className="text-xs font-bold text-slate-400 hover:text-slate-700 cursor-pointer"
              >
                Tutup
              </button>
            </div>

            <div className="space-y-3 bg-slate-50 p-4 rounded-2xl border border-slate-200/80 text-xs">
              <div className="flex items-center justify-between gap-2 pb-2 border-b border-slate-200">
                <span className="font-bold text-slate-500 uppercase text-[10px]">Mode Tampilan:</span>
                <span className="font-black text-slate-900">
                  {draftLayoutMode === "per_kategori"
                    ? "Masing-Masing 1 Kategori (Bisa Gabungan Pilihan)"
                    : draftScope === "kecil"
                      ? "Gabung Porsi Kecil"
                      : draftScope === "besar"
                        ? "Gabung Porsi Besar"
                        : "Gabung Semua Porsi (Kecil & Besar)"}
                </span>
              </div>

              {draftLayoutMode === "per_kategori" ? (
                (() => {
                  const groups = (
                    draftCategoryGroupsByGroup[activeGroupKey] || buildDefaultGroupsForTabGroup(activeGroupKey)
                  ).filter((g) => draftScope === "all" || g.porsi === draftScope);
                  const totalPmPreview = groups.reduce(
                    (sum, g) => sum + getEntryAutoPmCount(g.sasaranIds, g.porsi),
                    0
                  );
                  return (
                    <div className="space-y-1.5 max-h-48 overflow-y-auto">
                      <div className="flex items-center justify-between">
                        <span className="font-black text-emerald-800 uppercase text-[10px]">
                          Panel Kategori Aktif ({groups.length})
                        </span>
                        <span className="font-mono font-bold text-emerald-800">{totalPmPreview} PM</span>
                      </div>
                      <div className="space-y-1">
                        {groups.map((g) => (
                          <div
                            key={g.id}
                            className="flex items-center justify-between bg-white px-2.5 py-1.5 rounded-lg border border-slate-200 text-[11px]"
                          >
                            <span className="font-bold text-slate-800">{g.title}</span>
                            <span className="font-mono font-bold text-slate-600">
                              {g.porsi === "besar" ? "Besar" : "Kecil"} · {getEntryAutoPmCount(g.sasaranIds, g.porsi)} PM
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()
              ) : (
                <>
                  {(draftScope === "kecil" || draftScope === "all") && (
                    <div className="space-y-1">
                      <div className="flex items-center justify-between">
                        <span className="font-black text-indigo-800 uppercase text-[10px]">
                          Kategori Porsi Kecil ({draftKecilIds.length})
                        </span>
                        <span className="font-mono font-bold text-indigo-800">
                          {draftKecilIds.reduce((sum, id) => sum + getSasaranPmCount(id, "kecil"), 0)} PM
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-700 font-medium">
                        {draftKecilIds.length > 0
                          ? draftKecilIds.map((id) => getSasaranLabel(id)).join(", ")
                          : "-"}
                      </p>
                    </div>
                  )}

                  {(draftScope === "besar" || draftScope === "all") && (
                    <div className="space-y-1 pt-1">
                      <div className="flex items-center justify-between">
                        <span className="font-black text-cyan-800 uppercase text-[10px]">
                          Kategori Porsi Besar ({draftBesarIds.length})
                        </span>
                        <span className="font-mono font-bold text-cyan-800">
                          {draftBesarIds.reduce((sum, id) => sum + getSasaranPmCount(id, "besar"), 0)} PM
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-700 font-medium">
                        {draftBesarIds.length > 0
                          ? draftBesarIds.map((id) => getSasaranLabel(id)).join(", ")
                          : "-"}
                      </p>
                    </div>
                  )}
                </>
              )}
            </div>

            {selectedGroup !== "custom" && namaMenuHariIni && (
              <label className="flex items-start gap-2.5 p-3 rounded-xl bg-indigo-50/70 border border-indigo-200 cursor-pointer text-xs">
                <input
                  type="checkbox"
                  checked={autoFillOnConfirm}
                  onChange={(e) => setAutoFillOnConfirm(e.target.checked)}
                  className="mt-0.5 rounded border-indigo-300 text-indigo-600 focus:ring-indigo-500"
                />
                <span className="text-indigo-950 font-semibold leading-snug">
                  Otomatis salin &amp; petakan bahan makanan dari Master Menu (<strong>{namaMenuHariIni}</strong>) jika tabel masih kosong.
                </span>
              </label>
            )}

            <div className="flex items-center justify-end gap-2.5 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setShowCategoryConfirmModal(false)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition cursor-pointer"
              >
                Periksa Kembali
              </button>
              <button
                type="button"
                id="btn-modal-confirm-apply-fc"
                onClick={handleConfirmApplyFoodCost}
                className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-black transition shadow-sm cursor-pointer flex items-center gap-1.5"
              >
                <CheckCircle2 className="w-4 h-4" />
                Ya, Konfirmasi &amp; Tampilkan Otomatis
              </button>
            </div>
          </div>
        </div>
      )}

      {/* PASTE CONFIRMATION MODAL WHEN TARGET ALREADY HAS DATA */}
      {pasteConfirmModal?.isOpen && (
        <div
          className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-[9999] p-4 no-print"
          onClick={() => setPasteConfirmModal(null)}
        >
          <div
            className="bg-white border border-slate-200 rounded-3xl p-6 shadow-2xl w-full max-w-md space-y-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-start gap-3">
              <div className="p-2.5 rounded-xl bg-amber-50 text-amber-700 shrink-0">
                <ClipboardCheck className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h4 className="font-black text-slate-900 text-sm uppercase tracking-tight">
                  Tabel Tujuan Sudah Terisi
                </h4>
                <p className="text-xs text-slate-600 leading-relaxed">
                  Food Cost <strong>&ldquo;{pasteConfirmModal.targetLabel}&rdquo;</strong> saat ini sudah memiliki{" "}
                  <strong>{pasteConfirmModal.existingCount} bahan</strong>. Pilih cara menempel isian yang disalin:
                </p>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row sm:items-center justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setPasteConfirmModal(null)}
                className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={pasteConfirmModal.onAppend}
                className="px-3.5 py-2 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-800 rounded-xl text-xs font-black transition cursor-pointer"
              >
                Tambahkan di Bawahnya
              </button>
              <button
                type="button"
                onClick={pasteConfirmModal.onReplace}
                className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-xl text-xs font-black transition shadow-xs cursor-pointer"
              >
                Ganti Semua Isian
              </button>
            </div>
          </div>
        </div>
      )}

      <PrintPreviewModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title={`Pratinjau Cetak: Food Cost ${selectedType} — Hari Kerja ${selectedDay}`}
        elementId="print-area-food-cost"
        filename={`Food_Cost_Hari_${selectedDay}_${selectedType}`}
        defaultScale={100}
        defaultPaperSize={paperSize as "A4" | "F4"}
        defaultOrientation="auto"
      />
    </div>
  );
}
