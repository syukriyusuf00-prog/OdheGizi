/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  SPPGProfile,
  HariPM,
  FoodCostDay,
  TKPIItem,
  MasterMenu,
  UnifiedKopConfig,
  DEFAULT_UNIFIED_KOP_CONFIG,
} from "../types";
import { calculateDay, formatRupiah } from "../utils/calc";
import {
  Printer,
  Download,
  RefreshCw,
  Plus,
  Trash2,
  FileText,
  Combine,
  Eye,
  Maximize2,
  Minimize2,
  FileDown,
  Zap,
  Layers,
  X,
  Edit3,
  Check,
  Sparkles,
} from "lucide-react";
import { PriceCalculatorPopover } from "./PriceCalculatorPopover";
import * as XLSX from "xlsx";
import { printElementSmart, exportToPDF } from "../lib/printUtils";
import { KopSuratRenderHeader, LogoCrop } from "./KopSuratConfigSection";
import AutoSignatureBlock from "./AutoSignatureBlock";
import PrintPreviewModal from "./PrintPreviewModal";

const formatThousandSeparator = (val: number) => {
  return new Intl.NumberFormat("id-ID").format(Math.round(val || 0));
};

interface NotaPesananLogistikTabProps {
  profile: SPPGProfile;
  foodCostDays: FoodCostDay[];
  tkpiList: TKPIItem[];
  harianPM: HariPM[];
  masterMenu: MasterMenu;
  kopLine1: string;
  setKopLine1?: (val: string) => void;
  kopLine2: string;
  setKopLine2?: (val: string) => void;
  kopLine3: string;
  setKopLine3?: (val: string) => void;
  kopLine4: string;
  setKopLine4?: (val: string) => void;
  leftLogo: string;
  setLeftLogo?: (val: string) => void;
  rightLogo: string;
  setRightLogo?: (val: string) => void;
  leftLogoCrop?: LogoCrop;
  setLeftLogoCrop?: (val: LogoCrop) => void;
  rightLogoCrop?: LogoCrop;
  setRightLogoCrop?: (val: LogoCrop) => void;
  paperSize?: "A4" | "F4";
  setPaperSize?: (val: "A4" | "F4") => void;
  kopConfig?: UnifiedKopConfig;
  bumbuPct?: number;
  isAdmin?: boolean;
}

interface NotaItem {
  id: string;
  nama: string;
  jumlah: number;
  satuan: string;
  hargaSatuan: number;
  sumKg?: number;
  sumPotong?: number;
  sumEkor?: number;
  sumButir?: number;
  defaultUnit?: string;
  selectedUnitType?: "foodcost" | "kg" | "potong" | "ekor" | "custom";
  customUnit?: string;
}

type BuiltinSheetKind = "otomatis" | "bahan_baku" | "bumbu" | "custom";

interface NotaSheetMeta {
  id: string;
  title: string;
  kind: BuiltinSheetKind;
  kepada: string;
  subtitle: string;
}

const DEFAULT_BUILTIN_SHEETS: NotaSheetMeta[] = [
  {
    id: "sheet_otomatis",
    title: "Nota Otomatis (Semua Bahan)",
    kind: "otomatis",
    kepada: "Penyedia Logistik Bahan Makanan",
    subtitle: "NOTA PESANAN BAHAN MAKANAN (OTOMATIS)",
  },
  {
    id: "sheet_bahan_baku",
    title: "Nota Bahan Baku Utama",
    kind: "bahan_baku",
    kepada: "Penyedia Bahan Baku Utama",
    subtitle: "NOTA PESANAN BAHAN BAKU UTAMA",
  },
  {
    id: "sheet_bumbu",
    title: "Nota Bahan Baku Bumbu",
    kind: "bumbu",
    kepada: "Penyedia Bahan Baku Bumbu & Rempah",
    subtitle: "NOTA PESANAN BAHAN BAKU BUMBU",
  },
];

const QUICK_BUMBU_PRESETS: Array<{ nama: string; satuan: string; harga: number }> = [
  { nama: "Bawang Merah", satuan: "Kg", harga: 40000 },
  { nama: "Bawang Putih", satuan: "Kg", harga: 42000 },
  { nama: "Minyak Goreng", satuan: "Liter", harga: 18000 },
  { nama: "Garam Halus Beryodium", satuan: "Bks", harga: 3500 },
  { nama: "Gula Pasir", satuan: "Kg", harga: 18000 },
  { nama: "Kecap Manis", satuan: "Botol", harga: 22000 },
  { nama: "Saus Tiram", satuan: "Botol", harga: 20000 },
  { nama: "Merica / Lada Bubuk", satuan: "Bks", harga: 12000 },
  { nama: "Ketumbar Bubuk", satuan: "Bks", harga: 8000 },
  { nama: "Kaldu Bubuk (Ayam/Sapi)", satuan: "Bks", harga: 15000 },
  { nama: "Kunyit Segar", satuan: "Kg", harga: 18000 },
  { nama: "Jahe & Lengkuas", satuan: "Kg", harga: 22000 },
  { nama: "Kemiri", satuan: "Kg", harga: 55000 },
  { nama: "Serai & Daun Salam / Jeruk", satuan: "Ikat", harga: 5000 },
  { nama: "Cabai Merah / Keriting", satuan: "Kg", harga: 45000 },
  { nama: "Tomat Segar", satuan: "Kg", harga: 16000 },
];

const BUMBU_KEYWORDS = [
  "bumbu",
  "bawang",
  "garam",
  "gula",
  "merica",
  "lada",
  "ketumbar",
  "kunyit",
  "jahe",
  "lengkuas",
  "laos",
  "serai",
  "sereh",
  "daun salam",
  "daun jeruk",
  "kemiri",
  "kecap",
  "saus",
  "saos",
  "kaldu",
  "royco",
  "masako",
  "penyedap",
  "vetsin",
  "micin",
  "minyak",
  "cuka",
  "asam jawa",
  "terasi",
  "cabai",
  "cabe",
  "jinten",
  "jintan",
  "kayu manis",
  "cengkeh",
  "pala",
  "kapulaga",
  "wijen",
  "mentega",
  "margarin",
  "santan",
  "tepung bumbu",
  "bumbu racik",
  "oregano",
  "peterseli",
];

function isBumbuIngredient(name: string): boolean {
  const lower = (name || "").toLowerCase();
  return BUMBU_KEYWORDS.some((kw) => lower.includes(kw));
}

const formatNumber = (num: number, decimalPlaces = 0) => {
  return new Intl.NumberFormat("id-ID", {
    minimumFractionDigits: decimalPlaces,
    maximumFractionDigits: decimalPlaces,
  }).format(num);
};

const formatIndonesianDate = (dateStr: string): string => {
  if (!dateStr) return "";
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
};

const guessSatuan = (name: string): string => {
  const lower = name.toLowerCase();
  if (lower.includes("minyak")) return "Liter";
  if (lower.includes("susu") && !lower.includes("bubuk")) return "Liter";
  if (lower.includes("kecap") || lower.includes("saus") || lower.includes("tiram")) return "Botol";
  if (lower.includes("pisang") && lower.includes("mas")) return "Rak";
  if (lower.includes("telur")) return "Butir";
  if (lower.includes("tahu") || lower.includes("tempe")) return "Pcs";
  if (
    lower.includes("garam") ||
    lower.includes("royco") ||
    lower.includes("masako") ||
    lower.includes("bumbu")
  ) {
    return "Bks";
  }
  return "Kg";
};

const formatJumlahDisplay = (val: number | string | undefined | null): string => {
  if (val === undefined || val === null || val === "" || val === 0) return "";
  const num = typeof val === "number" ? val : parseFloat(String(val));
  if (isNaN(num) || num === 0) return "";
  const rounded = Math.round(num * 10) / 10;
  return Number.isInteger(rounded) ? rounded.toString() : rounded.toFixed(1);
};

function padNotaItems(list: NotaItem[], minRows = 18): NotaItem[] {
  const padded = [...list];
  let idx = padded.length;
  while (padded.length < minRows) {
    padded.push({
      id: `pad_${idx}_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      nama: "",
      jumlah: 0,
      satuan: "",
      hargaSatuan: 0,
      sumKg: 0,
      sumPotong: 0,
      sumEkor: 0,
      sumButir: 0,
      defaultUnit: "",
      selectedUnitType: "custom",
      customUnit: "",
    });
    idx++;
  }
  return padded;
}

export default function NotaPesananLogistikTab({
  profile,
  foodCostDays,
  tkpiList,
  harianPM,
  kopLine1,
  kopLine2,
  kopLine3,
  kopLine4,
  leftLogo,
  rightLogo,
  leftLogoCrop = { top: 0, bottom: 0, left: 0, right: 0 },
  rightLogoCrop = { top: 0, bottom: 0, left: 0, right: 0 },
  paperSize = "A4",
  kopConfig = DEFAULT_UNIFIED_KOP_CONFIG,
  bumbuPct = 10,
}: NotaPesananLogistikTabProps) {
  const totalDays = Math.max(profile?.periodeDates?.length || 10, 1);

  const [selectedDay, setSelectedDay] = useState<number>(1);
  const [mode, setMode] = useState<"harian" | "gabungan">("harian");
  const [selectedDays, setSelectedDays] = useState<number[]>(() =>
    Array.from({ length: totalDays }, (_, i) => i + 1)
  );

  // Multi-Sheet State (Nota Otomatis, Bahan Baku Utama, Bahan Baku Bumbu, + Custom Sheets)
  const [sheets, setSheets] = useState<NotaSheetMeta[]>(() => {
    try {
      const saved = localStorage.getItem("sisper_nota_sheets_meta_v2");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length >= 3) return parsed;
      }
    } catch {}
    return DEFAULT_BUILTIN_SHEETS;
  });

  const [activeSheetId, setActiveSheetId] = useState<string>(() => {
    return localStorage.getItem("sisper_nota_active_sheet_v2") || "sheet_otomatis";
  });

  // Auto-sync Fix Food Cost mode toggle (when ON, Nota Otomatis always pulls live directly from Food Cost)
  const [autoModeFix, setAutoModeFix] = useState<boolean>(() => {
    const saved = localStorage.getItem("sisper_nota_auto_mode_fix");
    return saved === null ? true : saved === "true";
  });

  // Add/Edit/Delete Sheet UI states
  const [showAddSheetModal, setShowAddSheetModal] = useState(false);
  const [newSheetTitle, setNewSheetTitle] = useState("");
  const [newSheetKepada, setNewSheetKepada] = useState("Penyedia Logistik Bahan Makanan");
  const [newSheetTemplate, setNewSheetTemplate] = useState<"blank" | "bahan_baku" | "bumbu">("blank");
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState("");
  const [confirmDeleteSheet, setConfirmDeleteSheet] = useState<NotaSheetMeta | null>(null);

  const [items, setItems] = useState<NotaItem[]>([]);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingVal, setEditingVal] = useState<string>("");
  const [showPrintModal, setShowPrintModal] = useState<boolean>(false);
  const [previewWidth, setPreviewWidth] = useState<"standard" | "full">(
    () => (localStorage.getItem("nota_preview_width") as "standard" | "full") || "full"
  );

  // Metadata automatically synced from Profil SPPG (Kepada & Tanggal editable inline cleanly)
  const [tanggalOverride, setTanggalOverride] = useState<string>("");

  const activeSheet = useMemo(() => {
    return sheets.find((s) => s.id === activeSheetId) || sheets[0] || DEFAULT_BUILTIN_SHEETS[0];
  }, [sheets, activeSheetId]);

  useEffect(() => {
    localStorage.setItem("sisper_nota_sheets_meta_v2", JSON.stringify(sheets));
  }, [sheets]);

  useEffect(() => {
    localStorage.setItem("sisper_nota_active_sheet_v2", activeSheetId);
  }, [activeSheetId]);

  useEffect(() => {
    localStorage.setItem("sisper_nota_auto_mode_fix", String(autoModeFix));
  }, [autoModeFix]);

  useEffect(() => {
    localStorage.setItem("nota_preview_width", previewWidth);
  }, [previewWidth]);

  useEffect(() => {
    if (selectedDay > totalDays) setSelectedDay(totalDays);
    setSelectedDays((prev) => {
      const valid = prev.filter((d) => d >= 1 && d <= totalDays);
      return valid.length > 0 ? valid : Array.from({ length: totalDays }, (_, i) => i + 1);
    });
  }, [totalDays, selectedDay]);

  // Settings for Target Price and Classifications
  const [settings] = useState(() => {
    const saved = localStorage.getItem("sisper_pm_settings");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch {}
    }
    return {
      porsiKecilHarga: 8000,
      porsiBesarHarga: 10000,
      porsiKecilSasaranIds: [
        "tk_paud_lb",
        "sd_kelas_1_3",
        "anak_balita",
        "anak_balita_13_59",
        "balita_6_11",
      ],
      porsiBesarSasaranIds: [
        "sd_kelas_4_6",
        "smp_mts_smplb",
        "sma_smk_ma",
        "pendidik",
        "tenaga_kependidikan",
        "ibu_hamil",
        "ibu_menyusui",
      ],
    };
  });

  // Auto-derived Tanggal from Profil SPPG dates
  const autoTanggal = useMemo(() => {
    if (mode === "gabungan") {
      const datesList = selectedDays
        .map((d) => profile.periodeDates?.[d - 1])
        .filter(Boolean);
      if (datesList.length > 0) {
        const first = formatIndonesianDate(datesList[0]);
        const last = formatIndonesianDate(datesList[datesList.length - 1]);
        return first === last ? first : `Gabungan Siklus: ${first} s.d. ${last}`;
      }
      return "";
    }
    const defaultDateStr = profile.periodeDates?.[selectedDay - 1] || "";
    return defaultDateStr ? formatIndonesianDate(defaultDateStr) : "";
  }, [selectedDay, selectedDays, mode, profile.periodeDates]);

  // Clear manual date override whenever day/mode/profile dates change
  useEffect(() => {
    setTanggalOverride("");
  }, [selectedDay, selectedDays, mode, profile.periodeDates]);

  const effectiveTanggal = tanggalOverride || autoTanggal;
  const dari = profile.namaLembaga || "SPPG";
  const alamat = profile.alamat || "-";

  // Compute RAB, Realisasi Belanja, Bumbu Allocation, and Surplus
  const summaryStats = useMemo(() => {
    const isCombined = mode === "gabungan";
    const daysToProcess = isCombined ? selectedDays : [selectedDay];

    let totalRAB = 0;
    let rabPorsiKecilTotal = 0;
    let rabPorsiBesarTotal = 0;
    let totalPMKecil = 0;
    let totalPMBesar = 0;

    let totalActualCost = 0;
    let porsiKecilActualCostTotal = 0;
    let porsiBesarActualCostTotal = 0;

    let totalBumbu10Nominal = 0;
    let bumbuKecilNominalTotal = 0;
    let bumbuBesarNominalTotal = 0;

    daysToProcess.forEach((d) => {
      const dayPM = harianPM.find((h) => h.hariKe === d) || harianPM[0] || { sasaran: [] };
      let pmKecil = 0;
      let pmBesar = 0;

      (settings.porsiKecilSasaranIds || []).forEach((sId: string) => {
        const sasItem = dayPM.sasaran.find((item) => item.id === sId);
        if (sasItem) {
          pmKecil += (Number(sasItem.porsiKecil) || 0) + (Number(sasItem.porsiBesar) || 0);
        }
      });

      (settings.porsiBesarSasaranIds || []).forEach((sId: string) => {
        const sasItem = dayPM.sasaran.find((item) => item.id === sId);
        if (sasItem) {
          pmBesar += (Number(sasItem.porsiKecil) || 0) + (Number(sasItem.porsiBesar) || 0);
        }
      });

      const dayFC = foodCostDays.find((fd) => fd.hariKe === d);
      if (dayFC && dayFC.customPmBesarCount !== undefined) pmBesar = dayFC.customPmBesarCount;
      if (dayFC && dayFC.customPmKecilCount !== undefined) pmKecil = dayFC.customPmKecilCount;

      totalPMKecil += pmKecil;
      totalPMBesar += pmBesar;

      const rabKecilDay = pmKecil * (settings.porsiKecilHarga || 8000);
      const rabBesarDay = pmBesar * (settings.porsiBesarHarga || 10000);
      rabPorsiKecilTotal += rabKecilDay;
      rabPorsiBesarTotal += rabBesarDay;
      totalRAB += rabKecilDay + rabBesarDay;

      const calcDay = calculateDay(
        dayFC?.porsiBesarBahan || [],
        dayFC?.porsiKecilBahan || [],
        pmBesar,
        pmKecil,
        dayFC?.bufferPct ?? 3,
        tkpiList,
        dayFC?.bumbuPct ?? bumbuPct
      );

      porsiKecilActualCostTotal += calcDay.subtotalKecilCost;
      porsiBesarActualCostTotal += calcDay.subtotalBesarCost;
      totalActualCost += calcDay.subtotalKecilCost + calcDay.subtotalBesarCost;

      bumbuKecilNominalTotal += calcDay.bumbuKecilCost;
      bumbuBesarNominalTotal += calcDay.bumbuBesarCost;
      totalBumbu10Nominal += calcDay.bumbuKecilCost + calcDay.bumbuBesarCost;
    });

    const sisaEfisiensiAnggaran = totalRAB - totalActualCost;
    const porsiKecilPerPorsi =
      totalPMKecil > 0 ? Math.round(porsiKecilActualCostTotal / totalPMKecil) : 0;
    const porsiBesarPerPorsi =
      totalPMBesar > 0 ? Math.round(porsiBesarActualCostTotal / totalPMBesar) : 0;

    return {
      isCombined,
      daysToProcess,
      totalRAB,
      rabPorsiKecilTotal,
      rabPorsiBesarTotal,
      totalPMKecil,
      totalPMBesar,
      totalActualCost,
      porsiKecilActualCostTotal,
      porsiBesarActualCostTotal,
      porsiKecilPerPorsi,
      porsiBesarPerPorsi,
      totalBumbu10Nominal,
      bumbuKecilNominalTotal,
      bumbuBesarNominalTotal,
      sisaEfisiensiAnggaran,
    };
  }, [mode, selectedDays, selectedDay, harianPM, settings, foodCostDays, tkpiList, bumbuPct]);

  const calculateRowTotal = (row: NotaItem): number => {
    if (!row.nama || row.nama.trim() === "") return 0;
    const qty = Number(row.jumlah) || 0;
    const price = Number(row.hargaSatuan) || 0;
    return qty * price;
  };

  // Main Food Cost ingredient extraction (returns ALL unpadded items)
  const computeRawMergedIngredients = useCallback(
    (dayNums: number[]): NotaItem[] => {
      const parseVal = (val: number | string | undefined | null): number => {
        if (val === undefined || val === null) return 0;
        if (typeof val === "number") return val;
        const str = String(val).trim();
        if (!str) return 0;
        try {
          const sanitized = str.replace(/[^0-9+\-*/().\s]/g, "");
          if (!sanitized.trim()) return 0;
          const evaluated = new Function(`return (${sanitized})`)();
          if (typeof evaluated === "number" && !isNaN(evaluated) && isFinite(evaluated)) {
            return evaluated;
          }
        } catch {}
        const parsed = parseFloat(str);
        return isNaN(parsed) ? 0 : parsed;
      };

      const aggregated: Record<
        string,
        {
          nama: string;
          sumKg: number;
          sumPotong: number;
          sumEkor: number;
          sumButir: number;
          hargaSatuan: number;
          defaultUnit: string;
        }
      > = {};

      const getFoodCostUnit = (item: any) => {
        const jChoice = item.jumlahBufferChoice || "auto";
        let unit = "Kg";
        if (jChoice === "auto") {
          const selectedBase = item.bufferBase || "auto";
          if (selectedBase === "kg") unit = "Kg";
          else if (selectedBase === "potong") unit = "Potong";
          else if (selectedBase === "ekor") unit = "Ekor";
          else if (selectedBase === "custom") unit = "Custom";
          else {
            if (parseVal(item.potong) > 0) unit = "Potong";
            else if (parseVal(item.ekor) > 0) unit = "Ekor";
            else unit = "Kg";
          }
        } else {
          const baseType = jChoice.split("_")[0];
          if (baseType === "kg") unit = "Kg";
          else if (baseType === "potong") unit = "Potong";
          else if (baseType === "ekor") unit = "Ekor";
          else if (baseType === "custom") unit = "Custom";
        }
        if (unit === "Custom") return guessSatuan(item.nama);
        return unit;
      };

      const processItems = (itemsList: any[], pmCount: number, bufferPct: number) => {
        itemsList.forEach((item) => {
          const rawName = item.nama;
          if (!rawName || rawName.trim() === "" || rawName.trim().toLowerCase() === "ketik bahan...") {
            return;
          }
          if ((item.butir || 0) <= 0 && (item.totalKebutuhanKg || 0) <= 0) return;

          const key =
            item.tkpiId && item.tkpiId !== "unknown"
              ? item.tkpiId
              : rawName.trim().toLowerCase();

          const addBuffer =
            !item.jumlahBufferChoice ||
            item.jumlahBufferChoice.endsWith("_with") ||
            item.jumlahBufferChoice === "auto";
          const rowBufferPct =
            item.bufferBase === "custom" && item.bufferCustomVal !== undefined
              ? parseVal(item.bufferCustomVal)
              : bufferPct;
          const bufMult = addBuffer ? 1 + rowBufferPct / 100 : 1;

          const kgVal = (item.totalKebutuhanKg || 0) * bufMult;
          const potongVal = parseVal(item.potong) * pmCount * bufMult;
          const ekorVal = parseVal(item.ekor) * pmCount * bufMult;
          const activeVal =
            typeof item.butir === "number" && item.butir > 0
              ? item.butir
              : (item.totalKebutuhanKg || 0) * bufMult;
          const activeUnit = getFoodCostUnit(item);

          if (!aggregated[key]) {
            aggregated[key] = {
              nama: rawName.trim(),
              sumKg: kgVal,
              sumPotong: potongVal,
              sumEkor: ekorVal,
              sumButir: activeVal,
              hargaSatuan: item.hargaSatuan,
              defaultUnit: activeUnit,
            };
          } else {
            aggregated[key].sumKg += kgVal;
            aggregated[key].sumPotong += potongVal;
            aggregated[key].sumEkor += ekorVal;
            aggregated[key].sumButir += activeVal;
            if (item.hargaSatuan > aggregated[key].hargaSatuan) {
              aggregated[key].hargaSatuan = item.hargaSatuan;
            }
          }
        });
      };

      const besarSasaranIds = settings?.porsiBesarSasaranIds || [
        "sd_kelas_4_6",
        "smp_mts_smplb",
        "sma_smk_ma",
        "pendidik",
        "tenaga_kependidikan",
        "ibu_hamil",
        "ibu_menyusui",
      ];
      const kecilSasaranIds = settings?.porsiKecilSasaranIds || [
        "tk_paud_lb",
        "sd_kelas_1_3",
        "anak_balita",
        "anak_balita_13_59",
        "balita_6_11",
      ];

      dayNums.forEach((dayNum) => {
        const dayPM = harianPM.find((h) => h.hariKe === dayNum) || harianPM[0] || { sasaran: [] };
        const matchingDays = foodCostDays.filter((d) => d.hariKe === dayNum);

        matchingDays.forEach((dayData) => {
          let pmBesar = 0;
          let pmKecil = 0;
          const isAlergi = dayData.jenisMenu.includes("Alergi");

          besarSasaranIds.forEach((sId: string) => {
            const sasItem = dayPM.sasaran?.find((item) => item.id === sId);
            if (sasItem) {
              if (isAlergi) {
                const cAlergi = Number(sasItem.alergiBesar) || 0;
                const cNormal = (Number(sasItem.porsiKecil) || 0) + (Number(sasItem.porsiBesar) || 0);
                pmBesar += cAlergi > 0 ? cAlergi : cNormal;
              } else {
                pmBesar += (Number(sasItem.porsiKecil) || 0) + (Number(sasItem.porsiBesar) || 0);
              }
            }
          });

          kecilSasaranIds.forEach((sId: string) => {
            const sasItem = dayPM.sasaran?.find((item) => item.id === sId);
            if (sasItem) {
              if (isAlergi) {
                const cAlergi = Number(sasItem.alergiKecil) || 0;
                const cNormal = (Number(sasItem.porsiKecil) || 0) + (Number(sasItem.porsiBesar) || 0);
                pmKecil += cAlergi > 0 ? cAlergi : cNormal;
              } else {
                pmKecil += (Number(sasItem.porsiKecil) || 0) + (Number(sasItem.porsiBesar) || 0);
              }
            }
          });

          if (dayData.customPmBesarCount !== undefined) pmBesar = dayData.customPmBesarCount;
          if (dayData.customPmKecilCount !== undefined) pmKecil = dayData.customPmKecilCount;

          const result = calculateDay(
            dayData.porsiBesarBahan,
            dayData.porsiKecilBahan,
            pmBesar,
            pmKecil,
            dayData.bufferPct,
            tkpiList,
            dayData.bumbuPct ?? bumbuPct
          );
          processItems(result.porsiBesarItems, result.jumlahPMBesar, dayData.bufferPct);
          processItems(result.porsiKecilItems, result.jumlahPMKecil, dayData.bufferPct);
        });
      });

      // Custom Food Cost tables from localStorage
      let customTables: any[] = [];
      try {
        const saved = localStorage.getItem("sppg_custom_fc_tables_v1");
        if (saved) customTables = JSON.parse(saved);
      } catch {}

      customTables.forEach((table) => {
        const resultsCustom =
          table.porsi === "besar"
            ? calculateDay(table.bahanList, [], table.pmCount, 0, table.bufferPct || 5, tkpiList, bumbuPct)
            : calculateDay([], table.bahanList, 0, table.pmCount, table.bufferPct || 5, tkpiList, bumbuPct);
        const itemsCustom =
          table.porsi === "besar" ? resultsCustom.porsiBesarItems : resultsCustom.porsiKecilItems;
        processItems(itemsCustom, table.pmCount, table.bufferPct || 5);
      });

      return Object.values(aggregated).map((x, idx) => {
        const initialUnit = x.defaultUnit || guessSatuan(x.nama);
        const initialJumlah = x.sumButir > 0 ? x.sumButir : x.sumKg;
        return {
          id: `merged_${idx}_${Date.now()}`,
          nama: x.nama,
          jumlah: parseFloat(initialJumlah.toFixed(1)),
          satuan: initialUnit,
          hargaSatuan: x.hargaSatuan,
          sumKg: parseFloat(x.sumKg.toFixed(1)),
          sumPotong: parseFloat(x.sumPotong.toFixed(1)),
          sumEkor: parseFloat(x.sumEkor.toFixed(1)),
          sumButir: parseFloat(x.sumButir.toFixed(1)),
          defaultUnit: initialUnit,
          selectedUnitType: "foodcost" as const,
          customUnit: "",
        };
      });
    },
    [foodCostDays, harianPM, tkpiList, settings, bumbuPct]
  );

  // Build default rows for a given sheet kind
  const buildDefaultItemsForSheet = useCallback(
    (sheet: NotaSheetMeta, dayNums: number[]): NotaItem[] => {
      if (sheet.kind === "custom") {
        return padNotaItems([], 18);
      }
      const allRaw = computeRawMergedIngredients(dayNums);
      if (sheet.kind === "otomatis") {
        return padNotaItems(allRaw, 20);
      }
      if (sheet.kind === "bahan_baku") {
        const mainOnly = allRaw.filter((it) => !isBumbuIngredient(it.nama));
        return padNotaItems(mainOnly, 20);
      }
      // kind === "bumbu"
      const bumbuOnly = allRaw.filter((it) => isBumbuIngredient(it.nama));
      return padNotaItems(bumbuOnly, 18);
    },
    [computeRawMergedIngredients]
  );

  const getStorageKey = useCallback(
    (sheetId: string, m: string, sDay: number, sDays: number[]) => {
      const periodKey =
        m === "harian"
          ? `harian_${sDay}`
          : `gabungan_${[...sDays].sort((a, b) => a - b).join("_")}`;
      if (sheetId === "sheet_otomatis") {
        // Maintain backwards-compatibility with legacy key when saving edits
        return `sisper_nota_items_${periodKey}`;
      }
      return `sisper_nota_sheet_${sheetId}_${periodKey}`;
    },
    []
  );

  const loadSheetItems = useCallback(
    (forceReset = false) => {
      const dayNums = mode === "gabungan" ? selectedDays : [selectedDay];
      const key = getStorageKey(activeSheet.id, mode, selectedDay, selectedDays);

      // When Auto Mode (Food Cost Fix) is enabled and viewing "Nota Otomatis", pull live straight from Food Cost!
      const shouldForceLiveAuto =
        activeSheet.kind === "otomatis" && autoModeFix && !forceReset;

      if (!forceReset && !shouldForceLiveAuto) {
        const saved = localStorage.getItem(key);
        if (saved) {
          try {
            const parsed = JSON.parse(saved);
            if (Array.isArray(parsed) && parsed.length > 0) {
              setItems(parsed);
              return;
            }
          } catch {}
        }
      } else if (forceReset) {
        localStorage.removeItem(key);
      }

      const fresh = buildDefaultItemsForSheet(activeSheet, dayNums);
      setItems(fresh);
    },
    [
      activeSheet,
      autoModeFix,
      mode,
      selectedDay,
      selectedDays,
      getStorageKey,
      buildDefaultItemsForSheet,
    ]
  );

  useEffect(() => {
    loadSheetItems(false);
  }, [loadSheetItems]);

  // Persist edits automatically (skip overwriting custom saved cache when locked to live autoModeFix on sheet_otomatis)
  useEffect(() => {
    if (!items || items.length === 0) return;
    if (activeSheet.kind === "otomatis" && autoModeFix) return;
    const key = getStorageKey(activeSheet.id, mode, selectedDay, selectedDays);
    localStorage.setItem(key, JSON.stringify(items));
  }, [items, activeSheet, autoModeFix, mode, selectedDay, selectedDays, getStorageKey]);

  // Sheet CRUD handlers
  const handleCreateSheet = (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = newSheetTitle.trim();
    if (!trimmed) return;
    const id = `sheet_custom_${Date.now()}`;
    const newMeta: NotaSheetMeta = {
      id,
      title: trimmed,
      kind: "custom",
      kepada: newSheetKepada.trim() || "Penyedia Logistik Bahan Makanan",
      subtitle: `NOTA PESANAN — ${trimmed.toUpperCase()}`,
    };

    const dayNums = mode === "gabungan" ? selectedDays : [selectedDay];
    const allRaw = computeRawMergedIngredients(dayNums);
    let initialRows: NotaItem[];
    if (newSheetTemplate === "bahan_baku") {
      initialRows = padNotaItems(
        allRaw.filter((it) => !isBumbuIngredient(it.nama)),
        18
      );
    } else if (newSheetTemplate === "bumbu") {
      initialRows = padNotaItems(
        allRaw.filter((it) => isBumbuIngredient(it.nama)),
        18
      );
    } else {
      initialRows = padNotaItems([], 18);
    }

    const key = getStorageKey(id, mode, selectedDay, selectedDays);
    localStorage.setItem(key, JSON.stringify(initialRows));

    setSheets((prev) => [...prev, newMeta]);
    setActiveSheetId(id);
    setNewSheetTitle("");
    setNewSheetKepada("Penyedia Logistik Bahan Makanan");
    setNewSheetTemplate("blank");
    setShowAddSheetModal(false);
  };

  const handleUpdateActiveSheetMeta = (patch: Partial<NotaSheetMeta>) => {
    setSheets((prev) =>
      prev.map((s) => (s.id === activeSheet.id ? { ...s, ...patch } : s))
    );
  };

  const handleSaveSheetRename = (id: string) => {
    const trimmed = renameDraft.trim();
    if (trimmed) {
      setSheets((prev) =>
        prev.map((s) =>
          s.id === id
            ? { ...s, title: trimmed, subtitle: `NOTA PESANAN — ${trimmed.toUpperCase()}` }
            : s
        )
      );
    }
    setRenamingSheetId(null);
    setRenameDraft("");
  };

  const handleConfirmDeleteSheet = () => {
    if (!confirmDeleteSheet) return;
    const delId = confirmDeleteSheet.id;
    setSheets((prev) => prev.filter((s) => s.id !== delId));
    if (activeSheetId === delId) {
      setActiveSheetId("sheet_otomatis");
    }
    setConfirmDeleteSheet(null);
  };

  // Row edit handlers (if editing on Nota Otomatis while autoModeFix is on, gently disengage autoModeFix so user edits stick)
  const ensureEditableMode = () => {
    if (activeSheet.kind === "otomatis" && autoModeFix) {
      setAutoModeFix(false);
    }
  };

  const handleEditItem = (id: string, field: keyof NotaItem, value: any) => {
    ensureEditableMode();
    setItems((prevItems) =>
      prevItems.map((it) => {
        if (it.id !== id) return it;
        let parsedValue = value;
        if (field === "jumlah") {
          if (typeof value === "string") {
            parsedValue = value.endsWith(".") ? value : value === "" ? 0 : Number(value);
          } else {
            parsedValue = Number(value);
          }
        } else if (field === "hargaSatuan") {
          parsedValue = value === "" ? 0 : Number(value);
        }
        return { ...it, [field]: parsedValue };
      })
    );
  };

  const handleUnitTypeChange = (
    id: string,
    type: "foodcost" | "kg" | "potong" | "ekor" | "custom",
    customUnitValue?: string
  ) => {
    ensureEditableMode();
    setItems((prevItems) =>
      prevItems.map((it) => {
        if (it.id !== id) return it;
        const updated = { ...it, selectedUnitType: type };
        if (customUnitValue !== undefined) updated.customUnit = customUnitValue;

        if (type === "foodcost") {
          updated.satuan = it.defaultUnit || "Kg";
          updated.jumlah = parseFloat((it.sumButir || 0).toFixed(1));
        } else if (type === "kg") {
          updated.satuan = "Kg";
          updated.jumlah = parseFloat((it.sumKg || 0).toFixed(1));
        } else if (type === "potong") {
          updated.satuan = "Potong";
          updated.jumlah = parseFloat((it.sumPotong || 0).toFixed(1));
        } else if (type === "ekor") {
          updated.satuan = "Ekor";
          updated.jumlah = parseFloat((it.sumEkor || 0).toFixed(1));
        } else {
          updated.satuan =
            customUnitValue !== undefined ? customUnitValue : it.customUnit || "Kg";
        }
        return updated;
      })
    );
  };

  const handleAddItem = () => {
    ensureEditableMode();
    const newItem: NotaItem = {
      id: `manual_${Date.now()}`,
      nama: "",
      jumlah: 0,
      satuan: "",
      hargaSatuan: 0,
      sumKg: 0,
      sumPotong: 0,
      sumEkor: 0,
      sumButir: 0,
      defaultUnit: "",
      selectedUnitType: "custom",
      customUnit: "",
    };
    setItems((prev) => [...prev, newItem]);
  };

  const handleQuickAddBumbu = (preset: { nama: string; satuan: string; harga: number }) => {
    ensureEditableMode();
    setItems((prev) => {
      // Fill first empty padded row if available, else append
      const emptyIdx = prev.findIndex((r) => !r.nama || r.nama.trim() === "");
      const newRow: NotaItem = {
        id: `bumbu_${Date.now()}_${Math.random().toString(36).slice(2, 5)}`,
        nama: preset.nama,
        jumlah: 1,
        satuan: preset.satuan,
        hargaSatuan: preset.harga,
        sumKg: 1,
        sumPotong: 0,
        sumEkor: 0,
        sumButir: 1,
        defaultUnit: preset.satuan,
        selectedUnitType: "custom",
        customUnit: preset.satuan,
      };
      if (emptyIdx >= 0) {
        const next = [...prev];
        next[emptyIdx] = newRow;
        return next;
      }
      return [...prev, newRow];
    });
  };

  const handleDeleteItem = (id: string) => {
    ensureEditableMode();
    setItems((prev) => prev.filter((it) => it.id !== id));
  };

  const handleMergeDuplicates = () => {
    ensureEditableMode();
    const aggregated: Record<string, NotaItem> = {};
    items.forEach((item) => {
      if (!item.nama || item.nama.trim() === "") return;
      const key = item.nama.trim().toLowerCase();
      if (!aggregated[key]) {
        aggregated[key] = { ...item, nama: item.nama.trim() };
      } else {
        aggregated[key].jumlah = parseFloat((aggregated[key].jumlah + item.jumlah).toFixed(1));
        aggregated[key].sumKg = parseFloat(
          ((aggregated[key].sumKg || 0) + (item.sumKg || 0)).toFixed(1)
        );
        aggregated[key].sumPotong = parseFloat(
          ((aggregated[key].sumPotong || 0) + (item.sumPotong || 0)).toFixed(1)
        );
        aggregated[key].sumEkor = parseFloat(
          ((aggregated[key].sumEkor || 0) + (item.sumEkor || 0)).toFixed(1)
        );
        aggregated[key].sumButir = parseFloat(
          ((aggregated[key].sumButir || 0) + (item.sumButir || 0)).toFixed(1)
        );
        if (item.hargaSatuan > aggregated[key].hargaSatuan) {
          aggregated[key].hargaSatuan = item.hargaSatuan;
        }
      }
    });
    setItems(padNotaItems(Object.values(aggregated), 18));
  };

  const totalBelanja = useMemo(
    () => items.reduce((acc, it) => acc + calculateRowTotal(it), 0),
    [items]
  );

  const filledRowCount = useMemo(
    () => items.filter((r) => r.nama && r.nama.trim() !== "").length,
    [items]
  );

  const sheetSlug = activeSheet.title.replace(/[^A-Za-z0-9]+/g, "_");
  const periodSlug = mode === "gabungan" ? "Gabungan" : `Hari_${selectedDay}`;

  const handlePrint = () => {
    printElementSmart("print-area-nota-logistik", {
      title: `${activeSheet.title} — ${
        mode === "harian" ? `Hari Ke-${selectedDay}` : `Gabungan ${selectedDays.length} Hari`
      }`,
      paperSize,
      orientation: "auto",
    });
  };

  const handleExportPdf = async () => {
    await exportToPDF("print-area-nota-logistik", `${sheetSlug}_${periodSlug}`, 100, undefined, {
      paperSize,
      orientation: "auto",
    });
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const sheetData = [
      [kopLine1],
      [kopLine2],
      [kopLine3],
      [kopLine4],
      [],
      [activeSheet.subtitle],
      [],
      [`Dari : ${dari}`],
      [`Kepada : ${activeSheet.kepada}`],
      [`Alamat : ${alamat}`],
      [`Tanggal : ${effectiveTanggal}`],
      [],
      ["No.", "Uraian Jenis Bahan Makanan", "Jumlah", "Satuan", "Harga", "Total"],
    ];

    items.forEach((row, idx) => {
      const isEmpty = !row.nama || row.nama.trim() === "";
      if (isEmpty) {
        sheetData.push(["", "", "", "", "", ""]);
      } else {
        const rowTotal = calculateRowTotal(row);
        sheetData.push([
          String(idx + 1),
          row.nama,
          row.jumlah === 0 ? "" : formatJumlahDisplay(row.jumlah),
          row.satuan,
          row.hargaSatuan === 0 ? "Rp -" : `Rp ${formatNumber(row.hargaSatuan)}`,
          rowTotal === 0 ? "Rp -" : `Rp ${formatNumber(rowTotal)}`,
        ]);
      }
    });

    sheetData.push(["TOTAL BELANJA", "", "", "", "", `Rp ${formatNumber(totalBelanja)}`]);
    sheetData.push([]);
    sheetData.push(["", "", "", "Mengetahui,"]);
    sheetData.push([
      "",
      "",
      "",
      profile.ttdConfig?.jabatanKepala || "Kepala Satuan Pelayanan Pemenuhan Gizi",
    ]);
    sheetData.push([]);
    sheetData.push([]);
    sheetData.push(["", "", "", profile.namaKepala || ""]);

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws["!cols"] = [
      { wch: 8 },
      { wch: 38 },
      { wch: 14 },
      { wch: 12 },
      { wch: 18 },
      { wch: 20 },
    ];
    XLSX.utils.book_append_sheet(wb, ws, activeSheet.title.slice(0, 31));
    XLSX.writeFile(wb, `${sheetSlug}_${periodSlug}.xlsx`);
  };

  const activeBumbuDisplay =
    mode === "harian"
      ? foodCostDays.find((fd) => fd.hariKe === selectedDay)?.bumbuPct ?? bumbuPct
      : bumbuPct;

  return (
    <div className="space-y-5 font-sans" id="nota-pesanan-logistik-tab-container">
      {/* Dynamic Print CSS */}
      <style>{`
        @media print {
          body {
            background-color: white !important;
            color: black !important;
            font-family: Arial, Helvetica, sans-serif !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          #nota-pesanan-logistik-tab-container .no-print {
            display: none !important;
          }
          .print-container {
            width: 100% !important;
            max-width: 100% !important;
            min-height: auto !important;
            padding: 0 !important;
            margin: 0 auto !important;
            box-shadow: none !important;
            border: none !important;
            background: white !important;
          }
          .print-container table {
            border-collapse: collapse !important;
            width: 100% !important;
          }
          .print-container th, .print-container td {
            border: 1px solid black !important;
            padding: 4px 6px !important;
            font-size: 10.5px !important;
            line-height: 1.2 !important;
          }
          .print-container input {
            border: none !important;
            background: transparent !important;
            color: black !important;
          }
        }
      `}</style>

      {/* CONTROL PANEL (NO-PRINT) */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4 no-print">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <h3 className="text-base font-extrabold text-slate-900 flex items-center gap-2">
              <FileText className="w-5 h-5 text-orange-600" />
              Nota Pesanan Logistik (Otomatis &amp; Pisah Lembaran Bahan Baku / Bumbu)
            </h3>
            <p className="text-xs text-slate-500">
              Identitas, KOP surat, tanggal periode, dan penandatangan (TTD) tersinkron otomatis dari <strong>Profil SPPG</strong>. Pisahkan nota Bahan Baku Utama, Bahan Baku Bumbu, tambah lembaran kustom, atau aktifkan <strong>Nota Otomatis</strong> saat perencanaan sudah fix.
            </p>
          </div>

          {/* Mode Harian vs Gabungan */}
          <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 self-start lg:self-center shrink-0">
            <button
              type="button"
              onClick={() => setMode("harian")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition cursor-pointer ${
                mode === "harian"
                  ? "bg-white text-indigo-600 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              Mode Harian
            </button>
            <button
              type="button"
              onClick={() => setMode("gabungan")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition cursor-pointer ${
                mode === "gabungan"
                  ? "bg-white text-indigo-600 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              Gabungan {totalDays} Hari
            </button>
          </div>
        </div>

        {/* MULTI-SHEET BAR: NOTA OTOMATIS | BAHAN BAKU | BUMBU | + TAMBAH LEMBARAN */}
        <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200/80 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider flex items-center gap-1.5">
              <Layers className="w-3.5 h-3.5 text-orange-600" />
              Pilih atau Tambah Lembaran Nota Pesanan
            </span>

            {/* Auto Mode Fix Switch */}
            <label className="inline-flex items-center gap-2 px-3 py-1 rounded-lg bg-white border border-emerald-200 text-xs font-bold text-emerald-800 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={autoModeFix}
                onChange={(e) => {
                  const checked = e.target.checked;
                  setAutoModeFix(checked);
                  if (checked) {
                    setActiveSheetId("sheet_otomatis");
                  }
                }}
                className="w-3.5 h-3.5 accent-emerald-600 cursor-pointer"
              />
              <Zap className="w-3.5 h-3.5 text-emerald-600" />
              <span>Gunakan Nota Pesanan Otomatis (Food Cost Fix)</span>
            </label>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {sheets.map((sheet) => {
              const isActive = sheet.id === activeSheet.id;
              const isRenaming = renamingSheetId === sheet.id;
              return (
                <div
                  key={sheet.id}
                  className={`inline-flex items-center gap-1.5 rounded-xl border transition ${
                    isActive
                      ? "bg-slate-900 border-slate-900 text-white shadow-xs"
                      : "bg-white border-slate-200 text-slate-700 hover:border-slate-300"
                  }`}
                >
                  {isRenaming ? (
                    <div className="flex items-center gap-1 px-2 py-1">
                      <input
                        type="text"
                        value={renameDraft}
                        onChange={(e) => setRenameDraft(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") handleSaveSheetRename(sheet.id);
                          if (e.key === "Escape") setRenamingSheetId(null);
                        }}
                        autoFocus
                        className="w-36 px-1.5 py-0.5 text-xs rounded bg-white text-slate-900 border border-slate-300 focus:outline-none"
                      />
                      <button
                        type="button"
                        onClick={() => handleSaveSheetRename(sheet.id)}
                        className="p-1 text-emerald-400 hover:text-emerald-300 cursor-pointer"
                      >
                        <Check className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ) : (
                    <>
                      <button
                        type="button"
                        onClick={() => setActiveSheetId(sheet.id)}
                        className="px-3 py-2 text-xs font-bold cursor-pointer flex items-center gap-1.5"
                      >
                        {sheet.kind === "otomatis" && (
                          <Zap className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
                        )}
                        <span>{sheet.title}</span>
                      </button>

                      {sheet.kind === "custom" && (
                        <div className="flex items-center pr-1.5 gap-0.5">
                          <button
                            type="button"
                            onClick={() => {
                              setRenamingSheetId(sheet.id);
                              setRenameDraft(sheet.title);
                            }}
                            title="Ubah nama lembaran"
                            className={`p-1 rounded hover:bg-white/15 cursor-pointer ${
                              isActive ? "text-slate-300 hover:text-white" : "text-slate-400 hover:text-slate-700"
                            }`}
                          >
                            <Edit3 className="w-3 h-3" />
                          </button>
                          <button
                            type="button"
                            onClick={() => setConfirmDeleteSheet(sheet)}
                            title="Hapus lembaran"
                            className={`p-1 rounded hover:bg-rose-500/20 cursor-pointer ${
                              isActive ? "text-rose-300 hover:text-rose-200" : "text-rose-500"
                            }`}
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        </div>
                      )}
                    </>
                  )}
                </div>
              );
            })}

            <button
              type="button"
              onClick={() => setShowAddSheetModal(true)}
              className="px-3 py-2 rounded-xl border border-dashed border-orange-400 bg-orange-50 hover:bg-orange-100 text-orange-800 text-xs font-bold inline-flex items-center gap-1.5 transition cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Tambah Lembaran Nota</span>
            </button>
          </div>
        </div>

        {/* DAY SELECTOR & ACTIONS */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pt-2 border-t border-slate-100">
          <div className="flex flex-wrap items-center gap-3">
            {mode === "harian" ? (
              <div className="flex items-center gap-2 bg-slate-50 p-1.5 rounded-xl border border-slate-200">
                <span className="text-xs text-slate-600 font-bold px-1.5">Pilih Hari:</span>
                <select
                  id="select-nota-logistik-day"
                  value={selectedDay}
                  onChange={(e) => setSelectedDay(Number(e.target.value))}
                  className="bg-white border border-slate-200 rounded-lg text-xs sm:text-sm px-3 py-1 font-mono font-bold text-indigo-700 focus:ring-1 focus:ring-indigo-500"
                >
                  {Array.from({ length: totalDays }).map((_, i) => (
                    <option key={i} value={i + 1}>
                      Hari Ke-{i + 1}
                      {profile?.periodeDates?.[i]
                        ? ` (${formatIndonesianDate(profile.periodeDates[i])})`
                        : ""}
                    </option>
                  ))}
                </select>
              </div>
            ) : (
              <div className="space-y-1.5">
                <div className="flex items-center gap-3 text-xs">
                  <span className="font-bold text-slate-700">Pilih Hari Gabungan:</span>
                  <button
                    type="button"
                    onClick={() =>
                      setSelectedDays(Array.from({ length: totalDays }, (_, i) => i + 1))
                    }
                    className="text-indigo-600 font-bold hover:underline cursor-pointer"
                  >
                    Semua
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedDays([])}
                    className="text-rose-500 font-bold hover:underline cursor-pointer"
                  >
                    Kosongkan
                  </button>
                </div>
                <div className="flex flex-wrap gap-1.5">
                  {Array.from({ length: totalDays }).map((_, i) => {
                    const dNum = i + 1;
                    const isSelected = selectedDays.includes(dNum);
                    return (
                      <button
                        key={i}
                        type="button"
                        onClick={() =>
                          setSelectedDays((prev) =>
                            isSelected
                              ? prev.filter((x) => x !== dNum)
                              : [...prev, dNum].sort((a, b) => a - b)
                          )
                        }
                        className={`px-2.5 py-1 text-xs font-mono font-bold rounded-lg border transition cursor-pointer ${
                          isSelected
                            ? "bg-indigo-50 border-indigo-300 text-indigo-700"
                            : "bg-white border-slate-200 text-slate-500 hover:bg-slate-50"
                        }`}
                      >
                        H-{dNum}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() =>
                setPreviewWidth((w) => (w === "full" ? "standard" : "full"))
              }
              className="px-3 py-2 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
              title="Ubah lebar pratinjau kertas"
            >
              {previewWidth === "full" ? (
                <>
                  <Minimize2 className="w-3.5 h-3.5" />
                  <span>A4 Standar</span>
                </>
              ) : (
                <>
                  <Maximize2 className="w-3.5 h-3.5" />
                  <span>Lebar Penuh</span>
                </>
              )}
            </button>

            <button
              type="button"
              onClick={() => loadSheetItems(true)}
              className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl border border-slate-200 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
              title="Tarik ulang bahan otomatis dari Food Cost"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Sinkron Food Cost</span>
            </button>

            <button
              type="button"
              onClick={handleAddItem}
              className="px-3 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl border border-indigo-200 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Tambah Baris</span>
            </button>

            <button
              type="button"
              onClick={handleMergeDuplicates}
              className="px-3 py-2 bg-teal-50 hover:bg-teal-100 text-teal-700 rounded-xl border border-teal-200 text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Combine className="w-3.5 h-3.5" />
              <span>Gabung Duplikat</span>
            </button>

            <button
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="px-3 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Eye className="w-3.5 h-3.5" />
              <span>Pratinjau</span>
            </button>

            <button
              type="button"
              onClick={handleExportPdf}
              className="px-3 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <FileDown className="w-3.5 h-3.5" />
              <span>Unduh PDF</span>
            </button>

            <button
              type="button"
              onClick={handleExportExcel}
              className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Excel</span>
            </button>

            <button
              type="button"
              onClick={handlePrint}
              className="px-3 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Printer className="w-3.5 h-3.5" />
              <span>Cetak</span>
            </button>
          </div>
        </div>

        {/* QUICK BUMBU ADDER BAR (ONLY SHOWN WHEN ON BUMBU SHEET) */}
        {activeSheet.kind === "bumbu" && (
          <div className="p-3.5 rounded-xl bg-amber-50/70 border border-amber-200 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2 text-xs font-bold text-amber-900">
                <Sparkles className="w-4 h-4 text-amber-600" />
                <span>Tambah Cepat Bahan Baku Bumbu &amp; Rempah Dapur</span>
              </div>
              <div className="text-xs font-bold tabular-nums text-amber-900 bg-white px-2.5 py-1 rounded-lg border border-amber-200">
                Alokasi Anggaran Bumbu ({activeBumbuDisplay}%):{" "}
                <span className="text-emerald-700">
                  Rp {formatThousandSeparator(summaryStats.totalBumbu10Nominal)}
                </span>{" "}
                • Belanja Nota Ini:{" "}
                <span
                  className={
                    totalBelanja > summaryStats.totalBumbu10Nominal &&
                    summaryStats.totalBumbu10Nominal > 0
                      ? "text-rose-600"
                      : "text-indigo-700"
                  }
                >
                  Rp {formatThousandSeparator(totalBelanja)}
                </span>
              </div>
            </div>
            <div className="flex flex-wrap gap-1.5">
              {QUICK_BUMBU_PRESETS.map((bp) => (
                <button
                  key={bp.nama}
                  type="button"
                  onClick={() => handleQuickAddBumbu(bp)}
                  className="px-2.5 py-1 rounded-lg bg-white hover:bg-amber-100 border border-amber-200 text-amber-950 text-[11px] font-semibold transition cursor-pointer inline-flex items-center gap-1"
                >
                  <Plus className="w-3 h-3 text-amber-600" />
                  <span>{bp.nama}</span>
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* RAB & REALISASI BELANJA KPI STRIP (NO-PRINT) */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 no-print">
        <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs flex flex-col justify-between">
          <div>
            <span className="text-[11px] font-bold text-emerald-700 uppercase tracking-wider block">
              Rencana Anggaran Biaya ({summaryStats.isCombined ? "RAB Gabungan" : "RAB Harian"})
            </span>
            <div className="mt-1.5 text-xl font-black text-slate-900 tabular-nums">
              Rp {formatThousandSeparator(summaryStats.totalRAB)}
            </div>
          </div>
          <div className="text-[11px] text-slate-500 font-medium mt-2.5 space-y-0.5 border-t border-slate-100 pt-2 tabular-nums">
            <div>
              Kecil: {summaryStats.totalPMKecil} PM × Rp{" "}
              {formatThousandSeparator(settings.porsiKecilHarga)} = Rp{" "}
              {formatThousandSeparator(summaryStats.rabPorsiKecilTotal)}
            </div>
            <div>
              Besar: {summaryStats.totalPMBesar} PM × Rp{" "}
              {formatThousandSeparator(settings.porsiBesarHarga)} = Rp{" "}
              {formatThousandSeparator(summaryStats.rabPorsiBesarTotal)}
            </div>
          </div>
        </div>

        <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between gap-2">
              <span className="text-[11px] font-bold text-indigo-700 uppercase tracking-wider">
                Realisasi Belanja Food Cost
              </span>
              <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-indigo-50 text-indigo-700 border border-indigo-200 tabular-nums">
                +{activeBumbuDisplay}% Bumbu
              </span>
            </div>
            <div className="mt-1.5 flex items-baseline justify-between gap-2">
              <span className="text-xl font-black text-slate-900 tabular-nums">
                Rp {formatThousandSeparator(summaryStats.totalActualCost)}
              </span>
              <span className="text-xs font-bold text-indigo-700 tabular-nums">
                Bumbu: Rp {formatThousandSeparator(summaryStats.totalBumbu10Nominal)}
              </span>
            </div>
          </div>
          <div className="text-[11px] text-slate-500 font-medium mt-2.5 space-y-0.5 border-t border-slate-100 pt-2 tabular-nums">
            <div>
              Kecil: Rp {formatThousandSeparator(summaryStats.porsiKecilActualCostTotal)} (Rp{" "}
              {formatThousandSeparator(summaryStats.porsiKecilPerPorsi)}/porsi)
            </div>
            <div>
              Besar: Rp {formatThousandSeparator(summaryStats.porsiBesarActualCostTotal)} (Rp{" "}
              {formatThousandSeparator(summaryStats.porsiBesarPerPorsi)}/porsi)
            </div>
          </div>
        </div>

        <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs flex flex-col justify-between">
          <div>
            <span
              className={`text-[11px] font-bold uppercase tracking-wider block ${
                summaryStats.sisaEfisiensiAnggaran >= 0 ? "text-teal-700" : "text-rose-700"
              }`}
            >
              {summaryStats.sisaEfisiensiAnggaran >= 0
                ? "Sisa / Efisiensi Anggaran"
                : "Defisit Anggaran"}
            </span>
            <div
              className={`mt-1.5 text-xl font-black tabular-nums ${
                summaryStats.sisaEfisiensiAnggaran >= 0 ? "text-teal-700" : "text-rose-600"
              }`}
            >
              {summaryStats.sisaEfisiensiAnggaran >= 0 ? "+" : ""}
              {formatRupiah(summaryStats.sisaEfisiensiAnggaran)}
            </div>
          </div>
          <div className="text-[11px] text-slate-500 font-medium mt-2.5 border-t border-slate-100 pt-2 flex items-center justify-between">
            <span>
              Lembar Aktif: <strong className="text-slate-800">{activeSheet.title}</strong>
            </span>
            <span className="tabular-nums font-bold text-slate-700">{filledRowCount} item</span>
          </div>
        </div>
      </div>

      {/* FULL-WIDTH PRINTABLE NOTA SHEET (GAMBAR 4 SIDEBAR COMPLETELY REMOVED) */}
      <div className="flex justify-center w-full" id="print-area-wrapper">
        <div
          id="print-area-nota-logistik"
          style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
          className={`print-container w-full ${
            previewWidth === "full" ? "max-w-full" : "max-w-[210mm]"
          } min-h-[297mm] bg-white text-black p-6 sm:p-[14mm] border border-slate-200 rounded-2xl shadow-sm relative flex flex-col gap-4 overflow-x-auto`}
        >
          {/* Centralized KOP Surat from Profil SPPG */}
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
            docKey="nota"
          />

          {/* Sheet Title Banner */}
          <div className="text-center -mt-1">
            <input
              type="text"
              value={activeSheet.subtitle}
              onChange={(e) => handleUpdateActiveSheetMeta({ subtitle: e.target.value })}
              className="w-full text-center text-sm font-bold uppercase tracking-wide text-black bg-transparent border-0 focus:outline-none focus:bg-slate-50 rounded py-0.5"
            />
          </div>

          {/* Inline Metadata Block (Dari & Alamat auto from Profil SPPG; Kepada & Tanggal editable inline) */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-1.5 text-xs font-medium my-1 leading-normal">
            <div className="space-y-1">
              <div className="flex items-start">
                <span className="w-20 inline-block font-bold shrink-0">Dari</span>
                <span className="mr-2">:</span>
                <span className="flex-1 font-semibold">{dari}</span>
              </div>
              <div className="flex items-center">
                <span className="w-20 inline-block font-bold shrink-0">Kepada</span>
                <span className="mr-2">:</span>
                <input
                  type="text"
                  value={activeSheet.kepada}
                  onChange={(e) => handleUpdateActiveSheetMeta({ kepada: e.target.value })}
                  className="flex-1 font-semibold bg-transparent border-b border-transparent hover:border-slate-300 focus:border-indigo-500 focus:outline-none py-0.5"
                  placeholder="Nama Supplier / Penyedia..."
                />
              </div>
            </div>

            <div className="space-y-1">
              <div className="flex items-start">
                <span className="w-20 inline-block font-bold shrink-0">Alamat</span>
                <span className="mr-2">:</span>
                <span className="flex-1 font-semibold">{alamat}</span>
              </div>
              <div className="flex items-center">
                <span className="w-20 inline-block font-bold shrink-0">Tanggal</span>
                <span className="mr-2">:</span>
                <input
                  type="text"
                  value={effectiveTanggal}
                  onChange={(e) => setTanggalOverride(e.target.value)}
                  className="flex-1 font-semibold bg-transparent border-b border-transparent hover:border-slate-300 focus:border-indigo-500 focus:outline-none py-0.5"
                  placeholder="Otomatis dari Profil SPPG..."
                />
              </div>
            </div>
          </div>

          {/* Main Order Table */}
          <div className="w-full">
            <table className="w-full text-[12px] border-collapse border border-black leading-relaxed">
              <thead>
                <tr className="bg-[#92D050] text-black font-bold border-b border-black">
                  <th className="border border-black p-2 text-center w-12 uppercase">No.</th>
                  <th className="border border-black p-2 text-left uppercase">
                    Uraian Jenis Bahan Makanan
                  </th>
                  <th className="border border-black p-2 text-center w-24 uppercase">Jumlah</th>
                  <th className="border border-black p-2 text-center w-36 uppercase">Satuan</th>
                  <th className="border border-black p-2 text-center w-36 uppercase">Harga</th>
                  <th className="border border-black p-2 text-center w-40 uppercase">Total</th>
                </tr>
              </thead>
              <tbody>
                {items.map((row, idx) => {
                  const rowTotal = calculateRowTotal(row);
                  return (
                    <tr key={row.id} className="hover:bg-slate-50/70 print:hover:bg-transparent group">
                      <td className="border border-black text-center py-1 font-bold tabular-nums">
                        {idx + 1}
                      </td>

                      <td className="border border-black px-2 py-0.5 relative">
                        <input
                          type="text"
                          value={row.nama}
                          placeholder="Ketik bahan..."
                          onChange={(e) => handleEditItem(row.id, "nama", e.target.value)}
                          className="w-full bg-transparent border-0 p-0 focus:outline-none font-medium text-slate-900 print:text-black placeholder:text-slate-300 print:placeholder:text-transparent pr-6"
                        />
                        <button
                          type="button"
                          onClick={() => handleDeleteItem(row.id)}
                          className="absolute right-1 top-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100 p-0.5 hover:bg-rose-100 text-rose-600 rounded transition no-print cursor-pointer"
                          title="Hapus baris ini"
                        >
                          <Trash2 className="w-3 h-3" />
                        </button>
                      </td>

                      <td className="border border-black px-1.5 py-0.5 text-center font-semibold tabular-nums">
                        <input
                          type="text"
                          value={
                            editingId === row.id
                              ? editingVal
                              : row.jumlah === 0
                              ? ""
                              : formatJumlahDisplay(row.jumlah)
                          }
                          placeholder="-"
                          onFocus={() => {
                            setEditingId(row.id);
                            setEditingVal(
                              row.jumlah === 0 ? "" : formatJumlahDisplay(row.jumlah)
                            );
                          }}
                          onBlur={() => {
                            setEditingId(null);
                            const num = editingVal === "" ? 0 : parseFloat(editingVal);
                            handleEditItem(
                              row.id,
                              "jumlah",
                              isNaN(num) ? 0 : parseFloat((Math.round(num * 10) / 10).toFixed(1))
                            );
                          }}
                          onChange={(e) => setEditingVal(e.target.value.replace(/[^0-9.]/g, ""))}
                          className="w-full bg-transparent border-0 p-0 text-center focus:outline-none font-bold placeholder:text-slate-300 print:placeholder:text-transparent text-xs tabular-nums"
                        />
                      </td>

                      <td className="border border-black px-1.5 py-0.5">
                        <div className="flex items-center justify-center gap-1">
                          <span className="font-semibold text-center hidden print:inline">
                            {row.satuan}
                          </span>
                          <div className="flex items-center gap-1 no-print w-full justify-center">
                            <input
                              type="text"
                              placeholder="Satuan"
                              value={row.satuan || ""}
                              onChange={(e) => {
                                handleEditItem(row.id, "satuan", e.target.value);
                                handleEditItem(row.id, "customUnit", e.target.value);
                                handleEditItem(row.id, "selectedUnitType", "custom");
                              }}
                              className="w-16 bg-transparent border border-transparent hover:border-slate-200 focus:border-indigo-500 rounded px-1 py-0.5 text-center focus:outline-none font-semibold text-slate-900 text-xs"
                            />
                            <select
                              value={row.selectedUnitType || "foodcost"}
                              onChange={(e) => handleUnitTypeChange(row.id, e.target.value as any)}
                              className="text-[10px] py-0.5 px-1 border border-slate-200 rounded bg-white text-slate-700 cursor-pointer font-semibold"
                            >
                              <option value="foodcost">Asli ({row.defaultUnit || "Kg"})</option>
                              <option value="kg">Kg</option>
                              <option value="ekor">Ekor</option>
                              <option value="potong">Potong</option>
                              <option value="custom">Kustom</option>
                            </select>
                          </div>
                        </div>
                      </td>

                      <td className="border border-black px-1.5 py-0.5 tabular-nums">
                        <div className="flex justify-between items-center w-full gap-1">
                          <span className="text-slate-400 print:text-black font-bold">Rp</span>
                          <div className="flex-1 flex items-center justify-end gap-1">
                            <input
                              type="text"
                              value={
                                row.hargaSatuan === 0
                                  ? ""
                                  : Number(row.hargaSatuan).toLocaleString("id-ID")
                              }
                              placeholder="-"
                              onChange={(e) => {
                                const rawVal = e.target.value.replace(/\D/g, "");
                                handleEditItem(
                                  row.id,
                                  "hargaSatuan",
                                  rawVal ? parseInt(rawVal, 10) : 0
                                );
                              }}
                              className="w-full bg-transparent border-0 p-0 text-right focus:outline-none font-bold placeholder:text-slate-400 print:placeholder:text-transparent text-xs text-slate-900 tabular-nums"
                            />
                            <PriceCalculatorPopover
                              initialValue={row.hargaSatuan || 0}
                              onApply={(val) => handleEditItem(row.id, "hargaSatuan", val)}
                              placeholder="Hitung Harga Satuan"
                              className="no-print"
                              ingredientName={row.nama}
                            />
                          </div>
                        </div>
                      </td>

                      <td className="border border-black px-2 py-0.5 font-bold tabular-nums bg-slate-50/30">
                        <div className="flex justify-between w-full">
                          <span>Rp</span>
                          <span>{rowTotal === 0 ? "-" : formatNumber(rowTotal)}</span>
                        </div>
                      </td>
                    </tr>
                  );
                })}

                {/* TOTAL BELANJA */}
                <tr className="bg-slate-200 font-bold text-black border-t-2 border-black">
                  <td
                    colSpan={4}
                    className="border border-black px-4 py-2 text-center text-xs font-extrabold uppercase tracking-wider"
                  >
                    TOTAL BELANJA
                  </td>
                  <td className="border border-black px-2 py-2 text-xs font-bold text-center text-slate-500">
                    Rp -
                  </td>
                  <td className="border border-black px-3 py-2 text-xs font-extrabold tabular-nums">
                    <div className="flex justify-between w-full">
                      <span>Rp</span>
                      <span>{formatNumber(totalBelanja)}</span>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Centralized Auto Signature & ID SPPG Block from Profil SPPG */}
          <AutoSignatureBlock
            profile={profile}
            docKey="nota"
            dateText={effectiveTanggal}
          />
        </div>
      </div>

      {/* ADD NEW NOTA SHEET MODAL */}
      {showAddSheetModal && (
        <div className="fixed inset-0 z-[99999] bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 no-print">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h4 className="text-base font-bold text-slate-900">
                  Tambah Lembaran Nota Pesanan Baru
                </h4>
                <p className="text-xs text-slate-500 mt-0.5">
                  Buat lembaran nota terpisah untuk supplier tertentu, bahan bumbu tambahan, atau kategori khusus.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setShowAddSheetModal(false)}
                className="p-1 text-slate-400 hover:text-slate-700 rounded-lg cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleCreateSheet} className="space-y-3.5">
              <div className="space-y-1">
                <label className="text-xs font-bold text-slate-600 block">
                  Nama Lembaran Nota
                </label>
                <input
                  type="text"
                  required
                  value={newSheetTitle}
                  onChange={(e) => setNewSheetTitle(e.target.value)}
                  placeholder="Contoh: Nota Sayur & Buah / Nota Supplier Protein"
                  className="w-full px-3.5 py-2 border border-slate-200 rounded-xl text-xs font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
                />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-bold text-slate-600 block">
                  Tujuan / Kepada (Supplier)
                </label>
                <input
                  type="text"
                  value={newSheetKepada}
                  onChange={(e) => setNewSheetKepada(e.target.value)}
                  placeholder="Contoh: CV Sumber Pangan / Toko Bumbu Jaya"
                  className="w-full px-3.5 py-2 border border-slate-200 rounded-xl text-xs font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
                />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-bold text-slate-600 block">
                  Isi Awal Lembaran
                </label>
                <div className="grid grid-cols-3 gap-2">
                  {[
                    { id: "blank" as const, label: "Kosong (Manual)" },
                    { id: "bahan_baku" as const, label: "Salin Bahan Baku" },
                    { id: "bumbu" as const, label: "Salin Bumbu" },
                  ].map((opt) => (
                    <button
                      key={opt.id}
                      type="button"
                      onClick={() => setNewSheetTemplate(opt.id)}
                      className={`px-2 py-2 rounded-xl border text-[11px] font-bold transition cursor-pointer ${
                        newSheetTemplate === opt.id
                          ? "bg-orange-600 border-orange-600 text-white"
                          : "bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100"
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>

              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setShowAddSheetModal(false)}
                  className="px-4 py-2 rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-50 text-xs font-bold cursor-pointer"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 rounded-xl bg-orange-600 hover:bg-orange-700 text-white text-xs font-bold cursor-pointer"
                >
                  Buat Lembaran Nota
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* CONFIRM DELETE CUSTOM SHEET MODAL */}
      {confirmDeleteSheet && (
        <div className="fixed inset-0 z-[99999] bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 no-print">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h4 className="text-base font-bold text-slate-900">Hapus Lembaran Nota?</h4>
                <p className="text-xs text-slate-500 mt-1">
                  Lembaran <strong>"{confirmDeleteSheet.title}"</strong> akan dihapus dari daftar nota pesanan.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setConfirmDeleteSheet(null)}
                className="p-1 text-slate-400 hover:text-slate-700 rounded-lg cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setConfirmDeleteSheet(null)}
                className="px-4 py-2 rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-50 text-xs font-bold cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmDeleteSheet}
                className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold cursor-pointer"
              >
                Ya, Hapus Lembaran
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Print Preview Modal */}
      <PrintPreviewModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title={`Pratinjau Cetak: ${activeSheet.title} (${
          mode === "harian" ? `Hari Ke-${selectedDay}` : `Gabungan ${selectedDays.length} Hari`
        })`}
        elementId="print-area-nota-logistik"
        filename={`${sheetSlug}_${periodSlug}`}
        defaultScale={100}
        defaultPaperSize={paperSize as "A4" | "F4"}
        defaultOrientation="auto"
      />
    </div>
  );
}
