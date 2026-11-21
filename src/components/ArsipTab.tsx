/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useEffect, useMemo, useState } from "react";
import { onAuthStateChanged } from "firebase/auth";
import {
  Archive,
  Calendar,
  Check,
  CheckCircle2,
  ChefHat,
  ChevronDown,
  ChevronRight,
  Clock,
  CloudUpload,
  DollarSign,
  Download,
  ExternalLink,
  FileText,
  Folder,
  FolderOpen,
  FolderPlus,
  Loader2,
  Package,
  Save,
  Search,
  ShoppingCart,
  Trash2,
  Users,
  X,
  AlertCircle,
  ClipboardCheck,
  Printer,
  ArrowUpRight,
} from "lucide-react";
import {
  FoodCostDay,
  HariPM,
  MasterMenu,
  SavedDocArchive,
  SekolahPM,
  SPPGProfile,
  TigaBPM,
  TKPIItem,
  UnifiedKopConfig,
  DEFAULT_UNIFIED_KOP_CONFIG,
} from "../types";
import { printElementSmart } from "../lib/printUtils";
import { calculateDay, createTkpiMap, formatRupiah } from "../utils/calc";
import {
  buildCategoryLabelFromIds,
  composeMenuName,
  ensureMasterMenuPlans,
  formatIndonesianDate,
  formatIndonesianMonthYear,
  resolveSasaranPmFromTable2,
} from "../utils/menuSchedule";
import {
  ArchiveCategory,
  ArchiveFoodCostSection,
  ArchiveMenuPlanRecord,
  ArchiveNotaItemRecord,
  ArchivePMRecord,
  buildArchivePdfFilename,
  DailyArchiveEntry,
  generateArchivePdfBlob,
} from "../lib/archivePdf";
import { auth } from "../lib/firebase";
import {
  cancelGoogleDriveConnect,
  clearDriveSession,
  connectGoogleDrive,
  DRIVE_OAUTH_AUDIENCE_URL,
  DriveAuthErrorCode,
  DriveConnectionError,
  DriveSessionInfo,
  DriveUploadResult,
  findOrCreateDriveFolder,
  getDriveSession,
  sanitizeDriveName,
  uploadPdfToDriveFolder,
} from "../lib/googleDrive";

interface ArsipTabProps {
  profile: SPPGProfile;
  harianPM: HariPM[];
  masterMenu: MasterMenu;
  foodCostDays: FoodCostDay[];
  tkpiList: TKPIItem[];
  pmSettings: any;
  savedPeriods: any[];
  setSavedPeriods?: (updated: any[]) => void;
  loadPeriod?: (period: any) => void;
  activeSekolahPM?: SekolahPM[];
  activeTigaBPM?: TigaBPM[];
  kopLines: [string, string, string, string];
  leftLogo?: string;
  rightLogo?: string;
  kopConfig?: UnifiedKopConfig;
  paperSize?: "A4" | "F4";
  bumbuPct?: number;
  userEmail?: string | null;
  isDemo?: boolean;
  docArchives?: SavedDocArchive[];
  setDocArchives?: React.Dispatch<React.SetStateAction<SavedDocArchive[]>>;
  onOpenDocArchive?: (archive: SavedDocArchive) => void;
}

interface DriveArchiveLogItem {
  id: string;
  sentAt: string;
  driveEmail: string;
  monthLabel: string;
  folderName: string;
  folderLink?: string;
  fileCount: number;
  files: Array<{ name: string; webViewLink?: string; menuFolder: string }>;
}

const ALL_CATEGORIES: Array<{
  id: ArchiveCategory;
  label: string;
  shortLabel: string;
  colorClass: string;
  badgeClass: string;
}> = [
  {
    id: "jumlah_pm",
    label: "Arsip Jumlah PM",
    shortLabel: "Jumlah PM",
    colorClass: "text-cyan-700",
    badgeClass: "bg-cyan-50 text-cyan-700 border-cyan-200",
  },
  {
    id: "menu",
    label: "Arsip Menu",
    shortLabel: "Menu",
    colorClass: "text-emerald-700",
    badgeClass: "bg-emerald-50 text-emerald-700 border-emerald-200",
  },
  {
    id: "food_cost",
    label: "Arsip Food Cost",
    shortLabel: "Food Cost",
    colorClass: "text-amber-700",
    badgeClass: "bg-amber-50 text-amber-700 border-amber-200",
  },
  {
    id: "nota_pesanan",
    label: "Arsip Nota Pesanan",
    shortLabel: "Nota Pesanan",
    colorClass: "text-indigo-700",
    badgeClass: "bg-indigo-50 text-indigo-700 border-indigo-200",
  },
];

const DEFAULT_BESAR_IDS = [
  "sd_kelas_4_6",
  "smp_mts_smplb",
  "sma_smk_ma",
  "pendidik",
  "tenaga_kependidikan",
  "ibu_hamil",
  "ibu_menyusui",
];

const DEFAULT_KECIL_IDS = [
  "tk_paud_lb",
  "sd_kelas_1_3",
  "anak_balita",
  "anak_balita_13_59",
  "balita_6_11",
];

const guessSatuanNota = (name: string): string => {
  const lower = (name || "").toLowerCase();
  if (lower.includes("minyak")) return "Liter";
  if (lower.includes("susu") && !lower.includes("bubuk")) return "Liter";
  if (lower.includes("kecap") || lower.includes("saus") || lower.includes("tiram")) return "Pcs";
  if (lower.includes("telur")) return "Butir";
  if (lower.includes("tahu") || lower.includes("tempe")) return "Pcs";
  if (lower.includes("garam") || lower.includes("royco") || lower.includes("masako") || lower.includes("bumbu")) {
    return "Bks";
  }
  return "Kg";
};

function computePMForDay(harianPM: HariPM[], hariKe: number): ArchivePMRecord {
  const dayPM = harianPM.find((h) => h.hariKe === hariKe) || harianPM[0] || { sasaran: [] };
  let totalKecil = 0;
  let totalBesar = 0;
  let totalAlergi = 0;

  const rows = (dayPM.sasaran || []).map((s) => {
    const pk = Number(s.porsiKecil) || 0;
    const pb = Number(s.porsiBesar) || 0;
    const ak = Number(s.alergiKecil) || 0;
    const ab = Number(s.alergiBesar) || 0;
    totalKecil += pk;
    totalBesar += pb;
    totalAlergi += ak + ab;
    return {
      label: s.label,
      porsiKecil: pk,
      porsiBesar: pb,
      alergiKecil: ak,
      alergiBesar: ab,
      total: pk + pb,
    };
  });

  return {
    totalPM: totalKecil + totalBesar,
    totalKecil,
    totalBesar,
    totalAlergi,
    rows,
  };
}

function buildFoodCostAndNotaForDay(
  dayNum: number,
  dateStr: string,
  harianPM: HariPM[],
  foodCostDays: FoodCostDay[],
  tkpiMap: Map<string, TKPIItem>,
  pmSettings: any,
  useLocalNotaOverrides: boolean,
  defaultBumbuPct: number = 10
): {
  foodCosts: ArchiveFoodCostSection[];
  notaItems: ArchiveNotaItemRecord[];
  notaTotalBelanja: number;
} {
  const dayPM = harianPM.find((h) => h.hariKe === dayNum) || harianPM[0] || { sasaran: [] };
  const besarSasaranIds: string[] = pmSettings?.porsiBesarSasaranIds || DEFAULT_BESAR_IDS;
  const kecilSasaranIds: string[] = pmSettings?.porsiKecilSasaranIds || DEFAULT_KECIL_IDS;
  const porsiBesarHarga = Number(pmSettings?.porsiBesarHarga) || 10000;
  const porsiKecilHarga = Number(pmSettings?.porsiKecilHarga) || 8000;

  const matchingFC = foodCostDays.filter(
    (d) => (dateStr && d.tanggal === dateStr) || d.hariKe === dayNum
  );

  // Deduplicate if same jenisMenu & hariKe appeared
  const seenKey = new Set<string>();
  const uniqueFC = matchingFC.filter((fc) => {
    const k = `${fc.hariKe}_${fc.jenisMenu}_${fc.masterMenuPlanId || ""}`;
    if (seenKey.has(k)) return false;
    seenKey.add(k);
    return true;
  });

  const notaAgg: Record<string, ArchiveNotaItemRecord> = {};
  const foodCosts: ArchiveFoodCostSection[] = [];

  uniqueFC.forEach((fc) => {
    const isAlergi = fc.jenisMenu.includes("Alergi");
    const effectiveBumbuPct = fc.bumbuPct !== undefined ? fc.bumbuPct : defaultBumbuPct;

    const mapBahan = (items: ReturnType<typeof calculateDay>["porsiBesarItems"]) =>
      items
        .filter((it) => it.nama && it.nama.trim() !== "")
        .map((it) => ({
          nama: it.nama,
          beratBB: it.beratBB,
          urt: it.urt || "-",
          kebutuhanKg: it.totalKebutuhanKg || 0,
          jumlahBuffer: Number(it.butir || 0),
          hargaSatuan: it.hargaSatuan || 0,
          hargaTotal: it.hargaTotal || 0,
        }));

    const aggregateNotaItems = (items: ReturnType<typeof calculateDay>["porsiBesarItems"]) => {
      items.forEach((it) => {
        const rawName = (it.nama || "").trim();
        if (!rawName || rawName.toLowerCase() === "ketik bahan...") return;
        const qty = Number(it.butir || 0) > 0 ? Number(it.butir) : Number(it.totalKebutuhanKg || 0);
        if (qty <= 0) return;

        const key = rawName.toLowerCase();
        let satuan = "Kg";
        if (it.bufferBase === "potong") satuan = "Potong";
        else if (it.bufferBase === "ekor") satuan = "Ekor";
        else if (it.bufferBase === "custom") satuan = guessSatuanNota(rawName);

        if (!notaAgg[key]) {
          notaAgg[key] = {
            nama: rawName,
            jumlah: Number(qty.toFixed(1)),
            satuan,
            hargaSatuan: it.hargaSatuan || 0,
            total: Number((qty * (it.hargaSatuan || 0)).toFixed(0)),
          };
        } else {
          notaAgg[key].jumlah = Number((notaAgg[key].jumlah + qty).toFixed(1));
          if ((it.hargaSatuan || 0) > notaAgg[key].hargaSatuan) {
            notaAgg[key].hargaSatuan = it.hargaSatuan || 0;
          }
          notaAgg[key].total = Number((notaAgg[key].jumlah * notaAgg[key].hargaSatuan).toFixed(0));
        }
      });
    };

    if (
      fc.activeMode === "kategori" &&
      Array.isArray(fc.categoryEntries) &&
      fc.categoryEntries.length > 0
    ) {
      fc.categoryEntries.forEach((ce) => {
        if (!ce.bahanList || ce.bahanList.length === 0) return;
        let entryPm = 0;
        if (ce.customPmCount !== undefined) {
          entryPm = Math.max(0, Number(ce.customPmCount) || 0);
        } else {
          const sharedIds = [
            "sampel",
            ...besarSasaranIds.filter((id) => kecilSasaranIds.includes(id)),
          ];
          entryPm = (ce.sasaranIds || []).reduce(
            (sum, sId) =>
              sum + resolveSasaranPmFromTable2(dayPM.sasaran, sId, ce.porsi, isAlergi, sharedIds),
            0
          );
        }

        const entryBuf = ce.bufferPct ?? fc.bufferPct ?? 5;
        const calcCat =
          ce.porsi === "besar"
            ? calculateDay(ce.bahanList, [], entryPm, 0, entryBuf, tkpiMap, effectiveBumbuPct)
            : calculateDay([], ce.bahanList, 0, entryPm, entryBuf, tkpiMap, effectiveBumbuPct);

        const hargaPorsi = ce.porsi === "besar" ? porsiBesarHarga : porsiKecilHarga;
        const rabTotal = entryPm * hargaPorsi;
        const actualTotal =
          ce.porsi === "besar" ? calcCat.subtotalBesarCost : calcCat.subtotalKecilCost;

        foodCosts.push({
          jenisMenu: `${fc.jenisMenu} · ${ce.title}`,
          namaMenu: fc.namaMenu || "",
          bufferPct: entryBuf,
          bumbuPct: effectiveBumbuPct,
          pmKecil: ce.porsi === "kecil" ? calcCat.jumlahPMKecil : 0,
          pmBesar: ce.porsi === "besar" ? calcCat.jumlahPMBesar : 0,
          rabTotal,
          actualTotal,
          costPerPorsiKecil: ce.porsi === "kecil" ? calcCat.costPerPorsiKecil : 0,
          costPerPorsiBesar: ce.porsi === "besar" ? calcCat.costPerPorsiBesar : 0,
          energiBesar: ce.porsi === "besar" ? calcCat.nutrisiPorsiBesar?.energi || 0 : 0,
          proteinBesar: ce.porsi === "besar" ? calcCat.nutrisiPorsiBesar?.protein || 0 : 0,
          energiKecil: ce.porsi === "kecil" ? calcCat.nutrisiPorsiKecil?.energi || 0 : 0,
          proteinKecil: ce.porsi === "kecil" ? calcCat.nutrisiPorsiKecil?.protein || 0 : 0,
          bahanBesar: ce.porsi === "besar" ? mapBahan(calcCat.porsiBesarItems) : [],
          bahanKecil: ce.porsi === "kecil" ? mapBahan(calcCat.porsiKecilItems) : [],
        });

        aggregateNotaItems([...calcCat.porsiBesarItems, ...calcCat.porsiKecilItems]);
      });
      return;
    }

    let pmBesar = 0;
    let pmKecil = 0;

    besarSasaranIds.forEach((sId) => {
      const sasItem = (dayPM.sasaran || []).find((item) => item.id === sId);
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

    kecilSasaranIds.forEach((sId) => {
      const sasItem = (dayPM.sasaran || []).find((item) => item.id === sId);
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

    if (fc.customPmBesarCount !== undefined) pmBesar = fc.customPmBesarCount;
    if (fc.customPmKecilCount !== undefined) pmKecil = fc.customPmKecilCount;

    const calc = calculateDay(
      fc.porsiBesarBahan || [],
      fc.porsiKecilBahan || [],
      pmBesar,
      pmKecil,
      fc.bufferPct ?? 5,
      tkpiMap,
      effectiveBumbuPct
    );

    const rabTotal = pmBesar * porsiBesarHarga + pmKecil * porsiKecilHarga;
    const actualTotal = calc.subtotalBesarCost + calc.subtotalKecilCost;

    foodCosts.push({
      jenisMenu: fc.jenisMenu,
      namaMenu: fc.namaMenu || "",
      bufferPct: fc.bufferPct ?? 5,
      bumbuPct: effectiveBumbuPct,
      pmKecil: calc.jumlahPMKecil,
      pmBesar: calc.jumlahPMBesar,
      rabTotal,
      actualTotal,
      costPerPorsiKecil: calc.costPerPorsiKecil,
      costPerPorsiBesar: calc.costPerPorsiBesar,
      energiBesar: calc.nutrisiPorsiBesar?.energi || 0,
      proteinBesar: calc.nutrisiPorsiBesar?.protein || 0,
      energiKecil: calc.nutrisiPorsiKecil?.energi || 0,
      proteinKecil: calc.nutrisiPorsiKecil?.protein || 0,
      bahanBesar: mapBahan(calc.porsiBesarItems),
      bahanKecil: mapBahan(calc.porsiKecilItems),
    });

    // Aggregate into Nota Pesanan
    aggregateNotaItems([...calc.porsiBesarItems, ...calc.porsiKecilItems]);
  });

  // Check manual saved Nota Pesanan items from localStorage for active day
  let notaItems: ArchiveNotaItemRecord[] = Object.values(notaAgg);
  if (useLocalNotaOverrides) {
    try {
      const savedNotaRaw = localStorage.getItem(`sisper_nota_items_harian_${dayNum}`);
      if (savedNotaRaw) {
        const parsed = JSON.parse(savedNotaRaw);
        if (Array.isArray(parsed)) {
          const validSaved = parsed
            .filter((row: any) => row && row.nama && String(row.nama).trim() !== "")
            .map((row: any) => {
              const jml = Number(row.jumlah) || 0;
              const hrg = Number(row.hargaSatuan) || 0;
              return {
                nama: String(row.nama).trim(),
                jumlah: jml,
                satuan: row.satuan || "Kg",
                hargaSatuan: hrg,
                total: jml * hrg,
              };
            });
          if (validSaved.length > 0) {
            notaItems = validSaved;
          }
        }
      }
    } catch {
      // fallback to computed aggregation
    }
  }

  const notaTotalBelanja = notaItems.reduce((acc, it) => acc + (it.total || 0), 0);
  return { foodCosts, notaItems, notaTotalBelanja };
}

function buildEntriesFromSource(
  sourceProfile: SPPGProfile,
  sourceHarianPM: HariPM[],
  sourceMenu: MasterMenu,
  sourceFoodCost: FoodCostDay[],
  tkpiMap: Map<string, TKPIItem>,
  pmSettings: any,
  sourceLabel: string,
  sourcePrefixId: string,
  useLocalNotaOverrides: boolean,
  defaultBumbuPct: number = 10
): DailyArchiveEntry[] {
  const configuredDates =
    sourceProfile?.periodeDates && sourceProfile.periodeDates.length > 0
      ? sourceProfile.periodeDates
      : sourceMenu?.scheduleConfig?.periodeDates || [];

  const normMenu = ensureMasterMenuPlans(sourceMenu, configuredDates);
  const resolvedDates =
    configuredDates.length > 0
      ? configuredDates
      : normMenu.scheduleConfig?.periodeDates || [];

  const dayCount = Math.max(
    resolvedDates.length,
    sourceHarianPM?.length || 0,
    1
  );

  const entries: DailyArchiveEntry[] = [];

  for (let i = 0; i < dayCount; i++) {
    const hariKe = i + 1;
    const rawDate = (resolvedDates[i] || "").trim();
    const fallbackMonth = normMenu.scheduleConfig?.calendarMonth || "2026-07";
    const dateStr = rawDate || `${fallbackMonth}-${String(Math.min(i + 1, 28)).padStart(2, "0")}`;
    const monthKey = /^\d{4}-\d{2}/.test(dateStr) ? dateStr.slice(0, 7) : fallbackMonth;

    const pm = computePMForDay(sourceHarianPM, hariKe);

    const menus: ArchiveMenuPlanRecord[] = (normMenu.plans || []).map((plan) => {
      const item = plan.days?.[i];
      const composed = composeMenuName(item);
      return {
        planTitle: plan.title || "Paket Menu",
        dietType: plan.dietType || "Standar",
        sasaranLabel: buildCategoryLabelFromIds(plan.sasaranIds || []),
        namaMenu: composed,
        karbohidrat: item?.karbohidrat || "",
        laukHewani: item?.laukHewani || "",
        laukNabati: item?.laukNabati || "",
        sayur: item?.sayur || "",
        buah: item?.buah || "",
        susu: item?.susu || "",
      };
    });

    const { foodCosts, notaItems, notaTotalBelanja } = buildFoodCostAndNotaForDay(
      hariKe,
      dateStr,
      sourceHarianPM,
      sourceFoodCost,
      tkpiMap,
      pmSettings,
      useLocalNotaOverrides,
      defaultBumbuPct
    );

    const firstNonEmptyMenu =
      menus.find((m) => m.namaMenu && m.namaMenu.trim() !== "")?.namaMenu ||
      foodCosts.find((fc) => fc.namaMenu && fc.namaMenu.trim() !== "")?.namaMenu ||
      `Menu Hari Ke-${hariKe}`;

    entries.push({
      id: `${sourcePrefixId}_${dateStr}_day_${hariKe}`,
      monthKey,
      dateStr,
      hariKe,
      primaryMenuName: firstNonEmptyMenu,
      sourceLabel,
      pm,
      menus,
      foodCosts,
      notaItems,
      notaTotalBelanja,
    });
  }

  return entries;
}

export default function ArsipTab({
  profile,
  harianPM,
  masterMenu,
  foodCostDays,
  tkpiList,
  pmSettings,
  savedPeriods = [],
  setSavedPeriods,
  loadPeriod,
  activeSekolahPM = [],
  activeTigaBPM = [],
  kopLines,
  leftLogo = "/assets/images/app_logo_sppg.png",
  rightLogo = "",
  kopConfig = DEFAULT_UNIFIED_KOP_CONFIG,
  paperSize = "A4",
  bumbuPct = 10,
  userEmail,
  isDemo = false,
  docArchives = [],
  setDocArchives,
  onOpenDocArchive,
}: ArsipTabProps) {
  const [confirmPeriodAction, setConfirmPeriodAction] = useState<{
    kind: "load" | "delete";
    period: any;
  } | null>(null);
  const [showSavedHistory, setShowSavedHistory] = useState<boolean>(true);
  const [docSourceFilter, setDocSourceFilter] = useState<"all" | "ikss" | "sop">("all");
  const [docSearchQuery, setDocSearchQuery] = useState<string>("");
  const [expandedDocArchiveIds, setExpandedDocArchiveIds] = useState<Record<string, boolean>>({});
  const [confirmDeleteDocArchive, setConfirmDeleteDocArchive] = useState<SavedDocArchive | null>(null);
  const tkpiMap = useMemo(() => createTkpiMap(tkpiList), [tkpiList]);

  const filteredDocArchives = useMemo(() => {
    const q = docSearchQuery.trim().toLowerCase();
    return docArchives.filter((arc) => {
      if (docSourceFilter !== "all" && arc.source !== docSourceFilter) return false;
      if (!q) return true;
      return (
        arc.title.toLowerCase().includes(q) ||
        (arc.sasaranSummary || "").toLowerCase().includes(q) ||
        (arc.periodeLabel || "").toLowerCase().includes(q) ||
        (arc.worksheets || []).some((ws) => (ws.title || "").toLowerCase().includes(q))
      );
    });
  }, [docArchives, docSourceFilter, docSearchQuery]);

  const docArchiveCounts = useMemo(() => {
    let ikss = 0;
    let sop = 0;
    docArchives.forEach((a) => {
      if (a.source === "ikss") ikss += 1;
      else if (a.source === "sop") sop += 1;
    });
    return { total: docArchives.length, ikss, sop };
  }, [docArchives]);

  const formatDocSavedTimestamp = (arc: SavedDocArchive): string => {
    const raw = arc.savedAt || arc.createdAtIso || arc.createdAt || "";
    if (!raw) return "-";
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return raw;
    try {
      return (
        d.toLocaleString("id-ID", {
          day: "numeric",
          month: "short",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        }) + " WIB"
      );
    } catch {
      return raw;
    }
  };

  const handlePrintDocArchive = (archive: SavedDocArchive) => {
    const printId = `print-doc-archive-${archive.id}`;
    if (!expandedDocArchiveIds[archive.id]) {
      setExpandedDocArchiveIds((prev) => ({ ...prev, [archive.id]: true }));
      setTimeout(() => {
        printElementSmart(printId, { title: archive.title, paperSize });
      }, 120);
    } else {
      printElementSmart(printId, { title: archive.title, paperSize });
    }
  };

  // Combine Active Planning Cycle + Historical Saved Periods into a monthly/daily index
  const allEntries = useMemo(() => {
    const activeEntries = buildEntriesFromSource(
      profile,
      harianPM,
      masterMenu,
      foodCostDays,
      tkpiMap,
      pmSettings,
      "Perencanaan Aktif",
      "active",
      true,
      bumbuPct
    );

    const byDateKey = new Map<string, DailyArchiveEntry>();
    // Saved periods first (historical baseline)
    savedPeriods.forEach((period: any) => {
      if (!period || !period.profile) return;
      const periodEntries = buildEntriesFromSource(
        period.profile,
        period.harianPM || [],
        period.masterMenu || masterMenu,
        period.foodCostDays || [],
        tkpiMap,
        pmSettings,
        period.namaPeriode || "Arsip Periode Tersimpan",
        period.id || "saved",
        false,
        bumbuPct
      );
      periodEntries.forEach((entry) => {
        byDateKey.set(`${entry.dateStr}_${entry.hariKe}`, entry);
      });
    });

    // Active state takes precedence for live planned dates
    activeEntries.forEach((entry) => {
      byDateKey.set(`${entry.dateStr}_${entry.hariKe}`, entry);
    });

    return Array.from(byDateKey.values()).sort((a, b) => {
      const cmp = a.dateStr.localeCompare(b.dateStr);
      return cmp !== 0 ? cmp : a.hariKe - b.hariKe;
    });
  }, [profile, harianPM, masterMenu, foodCostDays, tkpiMap, pmSettings, savedPeriods, bumbuPct]);

  const availableMonths = useMemo(() => {
    const map = new Map<string, { monthKey: string; count: number; totalPM: number }>();
    allEntries.forEach((e) => {
      const cur = map.get(e.monthKey) || { monthKey: e.monthKey, count: 0, totalPM: 0 };
      cur.count += 1;
      cur.totalPM += e.pm.totalPM;
      map.set(e.monthKey, cur);
    });
    return Array.from(map.values()).sort((a, b) => b.monthKey.localeCompare(a.monthKey));
  }, [allEntries]);

  const [selectedMonth, setSelectedMonth] = useState<string>(() => {
    return availableMonths[0]?.monthKey || "all";
  });
  const [categoryFilter, setCategoryFilter] = useState<"all" | ArchiveCategory>("all");
  const [searchQuery, setSearchQuery] = useState("");
  const [expandedIds, setExpandedIds] = useState<Record<string, boolean>>({});
  const [snapshotMessage, setSnapshotMessage] = useState<{ type: "success" | "error"; text: string } | null>(null);

  // Google Drive State (Access token strictly in-memory via googleDrive.ts)
  const [driveSession, setDriveSession] = useState<DriveSessionInfo | null>(() => getDriveSession());
  const [isConnectingDrive, setIsConnectingDrive] = useState(false);
  const [driveError, setDriveError] = useState<string | null>(null);
  const [driveErrorCode, setDriveErrorCode] = useState<DriveAuthErrorCode | null>(null);

  // Sync in-memory Drive session state if primary Firebase Auth signs out
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (firebaseUser) => {
      if (!firebaseUser) {
        clearDriveSession();
        setDriveSession(null);
      } else {
        setDriveSession(getDriveSession());
      }
    });
    return () => {
      unsubscribe();
    };
  }, []);

  // Selection state: map of `${entry.id}::${category}` -> boolean
  const [selectedItemsMap, setSelectedItemsMap] = useState<Record<string, boolean>>({});
  const [separateByMenuFolder, setSeparateByMenuFolder] = useState(true);

  // Confirmation modal & upload progress
  const [isConfirmModalOpen, setIsConfirmModalOpen] = useState(false);
  const [isUploadingToDrive, setIsUploadingToDrive] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<{ done: number; total: number; currentLabel: string } | null>(null);
  const [lastUploadResults, setLastUploadResults] = useState<{
    rootFolderName: string;
    monthFolderLink?: string;
    menuFolders: Array<{ menuName: string; webViewLink?: string; files: DriveUploadResult[] }>;
  } | null>(null);

  const [driveLogs, setDriveLogs] = useState<DriveArchiveLogItem[]>(() => {
    try {
      const raw = localStorage.getItem("sisper_drive_archive_logs");
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  });

  const activeMonth = useMemo(() => {
    if (selectedMonth === "all") return "all";
    if (availableMonths.some((m) => m.monthKey === selectedMonth)) return selectedMonth;
    return availableMonths[0]?.monthKey || "all";
  }, [selectedMonth, availableMonths]);

  const filteredEntries = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return allEntries.filter((entry) => {
      if (activeMonth !== "all" && entry.monthKey !== activeMonth) return false;
      if (!q) return true;
      return (
        entry.primaryMenuName.toLowerCase().includes(q) ||
        entry.dateStr.toLowerCase().includes(q) ||
        formatIndonesianDate(entry.dateStr).toLowerCase().includes(q) ||
        entry.menus.some((m) => m.namaMenu.toLowerCase().includes(q))
      );
    });
  }, [allEntries, activeMonth, searchQuery]);

  const monthStats = useMemo(() => {
    let totalPM = 0;
    let totalFoodCostActual = 0;
    let totalNota = 0;
    let menuCount = 0;
    filteredEntries.forEach((e) => {
      totalPM += e.pm.totalPM;
      totalFoodCostActual += e.foodCosts.reduce((acc, fc) => acc + fc.actualTotal, 0);
      totalNota += e.notaTotalBelanja;
      if (e.menus.some((m) => m.namaMenu)) menuCount += 1;
    });
    return {
      totalDates: filteredEntries.length,
      totalPM,
      menuCount,
      totalFoodCostActual,
      totalNota,
    };
  }, [filteredEntries]);

  const activeCategoriesList: ArchiveCategory[] = useMemo(() => {
    if (categoryFilter === "all") {
      return ["jumlah_pm", "menu", "food_cost", "nota_pesanan"];
    }
    return [categoryFilter];
  }, [categoryFilter]);

  const selectedUploadTasks = useMemo(() => {
    const tasks: Array<{
      key: string;
      entry: DailyArchiveEntry;
      category: ArchiveCategory;
      fileName: string;
      menuFolderName: string;
    }> = [];

    filteredEntries.forEach((entry) => {
      (["jumlah_pm", "menu", "food_cost", "nota_pesanan"] as ArchiveCategory[]).forEach((cat) => {
        const key = `${entry.id}::${cat}`;
        if (selectedItemsMap[key]) {
          tasks.push({
            key,
            entry,
            category: cat,
            fileName: buildArchivePdfFilename(entry, cat),
            menuFolderName: sanitizeDriveName(
              entry.primaryMenuName || `Menu Hari ${entry.hariKe}`,
              `Menu Hari ${entry.hariKe}`
            ),
          });
        }
      });
    });

    return tasks;
  }, [filteredEntries, selectedItemsMap]);

  const uniqueMenuFoldersPreview = useMemo(() => {
    const set = new Set<string>();
    selectedUploadTasks.forEach((t) => set.add(t.menuFolderName));
    return Array.from(set);
  }, [selectedUploadTasks]);

  const toggleExpand = (id: string) => {
    setExpandedIds((prev) => ({ ...prev, [id]: !prev[id] }));
  };

  const toggleItemSelection = (entryId: string, cat: ArchiveCategory) => {
    const key = `${entryId}::${cat}`;
    setSelectedItemsMap((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const toggleWholeDaySelection = (entryId: string, select: boolean) => {
    setSelectedItemsMap((prev) => {
      const next = { ...prev };
      activeCategoriesList.forEach((cat) => {
        next[`${entryId}::${cat}`] = select;
      });
      return next;
    });
  };

  const handleSelectAllFiltered = () => {
    const next: Record<string, boolean> = { ...selectedItemsMap };
    filteredEntries.forEach((entry) => {
      activeCategoriesList.forEach((cat) => {
        next[`${entry.id}::${cat}`] = true;
      });
    });
    setSelectedItemsMap(next);
  };

  const handleClearSelection = () => {
    setSelectedItemsMap({});
  };

  const handleDownloadLocalPdf = (entry: DailyArchiveEntry, category: ArchiveCategory) => {
    const blob = generateArchivePdfBlob({
      entry,
      category,
      profile,
      kopLines,
      leftLogo,
      rightLogo,
      kopConfig,
      paperSize,
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = buildArchivePdfFilename(entry, category);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleSnapshotMonthToArchive = () => {
    if (!setSavedPeriods) {
      setSnapshotMessage({ type: "error", text: "Fitur penyimpanan arsip belum tersedia." });
      return;
    }
    const targetMonthLabel =
      activeMonth === "all" ? "Siklus Perencanaan Aktif" : formatIndonesianMonthYear(activeMonth);
    const label = `Arsip ${targetMonthLabel} (${new Date().toLocaleDateString("id-ID")})`;
    const nowTime =
      new Date().toLocaleString("id-ID", {
        day: "numeric",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
      }) + " WIB";

    const newRecord = {
      id: `arsip_month_${Date.now()}`,
      namaPeriode: label,
      createdAt: nowTime,
      profile: JSON.parse(JSON.stringify(profile)),
      sekolahPM: JSON.parse(JSON.stringify(activeSekolahPM)),
      tigaBPM: JSON.parse(JSON.stringify(activeTigaBPM)),
      harianPM: JSON.parse(JSON.stringify(harianPM)),
      masterMenu: JSON.parse(JSON.stringify(masterMenu)),
      foodCostDays: JSON.parse(JSON.stringify(foodCostDays)),
    };

    setSavedPeriods([newRecord, ...savedPeriods]);
    setSnapshotMessage({
      type: "success",
      text: `Berhasil menyimpan snapshot "${label}" ke Riwayat Arsip Dokumen.`,
    });
    setTimeout(() => setSnapshotMessage(null), 5000);
  };

  const handleConnectDrive = async (skipLoginHint = false) => {
    setDriveError(null);
    setDriveErrorCode(null);
    setIsConnectingDrive(true);
    try {
      const session = await connectGoogleDrive(userEmail || undefined, { skipLoginHint });
      setDriveSession(session);
    } catch (err: any) {
      if (err instanceof DriveConnectionError) {
        setDriveErrorCode(err.code);
        setDriveError(err.message);
      } else {
        setDriveErrorCode("unknown");
        setDriveError(err?.message || "Gagal menghubungkan Google Drive.");
      }
    } finally {
      setIsConnectingDrive(false);
    }
  };

  const handleCancelConnectDrive = () => {
    cancelGoogleDriveConnect();
    setIsConnectingDrive(false);
  };

  const handleDisconnectDrive = () => {
    clearDriveSession();
    setDriveSession(null);
    setDriveError(null);
    setDriveErrorCode(null);
  };

  const handleOpenSendConfirmation = () => {
    setDriveError(null);
    if (!driveSession) {
      setDriveError("Silakan hubungkan akun Google Drive email aktif Anda terlebih dahulu sebelum mengirim arsip.");
      return;
    }
    if (selectedUploadTasks.length === 0) {
      setDriveError("Pilih minimal satu arsip (Jumlah PM / Menu / Food Cost / Nota Pesanan) yang ingin dikirim.");
      return;
    }
    setIsConfirmModalOpen(true);
  };

  const handleExecuteDriveUpload = async () => {
    if (!driveSession || selectedUploadTasks.length === 0) return;

    setIsUploadingToDrive(true);
    setDriveError(null);
    setLastUploadResults(null);
    setUploadProgress({ done: 0, total: selectedUploadTasks.length, currentLabel: "Menyiapkan folder di Google Drive..." });

    try {
      const token = driveSession.accessToken;
      const rootFolder = await findOrCreateDriveFolder(token, "Arsip SISPERMBG");
      const monthLabel =
        activeMonth === "all"
          ? "Semua Bulan"
          : formatIndonesianMonthYear(activeMonth);
      const monthFolderName = `Arsip ${monthLabel}`;
      const monthFolder = await findOrCreateDriveFolder(token, monthFolderName, rootFolder.id);

      const menuFolderCache: Record<string, { id: string; name: string; webViewLink?: string }> = {};
      const folderUploadsMap: Record<
        string,
        { menuName: string; webViewLink?: string; files: DriveUploadResult[] }
      > = {};

      for (let idx = 0; idx < selectedUploadTasks.length; idx++) {
        const task = selectedUploadTasks[idx];
        setUploadProgress({
          done: idx,
          total: selectedUploadTasks.length,
          currentLabel: `Mengonversi ke PDF & mengunggah: ${task.fileName}`,
        });

        let targetFolder = monthFolder;
        if (separateByMenuFolder) {
          if (!menuFolderCache[task.menuFolderName]) {
            const createdMenuFolder = await findOrCreateDriveFolder(
              token,
              task.menuFolderName,
              monthFolder.id
            );
            menuFolderCache[task.menuFolderName] = createdMenuFolder;
          }
          targetFolder = menuFolderCache[task.menuFolderName];
        }

        const pdfBlob = generateArchivePdfBlob({
          entry: task.entry,
          category: task.category,
          profile,
          kopLines,
          leftLogo,
          rightLogo,
          kopConfig,
          paperSize,
        });

        const uploaded = await uploadPdfToDriveFolder(
          token,
          targetFolder.id,
          task.fileName,
          pdfBlob
        );

        const groupKey = targetFolder.name;
        if (!folderUploadsMap[groupKey]) {
          folderUploadsMap[groupKey] = {
            menuName: groupKey,
            webViewLink: targetFolder.webViewLink,
            files: [],
          };
        }
        folderUploadsMap[groupKey].files.push({
          ...uploaded,
          folderWebViewLink: targetFolder.webViewLink,
        });
      }

      setUploadProgress({
        done: selectedUploadTasks.length,
        total: selectedUploadTasks.length,
        currentLabel: "Selesai mengunggah seluruh arsip PDF!",
      });

      const menuFoldersArr = Object.values(folderUploadsMap);
      setLastUploadResults({
        rootFolderName: `Arsip SISPERMBG / ${monthFolderName}`,
        monthFolderLink: monthFolder.webViewLink,
        menuFolders: menuFoldersArr,
      });

      const logItem: DriveArchiveLogItem = {
        id: `log_${Date.now()}`,
        sentAt: new Date().toLocaleString("id-ID"),
        driveEmail: driveSession.email,
        monthLabel,
        folderName: `Arsip SISPERMBG / ${monthFolderName}`,
        folderLink: monthFolder.webViewLink,
        fileCount: selectedUploadTasks.length,
        files: menuFoldersArr.flatMap((mf) =>
          mf.files.map((f) => ({
            name: f.name,
            webViewLink: f.webViewLink,
            menuFolder: mf.menuName,
          }))
        ),
      };

      const nextLogs = [logItem, ...driveLogs].slice(0, 20);
      setDriveLogs(nextLogs);
      localStorage.setItem("sisper_drive_archive_logs", JSON.stringify(nextLogs));
      setIsConfirmModalOpen(false);
    } catch (err: any) {
      setDriveError(err?.message || "Terjadi kesalahan saat mengunggah arsip ke Google Drive.");
      setIsConfirmModalOpen(false);
    } finally {
      setIsUploadingToDrive(false);
    }
  };

  const executeConfirmedPeriodAction = () => {
    if (!confirmPeriodAction) return;
    if (confirmPeriodAction.kind === "load" && loadPeriod) {
      loadPeriod(confirmPeriodAction.period);
      setSnapshotMessage({
        type: "success",
        text: `Berhasil membuka arsip "${confirmPeriodAction.period.namaPeriode}" ke lembar kerja aktif.`,
      });
    } else if (confirmPeriodAction.kind === "delete" && setSavedPeriods) {
      setSavedPeriods(
        savedPeriods.filter((p) => p.id !== confirmPeriodAction.period.id)
      );
      setSnapshotMessage({
        type: "success",
        text: `Arsip "${confirmPeriodAction.period.namaPeriode}" telah dihapus dari riwayat cloud.`,
      });
    }
    setConfirmPeriodAction(null);
  };

  return (
    <div id="arsip-dokumen-tab" className="space-y-6 font-sans">
      {/* HEADER & SNAPSHOT CONTROLS */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div className="space-y-1">
          <div className="flex flex-wrap items-center gap-2">
            <h2 className="text-lg font-extrabold text-slate-900 flex items-center gap-2">
              <Archive className="w-5 h-5 text-indigo-600" />
              Arsip Dokumen &amp; Sinkronisasi Google Drive
            </h2>
            {isDemo && (
              <span className="px-2.5 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-md text-[10px] font-extrabold uppercase tracking-wider">
                Mode Demo Aktif
              </span>
            )}
          </div>
          <p className="text-xs text-slate-500 max-w-3xl">
            Pusat <strong>Arsip Dokumen</strong> otomatis per bulan &amp; tanggal perencanaan mencakup <strong>Jumlah PM</strong>, <strong>Menu</strong>, <strong>Food Cost</strong>, <strong>Nota Pesanan</strong>, serta riwayat arsip cloud yang dapat dibuka kembali kapan saja atau dikirim dalam format <strong>PDF otomatis</strong> ke Google Drive Anda.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5 shrink-0">
          <button
            type="button"
            onClick={handleSnapshotMonthToArchive}
            className="px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition cursor-pointer shadow-xs"
          >
            <Save className="w-4 h-4 text-emerald-400" />
            Simpan Snapshot ke Arsip Dokumen
          </button>
        </div>
      </div>

      {snapshotMessage && (
        <div
          className={`p-3.5 rounded-xl border text-xs font-semibold flex items-center justify-between gap-3 ${
            snapshotMessage.type === "success"
              ? "bg-emerald-50 border-emerald-200 text-emerald-900"
              : "bg-rose-50 border-rose-200 text-rose-900"
          }`}
        >
          <span className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 shrink-0" />
            {snapshotMessage.text}
          </span>
          <button
            type="button"
            onClick={() => setSnapshotMessage(null)}
            className="font-bold hover:underline cursor-pointer"
          >
            Tutup
          </button>
        </div>
      )}

      {/* GOOGLE DRIVE SYNC PANEL */}
      <div className="bg-gradient-to-br from-slate-900 via-indigo-950 to-slate-900 text-white rounded-2xl p-5 shadow-md space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <CloudUpload className="w-5 h-5 text-emerald-400" />
              <h3 className="font-extrabold text-sm md:text-base">
                Kirim Arsip PDF ke Google Drive Email Aktif
              </h3>
            </div>
            <p className="text-xs text-slate-300 leading-relaxed max-w-2xl">
              Setiap arsip yang Anda pilih & konfirmasi akan dikonversi otomatis menjadi dokumen <strong>PDF</strong> resmi dan tersimpan di dalam folder terpisah sesuai <strong>Nama Menu</strong> pada Google Drive Anda.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            {driveSession ? (
              <div className="flex flex-wrap items-center gap-2">
                <div className="px-3 py-2 bg-emerald-500/15 border border-emerald-400/40 rounded-xl text-xs flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                  <span className="text-emerald-200 font-bold">
                    Drive Terhubung: {driveSession.email}
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => handleConnectDrive(true)}
                  disabled={isConnectingDrive}
                  className="px-3 py-2 bg-white/10 hover:bg-white/20 text-white rounded-xl text-xs font-bold border border-white/15 transition cursor-pointer disabled:opacity-60"
                >
                  {isConnectingDrive ? "Menghubungkan..." : "Ganti Email Drive"}
                </button>
                <button
                  type="button"
                  onClick={handleDisconnectDrive}
                  disabled={isConnectingDrive}
                  className="px-3 py-2 bg-rose-500/20 hover:bg-rose-500/30 text-rose-200 rounded-xl text-xs font-bold border border-rose-400/30 transition cursor-pointer"
                >
                  Putuskan
                </button>
              </div>
            ) : (
              <div className="flex flex-wrap items-center gap-2">
                {/* Official Google Sign-In Styled Button */}
                <button
                  type="button"
                  onClick={() => handleConnectDrive(false)}
                  disabled={isConnectingDrive}
                  className="inline-flex items-center gap-2.5 bg-white hover:bg-slate-50 text-[#1F1F1F] border border-[#747775] px-4 py-2 rounded-xl text-xs font-bold shadow-sm transition cursor-pointer disabled:opacity-60"
                >
                  {isConnectingDrive ? (
                    <Loader2 className="w-4 h-4 animate-spin text-blue-600" />
                  ) : (
                    <svg className="w-4 h-4 shrink-0" viewBox="0 0 24 24" aria-hidden="true">
                      <path
                        fill="#4285F4"
                        d="M23.745 12.27c0-.7-.06-1.4-.19-2.07H12v4.51h6.6c-.29 1.52-1.14 2.82-2.4 3.68v3.05h3.88c2.27-2.09 3.665-5.17 3.665-9.17z"
                      />
                      <path
                        fill="#34A853"
                        d="M12 24c3.24 0 5.95-1.08 7.93-2.91l-3.88-3.05c-1.08.72-2.45 1.16-4.05 1.16-3.12 0-5.77-2.11-6.72-4.96H1.29v3.14C3.26 21.3 7.31 24 12 24z"
                      />
                      <path
                        fill="#FBBC05"
                        d="M5.28 14.24c-.24-.72-.38-1.49-.38-2.24s.14-1.52.38-2.24V6.62H1.29C.47 8.24 0 10.06 0 12s.47 3.76 1.29 5.38l3.99-3.14z"
                      />
                      <path
                        fill="#EA4335"
                        d="M12 4.75c1.77 0 3.35.61 4.6 1.8l3.42-3.42C17.95 1.19 15.24 0 12 0 7.31 0 3.26 2.7 1.29 6.62l3.99 3.14c.95-2.85 3.6-4.96 6.72-4.96z"
                      />
                    </svg>
                  )}
                  <span>{isConnectingDrive ? "Menghubungkan..." : "Sign in with Google (Drive)"}</span>
                </button>
                {isConnectingDrive && (
                  <button
                    type="button"
                    onClick={handleCancelConnectDrive}
                    className="px-3 py-2 bg-white/10 hover:bg-white/20 text-white rounded-xl text-xs font-bold border border-white/20 transition cursor-pointer"
                  >
                    Batal
                  </button>
                )}
              </div>
            )}
          </div>
        </div>

        {/* Selection Action Bar */}
        <div className="bg-white/5 border border-white/10 rounded-xl p-3.5 flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-3 text-xs">
            <span className="px-2.5 py-1 bg-indigo-500/20 border border-indigo-400/30 rounded-lg font-extrabold text-indigo-200">
              {selectedUploadTasks.length} File PDF Dipilih
            </span>
            <label className="inline-flex items-center gap-2 cursor-pointer text-slate-200 font-semibold select-none">
              <input
                type="checkbox"
                checked={separateByMenuFolder}
                onChange={(e) => setSeparateByMenuFolder(e.target.checked)}
                className="w-4 h-4 accent-emerald-400 rounded cursor-pointer"
              />
              <span>Pisahkan folder otomatis sesuai Nama Menu</span>
            </label>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={handleSelectAllFiltered}
              className="px-3 py-1.5 bg-white/10 hover:bg-white/20 text-white rounded-lg text-xs font-bold transition cursor-pointer"
            >
              Pilih Semua ({filteredEntries.length} Tgl)
            </button>
            {selectedUploadTasks.length > 0 && (
              <button
                type="button"
                onClick={handleClearSelection}
                className="px-3 py-1.5 bg-rose-500/20 hover:bg-rose-500/30 text-rose-200 rounded-lg text-xs font-bold transition cursor-pointer"
              >
                Bersihkan Pilihan
              </button>
            )}
            <button
              type="button"
              onClick={handleOpenSendConfirmation}
              disabled={isUploadingToDrive}
              className="px-4 py-1.5 bg-emerald-500 hover:bg-emerald-400 text-slate-950 rounded-lg text-xs font-black flex items-center gap-1.5 transition cursor-pointer shadow-sm disabled:opacity-50"
            >
              <CloudUpload className="w-4 h-4" />
              Konfirmasi & Kirim ke Google Drive
            </button>
          </div>
        </div>

        {driveError && (
          <div className="bg-rose-500/15 border border-rose-400/40 text-rose-100 rounded-xl p-3.5 text-xs space-y-2.5">
            <div className="flex items-start justify-between gap-2">
              <span className="flex items-start gap-2">
                <AlertCircle className="w-4 h-4 text-rose-300 shrink-0 mt-0.5" />
                <span className="leading-relaxed">{driveError}</span>
              </span>
              <button
                type="button"
                onClick={() => {
                  setDriveError(null);
                  setDriveErrorCode(null);
                }}
                className="text-rose-200 font-bold hover:underline shrink-0 cursor-pointer"
              >
                Tutup
              </button>
            </div>

            {(driveErrorCode === "testing_access_denied" ||
              driveErrorCode === "popup_closed" ||
              driveErrorCode === "popup_timeout") && (
              <div className="bg-slate-950/60 border border-white/10 rounded-lg p-3 space-y-2 text-[11px] text-slate-200">
                <div className="font-bold text-amber-300">
                  Panduan Mengatasi Error 403: access_denied (Google Verification / Testing Mode):
                </div>
                <ol className="list-decimal list-inside space-y-1 text-slate-300 leading-relaxed">
                  <li>
                    Buka{" "}
                    <a
                      href={DRIVE_OAUTH_AUDIENCE_URL}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="text-emerald-300 font-bold underline inline-flex items-center gap-0.5"
                    >
                      Google Auth Platform → Audience
                      <ExternalLink className="w-3 h-3" />
                    </a>{" "}
                    (login sebagai pemilik project Google Cloud).
                  </li>
                  <li>
                    Di bagian <strong>Test users</strong>, klik <strong>+ Add users</strong> lalu tambahkan email Google yang ingin dipakai{userEmail && userEmail.includes("@") && !userEmail.endsWith("@sispermbg.id") ? ` (${userEmail})` : ""}, atau klik <strong>Publish App</strong> agar semua akun Google dapat menghubungkan Drive.
                  </li>
                  <li>
                    Atau gunakan akun Google tester/developer yang sudah terdaftar dengan menekan tombol di bawah ini:
                  </li>
                </ol>
                <div className="flex flex-wrap items-center gap-2 pt-1">
                  <button
                    type="button"
                    onClick={() => handleConnectDrive(true)}
                    disabled={isConnectingDrive}
                    className="px-3 py-1.5 bg-white text-slate-900 hover:bg-slate-100 rounded-lg font-bold text-[11px] inline-flex items-center gap-1.5 cursor-pointer transition disabled:opacity-60"
                  >
                    Pilih Akun Google Lain (Tanpa Hint Email)
                  </button>
                  <a
                    href={DRIVE_OAUTH_AUDIENCE_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="px-3 py-1.5 bg-emerald-500/20 hover:bg-emerald-500/30 border border-emerald-400/40 text-emerald-200 rounded-lg font-bold text-[11px] inline-flex items-center gap-1 transition"
                  >
                    Buka Google Auth Audience
                    <ExternalLink className="w-3 h-3" />
                  </a>
                </div>
              </div>
            )}
          </div>
        )}

        {uploadProgress && isUploadingToDrive && (
          <div className="bg-blue-500/15 border border-blue-400/30 rounded-xl p-3 space-y-2 text-xs">
            <div className="flex items-center justify-between font-bold text-blue-100">
              <span className="flex items-center gap-2">
                <Loader2 className="w-4 h-4 animate-spin text-emerald-400" />
                {uploadProgress.currentLabel}
              </span>
              <span>
                {uploadProgress.done}/{uploadProgress.total}
              </span>
            </div>
            <div className="w-full bg-slate-800 rounded-full h-2 overflow-hidden">
              <div
                className="bg-emerald-400 h-full transition-all duration-300"
                style={{
                  width: `${
                    uploadProgress.total > 0
                      ? Math.round((uploadProgress.done / uploadProgress.total) * 100)
                      : 0
                  }%`,
                }}
              />
            </div>
          </div>
        )}

        {lastUploadResults && (
          <div className="bg-emerald-500/15 border border-emerald-400/40 rounded-xl p-4 space-y-2.5 text-xs">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div className="flex items-center gap-2 text-emerald-200 font-extrabold">
                <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
                <span>
                  Arsip PDF Berhasil Dikirim ke Google Drive ({lastUploadResults.rootFolderName})
                </span>
              </div>
              {lastUploadResults.monthFolderLink && (
                <a
                  href={lastUploadResults.monthFolderLink}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1 px-2.5 py-1 bg-emerald-500 text-slate-950 font-black rounded-lg hover:bg-emerald-400 transition"
                >
                  Buka Folder Bulan di Drive
                  <ExternalLink className="w-3.5 h-3.5" />
                </a>
              )}
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
              {lastUploadResults.menuFolders.map((mf) => (
                <div
                  key={mf.menuName}
                  className="bg-slate-900/60 border border-white/10 rounded-lg p-2.5 space-y-1.5"
                >
                  <div className="flex items-center justify-between gap-2 font-bold text-amber-300">
                    <span className="flex items-center gap-1.5 truncate">
                      <Folder className="w-3.5 h-3.5 shrink-0" />
                      {mf.menuName}
                    </span>
                    {mf.webViewLink && (
                      <a
                        href={mf.webViewLink}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-[10px] text-emerald-300 hover:underline flex items-center gap-0.5 shrink-0"
                      >
                        Buka Folder <ExternalLink className="w-3 h-3" />
                      </a>
                    )}
                  </div>
                  <div className="space-y-1">
                    {mf.files.map((f) => (
                      <div key={f.fileId} className="flex items-center justify-between text-[11px] text-slate-200">
                        <span className="truncate pr-2">{f.name}</span>
                        {f.webViewLink && (
                          <a
                            href={f.webViewLink}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="text-sky-300 hover:underline shrink-0 flex items-center gap-0.5"
                          >
                            PDF <ExternalLink className="w-3 h-3" />
                          </a>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* MONTH & CATEGORY FILTERS */}
      <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          {/* Month Tabs */}
          <div className="flex flex-wrap items-center gap-1.5">
            <span className="text-xs font-bold text-slate-500 mr-1 flex items-center gap-1">
              <Calendar className="w-3.5 h-3.5 text-indigo-600" />
              Bulan Arsip:
            </span>
            <button
              type="button"
              onClick={() => setSelectedMonth("all")}
              className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer ${
                activeMonth === "all"
                  ? "bg-slate-900 text-white shadow-xs"
                  : "bg-slate-100 text-slate-600 hover:bg-slate-200"
              }`}
            >
              Semua Bulan ({allEntries.length})
            </button>
            {availableMonths.map((m) => (
              <button
                key={m.monthKey}
                type="button"
                onClick={() => setSelectedMonth(m.monthKey)}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer ${
                  activeMonth === m.monthKey
                    ? "bg-indigo-600 text-white shadow-xs"
                    : "bg-indigo-50 text-indigo-700 hover:bg-indigo-100"
                }`}
              >
                {formatIndonesianMonthYear(m.monthKey)} ({m.count} Tgl)
              </button>
            ))}
          </div>

          {/* Search */}
          <div className="relative w-full lg:w-72">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Cari nama menu atau tanggal..."
              className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
            />
          </div>
        </div>

        {/* 4 Category Differentiation Tabs */}
        <div className="flex flex-wrap items-center gap-2 pt-2 border-t border-slate-100">
          <button
            type="button"
            onClick={() => setCategoryFilter("all")}
            className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer ${
              categoryFilter === "all"
                ? "bg-slate-900 text-white"
                : "bg-slate-100 text-slate-600 hover:bg-slate-200"
            }`}
          >
            Semua Jenis Arsip (4 Kategori)
          </button>
          {ALL_CATEGORIES.map((cat) => (
            <button
              key={cat.id}
              type="button"
              onClick={() => setCategoryFilter(cat.id)}
              className={`px-3.5 py-2 rounded-xl text-xs font-bold border transition cursor-pointer flex items-center gap-1.5 ${
                categoryFilter === cat.id
                  ? "bg-white border-indigo-500 text-indigo-700 ring-2 ring-indigo-500/20 shadow-2xs"
                  : "bg-white border-slate-200 text-slate-600 hover:bg-slate-50"
              }`}
            >
              {cat.id === "jumlah_pm" && <Users className="w-3.5 h-3.5 text-cyan-600" />}
              {cat.id === "menu" && <ChefHat className="w-3.5 h-3.5 text-emerald-600" />}
              {cat.id === "food_cost" && <DollarSign className="w-3.5 h-3.5 text-amber-600" />}
              {cat.id === "nota_pesanan" && <ShoppingCart className="w-3.5 h-3.5 text-indigo-600" />}
              {cat.label}
            </button>
          ))}
        </div>
      </div>

      {/* MONTHLY SUMMARY CARDS */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
        <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
          <p className="text-[11px] font-bold text-slate-400 uppercase">Tanggal Terencana</p>
          <p className="text-lg font-black text-slate-900 mt-0.5">{monthStats.totalDates} Hari</p>
        </div>
        <div className="bg-white p-3.5 rounded-xl border border-cyan-100 shadow-2xs">
          <p className="text-[11px] font-bold text-cyan-700 uppercase">Total Arsip PM</p>
          <p className="text-lg font-black text-cyan-900 mt-0.5">
            {monthStats.totalPM.toLocaleString("id-ID")} <span className="text-xs font-bold">Porsi</span>
          </p>
        </div>
        <div className="bg-white p-3.5 rounded-xl border border-emerald-100 shadow-2xs">
          <p className="text-[11px] font-bold text-emerald-700 uppercase">Variasi Menu</p>
          <p className="text-lg font-black text-emerald-900 mt-0.5">{monthStats.menuCount} Hari Menu</p>
        </div>
        <div className="bg-white p-3.5 rounded-xl border border-amber-100 shadow-2xs">
          <p className="text-[11px] font-bold text-amber-700 uppercase">Realisasi Food Cost</p>
          <p className="text-base font-black text-amber-900 mt-0.5">
            {formatRupiah(monthStats.totalFoodCostActual)}
          </p>
        </div>
        <div className="bg-white p-3.5 rounded-xl border border-indigo-100 shadow-2xs">
          <p className="text-[11px] font-bold text-indigo-700 uppercase">Total Nota Pesanan</p>
          <p className="text-base font-black text-indigo-900 mt-0.5">
            {formatRupiah(monthStats.totalNota)}
          </p>
        </div>
      </div>

      {/* DAILY ARCHIVE LIST GROUPED BY PLANNED DATE */}
      <div className="space-y-4">
        {filteredEntries.length === 0 ? (
          <div className="bg-white rounded-2xl border border-slate-200 p-10 text-center space-y-2">
            <Archive className="w-10 h-10 text-slate-300 mx-auto" />
            <h4 className="font-bold text-slate-700 text-sm">Belum Ada Data Arsip pada Filter Ini</h4>
            <p className="text-xs text-slate-400 max-w-md mx-auto">
              Susun jadwal & tanggal di panel Master Menu lalu kirim ke Food Cost, atau ubah filter bulan/pencarian di atas.
            </p>
          </div>
        ) : (
          filteredEntries.map((entry) => {
            const isExpanded = Boolean(expandedIds[entry.id]);
            const allDayCatsSelected = activeCategoriesList.every(
              (c) => selectedItemsMap[`${entry.id}::${c}`]
            );
            const selectedCountForDay = activeCategoriesList.filter(
              (c) => selectedItemsMap[`${entry.id}::${c}`]
            ).length;

            return (
              <div
                key={entry.id}
                className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden"
              >
                {/* Date Card Header */}
                <div className="p-4 bg-slate-50/70 border-b border-slate-100 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
                  <div className="flex items-start gap-3">
                    <input
                      type="checkbox"
                      checked={allDayCatsSelected}
                      onChange={(e) => toggleWholeDaySelection(entry.id, e.target.checked)}
                      className="w-4 h-4 mt-1 accent-indigo-600 rounded cursor-pointer shrink-0"
                      title="Pilih semua kategori arsip di tanggal ini untuk dikirim ke Google Drive"
                    />
                    <div className="space-y-1">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="px-2.5 py-0.5 bg-indigo-600 text-white rounded-md text-[11px] font-extrabold">
                          {formatIndonesianDate(entry.dateStr) || entry.dateStr}
                        </span>
                        <span className="px-2 py-0.5 bg-slate-200 text-slate-700 rounded-md text-[10px] font-bold">
                          Hari Ke-{entry.hariKe}
                        </span>
                        <span className="px-2 py-0.5 bg-white border border-slate-200 text-slate-500 rounded-md text-[10px] font-semibold">
                          {entry.sourceLabel}
                        </span>
                        {selectedCountForDay > 0 && (
                          <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-md text-[10px] font-bold">
                            {selectedCountForDay} Kategori Siap Kirim Drive
                          </span>
                        )}
                      </div>
                      <h4 className="font-extrabold text-slate-900 text-sm">
                        Menu: {entry.primaryMenuName}
                      </h4>
                    </div>
                  </div>

                  <div className="flex flex-wrap items-center gap-2 shrink-0">
                    {/* Quick Category Checkboxes & Local PDF Download */}
                    <button
                      type="button"
                      onClick={() => toggleExpand(entry.id)}
                      className="px-3 py-1.5 bg-white hover:bg-slate-100 border border-slate-200 text-slate-700 rounded-xl text-xs font-bold flex items-center gap-1 transition cursor-pointer"
                    >
                      {isExpanded ? (
                        <>
                          <ChevronDown className="w-4 h-4" />
                          Tutup Detail
                        </>
                      ) : (
                        <>
                          <ChevronRight className="w-4 h-4" />
                          Lihat Rincian Tabel
                        </>
                      )}
                    </button>
                  </div>
                </div>

                {/* 4 Differentiated Category Cards Grid */}
                <div className="p-4 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-3">
                  {ALL_CATEGORIES.filter(
                    (c) => categoryFilter === "all" || categoryFilter === c.id
                  ).map((cat) => {
                    const isSelected = Boolean(selectedItemsMap[`${entry.id}::${cat.id}`]);
                    return (
                      <div
                        key={cat.id}
                        className={`rounded-xl border p-3 flex flex-col justify-between gap-2.5 transition ${
                          isSelected
                            ? "border-indigo-500 bg-indigo-50/30 ring-1 ring-indigo-500/20"
                            : "border-slate-200 bg-white"
                        }`}
                      >
                        <div className="flex items-start justify-between gap-2">
                          <label className="flex items-center gap-2 cursor-pointer">
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={() => toggleItemSelection(entry.id, cat.id)}
                              className="w-4 h-4 accent-indigo-600 rounded cursor-pointer"
                            />
                            <span className={`px-2 py-0.5 rounded-md border text-[11px] font-extrabold ${cat.badgeClass}`}>
                              {cat.shortLabel}
                            </span>
                          </label>

                          <button
                            type="button"
                            onClick={() => handleDownloadLocalPdf(entry, cat.id)}
                            title={`Unduh PDF ${cat.shortLabel} tanggal ${entry.dateStr}`}
                            className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[10px] font-bold flex items-center gap-1 transition cursor-pointer"
                          >
                            <Download className="w-3 h-3" />
                            PDF
                          </button>
                        </div>

                        {/* Category summary body */}
                        {cat.id === "jumlah_pm" && (
                          <div className="text-xs space-y-1 text-slate-600">
                            <div className="font-extrabold text-slate-900 text-sm">
                              {entry.pm.totalPM.toLocaleString("id-ID")} Penerima
                            </div>
                            <div className="flex justify-between text-[11px]">
                              <span>Porsi Kecil: <strong>{entry.pm.totalKecil}</strong></span>
                              <span>Porsi Besar: <strong>{entry.pm.totalBesar}</strong></span>
                            </div>
                            <div className="text-[11px] text-amber-700 font-semibold">
                              Alergi Khusus: {entry.pm.totalAlergi} PM
                            </div>
                          </div>
                        )}

                        {cat.id === "menu" && (
                          <div className="text-xs space-y-1 text-slate-600">
                            <div className="font-bold text-slate-900 line-clamp-2">
                              {entry.primaryMenuName}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {entry.menus.filter((m) => m.namaMenu).length} paket kategori terisi
                            </div>
                          </div>
                        )}

                        {cat.id === "food_cost" && (
                          <div className="text-xs space-y-1 text-slate-600">
                            <div className="font-extrabold text-slate-900 text-sm">
                              {formatRupiah(
                                entry.foodCosts.reduce((a, f) => a + f.actualTotal, 0)
                              )}
                            </div>
                            <div className="text-[11px]">
                              RAB: <strong>{formatRupiah(entry.foodCosts.reduce((a, f) => a + f.rabTotal, 0))}</strong>
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {entry.foodCosts.length} kategori food cost
                            </div>
                          </div>
                        )}

                        {cat.id === "nota_pesanan" && (
                          <div className="text-xs space-y-1 text-slate-600">
                            <div className="font-extrabold text-slate-900 text-sm">
                              {formatRupiah(entry.notaTotalBelanja)}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {entry.notaItems.length} item bahan pesanan
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>

                {/* Expandable Full Archive Detail Tables */}
                {isExpanded && (
                  <div className="px-4 pb-5 pt-2 border-t border-slate-100 bg-slate-50/40 space-y-4">
                    {(categoryFilter === "all" || categoryFilter === "jumlah_pm") && (
                      <div className="bg-white rounded-xl border border-slate-200 p-3.5 space-y-2">
                        <h5 className="font-extrabold text-xs text-cyan-800 flex items-center gap-1.5">
                          <Users className="w-4 h-4" />
                          Rincian Arsip Jumlah PM — {formatIndonesianDate(entry.dateStr)}
                        </h5>
                        <div className="overflow-x-auto">
                          <table className="w-full text-xs text-left border-collapse">
                            <thead className="bg-slate-100 text-slate-700 text-[11px]">
                              <tr>
                                <th className="p-2 border border-slate-200">Kelompok Sasaran</th>
                                <th className="p-2 border border-slate-200 text-center">Porsi Kecil</th>
                                <th className="p-2 border border-slate-200 text-center">Porsi Besar</th>
                                <th className="p-2 border border-slate-200 text-center">Alergi</th>
                                <th className="p-2 border border-slate-200 text-center">Total</th>
                              </tr>
                            </thead>
                            <tbody>
                              {entry.pm.rows.map((r, idx) => (
                                <tr key={idx} className="hover:bg-slate-50">
                                  <td className="p-2 border border-slate-200 font-semibold">{r.label}</td>
                                  <td className="p-2 border border-slate-200 text-center">{r.porsiKecil}</td>
                                  <td className="p-2 border border-slate-200 text-center">{r.porsiBesar}</td>
                                  <td className="p-2 border border-slate-200 text-center">{r.alergiKecil + r.alergiBesar}</td>
                                  <td className="p-2 border border-slate-200 text-center font-bold">{r.total}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {(categoryFilter === "all" || categoryFilter === "menu") && (
                      <div className="bg-white rounded-xl border border-slate-200 p-3.5 space-y-2">
                        <h5 className="font-extrabold text-xs text-emerald-800 flex items-center gap-1.5">
                          <ChefHat className="w-4 h-4" />
                          Rincian Arsip Menu — {formatIndonesianDate(entry.dateStr)}
                        </h5>
                        <div className="overflow-x-auto">
                          <table className="w-full text-xs text-left border-collapse">
                            <thead className="bg-slate-100 text-slate-700 text-[11px]">
                              <tr>
                                <th className="p-2 border border-slate-200">Paket / Kategori</th>
                                <th className="p-2 border border-slate-200">Nama Menu</th>
                                <th className="p-2 border border-slate-200">Pokok</th>
                                <th className="p-2 border border-slate-200">Lauk Hewani</th>
                                <th className="p-2 border border-slate-200">Lauk Nabati</th>
                                <th className="p-2 border border-slate-200">Sayur</th>
                                <th className="p-2 border border-slate-200">Buah / Susu</th>
                              </tr>
                            </thead>
                            <tbody>
                              {entry.menus.map((m, idx) => (
                                <tr key={idx} className="hover:bg-slate-50">
                                  <td className="p-2 border border-slate-200 font-semibold">
                                    <div>{m.planTitle}</div>
                                    <div className="text-[10px] text-slate-400">{m.dietType} • {m.sasaranLabel}</div>
                                  </td>
                                  <td className="p-2 border border-slate-200 font-bold text-slate-900">{m.namaMenu || "-"}</td>
                                  <td className="p-2 border border-slate-200">{m.karbohidrat || "-"}</td>
                                  <td className="p-2 border border-slate-200">{m.laukHewani || "-"}</td>
                                  <td className="p-2 border border-slate-200">{m.laukNabati || "-"}</td>
                                  <td className="p-2 border border-slate-200">{m.sayur || "-"}</td>
                                  <td className="p-2 border border-slate-200">{[m.buah, m.susu].filter(Boolean).join(" / ") || "-"}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {(categoryFilter === "all" || categoryFilter === "food_cost") && (
                      <div className="bg-white rounded-xl border border-slate-200 p-3.5 space-y-2">
                        <h5 className="font-extrabold text-xs text-amber-800 flex items-center gap-1.5">
                          <DollarSign className="w-4 h-4" />
                          Rincian Arsip Food Cost — {formatIndonesianDate(entry.dateStr)}
                        </h5>
                        {entry.foodCosts.length === 0 ? (
                          <p className="text-xs text-slate-400 italic">Belum ada rincian Food Cost pada tanggal ini.</p>
                        ) : (
                          <div className="space-y-3">
                            {entry.foodCosts.map((fc, idx) => (
                              <div key={idx} className="border border-amber-200 rounded-xl p-3 space-y-2 bg-amber-50/20">
                                <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
                                  <div className="font-extrabold text-slate-900">
                                    [{fc.jenisMenu}] {fc.namaMenu || entry.primaryMenuName}
                                  </div>
                                  <div className="flex flex-wrap gap-3 text-[11px] font-semibold">
                                    <span>PM Kecil: {fc.pmKecil} ({formatRupiah(fc.costPerPorsiKecil)}/porsi)</span>
                                    <span>PM Besar: {fc.pmBesar} ({formatRupiah(fc.costPerPorsiBesar)}/porsi)</span>
                                    <span className="text-emerald-700 font-bold">Realisasi: {formatRupiah(fc.actualTotal)}</span>
                                  </div>
                                </div>
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-2 text-[11px]">
                                  <div className="bg-white p-2 rounded-lg border border-slate-200">
                                    <div className="font-bold text-indigo-700 mb-1">
                                      Bahan Porsi Besar ({fc.bahanBesar.length} item • {fc.energiBesar.toFixed(0)} kkal)
                                    </div>
                                    {fc.bahanBesar.length === 0 ? (
                                      <span className="text-slate-400 italic">Kosong</span>
                                    ) : (
                                      <ul className="space-y-0.5">
                                        {fc.bahanBesar.map((b, bi) => (
                                          <li key={bi} className="flex justify-between gap-2">
                                            <span className="truncate">{b.nama} ({b.beratBB}g)</span>
                                            <span className="font-mono shrink-0">{formatRupiah(b.hargaTotal)}</span>
                                          </li>
                                        ))}
                                      </ul>
                                    )}
                                  </div>
                                  <div className="bg-white p-2 rounded-lg border border-slate-200">
                                    <div className="font-bold text-purple-700 mb-1">
                                      Bahan Porsi Kecil ({fc.bahanKecil.length} item • {fc.energiKecil.toFixed(0)} kkal)
                                    </div>
                                    {fc.bahanKecil.length === 0 ? (
                                      <span className="text-slate-400 italic">Kosong</span>
                                    ) : (
                                      <ul className="space-y-0.5">
                                        {fc.bahanKecil.map((b, ki) => (
                                          <li key={ki} className="flex justify-between gap-2">
                                            <span className="truncate">{b.nama} ({b.beratBB}g)</span>
                                            <span className="font-mono shrink-0">{formatRupiah(b.hargaTotal)}</span>
                                          </li>
                                        ))}
                                      </ul>
                                    )}
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    )}

                    {(categoryFilter === "all" || categoryFilter === "nota_pesanan") && (
                      <div className="bg-white rounded-xl border border-slate-200 p-3.5 space-y-2">
                        <div className="flex items-center justify-between">
                          <h5 className="font-extrabold text-xs text-indigo-800 flex items-center gap-1.5">
                            <ShoppingCart className="w-4 h-4" />
                            Rincian Arsip Nota Pesanan — {formatIndonesianDate(entry.dateStr)}
                          </h5>
                          <span className="text-xs font-black text-indigo-700">
                            Total: {formatRupiah(entry.notaTotalBelanja)}
                          </span>
                        </div>
                        {entry.notaItems.length === 0 ? (
                          <p className="text-xs text-slate-400 italic">Belum ada item pesanan bahan untuk tanggal ini.</p>
                        ) : (
                          <div className="overflow-x-auto">
                            <table className="w-full text-xs text-left border-collapse">
                              <thead className="bg-slate-100 text-slate-700 text-[11px]">
                                <tr>
                                  <th className="p-2 border border-slate-200 w-10 text-center">No</th>
                                  <th className="p-2 border border-slate-200">Uraian Bahan Makanan</th>
                                  <th className="p-2 border border-slate-200 text-center">Jumlah</th>
                                  <th className="p-2 border border-slate-200 text-center">Satuan</th>
                                  <th className="p-2 border border-slate-200 text-right">Harga Satuan</th>
                                  <th className="p-2 border border-slate-200 text-right">Total</th>
                                </tr>
                              </thead>
                              <tbody>
                                {entry.notaItems.map((n, idx) => (
                                  <tr key={idx} className="hover:bg-slate-50">
                                    <td className="p-2 border border-slate-200 text-center">{idx + 1}</td>
                                    <td className="p-2 border border-slate-200 font-semibold">{n.nama}</td>
                                    <td className="p-2 border border-slate-200 text-center font-mono">{n.jumlah}</td>
                                    <td className="p-2 border border-slate-200 text-center">{n.satuan}</td>
                                    <td className="p-2 border border-slate-200 text-right font-mono">{formatRupiah(n.hargaSatuan)}</td>
                                    <td className="p-2 border border-slate-200 text-right font-mono font-bold">{formatRupiah(n.total)}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>

      {/* DRIVE SEND HISTORY LOG */}
      {driveLogs.length > 0 && (
        <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-3">
          <h4 className="font-extrabold text-sm text-slate-800 flex items-center gap-2">
            <Package className="w-4 h-4 text-emerald-600" />
            Riwayat Pengiriman Arsip PDF ke Google Drive
          </h4>
          <div className="space-y-2">
            {driveLogs.map((log) => (
              <div
                key={log.id}
                className="p-3 bg-slate-50 border border-slate-200 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs"
              >
                <div>
                  <div className="font-bold text-slate-800">
                    {log.folderName} • <span className="text-emerald-700">{log.fileCount} File PDF</span>
                  </div>
                  <div className="text-[11px] text-slate-500">
                    Dikirim ke <strong>{log.driveEmail}</strong> pada {log.sentAt}
                  </div>
                </div>
                {log.folderLink && (
                  <a
                    href={log.folderLink}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="px-3 py-1.5 bg-white hover:bg-slate-100 border border-slate-200 rounded-lg font-bold text-indigo-700 inline-flex items-center gap-1 shrink-0"
                  >
                    Buka di Drive <ExternalLink className="w-3.5 h-3.5" />
                  </a>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ARSIP LEMBAR KERJA MENU IKSS & FORM SOP */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          <div className="flex items-start gap-3">
            <div className="p-2.5 bg-teal-50 rounded-xl text-teal-600 shrink-0">
              <ClipboardCheck className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-extrabold text-slate-900">
                Arsip Lembar Kerja Menu IKSS &amp; Form SOP
              </h3>
              <p className="text-xs text-slate-500">
                Kumpulan lembar kerja per kelompok sasaran yang disimpan dari tombol <strong>Simpan Arsip</strong> di Menu IKSS dan Form SOP.
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setDocSourceFilter("all")}
              className={`px-3 py-1.5 rounded-xl text-xs font-bold border transition cursor-pointer ${
                docSourceFilter === "all"
                  ? "bg-slate-900 text-white border-slate-900"
                  : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
              }`}
            >
              Semua ({docArchiveCounts.total})
            </button>
            <button
              type="button"
              onClick={() => setDocSourceFilter("ikss")}
              className={`px-3 py-1.5 rounded-xl text-xs font-bold border transition cursor-pointer ${
                docSourceFilter === "ikss"
                  ? "bg-emerald-600 text-white border-emerald-600"
                  : "bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100"
              }`}
            >
              Menu IKSS ({docArchiveCounts.ikss})
            </button>
            <button
              type="button"
              onClick={() => setDocSourceFilter("sop")}
              className={`px-3 py-1.5 rounded-xl text-xs font-bold border transition cursor-pointer ${
                docSourceFilter === "sop"
                  ? "bg-cyan-600 text-white border-cyan-600"
                  : "bg-cyan-50 text-cyan-700 border-cyan-200 hover:bg-cyan-100"
              }`}
            >
              Form SOP ({docArchiveCounts.sop})
            </button>
            <div className="relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                type="text"
                value={docSearchQuery}
                onChange={(e) => setDocSearchQuery(e.target.value)}
                placeholder="Cari arsip IKSS / SOP..."
                className="pl-8 pr-3 py-1.5 rounded-xl border border-slate-200 text-xs bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-teal-500/30 w-48 sm:w-56"
              />
            </div>
          </div>
        </div>

        {filteredDocArchives.length === 0 ? (
          <div className="p-6 text-center rounded-xl border border-dashed border-slate-200 bg-slate-50/50 text-xs text-slate-500">
            {docArchives.length === 0
              ? 'Belum ada arsip Menu IKSS atau Form SOP tersimpan. Buka tab Menu IKSS atau Form SOP, pilih kelompok sasaran yang diinginkan, lalu klik "Simpan Arsip".'
              : "Tidak ada arsip Menu IKSS / Form SOP yang cocok dengan filter atau pencarian."}
          </div>
        ) : (
          <div className="space-y-3">
            {filteredDocArchives.map((arc) => {
              const isExpanded = Boolean(expandedDocArchiveIds[arc.id]);
              const isIkss = arc.source === "ikss";
              return (
                <div
                  key={arc.id}
                  className="rounded-xl border border-slate-200 bg-slate-50/50 overflow-hidden"
                >
                  <div className="p-4 flex flex-col lg:flex-row lg:items-center justify-between gap-3 bg-white">
                    <div className="space-y-1 min-w-0">
                      <div className="flex flex-wrap items-center gap-2">
                        <span
                          className={`px-2 py-0.5 rounded-md text-[10px] font-extrabold uppercase tracking-wider border ${
                            isIkss
                              ? "bg-emerald-50 text-emerald-700 border-emerald-200"
                              : "bg-cyan-50 text-cyan-700 border-cyan-200"
                          }`}
                        >
                          {isIkss ? "Menu IKSS" : "Form SOP"}
                        </span>
                        <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-200">
                          {(arc.worksheets || []).length} Lembar Kerja
                        </span>
                        {arc.periodeLabel && (
                          <span className="px-2 py-0.5 rounded-md text-[10px] font-semibold bg-slate-100 text-slate-700">
                            {arc.periodeLabel}
                          </span>
                        )}
                      </div>
                      <h4 className="font-bold text-slate-900 text-xs sm:text-sm truncate">
                        {arc.title}
                      </h4>
                      <div className="flex flex-wrap items-center gap-3 text-[11px] text-slate-500">
                        <span className="flex items-center gap-1">
                          <Users className="w-3.5 h-3.5 text-slate-400" />
                          {arc.sasaranSummary || "Umum"}
                        </span>
                        <span className="flex items-center gap-1">
                          <Clock className="w-3.5 h-3.5 text-slate-400" />
                          Disimpan: {formatDocSavedTimestamp(arc)}
                        </span>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                      <button
                        type="button"
                        onClick={() =>
                          setExpandedDocArchiveIds((prev) => ({
                            ...prev,
                            [arc.id]: !prev[arc.id],
                          }))
                        }
                        className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold inline-flex items-center gap-1 transition cursor-pointer"
                      >
                        {isExpanded ? (
                          <>
                            <ChevronDown className="w-3.5 h-3.5" />
                            <span>Tutup Tabel</span>
                          </>
                        ) : (
                          <>
                            <ChevronRight className="w-3.5 h-3.5" />
                            <span>Pratinjau Tabel</span>
                          </>
                        )}
                      </button>

                      <button
                        type="button"
                        onClick={() => handlePrintDocArchive(arc)}
                        className="px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 rounded-xl text-xs font-bold inline-flex items-center gap-1 transition cursor-pointer"
                      >
                        <Printer className="w-3.5 h-3.5" />
                        <span>Cetak</span>
                      </button>

                      {onOpenDocArchive && (
                        <button
                          type="button"
                          onClick={() => onOpenDocArchive(arc)}
                          className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold inline-flex items-center gap-1 transition cursor-pointer"
                        >
                          <ArrowUpRight className="w-3.5 h-3.5" />
                          <span>Buka ke Lembar Kerja</span>
                        </button>
                      )}

                      {setDocArchives && (
                        <button
                          type="button"
                          onClick={() => setConfirmDeleteDocArchive(arc)}
                          title="Hapus Arsip"
                          className="p-1.5 text-rose-500 hover:bg-rose-50 rounded-xl transition cursor-pointer"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  </div>

                  {isExpanded && (
                    <div
                      id={`print-doc-archive-${arc.id}`}
                      className="p-4 border-t border-slate-200 bg-slate-50/70 space-y-5"
                    >
                      <div className="border-b border-slate-300 pb-2">
                        <h5 className="font-extrabold text-sm text-slate-900">{arc.title}</h5>
                        <p className="text-[11px] text-slate-600">
                          {arc.sppgName ? `${arc.sppgName} • ` : ""}
                          {arc.idSppg ? `ID SPPG: ${arc.idSppg} • ` : ""}
                          Periode: {arc.periodeLabel || "-"}
                        </p>
                      </div>

                      {(arc.worksheets || []).map((ws, wsIdx) => {
                        const cols = ws.columns || ws.headers || [];
                        return (
                          <div
                            key={ws.id || `${arc.id}_ws_${wsIdx}`}
                            className={`bg-white rounded-xl border border-slate-200 p-3.5 space-y-2 ${
                              wsIdx > 0 ? "print:break-before-page" : ""
                            }`}
                          >
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <div>
                                <div className="text-xs font-extrabold text-slate-900">
                                  {ws.title}
                                </div>
                                {(ws.subLabel || ws.subtitle || ws.sasaranLabel) && (
                                  <div className="text-[11px] text-slate-500">
                                    {[ws.sasaranLabel, ws.subLabel || ws.subtitle]
                                      .filter(Boolean)
                                      .join(" · ")}
                                  </div>
                                )}
                              </div>
                              <span className="text-[10px] font-bold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-md">
                                Lembar {wsIdx + 1} dari {arc.worksheets.length}
                              </span>
                            </div>

                            <div className="overflow-x-auto">
                              <table className="w-full text-[11px] text-left border-collapse">
                                {cols.length > 0 && (
                                  <thead className="bg-slate-100 text-slate-800">
                                    <tr>
                                      {cols.map((col, ci) => (
                                        <th
                                          key={ci}
                                          className="p-2 border border-slate-300 font-bold whitespace-nowrap"
                                        >
                                          {col}
                                        </th>
                                      ))}
                                    </tr>
                                  </thead>
                                )}
                                <tbody>
                                  {(ws.rows || []).length === 0 ? (
                                    <tr>
                                      <td
                                        colSpan={Math.max(cols.length, 1)}
                                        className="p-3 border border-slate-200 text-center text-slate-400 italic"
                                      >
                                        Tidak ada baris data pada lembar kerja ini.
                                      </td>
                                    </tr>
                                  ) : (
                                    (ws.rows || []).map((row, ri) => (
                                      <tr key={ri} className="hover:bg-slate-50/70">
                                        {row.map((cell, ci) => (
                                          <td
                                            key={ci}
                                            className="p-1.5 border border-slate-200 text-slate-700 align-top"
                                          >
                                            {cell === null || cell === undefined || cell === ""
                                              ? "-"
                                              : String(cell)}
                                          </td>
                                        ))}
                                      </tr>
                                    ))
                                  )}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* RIWAYAT SNAPSHOT ARSIP DOKUMEN CLOUD (BUKA / PULIHKAN PERENCANAAN) */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <button
            type="button"
            onClick={() => setShowSavedHistory((v) => !v)}
            className="flex items-center gap-2.5 text-left cursor-pointer"
          >
            <div className="p-2 bg-emerald-50 rounded-xl text-emerald-600 shrink-0">
              <Clock className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-900 flex items-center gap-1.5">
                Daftar Riwayat Arsip Perencanaan Cloud
                {showSavedHistory ? (
                  <ChevronDown className="w-4 h-4 text-slate-400" />
                ) : (
                  <ChevronRight className="w-4 h-4 text-slate-400" />
                )}
              </h3>
              <p className="text-xs text-slate-500">
                Otomatis terarsip setiap pergantian tanggal siklus di Master Menu atau saat klik Simpan Snapshot.
              </p>
            </div>
          </button>
          <span className="px-2.5 py-1 rounded-lg bg-indigo-50 text-indigo-700 text-xs font-bold tabular-nums">
            {savedPeriods.length} Arsip Tersimpan
          </span>
        </div>

        {showSavedHistory && (
          <div className="space-y-2.5">
            {savedPeriods.length === 0 ? (
              <div className="p-6 text-center rounded-xl border border-dashed border-slate-200 bg-slate-50/50 text-xs text-slate-500">
                Belum ada riwayat arsip perencanaan tersimpan. Klik <strong>Simpan Snapshot ke Arsip Dokumen</strong> atau ubah tanggal siklus di Master Menu untuk membuat arsip otomatis.
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {savedPeriods.map((period: any) => {
                  const totalSchools = period.sekolahPM?.length || 0;
                  const totalVillages = period.tigaBPM?.length || 0;
                  const dates: string[] = period.profile?.periodeDates || [];
                  const firstDate = dates[0] ? formatIndonesianDate(dates[0]) : "-";
                  const lastDate = dates[dates.length - 1]
                    ? formatIndonesianDate(dates[dates.length - 1])
                    : "-";
                  return (
                    <div
                      key={period.id}
                      className="p-3.5 rounded-xl border border-slate-200 hover:border-slate-300 bg-slate-50/40 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                    >
                      <div className="space-y-1 min-w-0">
                        <h4 className="font-bold text-slate-900 text-xs truncate">
                          {period.namaPeriode}
                        </h4>
                        <p className="text-[11px] text-slate-500 flex items-center gap-1">
                          <Clock className="w-3 h-3 text-slate-400 shrink-0" />
                          <span>Diarsipkan: {period.createdAt}</span>
                        </p>
                        <div className="flex flex-wrap gap-1.5 pt-1">
                          <span className="inline-flex items-center gap-1 bg-white border border-slate-200 text-slate-600 text-[10px] font-semibold px-2 py-0.5 rounded-md tabular-nums">
                            <Users className="w-3 h-3 text-slate-400" />
                            {totalSchools} Sekolah • {totalVillages} Desa
                          </span>
                          <span className="inline-flex items-center gap-1 bg-indigo-50 text-indigo-700 text-[10px] font-semibold px-2 py-0.5 rounded-md">
                            <Calendar className="w-3 h-3 text-indigo-400" />
                            {firstDate} s/d {lastDate}
                          </span>
                        </div>
                      </div>

                      <div className="flex items-center gap-1.5 shrink-0 self-end sm:self-center">
                        {setSavedPeriods && (
                          <button
                            type="button"
                            onClick={() =>
                              setConfirmPeriodAction({ kind: "delete", period })
                            }
                            title="Hapus Arsip"
                            className="p-2 text-rose-500 hover:bg-rose-50 rounded-xl transition cursor-pointer"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        )}
                        {loadPeriod && (
                          <button
                            type="button"
                            onClick={() =>
                              setConfirmPeriodAction({ kind: "load", period })
                            }
                            className="py-1.5 px-3 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl transition cursor-pointer inline-flex items-center gap-1.5"
                          >
                            <FolderOpen className="w-3.5 h-3.5" />
                            <span>Buka Arsip</span>
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
      </div>

      {/* IN-APP CONFIRMATION MODAL FOR DELETING SAVED IKSS / SOP ARCHIVE */}
      {confirmDeleteDocArchive && (
        <div className="fixed inset-0 z-[99999] bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h4 className="text-base font-bold text-slate-900">
                  Hapus Arsip Lembar Kerja?
                </h4>
                <p className="text-xs text-slate-500 mt-1">
                  Arsip &ldquo;{confirmDeleteDocArchive.title}&rdquo; akan dihapus permanen dari Menu Arsip.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setConfirmDeleteDocArchive(null)}
                className="p-1 text-slate-400 hover:text-slate-700 rounded-lg cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setConfirmDeleteDocArchive(null)}
                className="px-4 py-2 rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-50 text-xs font-bold cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={() => {
                  if (setDocArchives && confirmDeleteDocArchive) {
                    setDocArchives((prev) =>
                      prev.filter((item) => item.id !== confirmDeleteDocArchive.id)
                    );
                    setSnapshotMessage({
                      type: "success",
                      text: `Arsip "${confirmDeleteDocArchive.title}" telah dihapus.`,
                    });
                  }
                  setConfirmDeleteDocArchive(null);
                }}
                className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold cursor-pointer"
              >
                Ya, Hapus Permanen
              </button>
            </div>
          </div>
        </div>
      )}

      {/* IN-APP CONFIRMATION MODAL FOR OPENING OR DELETING SAVED ARCHIVE */}
      {confirmPeriodAction && (
        <div className="fixed inset-0 z-[99999] bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h4 className="text-base font-bold text-slate-900">
                  {confirmPeriodAction.kind === "load"
                    ? "Buka Arsip Perencanaan?"
                    : "Hapus Arsip Permanen?"}
                </h4>
                <p className="text-xs text-slate-500 mt-1">
                  {confirmPeriodAction.kind === "load"
                    ? `Lembar kerja aktif akan memuat arsip "${confirmPeriodAction.period.namaPeriode}". Pastikan perencanaan aktif saat ini sudah tersimpan jika masih diperlukan.`
                    : `Arsip "${confirmPeriodAction.period.namaPeriode}" akan dihapus permanen dari daftar arsip cloud.`}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setConfirmPeriodAction(null)}
                className="p-1 text-slate-400 hover:text-slate-700 rounded-lg cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setConfirmPeriodAction(null)}
                className="px-4 py-2 rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-50 text-xs font-bold cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={executeConfirmedPeriodAction}
                className={`px-4 py-2 rounded-xl text-white text-xs font-bold cursor-pointer ${
                  confirmPeriodAction.kind === "load"
                    ? "bg-emerald-600 hover:bg-emerald-700"
                    : "bg-rose-600 hover:bg-rose-700"
                }`}
              >
                {confirmPeriodAction.kind === "load"
                  ? "Ya, Buka Arsip"
                  : "Ya, Hapus Permanen"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* EXPLICIT CONFIRMATION MODAL BEFORE MUTATING GOOGLE DRIVE */}
      {isConfirmModalOpen && driveSession && (
        <div className="fixed inset-0 z-[99999] bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white w-full max-w-2xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
            <div className="p-5 bg-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <FolderPlus className="w-5 h-5 text-emerald-400" />
                <div>
                  <h3 className="font-extrabold text-sm">Konfirmasi Kirim Arsip PDF ke Google Drive</h3>
                  <p className="text-[11px] text-slate-300">
                    Pastikan daftar folder & berkas PDF di bawah sudah sesuai sebelum dikirim.
                  </p>
                </div>
              </div>
              <button
                type="button"
                disabled={isUploadingToDrive}
                onClick={() => setIsConfirmModalOpen(false)}
                className="p-1 text-slate-300 hover:text-white cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-4 text-xs">
              <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 space-y-1">
                <div className="font-bold text-emerald-900">
                  Email Google Drive Tujuan: <span className="underline">{driveSession.email}</span>
                </div>
                <div className="text-emerald-800">
                  Lokasi Folder Utama:{" "}
                  <strong>
                    Arsip SISPERMBG /{" "}
                    {activeMonth === "all" ? "Semua Bulan" : `Arsip ${formatIndonesianMonthYear(activeMonth)}`}
                  </strong>
                </div>
                <div className="text-emerald-700">
                  Mode Folder:{" "}
                  <strong>
                    {separateByMenuFolder
                      ? `Folder terpisah otomatis per Nama Menu (${uniqueMenuFoldersPreview.length} folder menu)`
                      : "Gabung dalam satu folder bulan"}
                  </strong>
                </div>
              </div>

              {separateByMenuFolder && (
                <div className="space-y-1.5">
                  <div className="font-bold text-slate-700">
                    Struktur Folder Menu yang Akan Dibuat / Dipakai:
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {uniqueMenuFoldersPreview.map((folderName) => (
                      <span
                        key={folderName}
                        className="px-2.5 py-1 bg-amber-50 border border-amber-200 text-amber-900 rounded-lg font-bold inline-flex items-center gap-1"
                      >
                        <Folder className="w-3.5 h-3.5 text-amber-600" />
                        {folderName}
                      </span>
                    ))}
                  </div>
                </div>
              )}

              <div className="space-y-1.5">
                <div className="font-bold text-slate-700">
                  Daftar File PDF yang Akan Dikonversi & Diunggah ({selectedUploadTasks.length} Berkas PDF):
                </div>
                <div className="max-h-56 overflow-y-auto border border-slate-200 rounded-xl divide-y divide-slate-100">
                  {selectedUploadTasks.map((t) => (
                    <div key={t.key} className="p-2.5 flex items-center justify-between gap-2 hover:bg-slate-50">
                      <div className="flex items-center gap-2 min-w-0">
                        <FileText className="w-4 h-4 text-rose-500 shrink-0" />
                        <div className="min-w-0">
                          <div className="font-bold text-slate-800 truncate">{t.fileName}</div>
                          <div className="text-[10px] text-slate-500 truncate">
                            {separateByMenuFolder ? `Folder: ${t.menuFolderName} • ` : ""}
                            {formatIndonesianDate(t.entry.dateStr)}
                          </div>
                        </div>
                      </div>
                      <button
                        type="button"
                        disabled={isUploadingToDrive}
                        onClick={() => toggleItemSelection(t.entry.id, t.category)}
                        className="text-[10px] font-bold text-rose-600 hover:underline shrink-0 cursor-pointer"
                      >
                        Hapus
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2.5">
              <button
                type="button"
                disabled={isUploadingToDrive}
                onClick={() => setIsConfirmModalOpen(false)}
                className="px-4 py-2 bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 rounded-xl text-xs font-bold cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                disabled={isUploadingToDrive || selectedUploadTasks.length === 0}
                onClick={handleExecuteDriveUpload}
                className="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-black flex items-center gap-1.5 shadow-sm cursor-pointer disabled:opacity-50"
              >
                {isUploadingToDrive ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    Mengonversi & Mengirim...
                  </>
                ) : (
                  <>
                    <Check className="w-4 h-4" />
                    Ya, Konfirmasi & Kirim PDF ke Drive
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
