/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from "react";
import {
  HariPM,
  KelompokSasaranPM,
  SPPGProfile,
  SekolahPM,
  TigaBPM,
  UnifiedKopConfig,
  DEFAULT_UNIFIED_KOP_CONFIG,
  UraianKategoriPM,
  UraianSasaranPM,
} from "../types";
import { EMPTY_SASARAN_LIST, normalizeSasaranList } from "../initialData";
import { formatIndonesianDate } from "../utils/menuSchedule";
import { computeSasaranBreakdown } from "../utils/calc";
import {
  Calendar,
  CalendarDays,
  Download,
  Upload,
  Copy,
  RotateCcw,
  AlertTriangle,
  Users,
  Settings,
  ChevronDown,
  ChevronUp,
  Printer,
  Image as ImageIcon,
  Eye,
  Plus,
  Trash2,
  Edit3,
  School,
  Home,
  GraduationCap,
  FlaskConical,
  ArrowUp,
  ArrowDown,
  RefreshCw,
  Check,
  X,
  Table2,
  Database,
} from "lucide-react";
import * as XLSX from "xlsx";
import { downloadElementAsImage, exportToPDF, printElementSmart } from "../lib/printUtils";
import { KopSuratRenderHeader, LogoCrop } from "./KopSuratConfigSection";
import AutoSignatureBlock from "./AutoSignatureBlock";
import PrintPreviewModal from "./PrintPreviewModal";

interface PenerimaManfaatTabProps {
  harianPM: HariPM[];
  onChange: (updated: HariPM[]) => void;
  profile?: SPPGProfile;
  sekolahPM?: SekolahPM[];
  tigaBPM?: TigaBPM[];
  pmSettings?: {
    porsiKecilHarga: number;
    porsiBesarHarga: number;
    porsiKecilSasaranIds: string[];
    porsiBesarSasaranIds: string[];
  };
  setPmSettings?: (updated: any) => void;
  kopLine1?: string;
  setKopLine1?: (val: string) => void;
  kopLine2?: string;
  setKopLine2?: (val: string) => void;
  kopLine3?: string;
  setKopLine3?: (val: string) => void;
  kopLine4?: string;
  setKopLine4?: (val: string) => void;
  leftLogo?: string;
  setLeftLogo?: (val: string) => void;
  rightLogo?: string;
  setRightLogo?: (val: string) => void;
  leftLogoCrop?: LogoCrop;
  setLeftLogoCrop?: (val: LogoCrop) => void;
  rightLogoCrop?: LogoCrop;
  setRightLogoCrop?: (val: LogoCrop) => void;
  paperSize?: "A4" | "F4";
  setPaperSize?: (val: "A4" | "F4") => void;
  kopConfig?: UnifiedKopConfig;
  isAdmin?: boolean;
}

type InputModeTab = "sekolah" | "tiga_b" | "tenaga" | "sampel";
type SampelChoiceMode = "none" | "besar" | "kecil" | "both";
type PrintScopeMode = "both" | "table1" | "table2";

const KATEGORI_BADGE_MAP: Record<UraianKategoriPM, { label: string; color: string }> = {
  tk_paud_lb: { label: "TK / PAUD / LB", color: "bg-sky-50 text-sky-700 border-sky-200" },
  sd_mi_slb: { label: "SD / MI / SLB", color: "bg-indigo-50 text-indigo-700 border-indigo-200" },
  smp_mts_smplb: { label: "SMP / MTS", color: "bg-violet-50 text-violet-700 border-violet-200" },
  sma_smk_ma: { label: "SMA / SMK / MA", color: "bg-purple-50 text-purple-700 border-purple-200" },
  pendidik: { label: "Pendidik", color: "bg-teal-50 text-teal-700 border-teal-200" },
  tenaga_kependidikan: { label: "Tenaga Pendidikan", color: "bg-slate-100 text-slate-700 border-slate-200" },
  posyandu_3b: { label: "3B Desa / Posyandu", color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
  sampel: { label: "Sampel", color: "bg-blue-50 text-blue-700 border-blue-200" },
};

function detectSchoolCategoryFromName(name: string): UraianKategoriPM {
  const upper = name.toUpperCase();
  if (/\b(PAUD|TK|KB|RA|TKN|TK-IT)\b/.test(upper)) return "tk_paud_lb";
  if (/\b(SMP|SMPN|MTS|SMPLB)\b/.test(upper)) return "smp_mts_smplb";
  if (/\b(SMA|SMAN|SMK|SMKS|SMKN|MA|MAS|MAN)\b/.test(upper)) return "sma_smk_ma";
  if (/\b(SD|SDN|MI|MIN|MIS|SLB)\b/.test(upper)) return "sd_mi_slb";
  return "sd_mi_slb";
}

/**
 * Rebuilds Tabel 2 (Rekapitulasi Keseluruhan — 13 kelompok sasaran termasuk Sampel)
 * accurately from Tabel 1 (Uraian Kelompok Sasaran per sekolah, 3B desa/posyandu, tenaga, sampel)
 * so Tabel 1 Sub. Total === Tabel 2 Sub. Total down to every single portion.
 */
function recalculateSasaranFromUraian(uraianList: UraianSasaranPM[]): KelompokSasaranPM[] {
  const map = new Map<string, KelompokSasaranPM>();
  EMPTY_SASARAN_LIST.forEach((def) => {
    map.set(def.id, { ...def, porsiBesar: 0, porsiKecil: 0, alergiBesar: 0, alergiKecil: 0 });
  });

  const addTo = (
    id: string,
    pb: number,
    pk: number,
    ab: number,
    ak: number
  ) => {
    const target = map.get(id);
    if (!target) return;
    target.porsiBesar += Math.max(0, Number(pb) || 0);
    target.porsiKecil += Math.max(0, Number(pk) || 0);
    target.alergiBesar += Math.max(0, Number(ab) || 0);
    target.alergiKecil += Math.max(0, Number(ak) || 0);
  };

  uraianList.forEach((u) => {
    const pb = Math.max(0, Number(u.porsiBesar) || 0);
    const pk = Math.max(0, Number(u.porsiKecil) || 0);
    const ab = Math.max(0, Number(u.alergiBesar) || 0);
    const ak = Math.max(0, Number(u.alergiKecil) || 0);
    const sb = Math.max(0, Number(u.sampelBesar) || 0);
    const sk = Math.max(0, Number(u.sampelKecil) || 0);

    switch (u.kategori) {
      case "tk_paud_lb":
        addTo("tk_paud_lb", pb, pk, ab, ak);
        break;
      case "sd_mi_slb":
        addTo("sd_kelas_1_3", 0, pk, 0, ak);
        addTo("sd_kelas_4_6", pb, 0, ab, 0);
        break;
      case "smp_mts_smplb":
        addTo("smp_mts_smplb", pb, pk, ab, ak);
        break;
      case "sma_smk_ma":
        addTo("sma_smk_ma", pb, pk, ab, ak);
        break;
      case "pendidik":
        addTo("pendidik", pb, pk, ab, ak);
        break;
      case "tenaga_kependidikan":
        addTo("tenaga_kependidikan", pb, pk, ab, ak);
        break;
      case "posyandu_3b": {
        const balitaId = u.balitaTargetId || "anak_balita";
        addTo(balitaId, 0, pk, 0, ak);

        const rawBumil = Math.max(0, Number(u.bumil) || 0);
        const rawBusui = Math.max(0, Number(u.busui) || 0);
        if (rawBumil + rawBusui === pb && pb > 0) {
          addTo("ibu_hamil", rawBumil, 0, ab, 0);
          addTo("ibu_menyusui", rawBusui, 0, 0, 0);
        } else if (pb > 0) {
          const busuiAllocated = Math.min(rawBusui, pb);
          const bumilAllocated = Math.max(0, pb - busuiAllocated);
          addTo("ibu_hamil", bumilAllocated, 0, ab, 0);
          addTo("ibu_menyusui", busuiAllocated, 0, 0, 0);
        } else if (ab > 0) {
          addTo("ibu_hamil", 0, 0, ab, 0);
        }
        break;
      }
      case "sampel":
        addTo("sampel", pb + sb, pk + sk, ab, ak);
        return; // already included sb & sk
      default:
        addTo("tenaga_kependidikan", pb, pk, ab, ak);
        break;
    }

    // Any sample portions tagged on school/3B/tenaga rows roll up into Row 13 "Sampel"
    if (sb > 0 || sk > 0) {
      addTo("sampel", sb, sk, 0, 0);
    }
  });

  return EMPTY_SASARAN_LIST.map((def) => map.get(def.id) || { ...def });
}

export default function PenerimaManfaatTab({
  harianPM,
  onChange,
  profile,
  sekolahPM = [],
  tigaBPM = [],
  pmSettings,
  setPmSettings,
  kopLine1 = "BADAN GIZI NASIONAL",
  kopLine2 = "SATUAN PELAYANAN PEMENUHAN GIZI",
  kopLine3 = "SPPG MUNA BARAT SAWERIGADI ONDOKE",
  kopLine4 = "Alamat : Jln. Poros Lagadi-Tondasi, Desa Ondoke, Kec. Sawerigadi, Kab. Muna Barat",
  leftLogo = "/assets/images/app_logo_sppg.png",
  rightLogo = "",
  leftLogoCrop = { top: 0, bottom: 0, left: 0, right: 0 },
  rightLogoCrop = { top: 0, bottom: 0, left: 0, right: 0 },
  paperSize = "A4",
  kopConfig = DEFAULT_UNIFIED_KOP_CONFIG,
}: PenerimaManfaatTabProps) {
  const [selectedDay, setSelectedDay] = useState<number>(1);
  const [alertMsg, setAlertMsg] = useState<{ type: "success" | "info" | "error"; text: string } | null>(null);
  const [isDownloading, setIsDownloading] = useState<string | null>(null);
  const [showPrintModal, setShowPrintModal] = useState<boolean>(false);
  const [confirmAction, setConfirmAction] = useState<"copy_all" | "reset_day" | null>(null);

  // Display / print options
  const [showSampelColInTable1, setShowSampelColInTable1] = useState<boolean>(true);
  const [printScope, setPrintScope] = useState<PrintScopeMode>("both");
  const [includeRabInPrint, setIncludeRabInPrint] = useState<boolean>(false);

  // Accordion state for classification editing
  const [showKecilSasaranEditor, setShowKecilSasaranEditor] = useState(false);
  const [showBesarSasaranEditor, setShowBesarSasaranEditor] = useState(false);

  // Dedicated Input Panel Form State
  const [inputMode, setInputMode] = useState<InputModeTab>("sekolah");
  const [editingRowId, setEditingRowId] = useState<string | null>(null);
  const [formNama, setFormNama] = useState<string>("");
  const [formKategori, setFormKategori] = useState<UraianKategoriPM>("sd_mi_slb");
  const [formBalitaTarget, setFormBalitaTarget] = useState<"anak_balita" | "anak_balita_13_59" | "balita_6_11">("anak_balita");
  const [formBumil, setFormBumil] = useState<string>("");
  const [formBusui, setFormBusui] = useState<string>("");
  const [formPorsiBesar, setFormPorsiBesar] = useState<string>("");
  const [formPorsiKecil, setFormPorsiKecil] = useState<string>("");
  const [formAlergiBesar, setFormAlergiBesar] = useState<string>("");
  const [formAlergiKecil, setFormAlergiKecil] = useState<string>("");
  const [formSampelChoice, setFormSampelChoice] = useState<SampelChoiceMode>("none");
  const [formSampelBesar, setFormSampelBesar] = useState<string>("");
  const [formSampelKecil, setFormSampelKecil] = useState<string>("");

  const totalDays = Math.max(profile?.periodeDates?.length || 0, 1);
  const effectiveTotalDays =
    profile?.periodeDates && profile.periodeDates.length > 0
      ? totalDays
      : Math.max(harianPM.length, 10);

  useEffect(() => {
    if (selectedDay > effectiveTotalDays) {
      setSelectedDay(effectiveTotalDays);
    }
  }, [effectiveTotalDays, selectedDay]);

  const rawActiveDay = harianPM.find((h) => h.hariKe === selectedDay) || {
    hariKe: selectedDay,
    sasaran: EMPTY_SASARAN_LIST,
    uraian: [],
  };

  const activeSasaran = useMemo(
    () => normalizeSasaranList(rawActiveDay.sasaran),
    [rawActiveDay.sasaran]
  );

  const activeUraian: UraianSasaranPM[] = useMemo(
    () => (Array.isArray(rawActiveDay.uraian) ? rawActiveDay.uraian : []),
    [rawActiveDay.uraian]
  );

  const activeDateStr = profile?.periodeDates?.[selectedDay - 1] || "";
  const formattedActiveDateBanner = activeDateStr
    ? formatIndonesianDate(activeDateStr)
    : `Hari Kerja Ke-${selectedDay}`;

  // Fallbacks for dynamic PM settings
  const defaultSettings = {
    porsiKecilHarga: 8000,
    porsiBesarHarga: 10000,
    porsiKecilSasaranIds: ["tk_paud_lb", "sd_kelas_1_3", "anak_balita", "anak_balita_13_59", "balita_6_11", "sampel"],
    porsiBesarSasaranIds: [
      "sd_kelas_4_6",
      "smp_mts_smplb",
      "sma_smk_ma",
      "pendidik",
      "tenaga_kependidikan",
      "ibu_hamil",
      "ibu_menyusui",
      "sampel",
    ],
  };

  const settings = pmSettings || defaultSettings;

  const updateSetting = (key: string, value: any) => {
    if (setPmSettings) {
      setPmSettings({
        ...settings,
        [key]: value,
      });
    }
  };

  const handleToggleKecilSasaran = (sasaranId: string) => {
    const alreadySelected = settings.porsiKecilSasaranIds.includes(sasaranId);
    const updatedIds = alreadySelected
      ? settings.porsiKecilSasaranIds.filter((id) => id !== sasaranId)
      : [...settings.porsiKecilSasaranIds, sasaranId];
    updateSetting("porsiKecilSasaranIds", updatedIds);
  };

  const handleToggleBesarSasaran = (sasaranId: string) => {
    const alreadySelected = settings.porsiBesarSasaranIds.includes(sasaranId);
    const updatedIds = alreadySelected
      ? settings.porsiBesarSasaranIds.filter((id) => id !== sasaranId)
      : [...settings.porsiBesarSasaranIds, sasaranId];
    updateSetting("porsiBesarSasaranIds", updatedIds);
  };

  const triggerAlert = (type: "success" | "info" | "error", text: string) => {
    setAlertMsg({ type, text });
    setTimeout(() => {
      setAlertMsg((prev) => (prev?.text === text ? null : prev));
    }, 4500);
  };

  // Persist active day change cleanly (ensuring day exists up to effectiveTotalDays)
  const commitActiveDayUpdate = (
    nextSasaran: KelompokSasaranPM[],
    nextUraian: UraianSasaranPM[]
  ) => {
    const existingLen = Math.max(harianPM.length, selectedDay);
    const updatedHarian: HariPM[] = Array.from({ length: existingLen }, (_, i) => {
      const dayNum = i + 1;
      const found = harianPM.find((h) => h.hariKe === dayNum);
      if (dayNum === selectedDay) {
        return {
          hariKe: dayNum,
          sasaran: normalizeSasaranList(nextSasaran),
          uraian: nextUraian,
        };
      }
      if (found) {
        return {
          ...found,
          sasaran: normalizeSasaranList(found.sasaran),
          uraian: Array.isArray(found.uraian) ? found.uraian : [],
        };
      }
      return {
        hariKe: dayNum,
        sasaran: EMPTY_SASARAN_LIST.map((s) => ({ ...s })),
        uraian: [],
      };
    });
    onChange(updatedHarian);
  };

  // Calculations for Active Day (Harian — clean 5-bucket & Tabel 2 column sums)
  const dailyBreakdown = useMemo(
    () => computeSasaranBreakdown(activeSasaran, settings.porsiKecilHarga, settings.porsiBesarHarga),
    [activeSasaran, settings.porsiKecilHarga, settings.porsiBesarHarga]
  );

  const totalPorsiBesar = dailyBreakdown.totalPorsiBesar;
  const totalPorsiKecil = dailyBreakdown.totalPorsiKecil;
  const totalAlergiBesar = dailyBreakdown.totalAlergiBesar;
  const totalAlergiKecil = dailyBreakdown.totalAlergiKecil;
  const totalAlergi = dailyBreakdown.totalAlergi;
  const totalPM = dailyBreakdown.totalPM;

  const sampelPorsiBesar = dailyBreakdown.sampelBesar;
  const sampelPorsiKecil = dailyBreakdown.sampelKecil;
  const totalSampel = dailyBreakdown.totalSampel;

  // Calculations for Tabel 1 (Uraian Kelompok Sasaran)
  const uraianTotals = useMemo(() => {
    let pb = 0;
    let pk = 0;
    let ab = 0;
    let ak = 0;
    let sb = 0;
    let sk = 0;
    activeUraian.forEach((u) => {
      pb += Number(u.porsiBesar) || 0;
      pk += Number(u.porsiKecil) || 0;
      ab += Number(u.alergiBesar) || 0;
      ak += Number(u.alergiKecil) || 0;
      sb += Number(u.sampelBesar) || 0;
      sk += Number(u.sampelKecil) || 0;
    });
    return {
      pb,
      pk,
      ab,
      ak,
      sb,
      sk,
      grandTotal: pb + pk + ab + ak + sb + sk,
    };
  }, [activeUraian]);

  // Dynamic RAB calculations (Column sums matching Tabel 2)
  const rabPorsiKecil = dailyBreakdown.rabPorsiKecil;
  const rabPorsiBesar = dailyBreakdown.rabPorsiBesar;
  const rabAlergiKecil = dailyBreakdown.rabAlergiKecil;
  const rabAlergiBesar = dailyBreakdown.rabAlergiBesar;
  const totalRABHarian = dailyBreakdown.totalRAB;

  const formatIDR = (num: number) =>
    new Intl.NumberFormat("id-ID", {
      style: "currency",
      currency: "IDR",
      maximumFractionDigits: 0,
    }).format(num);

  // Switch mode tab in Dedicated Input Panel
  const handleSelectInputMode = (mode: InputModeTab) => {
    setInputMode(mode);
    setEditingRowId(null);
    setFormNama("");
    setFormBumil("");
    setFormBusui("");
    setFormPorsiBesar("");
    setFormPorsiKecil("");
    setFormAlergiBesar("");
    setFormAlergiKecil("");

    if (mode === "sekolah") {
      setFormKategori("sd_mi_slb");
      setFormSampelChoice("none");
      setFormSampelBesar("");
      setFormSampelKecil("");
    } else if (mode === "tiga_b") {
      setFormKategori("posyandu_3b");
      setFormBalitaTarget("anak_balita");
      setFormSampelChoice("none");
      setFormSampelBesar("");
      setFormSampelKecil("");
    } else if (mode === "tenaga") {
      setFormKategori("tenaga_kependidikan");
      setFormNama("Tenaga Pendidikan");
      setFormSampelChoice("none");
      setFormSampelBesar("");
      setFormSampelKecil("");
    } else {
      setFormKategori("sampel");
      setFormNama("Sampel Makanan");
      setFormSampelChoice("both");
      setFormSampelBesar("1");
      setFormSampelKecil("1");
    }
  };

  const handleSampelChoiceChange = (choice: SampelChoiceMode) => {
    setFormSampelChoice(choice);
    if (choice === "none") {
      setFormSampelBesar("");
      setFormSampelKecil("");
    } else if (choice === "besar") {
      setFormSampelBesar((prev) => (Number(prev) > 0 ? prev : "1"));
      setFormSampelKecil("");
    } else if (choice === "kecil") {
      setFormSampelBesar("");
      setFormSampelKecil((prev) => (Number(prev) > 0 ? prev : "1"));
    } else {
      setFormSampelBesar((prev) => (Number(prev) > 0 ? prev : "1"));
      setFormSampelKecil((prev) => (Number(prev) > 0 ? prev : "1"));
    }
  };

  const resetInputForm = () => {
    setEditingRowId(null);
    setFormNama(inputMode === "tenaga" ? "Tenaga Pendidikan" : inputMode === "sampel" ? "Sampel Makanan" : "");
    setFormBumil("");
    setFormBusui("");
    setFormPorsiBesar("");
    setFormPorsiKecil("");
    setFormAlergiBesar("");
    setFormAlergiKecil("");
    if (inputMode === "sampel") {
      setFormSampelChoice("both");
      setFormSampelBesar("1");
      setFormSampelKecil("1");
    } else {
      setFormSampelChoice("none");
      setFormSampelBesar("");
      setFormSampelKecil("");
    }
  };

  // Add or Update row in Tabel 1 + auto-recalculate Tabel 2
  const handleSubmitUraianForm = (e: React.FormEvent) => {
    e.preventDefault();
    const cleanedName = formNama.trim();
    if (!cleanedName) {
      triggerAlert(
        "error",
        inputMode === "tiga_b"
          ? "Mohon isi Nama Desa / Posyandu (3B) terlebih dahulu."
          : inputMode === "sekolah"
          ? "Mohon isi Nama Sekolah terlebih dahulu."
          : "Mohon isi Nama Kelompok Sasaran terlebih dahulu."
      );
      return;
    }

    const parsedAlergiBesar = Math.max(0, parseInt(formAlergiBesar, 10) || 0);
    const parsedAlergiKecil = Math.max(0, parseInt(formAlergiKecil, 10) || 0);
    const parsedSampelBesar =
      formSampelChoice === "besar" || formSampelChoice === "both"
        ? Math.max(0, parseInt(formSampelBesar, 10) || 0)
        : 0;
    const parsedSampelKecil =
      formSampelChoice === "kecil" || formSampelChoice === "both"
        ? Math.max(0, parseInt(formSampelKecil, 10) || 0)
        : 0;

    let finalPorsiBesar = Math.max(0, parseInt(formPorsiBesar, 10) || 0);
    let finalPorsiKecil = Math.max(0, parseInt(formPorsiKecil, 10) || 0);
    let bumilVal: number | undefined;
    let busuiVal: number | undefined;

    if (inputMode === "tiga_b") {
      const b = Math.max(0, parseInt(formBumil, 10) || 0);
      const s = Math.max(0, parseInt(formBusui, 10) || 0);
      bumilVal = b;
      busuiVal = s;
      finalPorsiBesar = b + s;
    }

    const rowPayload: UraianSasaranPM = {
      id: editingRowId || `uraian_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      nama: cleanedName,
      tipeUnit: inputMode,
      kategori:
        inputMode === "tiga_b"
          ? "posyandu_3b"
          : inputMode === "sampel"
          ? "sampel"
          : formKategori,
      balitaTargetId: inputMode === "tiga_b" ? formBalitaTarget : undefined,
      bumil: bumilVal,
      busui: busuiVal,
      porsiBesar: inputMode === "sampel" ? 0 : finalPorsiBesar,
      porsiKecil: inputMode === "sampel" ? 0 : finalPorsiKecil,
      alergiBesar: parsedAlergiBesar,
      alergiKecil: parsedAlergiKecil,
      sampelBesar: parsedSampelBesar,
      sampelKecil: parsedSampelKecil,
    };

    const nextUraian = editingRowId
      ? activeUraian.map((u) => (u.id === editingRowId ? rowPayload : u))
      : [...activeUraian, rowPayload];

    const nextSasaran = recalculateSasaranFromUraian(nextUraian);
    commitActiveDayUpdate(nextSasaran, nextUraian);

    triggerAlert(
      "success",
      editingRowId
        ? `Data "${cleanedName}" diperbarui & Tabel 2 Rekapitulasi otomatis tersinkron.`
        : `"${cleanedName}" berhasil ditambahkan ke Tabel 1 & direkap otomatis ke Tabel 2!`
    );
    resetInputForm();
  };

  // Load a row into Panel Input Khusus for full editing
  const handleStartEditUraianRow = (row: UraianSasaranPM) => {
    setEditingRowId(row.id);
    setInputMode(row.tipeUnit);
    setFormNama(row.nama);
    setFormKategori(row.kategori);
    setFormBalitaTarget(row.balitaTargetId || "anak_balita");
    setFormBumil(row.bumil ? String(row.bumil) : "");
    setFormBusui(row.busui ? String(row.busui) : "");
    setFormPorsiBesar(row.porsiBesar ? String(row.porsiBesar) : "");
    setFormPorsiKecil(row.porsiKecil ? String(row.porsiKecil) : "");
    setFormAlergiBesar(row.alergiBesar ? String(row.alergiBesar) : "");
    setFormAlergiKecil(row.alergiKecil ? String(row.alergiKecil) : "");

    if (row.sampelBesar > 0 && row.sampelKecil > 0) {
      setFormSampelChoice("both");
    } else if (row.sampelBesar > 0) {
      setFormSampelChoice("besar");
    } else if (row.sampelKecil > 0) {
      setFormSampelChoice("kecil");
    } else {
      setFormSampelChoice("none");
    }
    setFormSampelBesar(row.sampelBesar ? String(row.sampelBesar) : "");
    setFormSampelKecil(row.sampelKecil ? String(row.sampelKecil) : "");
  };

  // Inline edit inside Tabel 1 -> auto-syncs Tabel 2 immediately
  const handleUraianInlineChange = (
    rowId: string,
    field: "nama" | "porsiBesar" | "porsiKecil" | "alergiBesar" | "alergiKecil" | "sampelBesar" | "sampelKecil",
    rawValue: string
  ) => {
    const nextUraian = activeUraian.map((u) => {
      if (u.id !== rowId) return u;
      if (field === "nama") {
        return { ...u, nama: rawValue };
      }
      const numVal = Math.max(0, parseInt(rawValue, 10) || 0);
      if (field === "porsiBesar" && u.kategori === "posyandu_3b") {
        const prevBusui = Math.max(0, Number(u.busui) || 0);
        const nextBusui = Math.min(prevBusui, numVal);
        const nextBumil = Math.max(0, numVal - nextBusui);
        return { ...u, porsiBesar: numVal, bumil: nextBumil, busui: nextBusui };
      }
      return { ...u, [field]: numVal };
    });
    const nextSasaran = recalculateSasaranFromUraian(nextUraian);
    commitActiveDayUpdate(nextSasaran, nextUraian);
  };

  const handleDeleteUraianRow = (rowId: string) => {
    const nextUraian = activeUraian.filter((u) => u.id !== rowId);
    const nextSasaran =
      nextUraian.length > 0 ? recalculateSasaranFromUraian(nextUraian) : activeSasaran;
    if (editingRowId === rowId) resetInputForm();
    commitActiveDayUpdate(nextSasaran, nextUraian);
  };

  const handleMoveUraianRow = (index: number, direction: -1 | 1) => {
    const targetIdx = index + direction;
    if (targetIdx < 0 || targetIdx >= activeUraian.length) return;
    const nextUraian = [...activeUraian];
    const [moved] = nextUraian.splice(index, 1);
    nextUraian.splice(targetIdx, 0, moved);
    commitActiveDayUpdate(activeSasaran, nextUraian);
  };

  const handleAddBlankRow = () => {
    const blank: UraianSasaranPM = {
      id: `blank_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      nama: "",
      tipeUnit: "sekolah",
      kategori: "sd_mi_slb",
      porsiBesar: 0,
      porsiKecil: 0,
      alergiBesar: 0,
      alergiKecil: 0,
      sampelBesar: 0,
      sampelKecil: 0,
    };
    commitActiveDayUpdate(activeSasaran, [...activeUraian, blank]);
  };

  const handleImportFromProfileLists = () => {
    const imported: UraianSasaranPM[] = [];
    sekolahPM.forEach((s, idx) => {
      const kat = detectSchoolCategoryFromName(s.namaSekolah);
      imported.push({
        id: `imp_sch_${idx}_${Date.now()}`,
        nama: s.namaSekolah,
        tipeUnit: "sekolah",
        kategori: kat,
        porsiBesar: Number(s.porsiBesar) || 0,
        porsiKecil: Number(s.porsiKecil) || 0,
        alergiBesar: kat === "tk_paud_lb" ? 0 : Number(s.alergi) || 0,
        alergiKecil: kat === "tk_paud_lb" ? Number(s.alergi) || 0 : 0,
        sampelBesar: 0,
        sampelKecil: 0,
      });
    });
    tigaBPM.forEach((t, idx) => {
      const bumil = Number(t.bumil) || 0;
      const busui = Number(t.busui) || 0;
      const balita = (Number(t.balita) || 0) + (Number(t.mpAsi) || 0);
      imported.push({
        id: `imp_3b_${idx}_${Date.now()}`,
        nama: t.namaDesa,
        tipeUnit: "tiga_b",
        kategori: "posyandu_3b",
        balitaTargetId: "anak_balita",
        bumil,
        busui,
        porsiBesar: bumil + busui,
        porsiKecil: balita,
        alergiBesar: 0,
        alergiKecil: Number(t.alergi) || 0,
        sampelBesar: 0,
        sampelKecil: 0,
      });
    });
    if (imported.length === 0) {
      triggerAlert("info", "Belum ada data sekolah / 3B tersimpan. Silakan input langsung melalui panel di atas.");
      return;
    }
    const nextSasaran = recalculateSasaranFromUraian(imported);
    commitActiveDayUpdate(nextSasaran, imported);
    triggerAlert("success", `Berhasil memuat ${imported.length} unit Sekolah & 3B ke Tabel 1 dan merekap ke Tabel 2.`);
  };

  const handleForceSyncTable1ToTable2 = () => {
    if (activeUraian.length === 0) {
      triggerAlert("info", "Tabel 1 (Uraian Kelompok Sasaran) masih kosong.");
      return;
    }
    const nextSasaran = recalculateSasaranFromUraian(activeUraian);
    commitActiveDayUpdate(nextSasaran, activeUraian);
    triggerAlert("success", "Tabel 2 (Rekapitulasi Keseluruhan) berhasil disinkronkan 100% dari Tabel 1.");
  };

  // Direct inline edit on Tabel 2 (Rekapitulasi Keseluruhan)
  const handleCellChangeTable2 = (
    id: string,
    field: "porsiKecil" | "porsiBesar" | "alergiKecil" | "alergiBesar",
    value: number
  ) => {
    const updatedSasaran = activeSasaran.map((s) =>
      s.id === id ? { ...s, [field]: Math.max(0, value) } : s
    );
    commitActiveDayUpdate(updatedSasaran, activeUraian);
  };

  // Copy current day (Tabel 1 + Tabel 2) to all planning days
  const executeCopyToAllDays = () => {
    const currentSasaranCopy = activeSasaran.map((s) => ({ ...s }));
    const currentUraianCopy = activeUraian.map((u) => ({ ...u }));
    const updatedHarian: HariPM[] = Array.from({
      length: Math.max(harianPM.length, effectiveTotalDays),
    }).map((_, i) => ({
      hariKe: i + 1,
      sasaran: currentSasaranCopy.map((s) => ({ ...s })),
      uraian: currentUraianCopy.map((u) => ({ ...u })),
    }));
    onChange(updatedHarian);
    setConfirmAction(null);
    triggerAlert(
      "success",
      `Sukses menyalin Tabel 1 & Tabel 2 Hari Ke-${selectedDay} ke seluruh ${effectiveTotalDays} Hari!`
    );
  };

  // Reset active day (Tabel 1 & Tabel 2) to 0
  const executeResetJumlahPM = () => {
    const clearedSasaran = activeSasaran.map((item) => ({
      ...item,
      porsiKecil: 0,
      porsiBesar: 0,
      alergiKecil: 0,
      alergiBesar: 0,
    }));
    const clearedUraian = activeUraian.map((u) => ({
      ...u,
      bumil: u.bumil !== undefined ? 0 : undefined,
      busui: u.busui !== undefined ? 0 : undefined,
      porsiBesar: 0,
      porsiKecil: 0,
      alergiBesar: 0,
      alergiKecil: 0,
      sampelBesar: 0,
      sampelKecil: 0,
    }));
    commitActiveDayUpdate(clearedSasaran, clearedUraian);
    setConfirmAction(null);
    triggerAlert("info", `Semua angka PM & Sampel Hari Ke-${selectedDay} telah direset menjadi 0.`);
  };

  // Export 2-Sheet Excel (Tabel 1 Uraian + Tabel 2 Rekapitulasi)
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    const uraianExport = activeUraian.map((u, idx) => ({
      No: idx + 1,
      "Kelompok Sasaran": u.nama,
      Kategori: KATEGORI_BADGE_MAP[u.kategori]?.label || u.kategori,
      "Porsi Besar": u.porsiBesar,
      "Porsi Kecil": u.porsiKecil,
      "Diet/Alergi Porsi Besar": u.alergiBesar,
      "Diet/Alergi Porsi Kecil": u.alergiKecil,
      "Sampel Porsi Besar": u.sampelBesar,
      "Sampel Porsi Kecil": u.sampelKecil,
      "Ibu Hamil (3B)": u.bumil ?? "",
      "Ibu Menyusui (3B)": u.busui ?? "",
      "Total Penerima Manfaat":
        (Number(u.porsiBesar) || 0) +
        (Number(u.porsiKecil) || 0) +
        (Number(u.alergiBesar) || 0) +
        (Number(u.alergiKecil) || 0) +
        (Number(u.sampelBesar) || 0) +
        (Number(u.sampelKecil) || 0),
    }));

    const rekapExport = activeSasaran.map((s, idx) => ({
      No: idx + 1,
      "Kelompok Sasaran": s.label,
      "Porsi Besar": s.porsiBesar,
      "Porsi Kecil": s.porsiKecil,
      "PM Alergi Porsi Besar": s.alergiBesar,
      "PM Alergi Porsi Kecil": s.alergiKecil,
      "Total Penerima Manfaat":
        (Number(s.porsiBesar) || 0) +
        (Number(s.porsiKecil) || 0) +
        (Number(s.alergiBesar) || 0) +
        (Number(s.alergiKecil) || 0),
    }));

    if (uraianExport.length > 0) {
      const ws1 = XLSX.utils.json_to_sheet(uraianExport);
      XLSX.utils.book_append_sheet(wb, ws1, `Tabel 1 Uraian H-${selectedDay}`);
    }
    const ws2 = XLSX.utils.json_to_sheet(rekapExport);
    XLSX.utils.book_append_sheet(wb, ws2, `Tabel 2 Rekap H-${selectedDay}`);

    XLSX.writeFile(wb, `Penerima_Manfaat_Hari_${selectedDay}.xlsx`);
    triggerAlert("success", "File Excel 2 Tabel Penerima Manfaat (termasuk Sampel) berhasil diunduh.");
  };

  // Import Excel (supports 2-sheet & legacy 1-sheet formats)
  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });

        // Check if Sheet 1 is Uraian format (has "Kategori" column or doesn't match the 13 fixed labels)
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const firstRows = XLSX.utils.sheet_to_json<any>(firstSheet);

        if (firstRows && firstRows.length > 0) {
          const looksLikeUraian = firstRows.some(
            (r: any) => r["Kategori"] !== undefined || r["Sampel Porsi Besar"] !== undefined
          );

          if (looksLikeUraian) {
            const parsedUraian: UraianSasaranPM[] = firstRows
              .filter((r: any) => String(r["Kelompok Sasaran"] || "").trim() !== "")
              .map((r: any, idx: number) => {
                const nama = String(r["Kelompok Sasaran"] || "").trim();
                const katRaw = String(r["Kategori"] || "").toLowerCase();
                let kategori: UraianKategoriPM = detectSchoolCategoryFromName(nama);
                let tipeUnit: InputModeTab = "sekolah";

                if (katRaw.includes("3b") || katRaw.includes("posyandu") || katRaw.includes("desa")) {
                  kategori = "posyandu_3b";
                  tipeUnit = "tiga_b";
                } else if (katRaw.includes("sampel")) {
                  kategori = "sampel";
                  tipeUnit = "sampel";
                } else if (katRaw.includes("tenaga")) {
                  kategori = "tenaga_kependidikan";
                  tipeUnit = "tenaga";
                } else if (katRaw.includes("pendidik")) {
                  kategori = "pendidik";
                  tipeUnit = "tenaga";
                } else if (katRaw.includes("paud") || katRaw.includes("tk")) {
                  kategori = "tk_paud_lb";
                } else if (katRaw.includes("smp") || katRaw.includes("mts")) {
                  kategori = "smp_mts_smplb";
                } else if (katRaw.includes("sma") || katRaw.includes("smk")) {
                  kategori = "sma_smk_ma";
                } else if (katRaw.includes("sd") || katRaw.includes("mi")) {
                  kategori = "sd_mi_slb";
                }

                const pb = Math.max(0, Number(r["Porsi Besar"]) || 0);
                const pk = Math.max(0, Number(r["Porsi Kecil"]) || 0);
                const bumil = r["Ibu Hamil (3B)"] !== "" && r["Ibu Hamil (3B)"] !== undefined ? Math.max(0, Number(r["Ibu Hamil (3B)"]) || 0) : undefined;
                const busui = r["Ibu Menyusui (3B)"] !== "" && r["Ibu Menyusui (3B)"] !== undefined ? Math.max(0, Number(r["Ibu Menyusui (3B)"]) || 0) : undefined;

                return {
                  id: `xls_${idx}_${Date.now()}`,
                  nama,
                  tipeUnit,
                  kategori,
                  balitaTargetId: tipeUnit === "tiga_b" ? "anak_balita" : undefined,
                  bumil,
                  busui,
                  porsiBesar: pb,
                  porsiKecil: pk,
                  alergiBesar: Math.max(0, Number(r["Diet/Alergi Porsi Besar"] ?? r["PM Alergi Porsi Besar"]) || 0),
                  alergiKecil: Math.max(0, Number(r["Diet/Alergi Porsi Kecil"] ?? r["PM Alergi Porsi Kecil"]) || 0),
                  sampelBesar: Math.max(0, Number(r["Sampel Porsi Besar"]) || 0),
                  sampelKecil: Math.max(0, Number(r["Sampel Porsi Kecil"]) || 0),
                };
              });

            const syncedSasaran = recalculateSasaranFromUraian(parsedUraian);
            commitActiveDayUpdate(syncedSasaran, parsedUraian);
            triggerAlert(
              "success",
              `Berhasil mengimpor ${parsedUraian.length} baris Uraian & menyinkronkan Tabel 2 Rekapitulasi!`
            );
          } else {
            // Legacy / Tabel 2 Rekap import
            const updatedSasaran = activeSasaran.map((s) => {
              const row = firstRows.find(
                (r: any) =>
                  String(r["Kelompok Sasaran"] || "").toLowerCase().trim() ===
                  s.label.toLowerCase().trim()
              );
              if (row) {
                const importedAlergiKecil =
                  row["PM Alergi Porsi Kecil"] ??
                  row["Diet/Alergi Porsi Kecil"] ??
                  row["PM Alergi"] ??
                  s.alergiKecil;
                const importedAlergiBesar =
                  row["PM Alergi Porsi Besar"] ??
                  row["Diet/Alergi Porsi Besar"] ??
                  s.alergiBesar;

                return {
                  ...s,
                  porsiKecil: Number(row["Porsi Kecil"] ?? s.porsiKecil) || 0,
                  porsiBesar: Number(row["Porsi Besar"] ?? s.porsiBesar) || 0,
                  alergiKecil: Number(importedAlergiKecil) || 0,
                  alergiBesar: Number(importedAlergiBesar) || 0,
                };
              }
              return s;
            });
            commitActiveDayUpdate(updatedSasaran, activeUraian);
            triggerAlert("success", `Berhasil memperbarui Tabel 2 Rekapitulasi Hari Ke-${selectedDay} dari file Excel!`);
          }
        }
      } catch {
        triggerAlert(
          "error",
          "Gagal memproses berkas Excel. Gunakan tombol 'Unduh Format' untuk mendapatkan template kolom yang sesuai."
        );
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  // Dynamic labels list for bottom classification settings
  const porsiKecilLabels = activeSasaran
    .filter((s) => settings.porsiKecilSasaranIds.includes(s.id))
    .map((s) => s.label)
    .join(", ");

  const porsiBesarLabels = activeSasaran
    .filter((s) => settings.porsiBesarSasaranIds.includes(s.id))
    .map((s) => s.label)
    .join(", ");

  const printTargetId = "print-area-penerima-manfaat";

  const printDocTitle =
    printScope === "table1"
      ? `Uraian Kelompok Sasaran PM — Hari Ke-${selectedDay}`
      : printScope === "table2"
      ? `Rekapitulasi Penerima Manfaat — Hari Ke-${selectedDay}`
      : `Data & Rekapitulasi Penerima Manfaat — Hari Ke-${selectedDay}`;

  const printFilename =
    printScope === "table1"
      ? `Uraian_Penerima_Manfaat_Hari_${selectedDay}`
      : printScope === "table2"
      ? `Rekap_Penerima_Manfaat_Hari_${selectedDay}`
      : `Data_dan_Rekap_Penerima_Manfaat_Hari_${selectedDay}`;

  // Helper to show blank on 0 in numeric inputs (matching Gambar 1 & Gambar 2 clean look)
  const numInputVal = (val: number) => (val > 0 ? String(val) : "");

  const previewBesar3B = (parseInt(formBumil, 10) || 0) + (parseInt(formBusui, 10) || 0);

  return (
    <div id="penerima-manfaat-container" className="space-y-6 font-sans">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
        <div>
          <h2 className="text-xl font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <Users className="w-5 h-5 text-indigo-600" />
            Penerima Manfaat ({effectiveTotalDays} Hari Perencanaan)
          </h2>
          <p className="text-sm text-slate-500">
            Input rincian Sekolah, 3B (Desa/Posyandu), Alergi, dan Sampel (Tabel 1) yang otomatis terekap ke Rekapitulasi Keseluruhan (Tabel 2).
          </p>
        </div>
      </div>

      {/* Alert Toast Banner */}
      {alertMsg && (
        <div
          className={`p-3.5 rounded-xl flex items-center justify-between gap-2 text-sm transition-all duration-200 shadow-xs border no-print ${
            alertMsg.type === "success"
              ? "bg-emerald-50 border-emerald-200 text-emerald-800"
              : alertMsg.type === "info"
              ? "bg-blue-50 border-blue-200 text-blue-800"
              : "bg-rose-50 border-rose-200 text-rose-800"
          }`}
        >
          <div className="flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-current shrink-0" />
            <span className="font-medium">{alertMsg.text}</span>
          </div>
          <button
            type="button"
            onClick={() => setAlertMsg(null)}
            className="p-1 rounded-lg hover:bg-black/5 cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* Inline Confirmation Banner (No native browser confirm dialog) */}
      {confirmAction && (
        <div className="p-4 rounded-2xl bg-amber-50 border border-amber-300 text-amber-950 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm no-print">
          <div className="flex items-start gap-2.5 text-sm">
            <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
            <div>
              <p className="font-bold">
                {confirmAction === "copy_all"
                  ? `Salin Tabel 1 & Tabel 2 Hari Ke-${selectedDay} ke seluruh ${effectiveTotalDays} hari perencanaan?`
                  : `Reset semua angka jumlah PM & Sampel Hari Ke-${selectedDay} menjadi 0?`}
              </p>
              <p className="text-xs text-amber-800 mt-0.5">
                {confirmAction === "copy_all"
                  ? "Seluruh daftar sekolah, 3B desa/posyandu, sampel, dan rekapitulasi hari lain akan disamakan dengan hari aktif ini."
                  : "Nama sekolah dan daftar posyandu tetap tersimpan, hanya angka porsi, alergi, dan sampel yang direset ke 0."}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={confirmAction === "copy_all" ? executeCopyToAllDays : executeResetJumlahPM}
              className={`px-3.5 py-2 rounded-xl text-xs font-bold text-white cursor-pointer transition ${
                confirmAction === "copy_all"
                  ? "bg-indigo-600 hover:bg-indigo-700"
                  : "bg-rose-600 hover:bg-rose-700"
              }`}
            >
              {confirmAction === "copy_all" ? "Ya, Samakan Semua Hari" : "Ya, Reset ke 0"}
            </button>
            <button
              type="button"
              onClick={() => setConfirmAction(null)}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-white border border-slate-200 text-slate-700 hover:bg-slate-50 cursor-pointer"
            >
              Batal
            </button>
          </div>
        </div>
      )}

      {/* Dynamic Day Selector Bar */}
      <div className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-xs space-y-3 no-print">
        <div className="flex items-center justify-between gap-2 flex-wrap">
          <div className="flex items-center gap-2 text-xs font-semibold text-slate-500 uppercase tracking-wider">
            <Calendar className="w-4 h-4 text-slate-400" />
            <span>Pilih Hari Kerja ({effectiveTotalDays} Hari Terjadwal)</span>
          </div>
          <span className="text-xs font-bold text-indigo-700 bg-indigo-50 border border-indigo-100 px-2.5 py-1 rounded-lg">
            Aktif: Hari {selectedDay} · {formattedActiveDateBanner}
          </span>
        </div>
        <div className="flex flex-wrap gap-2">
          {Array.from({ length: effectiveTotalDays }).map((_, i) => {
            const dayNum = i + 1;
            const isSelected = selectedDay === dayNum;
            const dateStr = profile?.periodeDates?.[i];
            return (
              <button
                key={dayNum}
                id={`btn-day-tab-${dayNum}`}
                type="button"
                onClick={() => setSelectedDay(dayNum)}
                className={`px-3.5 py-2 text-xs font-bold rounded-xl transition-all duration-150 border cursor-pointer text-left ${
                  isSelected
                    ? "bg-indigo-600 text-white border-indigo-600 shadow-xs"
                    : "bg-slate-50 text-slate-600 border-slate-200/80 hover:bg-slate-100 hover:text-slate-900"
                }`}
              >
                <span className="block leading-tight">HARI {dayNum}</span>
                {dateStr && (
                  <span
                    className={`text-[10px] font-normal block mt-0.5 ${
                      isSelected ? "text-indigo-100" : "text-slate-400"
                    }`}
                  >
                    {formatIndonesianDate(dateStr)}
                  </span>
                )}
              </button>
            );
          })}
        </div>
      </div>

      {/* Ringkasan Harian (Per Kategori + Total PM Harian & Estimasi RAB Harian dengan Tanggal) */}
      <div className="space-y-3 no-print">
        {/* Row 1: 4 Category Breakdown Cards (Non-Overlapping) */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          <div className="bg-emerald-50/75 p-4 rounded-2xl border border-emerald-200/80 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between gap-1">
                <span className="text-[11px] font-bold text-emerald-800 uppercase tracking-wider">
                  Porsi Besar (H-{selectedDay})
                </span>
                <span className="text-[10px] font-semibold text-emerald-700/80">Tanpa Alergi</span>
              </div>
              <div className="mt-1.5 flex items-baseline gap-1.5">
                <span className="text-2xl font-black text-emerald-950 tabular-nums">
                  {dailyBreakdown.nonSampelBesar.toLocaleString("id-ID")}
                </span>
                <span className="text-[11px] text-emerald-700 font-bold">Porsi</span>
              </div>
            </div>
            <div className="mt-2 pt-2 border-t border-emerald-200/70 flex items-center justify-between text-[11px]">
              <span className="text-emerald-700/90 font-medium">Estimasi RAB</span>
              <span className="font-extrabold font-mono text-emerald-950">
                {formatIDR(dailyBreakdown.rabNonSampelBesar)}
              </span>
            </div>
          </div>

          <div className="bg-sky-50/75 p-4 rounded-2xl border border-sky-200/80 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between gap-1">
                <span className="text-[11px] font-bold text-sky-800 uppercase tracking-wider">
                  Porsi Kecil (H-{selectedDay})
                </span>
                <span className="text-[10px] font-semibold text-sky-700/80">Tanpa Alergi</span>
              </div>
              <div className="mt-1.5 flex items-baseline gap-1.5">
                <span className="text-2xl font-black text-sky-950 tabular-nums">
                  {dailyBreakdown.nonSampelKecil.toLocaleString("id-ID")}
                </span>
                <span className="text-[11px] text-sky-700 font-bold">Porsi</span>
              </div>
            </div>
            <div className="mt-2 pt-2 border-t border-sky-200/70 flex items-center justify-between text-[11px]">
              <span className="text-sky-700/90 font-medium">Estimasi RAB</span>
              <span className="font-extrabold font-mono text-sky-950">
                {formatIDR(dailyBreakdown.rabNonSampelKecil)}
              </span>
            </div>
          </div>

          <div className="bg-orange-50/80 p-4 rounded-2xl border border-orange-200/80 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between gap-1">
                <span className="text-[11px] font-bold text-orange-800 uppercase tracking-wider">
                  Diet / Alergi (H-{selectedDay})
                </span>
                <span className="text-[10px] font-semibold text-orange-700">
                  B: {dailyBreakdown.nonSampelAlergiBesar} · K: {dailyBreakdown.nonSampelAlergiKecil}
                </span>
              </div>
              <div className="mt-1.5 flex items-baseline gap-1.5">
                <span className="text-2xl font-black text-orange-950 tabular-nums">
                  {(dailyBreakdown.nonSampelAlergiBesar + dailyBreakdown.nonSampelAlergiKecil).toLocaleString("id-ID")}
                </span>
                <span className="text-[11px] text-orange-700 font-bold">Porsi</span>
              </div>
            </div>
            <div className="mt-2 pt-2 border-t border-orange-200/70 flex items-center justify-between text-[11px]">
              <span className="text-orange-700/90 font-medium">
                B: {formatIDR(dailyBreakdown.rabNonSampelAlergiBesar)} · K: {formatIDR(dailyBreakdown.rabNonSampelAlergiKecil)}
              </span>
            </div>
          </div>

          <div className="bg-blue-50/80 p-4 rounded-2xl border border-blue-200/80 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between gap-1">
                <span className="text-[11px] font-bold text-blue-800 uppercase tracking-wider flex items-center gap-1">
                  <FlaskConical className="w-3.5 h-3.5" />
                  Sampel (H-{selectedDay})
                </span>
                <span className="text-[10px] font-semibold text-blue-700">
                  B: {sampelPorsiBesar} · K: {sampelPorsiKecil}
                </span>
              </div>
              <div className="mt-1.5 flex items-baseline gap-1.5">
                <span className="text-2xl font-black text-blue-950 tabular-nums">
                  {totalSampel.toLocaleString("id-ID")}
                </span>
                <span className="text-[11px] text-blue-700 font-bold">Porsi</span>
              </div>
            </div>
            <div className="mt-2 pt-2 border-t border-blue-200/70 flex items-center justify-between text-[11px]">
              <span className="text-blue-700/90 font-medium">Estimasi RAB</span>
              <span className="font-extrabold font-mono text-blue-950">
                {formatIDR(dailyBreakdown.rabSampel)}
              </span>
            </div>
          </div>
        </div>

        {/* Row 2: 2 Wide Daily Hero Cards WITH DATE (Total PM Harian & Estimasi RAB Harian) */}
        <div className="grid grid-cols-1 xl:grid-cols-2 gap-3">
          {/* TOTAL PM HARIAN */}
          <div className="bg-slate-900 p-4 rounded-2xl border border-slate-800 text-white space-y-3 shadow-xs">
            <div className="flex items-start justify-between gap-3 flex-wrap">
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-[11px] font-extrabold text-slate-300 uppercase tracking-wider">
                    Total PM Harian (H-{selectedDay})
                  </span>
                  <span className="inline-flex items-center gap-1 text-[11px] font-bold text-indigo-200 bg-slate-800 border border-slate-700 px-2 py-0.5 rounded-lg">
                    <CalendarDays className="w-3 h-3 text-indigo-400" />
                    {formattedActiveDateBanner}
                  </span>
                </div>
                <div className="mt-1.5 flex items-baseline gap-2">
                  <span className="text-3xl font-black text-white tabular-nums tracking-tight">
                    {totalPM.toLocaleString("id-ID")}
                  </span>
                  <span className="text-xs text-slate-300 font-bold">Porsi</span>
                  <span className="text-[11px] text-slate-400 ml-1">
                    ({activeUraian.length} unit di Tabel 1)
                  </span>
                </div>
              </div>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 pt-2.5 border-t border-slate-800 text-[11px]">
              <div className="bg-slate-800/70 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] text-emerald-300 font-semibold block">Porsi Besar</span>
                <span className="font-extrabold text-white tabular-nums">
                  {dailyBreakdown.nonSampelBesar.toLocaleString("id-ID")}
                </span>
              </div>
              <div className="bg-slate-800/70 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] text-orange-300 font-semibold block">Alergi Besar</span>
                <span className="font-extrabold text-white tabular-nums">
                  {dailyBreakdown.nonSampelAlergiBesar.toLocaleString("id-ID")}
                </span>
              </div>
              <div className="bg-slate-800/70 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] text-sky-300 font-semibold block">Porsi Kecil</span>
                <span className="font-extrabold text-white tabular-nums">
                  {dailyBreakdown.nonSampelKecil.toLocaleString("id-ID")}
                </span>
              </div>
              <div className="bg-slate-800/70 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] text-amber-300 font-semibold block">Alergi Kecil</span>
                <span className="font-extrabold text-white tabular-nums">
                  {dailyBreakdown.nonSampelAlergiKecil.toLocaleString("id-ID")}
                </span>
              </div>
              <div className="bg-slate-800/70 rounded-xl px-2.5 py-1.5 col-span-2 sm:col-span-1">
                <span className="text-[10px] text-blue-300 font-semibold block">Sampel</span>
                <span className="font-extrabold text-white tabular-nums">
                  {totalSampel.toLocaleString("id-ID")}{" "}
                  <span className="text-[9px] font-normal text-slate-300">
                    (B:{sampelPorsiBesar}/K:{sampelPorsiKecil})
                  </span>
                </span>
              </div>
            </div>
          </div>

          {/* ESTIMASI RAB HARIAN */}
          <div className="bg-gradient-to-br from-emerald-700 to-teal-800 p-4 rounded-2xl border border-emerald-600 text-white space-y-3 shadow-xs">
            <div className="flex items-start justify-between gap-3 flex-wrap">
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-[11px] font-extrabold text-emerald-100 uppercase tracking-wider">
                    Estimasi RAB Harian (H-{selectedDay})
                  </span>
                  <span className="inline-flex items-center gap-1 text-[11px] font-bold text-white bg-emerald-900/50 border border-emerald-500/50 px-2 py-0.5 rounded-lg">
                    <CalendarDays className="w-3 h-3 text-emerald-200" />
                    {formattedActiveDateBanner}
                  </span>
                </div>
                <div className="mt-1.5 flex items-baseline gap-2 flex-wrap">
                  <span className="text-3xl font-black font-mono text-white tracking-tight">
                    {formatIDR(totalRABHarian)}
                  </span>
                  <span className="text-[11px] text-emerald-100/90 font-medium">
                    (Besar {formatIDR(settings.porsiBesarHarga)} · Kecil {formatIDR(settings.porsiKecilHarga)})
                  </span>
                </div>
              </div>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 pt-2.5 border-t border-emerald-600/60 text-[11px]">
              <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] text-emerald-200 font-semibold block">Porsi Besar</span>
                <span className="font-extrabold font-mono text-white block truncate">
                  {formatIDR(dailyBreakdown.rabNonSampelBesar)}
                </span>
              </div>
              <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] text-amber-200 font-semibold block">Alergi Besar</span>
                <span className="font-extrabold font-mono text-white block truncate">
                  {formatIDR(dailyBreakdown.rabNonSampelAlergiBesar)}
                </span>
              </div>
              <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] text-sky-200 font-semibold block">Porsi Kecil</span>
                <span className="font-extrabold font-mono text-white block truncate">
                  {formatIDR(dailyBreakdown.rabNonSampelKecil)}
                </span>
              </div>
              <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] text-orange-200 font-semibold block">Alergi Kecil</span>
                <span className="font-extrabold font-mono text-white block truncate">
                  {formatIDR(dailyBreakdown.rabNonSampelAlergiKecil)}
                </span>
              </div>
              <div className="bg-emerald-950/35 rounded-xl px-2.5 py-1.5 col-span-2 sm:col-span-1">
                <span className="text-[10px] text-blue-200 font-semibold block">Sampel</span>
                <span className="font-extrabold font-mono text-white block truncate">
                  {formatIDR(dailyBreakdown.rabSampel)}
                </span>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* PANEL INPUT KHUSUS: SEKOLAH, 3B (DESA/POSYANDU), TENAGA PENDIDIKAN, & SAMPEL */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden no-print">
        <div className="px-5 py-4 bg-slate-50 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          <div>
            <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <Plus className="w-4 h-4 text-emerald-600" />
              {editingRowId
                ? "Edit Data Kelompok Sasaran (Tabel 1)"
                : "Panel Input Khusus Kelompok Sasaran (Sekolah, 3B Desa/Posyandu & Sampel)"}
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Setiap unit yang ditambahkan masuk ke <strong>Tabel 1 (Uraian Kelompok Sasaran)</strong> dan otomatis dihitung ke <strong>Tabel 2 (Rekapitulasi Keseluruhan)</strong>.
            </p>
          </div>

          {/* Input Mode Segmented Buttons */}
          <div className="flex flex-wrap items-center gap-1.5 bg-white p-1 rounded-xl border border-slate-200">
            <button
              type="button"
              onClick={() => handleSelectInputMode("sekolah")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition cursor-pointer ${
                inputMode === "sekolah"
                  ? "bg-emerald-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              <School className="w-3.5 h-3.5" />
              Sekolah
            </button>
            <button
              type="button"
              onClick={() => handleSelectInputMode("tiga_b")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition cursor-pointer ${
                inputMode === "tiga_b"
                  ? "bg-emerald-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              <Home className="w-3.5 h-3.5" />
              3B (Desa / Posyandu)
            </button>
            <button
              type="button"
              onClick={() => handleSelectInputMode("tenaga")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition cursor-pointer ${
                inputMode === "tenaga"
                  ? "bg-emerald-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              <GraduationCap className="w-3.5 h-3.5" />
              Pendidik / Tendik
            </button>
            <button
              type="button"
              onClick={() => handleSelectInputMode("sampel")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition cursor-pointer ${
                inputMode === "sampel"
                  ? "bg-blue-600 text-white shadow-xs"
                  : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              <FlaskConical className="w-3.5 h-3.5" />
              Sampel Khusus
            </button>
          </div>
        </div>

        <form onSubmit={handleSubmitUraianForm} className="p-5 space-y-4">
          {/* Row 1: Unit Name + Category Mapping */}
          <div className="grid grid-cols-1 md:grid-cols-12 gap-4">
            <div className="md:col-span-7">
              <label className="block text-xs font-bold text-slate-700 mb-1.5">
                {inputMode === "sekolah"
                  ? "Nama Sekolah *"
                  : inputMode === "tiga_b"
                  ? "Nama Desa / Posyandu (3B) *"
                  : inputMode === "tenaga"
                  ? "Nama Unit Pendidik / Tenaga Pendidikan *"
                  : "Keterangan Baris Sampel *"}
              </label>
              <input
                type="text"
                value={formNama}
                onChange={(e) => {
                  const val = e.target.value;
                  setFormNama(val);
                  if (inputMode === "sekolah" && !editingRowId) {
                    setFormKategori(detectSchoolCategoryFromName(val));
                  }
                }}
                placeholder={
                  inputMode === "sekolah"
                    ? "Contoh: Paud Bersama, SDN 3 Sawerigadi, SMPN 1 Sawerigadi, SMA 1 Sawerigadi"
                    : inputMode === "tiga_b"
                    ? "Contoh: Posyandu Kasih Ibu Desa Marobea, Mawar Ondoke, Kuncup Mekar Lakalamba"
                    : inputMode === "tenaga"
                    ? "Contoh: Tenaga Pendidikan / Pendidik"
                    : "Contoh: Sampel Makanan Harian"
                }
                className="w-full px-3.5 py-2.5 text-sm bg-white border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 font-medium"
              />
            </div>

            <div className="md:col-span-5">
              {inputMode === "sekolah" && (
                <>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">
                    Jenjang / Kelompok Rekap Tabel 2
                  </label>
                  <select
                    value={formKategori}
                    onChange={(e) => setFormKategori(e.target.value as UraianKategoriPM)}
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 font-semibold text-slate-800"
                  >
                    <option value="tk_paud_lb">PAUD / TK / KB / LB → Siswa TK/PAUD/LB</option>
                    <option value="sd_mi_slb">SD / MI / SLB → Kecil (Kls 1-3) & Besar (Kls 4-6)</option>
                    <option value="smp_mts_smplb">SMP / MTS / SMPLB → Siswa SMP/MTS/SMPLB</option>
                    <option value="sma_smk_ma">SMA / SMK / MK / MA → Siswa SMA/SMK/MA</option>
                  </select>
                </>
              )}

              {inputMode === "tiga_b" && (
                <>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">
                    Kategori Balita (untuk Porsi Kecil 3B di Tabel 2)
                  </label>
                  <select
                    value={formBalitaTarget}
                    onChange={(e) =>
                      setFormBalitaTarget(
                        e.target.value as "anak_balita" | "anak_balita_13_59" | "balita_6_11"
                      )
                    }
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 font-semibold text-slate-800"
                  >
                    <option value="anak_balita">Anak Balita (Baris 3)</option>
                    <option value="anak_balita_13_59">Anak Balita Usia 13-59 Bulan (Baris 2)</option>
                    <option value="balita_6_11">Balita 6-11 Bulan / MP-ASI (Baris 1)</option>
                  </select>
                </>
              )}

              {inputMode === "tenaga" && (
                <>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">
                    Target Baris Rekap Tabel 2
                  </label>
                  <select
                    value={formKategori}
                    onChange={(e) => {
                      const next = e.target.value as UraianKategoriPM;
                      setFormKategori(next);
                      if (!editingRowId) {
                        setFormNama(next === "pendidik" ? "Pendidik" : "Tenaga Pendidikan");
                      }
                    }}
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 font-semibold text-slate-800"
                  >
                    <option value="tenaga_kependidikan">Tenaga Pendidikan (Baris 10)</option>
                    <option value="pendidik">Pendidik (Baris 9)</option>
                  </select>
                </>
              )}

              {inputMode === "sampel" && (
                <div className="h-full flex items-end">
                  <div className="w-full px-3.5 py-2.5 rounded-xl bg-blue-50 border border-blue-200 text-xs text-blue-800 font-semibold">
                    Otomatis masuk ke Baris 13 <strong>Sampel</strong> pada Tabel 2 (Rekapitulasi Keseluruhan).
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Row 2: Portion Counts + Allergy Counts */}
          {inputMode !== "sampel" && (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
              {inputMode === "tiga_b" ? (
                <>
                  <div className="bg-emerald-50/60 p-3 rounded-xl border border-emerald-200 space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="text-[11px] font-bold text-emerald-900 uppercase">
                        Porsi Besar 3B (Bumil & Busui)
                      </span>
                      <span className="text-[11px] font-black text-emerald-700 tabular-nums">
                        Total: {previewBesar3B}
                      </span>
                    </div>
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="block text-[10px] font-semibold text-emerald-800 mb-0.5">
                          Ibu Hamil
                        </label>
                        <input
                          type="number"
                          min="0"
                          value={formBumil}
                          onChange={(e) => setFormBumil(e.target.value)}
                          placeholder="0"
                          className="w-full px-2.5 py-1.5 text-sm font-bold font-mono text-center bg-white border border-emerald-300 rounded-lg focus:outline-none focus:border-emerald-600"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-semibold text-emerald-800 mb-0.5">
                          Ibu Menyusui
                        </label>
                        <input
                          type="number"
                          min="0"
                          value={formBusui}
                          onChange={(e) => setFormBusui(e.target.value)}
                          placeholder="0"
                          className="w-full px-2.5 py-1.5 text-sm font-bold font-mono text-center bg-white border border-emerald-300 rounded-lg focus:outline-none focus:border-emerald-600"
                        />
                      </div>
                    </div>
                  </div>

                  <div className="bg-emerald-50/60 p-3 rounded-xl border border-emerald-200 flex flex-col justify-between">
                    <label className="block text-[11px] font-bold text-emerald-900 uppercase mb-1">
                      Jumlah Porsi Kecil (Balita)
                    </label>
                    <input
                      type="number"
                      min="0"
                      value={formPorsiKecil}
                      onChange={(e) => setFormPorsiKecil(e.target.value)}
                      placeholder="0"
                      className="w-full px-3 py-2 text-sm font-bold font-mono text-center bg-white border border-emerald-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                </>
              ) : (
                <>
                  <div className="bg-emerald-50/60 p-3 rounded-xl border border-emerald-200">
                    <label className="block text-[11px] font-bold text-emerald-900 uppercase mb-1">
                      Jumlah Porsi Besar (Tanpa Alergi)
                    </label>
                    <input
                      type="number"
                      min="0"
                      value={formPorsiBesar}
                      onChange={(e) => setFormPorsiBesar(e.target.value)}
                      placeholder="0"
                      className="w-full px-3 py-2 text-sm font-bold font-mono text-center bg-white border border-emerald-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                    {inputMode === "sekolah" && formKategori === "sd_mi_slb" && (
                      <span className="block text-[10px] text-emerald-700 mt-1">
                        Otomatis masuk SD Kelas 4-6
                      </span>
                    )}
                  </div>

                  <div className="bg-emerald-50/60 p-3 rounded-xl border border-emerald-200">
                    <label className="block text-[11px] font-bold text-emerald-900 uppercase mb-1">
                      Jumlah Porsi Kecil (Tanpa Alergi)
                    </label>
                    <input
                      type="number"
                      min="0"
                      value={formPorsiKecil}
                      onChange={(e) => setFormPorsiKecil(e.target.value)}
                      placeholder="0"
                      className="w-full px-3 py-2 text-sm font-bold font-mono text-center bg-white border border-emerald-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                    {inputMode === "sekolah" && formKategori === "sd_mi_slb" && (
                      <span className="block text-[10px] text-emerald-700 mt-1">
                        Otomatis masuk SD Kelas 1-3
                      </span>
                    )}
                  </div>
                </>
              )}

              <div className="bg-orange-50/70 p-3 rounded-xl border border-orange-200">
                <label className="block text-[11px] font-bold text-orange-900 uppercase mb-1">
                  Alergi Porsi Besar
                </label>
                <input
                  type="number"
                  min="0"
                  value={formAlergiBesar}
                  onChange={(e) => setFormAlergiBesar(e.target.value)}
                  placeholder="0"
                  className="w-full px-3 py-2 text-sm font-bold font-mono text-center bg-white border border-orange-300 rounded-lg focus:outline-none focus:border-orange-500 text-orange-950"
                />
              </div>

              <div className="bg-orange-50/70 p-3 rounded-xl border border-orange-200">
                <label className="block text-[11px] font-bold text-orange-900 uppercase mb-1">
                  Alergi Porsi Kecil
                </label>
                <input
                  type="number"
                  min="0"
                  value={formAlergiKecil}
                  onChange={(e) => setFormAlergiKecil(e.target.value)}
                  placeholder="0"
                  className="w-full px-3 py-2 text-sm font-bold font-mono text-center bg-white border border-orange-300 rounded-lg focus:outline-none focus:border-orange-500 text-orange-950"
                />
              </div>
            </div>
          )}

          {/* Row 3: Sampel Selector (Porsi Kecil / Porsi Besar / Keduanya) */}
          <div className="p-3.5 rounded-xl bg-blue-50/70 border border-blue-200 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
            <div className="space-y-1.5">
              <div className="flex items-center gap-1.5 text-xs font-bold text-blue-900">
                <FlaskConical className="w-4 h-4 text-blue-600" />
                <span>Pengaturan Sampel Makanan (Porsi Besar / Porsi Kecil)</span>
              </div>
              <div className="flex flex-wrap items-center gap-1.5">
                {(
                  [
                    { id: "none", label: "Tanpa Sampel" },
                    { id: "besar", label: "Sampel Porsi Besar" },
                    { id: "kecil", label: "Sampel Porsi Kecil" },
                    { id: "both", label: "Porsi Besar & Kecil" },
                  ] as { id: SampelChoiceMode; label: string }[]
                ).map((opt) => (
                  <button
                    key={opt.id}
                    type="button"
                    disabled={inputMode === "sampel" && opt.id === "none"}
                    onClick={() => handleSampelChoiceChange(opt.id)}
                    className={`px-2.5 py-1.5 rounded-lg text-xs font-bold border transition cursor-pointer disabled:opacity-40 ${
                      formSampelChoice === opt.id
                        ? "bg-blue-600 text-white border-blue-600 shadow-xs"
                        : "bg-white text-blue-900 border-blue-200 hover:bg-blue-100/70"
                    }`}
                  >
                    {opt.label}
                  </button>
                ))}
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-3">
              {(formSampelChoice === "besar" || formSampelChoice === "both") && (
                <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-lg border border-blue-200">
                  <label className="text-[11px] font-bold text-blue-900 whitespace-nowrap">
                    Sampel Porsi Besar:
                  </label>
                  <input
                    type="number"
                    min="0"
                    value={formSampelBesar}
                    onChange={(e) => setFormSampelBesar(e.target.value)}
                    placeholder="0"
                    className="w-20 px-2 py-1 text-sm font-bold font-mono text-center bg-blue-50/40 border border-blue-300 rounded-md focus:outline-none focus:border-blue-600"
                  />
                </div>
              )}

              {(formSampelChoice === "kecil" || formSampelChoice === "both") && (
                <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-lg border border-blue-200">
                  <label className="text-[11px] font-bold text-blue-900 whitespace-nowrap">
                    Sampel Porsi Kecil:
                  </label>
                  <input
                    type="number"
                    min="0"
                    value={formSampelKecil}
                    onChange={(e) => setFormSampelKecil(e.target.value)}
                    placeholder="0"
                    className="w-20 px-2 py-1 text-sm font-bold font-mono text-center bg-blue-50/40 border border-blue-300 rounded-md focus:outline-none focus:border-blue-600"
                  />
                </div>
              )}
            </div>
          </div>

          {/* Form Action Footer */}
          <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="submit"
                className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs transition cursor-pointer"
              >
                {editingRowId ? <Check className="w-4 h-4" /> : <Plus className="w-4 h-4" />}
                {editingRowId
                  ? "Simpan Perubahan Baris"
                  : "Tambah ke Tabel 1 & Rekap Otomatis"}
              </button>

              {editingRowId && (
                <button
                  type="button"
                  onClick={resetInputForm}
                  className="px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition cursor-pointer"
                >
                  Batal Edit
                </button>
              )}

              <button
                type="button"
                onClick={handleAddBlankRow}
                className="px-3.5 py-2.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer"
                title="Tambahkan baris kosong bernomor di bagian bawah Tabel 1 (seperti baris cadangan)"
              >
                <Plus className="w-3.5 h-3.5 text-slate-500" />
                Baris Kosong (Blanko)
              </button>

              {activeUraian.length === 0 && (sekolahPM.length > 0 || tigaBPM.length > 0) && (
                <button
                  type="button"
                  onClick={handleImportFromProfileLists}
                  className="px-3.5 py-2.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 rounded-xl text-xs font-bold flex items-center gap-1.5 transition cursor-pointer"
                >
                  <Database className="w-3.5 h-3.5" />
                  Muat {sekolahPM.length} Sekolah & {tigaBPM.length} Desa 3B
                </button>
              )}
            </div>

            {activeUraian.length > 0 && (
              <button
                type="button"
                onClick={handleForceSyncTable1ToTable2}
                className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer"
                title="Hitung ulang seluruh baris Tabel 2 berdasarkan data terbaru Tabel 1"
              >
                <RefreshCw className="w-3.5 h-3.5 text-emerald-600" />
                Sinkronkan Tabel 1 → Tabel 2
              </button>
            )}
          </div>
        </form>
      </div>

      {/* GLOBAL PRINT / EXPORT / VIEW TOOLBAR */}
      <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs space-y-3 no-print">
        <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-xs font-bold text-slate-600 flex items-center gap-1.5">
              <Table2 className="w-4 h-4 text-indigo-600" />
              Cetak / Unduh:
            </span>
            <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl border border-slate-200 text-xs font-bold">
              {(
                [
                  { id: "both", label: "Tabel 1 + Tabel 2" },
                  { id: "table1", label: "Hanya Tabel 1 (Uraian)" },
                  { id: "table2", label: "Hanya Tabel 2 (Rekap)" },
                ] as { id: PrintScopeMode; label: string }[]
              ).map((opt) => (
                <button
                  key={opt.id}
                  type="button"
                  onClick={() => setPrintScope(opt.id)}
                  className={`px-2.5 py-1.5 rounded-lg transition cursor-pointer ${
                    printScope === opt.id
                      ? "bg-indigo-600 text-white shadow-xs"
                      : "text-slate-600 hover:bg-white"
                  }`}
                >
                  {opt.label}
                </button>
              ))}
            </div>

            <label className="flex items-center gap-1.5 text-xs font-semibold text-slate-700 bg-blue-50/70 border border-blue-200 px-2.5 py-1.5 rounded-xl cursor-pointer select-none">
              <input
                type="checkbox"
                checked={showSampelColInTable1}
                onChange={(e) => setShowSampelColInTable1(e.target.checked)}
                className="rounded text-blue-600 focus:ring-blue-500 w-3.5 h-3.5"
              />
              <span>Kolom Sampel di Tabel 1</span>
            </label>

            <label className="flex items-center gap-1.5 text-xs font-semibold text-slate-700 bg-slate-50 border border-slate-200 px-2.5 py-1.5 rounded-xl cursor-pointer select-none">
              <input
                type="checkbox"
                checked={includeRabInPrint}
                onChange={(e) => setIncludeRabInPrint(e.target.checked)}
                className="rounded text-emerald-600 focus:ring-emerald-500 w-3.5 h-3.5"
              />
              <span>Sertakan RAB saat Cetak</span>
            </label>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              id="btn-preview-pm-modal"
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="flex items-center gap-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold transition cursor-pointer shadow-xs"
            >
              <Eye className="w-3.5 h-3.5" />
              Pratinjau Cetak
            </button>

            <button
              id="btn-export-pm-pdf"
              type="button"
              disabled={!!isDownloading}
              onClick={() =>
                exportToPDF(printTargetId, printFilename, 100, setIsDownloading, {
                  paperSize,
                  orientation: "auto",
                })
              }
              className="flex items-center gap-1.5 px-3 py-2 bg-rose-600 hover:bg-rose-700 disabled:opacity-50 text-white rounded-xl text-xs font-bold transition cursor-pointer shadow-xs"
            >
              <Download className="w-3.5 h-3.5" />
              {isDownloading ? isDownloading : "Ekspor PDF"}
            </button>

            <button
              id="btn-export-pm-img"
              type="button"
              disabled={!!isDownloading}
              onClick={() =>
                downloadElementAsImage(printTargetId, printFilename, setIsDownloading, {
                  paperSize,
                  orientation: "auto",
                })
              }
              className="flex items-center gap-1.5 px-3 py-2 bg-amber-600 hover:bg-amber-700 disabled:opacity-50 text-white rounded-xl text-xs font-bold transition cursor-pointer shadow-xs"
            >
              <ImageIcon className="w-3.5 h-3.5" />
              PNG
            </button>

            <button
              id="btn-print-pm-direct"
              type="button"
              onClick={() =>
                printElementSmart(printTargetId, {
                  title: printDocTitle,
                  paperSize,
                  orientation: "auto",
                })
              }
              className="flex items-center gap-1.5 px-3 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition cursor-pointer shadow-xs"
            >
              <Printer className="w-3.5 h-3.5" />
              Cetak
            </button>

            <button
              id="btn-copy-all-days"
              type="button"
              onClick={() => setConfirmAction("copy_all")}
              className="flex items-center gap-1.5 px-3 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl text-xs font-semibold transition cursor-pointer"
              title={`Salin Tabel 1 & Tabel 2 Hari Ke-${selectedDay} ke semua ${effectiveTotalDays} hari perencanaan`}
            >
              <Copy className="w-3.5 h-3.5" />
              Sama-kan {effectiveTotalDays} Hari
            </button>

            <label className="flex items-center gap-1.5 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold cursor-pointer transition">
              <Upload className="w-3.5 h-3.5" />
              Unggah Excel
              <input type="file" accept=".xlsx, .xls" onChange={handleImportExcel} className="hidden" />
            </label>

            <button
              id="btn-export-pm"
              type="button"
              onClick={handleExportExcel}
              className="flex items-center gap-1.5 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold transition cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Unduh Excel
            </button>

            <button
              id="btn-reset-pm"
              type="button"
              onClick={() => setConfirmAction("reset_day")}
              className="flex items-center gap-1.5 px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-xl text-xs font-bold transition cursor-pointer"
              title="Reset semua jumlah PM & Sampel pada hari ini menjadi 0"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Reset
            </button>
          </div>
        </div>
      </div>

      {/* PRINTABLE WORKSHEET AREA CONTAINING BOTH TABEL 1 & TABEL 2 */}
      <div
        id="print-area-penerima-manfaat"
        className="printable-area bg-white rounded-2xl border border-slate-200 shadow-sm p-5 sm:p-6 space-y-8"
        style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
      >
        {/* Printable KOP Surat Header */}
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
          docKey="penerima"
        />

        {/* ===================================================================== */}
        {/* TABEL 1 — URAIAN KELOMPOK SASARAN (PERSIS GAMBAR 1 + KOLOM SAMPEL)   */}
        {/* ===================================================================== */}
        <div
          id="print-area-pm-tabel-1"
          className={`space-y-3 ${printScope === "table2" ? "no-print" : ""}`}
        >

          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200 pb-2">
            <div>
              <h3 className="text-sm sm:text-base font-bold text-black uppercase tracking-tight">
                TABEL 1. URAIAN KELOMPOK SASARAN PENERIMA MANFAAT (HARI KE-{selectedDay})
              </h3>
              <p className="text-xs text-slate-600">
                {formattedActiveDateBanner} · Rincian per Sekolah, 3B (Desa / Posyandu), Pendidik/Tenaga Pendidikan &amp; Sampel
              </p>
            </div>
            <span className="text-xs font-bold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded-lg no-print self-start">
              {activeUraian.length} Unit Sasaran
            </span>
          </div>

          <div className="overflow-x-auto">
            <table
              className="w-full border-collapse text-xs text-black"
              style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
            >
              <thead>
                <tr>
                  <th
                    colSpan={2}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    DATA PENERIMA KELOMPOK SASARAN
                  </th>
                  <th
                    colSpan={showSampelColInTable1 ? 6 : 4}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    JUMLAH PENERIMA MANFAAT
                  </th>
                  <th
                    rowSpan={3}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle w-24"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    TOTAL PENERIMA MANFAAT
                  </th>
                  <th
                    rowSpan={3}
                    className="border border-black px-2 py-2 text-center font-bold text-slate-800 align-middle w-24 no-print"
                    style={{ backgroundColor: "#E2E8F0" }}
                  >
                    Aksi
                  </th>
                </tr>
                <tr>
                  <th
                    rowSpan={2}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle w-10"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    No.
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-2.5 py-2 text-center font-bold text-black align-middle min-w-[210px]"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    Kelompok Sasaran
                  </th>
                  <th
                    colSpan={2}
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    Tanpa Alergi
                  </th>
                  <th
                    colSpan={2}
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#FF6B00" }}
                  >
                    Diet/Alergi
                  </th>
                  {showSampelColInTable1 && (
                    <th
                      colSpan={2}
                      className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle"
                      style={{ backgroundColor: "#93C5FD" }}
                    >
                      Sampel
                    </th>
                  )}
                </tr>
                <tr>
                  <th
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-20"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    Porsi Besar
                  </th>
                  <th
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-20"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    Porsi Kecil
                  </th>
                  <th
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-20"
                    style={{ backgroundColor: "#FF6B00" }}
                  >
                    Porsi Besar
                  </th>
                  <th
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-20"
                    style={{ backgroundColor: "#FF6B00" }}
                  >
                    Porsi Kecil
                  </th>
                  {showSampelColInTable1 && (
                    <>
                      <th
                        className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-20"
                        style={{ backgroundColor: "#93C5FD" }}
                      >
                        Porsi Besar
                      </th>
                      <th
                        className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-20"
                        style={{ backgroundColor: "#93C5FD" }}
                      >
                        Porsi Kecil
                      </th>
                    </>
                  )}
                </tr>
              </thead>
              <tbody>
                {activeUraian.length === 0 ? (
                  <tr>
                    <td
                      colSpan={showSampelColInTable1 ? 10 : 8}
                      className="border border-black px-4 py-6 text-center text-slate-500"
                      style={{ backgroundColor: "#FFFFFF" }}
                    >
                      Belum ada daftar Sekolah / Desa-Posyandu (3B) / Sampel pada Hari Ke-{selectedDay}. Gunakan <strong>Panel Input Khusus</strong> di atas untuk menambahkan unit sasaran.
                    </td>
                  </tr>
                ) : (
                  activeUraian.map((u, idx) => {
                    const rowTotal =
                      (Number(u.porsiBesar) || 0) +
                      (Number(u.porsiKecil) || 0) +
                      (Number(u.alergiBesar) || 0) +
                      (Number(u.alergiKecil) || 0) +
                      (Number(u.sampelBesar) || 0) +
                      (Number(u.sampelKecil) || 0);
                    const badge = KATEGORI_BADGE_MAP[u.kategori];
                    const isEditingThis = editingRowId === u.id;

                    return (
                      <tr key={u.id}>
                        <td
                          className="border border-black px-2 py-1.5 text-center align-middle"
                          style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#FFFFFF" }}
                        >
                          {idx + 1}
                        </td>
                        <td
                          className="border border-black px-2 py-1.5 text-left align-middle"
                          style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#FFFFFF" }}
                        >
                          <div className="flex items-center justify-between gap-2">
                            <input
                              type="text"
                              value={u.nama}
                              onChange={(e) => handleUraianInlineChange(u.id, "nama", e.target.value)}
                              placeholder="Nama sekolah / desa / posyandu..."
                              className="w-full bg-transparent text-xs text-black focus:outline-none focus:bg-indigo-50/60 px-1 py-0.5 rounded"
                            />
                            {badge && (
                              <span
                                className={`text-[10px] font-semibold px-1.5 py-0.5 rounded border whitespace-nowrap shrink-0 no-print ${badge.color}`}
                              >
                                {badge.label}
                              </span>
                            )}
                          </div>
                          {u.kategori === "posyandu_3b" &&
                            ((u.bumil || 0) > 0 || (u.busui || 0) > 0) && (
                              <span className="block text-[10px] text-emerald-700 px-1 no-print">
                                Bumil: {u.bumil || 0} · Busui: {u.busui || 0}
                              </span>
                            )}
                        </td>
                        <td
                          className="border border-black px-1.5 py-1 text-center align-middle"
                          style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#FFFFFF" }}
                        >
                          <input
                            type="number"
                            min="0"
                            value={numInputVal(u.porsiBesar)}
                            onChange={(e) =>
                              handleUraianInlineChange(u.id, "porsiBesar", e.target.value)
                            }
                            className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-emerald-50 py-0.5 rounded tabular-nums"
                          />
                        </td>
                        <td
                          className="border border-black px-1.5 py-1 text-center align-middle"
                          style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#FFFFFF" }}
                        >
                          <input
                            type="number"
                            min="0"
                            value={numInputVal(u.porsiKecil)}
                            onChange={(e) =>
                              handleUraianInlineChange(u.id, "porsiKecil", e.target.value)
                            }
                            className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-emerald-50 py-0.5 rounded tabular-nums"
                          />
                        </td>
                        <td
                          className="border border-black px-1.5 py-1 text-center align-middle"
                          style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#FFFFFF" }}
                        >
                          <input
                            type="number"
                            min="0"
                            value={numInputVal(u.alergiBesar)}
                            onChange={(e) =>
                              handleUraianInlineChange(u.id, "alergiBesar", e.target.value)
                            }
                            className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-orange-50 py-0.5 rounded tabular-nums"
                          />
                        </td>
                        <td
                          className="border border-black px-1.5 py-1 text-center align-middle"
                          style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#FFFFFF" }}
                        >
                          <input
                            type="number"
                            min="0"
                            value={numInputVal(u.alergiKecil)}
                            onChange={(e) =>
                              handleUraianInlineChange(u.id, "alergiKecil", e.target.value)
                            }
                            className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-orange-50 py-0.5 rounded tabular-nums"
                          />
                        </td>
                        {showSampelColInTable1 && (
                          <>
                            <td
                              className="border border-black px-1.5 py-1 text-center align-middle"
                              style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#EFF6FF" }}
                            >
                              <input
                                type="number"
                                min="0"
                                value={numInputVal(u.sampelBesar)}
                                onChange={(e) =>
                                  handleUraianInlineChange(u.id, "sampelBesar", e.target.value)
                                }
                                className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-blue-100/70 py-0.5 rounded tabular-nums"
                              />
                            </td>
                            <td
                              className="border border-black px-1.5 py-1 text-center align-middle"
                              style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#EFF6FF" }}
                            >
                              <input
                                type="number"
                                min="0"
                                value={numInputVal(u.sampelKecil)}
                                onChange={(e) =>
                                  handleUraianInlineChange(u.id, "sampelKecil", e.target.value)
                                }
                                className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-blue-100/70 py-0.5 rounded tabular-nums"
                              />
                            </td>
                          </>
                        )}
                        <td
                          className="border border-black px-2 py-1.5 text-center font-medium text-black align-middle tabular-nums"
                          style={{ backgroundColor: isEditingThis ? "#FEF9C3" : "#FFFFFF" }}
                        >
                          {rowTotal}
                        </td>
                        <td
                          className="border border-black px-1 py-1 text-center align-middle no-print"
                          style={{ backgroundColor: "#F8FAFC" }}
                        >
                          <div className="flex items-center justify-center gap-0.5">
                            <button
                              type="button"
                              onClick={() => handleStartEditUraianRow(u)}
                              className="p-1 text-indigo-600 hover:bg-indigo-100 rounded cursor-pointer"
                              title="Edit detail baris"
                            >
                              <Edit3 className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleMoveUraianRow(idx, -1)}
                              disabled={idx === 0}
                              className="p-1 text-slate-500 hover:bg-slate-200 disabled:opacity-30 rounded cursor-pointer"
                              title="Geser ke atas"
                            >
                              <ArrowUp className="w-3 h-3" />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleMoveUraianRow(idx, 1)}
                              disabled={idx === activeUraian.length - 1}
                              className="p-1 text-slate-500 hover:bg-slate-200 disabled:opacity-30 rounded cursor-pointer"
                              title="Geser ke bawah"
                            >
                              <ArrowDown className="w-3 h-3" />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleDeleteUraianRow(u.id)}
                              className="p-1 text-rose-600 hover:bg-rose-100 rounded cursor-pointer"
                              title="Hapus baris"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                )}

                {/* Sub. Total Footer Row — Persis Gambar 1 (Abu-abu #A3A3A3) */}
                <tr>
                  <td
                    colSpan={2}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#A3A3A3" }}
                  >
                    Sub. Total
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#A3A3A3" }}
                  >
                    {uraianTotals.pb}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#A3A3A3" }}
                  >
                    {uraianTotals.pk}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#A3A3A3" }}
                  >
                    {uraianTotals.ab}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#A3A3A3" }}
                  >
                    {uraianTotals.ak}
                  </td>
                  {showSampelColInTable1 && (
                    <>
                      <td
                        className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                        style={{ backgroundColor: "#A3A3A3" }}
                      >
                        {uraianTotals.sb}
                      </td>
                      <td
                        className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                        style={{ backgroundColor: "#A3A3A3" }}
                      >
                        {uraianTotals.sk}
                      </td>
                    </>
                  )}
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#A3A3A3" }}
                  >
                    {uraianTotals.grandTotal}
                  </td>
                  <td
                    className="border border-black px-2 py-2 no-print"
                    style={{ backgroundColor: "#A3A3A3" }}
                  />
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* ===================================================================== */}
        {/* TABEL 2 — REKAPITULASI KESELURUHAN (PERSIS GAMBAR 2 + BARIS SAMPEL)  */}
        {/* ===================================================================== */}
        <div
          id="print-area-pm-tabel-2"
          className={`space-y-3 ${printScope === "table1" ? "no-print" : ""} ${
            printScope === "both" && activeUraian.length >= 16 ? "break-before-page" : ""
          }`}
        >

          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200 pb-2">
            <div>
              <h3 className="text-sm sm:text-base font-bold text-black uppercase tracking-tight">
                TABEL 2. REKAPITULASI KESELURUHAN PENERIMA MANFAAT (HARI KE-{selectedDay})
              </h3>
            </div>
            <span className="text-xs font-bold text-indigo-800 bg-indigo-50 border border-indigo-200 px-2.5 py-1 rounded-lg no-print self-start">
              13 Kelompok Sasaran (termasuk Sampel)
            </span>
          </div>

          <div className="overflow-x-auto">
            <table
              className="w-full border-collapse text-xs text-black"
              style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
            >
              <thead>
                {/* Yellow Date Banner Row — Persis Gambar 2 */}
                <tr>
                  <th
                    colSpan={7}
                    className="border border-black px-3 py-2 text-center font-bold text-sm text-black align-middle"
                    style={{ backgroundColor: "#E6FF00" }}
                  >
                    {formattedActiveDateBanner}
                  </th>
                </tr>
                <tr>
                  <th
                    colSpan={2}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    DATA PENERIMA KELOMPOK SASARAN
                  </th>
                  <th
                    colSpan={4}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    JUMLAH PENERIMA MANFAAT
                  </th>
                  <th
                    rowSpan={3}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle w-28"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    TOTAL PENERIMA MANFAAT
                  </th>
                </tr>
                <tr>
                  <th
                    rowSpan={2}
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle w-12"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    No.
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-3 py-2 text-center font-bold text-black align-middle min-w-[240px]"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    Kelompok Sasaran
                  </th>
                  <th
                    colSpan={2}
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    Tanpa Alergi
                  </th>
                  <th
                    colSpan={2}
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#FF6B00" }}
                  >
                    Diet/Alergi
                  </th>
                </tr>
                <tr>
                  <th
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-24"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    Porsi Besar
                  </th>
                  <th
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-24"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    Porsi Kecil
                  </th>
                  <th
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-24"
                    style={{ backgroundColor: "#FF6B00" }}
                  >
                    Porsi Besar
                  </th>
                  <th
                    className="border border-black px-2 py-1.5 text-center font-bold text-black align-middle w-24"
                    style={{ backgroundColor: "#FF6B00" }}
                  >
                    Porsi Kecil
                  </th>
                </tr>
              </thead>
              <tbody>
                {activeSasaran.map((s, idx) => {
                  const rowTotal =
                    (Number(s.porsiBesar) || 0) +
                    (Number(s.porsiKecil) || 0) +
                    (Number(s.alergiBesar) || 0) +
                    (Number(s.alergiKecil) || 0);
                  const isSampelRow = s.id === "sampel";

                  return (
                    <tr key={s.id}>
                      <td
                        className="border border-black px-2 py-1.5 text-center align-middle"
                        style={{ backgroundColor: "#FFFFFF" }}
                      >
                        {idx + 1}
                      </td>
                      <td
                        className="border border-black px-3 py-1.5 text-left align-middle font-normal text-black"
                        style={{ backgroundColor: "#FFFFFF" }}
                      >
                        <span className={isSampelRow ? "font-semibold" : ""}>{s.label}</span>
                      </td>
                      <td
                        className="border border-black px-2 py-1 text-center align-middle"
                        style={{ backgroundColor: "#FFFFFF" }}
                      >
                        <input
                          id={`input-pm-besar-${s.id}`}
                          type="number"
                          min="0"
                          value={numInputVal(s.porsiBesar)}
                          onChange={(e) =>
                            handleCellChangeTable2(
                              s.id,
                              "porsiBesar",
                              Math.max(0, parseInt(e.target.value, 10) || 0)
                            )
                          }
                          className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-emerald-50 py-0.5 rounded tabular-nums"
                        />
                      </td>
                      <td
                        className="border border-black px-2 py-1 text-center align-middle"
                        style={{ backgroundColor: "#FFFFFF" }}
                      >
                        <input
                          id={`input-pm-kecil-${s.id}`}
                          type="number"
                          min="0"
                          value={numInputVal(s.porsiKecil)}
                          onChange={(e) =>
                            handleCellChangeTable2(
                              s.id,
                              "porsiKecil",
                              Math.max(0, parseInt(e.target.value, 10) || 0)
                            )
                          }
                          className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-emerald-50 py-0.5 rounded tabular-nums"
                        />
                      </td>
                      <td
                        className="border border-black px-2 py-1 text-center align-middle"
                        style={{ backgroundColor: "#FFFFFF" }}
                      >
                        <input
                          id={`input-pm-alergi-besar-${s.id}`}
                          type="number"
                          min="0"
                          value={numInputVal(s.alergiBesar)}
                          onChange={(e) =>
                            handleCellChangeTable2(
                              s.id,
                              "alergiBesar",
                              Math.max(0, parseInt(e.target.value, 10) || 0)
                            )
                          }
                          className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-orange-50 py-0.5 rounded tabular-nums"
                        />
                      </td>
                      <td
                        className="border border-black px-2 py-1 text-center align-middle"
                        style={{ backgroundColor: "#FFFFFF" }}
                      >
                        <input
                          id={`input-pm-alergi-kecil-${s.id}`}
                          type="number"
                          min="0"
                          value={numInputVal(s.alergiKecil)}
                          onChange={(e) =>
                            handleCellChangeTable2(
                              s.id,
                              "alergiKecil",
                              Math.max(0, parseInt(e.target.value, 10) || 0)
                            )
                          }
                          className="w-full text-center bg-transparent text-xs text-black focus:outline-none focus:bg-orange-50 py-0.5 rounded tabular-nums"
                        />
                      </td>
                      <td
                        className="border border-black px-2 py-1.5 text-center align-middle tabular-nums text-black"
                        style={{ backgroundColor: "#FFFFFF" }}
                      >
                        {rowTotal > 0 ? rowTotal : ""}
                      </td>
                    </tr>
                  );
                })}

                {/* Sub. Total Footer Row — Persis Gambar 2 (Kuning-Hijau-Oranye-Kuning) */}
                <tr>
                  <td
                    colSpan={2}
                    className="border border-black px-3 py-2 text-center font-bold text-black align-middle"
                    style={{ backgroundColor: "#FFFF00" }}
                  >
                    Sub. Total
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    {totalPorsiBesar}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#1B9E4B" }}
                  >
                    {totalPorsiKecil}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#FF6B00" }}
                  >
                    {totalAlergiBesar}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#FF6B00" }}
                  >
                    {totalAlergiKecil}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold text-black align-middle tabular-nums"
                    style={{ backgroundColor: "#FFFF00" }}
                  >
                    {totalPM}
                  </td>
                </tr>

                {/* Optional RAB & Period Summary Rows (shown on screen; included in print only if toggled ON) */}
                <tr className={includeRabInPrint ? "" : "no-print"}>
                  <td
                    colSpan={2}
                    className="border border-black px-3 py-2 text-left font-bold text-emerald-950 align-middle"
                    style={{ backgroundColor: "#ECFDF5" }}
                  >
                    <div>ESTIMASI RAB HARIAN (H-{selectedDay} · {formattedActiveDateBanner})</div>
                    <div className="text-[10px] font-normal text-emerald-800">
                      Besar {formatIDR(settings.porsiBesarHarga)}/porsi · Kecil {formatIDR(settings.porsiKecilHarga)}/porsi (termasuk Sampel: {formatIDR(dailyBreakdown.rabSampel)})
                    </div>
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold font-mono text-emerald-900 align-middle"
                    style={{ backgroundColor: "#ECFDF5" }}
                  >
                    {formatIDR(rabPorsiBesar)}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold font-mono text-emerald-900 align-middle"
                    style={{ backgroundColor: "#ECFDF5" }}
                  >
                    {formatIDR(rabPorsiKecil)}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold font-mono text-orange-900 align-middle"
                    style={{ backgroundColor: "#FFF7ED" }}
                  >
                    {formatIDR(rabAlergiBesar)}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-bold font-mono text-orange-900 align-middle"
                    style={{ backgroundColor: "#FFF7ED" }}
                  >
                    {formatIDR(rabAlergiKecil)}
                  </td>
                  <td
                    className="border border-black px-2 py-2 text-center font-extrabold font-mono text-emerald-950 align-middle"
                    style={{ backgroundColor: "#D1FAE5" }}
                  >
                    {formatIDR(totalRABHarian)}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Unified Auto Signature Block (always shown at bottom of printed/exported sheet) */}
        {profile && (
          <AutoSignatureBlock
            profile={profile}
            docKey="penerima"
            dateText={activeDateStr ? formatIndonesianDate(activeDateStr) : undefined}
          />
        )}
      </div>

      {/* Classification & Unit Cost Configuration (No Print) */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 space-y-4 no-print">
        <div className="flex items-start gap-2.5 text-xs text-indigo-800">
          <AlertTriangle className="w-4 h-4 text-indigo-600 shrink-0 mt-0.5" />
          <div className="space-y-1">
            <span className="font-bold text-slate-800 text-sm">
              Pedoman Klasifikasi Jenis Porsi &amp; Harga Satuan (Juknis BGN 2025):
            </span>
            <p className="text-slate-600">
              Atur harga satuan porsi dan pilih kelompok sasaran (termasuk <strong>Sampel</strong>) yang masuk ke kalkulasi RAB, Food Cost, dan Nota Pesanan.
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {/* PORSI KECIL CONFIG CARD */}
          <div className="bg-slate-50/70 p-4 rounded-2xl border border-slate-200 space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <span className="px-2.5 py-1 bg-indigo-50 text-indigo-700 text-xs font-bold rounded-lg uppercase tracking-wider self-start border border-indigo-100">
                Porsi Kecil
              </span>
              <div className="flex items-center gap-1.5 text-xs text-slate-600 font-semibold">
                <span>Target Rp</span>
                <input
                  id="input-harga-kecil"
                  type="number"
                  value={settings.porsiKecilHarga}
                  min="0"
                  onChange={(e) =>
                    updateSetting("porsiKecilHarga", Math.max(0, parseInt(e.target.value, 10) || 0))
                  }
                  className="w-24 px-2 py-1 border border-slate-300 rounded-lg text-center text-xs font-extrabold text-emerald-700 bg-white focus:outline-none focus:border-emerald-500 font-mono"
                />
                <span>/ porsi</span>
              </div>
            </div>

            <div className="text-xs text-slate-600 leading-relaxed bg-white p-3 rounded-xl border border-slate-200/80">
              <span className="font-bold text-slate-700 block mb-1">Klasifikasi Sasaran Terpilih:</span>
              <p className="text-slate-500 text-[11px]">
                {porsiKecilLabels || "Belum ada sasaran terpilih."}
              </p>
            </div>

            <div>
              <button
                type="button"
                onClick={() => setShowKecilSasaranEditor(!showKecilSasaranEditor)}
                className="w-full flex items-center justify-between px-3 py-2 bg-indigo-50 hover:bg-indigo-100/80 text-indigo-700 rounded-xl text-[11px] font-bold transition cursor-pointer"
              >
                <span className="flex items-center gap-1">
                  <Settings className="w-3.5 h-3.5" />
                  Pilih Klasifikasi Sasaran Porsi Kecil
                </span>
                {showKecilSasaranEditor ? (
                  <ChevronUp className="w-3.5 h-3.5" />
                ) : (
                  <ChevronDown className="w-3.5 h-3.5" />
                )}
              </button>

              {showKecilSasaranEditor && (
                <div className="mt-2 p-3 bg-white border border-slate-200 rounded-xl grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px] max-h-[210px] overflow-y-auto">
                  {activeSasaran.map((s) => {
                    const isChecked = settings.porsiKecilSasaranIds.includes(s.id);
                    return (
                      <label
                        key={s.id}
                        className={`flex items-center gap-2 p-1.5 rounded-lg border cursor-pointer transition ${
                          isChecked
                            ? "bg-indigo-50/50 border-indigo-200 text-indigo-950 font-semibold"
                            : "border-slate-100 text-slate-600 hover:bg-slate-50"
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => handleToggleKecilSasaran(s.id)}
                          className="rounded text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5"
                        />
                        <span className="truncate">{s.label}</span>
                      </label>
                    );
                  })}
                </div>
              )}
            </div>
          </div>

          {/* PORSI BESAR CONFIG CARD */}
          <div className="bg-slate-50/70 p-4 rounded-2xl border border-slate-200 space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <span className="px-2.5 py-1 bg-cyan-50 text-cyan-700 text-xs font-bold rounded-lg uppercase tracking-wider self-start border border-cyan-100">
                Porsi Besar
              </span>
              <div className="flex items-center gap-1.5 text-xs text-slate-600 font-semibold">
                <span>Target Rp</span>
                <input
                  id="input-harga-besar"
                  type="number"
                  value={settings.porsiBesarHarga}
                  min="0"
                  onChange={(e) =>
                    updateSetting("porsiBesarHarga", Math.max(0, parseInt(e.target.value, 10) || 0))
                  }
                  className="w-24 px-2 py-1 border border-slate-300 rounded-lg text-center text-xs font-extrabold text-emerald-700 bg-white focus:outline-none focus:border-emerald-500 font-mono"
                />
                <span>/ porsi</span>
              </div>
            </div>

            <div className="text-xs text-slate-600 leading-relaxed bg-white p-3 rounded-xl border border-slate-200/80">
              <span className="font-bold text-slate-700 block mb-1">Klasifikasi Sasaran Terpilih:</span>
              <p className="text-slate-500 text-[11px]">
                {porsiBesarLabels || "Belum ada sasaran terpilih."}
              </p>
            </div>

            <div>
              <button
                type="button"
                onClick={() => setShowBesarSasaranEditor(!showBesarSasaranEditor)}
                className="w-full flex items-center justify-between px-3 py-2 bg-cyan-50 hover:bg-cyan-100/80 text-cyan-700 rounded-xl text-[11px] font-bold transition cursor-pointer"
              >
                <span className="flex items-center gap-1">
                  <Settings className="w-3.5 h-3.5" />
                  Pilih Klasifikasi Sasaran Porsi Besar
                </span>
                {showBesarSasaranEditor ? (
                  <ChevronUp className="w-3.5 h-3.5" />
                ) : (
                  <ChevronDown className="w-3.5 h-3.5" />
                )}
              </button>

              {showBesarSasaranEditor && (
                <div className="mt-2 p-3 bg-white border border-slate-200 rounded-xl grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px] max-h-[210px] overflow-y-auto">
                  {activeSasaran.map((s) => {
                    const isChecked = settings.porsiBesarSasaranIds.includes(s.id);
                    return (
                      <label
                        key={s.id}
                        className={`flex items-center gap-2 p-1.5 rounded-lg border cursor-pointer transition ${
                          isChecked
                            ? "bg-cyan-50/50 border-cyan-200 text-cyan-950 font-semibold"
                            : "border-slate-100 text-slate-600 hover:bg-slate-50"
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => handleToggleBesarSasaran(s.id)}
                          className="rounded text-cyan-600 focus:ring-cyan-500 w-3.5 h-3.5"
                        />
                        <span className="truncate">{s.label}</span>
                      </label>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      <PrintPreviewModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title={`Pratinjau Cetak: ${printDocTitle}`}
        elementId={printTargetId}
        filename={printFilename}
        defaultScale={100}
        defaultPaperSize={paperSize as "A4" | "F4"}
        defaultOrientation="auto"
      />
    </div>
  );
}
