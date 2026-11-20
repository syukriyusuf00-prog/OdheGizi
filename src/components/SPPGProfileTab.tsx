/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from "react";
import {
  SPPGProfile,
  UnifiedKopConfig,
  DEFAULT_UNIFIED_KOP_CONFIG,
  DEFAULT_SIGNATURE_CONFIG,
  SPPGSignatureConfig,
  KopDocKey,
} from "../types";
import {
  Building,
  Trash2,
  AlertCircle,
  CheckCircle2,
  ImagePlus,
  Link2,
  RotateCcw,
  RefreshCw,
  FileText,
  Sliders,
  PenTool,
  Printer,
  Upload,
  X,
  ArrowRight,
  Save,
} from "lucide-react";
import AutoSignatureBlock, { resolveSignatureConfig } from "./AutoSignatureBlock";
import { KopSuratRenderHeader } from "./KopSuratConfigSection";
import { printElementSmart } from "../lib/printUtils";

const REQUIRED_SPPG_FIELDS: Array<{
  key: keyof SPPGProfile;
  label: string;
  inputId: string;
}> = [
  { key: "idSPPG", label: "ID SPPG", inputId: "input-id-sppg" },
  { key: "namaLembaga", label: "Nama Lembaga / SPPG", inputId: "input-nama-lembaga" },
  { key: "alamat", label: "Alamat Lengkap", inputId: "input-alamat" },
  { key: "kabKota", label: "Kabupaten / Kota", inputId: "input-kab-kota" },
  { key: "provinsi", label: "Provinsi", inputId: "input-provinsi" },
  { key: "noHpSPPI", label: "Nomor Handphone (SPPI / AG)", inputId: "input-no-hp-sppi" },
  { key: "namaKepala", label: "Nama Kepala SPPG", inputId: "input-kepala" },
  { key: "namaAhliGizi", label: "Nama Ahli Gizi / SPPI", inputId: "input-ahli-gizi" },
  { key: "namaYayasan", label: "Nama Yayasan", inputId: "input-yayasan" },
  { key: "ketuaYayasan", label: "Ketua Yayasan", inputId: "input-ketua-yayasan" },
  { key: "namaAkuntan", label: "Nama Akuntan", inputId: "input-nama-akuntan" },
  { key: "tahunAnggaran", label: "Tahun Anggaran", inputId: "input-tahun-anggaran" },
];

interface SPPGProfileTabProps {
  profile: SPPGProfile;
  onChange: (updated: SPPGProfile) => void;
  isOnboardingStep?: boolean;
  onSaveProfile?: (
    updated: SPPGProfile,
    options?: { continueToPm?: boolean }
  ) => Promise<void> | void;
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
  kopConfig?: UnifiedKopConfig;
  setKopConfig?: (val: UnifiedKopConfig) => void;
  paperSize?: "A4" | "F4";
  setPaperSize?: (val: "A4" | "F4") => void;
}

const DOC_SYNC_OPTIONS: Array<{ key: KopDocKey; label: string }> = [
  { key: "profil", label: "Profil SPPG" },
  { key: "dashboard", label: "Dashboard & Output" },
  { key: "penerima", label: "Penerima Manfaat" },
  { key: "menu", label: "Master Menu" },
  { key: "foodcost", label: "Food Cost" },
  { key: "nota", label: "Nota Pesanan Logistik" },
  { key: "ikss", label: "Menu IKSS" },
  { key: "sop", label: "Form SOP" },
  { key: "cekgizi", label: "Cek Gizi" },
  { key: "arsip", label: "Arsip Dokumen" },
];

const TTD_DOC_OPTIONS: Array<{ key: KopDocKey; label: string }> = [
  { key: "profil", label: "Profil SPPG" },
  { key: "menu", label: "Master Menu" },
  { key: "foodcost", label: "Food Cost" },
  { key: "nota", label: "Nota Pesanan Logistik" },
  { key: "sop", label: "Form SOP" },
  { key: "ikss", label: "Menu IKSS" },
  { key: "penerima", label: "Penerima Manfaat" },
  { key: "dashboard", label: "Dashboard & Output" },
  { key: "cekgizi", label: "Sandbox / Cek Gizi" },
];

type TtdImageFieldKey =
  | "ttdImageKepala"
  | "ttdImageAhliGizi"
  | "ttdImageAkuntan"
  | "ttdImageYayasan";

/** Compress uploaded signature scan to a crisp lightweight PNG (<= ~25KB) preserving transparency */
async function compressSignatureImage(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Gagal membaca file gambar TTD."));
    reader.onload = () => {
      const src = typeof reader.result === "string" ? reader.result : "";
      if (!src) {
        reject(new Error("File TTD tidak valid."));
        return;
      }
      const img = new Image();
      img.onerror = () => reject(new Error("Format gambar TTD tidak didukung."));
      img.onload = () => {
        const maxW = 280;
        const maxH = 110;
        const scale = Math.min(1, maxW / img.width, maxH / img.height);
        const w = Math.max(1, Math.round(img.width * scale));
        const h = Math.max(1, Math.round(img.height * scale));
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          resolve(src);
          return;
        }
        ctx.clearRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        // Prefer PNG to keep transparent backgrounds crisp; fall back to JPEG if unusually large
        let out = canvas.toDataURL("image/png");
        if (out.length > 48000) {
          ctx.fillStyle = "#ffffff";
          ctx.globalCompositeOperation = "destination-over";
          ctx.fillRect(0, 0, w, h);
          out = canvas.toDataURL("image/jpeg", 0.78);
        }
        resolve(out);
      };
      img.src = src;
    };
    reader.readAsDataURL(file);
  });
}

export default function SPPGProfileTab({
  profile,
  onChange,
  isOnboardingStep = false,
  onSaveProfile,
  kopLine1 = "BADAN GIZI NASIONAL",
  setKopLine1,
  kopLine2 = "SATUAN PELAYANAN PEMENUHAN GIZI",
  setKopLine2,
  kopLine3 = "SPPG MUNA BARAT SAWERIGADI ONDOKE",
  setKopLine3,
  kopLine4 = "Alamat : Jl. Poros Sawerigadi No. 12, Ondoke, Muna Barat, Sulawesi Tenggara",
  setKopLine4,
  leftLogo = "",
  setLeftLogo,
  rightLogo = "",
  setRightLogo,
  kopConfig = DEFAULT_UNIFIED_KOP_CONFIG,
  setKopConfig,
  paperSize = "A4",
  setPaperSize,
}: SPPGProfileTabProps) {
  const [feedbackBanner, setFeedbackBanner] = useState<{
    type: "success" | "warning";
    text: string;
  } | null>(null);
  const [kopSubSection, setKopSubSection] = useState<"kop" | "ttd">("kop");
  const [showLeftWidthSlider, setShowLeftWidthSlider] = useState(false);
  const [showRightWidthSlider, setShowRightWidthSlider] = useState(false);
  const [showMissingHighlight, setShowMissingHighlight] = useState(false);
  const [isSavingProfile, setIsSavingProfile] = useState(false);

  const missingFields = REQUIRED_SPPG_FIELDS.filter(
    (f) => !String(profile[f.key] || "").trim()
  );
  const filledFieldCount = REQUIRED_SPPG_FIELDS.length - missingFields.length;
  const isAllProfileFilled = missingFields.length === 0;
  const isFieldMissing = (key: keyof SPPGProfile) =>
    showMissingHighlight && !String(profile[key] || "").trim();

  const showToast = (type: "success" | "warning", text: string) => {
    setFeedbackBanner({ type, text });
    setTimeout(() => {
      setFeedbackBanner((prev) => (prev?.text === text ? null : prev));
    }, 4500);
  };

  const handleChange = (field: keyof SPPGProfile, value: any) => {
    onChange({ ...profile, [field]: value });
  };

  const handleSaveAndProceed = async (continueToPm: boolean) => {
    if (!isAllProfileFilled) {
      setShowMissingHighlight(true);
      const firstMissing = missingFields[0];
      showToast(
        "warning",
        `Mohon lengkapi semua kolom Profil SPPG (${missingFields.length} kolom belum terisi: ${missingFields
          .slice(0, 3)
          .map((m) => m.label)
          .join(", ")}${missingFields.length > 3 ? ", ..." : ""}).`
      );
      if (firstMissing) {
        const el = document.getElementById(firstMissing.inputId);
        el?.scrollIntoView({ behavior: "smooth", block: "center" });
        setTimeout(() => el?.focus(), 250);
      }
      return;
    }

    setShowMissingHighlight(false);
    setIsSavingProfile(true);
    try {
      if (setKopLine3 && !kopLine3.trim() && profile.namaLembaga?.trim()) {
        setKopLine3(profile.namaLembaga.trim().toUpperCase());
      }
      if (setKopLine4 && !kopLine4.trim() && profile.alamat?.trim()) {
        setKopLine4(`Alamat : ${profile.alamat.trim()}`);
      }
      if (onSaveProfile) {
        await onSaveProfile(profile, { continueToPm });
      }
      if (!continueToPm) {
        showToast("success", "Seluruh data Profil SPPG berhasil disimpan & disinkronkan ke Cloud.");
      }
    } finally {
      setIsSavingProfile(false);
    }
  };

  const ttdConfig = resolveSignatureConfig(profile);

  const updateTtdConfig = (patch: Partial<SPPGSignatureConfig>) => {
    const next: SPPGSignatureConfig = {
      ...ttdConfig,
      ...patch,
      enabledDocs: {
        ...DEFAULT_SIGNATURE_CONFIG.enabledDocs,
        ...(ttdConfig.enabledDocs || {}),
        ...(patch.enabledDocs || {}),
      },
    };
    onChange({ ...profile, ttdConfig: next });
  };

  const toggleTtdDoc = (key: KopDocKey) => {
    const current = ttdConfig.enabledDocs?.[key] !== false;
    updateTtdConfig({
      enabledDocs: {
        ...ttdConfig.enabledDocs,
        [key]: !current,
      },
    });
  };

  const handleSignatureImageUpload = async (
    e: React.ChangeEvent<HTMLInputElement>,
    fieldKey: TtdImageFieldKey,
    roleLabel: string
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const compressed = await compressSignatureImage(file);
      updateTtdConfig({ [fieldKey]: compressed });
      showToast(
        "success",
        `Gambar TTD ${roleLabel} berhasil disimpan dan otomatis diterapkan ke lembar kerja aktif.`
      );
    } catch (err: any) {
      showToast("warning", err?.message || "Gagal memproses gambar TTD.");
    } finally {
      e.target.value = "";
    }
  };

  const updateKopConfig = (patch: Partial<UnifiedKopConfig>) => {
    if (!setKopConfig) return;
    setKopConfig({
      ...kopConfig,
      ...patch,
      linkedDocs: {
        ...DEFAULT_UNIFIED_KOP_CONFIG.linkedDocs,
        ...(kopConfig.linkedDocs || {}),
        ...(patch.linkedDocs || {}),
      },
    });
  };

  const toggleDocSync = (key: KopDocKey) => {
    if (!setKopConfig) return;
    const current = kopConfig.linkedDocs?.[key] !== false;
    updateKopConfig({
      linkedDocs: {
        ...DEFAULT_UNIFIED_KOP_CONFIG.linkedDocs,
        ...(kopConfig.linkedDocs || {}),
        [key]: !current,
      },
    });
  };

  const handleLogoFileChange = (
    e: React.ChangeEvent<HTMLInputElement>,
    side: "left" | "right"
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const result = event.target?.result;
      if (typeof result === "string") {
        if (side === "left" && setLeftLogo) {
          setLeftLogo(result);
          showToast(
            "success",
            "Logo Kiri diperbarui dan otomatis diterapkan ke seluruh lembar kerja terpilih."
          );
        } else if (side === "right" && setRightLogo) {
          setRightLogo(result);
          showToast(
            "success",
            "Logo Kanan berhasil diunggah dan dihubungkan ke kop dokumen."
          );
        }
      }
    };
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  const handleSyncKopFromProfile = () => {
    if (setKopLine3 && profile.namaLembaga?.trim()) {
      setKopLine3(profile.namaLembaga.trim().toUpperCase());
    }
    if (setKopLine4 && profile.alamat?.trim()) {
      setKopLine4(`Alamat : ${profile.alamat.trim()}`);
    }
    showToast(
      "success",
      "Baris KOP disinkronkan dari Nama Lembaga & Alamat Profil SPPG."
    );
  };

  const previewLeftPx = Math.min(120, Math.max(32, Number(kopConfig.leftLogoWidth) || 56));
  const previewRightPx = Math.min(120, Math.max(32, Number(kopConfig.rightLogoWidth) || 56));
  const previewShowLeft = kopConfig.linkLeftLogo && !!leftLogo;
  const previewShowRight = kopConfig.linkRightLogo && !!rightLogo;

  const signerItems: Array<{
    flagKey: "showKepala" | "showAhliGizi" | "showAkuntan" | "showYayasan";
    roleKey: "jabatanKepala" | "jabatanAhliGizi" | "jabatanAkuntan" | "jabatanYayasan";
    imageKey: TtdImageFieldKey;
    label: string;
    personName: string;
    placeholder: string;
  }> = [
    {
      flagKey: "showKepala",
      roleKey: "jabatanKepala",
      imageKey: "ttdImageKepala",
      label: "Kepala SPPG",
      personName: profile.namaKepala,
      placeholder: "Kepala Satuan Pelayanan Pemenuhan Gizi",
    },
    {
      flagKey: "showAhliGizi",
      roleKey: "jabatanAhliGizi",
      imageKey: "ttdImageAhliGizi",
      label: "Ahli Gizi / SPPI",
      personName: profile.namaAhliGizi,
      placeholder: "Ahli Gizi / SPPI",
    },
    {
      flagKey: "showAkuntan",
      roleKey: "jabatanAkuntan",
      imageKey: "ttdImageAkuntan",
      label: "Akuntan SPPG",
      personName: profile.namaAkuntan,
      placeholder: "Akuntan SPPG",
    },
    {
      flagKey: "showYayasan",
      roleKey: "jabatanYayasan",
      imageKey: "ttdImageYayasan",
      label: "Ketua Yayasan",
      personName: profile.ketuaYayasan,
      placeholder: "Ketua Yayasan",
    },
  ];

  return (
    <div id="sppg-profile-container" className="space-y-6 font-sans">
      {/* Guided First-Login Onboarding Step 1 Banner */}
      {isOnboardingStep && (
        <div
          id="sppg-onboarding-banner"
          className="bg-gradient-to-r from-indigo-700 via-blue-700 to-emerald-700 rounded-2xl p-5 text-white shadow-sm flex flex-col lg:flex-row lg:items-center justify-between gap-4 no-print"
        >
          <div className="space-y-1.5">
            <div className="flex items-center gap-2 text-[11px] font-extrabold uppercase tracking-wider text-indigo-100">
              <span>Langkah 1 dari 2 · Pengaturan Awal Akun SPPG</span>
              <span>·</span>
              <span className="font-mono">
                {filledFieldCount}/{REQUIRED_SPPG_FIELDS.length} Kolom Terisi
              </span>
            </div>
            <h2 className="text-lg font-extrabold tracking-tight">
              Lengkapi Seluruh Profil SPPG Anda Terlebih Dahulu
            </h2>
            <p className="text-xs text-indigo-100 leading-relaxed max-w-2xl">
              Isi seluruh 12 kolom Informasi Lembaga SPPG di bawah ini (ID SPPG, Nama Lembaga, Alamat, Kab/Kota, Provinsi, No. HP, Kepala SPPG, Ahli Gizi, Yayasan, Ketua Yayasan, Akuntan, Tahun Anggaran). Setelah lengkap, klik tombol <strong>Simpan Profil &amp; Lanjut Update PM</strong> untuk langsung diarahkan ke menu Penerima Manfaat.
            </p>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              id="btn-onboarding-save-sppg-top"
              disabled={isSavingProfile}
              onClick={() => handleSaveAndProceed(true)}
              className={`px-4 py-2.5 rounded-xl text-xs font-extrabold flex items-center gap-2 transition cursor-pointer whitespace-nowrap shadow-sm ${
                isAllProfileFilled
                  ? "bg-emerald-400 hover:bg-emerald-300 text-slate-950"
                  : "bg-white/20 hover:bg-white/30 text-white border border-white/30"
              }`}
            >
              <Save className="w-4 h-4" />
              <span>
                {isSavingProfile
                  ? "Menyimpan..."
                  : isAllProfileFilled
                    ? "Simpan Profil & Lanjut Update PM"
                    : `Lengkapi (${filledFieldCount}/${REQUIRED_SPPG_FIELDS.length})`}
              </span>
              <ArrowRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* Global Non-blocking Status Banner */}
      {feedbackBanner && (
        <div
          className={`p-3.5 rounded-xl border text-xs font-semibold flex items-center justify-between gap-3 no-print ${
            feedbackBanner.type === "success"
              ? "bg-emerald-50 border-emerald-200 text-emerald-800"
              : "bg-amber-50 border-amber-200 text-amber-800"
          }`}
        >
          <div className="flex items-center gap-2">
            {feedbackBanner.type === "success" ? (
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            ) : (
              <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
            )}
            <span>{feedbackBanner.text}</span>
          </div>
          <button
            type="button"
            onClick={() => setFeedbackBanner(null)}
            className="p-1 hover:bg-black/5 rounded-lg transition cursor-pointer"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* 1. INFORMASI LEMBAGA SPPG (OTOMATIS SINKRON KE SEMUA LEMBAR KERJA, IKSS, & FORM SOP) */}
      <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs no-print">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-5">
          <div>
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <Building className="w-5 h-5 text-indigo-600" />
              Informasi Lembaga SPPG
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Identitas di bawah ini terisi otomatis ke seluruh lembar kerja cetak, Master Menu, Food Cost, Nota Pesanan, Menu IKSS, Form SOP, serta footer ID SPPG dan blok TTD.
            </p>
          </div>
          <span
            className={`text-xs font-bold tabular-nums px-3 py-1 rounded-lg border shrink-0 self-start sm:self-auto ${
              isAllProfileFilled
                ? "bg-emerald-50 border-emerald-200 text-emerald-700"
                : "bg-amber-50 border-amber-200 text-amber-800"
            }`}
          >
            {filledFieldCount}/{REQUIRED_SPPG_FIELDS.length} Data Terisi
          </span>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              ID SPPG
            </label>
            <input
              id="input-id-sppg"
              type="text"
              value={profile.idSPPG || ""}
              onChange={(e) => handleChange("idSPPG", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-semibold tabular-nums placeholder:text-slate-400 ${
                isFieldMissing("idSPPG") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200"
              }`}
              placeholder="Contoh: AT9ENIQ6"
            />
          </div>

          <div className="space-y-1.5 lg:col-span-2">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Nama Lembaga / SPPG
            </label>
            <input
              id="input-nama-lembaga"
              type="text"
              value={profile.namaLembaga}
              onChange={(e) => handleChange("namaLembaga", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("namaLembaga") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200"
              }`}
              placeholder="SPPG Muna Barat Sawerigadi Ondoke"
            />
          </div>

          <div className="space-y-1.5 md:col-span-2 lg:col-span-3">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Alamat Lengkap
            </label>
            <input
              id="input-alamat"
              type="text"
              value={profile.alamat}
              onChange={(e) => handleChange("alamat", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("alamat") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200"
              }`}
              placeholder="Jl. Poros Sawerigadi No. 12, Ondoke, Muna Barat, Sulawesi Tenggara"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Kabupaten / Kota
            </label>
            <input
              id="input-kab-kota"
              type="text"
              value={profile.kabKota || ""}
              onChange={(e) => handleChange("kabKota", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("kabKota") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200"
              }`}
              placeholder="Contoh: MUNA BARAT"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Provinsi
            </label>
            <input
              id="input-provinsi"
              type="text"
              value={profile.provinsi || ""}
              onChange={(e) => handleChange("provinsi", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("provinsi") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200"
              }`}
              placeholder="Contoh: SULAWESI TENGGARA"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Nomor Handphone (SPPI / AG)
            </label>
            <input
              id="input-no-hp-sppi"
              type="text"
              value={profile.noHpSPPI || ""}
              onChange={(e) => handleChange("noHpSPPI", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 tabular-nums ${
                isFieldMissing("noHpSPPI") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200"
              }`}
              placeholder="Contoh: 082271095251"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Nama Kepala SPPG
            </label>
            <input
              id="input-kepala"
              type="text"
              value={profile.namaKepala}
              onChange={(e) => handleChange("namaKepala", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("namaKepala") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200 bg-slate-50/60"
              }`}
              placeholder="JABARNUR SALIM, S. H."
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Nama Ahli Gizi / SPPI
            </label>
            <input
              id="input-ahli-gizi"
              type="text"
              value={profile.namaAhliGizi}
              onChange={(e) => handleChange("namaAhliGizi", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("namaAhliGizi") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200 bg-slate-50/60"
              }`}
              placeholder="LA ODE MUHAMMAD SUKRI YUSUF, S.Tr.Gz"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Nama Yayasan
            </label>
            <input
              id="input-yayasan"
              type="text"
              value={profile.namaYayasan}
              onChange={(e) => handleChange("namaYayasan", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("namaYayasan") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200 bg-slate-50/60"
              }`}
              placeholder="LAWOROKU"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Ketua Yayasan
            </label>
            <input
              id="input-ketua-yayasan"
              type="text"
              value={profile.ketuaYayasan}
              onChange={(e) => handleChange("ketuaYayasan", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("ketuaYayasan") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200 bg-slate-50/60"
              }`}
              placeholder="ALI SAMDIN"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Nama Akuntan
            </label>
            <input
              id="input-nama-akuntan"
              type="text"
              value={profile.namaAkuntan}
              onChange={(e) => handleChange("namaAkuntan", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 ${
                isFieldMissing("namaAkuntan") ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300" : "border-slate-200 bg-slate-50/60"
              }`}
              placeholder="NUR ANISSA SA'BAN RINDO, S.Ak"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
              Tahun Anggaran
            </label>
            <input
              id="input-tahun-anggaran"
              type="text"
              value={profile.tahunAnggaran}
              onChange={(e) => handleChange("tahunAnggaran", e.target.value)}
              className={`w-full px-3.5 py-2.5 border rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 tabular-nums ${
                isFieldMissing("tahunAnggaran")
                  ? "border-amber-400 bg-amber-50/50 ring-1 ring-amber-300"
                  : "border-slate-200"
              }`}
              placeholder="2026"
            />
          </div>
        </div>

        {/* Save Action Bar right below Informasi Lembaga SPPG */}
        <div className="mt-6 pt-5 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="text-xs text-slate-600">
            {isAllProfileFilled ? (
              <span className="inline-flex items-center gap-1.5 font-semibold text-emerald-700">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                Semua 12 kolom Profil SPPG sudah lengkap dan siap disimpan.
              </span>
            ) : (
              <span className="inline-flex items-center gap-1.5 font-medium text-amber-700">
                <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
                Masih ada <strong>{missingFields.length}</strong> kolom Profil SPPG yang belum diisi.
              </span>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              type="button"
              id="btn-save-sppg-profile"
              disabled={isSavingProfile}
              onClick={() => handleSaveAndProceed(Boolean(isOnboardingStep))}
              className={`px-5 py-2.5 rounded-xl text-xs font-extrabold flex items-center gap-2 transition cursor-pointer whitespace-nowrap shadow-xs ${
                isAllProfileFilled
                  ? "bg-emerald-600 hover:bg-emerald-700 text-white"
                  : "bg-indigo-600 hover:bg-indigo-700 text-white"
              }`}
            >
              <Save className="w-4 h-4" />
              <span>
                {isSavingProfile
                  ? "Menyimpan..."
                  : isOnboardingStep
                    ? "Simpan Profil SPPG & Lanjut Update PM"
                    : "Simpan Profil SPPG"}
              </span>
              {isOnboardingStep && <ArrowRight className="w-4 h-4" />}
            </button>
            {!isOnboardingStep && onSaveProfile && (
              <button
                type="button"
                disabled={isSavingProfile}
                onClick={() => handleSaveAndProceed(true)}
                className="px-4 py-2.5 rounded-xl text-xs font-bold bg-indigo-50 hover:bg-indigo-100 text-indigo-700 flex items-center gap-1.5 transition cursor-pointer whitespace-nowrap"
              >
                <span>Simpan &amp; Buka Penerima Manfaat</span>
                <ArrowRight className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>
      </div>

      {/* 2. KONFIGURASI TERPUSAT KOP SURAT, LOGO, TTD & ID SPPG */}
      <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-5 no-print">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4">
          <div>
            <h3 className="text-base sm:text-lg font-bold text-slate-900">
              Konfigurasi KOP Surat, Tanda Tangan (TTD) &amp; ID SPPG Terpusat
            </h3>
            <p className="text-xs sm:text-sm text-slate-500 mt-0.5">
              Atur Logo, 4 baris KOP Surat, pejabat TTD, gambar TTD digital, dan footer ID SPPG untuk diterapkan otomatis ke seluruh lembar kerja cetak.
            </p>
          </div>

          <div className="inline-flex items-center gap-1 p-1 bg-slate-100 rounded-xl shrink-0 self-start md:self-auto">
            <button
              type="button"
              onClick={() => setKopSubSection("kop")}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition whitespace-nowrap cursor-pointer ${
                kopSubSection === "kop"
                  ? "bg-white text-slate-900 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              Logo &amp; Teks KOP Surat
            </button>
            <button
              type="button"
              onClick={() => setKopSubSection("ttd")}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition whitespace-nowrap cursor-pointer inline-flex items-center gap-1.5 ${
                kopSubSection === "ttd"
                  ? "bg-white text-indigo-700 shadow-xs"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <PenTool className="w-3.5 h-3.5" />
              TTD &amp; ID SPPG Lembar Kerja
            </button>
          </div>
        </div>

        {kopSubSection === "kop" ? (
          <div className="space-y-5">
            {/* Side-by-Side Logo Controls */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {/* Logo Kiri */}
              <div className="p-4 sm:p-5 rounded-xl border border-slate-200 bg-white space-y-4">
                <div className="flex items-start justify-between gap-3">
                  <div>
                    <h4 className="text-sm font-bold text-slate-900">Logo Kiri (Utama BGN)</h4>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Khusus KOP dokumen cetak; tidak mengubah logo header aplikasi
                    </p>
                  </div>
                  <span className="text-xs font-bold text-slate-600 tabular-nums shrink-0">
                    {previewLeftPx}px
                  </span>
                </div>

                <div className="flex flex-wrap items-center gap-4">
                  <div className="w-20 h-20 rounded-xl border border-dashed border-slate-300 bg-slate-50/50 flex items-center justify-center p-1.5 shrink-0 overflow-hidden">
                    {leftLogo ? (
                      <img
                        src={leftLogo}
                        alt="Logo Kiri"
                        className="max-w-full max-h-full object-contain"
                        referrerPolicy="no-referrer"
                      />
                    ) : (
                      <span className="text-[11px] font-bold text-slate-400 tracking-wider">KIRI</span>
                    )}
                  </div>

                  <div className="flex-1 min-w-[200px] space-y-2.5">
                    <div className="flex flex-wrap items-center gap-2">
                      <label className="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl border border-emerald-500 bg-emerald-50/40 hover:bg-emerald-50 text-emerald-800 text-xs font-bold transition cursor-pointer whitespace-nowrap">
                        <ImagePlus className="w-4 h-4 text-emerald-600" />
                        <span>Unggah Logo Kiri</span>
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) => handleLogoFileChange(e, "left")}
                          className="hidden"
                        />
                      </label>

                      <button
                        type="button"
                        onClick={() => setShowLeftWidthSlider((v) => !v)}
                        className="px-2.5 py-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-600 text-xs font-semibold transition cursor-pointer inline-flex items-center gap-1 whitespace-nowrap"
                        title="Atur lebar logo kiri"
                      >
                        <Sliders className="w-3.5 h-3.5" />
                        <span>Lebar</span>
                      </button>

                      {setLeftLogo && leftLogo && (
                        <button
                          type="button"
                          onClick={() => {
                            setLeftLogo("");
                            updateKopConfig({ leftLogoWidth: 56 });
                          }}
                          className="px-2.5 py-2 rounded-xl text-slate-500 hover:text-rose-600 hover:bg-rose-50 text-xs font-semibold transition cursor-pointer inline-flex items-center gap-1 whitespace-nowrap"
                        >
                          <RotateCcw className="w-3.5 h-3.5" />
                          <span>Hapus</span>
                        </button>
                      )}
                    </div>

                    {showLeftWidthSlider && (
                      <div className="flex items-center gap-3 pt-1">
                        <span className="text-[11px] font-semibold text-slate-500 tabular-nums">32px</span>
                        <input
                          type="range"
                          min={32}
                          max={120}
                          step={2}
                          value={previewLeftPx}
                          onChange={(e) =>
                            updateKopConfig({ leftLogoWidth: Number(e.target.value) })
                          }
                          className="flex-1 accent-emerald-600 cursor-pointer"
                        />
                        <span className="text-[11px] font-semibold text-slate-500 tabular-nums">120px</span>
                      </div>
                    )}

                    <label className="inline-flex items-center gap-2 text-xs text-slate-700 font-medium cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={kopConfig.linkLeftLogo}
                        onChange={(e) => updateKopConfig({ linkLeftLogo: e.target.checked })}
                        className="w-4 h-4 rounded border-slate-300 accent-emerald-600 cursor-pointer"
                      />
                      <Link2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                      <span>Hubungkan otomatis ke dokumen terpilih</span>
                    </label>
                  </div>
                </div>
              </div>

              {/* Logo Kanan */}
              <div className="p-4 sm:p-5 rounded-xl border border-slate-200 bg-white space-y-4">
                <div className="flex items-start justify-between gap-3">
                  <div>
                    <h4 className="text-sm font-bold text-slate-900">Logo Kanan (SPPG / Daerah)</h4>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Opsional di sisi kanan kop dokumen pilihan
                    </p>
                  </div>
                  <span className="text-xs font-bold text-slate-600 tabular-nums shrink-0">
                    {previewRightPx}px
                  </span>
                </div>

                <div className="flex flex-wrap items-center gap-4">
                  <div className="w-20 h-20 rounded-xl border border-dashed border-slate-300 bg-slate-50/50 flex items-center justify-center p-1.5 shrink-0 overflow-hidden">
                    {rightLogo ? (
                      <img
                        src={rightLogo}
                        alt="Logo Kanan"
                        className="max-w-full max-h-full object-contain"
                        referrerPolicy="no-referrer"
                      />
                    ) : (
                      <span className="text-[11px] font-bold text-slate-400 tracking-wider">KANAN</span>
                    )}
                  </div>

                  <div className="flex-1 min-w-[200px] space-y-2.5">
                    <div className="flex flex-wrap items-center gap-2">
                      <label className="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-800 text-xs font-bold transition cursor-pointer whitespace-nowrap">
                        <ImagePlus className="w-4 h-4 text-emerald-600" />
                        <span>Unggah Logo Kanan</span>
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) => handleLogoFileChange(e, "right")}
                          className="hidden"
                        />
                      </label>

                      <button
                        type="button"
                        onClick={() => setShowRightWidthSlider((v) => !v)}
                        className="px-2.5 py-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-600 text-xs font-semibold transition cursor-pointer inline-flex items-center gap-1 whitespace-nowrap"
                        title="Atur lebar logo kanan"
                      >
                        <Sliders className="w-3.5 h-3.5" />
                        <span>Lebar</span>
                      </button>

                      {rightLogo && setRightLogo && (
                        <button
                          type="button"
                          onClick={() => {
                            setRightLogo("");
                            showToast("success", "Logo Kanan telah dihapus.");
                          }}
                          className="px-2.5 py-2 rounded-xl text-rose-600 hover:bg-rose-50 text-xs font-semibold transition cursor-pointer inline-flex items-center gap-1 whitespace-nowrap"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          <span>Hapus</span>
                        </button>
                      )}
                    </div>

                    {showRightWidthSlider && (
                      <div className="flex items-center gap-3 pt-1">
                        <span className="text-[11px] font-semibold text-slate-500 tabular-nums">32px</span>
                        <input
                          type="range"
                          min={32}
                          max={120}
                          step={2}
                          value={previewRightPx}
                          onChange={(e) =>
                            updateKopConfig({ rightLogoWidth: Number(e.target.value) })
                          }
                          className="flex-1 accent-emerald-600 cursor-pointer"
                        />
                        <span className="text-[11px] font-semibold text-slate-500 tabular-nums">120px</span>
                      </div>
                    )}

                    <label className="inline-flex items-center gap-2 text-xs text-slate-700 font-medium cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={kopConfig.linkRightLogo}
                        onChange={(e) => updateKopConfig({ linkRightLogo: e.target.checked })}
                        className="w-4 h-4 rounded border-slate-300 accent-emerald-600 cursor-pointer"
                      />
                      <Link2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                      <span>Hubungkan otomatis ke dokumen terpilih</span>
                    </label>
                  </div>
                </div>
              </div>
            </div>

            {/* 4-Line KOP Text Editor */}
            <div className="space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                  Teks KOP Surat Serentak (4 Baris)
                </span>
                <button
                  type="button"
                  onClick={handleSyncKopFromProfile}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-bold transition cursor-pointer whitespace-nowrap"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Sinkronkan dari Profil SPPG</span>
                </button>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <input
                  type="text"
                  value={kopLine1}
                  onChange={(e) => setKopLine1 && setKopLine1(e.target.value)}
                  placeholder="BADAN GIZI NASIONAL"
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 bg-slate-50/50 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                />
                <input
                  type="text"
                  value={kopLine2}
                  onChange={(e) => setKopLine2 && setKopLine2(e.target.value)}
                  placeholder="SATUAN PELAYANAN PEMENUHAN GIZI"
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 bg-slate-50/50 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                />
                <input
                  type="text"
                  value={kopLine3}
                  onChange={(e) => setKopLine3 && setKopLine3(e.target.value)}
                  placeholder="SPPG MUNA BARAT SAWERIGADI ONDOKE"
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 bg-slate-50/50 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                />
                <input
                  type="text"
                  value={kopLine4}
                  onChange={(e) => setKopLine4 && setKopLine4(e.target.value)}
                  placeholder="Alamat : Jl. Poros Sawerigadi No. 12, Ondoke, Muna Barat, Sulawesi Tenggara"
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-medium text-slate-700 bg-slate-50/50 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                />
              </div>
            </div>

            {/* Document Selection Checkboxes + Paper Size */}
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 space-y-2.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                  Lembar Kerja yang Mengikuti KOP &amp; Logo Serentak
                </span>
                <div className="flex items-center gap-2 text-xs">
                  <span className="text-slate-500 font-semibold flex items-center gap-1">
                    <FileText className="w-3.5 h-3.5" /> Kertas Standar:
                  </span>
                  <button
                    type="button"
                    onClick={() => setPaperSize && setPaperSize("A4")}
                    className={`px-2.5 py-1 rounded-lg font-bold transition cursor-pointer whitespace-nowrap ${
                      paperSize === "A4"
                        ? "bg-indigo-600 text-white"
                        : "bg-white border border-slate-200 text-slate-600 hover:bg-slate-100"
                    }`}
                  >
                    A4
                  </button>
                  <button
                    type="button"
                    onClick={() => setPaperSize && setPaperSize("F4")}
                    className={`px-2.5 py-1 rounded-lg font-bold transition cursor-pointer whitespace-nowrap ${
                      paperSize === "F4"
                        ? "bg-indigo-600 text-white"
                        : "bg-white border border-slate-200 text-slate-600 hover:bg-slate-100"
                    }`}
                  >
                    F4 / Folio
                  </button>
                </div>
              </div>

              <div className="flex flex-wrap gap-2 pt-0.5">
                {DOC_SYNC_OPTIONS.map((doc) => {
                  const checked = kopConfig.linkedDocs?.[doc.key] !== false;
                  return (
                    <label
                      key={doc.key}
                      className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-lg border text-xs font-semibold cursor-pointer transition select-none whitespace-nowrap ${
                        checked
                          ? "bg-white border-emerald-300 text-emerald-900"
                          : "bg-slate-100/70 border-slate-200 text-slate-500"
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleDocSync(doc.key)}
                        className="w-3.5 h-3.5 accent-emerald-600 cursor-pointer"
                      />
                      <span>{doc.label}</span>
                    </label>
                  );
                })}
              </div>
            </div>

            {/* PRATINJAU KOP SURAT BERSAMA */}
            <div className="space-y-2">
              <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
                <span className="font-bold uppercase tracking-wider">
                  Pratinjau KOP Surat Bersama
                </span>
                <span>
                  Format Cetak Arial · Auto Potret / Lansekap · Tersinkron Cloud
                </span>
              </div>

              <div
                className="rounded-xl border border-slate-900 bg-white px-5 py-4"
                style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
              >
                <div className="flex items-center justify-between gap-4">
                  <div
                    className="flex items-center justify-center shrink-0"
                    style={{
                      width: `${Math.max(
                        previewShowLeft ? previewLeftPx + 8 : 0,
                        previewShowRight ? previewRightPx + 8 : 64
                      )}px`,
                    }}
                  >
                    {previewShowLeft && (
                      <img
                        src={leftLogo}
                        alt="Pratinjau Logo Kiri"
                        style={{ width: `${previewLeftPx}px`, height: `${previewLeftPx}px` }}
                        className="object-contain"
                        referrerPolicy="no-referrer"
                      />
                    )}
                  </div>

                  <div className="flex-1 text-center leading-snug">
                    <p className="text-sm sm:text-base font-bold uppercase text-slate-950 tracking-wide">
                      {kopLine1 || "BADAN GIZI NASIONAL"}
                    </p>
                    <p className="text-xs sm:text-sm font-bold uppercase text-slate-900">
                      {kopLine2 || "SATUAN PELAYANAN PEMENUHAN GIZI"}
                    </p>
                    <p className="text-xs sm:text-sm font-bold uppercase text-slate-950">
                      {kopLine3 || profile.namaLembaga || "SPPG MUNA BARAT SAWERIGADI ONDOKE"}
                    </p>
                    <p className="text-[11px] sm:text-xs italic underline underline-offset-2 text-slate-800 mt-1">
                      {kopLine4 ||
                        `Alamat : ${
                          profile.alamat ||
                          "Jl. Poros Sawerigadi No. 12, Ondoke, Muna Barat, Sulawesi Tenggara"
                        }`}
                    </p>
                  </div>

                  <div
                    className="flex items-center justify-center shrink-0"
                    style={{
                      width: `${Math.max(
                        previewShowLeft ? previewLeftPx + 8 : 0,
                        previewShowRight ? previewRightPx + 8 : 64
                      )}px`,
                    }}
                  >
                    {previewShowRight && (
                      <img
                        src={rightLogo}
                        alt="Pratinjau Logo Kanan"
                        style={{ width: `${previewRightPx}px`, height: `${previewRightPx}px` }}
                        className="object-contain"
                        referrerPolicy="no-referrer"
                      />
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : (
          /* INTEGRATED TTD & ID SPPG CONFIGURATION INSIDE KOP SURAT PANEL */
          <div className="space-y-5">
            {/* Kota TTD + ID SPPG Footer Toggle */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="space-y-1.5">
                <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
                  Kota / Tempat Tanda Tangan
                </label>
                <input
                  type="text"
                  value={ttdConfig.kotaTTD}
                  onChange={(e) => updateTtdConfig({ kotaTTD: e.target.value })}
                  placeholder={profile.kabKota || "Contoh: Muna Barat"}
                  className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm text-slate-800 font-medium focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                />
              </div>

              <div className="flex items-end">
                <label className="w-full flex items-start gap-3 p-3 rounded-xl border border-slate-200 bg-slate-50/70 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={ttdConfig.showIdSppgFooter !== false}
                    onChange={(e) =>
                      updateTtdConfig({ showIdSppgFooter: e.target.checked })
                    }
                    className="w-4 h-4 mt-0.5 accent-indigo-600 rounded cursor-pointer shrink-0"
                  />
                  <div className="text-xs leading-snug">
                    <span className="font-bold text-slate-800 block">
                      Tampilkan ID SPPG di Bagian Bawah Setiap Lembar Kerja
                    </span>
                    <span className="text-slate-500 tabular-nums">
                      ID Aktif: {profile.idSPPG?.trim() || "Belum diisi (isi pada Informasi Lembaga di atas)"}
                    </span>
                  </div>
                </label>
              </div>
            </div>

            {/* Signers Grid with Optional Digital TTD Upload */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {signerItems.map((item) => {
                const checked = ttdConfig[item.flagKey];
                const ttdImg = ttdConfig[item.imageKey] || "";
                return (
                  <div
                    key={item.flagKey}
                    className={`p-4 rounded-xl border transition space-y-3 ${
                      checked
                        ? "bg-white border-indigo-200"
                        : "bg-slate-50/70 border-slate-200/80 opacity-80"
                    }`}
                  >
                    <div className="flex items-center justify-between gap-2">
                      <label className="inline-flex items-center gap-2 text-xs font-bold text-slate-800 cursor-pointer select-none">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={(e) =>
                            updateTtdConfig({ [item.flagKey]: e.target.checked })
                          }
                          className="w-4 h-4 accent-indigo-600 rounded cursor-pointer"
                        />
                        <span>{item.label}</span>
                      </label>
                      <span className="text-[11px] font-semibold text-slate-500 truncate max-w-[200px]">
                        {item.personName || "Belum diisi"}
                      </span>
                    </div>

                    <input
                      type="text"
                      value={ttdConfig[item.roleKey]}
                      onChange={(e) =>
                        updateTtdConfig({ [item.roleKey]: e.target.value })
                      }
                      placeholder={item.placeholder}
                      className="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs text-slate-800 font-medium bg-slate-50/40 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                    />

                    {/* Optional Digital Signature Image */}
                    <div className="flex items-center justify-between gap-3 pt-1 border-t border-slate-100">
                      <div className="flex items-center gap-2.5 min-w-0">
                        <div className="w-20 h-10 rounded-lg border border-dashed border-slate-300 bg-slate-50 flex items-center justify-center overflow-hidden shrink-0 px-1">
                          {ttdImg ? (
                            <img
                              src={ttdImg}
                              alt={`TTD ${item.label}`}
                              className="max-h-full max-w-full object-contain"
                              referrerPolicy="no-referrer"
                            />
                          ) : (
                            <span className="text-[10px] text-slate-400">TTD Basah</span>
                          )}
                        </div>
                        <div className="flex flex-wrap items-center gap-1.5">
                          <label className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-700 text-[11px] font-semibold cursor-pointer whitespace-nowrap">
                            <Upload className="w-3 h-3 text-indigo-600" />
                            <span>{ttdImg ? "Ganti TTD" : "Unggah Scan TTD"}</span>
                            <input
                              type="file"
                              accept="image/*"
                              onChange={(e) =>
                                handleSignatureImageUpload(e, item.imageKey, item.label)
                              }
                              className="hidden"
                            />
                          </label>
                          {ttdImg && (
                            <button
                              type="button"
                              onClick={() => updateTtdConfig({ [item.imageKey]: "" })}
                              className="p-1.5 rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition cursor-pointer"
                              title="Hapus gambar TTD"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Which Worksheets Receive Auto TTD & ID SPPG */}
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 space-y-2.5">
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wider block">
                Lembar Kerja yang Menampilkan Blok TTD &amp; ID SPPG Otomatis
              </span>
              <div className="flex flex-wrap gap-2">
                {TTD_DOC_OPTIONS.map((doc) => {
                  const checked = ttdConfig.enabledDocs?.[doc.key] !== false;
                  return (
                    <label
                      key={doc.key}
                      className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-lg border text-xs font-semibold cursor-pointer transition select-none whitespace-nowrap ${
                        checked
                          ? "bg-white border-indigo-300 text-indigo-900"
                          : "bg-slate-100/70 border-slate-200 text-slate-500"
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleTtdDoc(doc.key)}
                        className="w-3.5 h-3.5 accent-indigo-600 cursor-pointer"
                      />
                      <span>{doc.label}</span>
                    </label>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 3. LEMBAR IDENTITAS RESMI PROFIL SPPG + TTD & ID SPPG OTOMATIS (SIAP CETAK) */}
      <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3 no-print">
          <div>
            <h3 className="text-sm sm:text-base font-bold text-slate-900">
              Lembar Identitas, KOP, TTD &amp; ID SPPG (Profil SPPG)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Pratinjau langsung bagian bawah lembar kerja yang terhubung otomatis dengan Profil SPPG, Master Menu, Food Cost, Nota Pesanan, Menu IKSS, dan Form SOP.
            </p>
          </div>
          <button
            type="button"
            onClick={() =>
              printElementSmart("print-area-profil-sppg", {
                title: `Profil SPPG - ${profile.namaLembaga || "SPPG"}`,
                paperSize,
                orientation: "portrait",
              })
            }
            className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition cursor-pointer whitespace-nowrap"
          >
            <Printer className="w-3.5 h-3.5" />
            <span>Cetak Profil SPPG ({paperSize})</span>
          </button>
        </div>

        <div
          id="print-area-profil-sppg"
          className="rounded-xl border border-slate-200 bg-white p-5 sm:p-6 text-slate-950"
          style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
        >
          <KopSuratRenderHeader
            kopLine1={kopLine1}
            kopLine2={kopLine2}
            kopLine3={kopLine3}
            kopLine4={kopLine4}
            leftLogo={leftLogo}
            rightLogo={rightLogo}
            kopConfig={kopConfig}
            docKey="profil"
          />

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-8 gap-y-1.5 text-xs mt-3 mb-2">
            <div className="flex gap-2">
              <span className="w-36 shrink-0 font-semibold text-slate-600">ID SPPG</span>
              <span className="font-bold text-slate-950 tabular-nums">: {profile.idSPPG || "-"}</span>
            </div>
            <div className="flex gap-2">
              <span className="w-36 shrink-0 font-semibold text-slate-600">Tahun Anggaran</span>
              <span className="font-bold text-slate-950 tabular-nums">: {profile.tahunAnggaran || "-"}</span>
            </div>
            <div className="flex gap-2">
              <span className="w-36 shrink-0 font-semibold text-slate-600">Nama Lembaga</span>
              <span className="font-bold text-slate-950">: {profile.namaLembaga || "-"}</span>
            </div>
            <div className="flex gap-2">
              <span className="w-36 shrink-0 font-semibold text-slate-600">Kab/Kota &amp; Provinsi</span>
              <span className="text-slate-900">
                : {[profile.kabKota, profile.provinsi].filter(Boolean).join(", ") || "-"}
              </span>
            </div>
            <div className="flex gap-2">
              <span className="w-36 shrink-0 font-semibold text-slate-600">Kepala SPPG</span>
              <span className="text-slate-900">: {profile.namaKepala || "-"}</span>
            </div>
            <div className="flex gap-2">
              <span className="w-36 shrink-0 font-semibold text-slate-600">Ahli Gizi / SPPI</span>
              <span className="text-slate-900">: {profile.namaAhliGizi || "-"}</span>
            </div>
            <div className="flex gap-2">
              <span className="w-36 shrink-0 font-semibold text-slate-600">Akuntan SPPG</span>
              <span className="text-slate-900">: {profile.namaAkuntan || "-"}</span>
            </div>
            <div className="flex gap-2">
              <span className="w-36 shrink-0 font-semibold text-slate-600">Yayasan / Ketua</span>
              <span className="text-slate-900">
                : {[profile.namaYayasan, profile.ketuaYayasan].filter(Boolean).join(" / ") || "-"}
              </span>
            </div>
          </div>

          <AutoSignatureBlock profile={profile} docKey="profil" />
        </div>
      </div>
    </div>
  );
}
