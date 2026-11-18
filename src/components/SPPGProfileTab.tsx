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
  X,
} from "lucide-react";
import AutoSignatureBlock, { resolveSignatureConfig } from "./AutoSignatureBlock";

interface SPPGProfileTabProps {
  profile: SPPGProfile;
  onChange: (updated: SPPGProfile) => void;
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
  { key: "dashboard", label: "Dashboard & Output" },
  { key: "penerima", label: "Penerima Manfaat" },
  { key: "menu", label: "Master Menu" },
  { key: "foodcost", label: "Food Cost" },
  { key: "nota", label: "Nota Pesanan Logistik" },
  { key: "ikss", label: "Menu IKSS" },
  { key: "cekgizi", label: "Cek Gizi" },
  { key: "arsip", label: "Arsip Dokumen" },
];

const TTD_DOC_OPTIONS: Array<{ key: KopDocKey; label: string }> = [
  { key: "dashboard", label: "Dashboard & Output" },
  { key: "penerima", label: "Penerima Manfaat" },
  { key: "menu", label: "Master Menu" },
  { key: "foodcost", label: "Food Cost" },
  { key: "nota", label: "Nota Pesanan Logistik" },
  { key: "ikss", label: "Menu IKSS" },
  { key: "cekgizi", label: "Sandbox / Cek Gizi" },
];

export default function SPPGProfileTab({
  profile,
  onChange,
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
  const [showLeftWidthSlider, setShowLeftWidthSlider] = useState(false);
  const [showRightWidthSlider, setShowRightWidthSlider] = useState(false);

  const showToast = (type: "success" | "warning", text: string) => {
    setFeedbackBanner({ type, text });
    setTimeout(() => {
      setFeedbackBanner((prev) => (prev?.text === text ? null : prev));
    }, 4000);
  };

  const handleChange = (field: keyof SPPGProfile, value: any) => {
    onChange({ ...profile, [field]: value });
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

  return (
    <div id="sppg-profile-container" className="space-y-6 font-sans">
      {/* Global Non-blocking Status Banner */}
      {feedbackBanner && (
        <div
          className={`p-3.5 rounded-xl border text-xs font-semibold flex items-center justify-between gap-3 ${
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

      {/* 1. INFORMASI LEMBAGA SPPG (OTOMATIS SINKRON KE SEMUA LEMBAR KERJA & IKSS) */}
      <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-5">
          <div>
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <Building className="w-5 h-5 text-indigo-600" />
              Informasi Lembaga SPPG
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Identitas di bawah ini terisi otomatis ke seluruh lembar kerja cetak, Menu IKSS (Nilai Gizi &amp; Foto Menu), Nota Pesanan, dan blok tanda tangan.
            </p>
          </div>
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 tabular-nums"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium bg-slate-50/60 placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium bg-slate-50/60 placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium bg-slate-50/60 placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium bg-slate-50/60 placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium bg-slate-50/60 placeholder:text-slate-400"
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
              className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-800 font-medium placeholder:text-slate-400 tabular-nums"
              placeholder="2026"
            />
          </div>
        </div>
      </div>

      {/* 2. PENGATURAN LOGO KIRI (BGN) & LOGO KANAN TERPUSAT + KOP SERENTAK */}
      <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-5">
        <div>
          <h3 className="text-base sm:text-lg font-bold text-slate-900">
            Pengaturan Logo Kiri (BGN) &amp; Logo Kanan Terpusat
          </h3>
          <p className="text-xs sm:text-sm text-slate-500 mt-1">
            Unggah logo KOP surat dan atur lebar ukurannya sebelum diterapkan ke seluruh lembar kerja &amp; cetak/PDF. Pengaturan ini khusus untuk KOP dokumen dan tidak mengubah logo utama aplikasi di header atas.
          </p>
        </div>

        {/* Side-by-Side Logo Cards */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {/* Logo Kiri Card */}
          <div className="p-4 sm:p-5 rounded-2xl border border-slate-200 bg-white space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h4 className="text-sm font-bold text-slate-900">Logo Kiri (Utama BGN)</h4>
                <p className="text-xs text-slate-500 mt-0.5">
                  Sekali ubah, seluruh lembar kerja ikut terupdate
                </p>
              </div>
              <span className="px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-700 text-xs font-bold tabular-nums shrink-0">
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
                  <label className="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl border border-emerald-500 bg-emerald-50/40 hover:bg-emerald-50 text-emerald-800 text-xs font-bold transition cursor-pointer">
                    <ImagePlus className="w-4 h-4 text-emerald-600" />
                    <span>Ubah / Atur Lebar Logo Kiri</span>
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
                    className="px-2.5 py-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-600 text-xs font-semibold transition cursor-pointer inline-flex items-center gap-1"
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
                      className="px-2.5 py-2 rounded-xl text-slate-500 hover:text-rose-600 hover:bg-rose-50 text-xs font-semibold transition cursor-pointer inline-flex items-center gap-1"
                    >
                      <RotateCcw className="w-3.5 h-3.5" />
                      <span>Hapus</span>
                    </button>
                  )}
                </div>

                {showLeftWidthSlider && (
                  <div className="flex items-center gap-3 pt-1">
                    <span className="text-[11px] font-semibold text-slate-500">32px</span>
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
                    <span className="text-[11px] font-semibold text-slate-500">120px</span>
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

          {/* Logo Kanan Card */}
          <div className="p-4 sm:p-5 rounded-2xl border border-slate-200 bg-white space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h4 className="text-sm font-bold text-slate-900">Logo Kanan (SPPG / Daerah)</h4>
                <p className="text-xs text-slate-500 mt-0.5">
                  Opsional di sisi kanan kop dokumen pilihan
                </p>
              </div>
              <span className="px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-200 text-slate-700 text-xs font-bold tabular-nums shrink-0">
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
                  <label className="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-800 text-xs font-bold transition cursor-pointer">
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
                    className="px-2.5 py-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-600 text-xs font-semibold transition cursor-pointer inline-flex items-center gap-1"
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
                      className="px-2.5 py-2 rounded-xl text-rose-600 hover:bg-rose-50 text-xs font-semibold transition cursor-pointer inline-flex items-center gap-1"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>Hapus</span>
                    </button>
                  )}
                </div>

                {showRightWidthSlider && (
                  <div className="flex items-center gap-3 pt-1">
                    <span className="text-[11px] font-semibold text-slate-500">32px</span>
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
                    <span className="text-[11px] font-semibold text-slate-500">120px</span>
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

        {/* Document Selection Checkboxes */}
        <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 space-y-2.5">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
              Dokumen Lembar Kerja yang Mengikuti KOP &amp; Logo Serentak
            </span>
            <div className="flex items-center gap-2 text-xs">
              <span className="text-slate-500 font-semibold flex items-center gap-1">
                <FileText className="w-3.5 h-3.5" /> Kertas Standar:
              </span>
              <button
                type="button"
                onClick={() => setPaperSize && setPaperSize("A4")}
                className={`px-2.5 py-1 rounded-lg font-bold transition cursor-pointer ${
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
                className={`px-2.5 py-1 rounded-lg font-bold transition cursor-pointer ${
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
                  className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-lg border text-xs font-semibold cursor-pointer transition select-none ${
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

        {/* 4-Line KOP Text Editor */}
        <div className="space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
              Teks KOP Surat Serentak (4 Baris)
            </span>
            <button
              type="button"
              onClick={handleSyncKopFromProfile}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-bold transition cursor-pointer"
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

        {/* PRATINJAU KOP SURAT BERSAMA */}
        <div className="space-y-2">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
              PRATINJAU KOP SURAT BERSAMA
            </span>
            <span className="text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-0.5 rounded-md">
              Format Cetak &amp; Unduh: Arial Presisi • Auto Potret / Lansekap • Sinkron Cloud
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

      {/* 3. KONFIGURASI PENANDATANGAN (TTD) OTOMATIS LEMBAR KERJA */}
      <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
          <div>
            <h3 className="text-base sm:text-lg font-bold text-slate-900 flex items-center gap-2">
              <PenTool className="w-5 h-5 text-indigo-600" />
              Konfigurasi Tanda Tangan (TTD) Otomatis Lembar Kerja
            </h3>
            <p className="text-xs sm:text-sm text-slate-500 mt-0.5">
              Atur pejabat penandatangan dan lembar kerja yang membutuhkan blok TTD secara serentak dari Profil SPPG.
            </p>
          </div>
        </div>

        {/* Kota/Tempat TTD */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
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
          <div className="md:col-span-2 flex items-end">
            <p className="text-xs text-slate-500 bg-slate-50 border border-slate-200/80 rounded-xl px-4 py-2.5 w-full">
              Nama pejabat diambil otomatis dari <strong>Informasi Lembaga SPPG</strong> di atas. Centang pejabat yang ditampilkan dan ubah redaksi jabatan jika diperlukan.
            </p>
          </div>
        </div>

        {/* Signers Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {[
            {
              flagKey: "showKepala" as const,
              roleKey: "jabatanKepala" as const,
              label: "Kepala SPPG",
              personName: profile.namaKepala,
              placeholder: "Kepala Satuan Pelayanan Pemenuhan Gizi",
            },
            {
              flagKey: "showAhliGizi" as const,
              roleKey: "jabatanAhliGizi" as const,
              label: "Ahli Gizi / SPPI",
              personName: profile.namaAhliGizi,
              placeholder: "Ahli Gizi / SPPI",
            },
            {
              flagKey: "showAkuntan" as const,
              roleKey: "jabatanAkuntan" as const,
              label: "Akuntan SPPG",
              personName: profile.namaAkuntan,
              placeholder: "Akuntan SPPG",
            },
            {
              flagKey: "showYayasan" as const,
              roleKey: "jabatanYayasan" as const,
              label: "Ketua Yayasan",
              personName: profile.ketuaYayasan,
              placeholder: "Ketua Yayasan",
            },
          ].map((item) => {
            const checked = ttdConfig[item.flagKey];
            return (
              <div
                key={item.flagKey}
                className={`p-4 rounded-xl border transition space-y-2.5 ${
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
              </div>
            );
          })}
        </div>

        {/* Which Worksheets Receive Auto TTD */}
        <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 space-y-2.5">
          <span className="text-xs font-bold text-slate-700 uppercase tracking-wider block">
            Lembar Kerja yang Menampilkan Blok TTD Otomatis
          </span>
          <div className="flex flex-wrap gap-2">
            {TTD_DOC_OPTIONS.map((doc) => {
              const checked = ttdConfig.enabledDocs?.[doc.key] !== false;
              return (
                <label
                  key={doc.key}
                  className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-lg border text-xs font-semibold cursor-pointer transition select-none ${
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

        {/* Live Preview of Auto TTD Block */}
        <div className="space-y-2">
          <span className="text-xs font-bold text-slate-500 uppercase tracking-wider block">
            Pratinjau Blok Tanda Tangan (Format Cetak Arial)
          </span>
          <div className="rounded-xl border border-slate-200 bg-white px-6 pb-5 pt-2">
            <AutoSignatureBlock profile={profile} docKey="nota" forceShow />
          </div>
        </div>
      </div>
    </div>
  );
}
