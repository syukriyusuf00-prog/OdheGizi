/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from "react";
import {
  DEFAULT_SIGNATURE_CONFIG,
  KopDocKey,
  SPPGProfile,
  SPPGSignatureConfig,
} from "../types";

export function resolveSignatureConfig(profile?: SPPGProfile): SPPGSignatureConfig {
  const cfg = profile?.ttdConfig;
  return {
    ...DEFAULT_SIGNATURE_CONFIG,
    ...(cfg || {}),
    showIdSppgFooter: cfg?.showIdSppgFooter ?? true,
    kotaTTD: (cfg?.kotaTTD ?? profile?.kabKota ?? "").trim(),
    enabledDocs: {
      ...DEFAULT_SIGNATURE_CONFIG.enabledDocs,
      ...(cfg?.enabledDocs || {}),
    },
  };
}

export interface SignerEntry {
  key: "kepala" | "ahliGizi" | "akuntan" | "yayasan";
  prefix: string;
  jabatan: string;
  nama: string;
  ttdImage?: string;
}

export function getActiveSigners(profile: SPPGProfile): SignerEntry[] {
  const cfg = resolveSignatureConfig(profile);
  const signers: SignerEntry[] = [];

  if (cfg.showYayasan) {
    signers.push({
      key: "yayasan",
      prefix: "Menyetujui,",
      jabatan: cfg.jabatanYayasan || "Ketua Yayasan",
      nama: profile.ketuaYayasan || "........................................",
      ttdImage: cfg.ttdImageYayasan || "",
    });
  }
  if (cfg.showAkuntan) {
    signers.push({
      key: "akuntan",
      prefix: "Diperiksa oleh,",
      jabatan: cfg.jabatanAkuntan || "Akuntan SPPG",
      nama: profile.namaAkuntan || "........................................",
      ttdImage: cfg.ttdImageAkuntan || "",
    });
  }
  if (cfg.showAhliGizi) {
    signers.push({
      key: "ahliGizi",
      prefix: "Disusun oleh,",
      jabatan: cfg.jabatanAhliGizi || "Ahli Gizi / SPPI",
      nama: profile.namaAhliGizi || "........................................",
      ttdImage: cfg.ttdImageAhliGizi || "",
    });
  }
  if (cfg.showKepala) {
    signers.push({
      key: "kepala",
      prefix: "Mengetahui,",
      jabatan: cfg.jabatanKepala || "Kepala Satuan Pelayanan Pemenuhan Gizi",
      nama: profile.namaKepala || "........................................",
      ttdImage: cfg.ttdImageKepala || "",
    });
  }

  return signers;
}

export function IdSppgFooterStrip({
  profile,
  forceShow = false,
  className = "",
}: {
  profile: SPPGProfile;
  forceShow?: boolean;
  className?: string;
}) {
  const cfg = resolveSignatureConfig(profile);
  if (!forceShow && cfg.showIdSppgFooter === false) return null;

  const idSppg = (profile.idSPPG || "").trim();
  const lembaga = (profile.namaLembaga || "").trim();
  const wilayah = [profile.kabKota, profile.provinsi]
    .map((s) => (s || "").trim())
    .filter(Boolean)
    .join(", ");
  const tahun = (profile.tahunAnggaran || "").trim();

  return (
    <div
      className={`mt-5 pt-2 border-t border-slate-300 flex flex-wrap items-center justify-between gap-x-4 gap-y-1 text-[10px] text-slate-700 print:text-black ${className}`}
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5">
        <span className="font-bold text-slate-900 print:text-black tabular-nums">
          ID SPPG: {idSppg || "Belum Diisi"}
        </span>
        {lembaga && (
          <>
            <span aria-hidden="true">·</span>
            <span className="font-semibold">{lembaga}</span>
          </>
        )}
      </div>
      <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 tabular-nums">
        {wilayah && <span>{wilayah}</span>}
        {wilayah && tahun && <span aria-hidden="true">·</span>}
        {tahun && <span>TA {tahun}</span>}
      </div>
    </div>
  );
}

interface AutoSignatureBlockProps {
  profile: SPPGProfile;
  docKey: KopDocKey;
  dateText?: string;
  forceShow?: boolean;
  showDigitalTtd?: boolean;
  hideIdSppgFooter?: boolean;
  className?: string;
}

export default function AutoSignatureBlock({
  profile,
  docKey,
  dateText,
  forceShow = false,
  showDigitalTtd = true,
  hideIdSppgFooter = false,
  className = "",
}: AutoSignatureBlockProps) {
  const cfg = resolveSignatureConfig(profile);
  const isTtdEnabled = forceShow || cfg.enabledDocs?.[docKey] !== false;
  const showFooter = !hideIdSppgFooter && cfg.showIdSppgFooter !== false;

  const signers = isTtdEnabled ? getActiveSigners(profile) : [];
  if (signers.length === 0 && !showFooter) return null;

  const formattedDate =
    dateText ||
    new Date().toLocaleDateString("id-ID", {
      day: "2-digit",
      month: "long",
      year: "numeric",
    });

  const locationLine = cfg.kotaTTD
    ? `${cfg.kotaTTD}, ${formattedDate}`
    : formattedDate;

  const gridColsClass =
    signers.length === 1
      ? "flex justify-end"
      : signers.length === 2
      ? "grid grid-cols-2 gap-8"
      : signers.length === 3
      ? "grid grid-cols-3 gap-6"
      : "grid grid-cols-2 sm:grid-cols-4 gap-5";

  return (
    <div
      className={`mt-6 pt-3 print:break-inside-avoid break-inside-avoid text-black ${className}`}
      style={{ fontFamily: "Arial, Helvetica, sans-serif" }}
    >
      {signers.length > 0 && (
        <div className={gridColsClass}>
          {signers.map((signer, idx) => {
            const isLast = idx === signers.length - 1;
            return (
              <div
                key={signer.key}
                className={`text-center text-[11px] leading-snug ${
                  signers.length === 1 ? "w-72" : "w-full"
                }`}
              >
                <p className="min-h-[16px] text-slate-800">
                  {isLast ? locationLine : "\u00A0"}
                </p>
                <p className="mt-0.5">{signer.prefix}</p>
                <p className="font-semibold min-h-[28px] flex items-start justify-center">
                  {signer.jabatan}
                </p>
                <div className="h-16 flex items-center justify-center my-0.5">
                  {showDigitalTtd && signer.ttdImage ? (
                    <img
                      src={signer.ttdImage}
                      alt={`TTD ${signer.nama}`}
                      className="max-h-14 max-w-[150px] object-contain"
                      referrerPolicy="no-referrer"
                    />
                  ) : null}
                </div>
                <p className="font-bold underline uppercase tracking-wide text-slate-950">
                  {signer.nama}
                </p>
              </div>
            );
          })}
        </div>
      )}

      {showFooter && <IdSppgFooterStrip profile={profile} />}
    </div>
  );
}
