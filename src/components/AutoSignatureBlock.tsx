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
    });
  }
  if (cfg.showAkuntan) {
    signers.push({
      key: "akuntan",
      prefix: "Diperiksa oleh,",
      jabatan: cfg.jabatanAkuntan || "Akuntan SPPG",
      nama: profile.namaAkuntan || "........................................",
    });
  }
  if (cfg.showAhliGizi) {
    signers.push({
      key: "ahliGizi",
      prefix: "Disusun oleh,",
      jabatan: cfg.jabatanAhliGizi || "Ahli Gizi / SPPI",
      nama: profile.namaAhliGizi || "........................................",
    });
  }
  if (cfg.showKepala) {
    signers.push({
      key: "kepala",
      prefix: "Mengetahui,",
      jabatan: cfg.jabatanKepala || "Kepala Satuan Pelayanan Pemenuhan Gizi",
      nama: profile.namaKepala || "........................................",
    });
  }

  return signers;
}

interface AutoSignatureBlockProps {
  profile: SPPGProfile;
  docKey: KopDocKey;
  dateText?: string;
  forceShow?: boolean;
  className?: string;
}

export default function AutoSignatureBlock({
  profile,
  docKey,
  dateText,
  forceShow = false,
  className = "",
}: AutoSignatureBlockProps) {
  const cfg = resolveSignatureConfig(profile);
  const isEnabled = forceShow || cfg.enabledDocs?.[docKey] !== false;
  if (!isEnabled) return null;

  const signers = getActiveSigners(profile);
  if (signers.length === 0) return null;

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
              <div className="h-16" />
              <p className="font-bold underline uppercase tracking-wide text-slate-950">
                {signer.nama}
              </p>
            </div>
          );
        })}
      </div>
    </div>
  );
}
