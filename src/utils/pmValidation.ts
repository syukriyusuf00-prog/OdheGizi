/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { SekolahPM, TigaBPM, KelompokSasaranPM, UraianSasaranPM } from "../types";

export interface PortionFieldValidation {
  raw: string;
  isValid: boolean;
  value: number;
  isNegative: boolean;
  isNonInteger: boolean;
  error: string | null;
}

export interface RegisteredSchoolEntry {
  sekolah: SekolahPM;
  besar: number;
  kecil: number;
  alergi: number;
  total: number;
}

export interface RegisteredSchoolCapacitySummary {
  schoolCount: number;
  registeredBesar: number;
  registeredKecil: number;
  registeredAlergi: number;
  registeredTotal: number;
  registered3BTotal: number;
  byName: Map<string, RegisteredSchoolEntry>;
  effectiveSchoolCapacity: number;
  effectiveCombinedCapacity: number;
  capacitySource: "registered" | "manual" | "unset";
}

export const SCHOOL_SASARAN_IDS = new Set<string>([
  "tk_paud_lb",
  "sd_kelas_1_3",
  "sd_kelas_4_6",
  "smp_mts_smplb",
  "sma_smk_ma",
]);

export const SCHOOL_URAIAN_KATEGORI = new Set<string>([
  "tk_paud_lb",
  "sd_mi_slb",
  "smp_mts_smplb",
  "sma_smk_ma",
]);

/**
 * Validates a string portion input (kecil / besar / alergi / sampel / bumil / busui).
 * Rejects negative values and non-integer numbers with descriptive Indonesian messages.
 */
export function validatePortionString(
  raw: string,
  fieldLabel: string
): PortionFieldValidation {
  const trimmed = (raw ?? "").trim();
  if (trimmed === "") {
    return {
      raw,
      isValid: true,
      value: 0,
      isNegative: false,
      isNonInteger: false,
      error: null,
    };
  }

  // Catch bare minus sign or negative numbers immediately
  if (trimmed === "-" || trimmed.startsWith("-")) {
    return {
      raw,
      isValid: false,
      value: 0,
      isNegative: true,
      isNonInteger: false,
      error: `${fieldLabel} tidak boleh negatif (minimal 0).`,
    };
  }

  const num = Number(trimmed);
  if (!Number.isFinite(num) || Number.isNaN(num)) {
    return {
      raw,
      isValid: false,
      value: 0,
      isNegative: false,
      isNonInteger: true,
      error: `${fieldLabel} harus berupa angka bulat ≥ 0.`,
    };
  }

  if (num < 0) {
    return {
      raw,
      isValid: false,
      value: 0,
      isNegative: true,
      isNonInteger: false,
      error: `${fieldLabel} tidak boleh negatif (minimal 0).`,
    };
  }

  if (!Number.isInteger(num)) {
    return {
      raw,
      isValid: false,
      value: Math.floor(num),
      isNegative: false,
      isNonInteger: true,
      error: `${fieldLabel} harus berupa bilangan bulat tanpa desimal.`,
    };
  }

  return {
    raw,
    isValid: true,
    value: num,
    isNegative: false,
    isNonInteger: false,
    error: null,
  };
}

/**
 * Checks whether a raw input string represents a negative number.
 */
export function isRawNegativeInput(raw: string): boolean {
  const trimmed = (raw ?? "").trim();
  if (!trimmed) return false;
  if (trimmed.startsWith("-")) return true;
  const parsed = Number(trimmed);
  return Number.isFinite(parsed) && parsed < 0;
}

/**
 * Computes registered school capacity from sekolahPM (plus 3B capacity from tigaBPM)
 * and resolves the effective school capacity ceiling (manual override takes precedence when set).
 */
export function computeRegisteredSchoolCapacity(
  sekolahPM: SekolahPM[] = [],
  tigaBPM: TigaBPM[] = [],
  manualSchoolCapacity?: number | null
): RegisteredSchoolCapacitySummary {
  let registeredBesar = 0;
  let registeredKecil = 0;
  let registeredAlergi = 0;
  const byName = new Map<string, RegisteredSchoolEntry>();

  (Array.isArray(sekolahPM) ? sekolahPM : []).forEach((s) => {
    const besar = Math.max(0, Number(s?.porsiBesar) || 0);
    const kecil = Math.max(0, Number(s?.porsiKecil) || 0);
    const alergi = Math.max(0, Number(s?.alergi) || 0);
    const total = besar + kecil + alergi;

    registeredBesar += besar;
    registeredKecil += kecil;
    registeredAlergi += alergi;

    const key = (s?.namaSekolah || "").trim().toLowerCase();
    if (key) {
      byName.set(key, { sekolah: s, besar, kecil, alergi, total });
    }
  });

  let registered3BTotal = 0;
  (Array.isArray(tigaBPM) ? tigaBPM : []).forEach((t) => {
    registered3BTotal +=
      Math.max(0, Number(t?.bumil) || 0) +
      Math.max(0, Number(t?.busui) || 0) +
      Math.max(0, Number(t?.balita) || 0) +
      Math.max(0, Number(t?.mpAsi) || 0) +
      Math.max(0, Number(t?.alergi) || 0);
  });

  const registeredTotal = registeredBesar + registeredKecil + registeredAlergi;
  const validManual =
    typeof manualSchoolCapacity === "number" &&
    Number.isFinite(manualSchoolCapacity) &&
    manualSchoolCapacity > 0
      ? Math.floor(manualSchoolCapacity)
      : 0;

  let effectiveSchoolCapacity = 0;
  let capacitySource: "registered" | "manual" | "unset" = "unset";

  if (validManual > 0) {
    effectiveSchoolCapacity = validManual;
    capacitySource = "manual";
  } else if (registeredTotal > 0) {
    effectiveSchoolCapacity = registeredTotal;
    capacitySource = "registered";
  }

  const effectiveCombinedCapacity =
    effectiveSchoolCapacity > 0
      ? effectiveSchoolCapacity + (capacitySource === "registered" ? registered3BTotal : 0)
      : 0;

  return {
    schoolCount: Array.isArray(sekolahPM) ? sekolahPM.length : 0,
    registeredBesar,
    registeredKecil,
    registeredAlergi,
    registeredTotal,
    registered3BTotal,
    byName,
    effectiveSchoolCapacity,
    effectiveCombinedCapacity,
    capacitySource,
  };
}

/**
 * Computes active day's school-related portions (Besar, Kecil, Alergi, Sampel) from Tabel 2 sasaran rows.
 */
export function computeActiveSchoolPortions(
  sasaranList: KelompokSasaranPM[] = []
): {
  schoolBesar: number;
  schoolKecil: number;
  schoolAlergiBesar: number;
  schoolAlergiKecil: number;
  schoolAlergi: number;
  sampelTotal: number;
  schoolOnlyTotal: number;
  schoolWithSampleTotal: number;
} {
  let schoolBesar = 0;
  let schoolKecil = 0;
  let schoolAlergiBesar = 0;
  let schoolAlergiKecil = 0;
  let sampelTotal = 0;

  (Array.isArray(sasaranList) ? sasaranList : []).forEach((s) => {
    const pb = Math.max(0, Number(s?.porsiBesar) || 0);
    const pk = Math.max(0, Number(s?.porsiKecil) || 0);
    const ab = Math.max(0, Number(s?.alergiBesar) || 0);
    const ak = Math.max(0, Number(s?.alergiKecil) || 0);

    if (SCHOOL_SASARAN_IDS.has(s.id)) {
      schoolBesar += pb;
      schoolKecil += pk;
      schoolAlergiBesar += ab;
      schoolAlergiKecil += ak;
    } else if (s.id === "sampel") {
      sampelTotal += pb + pk + ab + ak;
    }
  });

  const schoolAlergi = schoolAlergiBesar + schoolAlergiKecil;
  const schoolOnlyTotal = schoolBesar + schoolKecil + schoolAlergi;
  const schoolWithSampleTotal = schoolOnlyTotal + sampelTotal;

  return {
    schoolBesar,
    schoolKecil,
    schoolAlergiBesar,
    schoolAlergiKecil,
    schoolAlergi,
    sampelTotal,
    schoolOnlyTotal,
    schoolWithSampleTotal,
  };
}

/**
 * Computes total portions of a single Tabel 1 UraianSasaranPM row.
 */
export function getUraianRowTotal(u: UraianSasaranPM): number {
  return (
    Math.max(0, Number(u.porsiBesar) || 0) +
    Math.max(0, Number(u.porsiKecil) || 0) +
    Math.max(0, Number(u.alergiBesar) || 0) +
    Math.max(0, Number(u.alergiKecil) || 0) +
    Math.max(0, Number(u.sampelBesar) || 0) +
    Math.max(0, Number(u.sampelKecil) || 0)
  );
}
