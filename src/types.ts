/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export type KopDocKey =
  | "dashboard"
  | "penerima"
  | "menu"
  | "foodcost"
  | "nota"
  | "ikss"
  | "cekgizi"
  | "arsip";

export interface UnifiedKopConfig {
  leftLogoWidth: number; // px, 32..120
  rightLogoWidth: number; // px, 32..120
  linkLeftLogo: boolean;
  linkRightLogo: boolean;
  linkedDocs: Record<KopDocKey, boolean>;
  autoPrintFormat: boolean; // Auto Arial + adaptive font + auto Portrait/Landscape
}

export const DEFAULT_UNIFIED_KOP_CONFIG: UnifiedKopConfig = {
  leftLogoWidth: 56,
  rightLogoWidth: 56,
  linkLeftLogo: true,
  linkRightLogo: true,
  linkedDocs: {
    dashboard: true,
    penerima: true,
    menu: true,
    foodcost: true,
    nota: true,
    ikss: true,
    cekgizi: true,
    arsip: true,
  },
  autoPrintFormat: true,
};

export interface SPPGSignatureConfig {
  kotaTTD: string;
  showKepala: boolean;
  showAhliGizi: boolean;
  showAkuntan: boolean;
  showYayasan: boolean;
  jabatanKepala: string;
  jabatanAhliGizi: string;
  jabatanAkuntan: string;
  jabatanYayasan: string;
  enabledDocs: Record<KopDocKey, boolean>;
}

export const DEFAULT_SIGNATURE_CONFIG: SPPGSignatureConfig = {
  kotaTTD: "",
  showKepala: true,
  showAhliGizi: true,
  showAkuntan: false,
  showYayasan: false,
  jabatanKepala: "Kepala Satuan Pelayanan Pemenuhan Gizi",
  jabatanAhliGizi: "Ahli Gizi / SPPI",
  jabatanAkuntan: "Akuntan SPPG",
  jabatanYayasan: "Ketua Yayasan",
  enabledDocs: {
    dashboard: true,
    penerima: true,
    menu: true,
    foodcost: true,
    nota: true,
    ikss: true,
    cekgizi: false,
    arsip: false,
  },
};

export interface IkssAkgOverride {
  label?: string;
  energi?: number;
  protein?: number;
  lemak?: number;
  kh?: number;
  serat?: number;
}

export interface IkssDataState {
  selectedFoodCostTypes: string[];
  selectedSasaranIds: string[];
  customCategoryTitle?: string;
  customPeriodeLabel?: string;
  showKopOnIkss?: boolean;
  akgOverride?: IkssAkgOverride;
  /** Map keyed by `${groupKey}__day_${hariKe}` -> compressed JPEG data URL */
  photos: Record<string, string>;
}

export interface SPPGProfile {
  idSPPG?: string;
  namaLembaga: string;
  alamat: string;
  kabKota?: string;
  provinsi?: string;
  noHpSPPI?: string;
  namaKepala: string;
  namaAhliGizi: string;
  namaYayasan: string;
  ketuaYayasan: string;
  namaAkuntan: string;
  tahunAnggaran: string;
  periodeDates: string[]; // dynamic dates synchronized from Master Menu
  awalPeriodeBerikutnya: string;
  ttdConfig?: SPPGSignatureConfig;
}

export interface SekolahPM {
  id: string;
  namaSekolah: string;
  alamatDesa: string;
  porsiKecil: number; // TK/PAUD, SD 1-3
  porsiBesar: number; // SD 4-6, SMP, SMA
  alergi?: number;    // Jumlah siswa alergi
}

export interface TigaBPM {
  id: string;
  namaDesa: string;
  bumil: number;
  balita: number;
  busui: number;
  mpAsi?: number;    // PM PM-ASI
  alergi?: number;   // Jumlah PM 3B alergi
}

export interface KelompokSasaranPM {
  id: string;
  label: string;
  porsiKecil: number;
  porsiBesar: number;
  alergiKecil: number;
  alergiBesar: number;
}

export interface HariPM {
  hariKe: number;
  sasaran: KelompokSasaranPM[];
}

export interface MenuItem {
  namaMenu: string;
  karbohidrat: string;
  laukHewani: string;
  laukNabati: string;
  sayur: string;
  buah?: string;
  susu?: string;
  buahSusu: string;
}

export type MasterMenuDurationType =
  | "1_hari"
  | "1_minggu"
  | "2_minggu"
  | "3_minggu"
  | "4_minggu"
  | "1_bulan";

export interface MasterMenuScheduleConfig {
  durationType: MasterMenuDurationType;
  startDate: string; // YYYY-MM-DD
  calendarMonth: string; // YYYY-MM (used when durationType === "1_bulan")
  workDaysPerWeek: 5 | 6 | 7;
  periodeDates?: string[];
}

export type MasterMenuDietType = "Standar" | "Alergi" | "Kering" | "MP-ASI";

export interface MasterMenuPlan {
  id: string;
  title: string;
  sasaranIds: string[];
  dietType: MasterMenuDietType;
  days: MenuItem[];
}

export interface MasterMenu {
  usiaSekolah: MenuItem[];
  tigaB: MenuItem[];
  mpAsi: MenuItem[];
  usiaSekolahAlergi: MenuItem[];
  tigaBAlergi: MenuItem[];
  mpAsiAlergi: MenuItem[];
  scheduleConfig?: MasterMenuScheduleConfig;
  plans?: MasterMenuPlan[];
  activePlanId?: string;
  lastSentAt?: string;
}

export interface TKPIItem {
  id: string;
  no?: number;
  nama: string;
  sumber: string;
  kategori?: string;
  beratStandar?: number; // g
  bdd: number;       // %
  energi: number;    // kkal
  protein: number;   // g
  lemak: number;     // g
  kh: number;        // g
  serat: number;     // g
  abu: number;       // g
  ca: number;        // mg
  p: number;         // mg
  fe: number;        // mg
  na: number;        // mg
  k: number;         // mg
  cu: number;        // mg
  zn: number;        // mg
  retinol: number;   // mcg
  b_karoten: number; // mcg
  kartotal?: number; // mcg (KARTOTAL)
  thiamin: number;   // mg
  riboflavin: number;// mg
  niasin: number;    // mg
  vit_c: number;     // mg
  air?: number;      // g
  isCustom?: boolean;
}

export interface BahanMakananInput {
  id: string;
  tkpiId: string;
  nama?: string;
  beratBB: number; // gram
  urt: string;     // e.g. "1 gls", "1 ptg"
  hargaSatuan: number; // per Kg atau per Satuan
  bdd?: number;    // % (optional custom override)
  potong?: number | string; // potong per porsi
  ekor?: number | string;   // ekor per porsi
  buah?: number | string;   // buah per porsi
  butir?: number | string;  // butir per porsi
  formula?: "kg" | "potong" | "ekor" | "buah" | "butir"; // formula for total price calculation
  bufferBase?: "kg" | "potong" | "ekor" | "custom";
  bufferCustomVal?: string | number;
  jumlahBufferChoice?: string; // "auto" or format like "[kg|potong|ekor|custom]_[with|without]"
  jumlahBufferCustomVal?: string | number;
}

export interface FoodCostDay {
  jenisMenu: "Basah" | "Alergi" | "Kering" | "MP-ASI" | "3B-Basah" | "3B-Alergi" | "3B-Kering";
  hariKe: number; // 1 to N
  namaMenu?: string;
  tanggal?: string;
  sasaranIds?: string[];
  masterMenuPlanId?: string;
  porsiBesarBahan: BahanMakananInput[];
  porsiKecilBahan: BahanMakananInput[];
  bufferPct: number; // 2.5, 3, 5, or custom
  bumbuPct?: number; // default 10 or custom %
  customPmBesarCount?: number;
  customPmKecilCount?: number;
}

export interface CekGiziInput {
  id: string;
  namaMenu: string;
  tkpiId: string;
  berat: number; // gram
  urt: string;
}

export interface TargetAKG {
  label: string;
  energiMin: number;
  energiMax: number;
  proteinMin: number;
  lemakMin: number;
  khMin: number;
}
