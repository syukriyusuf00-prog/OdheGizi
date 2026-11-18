/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { GoogleAuthProvider, signInWithPopup } from "firebase/auth";
import { doc, getDoc, setDoc } from "firebase/firestore";
import { auth, db, reportFirestoreWriteResult } from "./firebase.ts";
import { GOOGLE_DRIVE_SCOPE } from "./googleDrive.ts";

export const GOOGLE_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";

export const WORKSPACE_SCOPES = [
  GOOGLE_DRIVE_SCOPE,
  GOOGLE_SHEETS_SCOPE,
] as const;

export const CANONICAL_SHEET_HEADERS = [
  "Email",
  "Nama Lengkap",
  "Profesi",
  "Nama SPPG",
  "No HP",
  "Peran",
  "Status ACC",
  "Tanggal Mulai Login",
  "Tanggal Berakhir Login",
  "Keterangan",
  "Login Pertama",
  "Login Terakhir",
  "Jumlah Login",
  "Sumber",
] as const;

export interface SheetSyncConfig {
  sheetUrl: string;
  spreadsheetId: string;
  gid: string;
  sheetTitle: string;
  spreadsheetTitle: string;
  webhookUrl?: string;
  autoWebhookEnabled?: boolean;
  lastWebhookPushAt?: string;
  lastWebhookStatus?: string;
  updatedAt: string;
  updatedBy?: string;
}

export interface SheetTabMeta {
  sheetId: number;
  title: string;
  index: number;
}

export interface SpreadsheetMetadata {
  spreadsheetId: string;
  title: string;
  sheets: SheetTabMeta[];
}

export interface SheetsSessionInfo {
  accessToken: string;
  email: string;
  displayName?: string;
  connectedAt: number;
}

export interface ParsedSheetUserRow {
  rowNumber: number;
  email: string;
  namaLengkap: string;
  profesi: string;
  namaSPPG: string;
  noHp: string;
  peran: "USER" | "ADMIN";
  statusPersetujuan: "aktif" | "menunggu" | "diblokir";
  mulaiPada: string;
  berakhirPada: string;
  keterangan: string;
  loginPertama: string;
  loginTerakhir: string;
  loginCount: number;
  sumber: string;
}

export interface SheetSyncPlan {
  sheetTitle: string;
  spreadsheetTitle: string;
  hasHeaderRow: boolean;
  existingRowCount: number;
  toAppend: Array<{ email: string; namaLengkap: string; statusPersetujuan: string }>;
  toUpdate: Array<{ email: string; namaLengkap: string; rowNumber: number; statusPersetujuan: string }>;
  unchangedCount: number;
}

export interface SheetWriteResult {
  appendedCount: number;
  updatedCount: number;
  totalInSheet: number;
  syncedAt: string;
}

const SHEET_CONFIG_LOCAL_KEY = "sisper_user_sheet_sync_config_v1";

// Strict in-memory OAuth token cache ONLY (never persisted to localStorage or sessionStorage)
let inMemorySheetsSession: SheetsSessionInfo | null = null;
let isSheetsConnectingFlag = false;

export function getSheetsSession(): SheetsSessionInfo | null {
  return inMemorySheetsSession;
}

export function clearSheetsSession(): void {
  inMemorySheetsSession = null;
}

export function getIsSheetsConnecting(): boolean {
  return isSheetsConnectingFlag;
}

/**
 * Extracts spreadsheetId and optional gid from any Google Sheets URL or raw spreadsheet ID.
 */
export function parseGoogleSheetUrl(input: string): { spreadsheetId: string; gid: string } | null {
  const trimmed = (input || "").trim();
  if (!trimmed) return null;

  const urlMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]{15,})/);
  const gidMatch = trimmed.match(/[?&#]gid=(\d+)/);

  if (urlMatch && urlMatch[1]) {
    return {
      spreadsheetId: urlMatch[1],
      gid: gidMatch?.[1] || "",
    };
  }

  // Raw spreadsheet ID support
  if (/^[a-zA-Z0-9_-]{20,}$/.test(trimmed)) {
    return {
      spreadsheetId: trimmed,
      gid: "",
    };
  }

  return null;
}

export function buildGoogleSheetUrl(spreadsheetId: string, gid?: string): string {
  const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/edit`;
  return gid ? `${base}#gid=${encodeURIComponent(gid)}` : base;
}

export function quoteSheetTabName(sheetTitle: string): string {
  const escaped = (sheetTitle || "").replace(/'/g, "''");
  return `'${escaped}'`;
}

const normalizeEmail = (val: unknown): string => {
  if (typeof val !== "string") return "";
  return val.trim().toLowerCase();
};

const isValidEmail = (email: string): boolean => {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) && email.length <= 256;
};

export const normalizeSheetDateValue = (raw: unknown): string => {
  const val = String(raw ?? "").trim();
  if (!val) return "";
  const lower = val.toLowerCase();
  if (
    lower === "tanpa batas" ||
    lower === "unlimited" ||
    lower === "selamanya" ||
    lower === "-" ||
    lower === "none" ||
    lower === "tidak ada"
  ) {
    return "";
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
  const slashMatch = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const [, a, b, y] = slashMatch;
    const first = Number(a);
    const second = Number(b);
    const month = first > 12 ? second : first;
    const day = first > 12 ? first : second;
    return `${y}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }
  return val.slice(0, 32);
};

export const normalizeSheetStatus = (raw: unknown): "aktif" | "menunggu" | "diblokir" => {
  const val = String(raw ?? "").trim().toLowerCase();
  if (
    val === "diblokir" ||
    val === "blokir" ||
    val === "blocked" ||
    val === "nonaktif" ||
    val === "non-aktif" ||
    val === "ditolak" ||
    val === "suspend" ||
    val === "suspended"
  ) {
    return "diblokir";
  }
  if (
    val === "menunggu" ||
    val === "pending" ||
    val === "wait" ||
    val === "waiting" ||
    val === "belum acc" ||
    val === "belum disetujui"
  ) {
    return "menunggu";
  }
  // Empty or explicit active values ("aktif", "acc", "disetujui", "active", "ya", "yes", "approved")
  // default to "aktif" so emails added directly in the Spreadsheet automatically gain dashboard access.
  return "aktif";
};

export const normalizeSheetRole = (raw: unknown): "USER" | "ADMIN" => {
  const val = String(raw ?? "").trim().toUpperCase();
  return val === "ADMIN" || val === "ADMINISTRATOR" ? "ADMIN" : "USER";
};

const formatDateValueForSheet = (ts: any): string => {
  if (!ts) return "";
  try {
    let dateObj: Date;
    if (typeof ts?.seconds === "number") {
      dateObj = new Date(ts.seconds * 1000);
    } else if (typeof ts?.toDate === "function") {
      dateObj = ts.toDate();
    } else {
      dateObj = new Date(ts);
    }
    if (Number.isNaN(dateObj.getTime())) return String(ts);
    return dateObj.toISOString();
  } catch {
    return String(ts || "");
  }
};

/**
 * Connects Admin's Google account with Google Sheets (+ Drive) permission via Firebase Auth popup.
 * Token is held strictly in memory only.
 */
export async function connectGoogleSheets(hintEmail?: string): Promise<SheetsSessionInfo> {
  const provider = new GoogleAuthProvider();
  WORKSPACE_SCOPES.forEach((scope) => provider.addScope(scope));

  const customParams: Record<string, string> = {
    prompt: "select_account",
  };
  if (hintEmail && hintEmail.includes("@") && !hintEmail.endsWith("@sispermbg.id")) {
    customParams.login_hint = hintEmail.trim();
  }
  provider.setCustomParameters(customParams);

  isSheetsConnectingFlag = true;
  try {
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    const accessToken = credential?.accessToken;

    if (!accessToken) {
      throw new Error(
        "Token akses Google Sheets tidak diterima. Pastikan izin akses Google Sheets dicentang saat masuk."
      );
    }

    const session: SheetsSessionInfo = {
      accessToken,
      email: result.user?.email || hintEmail || "Akun Google Terhubung",
      displayName: result.user?.displayName || undefined,
      connectedAt: Date.now(),
    };
    inMemorySheetsSession = session;
    return session;
  } catch (err: any) {
    const code = err?.code || "";
    if (code === "auth/unauthorized-domain") {
      const host = typeof window !== "undefined" ? window.location.hostname : "domain ini";
      throw new Error(
        `Domain "${host}" belum terdaftar di Firebase Authentication → Settings → Authorized domains. Tambahkan domain tersebut terlebih dahulu lalu coba lagi.`
      );
    }
    if (code === "auth/popup-blocked") {
      throw new Error(
        "Pop-up Google Sheets diblokir browser. Izinkan pop-up atau buka aplikasi di Tab Baru."
      );
    }
    if (code === "auth/popup-closed-by-user") {
      throw new Error("Jendela koneksi Google Sheets ditutup sebelum otorisasi selesai.");
    }
    if (code === "auth/network-request-failed") {
      throw new Error(
        "Koneksi Google Sheets terhalang iframe browser. Buka aplikasi di Tab Baru lalu hubungkan ulang."
      );
    }
    throw err;
  } finally {
    isSheetsConnectingFlag = false;
  }
}

async function handleSheetsApiResponse<T>(res: Response, actionLabel: string): Promise<T> {
  if (res.status === 401) {
    clearSheetsSession();
    throw new Error(
      `Sesi akses Google Sheets telah berakhir (401) saat ${actionLabel}. Silakan klik "Hubungkan Google Sheets" kembali.`
    );
  }
  if (res.status === 403) {
    let detail = "";
    try {
      const body = await res.json();
      detail = body?.error?.message || "";
    } catch {}
    throw new Error(
      `Akses Google Sheets ditolak (403) saat ${actionLabel}. Pastikan akun Google yang dihubungkan memiliki izin Editor pada spreadsheet ini. ${detail}`
    );
  }
  if (res.status === 404) {
    throw new Error(
      `Spreadsheet tidak ditemukan (404) saat ${actionLabel}. Periksa kembali tautan Google Sheets Anda.`
    );
  }
  if (!res.ok) {
    let detail = "";
    try {
      const body = await res.json();
      detail = body?.error?.message || JSON.stringify(body);
    } catch {
      detail = await res.text();
    }
    throw new Error(`Gagal ${actionLabel} (${res.status}): ${detail}`);
  }
  return res.json() as Promise<T>;
}

/**
 * Fetches spreadsheet metadata including all actual sheet tab names (never hardcoding "Sheet1").
 */
export async function fetchSpreadsheetMetadata(
  accessToken: string,
  spreadsheetId: string
): Promise<SpreadsheetMetadata> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
    spreadsheetId
  )}?fields=properties.title,sheets(properties(sheetId,title,index))`;

  const res = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  const data = await handleSheetsApiResponse<{
    properties?: { title?: string };
    sheets?: Array<{ properties?: { sheetId?: number; title?: string; index?: number } }>;
  }>(res, "membaca metadata spreadsheet");

  const sheets: SheetTabMeta[] = (data.sheets || [])
    .map((s, idx) => ({
      sheetId: typeof s.properties?.sheetId === "number" ? s.properties.sheetId : idx,
      title: s.properties?.title || `Tab ${idx + 1}`,
      index: typeof s.properties?.index === "number" ? s.properties.index : idx,
    }))
    .sort((a, b) => a.index - b.index);

  return {
    spreadsheetId,
    title: data.properties?.title || "Google Spreadsheet",
    sheets,
  };
}

interface DetectedColumns {
  hasHeaderRow: boolean;
  emailIdx: number;
  namaIdx: number;
  profesiIdx: number;
  sppgIdx: number;
  hpIdx: number;
  peranIdx: number;
  statusIdx: number;
  mulaiIdx: number;
  berakhirIdx: number;
  keteranganIdx: number;
  loginPertamaIdx: number;
  loginIdx: number;
  loginCountIdx: number;
  sumberIdx: number;
}

export function detectSheetColumns(rows: string[][]): DetectedColumns {
  if (!Array.isArray(rows) || rows.length === 0) {
    return {
      hasHeaderRow: false,
      emailIdx: 0,
      namaIdx: 1,
      profesiIdx: 2,
      sppgIdx: 3,
      hpIdx: 4,
      peranIdx: 5,
      statusIdx: 6,
      mulaiIdx: 7,
      berakhirIdx: 8,
      keteranganIdx: 9,
      loginPertamaIdx: 10,
      loginIdx: 11,
      loginCountIdx: 12,
      sumberIdx: 13,
    };
  }

  const firstRow = rows[0].map((c) => String(c ?? "").trim().toLowerCase());
  const firstRowHasEmailValue = firstRow.some((c) => isValidEmail(c));

  const findHeader = (patterns: RegExp[]): number => {
    for (let i = 0; i < firstRow.length; i++) {
      const cell = firstRow[i];
      if (!cell) continue;
      if (patterns.some((rx) => rx.test(cell))) return i;
    }
    return -1;
  };

  if (!firstRowHasEmailValue) {
    const emailIdx = findHeader([/\bemail\b/, /e-mail/, /gmail/, /alamat\s*email/, /\bakun\b/]);
    const namaIdx = findHeader([/nama\s*lengkap/, /^nama$/, /nama\s*pengguna/, /full\s*name/, /^name$/]);
    const profesiIdx = findHeader([/profesi/, /jabatan/, /pekerjaan/]);
    const sppgIdx = findHeader([/sppg/, /lembaga/, /instansi/, /dapur/, /unit/]);
    const hpIdx = findHeader([/no\.?\s*hp/, /^hp$/, /whatsapp/, /\bwa\b/, /telepon/, /phone/, /kontak/]);
    const peranIdx = findHeader([/^peran$/, /^role$/, /hak\s*akses/, /level/]);
    const statusIdx = findHeader([/status/, /\bacc\b/, /persetujuan/, /approval/]);
    const mulaiIdx = findHeader([/mulai/, /start\s*date/, /tanggal\s*mulai/]);
    const berakhirIdx = findHeader([/berakhir/, /masa\s*aktif/, /expir/, /kadaluarsa/, /kedaluwarsa/, /berlaku/]);
    const keteranganIdx = findHeader([/keterangan/, /catatan/, /note/, /alasan/]);
    const loginPertamaIdx = findHeader([/login\s*pertama/, /first\s*login/]);
    const loginIdx = findHeader([/login\s*terakhir/, /last\s*login/, /terakhir\s*masuk/]);
    const loginCountIdx = findHeader([/jumlah\s*login/, /login\s*count/, /total\s*login/]);
    const sumberIdx = findHeader([/sumber/, /source/]);

    if (emailIdx >= 0 || namaIdx >= 0 || statusIdx >= 0) {
      return {
        hasHeaderRow: true,
        emailIdx: emailIdx >= 0 ? emailIdx : 0,
        namaIdx: namaIdx >= 0 ? namaIdx : emailIdx === 0 ? 1 : 0,
        profesiIdx,
        sppgIdx,
        hpIdx,
        peranIdx,
        statusIdx,
        mulaiIdx,
        berakhirIdx,
        keteranganIdx,
        loginPertamaIdx,
        loginIdx,
        loginCountIdx,
        sumberIdx,
      };
    }
  }

  let detectedEmailCol = 0;
  for (let r = 0; r < Math.min(rows.length, 15); r++) {
    const row = rows[r] || [];
    const idx = row.findIndex((cell) => isValidEmail(normalizeEmail(cell)));
    if (idx >= 0) {
      detectedEmailCol = idx;
      break;
    }
  }

  return {
    hasHeaderRow: false,
    emailIdx: detectedEmailCol,
    namaIdx: detectedEmailCol === 0 ? 1 : 0,
    profesiIdx: 2,
    sppgIdx: 3,
    hpIdx: 4,
    peranIdx: 5,
    statusIdx: 6,
    mulaiIdx: 7,
    berakhirIdx: 8,
    keteranganIdx: 9,
    loginPertamaIdx: 10,
    loginIdx: 11,
    loginCountIdx: 12,
    sumberIdx: 13,
  };
}

export function parseRowsToUsers(rows: string[][]): {
  hasHeaderRow: boolean;
  columns: DetectedColumns;
  users: ParsedSheetUserRow[];
} {
  const columns = detectSheetColumns(rows);
  const startRow = columns.hasHeaderRow ? 1 : 0;
  const dedupByEmail = new Map<string, ParsedSheetUserRow>();

  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i] || [];
    let rawEmail = normalizeEmail(row[columns.emailIdx]);
    if (!isValidEmail(rawEmail)) {
      const foundCell = row.find((cell) => isValidEmail(normalizeEmail(cell)));
      rawEmail = normalizeEmail(foundCell);
    }
    if (!isValidEmail(rawEmail)) continue;

    const getCol = (idx: number): string => (idx >= 0 && idx < row.length ? String(row[idx] ?? "").trim() : "");
    const rawName = getCol(columns.namaIdx);
    const safeName =
      rawName && normalizeEmail(rawName) !== rawEmail
        ? rawName
        : rawEmail.split("@")[0] || "Pengguna SISPERMBG";
    const parsedCount = Number(getCol(columns.loginCountIdx));

    const parsed: ParsedSheetUserRow = {
      rowNumber: i + 1,
      email: rawEmail,
      namaLengkap: safeName,
      profesi: getCol(columns.profesiIdx) || "Ahli Gizi / Pengelola SPPG",
      namaSPPG: getCol(columns.sppgIdx),
      noHp: getCol(columns.hpIdx),
      peran: normalizeSheetRole(getCol(columns.peranIdx)),
      statusPersetujuan: normalizeSheetStatus(getCol(columns.statusIdx)),
      mulaiPada: normalizeSheetDateValue(getCol(columns.mulaiIdx)),
      berakhirPada: normalizeSheetDateValue(getCol(columns.berakhirIdx)),
      keterangan: getCol(columns.keteranganIdx),
      loginPertama: getCol(columns.loginPertamaIdx),
      loginTerakhir: getCol(columns.loginIdx),
      loginCount: Number.isFinite(parsedCount) && parsedCount > 0 ? Math.floor(parsedCount) : 1,
      sumber: getCol(columns.sumberIdx) || "Spreadsheet",
    };

    dedupByEmail.set(rawEmail, parsed);
  }

  return {
    hasHeaderRow: columns.hasHeaderRow,
    columns,
    users: Array.from(dedupByEmail.values()),
  };
}

/**
 * Reads all values from the selected sheet tab via Google Sheets API v4.
 */
export async function readUsersFromGoogleSheet(
  accessToken: string,
  spreadsheetId: string,
  sheetTitle: string
): Promise<{
  rawRows: string[][];
  hasHeaderRow: boolean;
  columns: DetectedColumns;
  users: ParsedSheetUserRow[];
}> {
  const range = `${quoteSheetTabName(sheetTitle)}!A1:N1000`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
    spreadsheetId
  )}/values/${encodeURIComponent(range)}`;

  const res = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  const data = await handleSheetsApiResponse<{ values?: string[][] }>(
    res,
    `membaca tab "${sheetTitle}"`
  );

  const rawRows: string[][] = Array.isArray(data.values)
    ? data.values.map((r) => (Array.isArray(r) ? r.map((c) => String(c ?? "")) : []))
    : [];

  const parsed = parseRowsToUsers(rawRows);
  return {
    rawRows,
    ...parsed,
  };
}

const buildCanonicalRowValues = (u: Record<string, any>): string[] => {
  const email = normalizeEmail(u.email);
  const berakhir = typeof u.berakhirPada === "string" ? u.berakhirPada.trim() : "";
  return [
    email,
    String(u.namaLengkap || email.split("@")[0] || "Pengguna SISPERMBG").trim(),
    String(u.profesi || "Ahli Gizi / Pengelola SPPG").trim(),
    String(u.namaSPPG || "").trim(),
    String(u.noHp || "").trim(),
    u.peran === "ADMIN" ? "ADMIN" : "USER",
    u.statusPersetujuan === "diblokir"
      ? "diblokir"
      : u.statusPersetujuan === "aktif"
      ? "aktif"
      : "menunggu",
    String(u.mulaiPada || "").trim(),
    berakhir || "Tanpa Batas",
    String(u.keterangan || "").trim(),
    formatDateValueForSheet(u.loginPertama || u.createdAt),
    formatDateValueForSheet(u.loginTerakhir),
    String(Number(u.loginCount) > 0 ? Math.floor(Number(u.loginCount)) : 1),
    String(u.sumber || (u.registeredViaSpreadsheet ? "Spreadsheet" : "Login Google")).trim(),
  ];
};

const mergeRowIntoExistingColumns = (
  existingRow: string[],
  columns: DetectedColumns,
  u: Record<string, any>
): string[] => {
  const canonical = buildCanonicalRowValues(u);
  if (!columns.hasHeaderRow) return canonical;

  const maxIdx = Math.max(
    existingRow.length - 1,
    columns.emailIdx,
    columns.namaIdx,
    columns.profesiIdx,
    columns.sppgIdx,
    columns.hpIdx,
    columns.peranIdx,
    columns.statusIdx,
    columns.mulaiIdx,
    columns.berakhirIdx,
    columns.keteranganIdx,
    columns.loginPertamaIdx,
    columns.loginIdx,
    columns.loginCountIdx,
    columns.sumberIdx,
    0
  );
  const out = Array.from({ length: maxIdx + 1 }, (_, i) => existingRow[i] ?? "");
  if (columns.emailIdx >= 0) out[columns.emailIdx] = canonical[0];
  if (columns.namaIdx >= 0) out[columns.namaIdx] = canonical[1];
  if (columns.profesiIdx >= 0) out[columns.profesiIdx] = canonical[2];
  if (columns.sppgIdx >= 0) out[columns.sppgIdx] = canonical[3];
  if (columns.hpIdx >= 0) out[columns.hpIdx] = canonical[4];
  if (columns.peranIdx >= 0) out[columns.peranIdx] = canonical[5];
  if (columns.statusIdx >= 0) out[columns.statusIdx] = canonical[6];
  if (columns.mulaiIdx >= 0) out[columns.mulaiIdx] = canonical[7];
  if (columns.berakhirIdx >= 0) out[columns.berakhirIdx] = canonical[8];
  if (columns.keteranganIdx >= 0) out[columns.keteranganIdx] = canonical[9];
  if (columns.loginPertamaIdx >= 0) out[columns.loginPertamaIdx] = canonical[10];
  if (columns.loginIdx >= 0) out[columns.loginIdx] = canonical[11];
  if (columns.loginCountIdx >= 0) out[columns.loginCountIdx] = canonical[12];
  if (columns.sumberIdx >= 0) out[columns.sumberIdx] = canonical[13];
  return out;
};

/**
 * Computes a human-readable preview plan before executing any write mutation to Google Sheets
 * (used to populate the mandatory confirmation dialog).
 */
export async function prepareSheetSyncPlan(
  accessToken: string,
  spreadsheetId: string,
  sheetTitle: string,
  spreadsheetTitle: string,
  appUsers: Record<string, any>[]
): Promise<SheetSyncPlan> {
  const { users: existingUsers, hasHeaderRow } = await readUsersFromGoogleSheet(
    accessToken,
    spreadsheetId,
    sheetTitle
  );
  const existingByEmail = new Map(existingUsers.map((u) => [u.email, u]));

  const toAppend: SheetSyncPlan["toAppend"] = [];
  const toUpdate: SheetSyncPlan["toUpdate"] = [];
  let unchangedCount = 0;

  const seenEmails = new Set<string>();
  for (const u of appUsers) {
    const email = normalizeEmail(u.email);
    if (!isValidEmail(email) || seenEmails.has(email)) continue;
    seenEmails.add(email);

    const existing = existingByEmail.get(email);
    const nextStatus = u.statusPersetujuan || "menunggu";
    const nextName = String(u.namaLengkap || email.split("@")[0] || "Pengguna");

    if (!existing) {
      toAppend.push({ email, namaLengkap: nextName, statusPersetujuan: nextStatus });
    } else {
      const nextMulai = String(u.mulaiPada || "").trim();
      const nextBerakhir = String(u.berakhirPada || "").trim();
      const nextKet = String(u.keterangan || "").trim();
      const hasDiff =
        existing.statusPersetujuan !== nextStatus ||
        existing.peran !== (u.peran === "ADMIN" ? "ADMIN" : "USER") ||
        (u.namaSPPG && existing.namaSPPG !== String(u.namaSPPG).trim()) ||
        existing.mulaiPada !== nextMulai ||
        existing.berakhirPada !== nextBerakhir ||
        existing.keterangan !== nextKet ||
        (u.namaLengkap && existing.namaLengkap !== nextName);

      if (hasDiff) {
        toUpdate.push({
          email,
          namaLengkap: nextName,
          rowNumber: existing.rowNumber,
          statusPersetujuan: nextStatus,
        });
      } else {
        unchangedCount++;
      }
    }
  }

  return {
    sheetTitle,
    spreadsheetTitle,
    hasHeaderRow,
    existingRowCount: existingUsers.length,
    toAppend,
    toUpdate,
    unchangedCount,
  };
}

/**
 * Writes app users into Google Sheets (updating existing email rows in-place and appending new email logins).
 * Caller MUST obtain explicit user confirmation before invoking.
 */
export async function syncUsersToGoogleSheet(
  accessToken: string,
  spreadsheetId: string,
  sheetTitle: string,
  appUsers: Record<string, any>[]
): Promise<SheetWriteResult> {
  const { rawRows, hasHeaderRow, columns, users: existingUsers } = await readUsersFromGoogleSheet(
    accessToken,
    spreadsheetId,
    sheetTitle
  );

  const quotedTab = quoteSheetTabName(sheetTitle);
  let effectiveColumns = columns;

  // Initialize header row if the sheet is completely empty
  if (rawRows.length === 0) {
    const headerRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
        spreadsheetId
      )}/values/${encodeURIComponent(`${quotedTab}!A1:N1`)}?valueInputOption=USER_ENTERED`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ values: [Array.from(CANONICAL_SHEET_HEADERS)] }),
      }
    );
    await handleSheetsApiResponse(headerRes, "membuat header kolom spreadsheet");
    effectiveColumns = {
      hasHeaderRow: true,
      emailIdx: 0,
      namaIdx: 1,
      profesiIdx: 2,
      sppgIdx: 3,
      hpIdx: 4,
      peranIdx: 5,
      statusIdx: 6,
      mulaiIdx: 7,
      berakhirIdx: 8,
      keteranganIdx: 9,
      loginPertamaIdx: 10,
      loginIdx: 11,
      loginCountIdx: 12,
      sumberIdx: 13,
    };
  }

  const existingByEmail = new Map(existingUsers.map((u) => [u.email, u]));
  const batchUpdates: Array<{ range: string; values: string[][] }> = [];
  const rowsToAppend: string[][] = [];

  const seenEmails = new Set<string>();
  for (const u of appUsers) {
    const email = normalizeEmail(u.email);
    if (!isValidEmail(email) || seenEmails.has(email)) continue;
    seenEmails.add(email);

    const existing = existingByEmail.get(email);
    if (existing) {
      const existingRaw = rawRows[existing.rowNumber - 1] || [];
      const updatedRow = mergeRowIntoExistingColumns(existingRaw, effectiveColumns, u);
      const changed =
        updatedRow.length !== existingRaw.length ||
        updatedRow.some((val, idx) => (existingRaw[idx] ?? "") !== val);
      if (changed) {
        batchUpdates.push({
          range: `${quotedTab}!A${existing.rowNumber}`,
          values: [updatedRow],
        });
      }
    } else {
      const newRow = hasHeaderRow
        ? mergeRowIntoExistingColumns([], effectiveColumns, u)
        : buildCanonicalRowValues(u);
      rowsToAppend.push(newRow);
    }
  }

  if (batchUpdates.length > 0) {
    const batchRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
        spreadsheetId
      )}/values:batchUpdate`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          valueInputOption: "USER_ENTERED",
          data: batchUpdates,
        }),
      }
    );
    await handleSheetsApiResponse(batchRes, "memperbarui baris pengguna di spreadsheet");
  }

  if (rowsToAppend.length > 0) {
    const appendRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
        spreadsheetId
      )}/values/${encodeURIComponent(`${quotedTab}!A1`)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ values: rowsToAppend }),
      }
    );
    await handleSheetsApiResponse(appendRes, "menambahkan pengguna baru ke spreadsheet");
  }

  return {
    appendedCount: rowsToAppend.length,
    updatedCount: batchUpdates.length,
    totalInSheet: existingUsers.length + rowsToAppend.length,
    syncedAt: new Date().toISOString(),
  };
}

/**
 * Loads the globally saved Spreadsheet Sync Config from Express Real-Time Hub, Firestore, or local cache.
 */
export async function loadGlobalSheetSyncConfig(): Promise<SheetSyncConfig | null> {
  // 1. Express Real-Time Hub (fast, works even during Firestore quota cap)
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 4000);
    const res = await fetch("/api/realtime/sheet-config", {
      method: "GET",
      cache: "no-store",
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (res.ok) {
      const data = await res.json();
      if (data?.config && (data.config.spreadsheetId || data.config.webhookUrl)) {
        try {
          localStorage.setItem(SHEET_CONFIG_LOCAL_KEY, JSON.stringify(data.config));
        } catch {}
        return data.config as SheetSyncConfig;
      }
    }
  } catch {}

  // 2. Firestore global config
  try {
    const snap = await getDoc(doc(db, "configs", "user_sheet_sync"));
    if (snap.exists()) {
      const fsCfg = snap.data() as SheetSyncConfig;
      if (fsCfg && (fsCfg.spreadsheetId || fsCfg.webhookUrl)) {
        try {
          localStorage.setItem(SHEET_CONFIG_LOCAL_KEY, JSON.stringify(fsCfg));
        } catch {}
        fetch("/api/realtime/sheet-config", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(fsCfg),
        }).catch(() => {});
        return fsCfg;
      }
    }
  } catch {}

  // 3. Local cache fallback
  try {
    const raw = localStorage.getItem(SHEET_CONFIG_LOCAL_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as SheetSyncConfig;
      if (parsed && (parsed.spreadsheetId || parsed.webhookUrl)) return parsed;
    }
  } catch {}

  return null;
}

/**
 * Saves Spreadsheet Sync & Apps Script Webhook Config globally across Express Real-Time Hub, Firestore, and local cache.
 */
export async function saveGlobalSheetSyncConfig(config: SheetSyncConfig): Promise<void> {
  const payload: SheetSyncConfig = {
    ...config,
    updatedAt: new Date().toISOString(),
  };

  try {
    localStorage.setItem(SHEET_CONFIG_LOCAL_KEY, JSON.stringify(payload));
  } catch {}

  await Promise.all([
    fetch("/api/realtime/sheet-config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    }).catch(() => {}),
    setDoc(doc(db, "configs", "user_sheet_sync"), payload, { merge: true })
      .then(() => reportFirestoreWriteResult(true))
      .catch((err) => {
        reportFirestoreWriteResult(false, err);
      }),
  ]);
}

/**
 * Pushes parsed Spreadsheet rows to the Express Real-Time Hub cache so any tenant logging in
 * from mobile or desktop is instantly matched — even if the Google Sheet itself is private.
 */
export async function cacheSheetUsersOnServer(
  users: ParsedSheetUserRow[],
  spreadsheetId: string,
  sheetTitle: string
): Promise<void> {
  try {
    await fetch("/api/realtime/sheet-users-cache", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        spreadsheetId,
        sheetTitle,
        users,
        syncedAt: new Date().toISOString(),
      }),
    });
  } catch {}
}

/**
 * Reads a publicly shared ("Anyone with the link can view") Google Sheet via the Express server proxy
 * without requiring an Admin OAuth session.
 */
export async function readPublicSpreadsheetViaServer(params: {
  spreadsheetId: string;
  gid?: string;
  sheetTitle?: string;
}): Promise<{
  ok: boolean;
  users: ParsedSheetUserRow[];
  requiresOAuth?: boolean;
  error?: string;
}> {
  const qs = new URLSearchParams();
  qs.set("spreadsheetId", params.spreadsheetId);
  if (params.gid) qs.set("gid", params.gid);
  if (params.sheetTitle) qs.set("sheetTitle", params.sheetTitle);

  const res = await fetch(`/api/sheets/read-public?${qs.toString()}`, {
    method: "GET",
    cache: "no-store",
  });
  const data = await res.json().catch(() => ({ ok: false, error: "Gagal membaca respons server." }));
  return {
    ok: Boolean(data?.ok),
    users: Array.isArray(data?.users) ? data.users : [],
    requiresOAuth: Boolean(data?.requiresOAuth),
    error: data?.error,
  };
}

/**
 * Sends user access database & live status summary to Admin's Google Sheet via server-side Apps Script Webhook proxy.
 */
export async function pushUsersToAppsScriptWebhook(params: {
  webhookUrl: string;
  users: Record<string, any>[];
}): Promise<{
  ok: boolean;
  pushedCount: number;
  syncedAt: string;
  message: string;
}> {
  const res = await fetch("/api/sheets/webhook-push", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      webhookUrl: params.webhookUrl.trim(),
      users: params.users,
    }),
  });
  const data = await res.json().catch(() => ({
    ok: false,
    pushedCount: 0,
    syncedAt: new Date().toISOString(),
    message: "Gagal membaca respons server webhook.",
  }));
  return {
    ok: Boolean(data?.ok),
    pushedCount: Number(data?.pushedCount) || 0,
    syncedAt: data?.syncedAt || new Date().toISOString(),
    message: data?.message || data?.error || "Gagal mengirim ke Apps Script Webhook.",
  };
}

const escapeCsvCell = (value: unknown): string => {
  const str = String(value ?? "");
  if (/[",\r\n]/.test(str)) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
};

/**
 * Generates an RFC-4180 UTF-8 BOM CSV export of the Admin User Access Database.
 */
export function buildUsersCsvContent(users: Record<string, any>[]): string {
  const headerLine = CANONICAL_SHEET_HEADERS.map(escapeCsvCell).join(",");
  const dataLines = users.map((u) => buildCanonicalRowValues(u).map(escapeCsvCell).join(","));
  return `\uFEFF${[headerLine, ...dataLines].join("\r\n")}`;
}

export function downloadUsersCsv(users: Record<string, any>[]): void {
  const csv = buildUsersCsvContent(users);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const today = new Date().toISOString().slice(0, 10);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Database_Pengguna_SISPERMBG_${today}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/**
 * Copy-ready Google Apps Script code for Admin's personal Google Spreadsheet.
 * Paste into Extensions -> Apps Script -> Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone).
 */
export const GOOGLE_APPS_SCRIPT_TEMPLATE = `function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Update Ringkasan Jumlah Pengguna
    var summarySheet = ss.getSheetByName("Ringkasan_Pengguna") || ss.insertSheet("Ringkasan_Pengguna");
    summarySheet.clear();
    summarySheet.appendRow(["Metrik", "Jumlah", "Sinkron Terakhir"]);
    var s = payload.summary || {};
    var ts = payload.syncedAt || new Date().toISOString();
    summarySheet.appendRow(["Total Pengguna Terdeteksi", s.total || 0, ts]);
    summarySheet.appendRow(["Daftar Tunggu (Perlu ACC)", s.menunggu || 0, ts]);
    summarySheet.appendRow(["Pengguna Aktif (ACC)", s.aktif || 0, ts]);
    summarySheet.appendRow(["Kedaluwarsa", s.kedaluwarsa || 0, ts]);
    summarySheet.appendRow(["Diblokir", s.diblokir || 0, ts]);

    // 2. Update Database Lengkap Pengguna
    var dbSheet = ss.getSheetByName("Database_Pengguna") || ss.insertSheet("Database_Pengguna");
    dbSheet.clear();
    dbSheet.appendRow([
      "Email", "Nama Lengkap", "Profesi", "Nama SPPG", "No HP", "Peran",
      "Status ACC", "Tanggal Mulai Login", "Tanggal Berakhir Login",
      "Keterangan", "Login Pertama", "Login Terakhir", "Jumlah Login", "Sumber"
    ]);
    var users = payload.users || [];
    if (users.length > 0) {
      var rows = users.map(function(u) {
        return [
          u.email || "", u.namaLengkap || "", u.profesi || "", u.namaSPPG || "",
          u.noHp || "", u.peran || "USER", u.statusLabel || u.statusPersetujuan || "",
          u.mulaiPada || "", u.berakhirPada || "Tanpa Batas", u.keterangan || "",
          u.loginPertama || "", u.loginTerakhir || "", u.loginCount || 1, u.sumber || ""
        ];
      });
      dbSheet.getRange(2, 1, rows.length, 14).setValues(rows);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, count: users.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;
