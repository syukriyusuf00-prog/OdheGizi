import "dotenv/config";
import express from "express";
import nodemailer from "nodemailer";
import crypto from "node:crypto";
import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 3000);
const OTP_TTL_MS = 10 * 60 * 1000; // 10 menit
const OTP_RESEND_COOLDOWN_MS = 15 * 1000; // 15 detik
const OTP_MAX_ATTEMPTS = 5;

const OTP_HMAC_SECRET =
  process.env.OTP_SECRET ||
  process.env.SMTP_PASS ||
  process.env.RESEND_API_KEY ||
  "sispermbg_gizisync_otp_signing_secret_v1";

interface OtpChallengeRecord {
  email: string;
  uid: string;
  nonce: string;
  codeHash: string;
  expiresAt: number;
  attempts: number;
  used: boolean;
  lastSentAt: number;
}

const otpChallenges = new Map<string, OtpChallengeRecord>();
const usedNonces = new Set<string>();

const normalizeEmail = (email: unknown): string => {
  if (typeof email !== "string") return "";
  return email.trim().toLowerCase();
};

const isValidEmail = (email: string): boolean => {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) && email.length <= 256;
};

const hashOtpCode = (email: string, code: string, nonce: string): string => {
  return crypto
    .createHmac("sha256", OTP_HMAC_SECRET)
    .update(`${email}|${code}|${nonce}`)
    .digest("hex");
};

const signVerificationToken = (payload: {
  email: string;
  uid: string;
  nonce: string;
  codeHash: string;
  expiresAt: number;
}): string => {
  const encodedPayload = Buffer.from(JSON.stringify(payload), "utf8").toString("base64url");
  const signature = crypto
    .createHmac("sha256", OTP_HMAC_SECRET)
    .update(encodedPayload)
    .digest("base64url");
  return `${encodedPayload}.${signature}`;
};

const verifyAndDecodeToken = (
  token: string
): {
  email: string;
  uid: string;
  nonce: string;
  codeHash: string;
  expiresAt: number;
} | null => {
  if (!token || typeof token !== "string" || !token.includes(".")) return null;
  const parts = token.split(".");
  if (parts.length !== 2) return null;
  const [encodedPayload, signature] = parts;

  const expectedSig = crypto
    .createHmac("sha256", OTP_HMAC_SECRET)
    .update(encodedPayload)
    .digest("base64url");

  const sigBuf = Buffer.from(signature, "utf8");
  const expectedBuf = Buffer.from(expectedSig, "utf8");
  if (sigBuf.length !== expectedBuf.length || !crypto.timingSafeEqual(sigBuf, expectedBuf)) {
    return null;
  }

  try {
    const parsed = JSON.parse(Buffer.from(encodedPayload, "base64url").toString("utf8"));
    if (
      typeof parsed?.email === "string" &&
      typeof parsed?.nonce === "string" &&
      typeof parsed?.codeHash === "string" &&
      typeof parsed?.expiresAt === "number"
    ) {
      return parsed;
    }
  } catch {
    return null;
  }
  return null;
};

const extractClientIp = (req: express.Request): string => {
  const forwarded = req.headers["x-forwarded-for"];
  if (typeof forwarded === "string" && forwarded.trim()) {
    const first = forwarded.split(",")[0].trim();
    return first.replace(/^::ffff:/i, "");
  }
  if (Array.isArray(forwarded) && forwarded.length > 0) {
    return forwarded[0].trim().replace(/^::ffff:/i, "");
  }
  const realIp = req.headers["x-real-ip"];
  if (typeof realIp === "string" && realIp.trim()) {
    return realIp.trim().replace(/^::ffff:/i, "");
  }
  const remote = req.socket?.remoteAddress || req.ip || "";
  const clean = remote.replace(/^::ffff:/i, "");
  if (clean === "::1" || clean === "127.0.0.1") return "127.0.0.1 (Lokal)";
  return clean || "Terdeteksi Otomatis";
};

const buildVerificationEmailHtml = (params: {
  namaLengkap: string;
  namaSPPG: string;
  email: string;
  code: string;
}) => {
  const safeName = (params.namaLengkap || "Pengguna SISPERMBG").replace(/[<>]/g, "");
  const safeSppg = (params.namaSPPG || "Satuan Pelayanan Pemenuhan Gizi").replace(/[<>]/g, "");
  return `
    <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden;">
      <div style="background: linear-gradient(135deg, #112461, #1e3a8a); padding: 24px 28px; color: #ffffff; text-align: center;">
        <div style="font-size: 12px; font-weight: 700; letter-spacing: 2px; color: #34d399; text-transform: uppercase;">GIZISYNC · SISPERMBG</div>
        <h2 style="margin: 8px 0 0; font-size: 20px; font-weight: 800;">Kode Verifikasi Pendaftaran Akun</h2>
      </div>
      <div style="padding: 28px; color: #1e293b; font-size: 14px; line-height: 1.6;">
        <p style="margin: 0 0 12px;">Halo <strong>${safeName}</strong>,</p>
        <p style="margin: 0 0 18px;">
          Terima kasih telah mendaftar pada sistem <strong>SISPERMBG (${safeSppg})</strong>.
          Masukkan kode verifikasi 6 digit berikut pada halaman pendaftaran untuk mengaktifkan akun Anda secara langsung:
        </p>
        <div style="background: #f1f5f9; border: 2px dashed #4f46e5; border-radius: 12px; padding: 18px; text-align: center; margin: 20px 0;">
          <div style="font-size: 11px; text-transform: uppercase; letter-spacing: 1.5px; color: #64748b; font-weight: 700; margin-bottom: 6px;">Kode Verifikasi Anda</div>
          <div style="font-size: 34px; font-weight: 900; letter-spacing: 10px; color: #112461; font-family: 'Courier New', monospace;">${params.code}</div>
        </div>
        <p style="margin: 0 0 10px; font-size: 13px; color: #475569;">
          Kode ini berlaku selama <strong>10 menit</strong>. Setelah berhasil divalidasi, akun Anda otomatis aktif dan terhubung di Panel Admin.
        </p>
        <p style="margin: 0; font-size: 12px; color: #94a3b8;">
          Jika Anda tidak merasa melakukan pendaftaran menggunakan email ${params.email}, abaikan pesan ini.
        </p>
      </div>
      <div style="background: #f8fafc; padding: 14px 28px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0;">
        © 2026 GiziSync · Sistem Perencanaan & Pemantauan Gizi Nasional (SISPERMBG)
      </div>
    </div>
  `;
};

const sendVerificationEmail = async (params: {
  email: string;
  namaLengkap: string;
  namaSPPG: string;
  code: string;
}): Promise<{
  delivered: boolean;
  provider: "resend" | "smtp" | "preview";
  errorDetail?: string;
}> => {
  const subject = `[SISPERMBG] Kode Verifikasi Pendaftaran Anda: ${params.code}`;
  const html = buildVerificationEmailHtml(params);
  const text = `Halo ${params.namaLengkap || "Pengguna"},\n\nKode verifikasi pendaftaran SISPERMBG (${params.namaSPPG || "SPPG"}) Anda adalah: ${params.code}\n\nKode berlaku selama 10 menit. Masukkan kode tersebut di aplikasi untuk mengaktifkan akun Anda secara otomatis.`;

  // 1. Try Resend HTTPS API if configured
  const resendApiKey = process.env.RESEND_API_KEY?.trim();
  if (resendApiKey && resendApiKey !== "MY_RESEND_API_KEY") {
    try {
      const fromAddress = process.env.SMTP_FROM?.trim() || "SISPERMBG GiziSync <onboarding@resend.dev>";
      const resp = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${resendApiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          from: fromAddress,
          to: [params.email],
          subject,
          html,
          text,
        }),
      });
      if (resp.ok) {
        return { delivered: true, provider: "resend" };
      }
      const errText = await resp.text();
      console.warn("[OTP Email] Resend API error:", errText);
    } catch (err) {
      console.warn("[OTP Email] Resend fetch error:", err);
    }
  }

  // 2. Try Nodemailer SMTP if SMTP credentials are configured
  const smtpUser = process.env.SMTP_USER?.trim();
  const smtpPass = process.env.SMTP_PASS?.trim();
  if (smtpUser && smtpPass && smtpUser !== "MY_SMTP_USER" && smtpPass !== "MY_SMTP_PASS") {
    try {
      const smtpHost = process.env.SMTP_HOST?.trim() || "smtp.gmail.com";
      const smtpPort = Number(process.env.SMTP_PORT || 465);
      const secure = process.env.SMTP_SECURE
        ? process.env.SMTP_SECURE === "true"
        : smtpPort === 465;

      const transporter = nodemailer.createTransport({
        host: smtpHost,
        port: smtpPort,
        secure,
        auth: {
          user: smtpUser,
          pass: smtpPass,
        },
      });

      const fromAddress = process.env.SMTP_FROM?.trim() || `SISPERMBG GiziSync <${smtpUser}>`;
      await transporter.sendMail({
        from: fromAddress,
        to: params.email,
        subject,
        text,
        html,
      });

      return { delivered: true, provider: "smtp" };
    } catch (err: any) {
      console.warn("[OTP Email] SMTP error:", err?.message || err);
      return {
        delivered: false,
        provider: "preview",
        errorDetail: err?.message || "Gagal terhubung ke server SMTP.",
      };
    }
  }

  // 3. Fallback preview relay when no SMTP/Resend secrets are set in runtime environment
  return { delivered: false, provider: "preview" };
};

interface ServerSheetSyncConfig {
  sheetUrl: string;
  spreadsheetId: string;
  gid: string;
  sheetTitle: string;
  spreadsheetTitle: string;
  updatedAt: string;
  updatedBy?: string;
}

interface ServerSheetUserRow {
  rowNumber: number;
  email: string;
  namaLengkap: string;
  profesi: string;
  namaSPPG: string;
  noHp: string;
  peran: "USER" | "ADMIN";
  statusPersetujuan: "aktif" | "menunggu" | "diblokir";
  berakhirPada: string;
  loginTerakhir: string;
  sumber: string;
}

const REALTIME_CACHE_FILE = path.join("/tmp", "sispermbg_realtime_hub_v1.json");
const realtimeUsers = new Map<string, Record<string, any>>();
const realtimeDevices = new Map<string, Record<string, any>>();
const cachedSheetUsers = new Map<string, ServerSheetUserRow>();
let sheetSyncConfig: ServerSheetSyncConfig | null = null;
let lastPublicSheetFetchAt = 0;
const sseClients = new Set<express.Response>();
let persistTimer: ReturnType<typeof setTimeout> | null = null;

const loadRealtimeCacheFromDisk = () => {
  try {
    if (!fs.existsSync(REALTIME_CACHE_FILE)) return;
    const raw = fs.readFileSync(REALTIME_CACHE_FILE, "utf8");
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed?.users)) {
      parsed.users.forEach((u: any) => {
        if (u && typeof u.uid === "string" && u.uid) {
          realtimeUsers.set(u.uid, u);
        }
      });
    }
    if (Array.isArray(parsed?.devices)) {
      parsed.devices.forEach((d: any) => {
        if (d && typeof d.deviceId === "string" && d.deviceId) {
          realtimeDevices.set(d.deviceId, d);
        }
      });
    }
    if (parsed?.sheetSyncConfig && typeof parsed.sheetSyncConfig.spreadsheetId === "string") {
      sheetSyncConfig = parsed.sheetSyncConfig;
    }
    if (Array.isArray(parsed?.cachedSheetUsers)) {
      parsed.cachedSheetUsers.forEach((su: any) => {
        const em = normalizeEmail(su?.email);
        if (em && isValidEmail(em)) {
          cachedSheetUsers.set(em, { ...su, email: em });
        }
      });
    }
  } catch (err) {
    console.warn("[RealtimeHub] Notice loading cache:", err);
  }
};

const schedulePersistRealtimeCache = () => {
  if (persistTimer) return;
  persistTimer = setTimeout(() => {
    persistTimer = null;
    try {
      const payload = {
        updatedAt: new Date().toISOString(),
        users: Array.from(realtimeUsers.values()).slice(0, 500),
        devices: Array.from(realtimeDevices.values()).slice(0, 500),
        sheetSyncConfig,
        cachedSheetUsers: Array.from(cachedSheetUsers.values()).slice(0, 1000),
      };
      fs.writeFileSync(REALTIME_CACHE_FILE, JSON.stringify(payload), "utf8");
    } catch (err) {
      console.warn("[RealtimeHub] Notice writing cache:", err);
    }
  }, 300);
};

const parseCsvRows = (csvText: string): string[][] => {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < csvText.length; i++) {
    const ch = csvText[i];
    if (inQuotes) {
      if (ch === '"') {
        if (csvText[i + 1] === '"') {
          cell += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cell += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\r") {
      if (csvText[i + 1] === "\n") i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += ch;
    }
  }
  if (cell.length > 0 || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
};

const normalizeServerSheetStatus = (raw: unknown): "aktif" | "menunggu" | "diblokir" => {
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
  return "aktif";
};

const getDefaultSheetExpDate = (): string => {
  const d = new Date();
  d.setFullYear(d.getFullYear() + 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};

const parseSheetRowsOnServer = (rows: string[][]): ServerSheetUserRow[] => {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const firstRow = rows[0].map((c) => String(c ?? "").trim().toLowerCase());
  const firstRowHasEmail = firstRow.some((c) => isValidEmail(c));

  const findHeader = (patterns: RegExp[]): number => {
    for (let i = 0; i < firstRow.length; i++) {
      if (firstRow[i] && patterns.some((rx) => rx.test(firstRow[i]))) return i;
    }
    return -1;
  };

  let hasHeader = false;
  let emailIdx = 0;
  let namaIdx = 1;
  let profesiIdx = 2;
  let sppgIdx = 3;
  let hpIdx = 4;
  let peranIdx = 5;
  let statusIdx = 6;
  let berakhirIdx = 7;
  let loginIdx = 8;
  let sumberIdx = 9;

  if (!firstRowHasEmail) {
    const hEmail = findHeader([/\bemail\b/, /e-mail/, /gmail/, /alamat\s*email/, /\bakun\b/]);
    const hNama = findHeader([/nama\s*lengkap/, /^nama$/, /nama\s*pengguna/, /full\s*name/, /^name$/]);
    const hStatus = findHeader([/status/, /\bacc\b/, /persetujuan/, /approval/]);
    if (hEmail >= 0 || hNama >= 0 || hStatus >= 0) {
      hasHeader = true;
      emailIdx = hEmail >= 0 ? hEmail : 0;
      namaIdx = hNama >= 0 ? hNama : emailIdx === 0 ? 1 : 0;
      profesiIdx = findHeader([/profesi/, /jabatan/, /pekerjaan/]);
      sppgIdx = findHeader([/sppg/, /lembaga/, /instansi/, /dapur/, /unit/]);
      hpIdx = findHeader([/no\.?\s*hp/, /^hp$/, /whatsapp/, /\bwa\b/, /telepon/, /phone/, /kontak/]);
      peranIdx = findHeader([/^peran$/, /^role$/, /hak\s*akses/, /level/]);
      statusIdx = hStatus;
      berakhirIdx = findHeader([/berakhir/, /masa\s*aktif/, /expir/, /kadaluarsa/, /berlaku/]);
      loginIdx = findHeader([/login\s*terakhir/, /last\s*login/, /terakhir\s*masuk/]);
      sumberIdx = findHeader([/sumber/, /source/]);
    }
  }

  if (!hasHeader) {
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
      const idx = (rows[r] || []).findIndex((c) => isValidEmail(normalizeEmail(c)));
      if (idx >= 0) {
        emailIdx = idx;
        namaIdx = idx === 0 ? 1 : 0;
        break;
      }
    }
  }

  const startIdx = hasHeader ? 1 : 0;
  const byEmail = new Map<string, ServerSheetUserRow>();

  for (let i = startIdx; i < rows.length; i++) {
    const row = rows[i] || [];
    let email = normalizeEmail(row[emailIdx]);
    if (!isValidEmail(email)) {
      const fallbackCell = row.find((c) => isValidEmail(normalizeEmail(c)));
      email = normalizeEmail(fallbackCell);
    }
    if (!isValidEmail(email)) continue;

    const getVal = (colIdx: number) => (colIdx >= 0 && colIdx < row.length ? String(row[colIdx] ?? "").trim() : "");
    const rawName = getVal(namaIdx);
    const safeName =
      rawName && normalizeEmail(rawName) !== email ? rawName : email.split("@")[0] || "Pengguna SISPERMBG";
    const rawRole = getVal(peranIdx).toUpperCase();

    byEmail.set(email, {
      rowNumber: i + 1,
      email,
      namaLengkap: safeName.slice(0, 160),
      profesi: (getVal(profesiIdx) || "Ahli Gizi / Pengelola SPPG").slice(0, 120),
      namaSPPG: getVal(sppgIdx).slice(0, 160),
      noHp: getVal(hpIdx).slice(0, 40),
      peran: rawRole === "ADMIN" || rawRole === "ADMINISTRATOR" ? "ADMIN" : "USER",
      statusPersetujuan: normalizeServerSheetStatus(getVal(statusIdx)),
      berakhirPada: (getVal(berakhirIdx) || getDefaultSheetExpDate()).slice(0, 32),
      loginTerakhir: getVal(loginIdx).slice(0, 64),
      sumber: (getVal(sumberIdx) || "Spreadsheet").slice(0, 64),
    });
  }

  return Array.from(byEmail.values());
};

const mergeSheetUsersIntoRealtimeHub = (sheetUsers: ServerSheetUserRow[]) => {
  const nowIso = new Date().toISOString();
  for (const su of sheetUsers) {
    const email = normalizeEmail(su.email);
    if (!isValidEmail(email)) continue;
    cachedSheetUsers.set(email, su);

    let existingKey = "";
    let existingVal: Record<string, any> | null = null;
    for (const [k, u] of realtimeUsers.entries()) {
      if (normalizeEmail(u.email) === email) {
        existingKey = k;
        existingVal = u;
        break;
      }
    }

    if (existingVal && existingKey) {
      // Preserve real Firebase UID; promote menunggu -> aktif if spreadsheet explicitly grants aktif
      const mergedUser: Record<string, any> = {
        ...existingVal,
        namaLengkap: existingVal.namaLengkap || su.namaLengkap,
        profesi: existingVal.profesi || su.profesi,
        namaSPPG: existingVal.namaSPPG || su.namaSPPG,
        noHp: existingVal.noHp || su.noHp,
        berakhirPada: existingVal.berakhirPada || su.berakhirPada,
        peran: existingVal.peran === "ADMIN" ? "ADMIN" : su.peran,
        statusPersetujuan:
          existingVal.statusPersetujuan === "menunggu" && su.statusPersetujuan === "aktif"
            ? "aktif"
            : existingVal.statusPersetujuan || su.statusPersetujuan,
        registeredViaSpreadsheet: true,
        sumber: existingVal.sumber || su.sumber || "Spreadsheet",
        updatedAt: nowIso,
      };
      realtimeUsers.set(existingKey, mergedUser);
      broadcastRealtimeEvent("user_upsert", mergedUser);
    } else {
      const sheetUid = `sheet_user_${email.replace(/[^a-z0-9]/g, "_")}`;
      const createdUser: Record<string, any> = {
        uid: sheetUid,
        id: sheetUid,
        email,
        namaLengkap: su.namaLengkap,
        profesi: su.profesi,
        namaSPPG: su.namaSPPG,
        noHp: su.noHp,
        peran: su.peran,
        statusPersetujuan: su.statusPersetujuan,
        berakhirPada: su.berakhirPada,
        authProvider: "google",
        emailVerified: true,
        registeredViaSpreadsheet: true,
        sumber: su.sumber || "Spreadsheet",
        createdAt: nowIso,
        updatedAt: nowIso,
        loginTerakhir: su.loginTerakhir || null,
      };
      realtimeUsers.set(sheetUid, createdUser);
      broadcastRealtimeEvent("user_upsert", createdUser);
    }
  }
  schedulePersistRealtimeCache();
};

const fetchPublicGoogleSheetUsers = async (params: {
  spreadsheetId: string;
  gid?: string;
  sheetTitle?: string;
}): Promise<{
  ok: boolean;
  users: ServerSheetUserRow[];
  requiresOAuth?: boolean;
  error?: string;
}> => {
  const spreadsheetId = clampField(params.spreadsheetId, 128);
  if (!spreadsheetId || !/^[a-zA-Z0-9_-]+$/.test(spreadsheetId)) {
    return { ok: false, users: [], error: "ID Spreadsheet tidak valid." };
  }

  const query = new URLSearchParams({ tqx: "out:csv" });
  if (params.sheetTitle) {
    query.set("sheet", params.sheetTitle);
  } else if (params.gid) {
    query.set("gid", params.gid);
  }

  const csvUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
    spreadsheetId
  )}/gviz/tq?${query.toString()}`;

  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 7000);
    const res = await fetch(csvUrl, {
      method: "GET",
      redirect: "follow",
      signal: controller.signal,
    });
    clearTimeout(timer);

    const contentType = (res.headers.get("content-type") || "").toLowerCase();
    const text = await res.text();

    // Private sheets redirect to Google login HTML page instead of returning CSV
    if (
      !res.ok ||
      contentType.includes("text/html") ||
      text.trim().startsWith("<!DOCTYPE") ||
      text.trim().startsWith("<html")
    ) {
      return {
        ok: false,
        users: [],
        requiresOAuth: true,
        error:
          "Spreadsheet bersifat privat. Hubungkan Google Sheets di Panel Admin atau ubah izin berbagi spreadsheet menjadi 'Siapa saja yang memiliki link dapat melihat'.",
      };
    }

    const parsedRows = parseCsvRows(text);
    const users = parseSheetRowsOnServer(parsedRows);
    lastPublicSheetFetchAt = Date.now();
    if (users.length > 0) {
      mergeSheetUsersIntoRealtimeHub(users);
    }
    return { ok: true, users };
  } catch (err: any) {
    return {
      ok: false,
      users: [],
      error: err?.message || "Gagal mengambil data spreadsheet.",
    };
  }
};

const broadcastRealtimeEvent = (eventType: string, payload: unknown) => {
  if (sseClients.size === 0) return;
  const serialized = `event: ${eventType}\ndata: ${JSON.stringify(payload)}\n\n`;
  for (const client of sseClients) {
    try {
      client.write(serialized);
    } catch {
      sseClients.delete(client);
    }
  }
};

const parseJsonBodySafe = (body: unknown): Record<string, any> => {
  if (body && typeof body === "object") return body as Record<string, any>;
  if (typeof body === "string" && body.trim()) {
    try {
      const parsed = JSON.parse(body);
      if (parsed && typeof parsed === "object") return parsed as Record<string, any>;
    } catch {}
  }
  return {};
};

const clampField = (val: unknown, maxLen: number): string => {
  if (typeof val !== "string") return "";
  return val.trim().slice(0, maxLen);
};

const sanitizeRealtimeUser = (raw: Record<string, any>): Record<string, any> | null => {
  const email = normalizeEmail(raw.email);
  const uid = clampField(raw.uid || raw.id || (email ? `custom_user_${email.replace(/[@.]/g, "_")}` : ""), 128);
  if (!uid || !email) return null;
  const nowIso = new Date().toISOString();
  let prev = realtimeUsers.get(uid) || {};
  if (!realtimeUsers.has(uid)) {
    for (const existingUser of realtimeUsers.values()) {
      if (normalizeEmail(existingUser.email) === email) {
        prev = existingUser;
        break;
      }
    }
  }
  const merged: Record<string, any> = {
    ...prev,
    uid,
    id: uid,
    email,
    namaLengkap: clampField(raw.namaLengkap ?? prev.namaLengkap ?? "", 160),
    profesi: clampField(raw.profesi ?? prev.profesi ?? "", 120),
    namaSPPG: clampField(raw.namaSPPG ?? prev.namaSPPG ?? "", 160),
    noHp: clampField(raw.noHp ?? prev.noHp ?? "", 40),
    peran: clampField(raw.peran ?? prev.peran ?? "USER", 32) || "USER",
    statusPersetujuan: clampField(raw.statusPersetujuan ?? prev.statusPersetujuan ?? "menunggu", 32) || "menunggu",
    berakhirPada: clampField(raw.berakhirPada ?? prev.berakhirPada ?? "", 32),
    authProvider: clampField(raw.authProvider ?? prev.authProvider ?? "", 32) || undefined,
    sumber: clampField(raw.sumber ?? prev.sumber ?? "", 64) || undefined,
    registeredViaSpreadsheet:
      typeof raw.registeredViaSpreadsheet === "boolean"
        ? raw.registeredViaSpreadsheet
        : Boolean(prev.registeredViaSpreadsheet),
    emailVerified: typeof raw.emailVerified === "boolean" ? raw.emailVerified : Boolean(prev.emailVerified),
    verificationStatus: clampField(raw.verificationStatus ?? prev.verificationStatus ?? "", 32) || undefined,
    emailVerifiedAt: clampField(raw.emailVerifiedAt ?? prev.emailVerifiedAt ?? "", 64) || undefined,
    activatedBy: clampField(raw.activatedBy ?? prev.activatedBy ?? "", 40) || undefined,
    lastDeviceId: clampField(raw.lastDeviceId ?? prev.lastDeviceId ?? "", 64) || undefined,
    lastDeviceLabel: clampField(raw.lastDeviceLabel ?? prev.lastDeviceLabel ?? "", 160) || undefined,
    lastDeviceSeenAt: clampField(raw.lastDeviceSeenAt ?? prev.lastDeviceSeenAt ?? "", 64) || undefined,
    lastDevice: raw.lastDevice && typeof raw.lastDevice === "object" ? raw.lastDevice : prev.lastDevice,
    createdAt: typeof raw.createdAt === "string" && raw.createdAt ? raw.createdAt : prev.createdAt || nowIso,
    updatedAt: nowIso,
    loginTerakhir:
      raw.loginTerakhir === null
        ? prev.loginTerakhir ?? null
        : typeof raw.loginTerakhir === "string" && raw.loginTerakhir
        ? raw.loginTerakhir
        : prev.loginTerakhir ?? null,
  };
  if (typeof raw.sandi === "string" && raw.sandi.trim()) {
    merged.sandi = raw.sandi.trim().slice(0, 128);
  }
  return merged;
};

const sanitizeRealtimeDevice = (raw: Record<string, any>, observedIp: string): Record<string, any> | null => {
  const deviceId = clampField(raw.deviceId, 64);
  if (!deviceId || !/^[a-zA-Z0-9_-]+$/.test(deviceId)) return null;
  const prev = realtimeDevices.get(deviceId) || {};
  const nowIso = new Date().toISOString();
  const validTypes = new Set(["Mobile", "Tablet", "Desktop / Laptop"]);
  const validStates = new Set(["guest", "registering", "active", "demo"]);
  const deviceType = validTypes.has(raw.deviceType) ? raw.deviceType : (prev.deviceType || "Desktop / Laptop");
  const sessionState = validStates.has(raw.sessionState) ? raw.sessionState : (prev.sessionState || "guest");

  const resolvedIp =
    observedIp && observedIp !== "Terdeteksi Otomatis"
      ? observedIp
      : clampField(raw.ipAddress ?? prev.ipAddress ?? "Terdeteksi", 64);

  return {
    ...prev,
    deviceId,
    deviceType,
    deviceModel: clampField(raw.deviceModel ?? prev.deviceModel ?? deviceType, 100),
    osName: clampField(raw.osName ?? prev.osName ?? "Tidak Diketahui", 80),
    browserName: clampField(raw.browserName ?? prev.browserName ?? "Browser Web", 80),
    screenResolution: clampField(raw.screenResolution ?? prev.screenResolution ?? "-", 40),
    viewportSize: clampField(raw.viewportSize ?? prev.viewportSize ?? "-", 40),
    language: clampField(raw.language ?? prev.language ?? "id-ID", 32),
    timezone: clampField(raw.timezone ?? prev.timezone ?? "Asia/Jakarta", 80),
    ipAddress: resolvedIp,
    userAgent: clampField(raw.userAgent ?? prev.userAgent ?? "", 320),
    sessionState,
    userUid: raw.userUid === null ? null : clampField(raw.userUid ?? prev.userUid ?? "", 128) || null,
    userEmail:
      raw.userEmail === null
        ? null
        : normalizeEmail(raw.userEmail ?? prev.userEmail ?? "") || null,
    userName: raw.userName === null ? null : clampField(raw.userName ?? prev.userName ?? "", 160) || null,
    namaSPPG: raw.namaSPPG === null ? null : clampField(raw.namaSPPG ?? prev.namaSPPG ?? "", 160) || null,
    peran: raw.peran === null ? null : clampField(raw.peran ?? prev.peran ?? "", 32) || null,
    statusPersetujuan:
      raw.statusPersetujuan === null
        ? null
        : clampField(raw.statusPersetujuan ?? prev.statusPersetujuan ?? "", 32) || null,
    isOnline: typeof raw.isOnline === "boolean" ? raw.isOnline : true,
    firstSeenAt: clampField(raw.firstSeenAt ?? prev.firstSeenAt ?? nowIso, 64) || nowIso,
    lastSeenAt: nowIso,
  };
};

async function startServer() {
  loadRealtimeCacheFromDisk();
  const app = express();
  app.set("trust proxy", true);
  app.use(express.json({ limit: "2mb" }));
  app.use(express.text({ type: "text/plain", limit: "1mb" }));

  // Periodic cleanup of expired OTP challenges
  setInterval(() => {
    const now = Date.now();
    for (const [key, rec] of otpChallenges.entries()) {
      if (now > rec.expiresAt + 60_000) {
        otpChallenges.delete(key);
      }
    }
    if (usedNonces.size > 5000) {
      usedNonces.clear();
    }
  }, 5 * 60 * 1000).unref();

  // 1. Device Telemetry Enrichment Endpoint (returns client IP & server timestamp)
  app.get("/api/device/info", (req, res) => {
    res.json({
      ipAddress: extractClientIp(req),
      userAgent: typeof req.headers["user-agent"] === "string" ? req.headers["user-agent"].slice(0, 280) : "",
      serverTime: new Date().toISOString(),
    });
  });

  // 1b. Cross-Device Real-Time Sync Hub (SSE + Snapshot + Upsert/Delete)
  // Ensures Mobile -> Laptop Admin Panel updates arrive in <1s even if Firestore free daily quota hits limit
  app.get("/api/realtime/stream", (req, res) => {
    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.setHeader("X-Accel-Buffering", "no");
    if (typeof res.flushHeaders === "function") {
      res.flushHeaders();
    }

    sseClients.add(res);
    const initialSnapshot = {
      users: Array.from(realtimeUsers.values()),
      devices: Array.from(realtimeDevices.values()),
      serverTime: new Date().toISOString(),
    };
    res.write(`event: snapshot\ndata: ${JSON.stringify(initialSnapshot)}\n\n`);

    const keepAlive = setInterval(() => {
      try {
        res.write(`: keepalive ${Date.now()}\n\n`);
      } catch {
        clearInterval(keepAlive);
        sseClients.delete(res);
      }
    }, 20_000);

    req.on("close", () => {
      clearInterval(keepAlive);
      sseClients.delete(res);
    });
  });

  app.get("/api/realtime/snapshot", (_req, res) => {
    res.setHeader("Cache-Control", "no-store");
    res.json({
      ok: true,
      users: Array.from(realtimeUsers.values()),
      devices: Array.from(realtimeDevices.values()),
      connectedListeners: sseClients.size,
      serverTime: new Date().toISOString(),
    });
  });

  app.get("/api/realtime/user", (req, res) => {
    res.setHeader("Cache-Control", "no-store");
    const email = normalizeEmail(req.query.email);
    const uid = clampField(req.query.uid, 128);
    let found: Record<string, any> | null = null;
    if (uid && realtimeUsers.has(uid)) {
      found = realtimeUsers.get(uid) || null;
    }
    if (!found && email) {
      for (const u of realtimeUsers.values()) {
        if (normalizeEmail(u.email) === email) {
          found = u;
          break;
        }
      }
    }
    res.json({ ok: true, user: found });
  });

  app.post("/api/realtime/sync-user", (req, res) => {
    const body = parseJsonBodySafe(req.body);
    const sanitized = sanitizeRealtimeUser(body);
    if (!sanitized) {
      res.status(400).json({ ok: false, error: "Data pengguna tidak valid." });
      return;
    }
    // Remove duplicate keys with same email so a Google Firebase UID replaces legacy custom_user_* or sheet_user_* cleanly
    const targetEmail = normalizeEmail(sanitized.email);
    const isRealFirebaseUid =
      !sanitized.uid.startsWith("custom_user_") && !sanitized.uid.startsWith("sheet_user_");
    for (const [existingKey, existingVal] of realtimeUsers.entries()) {
      if (existingKey !== sanitized.uid && normalizeEmail(existingVal.email) === targetEmail) {
        if (
          isRealFirebaseUid ||
          existingKey.startsWith("custom_user_") ||
          existingKey.startsWith("sheet_user_")
        ) {
          realtimeUsers.delete(existingKey);
        }
      }
    }
    realtimeUsers.set(sanitized.uid, sanitized);
    schedulePersistRealtimeCache();
    broadcastRealtimeEvent("user_upsert", sanitized);
    res.json({ ok: true, user: sanitized, serverTime: new Date().toISOString() });
  });

  app.post("/api/realtime/delete-user", (req, res) => {
    const body = parseJsonBodySafe(req.body);
    const uid = clampField(body.uid, 128);
    const email = normalizeEmail(body.email);
    const removedUids: string[] = [];
    for (const [k, u] of realtimeUsers.entries()) {
      if ((uid && (k === uid || u.uid === uid)) || (email && normalizeEmail(u.email) === email)) {
        realtimeUsers.delete(k);
        removedUids.push(k);
      }
    }
    if (email) {
      cachedSheetUsers.delete(email);
    }
    schedulePersistRealtimeCache();
    broadcastRealtimeEvent("user_delete", { uid, email, removedUids });
    res.json({ ok: true, removedUids });
  });

  app.post("/api/realtime/sync-device", (req, res) => {
    const body = parseJsonBodySafe(req.body);
    const observedIp = extractClientIp(req);
    const sanitized = sanitizeRealtimeDevice(body, observedIp);
    if (!sanitized) {
      res.status(400).json({ ok: false, error: "Data perangkat tidak valid." });
      return;
    }
    realtimeDevices.set(sanitized.deviceId, sanitized);
    schedulePersistRealtimeCache();
    broadcastRealtimeEvent("device_upsert", sanitized);
    res.json({ ok: true, device: sanitized, serverTime: new Date().toISOString() });
  });

  app.post("/api/realtime/delete-device", (req, res) => {
    const body = parseJsonBodySafe(req.body);
    const deviceId = clampField(body.deviceId, 64);
    if (deviceId) {
      realtimeDevices.delete(deviceId);
      schedulePersistRealtimeCache();
      broadcastRealtimeEvent("device_delete", { deviceId });
    }
    res.json({ ok: true, deviceId });
  });

  // 1c. Google Sheets Bidirectional User Sync Config & Live Lookup Routes
  app.get("/api/realtime/sheet-config", (_req, res) => {
    res.setHeader("Cache-Control", "no-store");
    res.json({
      ok: true,
      config: sheetSyncConfig,
      cachedSheetUsersCount: cachedSheetUsers.size,
      serverTime: new Date().toISOString(),
    });
  });

  app.post("/api/realtime/sheet-config", (req, res) => {
    const body = parseJsonBodySafe(req.body);
    const spreadsheetId = clampField(body.spreadsheetId, 128);
    if (!spreadsheetId || !/^[a-zA-Z0-9_-]+$/.test(spreadsheetId)) {
      res.status(400).json({ ok: false, error: "ID Google Spreadsheet tidak valid." });
      return;
    }
    sheetSyncConfig = {
      sheetUrl: clampField(body.sheetUrl, 512) || `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
      spreadsheetId,
      gid: clampField(body.gid, 40),
      sheetTitle: clampField(body.sheetTitle, 120),
      spreadsheetTitle: clampField(body.spreadsheetTitle, 200),
      updatedAt: new Date().toISOString(),
      updatedBy: clampField(body.updatedBy, 160) || undefined,
    };
    lastPublicSheetFetchAt = 0; // Force fresh lookup on next login
    schedulePersistRealtimeCache();
    broadcastRealtimeEvent("sheet_config_update", sheetSyncConfig);
    res.json({ ok: true, config: sheetSyncConfig });
  });

  app.post("/api/realtime/sheet-users-cache", (req, res) => {
    const body = parseJsonBodySafe(req.body);
    const incomingUsers = Array.isArray(body.users) ? body.users : [];
    const normalized: ServerSheetUserRow[] = [];
    for (const item of incomingUsers) {
      const email = normalizeEmail(item?.email);
      if (!isValidEmail(email)) continue;
      normalized.push({
        rowNumber: Number(item?.rowNumber) || 0,
        email,
        namaLengkap: clampField(item?.namaLengkap, 160) || email.split("@")[0] || "Pengguna SISPERMBG",
        profesi: clampField(item?.profesi, 120) || "Ahli Gizi / Pengelola SPPG",
        namaSPPG: clampField(item?.namaSPPG, 160),
        noHp: clampField(item?.noHp, 40),
        peran: String(item?.peran || "").toUpperCase() === "ADMIN" ? "ADMIN" : "USER",
        statusPersetujuan: normalizeServerSheetStatus(item?.statusPersetujuan),
        berakhirPada: clampField(item?.berakhirPada, 32) || getDefaultSheetExpDate(),
        loginTerakhir: clampField(item?.loginTerakhir, 64),
        sumber: clampField(item?.sumber, 64) || "Spreadsheet",
      });
    }
    if (normalized.length > 0) {
      mergeSheetUsersIntoRealtimeHub(normalized);
    }
    res.json({ ok: true, cachedCount: normalized.length, totalCached: cachedSheetUsers.size });
  });

  app.get("/api/sheets/read-public", async (req, res) => {
    res.setHeader("Cache-Control", "no-store");
    const spreadsheetId = clampField(req.query.spreadsheetId, 128) || sheetSyncConfig?.spreadsheetId || "";
    const gid = clampField(req.query.gid, 40) || sheetSyncConfig?.gid || "";
    const sheetTitle = clampField(req.query.sheetTitle, 120) || sheetSyncConfig?.sheetTitle || "";

    if (!spreadsheetId) {
      res.status(400).json({ ok: false, error: "Link Google Spreadsheet belum dikonfigurasi." });
      return;
    }

    const result = await fetchPublicGoogleSheetUsers({ spreadsheetId, gid, sheetTitle });
    res.json({
      ...result,
      rowCount: result.users.length,
      serverTime: new Date().toISOString(),
    });
  });

  app.get("/api/sheets/lookup-email", async (req, res) => {
    res.setHeader("Cache-Control", "no-store");
    const email = normalizeEmail(req.query.email);
    if (!isValidEmail(email)) {
      res.status(400).json({ ok: false, found: false, error: "Email tidak valid." });
      return;
    }

    // Refresh live from public/shared Google Sheet (throttled to every 15s) so newly added spreadsheet emails take effect immediately
    if (sheetSyncConfig?.spreadsheetId && Date.now() - lastPublicSheetFetchAt > 15_000) {
      await fetchPublicGoogleSheetUsers({
        spreadsheetId: sheetSyncConfig.spreadsheetId,
        gid: sheetSyncConfig.gid,
        sheetTitle: sheetSyncConfig.sheetTitle,
      }).catch(() => {});
    }

    const match = cachedSheetUsers.get(email) || null;
    res.json({
      ok: true,
      found: Boolean(match),
      sheetUser: match,
      hasLinkedSpreadsheet: Boolean(sheetSyncConfig?.spreadsheetId),
    });
  });

  // 2. Send 6-Digit Email Verification Code on Registration
  app.post("/api/auth/send-verification-code", async (req, res) => {
    try {
      const email = normalizeEmail(req.body?.email);
      const namaLengkap = typeof req.body?.namaLengkap === "string" ? req.body.namaLengkap.trim().slice(0, 140) : "";
      const namaSPPG = typeof req.body?.namaSPPG === "string" ? req.body.namaSPPG.trim().slice(0, 140) : "";
      const uid =
        typeof req.body?.uid === "string" && req.body.uid.trim()
          ? req.body.uid.trim().slice(0, 128)
          : `custom_user_${email.replace(/[@.]/g, "_")}`;

      if (!isValidEmail(email)) {
        res.status(400).json({ ok: false, error: "Format alamat email tidak valid." });
        return;
      }

      const existing = otpChallenges.get(email);
      const now = Date.now();
      if (existing && !existing.used && now - existing.lastSentAt < OTP_RESEND_COOLDOWN_MS) {
        const retryAfter = Math.ceil((OTP_RESEND_COOLDOWN_MS - (now - existing.lastSentAt)) / 1000);
        res.status(429).json({
          ok: false,
          error: `Mohon tunggu ${retryAfter} detik sebelum meminta kirim ulang kode verifikasi.`,
          retryAfterSeconds: retryAfter,
        });
        return;
      }

      const code = String(crypto.randomInt(100000, 1000000));
      const nonce = crypto.randomBytes(10).toString("hex");
      const expiresAt = now + OTP_TTL_MS;
      const codeHash = hashOtpCode(email, code, nonce);

      const record: OtpChallengeRecord = {
        email,
        uid,
        nonce,
        codeHash,
        expiresAt,
        attempts: 0,
        used: false,
        lastSentAt: now,
      };
      otpChallenges.set(email, record);

      const verificationToken = signVerificationToken({
        email,
        uid,
        nonce,
        codeHash,
        expiresAt,
      });

      const emailResult = await sendVerificationEmail({
        email,
        namaLengkap,
        namaSPPG,
        code,
      });

      if (emailResult.delivered) {
        res.json({
          ok: true,
          deliveryMode: "email",
          provider: emailResult.provider,
          recipientEmail: email,
          verificationToken,
          expiresAt,
          expiresInSeconds: Math.floor(OTP_TTL_MS / 1000),
          message: `Kode verifikasi 6 digit telah dikirim otomatis ke email ${email}. Silakan cek Kotak Masuk atau folder Spam.`,
        });
        return;
      }

      // When SMTP/Resend secrets are not yet configured in environment, deliver in-app email preview so verification still completes end-to-end
      res.json({
        ok: true,
        deliveryMode: "preview",
        provider: "preview",
        recipientEmail: email,
        verificationToken,
        expiresAt,
        expiresInSeconds: Math.floor(OTP_TTL_MS / 1000),
        emailSubject: `[SISPERMBG] Kode Verifikasi Pendaftaran Anda: ${code}`,
        previewCode: code,
        smtpWarning: emailResult.errorDetail || null,
        message: `Kode verifikasi 6 digit untuk ${email} telah dibuat otomatis. Masukkan kode pada pratinjau email di bawah untuk mengaktifkan akun.`,
      });
    } catch (err: any) {
      console.error("[OTP Send Error]:", err);
      res.status(500).json({
        ok: false,
        error: "Terjadi kesalahan saat mengirim kode verifikasi email. Silakan coba lagi.",
      });
    }
  });

  // 3. Validate 6-Digit Verification Code
  app.post("/api/auth/verify-code", (req, res) => {
    try {
      const email = normalizeEmail(req.body?.email);
      const rawCode = typeof req.body?.code === "string" ? req.body.code.replace(/\D/g, "").trim() : "";
      const verificationToken = typeof req.body?.verificationToken === "string" ? req.body.verificationToken.trim() : "";

      if (!isValidEmail(email)) {
        res.status(400).json({ ok: false, error: "Alamat email tidak valid." });
        return;
      }
      if (!/^\d{6}$/.test(rawCode)) {
        res.status(400).json({ ok: false, error: "Kode verifikasi harus terdiri dari 6 digit angka." });
        return;
      }

      const tokenData = verifyAndDecodeToken(verificationToken);
      const memRecord = otpChallenges.get(email);

      const challenge = tokenData && tokenData.email === email
        ? {
            email: tokenData.email,
            uid: tokenData.uid,
            nonce: tokenData.nonce,
            codeHash: tokenData.codeHash,
            expiresAt: tokenData.expiresAt,
          }
        : memRecord
        ? {
            email: memRecord.email,
            uid: memRecord.uid,
            nonce: memRecord.nonce,
            codeHash: memRecord.codeHash,
            expiresAt: memRecord.expiresAt,
          }
        : null;

      if (!challenge) {
        res.status(400).json({
          ok: false,
          error: "Sesi kode verifikasi tidak ditemukan atau tidak valid. Silakan kirim ulang kode baru.",
        });
        return;
      }

      if (usedNonces.has(challenge.nonce) || (memRecord && memRecord.nonce === challenge.nonce && memRecord.used)) {
        res.status(400).json({
          ok: false,
          error: "Kode verifikasi ini sudah digunakan. Silakan masuk menggunakan akun Anda.",
        });
        return;
      }

      if (Date.now() > challenge.expiresAt) {
        res.status(400).json({
          ok: false,
          error: "Kode verifikasi telah kedaluwarsa (lebih dari 10 menit). Silakan klik Kirim Ulang Kode.",
        });
        return;
      }

      if (memRecord && memRecord.nonce === challenge.nonce && memRecord.attempts >= OTP_MAX_ATTEMPTS) {
        res.status(429).json({
          ok: false,
          error: "Batas percobaan kode verifikasi tercapai (5 kali). Silakan minta kode verifikasi baru.",
        });
        return;
      }

      const expectedHash = hashOtpCode(email, rawCode, challenge.nonce);
      const expectedBuf = Buffer.from(expectedHash, "hex");
      const actualBuf = Buffer.from(challenge.codeHash, "hex");
      const isMatch =
        expectedBuf.length === actualBuf.length && crypto.timingSafeEqual(expectedBuf, actualBuf);

      if (!isMatch) {
        let remaining = OTP_MAX_ATTEMPTS - 1;
        if (memRecord && memRecord.nonce === challenge.nonce) {
          memRecord.attempts += 1;
          remaining = Math.max(0, OTP_MAX_ATTEMPTS - memRecord.attempts);
        }
        res.status(400).json({
          ok: false,
          error:
            remaining > 0
              ? `Kode verifikasi tidak sesuai. Sisa percobaan: ${remaining} kali.`
              : "Kode verifikasi salah dan batas percobaan habis. Silakan kirim ulang kode baru.",
        });
        return;
      }

      // Mark OTP challenge as used so it cannot be replayed
      usedNonces.add(challenge.nonce);
      if (memRecord && memRecord.nonce === challenge.nonce) {
        memRecord.used = true;
      }

      res.json({
        ok: true,
        verified: true,
        email: challenge.email,
        uid: challenge.uid,
        verifiedAt: new Date().toISOString(),
      });
    } catch (err: any) {
      console.error("[OTP Verify Error]:", err);
      res.status(500).json({
        ok: false,
        error: "Terjadi kesalahan saat memvalidasi kode verifikasi.",
      });
    }
  });

  // Mount Vite middleware in development or serve static dist in production
  const isProduction = process.env.NODE_ENV === "production";
  if (!isProduction) {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(__dirname, "dist");
    app.use(express.static(distDir));
    app.get("*", (_req, res) => {
      const indexFile = path.join(distDir, "index.html");
      if (fs.existsSync(indexFile)) {
        res.sendFile(indexFile);
      } else {
        res.status(404).send("Build output not found. Run npm run build first.");
      }
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`[SISPERMBG] Server aktif pada http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error("[SISPERMBG] Gagal menjalankan server:", err);
  process.exit(1);
});
