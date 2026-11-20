/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { getApps, initializeApp } from "firebase/app";
import {
  getAuth,
  GoogleAuthProvider,
  inMemoryPersistence,
  setPersistence,
  signInWithPopup,
  signOut,
} from "firebase/auth";
import firebaseConfig from "../../firebase-applet-config.json";

export const GOOGLE_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";
export const DRIVE_FIREBASE_PROJECT_ID = firebaseConfig.projectId;
export const DRIVE_OAUTH_AUDIENCE_URL = `https://console.cloud.google.com/auth/audience?project=${encodeURIComponent(
  DRIVE_FIREBASE_PROJECT_ID
)}`;

export interface DriveSessionInfo {
  accessToken: string;
  email: string;
  displayName?: string;
  connectedAt: number;
}

export interface DriveUploadResult {
  fileId: string;
  name: string;
  webViewLink?: string;
  folderId: string;
  folderWebViewLink?: string;
}

export type DriveAuthErrorCode =
  | "testing_access_denied"
  | "unauthorized_domain"
  | "popup_blocked"
  | "popup_closed"
  | "popup_timeout"
  | "cancelled"
  | "network_blocked"
  | "missing_token"
  | "unknown";

export class DriveConnectionError extends Error {
  code: DriveAuthErrorCode;
  attemptedEmail?: string;

  constructor(code: DriveAuthErrorCode, message: string, attemptedEmail?: string) {
    super(message);
    this.name = "DriveConnectionError";
    this.code = code;
    this.attemptedEmail = attemptedEmail;
  }
}

export interface ConnectDriveOptions {
  skipLoginHint?: boolean;
}

// Dedicated secondary Firebase App with strictly in-memory persistence so connecting Drive
// never overwrites or hijacks the user's primary SISPERMBG authentication session.
const DRIVE_APP_NAME = "sisper_drive_oauth";
function getDriveAuthInstance() {
  const existing = getApps().find((a) => a.name === DRIVE_APP_NAME);
  const driveApp = existing || initializeApp(firebaseConfig, DRIVE_APP_NAME);
  return getAuth(driveApp);
}

// Strict in-memory OAuth token storage (never persisted to localStorage or sessionStorage)
let inMemoryDriveSession: DriveSessionInfo | null = null;
let isDriveConnectingFlag = false;
let activeCancelReject: ((err: DriveConnectionError) => void) | null = null;

export function getDriveSession(): DriveSessionInfo | null {
  return inMemoryDriveSession;
}

export function clearDriveSession(): void {
  inMemoryDriveSession = null;
  try {
    const driveAuth = getDriveAuthInstance();
    void signOut(driveAuth).catch(() => {});
  } catch {
    // ignore secondary auth cleanup errors
  }
}

export function getIsDriveConnecting(): boolean {
  return isDriveConnectingFlag;
}

export function cancelGoogleDriveConnect(): void {
  if (activeCancelReject) {
    const rejectFn = activeCancelReject;
    activeCancelReject = null;
    isDriveConnectingFlag = false;
    rejectFn(
      new DriveConnectionError(
        "cancelled",
        "Proses menghubungkan Google Drive dibatalkan. Anda dapat mencoba lagi kapan saja."
      )
    );
  } else {
    isDriveConnectingFlag = false;
  }
}

export function sanitizeDriveName(name: string, fallback = "Arsip"): string {
  const cleaned = (name || "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned) return fallback;
  return cleaned.slice(0, 90);
}

function classifyDriveAuthError(err: any, attemptedEmail?: string): DriveConnectionError {
  if (err instanceof DriveConnectionError) return err;

  const code = String(err?.code || "").toLowerCase();
  const rawMsg = String(err?.message || err || "");
  const lowerMsg = rawMsg.toLowerCase();

  if (
    lowerMsg.includes("access_denied") ||
    lowerMsg.includes("403") ||
    lowerMsg.includes("verification process") ||
    lowerMsg.includes("being tested") ||
    lowerMsg.includes("developer-approved") ||
    lowerMsg.includes("org_internal") ||
    lowerMsg.includes("admin_policy_enforced") ||
    code === "auth/admin-restricted-operation"
  ) {
    const emailLabel = attemptedEmail ? ` (${attemptedEmail})` : "";
    return new DriveConnectionError(
      "testing_access_denied",
      `Akses Google Drive diblokir (Error 403: access_denied) untuk akun${emailLabel} karena aplikasi OAuth Google (${DRIVE_FIREBASE_PROJECT_ID}) masih berstatus Testing. Tambahkan email tersebut di Google Auth Platform → Audience → Test users (atau klik Publish App), lalu coba lagi atau pilih akun developer yang sudah terdaftar.`,
      attemptedEmail
    );
  }

  if (code === "auth/unauthorized-domain") {
    const host = typeof window !== "undefined" ? window.location.hostname : "domain ini";
    return new DriveConnectionError(
      "unauthorized_domain",
      `Domain "${host}" belum terdaftar di Firebase Authentication → Settings → Authorized domains. Tambahkan domain tersebut di Firebase Console lalu coba hubungkan Google Drive kembali.`,
      attemptedEmail
    );
  }

  if (code === "auth/popup-blocked") {
    return new DriveConnectionError(
      "popup_blocked",
      "Pop-up Google Drive diblokir browser. Izinkan pop-up untuk situs ini atau buka aplikasi di Tab Baru lalu coba kembali.",
      attemptedEmail
    );
  }

  if (code === "auth/popup-closed-by-user" || code === "auth/cancelled-popup-request") {
    return new DriveConnectionError(
      "popup_closed",
      "Jendela login Google Drive ditutup sebelum otorisasi selesai. Jika di pop-up muncul pesan 'Access blocked / Error 403', tambahkan email Anda ke daftar Test Users Google Cloud atau pilih akun Google lain.",
      attemptedEmail
    );
  }

  if (code === "auth/network-request-failed") {
    return new DriveConnectionError(
      "network_blocked",
      "Koneksi ke Google Auth terhalang iframe atau jaringan. Silakan buka aplikasi di Tab Baru lalu hubungkan Google Drive.",
      attemptedEmail
    );
  }

  return new DriveConnectionError(
    "unknown",
    rawMsg || "Gagal menghubungkan Google Drive. Silakan coba lagi.",
    attemptedEmail
  );
}

/**
 * Connects user's Google account for Google Drive access using an isolated in-memory Firebase Auth popup.
 * Stores the OAuth access token strictly in memory so it never clobbers the primary SISPERMBG session.
 */
export async function connectGoogleDrive(
  hintEmail?: string,
  options?: ConnectDriveOptions
): Promise<DriveSessionInfo> {
  // Clean up any stale in-flight connection first so UI can never stay stuck
  if (activeCancelReject) {
    try {
      activeCancelReject(
        new DriveConnectionError("cancelled", "Permintaan login sebelumnya digantikan permintaan baru.")
      );
    } catch {
      // ignore
    }
    activeCancelReject = null;
  }

  const cleanHint =
    !options?.skipLoginHint &&
    hintEmail &&
    hintEmail.includes("@") &&
    !hintEmail.endsWith("@sispermbg.id")
      ? hintEmail.trim()
      : undefined;

  const provider = new GoogleAuthProvider();
  provider.addScope(GOOGLE_DRIVE_SCOPE);

  const customParams: Record<string, string> = {
    prompt: "select_account consent",
  };
  if (cleanHint) {
    customParams.login_hint = cleanHint;
  }
  provider.setCustomParameters(customParams);

  const driveAuth = getDriveAuthInstance();
  isDriveConnectingFlag = true;

  let timeoutId: ReturnType<typeof setTimeout> | null = null;

  try {
    await setPersistence(driveAuth, inMemoryPersistence).catch(() => {});

    const popupPromise = signInWithPopup(driveAuth, provider);
    const guardPromise = new Promise<never>((_, reject) => {
      activeCancelReject = reject;
      timeoutId = setTimeout(() => {
        activeCancelReject = null;
        reject(
          new DriveConnectionError(
            "popup_timeout",
            "Waktu tunggu otorisasi Google Drive habis (popup tertutup atau tertahan). Silakan klik Hubungkan ulang atau pilih akun Google lain.",
            cleanHint
          )
        );
      }, 120_000);
    });

    const result = await Promise.race([popupPromise, guardPromise]);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    const accessToken = credential?.accessToken;

    if (!accessToken) {
      throw new DriveConnectionError(
        "missing_token",
        "Token akses Google Drive tidak diterima. Pastikan kotak izin Google Drive dicentang saat login Google.",
        result.user?.email || cleanHint
      );
    }

    const session: DriveSessionInfo = {
      accessToken,
      email: result.user?.email || cleanHint || "Akun Google Terhubung",
      displayName: result.user?.displayName || undefined,
      connectedAt: Date.now(),
    };
    inMemoryDriveSession = session;
    return session;
  } catch (err: any) {
    throw classifyDriveAuthError(err, cleanHint);
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
    activeCancelReject = null;
    isDriveConnectingFlag = false;
  }
}

function escapeDriveQueryValue(val: string): string {
  return val.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

async function handleDriveApiResponse<T>(res: Response, actionLabel: string): Promise<T> {
  if (res.status === 401 || res.status === 403) {
    if (res.status === 401) {
      clearDriveSession();
    }
    let detail = "";
    try {
      const errBody = await res.json();
      detail = errBody?.error?.message || "";
    } catch {
      // ignore parse failure
    }
    throw new Error(
      `Otorisasi Google Drive gagal (${res.status}) saat ${actionLabel}. ${detail || "Silakan hubungkan ulang Google Drive Anda."}`
    );
  }
  if (!res.ok) {
    let detail = "";
    try {
      const errBody = await res.json();
      detail = errBody?.error?.message || JSON.stringify(errBody);
    } catch {
      detail = await res.text();
    }
    throw new Error(`Gagal ${actionLabel} (${res.status}): ${detail}`);
  }
  return res.json() as Promise<T>;
}

/**
 * Finds an existing Google Drive folder by name (and optional parentId) or creates a new one.
 */
export async function findOrCreateDriveFolder(
  accessToken: string,
  folderName: string,
  parentId?: string
): Promise<{ id: string; name: string; webViewLink?: string }> {
  const safeName = sanitizeDriveName(folderName, "Folder_Arsip");
  const queryParts = [
    `mimeType = 'application/vnd.google-apps.folder'`,
    `name = '${escapeDriveQueryValue(safeName)}'`,
    `trashed = false`,
  ];
  if (parentId) {
    queryParts.push(`'${escapeDriveQueryValue(parentId)}' in parents`);
  }

  const searchUrl =
    `https://www.googleapis.com/drive/v3/files?` +
    new URLSearchParams({
      q: queryParts.join(" and "),
      fields: "files(id, name, webViewLink)",
      spaces: "drive",
      pageSize: "5",
    }).toString();

  const searchRes = await fetch(searchUrl, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const searchData = await handleDriveApiResponse<{
    files?: Array<{ id: string; name: string; webViewLink?: string }>;
  }>(searchRes, `mencari folder "${safeName}"`);

  if (searchData.files && searchData.files.length > 0) {
    return searchData.files[0];
  }

  const metadata: Record<string, any> = {
    name: safeName,
    mimeType: "application/vnd.google-apps.folder",
  };
  if (parentId) {
    metadata.parents = [parentId];
  }

  const createRes = await fetch(
    "https://www.googleapis.com/drive/v3/files?fields=id,name,webViewLink",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(metadata),
    }
  );

  return handleDriveApiResponse<{ id: string; name: string; webViewLink?: string }>(
    createRes,
    `membuat folder "${safeName}"`
  );
}

/**
 * Uploads a PDF Blob into a specific Google Drive folder using multipart/related upload.
 */
export async function uploadPdfToDriveFolder(
  accessToken: string,
  folderId: string,
  fileName: string,
  pdfBlob: Blob
): Promise<DriveUploadResult> {
  const safeFileName = sanitizeDriveName(fileName.replace(/\.pdf$/i, ""), "Arsip") + ".pdf";
  const metadata = {
    name: safeFileName,
    mimeType: "application/pdf",
    parents: [folderId],
  };

  const boundary = `sispermbg_drive_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  const metadataPart =
    `--${boundary}\r\n` +
    `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
    `${JSON.stringify(metadata)}\r\n`;
  const mediaHeader =
    `--${boundary}\r\n` +
    `Content-Type: application/pdf\r\n\r\n`;
  const closingBoundary = `\r\n--${boundary}--`;

  const bodyBlob = new Blob([metadataPart, mediaHeader, pdfBlob, closingBoundary], {
    type: `multipart/related; boundary=${boundary}`,
  });

  const uploadRes = await fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
      body: bodyBlob,
    }
  );

  const created = await handleDriveApiResponse<{
    id: string;
    name: string;
    webViewLink?: string;
  }>(uploadRes, `mengunggah file PDF "${safeFileName}"`);

  return {
    fileId: created.id,
    name: created.name,
    webViewLink: created.webViewLink,
    folderId,
  };
}
