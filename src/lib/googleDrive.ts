/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { GoogleAuthProvider, signInWithPopup } from "firebase/auth";
import { auth } from "./firebase";

export const GOOGLE_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";

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

// Strict in-memory OAuth token storage (never persisted to localStorage or sessionStorage)
let inMemoryDriveSession: DriveSessionInfo | null = null;
let isDriveConnectingFlag = false;

export function getDriveSession(): DriveSessionInfo | null {
  return inMemoryDriveSession;
}

export function clearDriveSession(): void {
  inMemoryDriveSession = null;
}

export function getIsDriveConnecting(): boolean {
  return isDriveConnectingFlag;
}

export function sanitizeDriveName(name: string, fallback = "Arsip"): string {
  const cleaned = (name || "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned) return fallback;
  return cleaned.slice(0, 90);
}

/**
 * Connects user's active Google account for Google Drive access using Firebase Auth popup.
 * Uses only the existing Firebase app/auth instance and stores the OAuth access token in memory.
 */
export async function connectGoogleDrive(hintEmail?: string): Promise<DriveSessionInfo> {
  const provider = new GoogleAuthProvider();
  provider.addScope(GOOGLE_DRIVE_SCOPE);

  const customParams: Record<string, string> = {
    prompt: "select_account",
  };
  if (hintEmail && hintEmail.includes("@") && !hintEmail.endsWith("@sispermbg.id")) {
    customParams.login_hint = hintEmail.trim();
  }
  provider.setCustomParameters(customParams);

  isDriveConnectingFlag = true;
  try {
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    const accessToken = credential?.accessToken;

    if (!accessToken) {
      throw new Error(
        "Token akses Google Drive tidak diterima. Pastikan izin akses Google Drive dicentang saat login Google."
      );
    }

    const session: DriveSessionInfo = {
      accessToken,
      email: result.user?.email || hintEmail || "Akun Google Terhubung",
      displayName: result.user?.displayName || undefined,
      connectedAt: Date.now(),
    };
    inMemoryDriveSession = session;
    return session;
  } catch (err: any) {
    const code = err?.code || "";
    if (code === "auth/unauthorized-domain") {
      const host = typeof window !== "undefined" ? window.location.hostname : "domain ini";
      throw new Error(
        `Domain "${host}" belum terdaftar di Firebase Authentication → Settings → Authorized domains. Tambahkan domain tersebut di Firebase Console lalu coba hubungkan Google Drive kembali.`
      );
    }
    if (code === "auth/popup-blocked") {
      throw new Error(
        "Pop-up Google Drive diblokir browser. Izinkan pop-up atau klik 'Buka di Tab Baru' di kanan atas pratinjau."
      );
    }
    if (code === "auth/popup-closed-by-user") {
      throw new Error("Jendela login Google Drive ditutup sebelum otorisasi selesai.");
    }
    if (code === "auth/network-request-failed") {
      throw new Error(
        "Koneksi ke Google Auth terhalang iframe browser. Silakan buka aplikasi di Tab Baru lalu hubungkan Google Drive."
      );
    }
    throw err;
  } finally {
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
