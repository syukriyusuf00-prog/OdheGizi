/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { doc, getDoc, setDoc, deleteDoc, serverTimestamp } from "firebase/firestore";
import { db } from "../lib/firebase";

const IDB_NAME = "sispermbg_ikss_photos_db";
const IDB_VERSION = 1;
const STORE_NAME = "photos";

// Stay safely below firestore.rules 950,000 byte cap per state doc
const MAX_BATCH_CHARS = 520_000;
const MAX_BATCH_COUNT = 24;
const CLOUD_TIMEOUT_MS = 8500;

interface PhotoManifest {
  version: number;
  batchCount: number;
  totalPhotos: number;
  updatedAt: string;
}

function withTimeout<T>(promise: Promise<T>, ms: number, fallback: T): Promise<T> {
  return new Promise((resolve) => {
    let settled = false;
    const timer = window.setTimeout(() => {
      if (!settled) {
        settled = true;
        resolve(fallback);
      }
    }, ms);
    promise
      .then((val) => {
        if (!settled) {
          settled = true;
          window.clearTimeout(timer);
          resolve(val);
        }
      })
      .catch(() => {
        if (!settled) {
          settled = true;
          window.clearTimeout(timer);
          resolve(fallback);
        }
      });
  });
}

function openPhotoDb(): Promise<IDBDatabase | null> {
  if (typeof indexedDB === "undefined") return Promise.resolve(null);
  return new Promise((resolve) => {
    try {
      const req = indexedDB.open(IDB_NAME, IDB_VERSION);
      req.onupgradeneeded = () => {
        const database = req.result;
        if (!database.objectStoreNames.contains(STORE_NAME)) {
          database.createObjectStore(STORE_NAME);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

function getIdbRecordKey(scopeId: string): string {
  return `ikss_photos_${scopeId || "local"}`;
}

export async function loadIkssPhotosLocal(scopeId: string): Promise<Record<string, string>> {
  const database = await openPhotoDb();
  if (!database) return {};
  return new Promise((resolve) => {
    try {
      const tx = database.transaction(STORE_NAME, "readonly");
      const store = tx.objectStore(STORE_NAME);
      const req = store.get(getIdbRecordKey(scopeId));
      req.onsuccess = () => {
        const res = req.result;
        if (res && typeof res === "object" && !Array.isArray(res)) {
          resolve(res as Record<string, string>);
        } else {
          resolve({});
        }
      };
      req.onerror = () => resolve({});
    } catch {
      resolve({});
    }
  });
}

export async function saveIkssPhotosLocal(
  scopeId: string,
  photos: Record<string, string>
): Promise<boolean> {
  const database = await openPhotoDb();
  if (!database) return false;
  return new Promise((resolve) => {
    try {
      const tx = database.transaction(STORE_NAME, "readwrite");
      const store = tx.objectStore(STORE_NAME);
      store.put(photos, getIdbRecordKey(scopeId));
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    } catch {
      resolve(false);
    }
  });
}

export async function clearIkssPhotosLocal(scopeId: string): Promise<void> {
  const database = await openPhotoDb();
  if (!database) return;
  await new Promise<void>((resolve) => {
    try {
      const tx = database.transaction(STORE_NAME, "readwrite");
      tx.objectStore(STORE_NAME).delete(getIdbRecordKey(scopeId));
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    } catch {
      resolve();
    }
  });
}

/**
 * Loads all chunked IKSS menu photos from Firestore `users/{uid}/states/ikss_photo_manifest`
 * and `ikss_photo_batch_0..N`.
 */
export async function loadIkssPhotosFromCloud(uid: string): Promise<Record<string, string> | null> {
  if (!uid) return null;
  try {
    const manifestRef = doc(db, "users", uid, "states", "ikss_photo_manifest");
    const manifestSnap = await withTimeout(getDoc(manifestRef), CLOUD_TIMEOUT_MS, null);
    if (!manifestSnap || !manifestSnap.exists()) return null;

    const rawManifest = manifestSnap.data()?.data;
    if (typeof rawManifest !== "string") return null;
    const manifest = JSON.parse(rawManifest) as PhotoManifest;
    const count = Math.min(Math.max(0, Number(manifest.batchCount) || 0), MAX_BATCH_COUNT);
    if (count === 0) return {};

    const batchPromises: Promise<Record<string, string>>[] = [];
    for (let i = 0; i < count; i++) {
      const batchKey = `ikss_photo_batch_${i}`;
      const batchRef = doc(db, "users", uid, "states", batchKey);
      batchPromises.push(
        withTimeout(getDoc(batchRef), CLOUD_TIMEOUT_MS, null).then((snap) => {
          if (!snap || !snap.exists()) return {};
          const raw = snap.data()?.data;
          if (typeof raw !== "string") return {};
          try {
            const parsed = JSON.parse(raw);
            return parsed && typeof parsed === "object" ? (parsed as Record<string, string>) : {};
          } catch {
            return {};
          }
        })
      );
    }

    const chunks = await Promise.all(batchPromises);
    const merged: Record<string, string> = {};
    for (const chunk of chunks) {
      Object.assign(merged, chunk);
    }
    return merged;
  } catch (err) {
    console.warn("Gagal memuat batch foto IKSS dari cloud:", err);
    return null;
  }
}

/**
 * Splits IKSS photos into <=520KB JSON chunks and writes them safely under
 * `users/{uid}/states/ikss_photo_batch_{i}` + `ikss_photo_manifest` in full compliance with firestore.rules.
 */
export async function saveIkssPhotosToCloud(
  uid: string,
  photos: Record<string, string>
): Promise<{ ok: boolean; batchCount: number; savedCount: number; error?: string }> {
  if (!uid) {
    return { ok: false, batchCount: 0, savedCount: 0, error: "Akun cloud belum aktif." };
  }

  try {
    const entries = Object.entries(photos).filter(
      ([k, v]) => typeof k === "string" && k.length > 0 && typeof v === "string" && v.startsWith("data:image/")
    );

    const batches: Record<string, string>[] = [];
    let currentBatch: Record<string, string> = {};
    let currentChars = 2; // "{}"

    for (const [k, v] of entries) {
      const entryCost = k.length + v.length + 8;
      if (currentChars + entryCost > MAX_BATCH_CHARS && Object.keys(currentBatch).length > 0) {
        if (batches.length < MAX_BATCH_COUNT - 1) {
          batches.push(currentBatch);
          currentBatch = {};
          currentChars = 2;
        } else {
          break;
        }
      }
      currentBatch[k] = v;
      currentChars += entryCost;
    }

    if (Object.keys(currentBatch).length > 0) {
      batches.push(currentBatch);
    }

    // Check previous batch count so we can clean up stale trailing batches
    let prevBatchCount = 0;
    try {
      const prevSnap = await withTimeout(
        getDoc(doc(db, "users", uid, "states", "ikss_photo_manifest")),
        3500,
        null
      );
      if (prevSnap && prevSnap.exists()) {
        const raw = prevSnap.data()?.data;
        if (typeof raw === "string") {
          const parsed = JSON.parse(raw) as PhotoManifest;
          prevBatchCount = Math.min(MAX_BATCH_COUNT, Math.max(0, Number(parsed.batchCount) || 0));
        }
      }
    } catch {
      prevBatchCount = 0;
    }

    // Write each batch sequentially or in small groups of 2 so network/browser stays silky smooth
    let savedCount = 0;
    for (let i = 0; i < batches.length; i++) {
      const batchKey = `ikss_photo_batch_${i}`;
      const serialized = JSON.stringify(batches[i]);
      if (serialized.length > 940_000) {
        continue;
      }
      const batchRef = doc(db, "users", uid, "states", batchKey);
      const written = await withTimeout(
        setDoc(batchRef, {
          userId: uid,
          stateKey: batchKey,
          data: serialized,
          updatedAt: serverTimestamp(),
        }).then(() => true),
        CLOUD_TIMEOUT_MS,
        false
      );
      if (!written) {
        return {
          ok: false,
          batchCount: i,
          savedCount,
          error: "Waktu unggah cloud habis. Foto tetap aman tersimpan di perangkat ini.",
        };
      }
      savedCount += Object.keys(batches[i]).length;
    }

    // Write manifest
    const manifest: PhotoManifest = {
      version: 1,
      batchCount: batches.length,
      totalPhotos: savedCount,
      updatedAt: new Date().toISOString(),
    };
    const manifestWritten = await withTimeout(
      setDoc(doc(db, "users", uid, "states", "ikss_photo_manifest"), {
        userId: uid,
        stateKey: "ikss_photo_manifest",
        data: JSON.stringify(manifest),
        updatedAt: serverTimestamp(),
      }).then(() => true),
      CLOUD_TIMEOUT_MS,
      false
    );

    // Clean up stale batches in background
    if (prevBatchCount > batches.length) {
      for (let i = batches.length; i < prevBatchCount; i++) {
        deleteDoc(doc(db, "users", uid, "states", `ikss_photo_batch_${i}`)).catch(() => {});
      }
    }

    if (!manifestWritten) {
      return {
        ok: false,
        batchCount: batches.length,
        savedCount,
        error: "Sebagian foto tersimpan lokal; sinkronisasi manifest cloud tertunda.",
      };
    }

    return { ok: true, batchCount: batches.length, savedCount };
  } catch (err: any) {
    console.warn("Gagal menyinkronkan foto IKSS ke cloud:", err);
    return {
      ok: false,
      batchCount: 0,
      savedCount: 0,
      error: err?.message || "Sinkronisasi cloud tertunda; foto aman di penyimpanan lokal.",
    };
  }
}

/**
 * Non-blocking photo compression that avoids FileReader.readAsDataURL heap spikes.
 * Prefers createImageBitmap (off-main-thread decode + EXIF orientation) with ObjectURL fallback.
 */
export async function compressMenuPhoto(
  file: File,
  options?: { maxWidth?: number; maxHeight?: number; targetMaxChars?: number }
): Promise<string> {
  if (!file || !file.type.startsWith("image/")) {
    throw new Error("Format file tidak didukung. Pilih file foto (JPG, PNG, WEBP).");
  }
  if (file.size > 30 * 1024 * 1024) {
    throw new Error("Ukuran file terlalu besar (maksimal 30 MB).");
  }

  const maxWidth = options?.maxWidth ?? 640;
  const maxHeight = options?.maxHeight ?? 480;
  const targetMaxChars = options?.targetMaxChars ?? 72_000;

  // Yield a frame so UI spinners render before canvas encode
  await new Promise<void>((r) => requestAnimationFrame(() => r()));

  let srcWidth = 0;
  let srcHeight = 0;
  let drawSource: CanvasImageSource | null = null;
  let bitmapToClose: ImageBitmap | null = null;
  let objectUrlToRevoke: string | null = null;

  try {
    if (typeof createImageBitmap === "function") {
      try {
        bitmapToClose = await createImageBitmap(file);
        srcWidth = bitmapToClose.width;
        srcHeight = bitmapToClose.height;
        drawSource = bitmapToClose;
      } catch {
        bitmapToClose = null;
      }
    }

    if (!drawSource) {
      objectUrlToRevoke = URL.createObjectURL(file);
      const img = await new Promise<HTMLImageElement>((resolve, reject) => {
        const el = new Image();
        el.decoding = "async";
        el.onload = () => resolve(el);
        el.onerror = () => reject(new Error("File gambar tidak dapat dibaca atau rusak."));
        el.src = objectUrlToRevoke!;
      });
      srcWidth = img.naturalWidth || img.width;
      srcHeight = img.naturalHeight || img.height;
      drawSource = img;
    }

    if (srcWidth <= 0 || srcHeight <= 0) {
      throw new Error("Dimensi foto tidak valid.");
    }

    const scale = Math.min(1, maxWidth / srcWidth, maxHeight / srcHeight);
    const outW = Math.max(1, Math.round(srcWidth * scale));
    const outH = Math.max(1, Math.round(srcHeight * scale));

    const canvas = document.createElement("canvas");
    canvas.width = outW;
    canvas.height = outH;
    const ctx = canvas.getContext("2d", { alpha: false });
    if (!ctx) {
      throw new Error("Browser tidak mendukung pemrosesan gambar kanvas.");
    }

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, outW, outH);
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(drawSource, 0, 0, outW, outH);

    // Adaptive JPEG quality so photos remain crisp while staying lightweight (~25-65KB)
    const qualities = [0.8, 0.7, 0.58, 0.46];
    let result = canvas.toDataURL("image/jpeg", qualities[0]);
    for (let i = 1; i < qualities.length && result.length > targetMaxChars; i++) {
      result = canvas.toDataURL("image/jpeg", qualities[i]);
    }
    return result;
  } finally {
    if (bitmapToClose && typeof bitmapToClose.close === "function") {
      try {
        bitmapToClose.close();
      } catch {}
    }
    if (objectUrlToRevoke) {
      try {
        URL.revokeObjectURL(objectUrlToRevoke);
      } catch {}
    }
  }
}

/**
 * Rotates an already-compressed menu photo 90 degrees clockwise without blocking the UI.
 */
export async function rotatePhotoDataUrl90(dataUrl: string): Promise<string> {
  await new Promise<void>((r) => requestAnimationFrame(() => r()));
  const img = await new Promise<HTMLImageElement>((resolve, reject) => {
    const el = new Image();
    el.decoding = "async";
    el.onload = () => resolve(el);
    el.onerror = () => reject(new Error("Gagal memutar foto."));
    el.src = dataUrl;
  });

  const w = img.naturalWidth || img.width;
  const h = img.naturalHeight || img.height;
  const canvas = document.createElement("canvas");
  canvas.width = h;
  canvas.height = w;
  const ctx = canvas.getContext("2d", { alpha: false });
  if (!ctx) throw new Error("Gagal memproses rotasi foto.");

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate(Math.PI / 2);
  ctx.drawImage(img, -w / 2, -h / 2, w, h);

  return canvas.toDataURL("image/jpeg", 0.8);
}
