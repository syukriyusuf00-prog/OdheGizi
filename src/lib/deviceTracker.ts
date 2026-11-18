import { doc, setDoc, serverTimestamp } from "firebase/firestore";
import { db, reportFirestoreWriteResult } from "./firebase.ts";

export type DeviceSessionState = "guest" | "registering" | "active" | "demo";

export interface DeviceRecord {
  deviceId: string;
  deviceType: "Mobile" | "Tablet" | "Desktop / Laptop";
  deviceModel: string;
  osName: string;
  browserName: string;
  screenResolution: string;
  viewportSize: string;
  language: string;
  timezone: string;
  ipAddress: string;
  userAgent: string;
  sessionState: DeviceSessionState;
  userUid: string | null;
  userEmail: string | null;
  userName: string | null;
  namaSPPG: string | null;
  peran: string | null;
  statusPersetujuan: string | null;
  isOnline: boolean;
  firstSeenAt: string;
  lastSeenAt: string;
  updatedAt?: unknown;
}

export interface CompactDeviceSummary {
  deviceId: string;
  deviceType: "Mobile" | "Tablet" | "Desktop / Laptop";
  deviceModel: string;
  osName: string;
  browserName: string;
  ipAddress: string;
  screenResolution: string;
  timezone: string;
  lastSeenAt: string;
  isOnline: boolean;
}

export interface DeviceSyncContext {
  sessionState: DeviceSessionState;
  userUid?: string | null;
  userEmail?: string | null;
  userName?: string | null;
  namaSPPG?: string | null;
  peran?: string | null;
  statusPersetujuan?: string | null;
  isOnline?: boolean;
  forceFirestore?: boolean;
  useBeacon?: boolean;
}

const DEVICE_ID_STORAGE_KEY = "sisper_device_id";
const DEVICE_FIRST_SEEN_KEY = "sisper_device_first_seen";
const DEVICE_CACHE_KEY = "sisper_current_device_info";

// In-memory fallbacks for restricted mobile Private Mode / WebViews where localStorage may throw
let inMemoryDeviceId: string | null = null;
let inMemoryFirstSeenAt: string | null = null;
let cachedServerMeta: { ipAddress: string } | null = null;
let serverMetaPromise: Promise<{ ipAddress: string }> | null = null;

// Quota protection: track last Firestore write signature & time so heartbeats do NOT exhaust free daily write units
let lastFirestoreDeviceSig = "";
let lastFirestoreDeviceWriteMs = 0;
const FIRESTORE_DEVICE_THROTTLE_MS = 5 * 60 * 1000; // Max once every 5 minutes unless session state / user changes

const clampStr = (val: unknown, maxLen: number): string => {
  if (typeof val !== "string") return "";
  return val.trim().slice(0, maxLen);
};

export const getOrCreateDeviceId = (): string => {
  if (inMemoryDeviceId) return inMemoryDeviceId;

  try {
    const existing = localStorage.getItem(DEVICE_ID_STORAGE_KEY);
    if (existing && /^[a-zA-Z0-9_-]{8,64}$/.test(existing)) {
      inMemoryDeviceId = existing;
      return existing;
    }
  } catch {}

  let randomPart = "";
  try {
    if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
      randomPart = crypto.randomUUID().replace(/-/g, "").slice(0, 20);
    }
  } catch {}

  if (!randomPart) {
    randomPart = `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 12)}`;
  }

  const newId = `dev_${randomPart}`.replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 48);
  inMemoryDeviceId = newId;
  try {
    localStorage.setItem(DEVICE_ID_STORAGE_KEY, newId);
  } catch {}
  return newId;
};

const getOrCreateFirstSeenAt = (): string => {
  if (inMemoryFirstSeenAt) return inMemoryFirstSeenAt;
  try {
    const existing = localStorage.getItem(DEVICE_FIRST_SEEN_KEY);
    if (existing) {
      inMemoryFirstSeenAt = existing;
      return existing;
    }
    const now = new Date().toISOString();
    inMemoryFirstSeenAt = now;
    localStorage.setItem(DEVICE_FIRST_SEEN_KEY, now);
    return now;
  } catch {
    const now = new Date().toISOString();
    inMemoryFirstSeenAt = now;
    return now;
  }
};

const parseAndroidModel = (ua: string): string => {
  // Typical Android UA: "Mozilla/5.0 (Linux; Android 14; SM-S918B Build/...) ..."
  const match = ua.match(/Android\s[\d.]+;\s*([^;)]+?)(?:\s+Build\/|[;)])/i);
  if (match && match[1]) {
    const raw = match[1].trim();
    if (raw && raw.toUpperCase() !== "K" && raw.toLowerCase() !== "wv") {
      return `Android (${raw.slice(0, 32)})`;
    }
  }
  return "Perangkat Android";
};

export const detectClientDeviceInfo = () => {
  const ua = typeof navigator !== "undefined" ? navigator.userAgent || "" : "";
  const platform = typeof navigator !== "undefined" ? navigator.platform || "" : "";
  const maxTouchPoints = typeof navigator !== "undefined" ? navigator.maxTouchPoints || 0 : 0;

  let deviceType: "Mobile" | "Tablet" | "Desktop / Laptop" = "Desktop / Laptop";
  let deviceModel = "Komputer / Laptop";
  let osName = "Tidak Diketahui";
  let browserName = "Browser Web";

  // OS & Device Model detection
  if (/iPhone/i.test(ua)) {
    deviceType = "Mobile";
    deviceModel = "iPhone";
    const iosMatch = ua.match(/OS (\d+[_.\d]*)/i);
    osName = iosMatch ? `iOS ${iosMatch[1].replace(/_/g, ".")}` : "iOS";
  } else if (/iPad/i.test(ua) || (/Macintosh/i.test(ua) && maxTouchPoints > 1)) {
    deviceType = "Tablet";
    deviceModel = "iPad";
    const iosMatch = ua.match(/OS (\d+[_.\d]*)/i);
    osName = iosMatch ? `iPadOS ${iosMatch[1].replace(/_/g, ".")}` : "iPadOS";
  } else if (/Android/i.test(ua)) {
    const isMobile = /Mobile/i.test(ua);
    deviceType = isMobile ? "Mobile" : "Tablet";
    deviceModel = parseAndroidModel(ua);
    const andMatch = ua.match(/Android\s([\d.]+)/i);
    osName = andMatch ? `Android ${andMatch[1]}` : "Android";
  } else if (/CrOS/i.test(ua)) {
    deviceType = "Desktop / Laptop";
    deviceModel = "Chromebook";
    osName = "ChromeOS";
  } else if (/Windows NT/i.test(ua)) {
    deviceType = "Desktop / Laptop";
    deviceModel = "PC / Laptop Windows";
    if (/Windows NT 10\.0/i.test(ua)) osName = "Windows 10/11";
    else if (/Windows NT 6\.3/i.test(ua)) osName = "Windows 8.1";
    else if (/Windows NT 6\.2/i.test(ua)) osName = "Windows 8";
    else if (/Windows NT 6\.1/i.test(ua)) osName = "Windows 7";
    else osName = "Windows";
  } else if (/Mac OS X/i.test(ua) || /MacIntel/i.test(platform)) {
    deviceType = "Desktop / Laptop";
    deviceModel = "MacBook / Mac";
    const macMatch = ua.match(/Mac OS X\s([\d_]+)/i);
    osName = macMatch ? `macOS ${macMatch[1].replace(/_/g, ".")}` : "macOS";
  } else if (/Linux/i.test(ua)) {
    deviceType = "Desktop / Laptop";
    deviceModel = "PC / Laptop Linux";
    osName = "Linux";
  }

  // Browser detection
  const edg = ua.match(/Edg(?:A|iOS)?\/([\d.]+)/i);
  const opr = ua.match(/OPR\/([\d.]+)/i);
  const samsung = ua.match(/SamsungBrowser\/([\d.]+)/i);
  const firefox = ua.match(/(?:Firefox|FxiOS)\/([\d.]+)/i);
  const chrome = ua.match(/(?:Chrome|CriOS)\/([\d.]+)/i);
  const safari = ua.match(/Version\/([\d.]+).*Safari/i);

  if (edg) browserName = `Edge ${edg[1].split(".")[0]}`;
  else if (opr) browserName = `Opera ${opr[1].split(".")[0]}`;
  else if (samsung) browserName = `Samsung Browser ${samsung[1].split(".")[0]}`;
  else if (firefox) browserName = `Firefox ${firefox[1].split(".")[0]}`;
  else if (chrome) browserName = `Chrome ${chrome[1].split(".")[0]}`;
  else if (safari) browserName = `Safari ${safari[1].split(".")[0]}`;
  else if (/Safari/i.test(ua)) browserName = "Safari";

  const screenResolution =
    typeof window !== "undefined" && window.screen
      ? `${window.screen.width}x${window.screen.height}`
      : "-";
  const viewportSize =
    typeof window !== "undefined" ? `${window.innerWidth}x${window.innerHeight}` : "-";
  const language = typeof navigator !== "undefined" ? navigator.language || "id-ID" : "id-ID";

  let timezone = "Asia/Jakarta";
  try {
    timezone = Intl.DateTimeFormat().resolvedOptions().timeZone || "Asia/Jakarta";
  } catch {}

  return {
    deviceType,
    deviceModel: clampStr(deviceModel, 80),
    osName: clampStr(osName, 60),
    browserName: clampStr(browserName, 60),
    screenResolution: clampStr(screenResolution, 32),
    viewportSize: clampStr(viewportSize, 32),
    language: clampStr(language, 24),
    timezone: clampStr(timezone, 64),
    userAgent: clampStr(ua, 280),
  };
};

export const fetchServerDeviceMeta = async (): Promise<{ ipAddress: string }> => {
  if (cachedServerMeta) return cachedServerMeta;
  if (serverMetaPromise) return serverMetaPromise;

  serverMetaPromise = (async () => {
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 3500);
      const res = await fetch("/api/device/info", {
        method: "GET",
        signal: controller.signal,
        headers: { Accept: "application/json" },
      });
      clearTimeout(timeout);
      if (res.ok) {
        const json = await res.json();
        const ipAddress = clampStr(json?.ipAddress || "Terdeteksi", 64);
        cachedServerMeta = { ipAddress };
        return cachedServerMeta;
      }
    } catch {
      // Fallback cleanly if endpoint is temporarily unreachable
    }
    return { ipAddress: "Jaringan Lokal" };
  })();

  const result = await serverMetaPromise;
  serverMetaPromise = null;
  return result;
};

export const buildDeviceRecord = async (context: DeviceSyncContext): Promise<DeviceRecord> => {
  const deviceId = getOrCreateDeviceId();
  const firstSeenAt = getOrCreateFirstSeenAt();
  const nowIso = new Date().toISOString();
  const clientInfo = detectClientDeviceInfo();
  const serverMeta = context.useBeacon
    ? cachedServerMeta || { ipAddress: "Terdeteksi" }
    : await fetchServerDeviceMeta();

  return {
    deviceId,
    deviceType: clientInfo.deviceType,
    deviceModel: clientInfo.deviceModel,
    osName: clientInfo.osName,
    browserName: clientInfo.browserName,
    screenResolution: clientInfo.screenResolution,
    viewportSize: clientInfo.viewportSize,
    language: clientInfo.language,
    timezone: clientInfo.timezone,
    ipAddress: serverMeta.ipAddress,
    userAgent: clientInfo.userAgent,
    sessionState: context.sessionState,
    userUid: context.userUid ? clampStr(context.userUid, 128) : null,
    userEmail: context.userEmail ? clampStr(context.userEmail.toLowerCase(), 256) : null,
    userName: context.userName ? clampStr(context.userName, 160) : null,
    namaSPPG: context.namaSPPG ? clampStr(context.namaSPPG, 160) : null,
    peran: context.peran ? clampStr(context.peran, 32) : null,
    statusPersetujuan: context.statusPersetujuan ? clampStr(context.statusPersetujuan, 32) : null,
    isOnline: context.isOnline !== false,
    firstSeenAt,
    lastSeenAt: nowIso,
  };
};

export const toCompactDeviceSummary = (record: DeviceRecord): CompactDeviceSummary => ({
  deviceId: record.deviceId,
  deviceType: record.deviceType,
  deviceModel: record.deviceModel,
  osName: record.osName,
  browserName: record.browserName,
  ipAddress: record.ipAddress,
  screenResolution: record.screenResolution,
  timezone: record.timezone,
  lastSeenAt: record.lastSeenAt,
  isOnline: record.isOnline,
});

export const formatDeviceShortLabel = (
  d?: Partial<DeviceRecord> | Partial<CompactDeviceSummary> | null
): string => {
  if (!d) return "-";
  const parts = [d.deviceModel || d.osName, d.browserName].filter(Boolean);
  return parts.join(" · ") || "Perangkat Web";
};

const syncDeviceToRealtimeServer = async (
  record: DeviceRecord,
  useBeacon = false
): Promise<DeviceRecord | null> => {
  try {
    const serialized = JSON.stringify(record);
    if (useBeacon && typeof navigator !== "undefined" && typeof navigator.sendBeacon === "function") {
      navigator.sendBeacon("/api/realtime/sync-device", new Blob([serialized], { type: "text/plain" }));
      return record;
    }

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 4500);
    const resp = await fetch("/api/realtime/sync-device", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: serialized,
      keepalive: true,
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (resp.ok) {
      const json = await resp.json().catch(() => null);
      if (json?.device?.ipAddress) {
        cachedServerMeta = { ipAddress: json.device.ipAddress };
        return json.device as DeviceRecord;
      }
    }
  } catch {
    // Non-fatal fallback
  }
  return record;
};

export const syncDeviceToFirestore = async (
  context: DeviceSyncContext
): Promise<DeviceRecord | null> => {
  try {
    const record = await buildDeviceRecord(context);

    // 1. Always push to Express Real-Time Hub first so Mobile <-> Laptop Admin Panel updates in < 1 second
    const syncedRecord = (await syncDeviceToRealtimeServer(record, Boolean(context.useBeacon))) || record;

    try {
      localStorage.setItem(DEVICE_CACHE_KEY, JSON.stringify(syncedRecord));
    } catch {}

    try {
      window.dispatchEvent(new CustomEvent("sisper_device_updated", { detail: syncedRecord }));
    } catch {}

    // Skip Firestore write during synchronous pagehide beacon to avoid mobile abort noise
    if (context.useBeacon) {
      return syncedRecord;
    }

    // 2. Quota-smart Firestore sync: only write to Firestore when session/user changes, forced, or > 5 mins elapsed
    const stateSig = `${syncedRecord.sessionState}|${syncedRecord.userUid || ""}|${syncedRecord.userEmail || ""}|${syncedRecord.statusPersetujuan || ""}|${syncedRecord.isOnline}`;
    const nowMs = Date.now();
    const shouldWriteFirestore =
      Boolean(context.forceFirestore) ||
      stateSig !== lastFirestoreDeviceSig ||
      nowMs - lastFirestoreDeviceWriteMs >= FIRESTORE_DEVICE_THROTTLE_MS;

    if (shouldWriteFirestore) {
      lastFirestoreDeviceSig = stateSig;
      lastFirestoreDeviceWriteMs = nowMs;

      const deviceDocRef = doc(db, "devices", syncedRecord.deviceId);
      const firestorePayload = {
        ...syncedRecord,
        updatedAt: serverTimestamp(),
      };

      const writeTask = setDoc(deviceDocRef, firestorePayload, { merge: true })
        .then(() => reportFirestoreWriteResult(true))
        .catch((err) => {
          reportFirestoreWriteResult(false, err);
          console.warn("Notice syncing device to /devices:", err);
        });

      await Promise.race([
        writeTask,
        new Promise((resolve) => setTimeout(resolve, 3500)),
      ]);
    }

    return syncedRecord;
  } catch (err) {
    console.warn("Notice during automatic device sync:", err);
    return null;
  }
};

export const getCachedLocalDevice = (): DeviceRecord | null => {
  try {
    const raw = localStorage.getItem(DEVICE_CACHE_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as DeviceRecord;
  } catch {
    return null;
  }
};
