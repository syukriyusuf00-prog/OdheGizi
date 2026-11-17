import { doc, setDoc, serverTimestamp } from "firebase/firestore";
import { db } from "./firebase.ts";

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
}

const DEVICE_ID_STORAGE_KEY = "sisper_device_id";
const DEVICE_FIRST_SEEN_KEY = "sisper_device_first_seen";
const DEVICE_CACHE_KEY = "sisper_current_device_info";

let cachedServerMeta: { ipAddress: string } | null = null;
let serverMetaPromise: Promise<{ ipAddress: string }> | null = null;

const clampStr = (val: unknown, maxLen: number): string => {
  if (typeof val !== "string") return "";
  return val.trim().slice(0, maxLen);
};

export const getOrCreateDeviceId = (): string => {
  try {
    const existing = localStorage.getItem(DEVICE_ID_STORAGE_KEY);
    if (existing && /^[a-zA-Z0-9_-]{8,64}$/.test(existing)) {
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
  try {
    localStorage.setItem(DEVICE_ID_STORAGE_KEY, newId);
  } catch {}
  return newId;
};

const getOrCreateFirstSeenAt = (): string => {
  try {
    const existing = localStorage.getItem(DEVICE_FIRST_SEEN_KEY);
    if (existing) return existing;
    const now = new Date().toISOString();
    localStorage.setItem(DEVICE_FIRST_SEEN_KEY, now);
    return now;
  } catch {
    return new Date().toISOString();
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
  const serverMeta = await fetchServerDeviceMeta();

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

export const formatDeviceShortLabel = (d?: Partial<DeviceRecord> | Partial<CompactDeviceSummary> | null): string => {
  if (!d) return "-";
  const parts = [d.deviceModel || d.osName, d.browserName].filter(Boolean);
  return parts.join(" · ") || "Perangkat Web";
};

export const syncDeviceToFirestore = async (context: DeviceSyncContext): Promise<DeviceRecord | null> => {
  try {
    const record = await buildDeviceRecord(context);

    try {
      localStorage.setItem(DEVICE_CACHE_KEY, JSON.stringify(record));
    } catch {}

    try {
      window.dispatchEvent(new CustomEvent("sisper_device_updated", { detail: record }));
    } catch {}

    // 1. Write to top-level /devices/{deviceId} so AdminPanel reads EVERY device live
    const deviceDocRef = doc(db, "devices", record.deviceId);
    const firestorePayload = {
      ...record,
      updatedAt: serverTimestamp(),
    };

    const writeTasks: Promise<unknown>[] = [
      setDoc(deviceDocRef, firestorePayload, { merge: true }).catch((err) => {
        console.warn("Notice syncing device to /devices:", err);
      }),
    ];

    // 2. Also mirror device onto the authenticated user's profile & /users/{uid}/devices/{deviceId}
    if (record.userUid && record.userUid !== "demo_user_5min") {
      const compact = toCompactDeviceSummary(record);
      const userDeviceRef = doc(db, "users", record.userUid, "devices", record.deviceId);
      const userRef = doc(db, "users", record.userUid);

      writeTasks.push(
        setDoc(userDeviceRef, firestorePayload, { merge: true }).catch((err) => {
          console.warn("Notice syncing user device subcollection:", err);
        }),
        setDoc(
          userRef,
          {
            lastDevice: compact,
            lastDeviceLabel: formatDeviceShortLabel(compact),
            lastDeviceId: record.deviceId,
            lastDeviceSeenAt: record.lastSeenAt,
          },
          { merge: true }
        ).catch((err) => {
          console.warn("Notice updating lastDevice on user profile:", err);
        })
      );
    }

    await Promise.all(writeTasks);
    return record;
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
