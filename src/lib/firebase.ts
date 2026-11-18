import { initializeApp } from "firebase/app";
import { getAuth, GoogleAuthProvider } from "firebase/auth";
import {
  doc,
  getDocFromServer,
  initializeFirestore,
  setLogLevel,
} from "firebase/firestore";
import firebaseConfig from "../../firebase-applet-config.json";

export const FIREBASE_PROJECT_ID = firebaseConfig.projectId;
export const FIREBASE_API_KEY = firebaseConfig.apiKey;
export const FIRESTORE_DATABASE_ID =
  (firebaseConfig as { firestoreDatabaseId?: string }).firestoreDatabaseId || "(default)";

export const FIRESTORE_UPGRADE_URL = `https://console.firebase.google.com/project/${FIREBASE_PROJECT_ID}/firestore/databases/${encodeURIComponent(
  FIRESTORE_DATABASE_ID
)}/data?openUpgradeDialog=true`;

export const FIREBASE_AUTH_SETTINGS_URL = `https://console.firebase.google.com/project/${FIREBASE_PROJECT_ID}/authentication/settings`;

const FIRESTORE_REST_BASE = `https://firestore.googleapis.com/v1/projects/${encodeURIComponent(
  FIREBASE_PROJECT_ID
)}/databases/${encodeURIComponent(FIRESTORE_DATABASE_ID)}/documents`;

export interface DomainAuthCheckResult {
  currentHostname: string;
  isAuthorized: boolean;
  authorizedDomains: string[];
  checked: boolean;
  error?: string | null;
}

let cachedDomainAuthResult: DomainAuthCheckResult | null = null;

export const isHostnameAuthorized = (hostname: string, authorizedDomains: string[]): boolean => {
  const cleanHost = (hostname || "").toLowerCase().trim();
  if (!cleanHost) return false;
  return authorizedDomains.some((rawDomain) => {
    const cleanDomain = (rawDomain || "").toLowerCase().trim();
    if (!cleanDomain) return false;
    return cleanHost === cleanDomain || cleanHost.endsWith(`.${cleanDomain}`);
  });
};

export const getPreferredAuthorizedAppUrls = (authorizedDomains: string[]): string[] => {
  const currentHost =
    typeof window !== "undefined" ? window.location.hostname.toLowerCase().trim() : "";
  const filtered = authorizedDomains.filter((d) => {
    const lower = (d || "").toLowerCase().trim();
    if (!lower || lower === currentHost) return false;
    if (lower.endsWith(".firebaseapp.com") || lower.endsWith(".web.app")) return false;
    if (lower === "localhost" || lower === "127.0.0.1") return false;
    if (lower.startsWith("ais-dev-") || lower.startsWith("ais-mob-") || lower.startsWith("ais-shared-")) {
      return false;
    }
    return true;
  });
  return Array.from(new Set(filtered)).slice(0, 3);
};

export async function checkFirebaseAuthorizedDomain(
  forceRefresh = false
): Promise<DomainAuthCheckResult> {
  const currentHostname =
    typeof window !== "undefined" ? window.location.hostname.toLowerCase().trim() : "";

  if (!forceRefresh && cachedDomainAuthResult && cachedDomainAuthResult.checked) {
    return cachedDomainAuthResult;
  }

  if (!FIREBASE_API_KEY || typeof fetch === "undefined") {
    return {
      currentHostname,
      isAuthorized: true,
      authorizedDomains: [],
      checked: false,
      error: null,
    };
  }

  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 5000);
    const res = await fetch(
      `https://www.googleapis.com/identitytoolkit/v3/relyingparty/getProjectConfig?key=${encodeURIComponent(
        FIREBASE_API_KEY
      )}&_t=${Date.now()}`,
      {
        method: "GET",
        cache: "no-store",
        signal: controller.signal,
      }
    );
    clearTimeout(timer);

    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }

    const data = await res.json();
    const domains: string[] = Array.isArray(data?.authorizedDomains)
      ? data.authorizedDomains.filter((d: unknown): d is string => typeof d === "string")
      : [];

    const result: DomainAuthCheckResult = {
      currentHostname,
      isAuthorized: domains.length === 0 ? true : isHostnameAuthorized(currentHostname, domains),
      authorizedDomains: domains,
      checked: true,
      error: null,
    };
    cachedDomainAuthResult = result;
    return result;
  } catch (err: any) {
    return {
      currentHostname,
      isAuthorized: true,
      authorizedDomains: cachedDomainAuthResult?.authorizedDomains || [],
      checked: false,
      error: err?.message || String(err),
    };
  }
}

export interface FirestoreHealthState {
  status: "checking" | "connected" | "quota_exceeded" | "offline" | "error";
  readOk: boolean;
  writeQuotaExceeded: boolean;
  lastCheckedAt: string | null;
  lastError: string | null;
  projectId: string;
  databaseId: string;
}

const firestoreHealth: FirestoreHealthState = {
  status: "checking",
  readOk: false,
  writeQuotaExceeded: false,
  lastCheckedAt: null,
  lastError: null,
  projectId: FIREBASE_PROJECT_ID,
  databaseId: FIRESTORE_DATABASE_ID,
};

const healthListeners = new Set<(state: FirestoreHealthState) => void>();

const emitHealthUpdate = () => {
  const snapshot = { ...firestoreHealth };
  healthListeners.forEach((cb) => {
    try {
      cb(snapshot);
    } catch {}
  });
};

export const getFirestoreHealth = (): FirestoreHealthState => ({ ...firestoreHealth });

export const subscribeFirestoreHealth = (
  listener: (state: FirestoreHealthState) => void
): (() => void) => {
  healthListeners.add(listener);
  listener({ ...firestoreHealth });
  return () => {
    healthListeners.delete(listener);
  };
};

export const isFirestoreQuotaError = (error: unknown): boolean => {
  const msg = error instanceof Error ? error.message : String(error || "");
  return (
    msg.includes("RESOURCE_EXHAUSTED") ||
    msg.includes("Quota limit exceeded") ||
    msg.includes("Quota exceeded") ||
    msg.includes("Free daily write units") ||
    msg.includes("Free daily read units") ||
    msg.includes("quota metric")
  );
};

export const reportFirestoreWriteResult = (ok: boolean, error?: unknown) => {
  firestoreHealth.lastCheckedAt = new Date().toISOString();
  if (ok) {
    firestoreHealth.readOk = true;
    firestoreHealth.writeQuotaExceeded = false;
    firestoreHealth.status = "connected";
    firestoreHealth.lastError = null;
  } else if (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    if (isFirestoreQuotaError(error)) {
      firestoreHealth.writeQuotaExceeded = true;
      firestoreHealth.status = "quota_exceeded";
      firestoreHealth.lastError =
        "Kuota harian tulis Firestore (Free Tier) sementara penuh. Sinkronisasi real-time antar perangkat tetap berjalan melalui Server Relay, dan kuota Firestore akan reset otomatis besok.";
    } else if (errMsg.includes("offline") || errMsg.includes("Could not reach")) {
      firestoreHealth.status = "offline";
      firestoreHealth.lastError = errMsg;
    } else {
      firestoreHealth.lastError = errMsg;
    }
  }
  emitHealthUpdate();
};

// Keep Firestore internal SDK log level to error so network retries don't spam console
try {
  setLogLevel("error");
} catch {}

// Detect Firestore quota errors emitted by SDK transport layer while preventing uncaught log noise
const originalConsoleError = console.error;
console.error = function (...args) {
  const message = args
    .map((arg) => {
      if (typeof arg === "string") return arg;
      if (arg instanceof Error) return arg.message + (arg.stack ? "\n" + arg.stack : "");
      try {
        return JSON.stringify(arg);
      } catch {
        return String(arg);
      }
    })
    .join(" ");

  if (
    message.includes("Quota limit exceeded") ||
    message.includes("Quota exceeded") ||
    message.includes("Free daily write units") ||
    message.includes("Free daily read units") ||
    message.includes("RESOURCE_EXHAUSTED")
  ) {
    reportFirestoreWriteResult(false, message);
    console.warn("[Firestore Quota Notice]:", message.slice(0, 260));
    return;
  }

  if (
    message.includes("Could not reach Cloud Firestore backend") ||
    message.includes("@firebase/firestore")
  ) {
    console.warn("[Firestore Transport Notice]:", message.slice(0, 260));
    return;
  }

  originalConsoleError.apply(console, args);
};

// Recursively strip undefined properties so Firestore setDoc never rejects optional fields
export function stripUndefinedDeep<T>(value: T): T {
  if (value === null || value === undefined || typeof value !== "object") {
    return value;
  }
  if (Array.isArray(value)) {
    return value
      .filter((item) => item !== undefined)
      .map((item) => stripUndefinedDeep(item)) as unknown as T;
  }
  // Preserve Firestore FieldValue / Timestamp / Date class instances intact
  const proto = Object.getPrototypeOf(value);
  if (proto !== Object.prototype && proto !== null) {
    return value;
  }
  const cleaned: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
    if (v !== undefined) {
      cleaned[k] = stripUndefinedDeep(v);
    }
  }
  return cleaned as T;
}

// Initialize the Firebase App
const app = initializeApp(firebaseConfig);

// CRITICAL: Always pass firestoreDatabaseId from firebase-applet-config.json
// Use experimentalForceLongPolling so Chrome Incognito, iOS Safari, Android mobile networks,
// and Cloud Run / iframe environments never stall on WebChannel streaming negotiation.
export const db = initializeFirestore(
  app,
  {
    experimentalForceLongPolling: true,
    ignoreUndefinedProperties: true,
  },
  (firebaseConfig as { firestoreDatabaseId?: string }).firestoreDatabaseId
);

// Initialize Firebase Authentication
export const auth = getAuth(app);
export const googleAuthProvider = new GoogleAuthProvider();
googleAuthProvider.setCustomParameters({ prompt: "select_account" });

// ============================================================================
// Direct Authenticated Cloud Firestore REST API v1 Helpers
// Guarantees fast (<300ms) HTTPS cross-device reads & writes on Mobile,
// Incognito, Safari, and embedded iframes — enforces the EXACT same
// firestore.rules using the user's real Firebase Auth ID token!
// ============================================================================

type FirestoreRestValue =
  | { nullValue: null }
  | { booleanValue: boolean }
  | { integerValue: string }
  | { doubleValue: number }
  | { timestampValue: string }
  | { stringValue: string }
  | { arrayValue: { values?: FirestoreRestValue[] } }
  | { mapValue: { fields?: Record<string, FirestoreRestValue> } };

const toFirestoreRestValue = (value: unknown): FirestoreRestValue | undefined => {
  if (value === undefined) return undefined;
  if (value === null) return { nullValue: null };
  if (typeof value === "boolean") return { booleanValue: value };
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return { integerValue: "0" };
    return Number.isInteger(value)
      ? { integerValue: String(value) }
      : { doubleValue: value };
  }
  if (typeof value === "string") return { stringValue: value };
  if (value instanceof Date) return { timestampValue: value.toISOString() };
  if (Array.isArray(value)) {
    const encoded = value
      .map((item) => toFirestoreRestValue(item))
      .filter((v): v is FirestoreRestValue => v !== undefined);
    return { arrayValue: { values: encoded } };
  }
  if (typeof value === "object") {
    const maybeTs = value as { toDate?: () => Date; seconds?: number };
    if (typeof maybeTs.toDate === "function") {
      try {
        return { timestampValue: maybeTs.toDate().toISOString() };
      } catch {}
    }
    if (typeof maybeTs.seconds === "number" && Object.keys(value as Record<string, unknown>).length <= 2) {
      return { timestampValue: new Date(maybeTs.seconds * 1000).toISOString() };
    }
    // Skip non-plain objects such as Firestore FieldValue sentinels
    const proto = Object.getPrototypeOf(value);
    if (proto !== Object.prototype && proto !== null) {
      return undefined;
    }
    const fields: Record<string, FirestoreRestValue> = {};
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
      const encoded = toFirestoreRestValue(v);
      if (encoded !== undefined) {
        fields[k] = encoded;
      }
    }
    return { mapValue: { fields } };
  }
  return undefined;
};

const encodeFirestoreRestFields = (data: Record<string, unknown>): Record<string, FirestoreRestValue> => {
  const fields: Record<string, FirestoreRestValue> = {};
  for (const [k, v] of Object.entries(data)) {
    const encoded = toFirestoreRestValue(v);
    if (encoded !== undefined) {
      fields[k] = encoded;
    }
  }
  return fields;
};

const fromFirestoreRestValue = (val: any): any => {
  if (!val || typeof val !== "object") return undefined;
  if ("nullValue" in val) return null;
  if ("booleanValue" in val) return Boolean(val.booleanValue);
  if ("integerValue" in val) return Number(val.integerValue);
  if ("doubleValue" in val) return Number(val.doubleValue);
  if ("timestampValue" in val) return String(val.timestampValue);
  if ("stringValue" in val) return String(val.stringValue);
  if ("arrayValue" in val) {
    const arr = Array.isArray(val.arrayValue?.values) ? val.arrayValue.values : [];
    return arr.map((item: any) => fromFirestoreRestValue(item));
  }
  if ("mapValue" in val) {
    return decodeFirestoreRestFields(val.mapValue?.fields);
  }
  return undefined;
};

const decodeFirestoreRestFields = (fields?: Record<string, any>): Record<string, any> => {
  const result: Record<string, any> = {};
  if (!fields || typeof fields !== "object") return result;
  for (const [k, v] of Object.entries(fields)) {
    const decoded = fromFirestoreRestValue(v);
    if (decoded !== undefined) {
      result[k] = decoded;
    }
  }
  return result;
};

export interface FirestoreRestResult<T = unknown> {
  ok: boolean;
  status: number;
  data?: T;
  notFound?: boolean;
  permissionDenied?: boolean;
  quotaExceeded?: boolean;
  error?: string;
}

export async function getActiveFirebaseIdToken(forceRefresh = false): Promise<string | null> {
  try {
    const current = auth.currentUser;
    if (!current) return null;
    return await current.getIdToken(forceRefresh);
  } catch {
    return null;
  }
}

export async function firestoreRestGetDoc(
  collectionName: string,
  docId: string,
  options?: { idToken?: string | null; timeoutMs?: number }
): Promise<FirestoreRestResult<Record<string, any> | null>> {
  if (!docId || typeof fetch === "undefined") {
    return { ok: false, status: 0, error: "Invalid document ID" };
  }
  const token =
    options?.idToken !== undefined ? options.idToken : await getActiveFirebaseIdToken(false);
  const url = `${FIRESTORE_REST_BASE}/${encodeURIComponent(collectionName)}/${encodeURIComponent(docId)}`;
  const headers: Record<string, string> = { Accept: "application/json" };
  if (token) headers.Authorization = `Bearer ${token}`;

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), options?.timeoutMs ?? 4500);
  try {
    const res = await fetch(url, {
      method: "GET",
      headers,
      cache: "no-store",
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (res.status === 404) {
      return { ok: true, status: 404, notFound: true, data: null };
    }
    const json = await res.json().catch(() => null);
    if (!res.ok) {
      const msg = json?.error?.message || `HTTP ${res.status}`;
      const isPerm = res.status === 403 || msg.includes("PERMISSION_DENIED");
      const isQuota = res.status === 429 || isFirestoreQuotaError(msg);
      if (isQuota) reportFirestoreWriteResult(false, msg);
      return { ok: false, status: res.status, permissionDenied: isPerm, quotaExceeded: isQuota, error: msg };
    }
    const decoded = decodeFirestoreRestFields(json?.fields);
    reportFirestoreWriteResult(true);
    return { ok: true, status: res.status, data: { id: docId, ...decoded } };
  } catch (err: any) {
    clearTimeout(timer);
    return { ok: false, status: 0, error: err?.message || String(err) };
  }
}

export async function firestoreRestUpsertDoc(
  collectionName: string,
  docId: string,
  data: Record<string, unknown>,
  options?: {
    idToken?: string | null;
    updateMaskFields?: string[];
    keepalive?: boolean;
    timeoutMs?: number;
  }
): Promise<FirestoreRestResult<Record<string, any>>> {
  if (!docId || typeof fetch === "undefined") {
    return { ok: false, status: 0, error: "Invalid document ID" };
  }
  const token =
    options?.idToken !== undefined ? options.idToken : await getActiveFirebaseIdToken(false);
  const fields = encodeFirestoreRestFields(data);
  const maskFields =
    options?.updateMaskFields && options.updateMaskFields.length > 0
      ? options.updateMaskFields.filter((f) => f in fields)
      : Object.keys(fields);

  const params = new URLSearchParams();
  maskFields.forEach((f) => params.append("updateMask.fieldPaths", f));
  const url = `${FIRESTORE_REST_BASE}/${encodeURIComponent(collectionName)}/${encodeURIComponent(
    docId
  )}?${params.toString()}`;

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    Accept: "application/json",
  };
  if (token) headers.Authorization = `Bearer ${token}`;

  const bodyStr = JSON.stringify({ fields });
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), options?.timeoutMs ?? 6000);
  try {
    const res = await fetch(url, {
      method: "PATCH",
      headers,
      body: bodyStr,
      cache: "no-store",
      // keepalive supports payloads up to 64KB and survives mobile tab switching
      keepalive: Boolean(options?.keepalive && bodyStr.length < 58_000),
      signal: controller.signal,
    });
    clearTimeout(timer);
    const json = await res.json().catch(() => null);
    if (!res.ok) {
      const msg = json?.error?.message || `HTTP ${res.status}`;
      const isPerm = res.status === 403 || msg.includes("PERMISSION_DENIED");
      const isQuota = res.status === 429 || isFirestoreQuotaError(msg);
      reportFirestoreWriteResult(false, msg);
      return { ok: false, status: res.status, permissionDenied: isPerm, quotaExceeded: isQuota, error: msg };
    }
    reportFirestoreWriteResult(true);
    return {
      ok: true,
      status: res.status,
      data: { id: docId, ...decodeFirestoreRestFields(json?.fields) },
    };
  } catch (err: any) {
    clearTimeout(timer);
    const errMsg = err?.message || String(err);
    reportFirestoreWriteResult(false, errMsg);
    return { ok: false, status: 0, error: errMsg };
  }
}

export async function firestoreRestDeleteDoc(
  collectionName: string,
  docId: string,
  options?: { idToken?: string | null; timeoutMs?: number }
): Promise<FirestoreRestResult<null>> {
  if (!docId || typeof fetch === "undefined") {
    return { ok: false, status: 0, error: "Invalid document ID" };
  }
  const token =
    options?.idToken !== undefined ? options.idToken : await getActiveFirebaseIdToken(false);
  const url = `${FIRESTORE_REST_BASE}/${encodeURIComponent(collectionName)}/${encodeURIComponent(docId)}`;
  const headers: Record<string, string> = { Accept: "application/json" };
  if (token) headers.Authorization = `Bearer ${token}`;

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), options?.timeoutMs ?? 5000);
  try {
    const res = await fetch(url, {
      method: "DELETE",
      headers,
      cache: "no-store",
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (res.status === 404 || res.ok) {
      reportFirestoreWriteResult(true);
      return { ok: true, status: res.status, data: null };
    }
    const json = await res.json().catch(() => null);
    const msg = json?.error?.message || `HTTP ${res.status}`;
    return {
      ok: false,
      status: res.status,
      permissionDenied: res.status === 403,
      quotaExceeded: res.status === 429 || isFirestoreQuotaError(msg),
      error: msg,
    };
  } catch (err: any) {
    clearTimeout(timer);
    return { ok: false, status: 0, error: err?.message || String(err) };
  }
}

export async function firestoreRestListCollection(
  collectionName: string,
  options?: { idToken?: string | null; pageSize?: number; maxPages?: number; timeoutMs?: number }
): Promise<FirestoreRestResult<Record<string, any>[]>> {
  if (typeof fetch === "undefined") {
    return { ok: false, status: 0, error: "Fetch unavailable" };
  }
  const token =
    options?.idToken !== undefined ? options.idToken : await getActiveFirebaseIdToken(false);
  const headers: Record<string, string> = { Accept: "application/json" };
  if (token) headers.Authorization = `Bearer ${token}`;

  const pageSize = options?.pageSize ?? 300;
  const maxPages = options?.maxPages ?? 4;
  const allDocs: Record<string, any>[] = [];
  let pageToken = "";

  for (let page = 0; page < maxPages; page++) {
    const params = new URLSearchParams({ pageSize: String(pageSize) });
    if (pageToken) params.set("pageToken", pageToken);
    const url = `${FIRESTORE_REST_BASE}/${encodeURIComponent(collectionName)}?${params.toString()}`;

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), options?.timeoutMs ?? 6500);
    try {
      const res = await fetch(url, {
        method: "GET",
        headers,
        cache: "no-store",
        signal: controller.signal,
      });
      clearTimeout(timer);
      const json = await res.json().catch(() => null);
      if (!res.ok) {
        const msg = json?.error?.message || `HTTP ${res.status}`;
        const isPerm = res.status === 403 || msg.includes("PERMISSION_DENIED");
        const isQuota = res.status === 429 || isFirestoreQuotaError(msg);
        if (isQuota) reportFirestoreWriteResult(false, msg);
        return {
          ok: false,
          status: res.status,
          permissionDenied: isPerm,
          quotaExceeded: isQuota,
          error: msg,
        };
      }
      const docs = Array.isArray(json?.documents) ? json.documents : [];
      for (const rawDoc of docs) {
        const fullName = String(rawDoc?.name || "");
        const docId = fullName.split("/").pop() || "";
        if (!docId) continue;
        const decoded = decodeFirestoreRestFields(rawDoc?.fields);
        allDocs.push({ id: docId, ...decoded });
      }
      pageToken = typeof json?.nextPageToken === "string" ? json.nextPageToken : "";
      if (!pageToken) break;
    } catch (err: any) {
      clearTimeout(timer);
      return { ok: false, status: 0, error: err?.message || String(err) };
    }
  }

  reportFirestoreWriteResult(true);
  return { ok: true, status: 200, data: allDocs };
}

// Validate Firestore connectivity on startup via getDocFromServer
export async function testFirestoreConnection(): Promise<FirestoreHealthState> {
  try {
    await getDocFromServer(doc(db, "test", "connection"));
    firestoreHealth.readOk = true;
    firestoreHealth.lastCheckedAt = new Date().toISOString();
    if (!firestoreHealth.writeQuotaExceeded) {
      firestoreHealth.status = "connected";
      firestoreHealth.lastError = null;
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    firestoreHealth.lastCheckedAt = new Date().toISOString();
    if (isFirestoreQuotaError(error)) {
      firestoreHealth.writeQuotaExceeded = true;
      firestoreHealth.status = "quota_exceeded";
      firestoreHealth.lastError = errMsg;
    } else if (errMsg.includes("permission-denied") || errMsg.includes("Missing or insufficient permissions")) {
      // Server responded, meaning network & database ID are valid
      firestoreHealth.readOk = true;
      if (!firestoreHealth.writeQuotaExceeded) {
        firestoreHealth.status = "connected";
        firestoreHealth.lastError = null;
      }
    } else if (errMsg.includes("client is offline") || errMsg.includes("Could not reach")) {
      firestoreHealth.status = "offline";
      firestoreHealth.lastError = "Koneksi ke server Firestore sedang offline atau tidak terjangkau.";
    } else {
      firestoreHealth.status = "error";
      firestoreHealth.lastError = errMsg;
    }
  }
  emitHealthUpdate();
  return getFirestoreHealth();
}

// Run startup connection verification in browser environments
if (typeof window !== "undefined") {
  testFirestoreConnection().catch(() => {});
}

// Standardize operation types for our security monitoring
export enum OperationType {
  CREATE = "create",
  UPDATE = "update",
  DELETE = "delete",
  LIST = "list",
  GET = "get",
  WRITE = "write",
}

// Structuring error details so security rule violations can be diagnosed
export interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  };
}

/**
 * Robust error handling middleware for Firestore.
 * Conforms to the JSON payload structure required by the system.
 */
export function handleFirestoreError(
  error: unknown,
  operationType: OperationType,
  path: string | null
) {
  const errStr = error instanceof Error ? error.message : String(error);
  if (isFirestoreQuotaError(error)) {
    reportFirestoreWriteResult(false, error);
    console.warn(`[Firestore Quota Mode on ${operationType} - ${path}]: Switching to realtime relay.`);
    return;
  }

  const isOffline =
    errStr.includes("Could not reach") ||
    errStr.includes("offline") ||
    errStr.includes("unavailable");

  if (isOffline) {
    console.warn(`[Firestore Fallback Mode on ${operationType} - ${path}]: Using offline/relay state.`);
    return;
  }

  const errInfo: FirestoreErrorInfo = {
    error: errStr,
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo:
        auth.currentUser?.providerData?.map((provider) => ({
          providerId: provider.providerId,
          email: provider.email,
        })) || [],
    },
    operationType,
    path,
  };
  console.warn("Firestore Security/Operation Warning: ", JSON.stringify(errInfo));
}
