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
export const FIRESTORE_DATABASE_ID =
  (firebaseConfig as { firestoreDatabaseId?: string }).firestoreDatabaseId || "(default)";

export const FIRESTORE_UPGRADE_URL = `https://console.firebase.google.com/project/${FIREBASE_PROJECT_ID}/firestore/databases/${encodeURIComponent(
  FIRESTORE_DATABASE_ID
)}/data?openUpgradeDialog=true`;

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

// Initialize the Firebase App
const app = initializeApp(firebaseConfig);

// CRITICAL: Always pass firestoreDatabaseId from firebase-applet-config.json
// Use experimentalAutoDetectLongPolling so mobile browsers (iOS Safari / Android Chrome)
// and desktop browsers negotiate fast WebChannel/HTTP2 with automatic long-polling fallback.
export const db = initializeFirestore(
  app,
  {
    experimentalAutoDetectLongPolling: true,
  },
  (firebaseConfig as { firestoreDatabaseId?: string }).firestoreDatabaseId
);

// Initialize Firebase Authentication
export const auth = getAuth(app);
export const googleAuthProvider = new GoogleAuthProvider();
googleAuthProvider.setCustomParameters({ prompt: "select_account" });

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
