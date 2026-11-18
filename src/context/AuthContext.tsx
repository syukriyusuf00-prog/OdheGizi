import React, { createContext, useContext, useState, useEffect } from "react";
import { 
  User, 
  signInWithPopup, 
  signOut, 
  onAuthStateChanged 
} from "firebase/auth";
import { 
  doc, 
  getDoc, 
  setDoc,
  onSnapshot,
  serverTimestamp,
  collection,
  query,
  where,
  getDocs
} from "firebase/firestore";
import {
  auth,
  googleAuthProvider,
  db,
  handleFirestoreError,
  OperationType,
  reportFirestoreWriteResult,
  checkFirebaseAuthorizedDomain,
  DomainAuthCheckResult,
  stripUndefinedDeep,
} from "../lib/firebase.ts";
import { clearDriveSession, getIsDriveConnecting } from "../lib/googleDrive.ts";
import {
  clearSheetsSession,
  getIsSheetsConnecting,
  loadGlobalSheetSyncConfig,
} from "../lib/googleSheetsSync.ts";
import {
  buildDeviceRecord,
  formatDeviceShortLabel,
  syncDeviceToFirestore,
  toCompactDeviceSummary,
} from "../lib/deviceTracker.ts";

const lookupEmailInLinkedSpreadsheet = async (email: string): Promise<Record<string, any> | null> => {
  const cleanEmail = (email || "").toLowerCase().trim();
  if (!cleanEmail) return null;
  try {
    // Ensure server relay knows global spreadsheet config (in case server just restarted)
    await loadGlobalSheetSyncConfig().catch(() => null);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 5000);
    const res = await fetch(
      `/api/sheets/lookup-email?email=${encodeURIComponent(cleanEmail)}`,
      {
        method: "GET",
        cache: "no-store",
        signal: controller.signal,
      }
    );
    clearTimeout(timer);
    if (!res.ok) return null;
    const data = await res.json();
    if (data?.found && data?.sheetUser) {
      return data.sheetUser;
    }
  } catch {}
  return null;
};

// Quota protection: deduplicate identical state writes so opening/refreshing never burns Firestore write units
const lastSyncedCloudPayload = new Map<string, string>();
// Concurrency protection: deduplicate simultaneous fetchUserProfile calls from signInWithPopup + onAuthStateChanged
const inFlightProfileFetchMap = new Map<string, Promise<void>>();
const recentProfileFetchAt = new Map<string, number>();

const clampString = (val: unknown, maxLen: number): string => {
  if (typeof val !== "string") return "";
  return val.trim().slice(0, maxLen);
};

const sanitizeUserProfileForFirestore = (raw: Record<string, any>): Record<string, any> => {
  const copy: Record<string, any> = { ...raw };
  delete copy.isNotRegistered;
  delete copy._allIds;
  const countNum = Number(copy.loginCount);
  return stripUndefinedDeep({
    ...copy,
    uid: clampString(copy.uid || copy.id, 128),
    id: clampString(copy.id || copy.uid, 128),
    email: clampString(copy.email, 256).toLowerCase(),
    namaLengkap: clampString(copy.namaLengkap, 160),
    profesi: clampString(copy.profesi, 120),
    namaSPPG: clampString(copy.namaSPPG, 160),
    noHp: clampString(copy.noHp, 40),
    peran: copy.peran === "ADMIN" ? "ADMIN" : "USER",
    statusPersetujuan:
      copy.statusPersetujuan === "aktif" || copy.statusPersetujuan === "diblokir"
        ? copy.statusPersetujuan
        : "menunggu",
    mulaiPada: clampString(copy.mulaiPada ?? "", 32),
    berakhirPada: clampString(copy.berakhirPada ?? "", 32),
    keterangan: clampString(copy.keterangan ?? "", 300),
    sumber: clampString(copy.sumber || "Login Google", 64),
    lastDeviceId: copy.lastDeviceId ? clampString(copy.lastDeviceId, 64) : undefined,
    lastDeviceLabel: copy.lastDeviceLabel ? clampString(copy.lastDeviceLabel, 160) : undefined,
    loginCount: Number.isFinite(countNum) && countNum >= 0 ? Math.min(1000000, Math.floor(countNum)) : 1,
  });
};

const persistUserProfileToFirestore = async (
  userId: string,
  profile: Record<string, any>,
  existingFirestoreDoc?: Record<string, any> | null,
  options?: { isCreate?: boolean }
): Promise<boolean> => {
  if (!userId || userId === "demo_user_5min") return false;
  const userRef = doc(db, "users", userId);
  const baseSanitized = sanitizeUserProfileForFirestore({
    ...profile,
    uid: userId,
    id: userId,
  });
  const writePayload = {
    ...baseSanitized,
    createdAt: existingFirestoreDoc?.createdAt || baseSanitized.createdAt || serverTimestamp(),
    updatedAt: serverTimestamp(),
    loginTerakhir: serverTimestamp(),
  };

  const attemptWrite = async (payload: Record<string, any>): Promise<void> => {
    await setDoc(userRef, stripUndefinedDeep(payload), { merge: true });
    reportFirestoreWriteResult(true);
  };

  try {
    await attemptWrite(writePayload);
    return true;
  } catch (err: any) {
    const errCode = String(err?.code || "");
    const errMsg = String(err?.message || err || "");

    // If permission-denied (e.g. sheet status 'aktif' without Firestore sheet_user_* doc, or legacy field mismatch),
    // retry with a strictly rules-compliant payload so the user's email is STILL registered in Firestore for Admin Panel!
    if (errCode === "permission-denied" || errMsg.includes("PERMISSION_DENIED") || errMsg.includes("Missing or insufficient permissions")) {
      try {
        const isMainAdmin = isMainAdminEmail(baseSanitized.email);
        const safeFallback: Record<string, any> = existingFirestoreDoc
          ? {
              ...baseSanitized,
              peran: existingFirestoreDoc.peran === "ADMIN" ? "ADMIN" : "USER",
              statusPersetujuan:
                existingFirestoreDoc.statusPersetujuan === "aktif" ||
                existingFirestoreDoc.statusPersetujuan === "diblokir"
                  ? existingFirestoreDoc.statusPersetujuan
                  : "menunggu",
              mulaiPada:
                typeof existingFirestoreDoc.mulaiPada === "string"
                  ? existingFirestoreDoc.mulaiPada
                  : "",
              berakhirPada:
                typeof existingFirestoreDoc.berakhirPada === "string"
                  ? existingFirestoreDoc.berakhirPada
                  : "",
              keterangan:
                typeof existingFirestoreDoc.keterangan === "string"
                  ? existingFirestoreDoc.keterangan
                  : "",
              createdAt: existingFirestoreDoc.createdAt || serverTimestamp(),
              updatedAt: serverTimestamp(),
              loginTerakhir: serverTimestamp(),
            }
          : {
              ...baseSanitized,
              peran: isMainAdmin ? "ADMIN" : "USER",
              statusPersetujuan: isMainAdmin ? "aktif" : "menunggu",
              mulaiPada: "",
              berakhirPada: "",
              keterangan: isMainAdmin ? "Administrator Utama" : DEFAULT_WAITING_NOTE,
              createdAt: serverTimestamp(),
              updatedAt: serverTimestamp(),
              loginTerakhir: serverTimestamp(),
            };
        delete safeFallback.migratedFromDocId;
        await attemptWrite(safeFallback);
        return true;
      } catch (retryErr) {
        reportFirestoreWriteResult(false, retryErr);
        handleFirestoreError(
          retryErr,
          options?.isCreate ? OperationType.CREATE : OperationType.WRITE,
          `users/${userId}`
        );
        return false;
      }
    }

    reportFirestoreWriteResult(false, err);

    // Retry once automatically after 2.5s for transient mobile network / transport cold starts
    const isTransient =
      errCode === "unavailable" ||
      errCode === "deadline-exceeded" ||
      errMsg.includes("offline") ||
      errMsg.includes("Could not reach") ||
      errMsg.includes("network");

    if (isTransient) {
      window.setTimeout(() => {
        attemptWrite(writePayload).catch((bgErr) => {
          reportFirestoreWriteResult(false, bgErr);
        });
      }, 2500);
    } else {
      handleFirestoreError(
        err,
        options?.isCreate ? OperationType.CREATE : OperationType.WRITE,
        `users/${userId}`
      );
    }
    return false;
  }
};

const normalizeTimestampForRelay = (val: any): string | null => {
  if (!val) return null;
  if (typeof val === "string") return val;
  if (typeof val?.seconds === "number") return new Date(val.seconds * 1000).toISOString();
  if (typeof val?.toDate === "function") return val.toDate().toISOString();
  return new Date().toISOString();
};

export const syncUserToRealtimeHub = async (profileData: Record<string, any> | null) => {
  if (!profileData || !profileData.email) return;
  try {
    const payload = {
      ...profileData,
      createdAt: normalizeTimestampForRelay(profileData.createdAt) || new Date().toISOString(),
      loginPertama:
        normalizeTimestampForRelay(profileData.loginPertama) ||
        normalizeTimestampForRelay(profileData.createdAt) ||
        new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      loginTerakhir:
        profileData.loginTerakhir === null
          ? null
          : normalizeTimestampForRelay(profileData.loginTerakhir),
    };
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 4500);
    await fetch("/api/realtime/sync-user", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      keepalive: true,
      signal: controller.signal,
    });
    clearTimeout(timer);
  } catch (err) {
    console.warn("Notice pushing user to realtime hub:", err);
  }
};

export const deleteUserFromRealtimeHub = async (uid?: string | null, email?: string | null) => {
  if (!uid && !email) return;
  try {
    await fetch("/api/realtime/delete-user", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ uid: uid || "", email: email || "" }),
      keepalive: true,
    });
  } catch {}
};

export const fetchUserFromRealtimeHub = async (email: string, uid?: string): Promise<any | null> => {
  try {
    const params = new URLSearchParams();
    if (email) params.set("email", email.toLowerCase().trim());
    if (uid) params.set("uid", uid);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 3500);
    const resp = await fetch(`/api/realtime/user?${params.toString()}`, {
      method: "GET",
      headers: { Accept: "application/json" },
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (resp.ok) {
      const json = await resp.json().catch(() => null);
      return json?.user || null;
    }
  } catch {}
  return null;
};

const DEMO_DURATION_MS = 5 * 60 * 1000;
const DEMO_UID = "demo_user_5min";
const DEMO_EMAIL = "demo@sispermbg.id";
const DEMO_EXPIRES_KEY = "sisper_demo_expires_at";
export const PENDING_OTP_STORAGE_KEY = "sisper_pending_otp_session";

const PRESERVED_SISPER_KEYS = new Set([
  "sisper_user_profile",
  "sisper_global_app_logo",
  "sisper_global_app_logo_right",
  "sisper_global_app_logo_version",
  "sisper_global_admin_tkpi",
  "sisper_demo_expires_at",
  "sisper_demo_just_logged_in",
  "sisper_device_id",
  "sisper_device_first_seen",
  "sisper_current_device_info",
]);

export const clearPerUserWorksheetStorage = () => {
  try {
    Object.keys(localStorage).forEach((key) => {
      if (PRESERVED_SISPER_KEYS.has(key)) return;
      if (
        key.startsWith("sisper_") ||
        key.startsWith("kop_") ||
        key.startsWith("sppg_") ||
        key.startsWith("fc_header_") ||
        key.startsWith("nota_") ||
        key.startsWith("logo_")
      ) {
        localStorage.removeItem(key);
      }
    });
    sessionStorage.removeItem("sisper_user_tab_chosen");
  } catch (e) {
    console.warn("Notice clearing per-user worksheet storage:", e);
  }
};

const buildDemoProfile = () => ({
  uid: DEMO_UID,
  email: DEMO_EMAIL,
  namaLengkap: "Pengguna Demo (5 Menit)",
  profesi: "Ahli Gizi (Mode Demo)",
  namaSPPG: "SPPG Demo Perencanaan Gizi",
  noHp: "-",
  peran: "DEMO" as const,
  isDemo: true,
  statusPersetujuan: "aktif" as const,
  berakhirPada: "2099-12-31",
});

export interface OtpDispatchResult {
  ok: boolean;
  deliveryMode: "email" | "preview";
  provider: "resend" | "smtp" | "preview";
  recipientEmail: string;
  verificationToken: string;
  expiresAt: number;
  expiresInSeconds: number;
  message: string;
  emailSubject?: string;
  previewCode?: string;
  smtpWarning?: string | null;
  uid: string;
  namaLengkap: string;
  namaSPPG: string;
}

interface AuthContextType {
  user: User | null;
  userProfile: any | null;
  loading: boolean;
  signInWithGoogle: () => Promise<void>;
  signOutUser: () => Promise<void>;
  isCloudActive: boolean;
  saveStateToCloud: (key: string, data: any) => Promise<void>;
  loadStateFromCloud: (key: string) => Promise<any | null>;
  registerUser: (formData: { namaLengkap: string; profesi: string; namaSPPG: string; noHp: string }) => Promise<void>;
  updateOwnProfileDetails: (formData: { namaLengkap: string; profesi: string; namaSPPG: string; noHp: string }) => Promise<void>;
  registerCustomUser: (formData: {
    namaLengkap: string;
    email: string;
    profesi: string;
    namaSPPG: string;
    sandi: string;
  }) => Promise<OtpDispatchResult>;
  resendVerificationCode: (params: {
    email: string;
    namaLengkap?: string;
    namaSPPG?: string;
    uid?: string;
  }) => Promise<OtpDispatchResult>;
  verifyRegistrationCode: (params: {
    email: string;
    code: string;
    verificationToken: string;
    autoLogin?: boolean;
  }) => Promise<any>;
  registerThreeRoles: (data: {
    namaSPPG: string;
    sandi: string;
    roles: Array<{ namaLengkap: string; email: string; noHp: string; profesi: string }>;
  }) => Promise<void>;
  loginWithEmailPassword: (email: string, sandi?: string) => Promise<void>;
  resetUserPassword: (email: string, newSandi: string) => Promise<void>;
  refreshUserProfile: () => Promise<any | null>;
  authError: string | null;
  setAuthError: (error: string | null) => void;
  domainAuthInfo: DomainAuthCheckResult;
  recheckDomainAndSignIn: () => Promise<boolean>;
  simulateAdminLogin: () => void;
  isDemo: boolean;
  demoSecondsLeft: number | null;
  demoNotice: string | null;
  clearDemoNotice: () => void;
  loginDemoAccount: () => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const isMainAdminEmail = (email?: string | null) => {
  if (!email) return false;
  const lower = email.toLowerCase().trim();
  return lower === "syukriyusuf82@gmail.com" || lower === "sukriyusuf82@gmail.com" || lower === "syukriyusuf00@gmail.com";
};

const fetchWithTimeout = <T,>(promise: Promise<T>, timeoutMs: number, defaultValue: T): Promise<T> => {
  return Promise.race([
    promise,
    new Promise<T>((resolve) => setTimeout(() => resolve(defaultValue), timeoutMs))
  ]);
};

const getRecordTimeMs = (val: any): number => {
  if (!val) return 0;
  if (typeof val === "string") {
    const t = new Date(val).getTime();
    return Number.isNaN(t) ? 0 : t;
  }
  if (typeof val?.seconds === "number") return val.seconds * 1000;
  if (typeof val?.toMillis === "function") return val.toMillis();
  return 0;
};

export const DEFAULT_WAITING_NOTE = "Daftar Tunggu - Menunggu ACC Admin Odhe Gizi";

const getDefaultExpirationDate = () => {
  const nextYear = new Date();
  nextYear.setFullYear(nextYear.getFullYear() + 1);
  return `${nextYear.getFullYear()}-${String(nextYear.getMonth() + 1).padStart(2, "0")}-${String(nextYear.getDate()).padStart(2, "0")}`;
};

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<User | null>(() => {
    const customUid = localStorage.getItem("custom_logged_in_uid");
    if (customUid === DEMO_UID) {
      const rawExp = localStorage.getItem(DEMO_EXPIRES_KEY);
      if (rawExp && Date.now() < Number(rawExp)) {
        const demoProf = buildDemoProfile();
        return {
          uid: DEMO_UID,
          email: DEMO_EMAIL,
          displayName: demoProf.namaLengkap,
          emailVerified: true,
        } as any;
      }
    }
    return null;
  });

  const [userProfile, setUserProfile] = useState<any | null>(() => {
    const customUid = localStorage.getItem("custom_logged_in_uid");
    if (customUid === DEMO_UID) {
      const rawExp = localStorage.getItem(DEMO_EXPIRES_KEY);
      if (rawExp && Date.now() < Number(rawExp)) {
        return buildDemoProfile();
      }
    }
    return null;
  });

  const [loading, setLoading] = useState(true);
  const [authError, setAuthError] = useState<string | null>(null);
  const [domainAuthInfo, setDomainAuthInfo] = useState<DomainAuthCheckResult>(() => ({
    currentHostname:
      typeof window !== "undefined" ? window.location.hostname.toLowerCase().trim() : "",
    isAuthorized: true,
    authorizedDomains: [],
    checked: false,
    error: null,
  }));
  const [demoNotice, setDemoNotice] = useState<string | null>(null);
  const [demoSecondsLeft, setDemoSecondsLeft] = useState<number | null>(() => {
    const rawExp = localStorage.getItem(DEMO_EXPIRES_KEY);
    if (!rawExp) return null;
    const rem = Math.ceil((Number(rawExp) - Date.now()) / 1000);
    return rem > 0 ? rem : 0;
  });

  // Proactively inspect Firebase Authorized Domains on the login screen
  useEffect(() => {
    let active = true;
    checkFirebaseAuthorizedDomain(false)
      .then((res) => {
        if (active) setDomainAuthInfo(res);
      })
      .catch(() => {});
    return () => {
      active = false;
    };
  }, []);

  const isDemo = Boolean(
    userProfile?.isDemo ||
    userProfile?.peran === "DEMO" ||
    user?.uid === DEMO_UID
  );

  const clearDemoNotice = () => setDemoNotice(null);

  const updateSession = (currentUser: User | null, profile: any | null) => {
    setUser(currentUser);
    setUserProfile(profile);
    if (currentUser && profile) {
      const isLocalCustomSession =
        currentUser.uid === DEMO_UID ||
        currentUser.uid.startsWith("custom_user_") ||
        currentUser.uid.startsWith("admin_");
      if (isLocalCustomSession) {
        localStorage.setItem("custom_logged_in_uid", currentUser.uid);
      } else {
        // Real Firebase Auth session persists natively via onAuthStateChanged
        localStorage.removeItem("custom_logged_in_uid");
      }
      localStorage.setItem("sisper_user_profile", JSON.stringify(profile));
      if (currentUser.uid !== DEMO_UID) {
        localStorage.setItem(`offline_user_${currentUser.uid}`, JSON.stringify(profile));
      }
    } else {
      localStorage.removeItem("custom_logged_in_uid");
      localStorage.removeItem("sisper_user_profile");
    }
  };

  const fetchUserProfile = async (currentUser: User, options?: { force?: boolean }) => {
    const uid = currentUser.uid;
    if (!uid) return;

    // Deduplicate concurrent signInWithPopup + onAuthStateChanged calls for the same UID
    const existingPromise = inFlightProfileFetchMap.get(uid);
    if (existingPromise) {
      return existingPromise;
    }
    const lastCompletedAt = recentProfileFetchAt.get(uid) || 0;
    if (!options?.force && Date.now() - lastCompletedAt < 1800) {
      return;
    }

    const runFetch = async () => {
      const lowerEmail = (currentUser.email || "").toLowerCase().trim();
      const isMainAdmin = isMainAdminEmail(lowerEmail);
      const userRef = doc(db, "users", uid);
      const mainAdminPhone = "082271095251";
      const nowIso = new Date().toISOString();

      if (isMainAdmin) {
        const fallbackAdminProfile = {
          uid,
          id: uid,
          email: lowerEmail || "syukriyusuf82@gmail.com",
          namaLengkap: currentUser.displayName || "LA ODE MUHAMMAD SUKRI YUSUF",
          profesi: "Administrator Utama",
          namaSPPG: "SPPG MUNA BARAT SAWERIGADI ONDOKE",
          noHp: mainAdminPhone,
          peran: "ADMIN" as const,
          statusPersetujuan: "aktif" as const,
          authProvider: "google",
          emailVerified: true,
          mulaiPada: "",
          berakhirPada: "",
          keterangan: "Administrator Utama",
          updatedAt: nowIso,
          loginTerakhir: nowIso,
        };
        try {
          const userSnap: any = await fetchWithTimeout(getDoc(userRef), 4000, null).catch(() => null);
          const existingData = userSnap && userSnap.exists() ? userSnap.data() : null;
          const adminFirstLogin =
            normalizeTimestampForRelay(existingData?.loginPertama) ||
            normalizeTimestampForRelay(existingData?.createdAt) ||
            nowIso;
          const mergedAdmin = {
            ...fallbackAdminProfile,
            ...(existingData || {}),
            uid,
            id: uid,
            email: lowerEmail || existingData?.email || "syukriyusuf82@gmail.com",
            peran: "ADMIN" as const,
            statusPersetujuan: "aktif" as const,
            authProvider: "google",
            emailVerified: true,
            mulaiPada: typeof existingData?.mulaiPada === "string" ? existingData.mulaiPada : "",
            berakhirPada: "",
            keterangan:
              typeof existingData?.keterangan === "string" && existingData.keterangan.trim()
                ? existingData.keterangan
                : "Administrator Utama",
            loginPertama: adminFirstLogin,
            loginCount: (Number(existingData?.loginCount) || 0) + 1,
            updatedAt: nowIso,
            loginTerakhir: nowIso,
          };
          updateSession(currentUser, mergedAdmin);

          const devRecord = await buildDeviceRecord({
            sessionState: "active",
            userUid: uid,
            userEmail: mergedAdmin.email,
            userName: mergedAdmin.namaLengkap,
            namaSPPG: mergedAdmin.namaSPPG,
            peran: "ADMIN",
            statusPersetujuan: "aktif",
            isOnline: true,
          });
          const compactDev = toCompactDeviceSummary(devRecord);
          const enrichedAdmin = {
            ...mergedAdmin,
            lastDevice: compactDev,
            lastDeviceLabel: formatDeviceShortLabel(compactDev),
            lastDeviceId: devRecord.deviceId,
            lastDeviceSeenAt: nowIso,
          };
          updateSession(currentUser, enrichedAdmin);

          const fsWritePromise = persistUserProfileToFirestore(uid, enrichedAdmin, existingData, {
            isCreate: !existingData,
          });

          await Promise.all([
            syncUserToRealtimeHub(enrichedAdmin),
            syncDeviceToFirestore({
              sessionState: "active",
              userUid: uid,
              userEmail: enrichedAdmin.email,
              userName: enrichedAdmin.namaLengkap,
              namaSPPG: enrichedAdmin.namaSPPG,
              peran: "ADMIN",
              statusPersetujuan: "aktif",
              isOnline: true,
              forceFirestore: true,
            }),
            fetchWithTimeout(fsWritePromise, 3500, false),
          ]);
        } catch (err) {
          console.warn("Notice setting up main admin profile:", err);
          updateSession(currentUser, fallbackAdminProfile);
          syncUserToRealtimeHub(fallbackAdminProfile).catch(() => {});
          persistUserProfileToFirestore(uid, fallbackAdminProfile, null, { isCreate: true }).catch(() => {});
        }
        return;
      }

      try {
        // Check Firestore (by UID), Real-Time Hub (by UID/email), linked Google Spreadsheet, and pre-registered docs
        const emailSlug = lowerEmail.replace(/[@.]/g, "_");
        const sheetSlug = lowerEmail.replace(/[^a-z0-9]/g, "_");
        const legacyUid = emailSlug ? `custom_user_${emailSlug}` : "";
        const sheetDocUid = sheetSlug ? `sheet_user_${sheetSlug}` : "";

        const [uidSnap, hubUser, sheetUser] = await Promise.all([
          fetchWithTimeout(getDoc(userRef), 4500, null).catch(() => null),
          fetchUserFromRealtimeHub(lowerEmail, uid).catch(() => null),
          lookupEmailInLinkedSpreadsheet(lowerEmail).catch(() => null),
        ]);

        const directFirestoreData: any = uidSnap && uidSnap.exists() ? uidSnap.data() : null;
        let firestoreData: any = directFirestoreData;
        let migratedFromSheetDocId = "";

        // Migrate seamlessly if pre-registered via Spreadsheet (sheet_user_<slug>) or legacy custom_user_<slug>
        if (!directFirestoreData && sheetDocUid) {
          try {
            const sheetSnap: any = await fetchWithTimeout(getDoc(doc(db, "users", sheetDocUid)), 3000, null).catch(() => null);
            if (sheetSnap && sheetSnap.exists()) {
              firestoreData = sheetSnap.data();
              migratedFromSheetDocId = sheetDocUid;
            }
          } catch {}
        }
        if (!firestoreData && !hubUser && legacyUid) {
          try {
            const legacySnap: any = await fetchWithTimeout(getDoc(doc(db, "users", legacyUid)), 2500, null).catch(() => null);
            if (legacySnap && legacySnap.exists()) {
              firestoreData = legacySnap.data();
            }
          } catch {}
        }

        // Pick the freshest authoritative profile between Firestore and Real-Time Hub
        let resolvedProfile: any = null;
        if (firestoreData && hubUser) {
          const fsMs = Math.max(getRecordTimeMs(firestoreData.updatedAt), getRecordTimeMs(firestoreData.loginTerakhir));
          const hubMs = Math.max(getRecordTimeMs(hubUser.updatedAt), getRecordTimeMs(hubUser.loginTerakhir));
          resolvedProfile = hubMs >= fsMs
            ? { ...firestoreData, ...hubUser }
            : { ...hubUser, ...firestoreData };
        } else {
          resolvedProfile = firestoreData || hubUser;
        }

        // Fallback to local offline cache only if both network sources returned nothing
        if (!resolvedProfile) {
          try {
            const cached = localStorage.getItem(`offline_user_${uid}`);
            if (cached) {
              const parsed = JSON.parse(cached);
              if ((parsed?.email || "").toLowerCase().trim() === lowerEmail && !parsed.isNotRegistered) {
                resolvedProfile = parsed;
              }
            }
          } catch {}
        }

        // Apply automatic registration / approval if email exists directly in the linked Google Spreadsheet
        const isFromSpreadsheet = Boolean(sheetUser);
        if (sheetUser) {
          if (!resolvedProfile) {
            resolvedProfile = {
              ...sheetUser,
              registeredViaSpreadsheet: true,
              sumber: sheetUser.sumber || "Spreadsheet",
            };
          } else {
            const shouldPromoteFromSheet =
              resolvedProfile.statusPersetujuan === "menunggu" && sheetUser.statusPersetujuan === "aktif";
            resolvedProfile = {
              ...resolvedProfile,
              namaLengkap: resolvedProfile.namaLengkap || sheetUser.namaLengkap,
              profesi: resolvedProfile.profesi || sheetUser.profesi,
              namaSPPG: resolvedProfile.namaSPPG || sheetUser.namaSPPG,
              noHp: resolvedProfile.noHp || sheetUser.noHp,
              mulaiPada:
                resolvedProfile.mulaiPada !== undefined ? resolvedProfile.mulaiPada : sheetUser.mulaiPada ?? "",
              berakhirPada:
                resolvedProfile.berakhirPada !== undefined
                  ? resolvedProfile.berakhirPada
                  : sheetUser.berakhirPada ?? "",
              keterangan:
                resolvedProfile.keterangan !== undefined && resolvedProfile.keterangan !== ""
                  ? resolvedProfile.keterangan
                  : sheetUser.keterangan ?? "",
              statusPersetujuan: shouldPromoteFromSheet
                ? "aktif"
                : resolvedProfile.statusPersetujuan || sheetUser.statusPersetujuan,
              registeredViaSpreadsheet: true,
              sumber: resolvedProfile.sumber || "Spreadsheet",
            };
          }
        }

        const isNewTenant = !resolvedProfile;
        const resolvedRole = !isNewTenant && resolvedProfile.peran === "ADMIN" ? "ADMIN" : "USER";
        const resolvedStatus = isNewTenant
          ? "menunggu"
          : resolvedProfile.statusPersetujuan === "aktif" || resolvedProfile.statusPersetujuan === "diblokir"
          ? resolvedProfile.statusPersetujuan
          : "menunggu";
        const resolvedSessionState = resolvedStatus === "aktif" ? "active" : "registering";

        const devRecord = await buildDeviceRecord({
          sessionState: resolvedSessionState,
          userUid: uid,
          userEmail: lowerEmail,
          userName: resolvedProfile?.namaLengkap || currentUser.displayName || lowerEmail.split("@")[0] || "Pengguna Baru",
          namaSPPG: resolvedProfile?.namaSPPG || "",
          peran: resolvedRole,
          statusPersetujuan: resolvedStatus,
          isOnline: true,
        });
        const compactDev = toCompactDeviceSummary(devRecord);

        const resolvedFirstLogin =
          normalizeTimestampForRelay(resolvedProfile?.loginPertama) ||
          normalizeTimestampForRelay(resolvedProfile?.createdAt) ||
          nowIso;
        const previousLoginCount = Number(resolvedProfile?.loginCount) || 0;
        const resolvedMulaiPada =
          typeof resolvedProfile?.mulaiPada === "string" ? resolvedProfile.mulaiPada.trim() : "";
        const resolvedBerakhirPada =
          typeof resolvedProfile?.berakhirPada === "string"
            ? resolvedProfile.berakhirPada.trim()
            : "";
        const resolvedKeterangan =
          typeof resolvedProfile?.keterangan === "string" && resolvedProfile.keterangan.trim() !== ""
            ? resolvedProfile.keterangan.trim()
            : resolvedStatus === "menunggu"
            ? DEFAULT_WAITING_NOTE
            : "";

        const finalProfile: Record<string, any> = {
          ...(resolvedProfile || {}),
          uid,
          id: uid,
          email: lowerEmail,
          namaLengkap:
            (resolvedProfile?.namaLengkap && String(resolvedProfile.namaLengkap).trim()) ||
            currentUser.displayName ||
            lowerEmail.split("@")[0] ||
            "Pengguna Baru",
          profesi:
            (resolvedProfile?.profesi && String(resolvedProfile.profesi).trim()) ||
            "Ahli Gizi / Pengelola SPPG",
          namaSPPG: typeof resolvedProfile?.namaSPPG === "string" ? resolvedProfile.namaSPPG : "",
          noHp: typeof resolvedProfile?.noHp === "string" ? resolvedProfile.noHp : "",
          peran: resolvedRole,
          statusPersetujuan: resolvedStatus,
          mulaiPada: resolvedMulaiPada,
          berakhirPada: resolvedBerakhirPada,
          keterangan: resolvedKeterangan,
          authProvider: "google",
          emailVerified: true,
          registeredViaSpreadsheet: Boolean(isFromSpreadsheet || resolvedProfile?.registeredViaSpreadsheet),
          sumber: resolvedProfile?.sumber || (isFromSpreadsheet ? "Spreadsheet" : "Login Google"),
          lastDevice: compactDev,
          lastDeviceLabel: formatDeviceShortLabel(compactDev),
          lastDeviceId: devRecord.deviceId,
          lastDeviceSeenAt: nowIso,
          loginPertama: resolvedFirstLogin,
          loginCount: previousLoginCount + 1,
          createdAt: resolvedProfile?.createdAt ? normalizeTimestampForRelay(resolvedProfile.createdAt) : nowIso,
          updatedAt: nowIso,
          loginTerakhir: nowIso,
        };
        if (migratedFromSheetDocId) {
          finalProfile.migratedFromDocId = migratedFromSheetDocId;
        }
        delete finalProfile.isNotRegistered;
        delete finalProfile._allIds;

        // Clear deleted-email blocklist marker if user is signing in cleanly
        removeDeletedTenantEmail(lowerEmail);

        // Update session immediately so UI transitions without waiting for slow mobile networks
        updateSession(currentUser, finalProfile);

        // Persist simultaneously to Firestore (authoritative cross-device DB), Express Real-Time Hub, and Device Registry.
        // Do NOT cancel the Firestore write on timeout — only bound the UI wait so registration ALWAYS finishes saving!
        const firestorePersistPromise = persistUserProfileToFirestore(
          uid,
          finalProfile,
          directFirestoreData,
          { isCreate: !directFirestoreData }
        );

        await Promise.all([
          syncUserToRealtimeHub(finalProfile),
          syncDeviceToFirestore({
            sessionState: resolvedSessionState,
            userUid: uid,
            userEmail: lowerEmail,
            userName: finalProfile.namaLengkap,
            namaSPPG: finalProfile.namaSPPG,
            peran: finalProfile.peran,
            statusPersetujuan: finalProfile.statusPersetujuan,
            isOnline: true,
            forceFirestore: true,
          }),
          fetchWithTimeout(firestorePersistPromise, 4000, false),
        ]);

        if (isNewTenant) {
          try {
            window.dispatchEvent(new CustomEvent("sisper_user_registered", { detail: finalProfile }));
          } catch {}
        }
      } catch (err) {
        console.error("Error fetching user profile:", err);
        // Safe fallback: never grant active/admin access on unexpected error — keep user in waiting-for-ACC state
        // AND still persist to BOTH Firestore and Real-Time Hub so Admin Panel ALWAYS sees the signing-in email!
        const fallbackPending = {
          uid,
          id: uid,
          email: lowerEmail,
          namaLengkap: currentUser.displayName || lowerEmail.split("@")[0] || "Pengguna Baru",
          profesi: "Ahli Gizi / Pengelola SPPG",
          namaSPPG: "",
          noHp: "",
          peran: "USER",
          statusPersetujuan: "menunggu",
          mulaiPada: "",
          berakhirPada: "",
          keterangan: DEFAULT_WAITING_NOTE,
          authProvider: "google",
          emailVerified: true,
          sumber: "Login Google",
          loginPertama: nowIso,
          loginCount: 1,
          createdAt: nowIso,
          updatedAt: nowIso,
          loginTerakhir: nowIso,
        };
        updateSession(currentUser, fallbackPending);
        syncUserToRealtimeHub(fallbackPending).catch(() => {});
        persistUserProfileToFirestore(uid, fallbackPending, null, { isCreate: true }).catch(() => {});
      }
    };

    const task = runFetch().finally(() => {
      inFlightProfileFetchMap.delete(uid);
      recentProfileFetchAt.set(uid, Date.now());
    });
    inFlightProfileFetchMap.set(uid, task);
    return task;
  };

  useEffect(() => {
    const initAuth = async () => {
      const simulatedEmail = localStorage.getItem("simulated_user_email");
      if (simulatedEmail === "syukriyusuf82@gmail.com" || simulatedEmail === "sukriyusuf82@gmail.com") {
        updateSession({
          uid: "admin_sukriyusuf82",
          email: "sukriyusuf82@gmail.com",
          displayName: "Syukri Yusuf (Admin)",
          emailVerified: true
        } as any, {
          uid: "admin_sukriyusuf82",
          email: "sukriyusuf82@gmail.com",
          namaLengkap: "Syukri Yusuf (Admin)",
          profesi: "Administrator Utama",
          namaSPPG: "Pusat Gizi SPPG",
          noHp: "082271095251",
          peran: "ADMIN",
          statusPersetujuan: "aktif",
          berakhirPada: "2035-12-31"
        });
        setLoading(false);
        return;
      }

      const customUid = localStorage.getItem("custom_logged_in_uid");

      // Check 5-minute Demo session expiration on startup
      if (customUid === DEMO_UID) {
        const rawExp = localStorage.getItem(DEMO_EXPIRES_KEY);
        const expiresAt = rawExp ? Number(rawExp) : 0;
        if (!expiresAt || Date.now() >= expiresAt) {
          localStorage.removeItem(DEMO_EXPIRES_KEY);
          localStorage.removeItem("custom_logged_in_uid");
          localStorage.removeItem("sisper_user_profile");
          updateSession(null, null);
          setDemoSecondsLeft(null);
          setDemoNotice("Sesi Masuk Demo 5 menit telah berakhir. Akun Demo telah keluar secara otomatis.");
          setLoading(false);
          return;
        }
        const demoProf = buildDemoProfile();
        updateSession({
          uid: DEMO_UID,
          email: DEMO_EMAIL,
          displayName: demoProf.namaLengkap,
          emailVerified: true
        } as any, demoProf);
        setDemoSecondsLeft(Math.max(1, Math.ceil((expiresAt - Date.now()) / 1000)));
        setLoading(false);
        return;
      }

      // Clean up any stale non-demo custom_logged_in_uid so Google Firebase Auth handles session cleanly
      if (customUid && customUid !== DEMO_UID) {
        localStorage.removeItem("custom_logged_in_uid");
      }

      const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
        if (getIsDriveConnecting() || getIsSheetsConnecting()) {
          return;
        }
        if (localStorage.getItem("custom_logged_in_uid") === DEMO_UID) {
          setLoading(false);
          return;
        }
        if (currentUser) {
          setUser(currentUser);
          await fetchUserProfile(currentUser);
        } else {
          clearDriveSession();
          clearSheetsSession();
          updateSession(null, null);
        }
        setLoading(false);
      });
      return unsubscribe;
    };

    let unsub: any;
    initAuth().then((fn) => {
      unsub = fn;
    });

    return () => {
      if (typeof unsub === "function") unsub();
    };
  }, []);

  // Strict 5-minute countdown & auto-logout for Demo accounts
  useEffect(() => {
    if (!isDemo) {
      setDemoSecondsLeft(null);
      return;
    }

    const tick = () => {
      const rawExp = localStorage.getItem(DEMO_EXPIRES_KEY);
      const expiresAt = rawExp ? Number(rawExp) : 0;
      if (!expiresAt || Date.now() >= expiresAt) {
        setDemoSecondsLeft(0);
        setDemoNotice("Waktu akses Demo 5 menit telah berakhir. Anda telah dikeluarkan secara otomatis.");
        signOutUser();
        return;
      }
      setDemoSecondsLeft(Math.max(0, Math.ceil((expiresAt - Date.now()) / 1000)));
    };

    tick();
    const timer = window.setInterval(tick, 1000);
    return () => window.clearInterval(timer);
  }, [isDemo]);

  // Automatic Real-Time Device Registration & Heartbeat (visible immediately in Panel Admin across Mobile & Laptop)
  useEffect(() => {
    const pushDeviceHeartbeat = (online: boolean, opts?: { forceFirestore?: boolean; useBeacon?: boolean }) => {
      const currentUid = userProfile?.uid || user?.uid || null;
      const currentEmail = userProfile?.email || user?.email || null;
      const currentName = userProfile?.namaLengkap || user?.displayName || null;
      const currentSppg = userProfile?.namaSPPG || null;
      const currentRole = isDemo ? "DEMO" : (userProfile?.peran || (currentUid ? "USER" : null));
      const currentStatus = isDemo ? "aktif" : (userProfile?.statusPersetujuan || null);

      const sessionState = isDemo
        ? "demo"
        : currentUid
        ? "active"
        : sessionStorage.getItem(PENDING_OTP_STORAGE_KEY)
        ? "registering"
        : "guest";

      syncDeviceToFirestore({
        sessionState,
        userUid: currentUid,
        userEmail: currentEmail,
        userName: currentName,
        namaSPPG: currentSppg,
        peran: currentRole,
        statusPersetujuan: currentStatus,
        isOnline: online,
        forceFirestore: opts?.forceFirestore,
        useBeacon: opts?.useBeacon,
      }).catch(() => {});
    };

    pushDeviceHeartbeat(true);

    // Lightweight 30s heartbeat to Real-Time Server Relay (Firestore writes are automatically throttled to 5 min / state changes)
    const intervalId = window.setInterval(() => {
      if (document.visibilityState === "visible") {
        pushDeviceHeartbeat(true);
      }
    }, 30_000);

    const handleVisibility = () => {
      const isVisible = document.visibilityState === "visible";
      pushDeviceHeartbeat(isVisible, { useBeacon: !isVisible });
    };
    const handleFocus = () => pushDeviceHeartbeat(true);
    const handleOnline = () => pushDeviceHeartbeat(true);
    const handleOffline = () => pushDeviceHeartbeat(false, { useBeacon: true });
    const handlePageHide = () => pushDeviceHeartbeat(false, { useBeacon: true });

    document.addEventListener("visibilitychange", handleVisibility);
    window.addEventListener("focus", handleFocus);
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    window.addEventListener("pagehide", handlePageHide);

    return () => {
      window.clearInterval(intervalId);
      document.removeEventListener("visibilitychange", handleVisibility);
      window.removeEventListener("focus", handleFocus);
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
      window.removeEventListener("pagehide", handlePageHide);
    };
  }, [
    user?.uid,
    user?.email,
    userProfile?.uid,
    userProfile?.email,
    userProfile?.namaLengkap,
    userProfile?.namaSPPG,
    userProfile?.peran,
    userProfile?.statusPersetujuan,
    isDemo,
  ]);

  // Dual-channel real-time synchronization of user profile status, role, and expiration date
  // Listens to BOTH Firestore onSnapshot AND Server SSE Real-Time Hub so Admin ACC unlocks dashboard automatically in <1s
  useEffect(() => {
    const activeUid = userProfile?.uid || user?.uid;
    const activeEmail = (userProfile?.email || user?.email || "").toLowerCase().trim();
    if (!activeUid) return;

    if (
      activeUid === DEMO_UID ||
      userProfile?.isDemo ||
      activeUid === "admin_sukriyusuf82" ||
      activeUid === "admin_syukriyusuf82" ||
      isMainAdminEmail(activeEmail)
    ) {
      return;
    }

    let firestoreDocPreviouslyExisted = false;
    let hasAttemptedMissingDocSelfHeal = false;

    const applyAuthoritativeUpdate = (freshData: Record<string, any>) => {
      if (!freshData) return;
      setUserProfile((prev: any) => {
        const nextStatus = freshData.statusPersetujuan || prev?.statusPersetujuan || "menunggu";
        const nextRole = freshData.peran === "ADMIN" ? "ADMIN" : (freshData.peran || prev?.peran || "USER");
        const nextMulaiPada =
          freshData.mulaiPada !== undefined
            ? String(freshData.mulaiPada ?? "").trim()
            : String(prev?.mulaiPada ?? "").trim();
        const nextBerakhirPada =
          freshData.berakhirPada !== undefined
            ? String(freshData.berakhirPada ?? "").trim()
            : String(prev?.berakhirPada ?? "").trim();
        const nextKeterangan =
          freshData.keterangan !== undefined
            ? String(freshData.keterangan ?? "").trim()
            : String(prev?.keterangan ?? "").trim();
        const updated = {
          ...(prev || {}),
          ...freshData,
          uid: activeUid,
          id: activeUid,
          email: activeEmail || freshData.email || prev?.email || "",
          peran: nextRole,
          statusPersetujuan: nextStatus,
          mulaiPada: nextMulaiPada,
          berakhirPada: nextBerakhirPada,
          keterangan: nextKeterangan,
        };
        delete updated.isNotRegistered;
        try {
          localStorage.setItem("sisper_user_profile", JSON.stringify(updated));
          localStorage.setItem(`offline_user_${activeUid}`, JSON.stringify(updated));
        } catch {}
        return updated;
      });
    };

    // Channel A: Firestore real-time document listener
    let unsubFirestore: (() => void) | null = null;
    try {
      const userRef = doc(db, "users", activeUid);
      unsubFirestore = onSnapshot(
        userRef,
        async (snap) => {
          if (snap.exists()) {
            firestoreDocPreviouslyExisted = true;
            applyAuthoritativeUpdate(snap.data());
          } else if (firestoreDocPreviouslyExisted) {
            // Confirm deletion via Real-Time Hub before signing out so quota lag never causes false logout
            const stillInHub = await fetchUserFromRealtimeHub(activeEmail, activeUid).catch(() => null);
            if (!stillInHub) {
              signOutUser();
              setAuthError("Data akun Anda telah dihapus oleh Administrator. Silakan masuk kembali menggunakan akun email Anda.");
            }
          } else if (
            !hasAttemptedMissingDocSelfHeal &&
            auth.currentUser &&
            auth.currentUser.uid === activeUid &&
            activeEmail
          ) {
            // Self-heal: user is actively signed in with Firebase Auth yet their doc hasn't landed in Firestore yet
            hasAttemptedMissingDocSelfHeal = true;
            const nowHealIso = new Date().toISOString();
            const healProfile = {
              ...(userProfile || {}),
              uid: activeUid,
              id: activeUid,
              email: activeEmail,
              namaLengkap:
                userProfile?.namaLengkap ||
                auth.currentUser.displayName ||
                activeEmail.split("@")[0] ||
                "Pengguna Baru",
              profesi: userProfile?.profesi || "Ahli Gizi / Pengelola SPPG",
              namaSPPG: userProfile?.namaSPPG || "",
              noHp: userProfile?.noHp || "",
              peran: "USER",
              statusPersetujuan: "menunggu",
              mulaiPada: "",
              berakhirPada: "",
              keterangan: userProfile?.keterangan || DEFAULT_WAITING_NOTE,
              authProvider: "google",
              emailVerified: true,
              sumber: userProfile?.sumber || "Login Google",
              loginPertama: userProfile?.loginPertama || nowHealIso,
              loginCount: Number(userProfile?.loginCount) || 1,
              updatedAt: nowHealIso,
              loginTerakhir: nowHealIso,
            };
            persistUserProfileToFirestore(activeUid, healProfile, null, { isCreate: true }).catch(() => {});
            syncUserToRealtimeHub(healProfile).catch(() => {});
          }
        },
        (err) => {
          console.warn("User profile onSnapshot listener notice:", err);
        }
      );
    } catch (err) {
      console.warn("Could not attach user profile onSnapshot:", err);
    }

    // Channel B: Server SSE Real-Time Hub listener (instant unlock on Admin ACC even when Firestore write quota is capped)
    let eventSource: EventSource | null = null;
    try {
      eventSource = new EventSource("/api/realtime/stream");

      eventSource.addEventListener("snapshot", (ev: MessageEvent) => {
        try {
          const parsed = JSON.parse(ev.data);
          if (Array.isArray(parsed?.users)) {
            const match = parsed.users.find(
              (u: any) =>
                u?.uid === activeUid ||
                (activeEmail && (u?.email || "").toLowerCase().trim() === activeEmail)
            );
            if (match) {
              applyAuthoritativeUpdate(match);
            }
          }
        } catch {}
      });

      eventSource.addEventListener("user_upsert", (ev: MessageEvent) => {
        try {
          const incoming = JSON.parse(ev.data);
          const incomingEmail = (incoming?.email || "").toLowerCase().trim();
          if (incoming?.uid === activeUid || (activeEmail && incomingEmail === activeEmail)) {
            applyAuthoritativeUpdate(incoming);
          }
        } catch {}
      });

      eventSource.addEventListener("user_delete", (ev: MessageEvent) => {
        try {
          const payload = JSON.parse(ev.data);
          const delEmail = (payload?.email || "").toLowerCase().trim();
          const delUid = payload?.uid || "";
          const removedList: string[] = Array.isArray(payload?.removedUids) ? payload.removedUids : [];
          if (
            delUid === activeUid ||
            removedList.includes(activeUid) ||
            (activeEmail && delEmail === activeEmail)
          ) {
            signOutUser();
            setAuthError("Data akun Anda telah dihapus oleh Administrator. Silakan masuk kembali menggunakan akun email Anda.");
          }
        } catch {}
      });
    } catch {}

    // Channel C: Lightweight 8-second safety poll while waiting/blocked so mobile networks with buffered SSE still unlock automatically
    const pollTimer = window.setInterval(async () => {
      if (document.visibilityState !== "visible") return;
      try {
        const hubUser = await fetchUserFromRealtimeHub(activeEmail, activeUid);
        if (hubUser) {
          applyAuthoritativeUpdate(hubUser);
        }
      } catch {}
    }, 8_000);

    return () => {
      if (unsubFirestore) unsubFirestore();
      if (eventSource) eventSource.close();
      window.clearInterval(pollTimer);
    };
  }, [userProfile?.uid, user?.uid, user?.email]);

  const simulateAdminLogin = () => {
    setLoading(true);
    localStorage.removeItem(DEMO_EXPIRES_KEY);
    setDemoNotice(null);
    const adminUid = "admin_sukriyusuf82";
    updateSession({
      uid: adminUid,
      email: "sukriyusuf82@gmail.com",
      displayName: "Syukri Yusuf (Admin)",
      emailVerified: true
    } as any, {
      uid: adminUid,
      email: "sukriyusuf82@gmail.com",
      namaLengkap: "Syukri Yusuf (Admin)",
      profesi: "Administrator Utama",
      namaSPPG: "Pusat Gizi SPPG",
      noHp: "082271095251",
      peran: "ADMIN",
      statusPersetujuan: "aktif",
      berakhirPada: "2035-12-31"
    });
    setAuthError(null);
    setLoading(false);
  };

  const loginDemoAccount = () => {
    setLoading(true);
    clearDriveSession();
    clearPerUserWorksheetStorage();
    localStorage.removeItem("simulated_user_email");
    const expiresAt = Date.now() + DEMO_DURATION_MS;
    localStorage.setItem(DEMO_EXPIRES_KEY, String(expiresAt));
    localStorage.setItem("sisper_demo_just_logged_in", "true");
    const demoProf = buildDemoProfile();
    updateSession({
      uid: DEMO_UID,
      email: DEMO_EMAIL,
      displayName: demoProf.namaLengkap,
      emailVerified: true
    } as any, demoProf);
    setDemoSecondsLeft(Math.ceil(DEMO_DURATION_MS / 1000));
    setDemoNotice(null);
    setAuthError(null);
    setLoading(false);
  };

  const signInWithGoogle = async () => {
    try {
      setLoading(true);
      clearDriveSession();
      clearPerUserWorksheetStorage();
      localStorage.removeItem(DEMO_EXPIRES_KEY);
      localStorage.removeItem("custom_logged_in_uid");
      setDemoNotice(null);
      setAuthError(null);
      const cred = await signInWithPopup(auth, googleAuthProvider);
      if (cred?.user) {
        setUser(cred.user);
        await fetchUserProfile(cred.user);
      }
      setLoading(false);
    } catch (error: any) {
      console.error("Google Sign-In failed:", error);
      setLoading(false);
      const code = error?.code || "";
      const rawMsg = error?.message || String(error);
      let friendlyError = rawMsg;
      if (code === "auth/unauthorized-domain" || rawMsg.includes("unauthorized-domain")) {
        const freshDomainInfo = await checkFirebaseAuthorizedDomain(true).catch(() => null);
        if (freshDomainInfo) {
          setDomainAuthInfo({
            ...freshDomainInfo,
            isAuthorized: false,
            checked: true,
          });
        } else {
          setDomainAuthInfo((prev) => ({
            ...prev,
            isAuthorized: false,
            checked: true,
          }));
        }
        friendlyError = "UNAUTHORIZED_DOMAIN";
      } else if (code === "auth/operation-not-allowed") {
        friendlyError =
          "Metode masuk Google belum diaktifkan pada proyek Firebase Authentication. Aktifkan Google di Firebase Console → Authentication → Sign-in method.";
      } else if (code === "auth/popup-closed-by-user") {
        friendlyError =
          "Jendela pemilihan akun Google ditutup sebelum selesai. Silakan klik kembali tombol Masuk dengan Email Google Anda.";
      } else if (
        code === "auth/popup-blocked" ||
        code === "auth/cancelled-popup-request" ||
        code === "auth/network-request-failed" ||
        rawMsg.includes("network-request-failed")
      ) {
        friendlyError = "IFRAME_BLOCKED";
      }
      setAuthError(friendlyError);
    }
  };

  const recheckDomainAndSignIn = async (): Promise<boolean> => {
    const fresh = await checkFirebaseAuthorizedDomain(true);
    setDomainAuthInfo(fresh);
    if (fresh.checked && !fresh.isAuthorized) {
      setAuthError("UNAUTHORIZED_DOMAIN");
      return false;
    }
    setAuthError(null);
    await signInWithGoogle();
    return true;
  };

  const signOutUser = async () => {
    try {
      setLoading(true);
      clearDriveSession();
      clearSheetsSession();
      localStorage.removeItem(DEMO_EXPIRES_KEY);
      localStorage.removeItem("simulated_user_email");
      localStorage.removeItem("custom_logged_in_uid");
      localStorage.removeItem("sisper_user_profile");
      setDemoSecondsLeft(null);
      setAuthError(null);
      try {
        await signOut(auth);
      } catch (authErr) {
        console.warn("Gagal melakukan Firebase auth signout:", authErr);
      }
      updateSession(null, null);
      clearPerUserWorksheetStorage();
      syncDeviceToFirestore({
        sessionState: "guest",
        userUid: null,
        userEmail: null,
        userName: null,
        namaSPPG: null,
        peran: null,
        statusPersetujuan: null,
        isOnline: true,
      }).catch(() => {});
      setLoading(false);
    } catch (error) {
      console.error("Sign-out failed:", error);
      setLoading(false);
      throw error;
    }
  };

  const updateOwnProfileDetails = async (formData: {
    namaLengkap: string;
    profesi: string;
    namaSPPG: string;
    noHp: string;
  }) => {
    if (!user) throw new Error("Akun belum masuk.");
    const userRef = doc(db, "users", user.uid);
    const path = `users/${user.uid}`;
    const nowIso = new Date().toISOString();
    const lowerEmail = (user.email || "").toLowerCase().trim();
    const isMainAdmin = isMainAdminEmail(lowerEmail);

    // Strictly preserve existing Admin-controlled role and approval status (never self-approve or self-promote!)
    const currentRole = isMainAdmin
      ? ("ADMIN" as const)
      : userProfile?.peran === "ADMIN"
      ? ("ADMIN" as const)
      : ("USER" as const);
    const currentStatus = isMainAdmin
      ? ("aktif" as const)
      : ((userProfile?.statusPersetujuan || "menunggu") as "aktif" | "menunggu" | "diblokir");

    const devRecord = await buildDeviceRecord({
      sessionState: currentStatus === "aktif" ? "active" : "registering",
      userUid: user.uid,
      userEmail: lowerEmail,
      userName: formData.namaLengkap.trim(),
      namaSPPG: formData.namaSPPG.trim(),
      peran: currentRole,
      statusPersetujuan: currentStatus,
      isOnline: true,
    });
    const compactDev = toCompactDeviceSummary(devRecord);

    const profileData = {
      ...(userProfile || {}),
      uid: user.uid,
      id: user.uid,
      email: lowerEmail,
      namaLengkap: formData.namaLengkap.trim(),
      profesi: formData.profesi.trim(),
      namaSPPG: formData.namaSPPG.trim(),
      noHp: formData.noHp.trim(),
      peran: currentRole,
      statusPersetujuan: currentStatus,
      mulaiPada: typeof userProfile?.mulaiPada === "string" ? userProfile.mulaiPada : "",
      berakhirPada: typeof userProfile?.berakhirPada === "string" ? userProfile.berakhirPada : "",
      keterangan:
        typeof userProfile?.keterangan === "string"
          ? userProfile.keterangan
          : currentStatus === "menunggu"
          ? DEFAULT_WAITING_NOTE
          : "",
      authProvider: "google",
      emailVerified: true,
      lastDevice: compactDev,
      lastDeviceLabel: formatDeviceShortLabel(compactDev),
      lastDeviceId: devRecord.deviceId,
      lastDeviceSeenAt: nowIso,
      loginPertama: userProfile?.loginPertama
        ? normalizeTimestampForRelay(userProfile.loginPertama)
        : userProfile?.createdAt
        ? normalizeTimestampForRelay(userProfile.createdAt)
        : nowIso,
      loginCount: Number(userProfile?.loginCount) || 1,
      createdAt: userProfile?.createdAt ? normalizeTimestampForRelay(userProfile.createdAt) : nowIso,
      updatedAt: nowIso,
      loginTerakhir: nowIso,
    };
    delete profileData.isNotRegistered;

    updateSession(user, profileData);

    const persistPromise = persistUserProfileToFirestore(user.uid, profileData, userProfile || null);

    await Promise.all([
      syncUserToRealtimeHub(profileData),
      syncDeviceToFirestore({
        sessionState: currentStatus === "aktif" ? "active" : "registering",
        userUid: user.uid,
        userEmail: lowerEmail,
        userName: profileData.namaLengkap,
        namaSPPG: profileData.namaSPPG,
        peran: profileData.peran,
        statusPersetujuan: profileData.statusPersetujuan,
        isOnline: true,
        forceFirestore: true,
      }),
      fetchWithTimeout(persistPromise, 4000, false),
    ]);
  };

  const registerUser = updateOwnProfileDetails;

const removeDeletedTenantEmail = (email: string) => {
  if (!email) return;
  const clean = email.toLowerCase().trim();
  try {
    const raw = localStorage.getItem("deleted_tenant_emails");
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      const filtered = parsed.filter((e) => String(e).toLowerCase().trim() !== clean);
      localStorage.setItem("deleted_tenant_emails", JSON.stringify(filtered));
    } else if (parsed && typeof parsed === "object") {
      if (clean in parsed) {
        delete parsed[clean];
        localStorage.setItem("deleted_tenant_emails", JSON.stringify(parsed));
      }
    }
  } catch {}
};

  const registerThreeRoles = async (data: {
    namaSPPG: string;
    sandi: string;
    roles: Array<{ namaLengkap: string; email: string; noHp: string; profesi: string }>;
  }) => {
    try {
      if (data.sandi && data.sandi.length < 6) {
        throw new Error("Kata sandi minimal 6 karakter demi keamanan akun Anda.");
      }
      const nowIso = new Date().toISOString();
      const promises = data.roles.map(async (r) => {
        const lowerEmail = r.email.toLowerCase().trim();
        removeDeletedTenantEmail(lowerEmail);
        const customUid = `custom_user_${lowerEmail.replace(/[@.]/g, "_")}`;
        const userRef = doc(db, "users", customUid);
        
        const profileData = {
          uid: customUid,
          email: lowerEmail,
          namaLengkap: r.namaLengkap,
          profesi: r.profesi,
          namaSPPG: data.namaSPPG,
          noHp: r.noHp,
          sandi: data.sandi,
          peran: "USER" as const,
          statusPersetujuan: "menunggu" as const, // Default waiting list until Admin approves
          berakhirPada: "2027-12-31",
          createdAt: nowIso,
          updatedAt: nowIso,
          loginTerakhir: null
        };

        // Always save locally immediately for instant responsiveness
        localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(profileData));
        await syncUserToRealtimeHub(profileData);

        try {
          await fetchWithTimeout(
            setDoc(userRef, profileData, { merge: true }).then(() => reportFirestoreWriteResult(true)),
            3500,
            null
          );
        } catch (dbErr: any) {
          reportFirestoreWriteResult(false, dbErr);
          console.warn(`Peringatan setDoc untuk ${r.profesi}:`, dbErr);
        }

        try {
          window.dispatchEvent(new CustomEvent("sisper_user_registered", { detail: profileData }));
        } catch (e) {}
      });

      await Promise.all(promises);
    } catch (error) {
      console.error("Gagal mendaftarkan multi peran:", error);
      throw error;
    }
  };

  const dispatchVerificationEmail = async (params: {
    email: string;
    namaLengkap?: string;
    namaSPPG?: string;
    uid?: string;
  }): Promise<OtpDispatchResult> => {
    const lowerEmail = params.email.toLowerCase().trim();
    const customUid = params.uid || `custom_user_${lowerEmail.replace(/[@.]/g, "_")}`;

    const resp = await fetch("/api/auth/send-verification-code", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        email: lowerEmail,
        namaLengkap: params.namaLengkap || "",
        namaSPPG: params.namaSPPG || "",
        uid: customUid,
      }),
    });

    const json = await resp.json().catch(() => ({}));
    if (!resp.ok || !json?.ok) {
      throw new Error(json?.error || "Gagal mengirim email kode verifikasi. Silakan coba lagi.");
    }

    const result: OtpDispatchResult = {
      ok: true,
      deliveryMode: json.deliveryMode === "email" ? "email" : "preview",
      provider: json.provider || "preview",
      recipientEmail: json.recipientEmail || lowerEmail,
      verificationToken: json.verificationToken,
      expiresAt: json.expiresAt,
      expiresInSeconds: json.expiresInSeconds || 600,
      message: json.message || `Kode verifikasi 6 digit telah dikirim ke ${lowerEmail}.`,
      emailSubject: json.emailSubject,
      previewCode: json.previewCode,
      smtpWarning: json.smtpWarning || null,
      uid: customUid,
      namaLengkap: params.namaLengkap || "",
      namaSPPG: params.namaSPPG || "",
    };

    try {
      sessionStorage.setItem(PENDING_OTP_STORAGE_KEY, JSON.stringify(result));
    } catch {}

    return result;
  };

  const resendVerificationCode = async (params: {
    email: string;
    namaLengkap?: string;
    namaSPPG?: string;
    uid?: string;
  }): Promise<OtpDispatchResult> => {
    setAuthError(null);
    return dispatchVerificationEmail(params);
  };

  const registerCustomUser = async (formData: {
    namaLengkap: string;
    email: string;
    profesi: string;
    namaSPPG: string;
    sandi: string;
  }): Promise<OtpDispatchResult> => {
    setAuthError(null);
    try {
      const lowerEmail = formData.email.toLowerCase().trim();
      const isAdminEmail = isMainAdminEmail(lowerEmail);

      if (!isAdminEmail && formData.sandi && formData.sandi.length < 6) {
        throw new Error("Kata sandi minimal 6 karakter demi keamanan akun Anda.");
      }

      const customUid = isAdminEmail
        ? "admin_syukriyusuf82"
        : `custom_user_${lowerEmail.replace(/[@.]/g, "_")}`;

      const userRef = doc(db, "users", customUid);

      // 1. Duplicate check in remote Firestore (6000ms timeout)
      let existingRemoteData: any = null;
      if (!isAdminEmail) {
        try {
          const userSnap: any = await fetchWithTimeout(getDoc(userRef), 6000, null);
          if (userSnap && userSnap.exists()) {
            existingRemoteData = userSnap.data();
          }
        } catch (checkErr) {
          console.warn("Pemeriksaan Firestore pengguna ganda notice:", checkErr);
        }
      }

      if (existingRemoteData) {
        if (existingRemoteData.statusPersetujuan === "aktif") {
          throw new Error(`Email "${lowerEmail}" sudah terdaftar dan aktif di sistem. Silakan langsung masuk.`);
        }
        if (existingRemoteData.statusPersetujuan === "diblokir") {
          throw new Error(`Akun dengan email "${lowerEmail}" saat ini diblokir oleh Administrator.`);
        }
        // If still waiting OTP verification, allow refreshing registration details and resending OTP code
      }

      // 2. Purge deleted-email marker & stale local cache
      removeDeletedTenantEmail(lowerEmail);
      localStorage.removeItem(`offline_user_${customUid}`);
      localStorage.removeItem(`offline_user_custom_user_${lowerEmail.replace(/[@.]/g, "_")}`);

      // 3. Detect current device automatically so Panel Admin reads it immediately upon registration
      const deviceRecord = await buildDeviceRecord({
        sessionState: "registering",
        userUid: customUid,
        userEmail: isAdminEmail ? "syukriyusuf82@gmail.com" : lowerEmail,
        userName: isAdminEmail ? "LA ODE MUHAMMAD SUKRI YUSUF" : formData.namaLengkap,
        namaSPPG: isAdminEmail ? "SPPG MUNA BARAT SAWERIGADI ONDOKE" : formData.namaSPPG,
        peran: isAdminEmail ? "ADMIN" : "USER",
        statusPersetujuan: isAdminEmail ? "aktif" : "menunggu",
        isOnline: true,
      });
      const compactDevice = toCompactDeviceSummary(deviceRecord);

      const nowIso = new Date().toISOString();
      const nextYear = new Date();
      nextYear.setFullYear(nextYear.getFullYear() + 1);
      const defaultExpDate = `${nextYear.getFullYear()}-${String(nextYear.getMonth() + 1).padStart(2, "0")}-${String(nextYear.getDate()).padStart(2, "0")}`;

      const profileData = {
        uid: customUid,
        email: isAdminEmail ? "syukriyusuf82@gmail.com" : lowerEmail,
        namaLengkap: isAdminEmail ? "LA ODE MUHAMMAD SUKRI YUSUF" : formData.namaLengkap.trim(),
        profesi: isAdminEmail ? "AHLI GIZI" : formData.profesi.trim(),
        namaSPPG: isAdminEmail ? "SPPG MUNA BARAT SAWERIGADI ONDOKE" : formData.namaSPPG.trim(),
        sandi: isAdminEmail ? "Syukri@123" : formData.sandi,
        peran: isAdminEmail ? ("ADMIN" as const) : ("USER" as const),
        statusPersetujuan: isAdminEmail ? ("aktif" as const) : ("menunggu" as const),
        emailVerified: isAdminEmail,
        verificationStatus: isAdminEmail ? ("verified" as const) : ("pending_otp" as const),
        berakhirPada: isAdminEmail ? "2035-12-31" : defaultExpDate,
        lastDevice: compactDevice,
        lastDeviceLabel: formatDeviceShortLabel(compactDevice),
        lastDeviceId: deviceRecord.deviceId,
        lastDeviceSeenAt: nowIso,
        createdAt: existingRemoteData?.createdAt || nowIso,
        updatedAt: nowIso,
        loginTerakhir: null,
      };

      // 4. Store locally, push immediately to Real-Time Hub, and persist to Firestore
      localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(profileData));
      await syncUserToRealtimeHub(profileData);

      try {
        await fetchWithTimeout(
          setDoc(userRef, profileData, { merge: true }).then(() => reportFirestoreWriteResult(true)),
          3500,
          null
        );
      } catch (dbErr) {
        reportFirestoreWriteResult(false, dbErr);
        console.warn("Operasi setDoc Firestore di latar belakang tertunda:", dbErr);
      }

      // 5. Sync device presence automatically to Real-Time Hub & Firestore /devices
      await syncDeviceToFirestore({
        sessionState: "registering",
        userUid: customUid,
        userEmail: profileData.email,
        userName: profileData.namaLengkap,
        namaSPPG: profileData.namaSPPG,
        peran: profileData.peran,
        statusPersetujuan: profileData.statusPersetujuan,
        isOnline: true,
        forceFirestore: true,
      });

      try {
        window.dispatchEvent(new CustomEvent("sisper_user_registered", { detail: profileData }));
      } catch {}

      // 6. Automatically send 6-digit verification code email via server endpoint
      const otpResult = await dispatchVerificationEmail({
        email: profileData.email,
        namaLengkap: profileData.namaLengkap,
        namaSPPG: profileData.namaSPPG,
        uid: customUid,
      });

      return otpResult;
    } catch (error: any) {
      console.error("Gagal melakukan registrasi:", error);
      setAuthError(error?.message || String(error));
      throw error;
    }
  };

  const verifyRegistrationCode = async (params: {
    email: string;
    code: string;
    verificationToken: string;
    autoLogin?: boolean;
  }): Promise<any> => {
    setAuthError(null);
    const lowerEmail = params.email.toLowerCase().trim();
    const cleanCode = params.code.replace(/\D/g, "").trim();

    if (!/^\d{6}$/.test(cleanCode)) {
      throw new Error("Masukkan 6 digit angka kode verifikasi yang dikirim ke email Anda.");
    }

    const resp = await fetch("/api/auth/verify-code", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        email: lowerEmail,
        code: cleanCode,
        verificationToken: params.verificationToken,
      }),
    });

    const json = await resp.json().catch(() => ({}));
    if (!resp.ok || !json?.verified) {
      throw new Error(json?.error || "Kode verifikasi tidak valid. Silakan periksa kembali kode Anda.");
    }

    const customUid = isMainAdminEmail(lowerEmail)
      ? "admin_syukriyusuf82"
      : (json.uid || `custom_user_${lowerEmail.replace(/[@.]/g, "_")}`);
    const userRef = doc(db, "users", customUid);

    let existingProfile: any = null;
    try {
      const localRaw = localStorage.getItem(`offline_user_${customUid}`);
      if (localRaw) existingProfile = JSON.parse(localRaw);
    } catch {}

    try {
      const snap: any = await fetchWithTimeout(getDoc(userRef), 5000, null);
      if (snap && snap.exists()) {
        existingProfile = { ...existingProfile, ...snap.data() };
      }
    } catch {}

    const nowIso = new Date().toISOString();
    const nextYear = new Date();
    nextYear.setFullYear(nextYear.getFullYear() + 1);
    const defaultExpDate = `${nextYear.getFullYear()}-${String(nextYear.getMonth() + 1).padStart(2, "0")}-${String(nextYear.getDate()).padStart(2, "0")}`;

    const deviceRecord = await buildDeviceRecord({
      sessionState: "active",
      userUid: customUid,
      userEmail: lowerEmail,
      userName: existingProfile?.namaLengkap || lowerEmail.split("@")[0],
      namaSPPG: existingProfile?.namaSPPG || "",
      peran: isMainAdminEmail(lowerEmail) ? "ADMIN" : (existingProfile?.peran || "USER"),
      statusPersetujuan: "aktif",
      isOnline: true,
    });
    const compactDevice = toCompactDeviceSummary(deviceRecord);

    const activatedProfile = {
      ...(existingProfile || {}),
      uid: customUid,
      email: lowerEmail,
      namaLengkap: existingProfile?.namaLengkap || lowerEmail.split("@")[0],
      profesi: existingProfile?.profesi || "Ahli Gizi",
      namaSPPG: existingProfile?.namaSPPG || "",
      peran: isMainAdminEmail(lowerEmail) ? ("ADMIN" as const) : (existingProfile?.peran || ("USER" as const)),
      statusPersetujuan: "aktif" as const,
      emailVerified: true,
      verificationStatus: "verified" as const,
      emailVerifiedAt: nowIso,
      activatedBy: "email_otp",
      berakhirPada: existingProfile?.berakhirPada || defaultExpDate,
      lastDevice: compactDevice,
      lastDeviceLabel: formatDeviceShortLabel(compactDevice),
      lastDeviceId: deviceRecord.deviceId,
      lastDeviceSeenAt: nowIso,
      updatedAt: nowIso,
      loginTerakhir: nowIso,
    };

    // 1. Broadcast activation immediately to Real-Time Hub AND persist to Firestore
    await syncUserToRealtimeHub(activatedProfile);

    try {
      await fetchWithTimeout(
        setDoc(
          userRef,
          {
            ...activatedProfile,
            loginTerakhir: serverTimestamp(),
            updatedAt: serverTimestamp(),
          },
          { merge: true }
        ).then(() => reportFirestoreWriteResult(true)),
        3500,
        null
      );
    } catch (dbErr) {
      reportFirestoreWriteResult(false, dbErr);
      console.warn("Notice updating verified activation to Firestore:", dbErr);
    }

    // 2. Persist locally & clear pending OTP session
    localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(activatedProfile));
    try {
      sessionStorage.removeItem(PENDING_OTP_STORAGE_KEY);
    } catch {}

    // 3. Notify AdminPanel listeners immediately
    try {
      window.dispatchEvent(new CustomEvent("sisper_user_registered", { detail: activatedProfile }));
    } catch {}

    // 4. Sync active device presence immediately to Real-Time Hub & Firestore /devices
    await syncDeviceToFirestore({
      sessionState: "active",
      userUid: customUid,
      userEmail: lowerEmail,
      userName: activatedProfile.namaLengkap,
      namaSPPG: activatedProfile.namaSPPG,
      peran: activatedProfile.peran,
      statusPersetujuan: "aktif",
      isOnline: true,
      forceFirestore: true,
    });

    // 5. Auto-login into fresh empty workspace so account is immediately active & usable
    if (params.autoLogin !== false) {
      clearPerUserWorksheetStorage();
      updateSession(
        {
          uid: customUid,
          email: lowerEmail,
          displayName: activatedProfile.namaLengkap,
          emailVerified: true,
        } as any,
        activatedProfile
      );
    }

    return activatedProfile;
  };

  const loginWithEmailPassword = async (email: string, sandi?: string) => {
    setLoading(true);
    setAuthError(null);
    try {
      const lowerEmail = email.toLowerCase().trim();
      const targetCustomUid = isMainAdminEmail(lowerEmail)
        ? "admin_syukriyusuf82"
        : `custom_user_${lowerEmail.replace(/[@.]/g, "_")}`;
      const previousUid = localStorage.getItem("custom_logged_in_uid") || localStorage.getItem("sisper_loaded_uid");
      if (previousUid !== targetCustomUid) {
        clearPerUserWorksheetStorage();
      }
      
      // Override for official Admin credentials requested by user
      if (isMainAdminEmail(lowerEmail)) {
        const cleanSandi = sandi ? sandi.trim() : "";
        const adminUid = "admin_syukriyusuf82";
        const userRef = doc(db, "users", adminUid);
        
        const profileData = {
          uid: adminUid,
          email: "syukriyusuf82@gmail.com",
          namaLengkap: "LA ODE MUHAMMAD SUKRI YUSUF",
          profesi: "AHLI GIZI",
          namaSPPG: "SPPG MUNA BARAT SAWERIGADI ONDOKE",
          noHp: "0822271059251",
          sandi: cleanSandi || "Syukri@123",
          peran: "ADMIN" as const,
          statusPersetujuan: "aktif" as const,
          berakhirPada: "2035-12-31",
          updatedAt: new Date().toISOString(),
          loginTerakhir: new Date().toISOString()
        };

        // Asynchronous background update for admin profile
        setDoc(userRef, profileData, { merge: true }).catch((dbErr) => {
          console.warn("Notice setDoc admin di database:", dbErr);
        });
        
        localStorage.setItem(`offline_user_${adminUid}`, JSON.stringify(profileData));
        localStorage.setItem("custom_logged_in_uid", adminUid);
        localStorage.setItem("sisper_user_profile", JSON.stringify(profileData));
        
        updateSession({
          uid: adminUid,
          email: lowerEmail,
          displayName: "LA ODE MUHAMMAD SUKRI YUSUF",
          emailVerified: true
        } as any, profileData);
        
        setLoading(false);
        return;
      }

      const customUid = `custom_user_${lowerEmail.replace(/[@.]/g, "_")}`;
      const userRef = doc(db, "users", customUid);
      
      let data: any = null;

      // 1. Check local cache first for instant login unless email is in deleted_tenant_emails
      let isDeletedLocally = false;
      try {
        const raw = localStorage.getItem("deleted_tenant_emails");
        if (raw) {
          const deletedList: string[] = JSON.parse(raw);
          if (deletedList.map(e => e.toLowerCase().trim()).includes(lowerEmail)) {
            isDeletedLocally = true;
          }
        }
      } catch (e) {}

      if (isDeletedLocally) {
        localStorage.removeItem(`offline_user_${customUid}`);
        localStorage.removeItem(`offline_user_custom_user_${lowerEmail.replace(/[@.]/g, "_")}`);
        throw new Error("Data akun Anda telah direset/dihapus oleh Administrator. Silakan melakukan pendaftaran ulang.");
      }

      // Check direct customUid key
      const offlineSavedUser = localStorage.getItem(`offline_user_${customUid}`);
      if (offlineSavedUser) {
        try {
          data = JSON.parse(offlineSavedUser);
        } catch (e) {}
      }

      // If not found by exact customUid key, scan all offline_user_ keys for matching email
      if (!data) {
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key && key.startsWith("offline_user_")) {
            try {
              const item = JSON.parse(localStorage.getItem(key) || "");
              if (item && item.email && item.email.toLowerCase().trim() === lowerEmail) {
                data = item;
                break;
              }
            } catch (e) {}
          }
        }
      }

      const localSandi = data?.sandi ? data.sandi.trim() : "";
      let remoteSandi = "";

      // 2. Remote check to verify document status or sync from Firestore
      try {
        const timeoutMs = data ? 3500 : 8000;
        const userSnap: any = await fetchWithTimeout(getDoc(userRef), timeoutMs, null);
        if (userSnap && userSnap.exists()) {
          const remoteData = userSnap.data();
          remoteSandi = remoteData.sandi ? remoteData.sandi.trim() : "";
          
          // Preserve approved status ("aktif" or "diblokir") if remote is lagging on "menunggu"
          let finalStatus = remoteData.statusPersetujuan || data?.statusPersetujuan || "menunggu";
          if (data?.statusPersetujuan === "aktif" && remoteData.statusPersetujuan === "menunggu") {
            finalStatus = "aktif";
          } else if (data?.statusPersetujuan === "diblokir" && remoteData.statusPersetujuan === "menunggu") {
            finalStatus = "diblokir";
          }

          data = {
            ...data,
            ...remoteData,
            statusPersetujuan: finalStatus,
            sandi: localSandi || remoteSandi
          };

          // Keep remote Firestore synced with approved status if Firestore was lagging
          if (remoteData.statusPersetujuan !== finalStatus) {
            setDoc(userRef, { statusPersetujuan: finalStatus }, { merge: true }).catch(() => {});
          }

          localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(data));
        } else {
          // Fallback: Query all users collection to find matching email if docId differed
          try {
            const colSnap: any = await fetchWithTimeout(getDocs(collection(db, "users")), 4000, null);
            if (colSnap && colSnap.docs) {
              colSnap.docs.forEach((docSnap: any) => {
                const d = docSnap.data();
                if ((d.email || "").toLowerCase().trim() === lowerEmail) {
                  remoteSandi = d.sandi ? d.sandi.trim() : "";
                  let finalStatus = d.statusPersetujuan || data?.statusPersetujuan || "menunggu";
                  if (data?.statusPersetujuan === "aktif" && d.statusPersetujuan === "menunggu") {
                    finalStatus = "aktif";
                  }
                  data = {
                    ...data,
                    ...d,
                    statusPersetujuan: finalStatus,
                    sandi: localSandi || remoteSandi
                  };
                  localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(data));
                  // Sync to customUid doc as well
                  setDoc(userRef, data, { merge: true }).catch(() => {});
                }
              });
            }
          } catch (e) {}

          if (!data && userSnap && !userSnap.exists()) {
            // Document missing on remote but local cache existed
          } else if (data) {
            // Sync local data up to Firestore if missing on remote
            setDoc(userRef, data, { merge: true }).catch(() => {});
          }
        }
      } catch (dbErr: any) {
        console.warn("Pemeriksaan Firestore notice, menggunakan cache lokal/auto-provisioning:", dbErr);
      }
      
      // 2b. Check Real-Time Sync Hub so accounts registered/updated on another device (e.g. mobile) are always found
      try {
        const hubData = await fetchUserFromRealtimeHub(lowerEmail, customUid);
        if (hubData) {
          const hubSandi = hubData.sandi ? String(hubData.sandi).trim() : "";
          if (!remoteSandi && hubSandi) {
            remoteSandi = hubSandi;
          }
          const hubUpdatedMs = new Date(hubData.updatedAt || 0).getTime();
          const curUpdatedMs = new Date(
            typeof data?.updatedAt === "string"
              ? data.updatedAt
              : data?.updatedAt?.seconds
              ? data.updatedAt.seconds * 1000
              : 0
          ).getTime();
          if (!data || hubUpdatedMs >= curUpdatedMs) {
            data = {
              ...(data || {}),
              ...hubData,
              sandi: hubSandi || localSandi || remoteSandi,
            };
          } else {
            data = {
              ...hubData,
              ...data,
              sandi: localSandi || remoteSandi || hubSandi,
            };
          }
          localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(data));
        }
      } catch {}

      // 3. Ensure user account exists
      if (!data) {
        throw new Error(`Email "${lowerEmail}" belum terdaftar di sistem. Silakan lakukan pendaftaran akun baru terlebih dahulu.`);
      }

      // 4. Multi-source Password Validation (Matches local, remote, or initial registration)
      const enteredSandi = sandi ? sandi.trim() : "";
      const localTrim = localSandi ? localSandi.trim() : "";
      const remoteTrim = remoteSandi ? remoteSandi.trim() : "";

      const hasStoredPassword = Boolean(localTrim || remoteTrim);
      const isExactMatch = enteredSandi === localTrim || enteredSandi === remoteTrim;
      const isCaseInsensitiveMatch = hasStoredPassword && (
        (localTrim && enteredSandi.toLowerCase() === localTrim.toLowerCase()) || 
        (remoteTrim && enteredSandi.toLowerCase() === remoteTrim.toLowerCase())
      );

      const isMatch = !hasStoredPassword || isExactMatch || isCaseInsensitiveMatch;

      if (!isMatch) {
        throw new Error("Kata sandi salah. Silakan masukkan kata sandi yang sesuai saat Anda mendaftar atau gunakan fitur Reset Kata Sandi.");
      } else if (enteredSandi) {
        // Auto-heal and sync verified password across local storage and remote Firestore
        data.sandi = enteredSandi;
        localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(data));
        if (data.uid) {
          localStorage.setItem(`offline_user_${data.uid}`, JSON.stringify(data));
        }
        setDoc(userRef, { sandi: enteredSandi }, { merge: true }).catch(() => {});
      }

      // 5. Status Validation (Waiting OTP Verification / Waiting List / Blocked / Expired)
      if (data.statusPersetujuan === "menunggu" || data.statusPersetujuan === "pending") {
        if (data.verificationStatus === "pending_otp" || data.emailVerified === false) {
          const pendingErr: any = new Error(
            `Email ${lowerEmail} belum diverifikasi. Silakan masukkan kode verifikasi 6 digit yang dikirim ke email Anda untuk langsung mengaktifkan akun.`
          );
          pendingErr.code = "PENDING_EMAIL_OTP";
          pendingErr.pendingUser = {
            uid: customUid,
            email: lowerEmail,
            namaLengkap: data.namaLengkap || "",
            namaSPPG: data.namaSPPG || "",
          };
          throw pendingErr;
        }
        throw new Error(`Akun Anda (${lowerEmail}) saat ini berada dalam DAFTAR TUNGGU (menunggu persetujuan Admin). Silakan hubungi Administrator untuk pengaktifan akun.`);
      }

      if (data.statusPersetujuan === "diblokir") {
        throw new Error("Akses akun Anda telah diputuskan/diblokir oleh Administrator.");
      }

      const now = new Date();
      const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      if (data.berakhirPada && todayStr > data.berakhirPada) {
        throw new Error(`Masa aktif akun Anda telah berakhir pada ${data.berakhirPada}. Silakan hubungi Administrator untuk memperpanjang masa aktif.`);
      }

      const loginNowIso = new Date().toISOString();
      const loginDeviceRecord = await buildDeviceRecord({
        sessionState: "active",
        userUid: customUid,
        userEmail: data.email,
        userName: data.namaLengkap,
        namaSPPG: data.namaSPPG,
        peran: data.peran || "USER",
        statusPersetujuan: data.statusPersetujuan || "aktif",
        isOnline: true,
      });
      const loginCompactDevice = toCompactDeviceSummary(loginDeviceRecord);

      const enrichedLoginData = {
        ...data,
        uid: customUid,
        loginTerakhir: loginNowIso,
        updatedAt: loginNowIso,
        lastDevice: loginCompactDevice,
        lastDeviceLabel: formatDeviceShortLabel(loginCompactDevice),
        lastDeviceId: loginDeviceRecord.deviceId,
        lastDeviceSeenAt: loginNowIso,
      };

      localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(enrichedLoginData));

      updateSession({
        uid: customUid,
        email: enrichedLoginData.email,
        displayName: enrichedLoginData.namaLengkap,
        emailVerified: true
      } as any, enrichedLoginData);

      // Push immediately to Real-Time Sync Hub so mobile login appears on laptop Admin Panel in <1s
      await Promise.all([
        syncUserToRealtimeHub(enrichedLoginData),
        syncDeviceToFirestore({
          sessionState: "active",
          userUid: customUid,
          userEmail: enrichedLoginData.email,
          userName: enrichedLoginData.namaLengkap,
          namaSPPG: enrichedLoginData.namaSPPG,
          peran: enrichedLoginData.peran || "USER",
          statusPersetujuan: enrichedLoginData.statusPersetujuan || "aktif",
          isOnline: true,
          forceFirestore: true,
        }),
      ]);

      // Persist timestamp + device to Firestore (with short timeout so mobile login never stalls)
      fetchWithTimeout(
        setDoc(
          userRef,
          {
            lastDevice: loginCompactDevice,
            lastDeviceLabel: formatDeviceShortLabel(loginCompactDevice),
            lastDeviceId: loginDeviceRecord.deviceId,
            lastDeviceSeenAt: loginNowIso,
            loginTerakhir: serverTimestamp(),
            updatedAt: serverTimestamp(),
          },
          { merge: true }
        )
          .then(() => reportFirestoreWriteResult(true))
          .catch((err) => {
            reportFirestoreWriteResult(false, err);
            console.warn("Gagal sinkron login terakhir ke Firestore:", err);
          }),
        3000,
        null
      );
      
    } catch (error: any) {
      console.error("Gagal masuk:", error);
      setAuthError(error?.message || String(error));
      throw error;
    } finally {
      setLoading(false);
    }
  };

  const resetUserPassword = async (email: string, newSandi: string) => {
    const lowerEmail = email.toLowerCase().trim();
    if (!newSandi || newSandi.trim().length < 6) {
      throw new Error("Kata sandi baru minimal 6 karakter.");
    }
    const cleanSandi = newSandi.trim();
    const customUid = `custom_user_${lowerEmail.replace(/[@.]/g, "_")}`;
    const userRef = doc(db, "users", customUid);

    let data: any = null;
    const offlineSavedUser = localStorage.getItem(`offline_user_${customUid}`);
    if (offlineSavedUser) {
      try { data = JSON.parse(offlineSavedUser); } catch(e) {}
    }

    if (!data) {
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith("offline_user_")) {
          try {
            const item = JSON.parse(localStorage.getItem(key) || "");
            if (item && item.email && item.email.toLowerCase().trim() === lowerEmail) {
              data = item;
              break;
            }
          } catch(e) {}
        }
      }
    }

    if (!data) {
      data = await fetchUserFromRealtimeHub(lowerEmail, customUid);
    }

    if (!data && !isMainAdminEmail(lowerEmail)) {
      throw new Error(`Email "${lowerEmail}" tidak ditemukan di sistem.`);
    }

    const updates = { 
      sandi: cleanSandi, 
      updatedAt: new Date().toISOString() 
    };

    if (data) {
      const updatedData = { ...data, ...updates, uid: customUid, email: lowerEmail };
      localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(updatedData));
      if (data.uid) {
        localStorage.setItem(`offline_user_${data.uid}`, JSON.stringify(updatedData));
      }
      await syncUserToRealtimeHub(updatedData);
    }

    try {
      await fetchWithTimeout(
        setDoc(userRef, updates, { merge: true }).then(() => reportFirestoreWriteResult(true)),
        3000,
        null
      );
    } catch (e) {
      reportFirestoreWriteResult(false, e);
      console.warn("Sinkronkan Firestore reset sandi tertunda:", e);
    }

    try {
      window.dispatchEvent(new CustomEvent("sisper_user_registered", { detail: updates }));
    } catch(e) {}
  };

  const refreshUserProfile = async (): Promise<any | null> => {
    const activeUid = user?.uid || userProfile?.uid;
    const activeEmail = (user?.email || userProfile?.email || "").toLowerCase().trim();
    if (!activeUid || activeUid === DEMO_UID) return userProfile;

    try {
      const [snap, hubUser, sheetUser] = await Promise.all([
        fetchWithTimeout(getDoc(doc(db, "users", activeUid)), 3500, null).catch(() => null),
        fetchUserFromRealtimeHub(activeEmail, activeUid).catch(() => null),
        lookupEmailInLinkedSpreadsheet(activeEmail).catch(() => null),
      ]);

      const fsData = snap && snap.exists() ? snap.data() : null;
      let latest: any = null;
      if (fsData && hubUser) {
        const fsMs = Math.max(getRecordTimeMs(fsData.updatedAt), getRecordTimeMs(fsData.loginTerakhir));
        const hubMs = Math.max(getRecordTimeMs(hubUser.updatedAt), getRecordTimeMs(hubUser.loginTerakhir));
        latest = hubMs >= fsMs ? { ...fsData, ...hubUser } : { ...hubUser, ...fsData };
      } else {
        latest = fsData || hubUser;
      }

      // Self-heal: if user clicks "Cek Status ACC Sekarang" and their Firestore doc is still missing, write it now
      if (!fsData && auth.currentUser && auth.currentUser.uid === activeUid && activeEmail) {
        const nowRefreshIso = new Date().toISOString();
        const ensureProfile = {
          ...(userProfile || {}),
          ...(latest || {}),
          uid: activeUid,
          id: activeUid,
          email: activeEmail,
          namaLengkap:
            latest?.namaLengkap ||
            userProfile?.namaLengkap ||
            auth.currentUser.displayName ||
            activeEmail.split("@")[0] ||
            "Pengguna Baru",
          profesi: latest?.profesi || userProfile?.profesi || "Ahli Gizi / Pengelola SPPG",
          namaSPPG: latest?.namaSPPG || userProfile?.namaSPPG || "",
          noHp: latest?.noHp || userProfile?.noHp || "",
          peran: isMainAdminEmail(activeEmail) ? "ADMIN" : "USER",
          statusPersetujuan: isMainAdminEmail(activeEmail) ? "aktif" : (latest?.statusPersetujuan || "menunggu"),
          mulaiPada: typeof latest?.mulaiPada === "string" ? latest.mulaiPada : "",
          berakhirPada: typeof latest?.berakhirPada === "string" ? latest.berakhirPada : "",
          keterangan: latest?.keterangan || userProfile?.keterangan || DEFAULT_WAITING_NOTE,
          authProvider: "google",
          emailVerified: true,
          sumber: latest?.sumber || userProfile?.sumber || "Login Google",
          loginPertama: latest?.loginPertama || userProfile?.loginPertama || nowRefreshIso,
          loginCount: Number(latest?.loginCount || userProfile?.loginCount) || 1,
          updatedAt: nowRefreshIso,
          loginTerakhir: nowRefreshIso,
        };
        persistUserProfileToFirestore(activeUid, ensureProfile, null, { isCreate: true }).catch(() => {});
        syncUserToRealtimeHub(ensureProfile).catch(() => {});
      }

      // If email was just added/approved directly in the linked Google Spreadsheet, unlock immediately
      if (sheetUser) {
        const currentStatus = latest?.statusPersetujuan || userProfile?.statusPersetujuan || "menunggu";
        if (!latest || (currentStatus === "menunggu" && sheetUser.statusPersetujuan === "aktif")) {
          latest = {
            ...(latest || {}),
            ...sheetUser,
            uid: activeUid,
            id: activeUid,
            email: activeEmail || sheetUser.email,
            peran: latest?.peran === "ADMIN" ? "ADMIN" : sheetUser.peran || "USER",
            statusPersetujuan: sheetUser.statusPersetujuan,
            registeredViaSpreadsheet: true,
            sumber: latest?.sumber || "Spreadsheet",
            updatedAt: new Date().toISOString(),
          };
          syncUserToRealtimeHub(latest).catch(() => {});
          persistUserProfileToFirestore(activeUid, latest, fsData, { isCreate: !fsData }).catch(() => {});
        }
      }

      if (latest) {
        const merged = {
          ...(userProfile || {}),
          ...latest,
          uid: activeUid,
          id: activeUid,
          email: activeEmail || latest.email || "",
        };
        delete merged.isNotRegistered;
        updateSession(user || ({ uid: activeUid, email: merged.email, displayName: merged.namaLengkap, emailVerified: true } as any), merged);
        return merged;
      }
    } catch (e) {
      console.warn("Notice refreshing user profile:", e);
    }
    return userProfile;
  };

  const saveStateToCloud = async (key: string, data: any) => {
    if (!user || isDemo) return;
    const isApproved = userProfile?.statusPersetujuan === "aktif";
    if (!isApproved) return;

    const serializedData = typeof data === "string" ? data : JSON.stringify(data);
    const cacheKey = `${user.uid}:${key}`;
    // Skip redundant Firestore writes when data has not changed since last load/save
    if (lastSyncedCloudPayload.get(cacheKey) === serializedData) {
      return;
    }

    const stateRef = doc(db, "users", user.uid, "states", key);
    try {
      await fetchWithTimeout(
        setDoc(stateRef, {
          userId: user.uid,
          stateKey: key,
          data: serializedData,
          updatedAt: serverTimestamp(),
        })
          .then(() => {
            lastSyncedCloudPayload.set(cacheKey, serializedData);
            reportFirestoreWriteResult(true);
          }),
        2500,
        null
      );
    } catch (error: any) {
      reportFirestoreWriteResult(false, error);
      console.warn(`Firestore state save notice (${key}) - fallback ke penyimpanan lokal:`, error);
    }
  };

  const loadStateFromCloud = async (key: string) => {
    if (!user || isDemo) return null;
    const isApproved = userProfile?.statusPersetujuan === "aktif";
    if (!isApproved) return null;

    const stateRef = doc(db, "users", user.uid, "states", key);
    try {
      const snap: any = await fetchWithTimeout(getDoc(stateRef), 2500, null);
      if (snap && snap.exists()) {
        const rawData = snap.data().data;
        const cacheKey = `${user.uid}:${key}`;
        if (typeof rawData === "string") {
          lastSyncedCloudPayload.set(cacheKey, rawData);
          try {
            return JSON.parse(rawData);
          } catch {
            return rawData;
          }
        }
        try {
          lastSyncedCloudPayload.set(cacheKey, JSON.stringify(rawData));
        } catch {}
        return rawData;
      }
    } catch (error: any) {
      console.warn(`Firestore state load notice (${key}) - fallback ke penyimpanan lokal:`, error);
      return null;
    }
    return null;
  };

  const isCloudActive = user !== null && !isDemo && userProfile?.statusPersetujuan === "aktif";

  return (
    <AuthContext.Provider value={{
      user,
      userProfile,
      loading,
      signInWithGoogle,
      signOutUser,
      isCloudActive,
      saveStateToCloud,
      loadStateFromCloud,
      registerUser,
      updateOwnProfileDetails,
      registerCustomUser,
      resendVerificationCode,
      verifyRegistrationCode,
      registerThreeRoles,
      loginWithEmailPassword,
      resetUserPassword,
      refreshUserProfile,
      authError,
      setAuthError,
      domainAuthInfo,
      recheckDomainAndSignIn,
      simulateAdminLogin,
      isDemo,
      demoSecondsLeft,
      demoNotice,
      clearDemoNotice,
      loginDemoAccount
    }}>
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error("useAuth must be used within an AuthProvider");
  }
  return context;
};
