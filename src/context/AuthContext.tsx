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
} from "../lib/firebase.ts";
import { clearDriveSession, getIsDriveConnecting } from "../lib/googleDrive.ts";
import {
  buildDeviceRecord,
  formatDeviceShortLabel,
  syncDeviceToFirestore,
  toCompactDeviceSummary,
} from "../lib/deviceTracker.ts";

// Quota protection: deduplicate identical state writes so opening/refreshing never burns Firestore write units
const lastSyncedCloudPayload = new Map<string, string>();

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
  refreshUserProfile: () => Promise<void>;
  authError: string | null;
  setAuthError: (error: string | null) => void;
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

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<User | null>(() => {
    const customUid = localStorage.getItem("custom_logged_in_uid");
    if (customUid) {
      const savedProfile = localStorage.getItem("sisper_user_profile");
      if (savedProfile) {
        try {
          const profile = JSON.parse(savedProfile);
          return {
            uid: customUid,
            email: profile.email,
            displayName: profile.namaLengkap,
            emailVerified: true
          } as any;
        } catch (e) {
          return null;
        }
      }
    }
    return null;
  });

  const [userProfile, setUserProfile] = useState<any | null>(() => {
    const savedProfile = localStorage.getItem("sisper_user_profile");
    if (savedProfile) {
      try {
        return JSON.parse(savedProfile);
      } catch (e) {
        return null;
      }
    }
    return null;
  });

  const [loading, setLoading] = useState(true);
  const [authError, setAuthError] = useState<string | null>(null);
  const [demoNotice, setDemoNotice] = useState<string | null>(null);
  const [demoSecondsLeft, setDemoSecondsLeft] = useState<number | null>(() => {
    const rawExp = localStorage.getItem(DEMO_EXPIRES_KEY);
    if (!rawExp) return null;
    const rem = Math.ceil((Number(rawExp) - Date.now()) / 1000);
    return rem > 0 ? rem : 0;
  });

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
      localStorage.setItem("custom_logged_in_uid", currentUser.uid);
      localStorage.setItem("sisper_user_profile", JSON.stringify(profile));
    } else {
      localStorage.removeItem("custom_logged_in_uid");
      localStorage.removeItem("sisper_user_profile");
    }
  };

  const fetchUserProfile = async (currentUser: User) => {
    const lowerEmail = currentUser.email?.toLowerCase();
    const isMainAdmin = isMainAdminEmail(lowerEmail);
    const userRef = doc(db, "users", currentUser.uid);
    const path = `users/${currentUser.uid}`;
    const mainAdminPhone = "0822271059251";

    if (isMainAdmin) {
      try {
        const userSnap = await getDoc(userRef);
        if (userSnap.exists()) {
          const data = userSnap.data();
          if (data.peran !== "ADMIN" || data.statusPersetujuan !== "aktif") {
            // Document exists but does not have active ADMIN role, update it!
            await setDoc(userRef, {
              peran: "ADMIN",
              statusPersetujuan: "aktif",
              namaLengkap: data.namaLengkap || "LA ODE MUHAMMAD SUKRI YUSUF",
              profesi: data.profesi || "Administrator Utama",
              namaSPPG: data.namaSPPG || "SPPG MUNA BARAT SAWERIGADI ONDOKE",
              noHp: data.noHp || mainAdminPhone,
              berakhirPada: "2035-12-31",
              updatedAt: serverTimestamp(),
              loginTerakhir: serverTimestamp()
            }, { merge: true });
            const snap = await getDoc(userRef);
            updateSession(currentUser, snap.data());
          } else {
            // Already active ADMIN, just update login timestamp
            await setDoc(userRef, {
              loginTerakhir: serverTimestamp(),
              updatedAt: serverTimestamp()
            }, { merge: true });
            const snap = await getDoc(userRef);
            updateSession(currentUser, snap.data());
          }
        } else {
          // Document does not exist yet, create it!
          const adminProfile = {
            uid: currentUser.uid,
            email: currentUser.email || "syukriyusuf82@gmail.com",
            namaLengkap: "LA ODE MUHAMMAD SUKRI YUSUF",
            profesi: "Administrator Utama",
            namaSPPG: "SPPG MUNA BARAT SAWERIGADI ONDOKE",
            noHp: mainAdminPhone,
            peran: "ADMIN" as const,
            statusPersetujuan: "aktif" as const,
            berakhirPada: "2035-12-31",
            createdAt: serverTimestamp(),
            updatedAt: serverTimestamp(),
            loginTerakhir: serverTimestamp()
          };
          await setDoc(userRef, adminProfile);
          const snap = await getDoc(userRef);
          updateSession(currentUser, snap.data());
        }
      } catch (err) {
        console.error("Error setting up main admin profile:", err);
        updateSession(currentUser, {
          uid: currentUser.uid,
          email: currentUser.email || "syukriyusuf82@gmail.com",
          namaLengkap: "LA ODE MUHAMMAD SUKRI YUSUF",
          profesi: "Administrator Utama",
          namaSPPG: "SPPG MUNA BARAT SAWERIGADI ONDOKE",
          noHp: mainAdminPhone,
          peran: "ADMIN",
          statusPersetujuan: "aktif",
          berakhirPada: "2035-12-31"
        });
      }
      return;
    }

    try {
      const userSnap = await getDoc(userRef);
      if (userSnap.exists()) {
        const data = userSnap.data();
        const nowIso = new Date().toISOString();
        const devRecord = await buildDeviceRecord({
          sessionState: "active",
          userUid: currentUser.uid,
          userEmail: data.email || currentUser.email || "",
          userName: data.namaLengkap || currentUser.displayName || "",
          namaSPPG: data.namaSPPG || "",
          peran: data.peran || "USER",
          statusPersetujuan: data.statusPersetujuan || "aktif",
          isOnline: true,
        });
        const compactDev = toCompactDeviceSummary(devRecord);
        const enriched: Record<string, any> = {
          ...data,
          lastDevice: compactDev,
          lastDeviceLabel: formatDeviceShortLabel(compactDev),
          lastDeviceId: devRecord.deviceId,
          lastDeviceSeenAt: nowIso,
          loginTerakhir: nowIso,
        };
        updateSession(currentUser, enriched);

        // Push immediately to real-time hub so laptop Admin Panel reflects mobile Google login in <1s
        await syncUserToRealtimeHub(enriched);
        await syncDeviceToFirestore({
          sessionState: "active",
          userUid: currentUser.uid,
          userEmail: enriched.email || currentUser.email || "",
          userName: enriched.namaLengkap || currentUser.displayName || "",
          namaSPPG: enriched.namaSPPG || "",
          peran: enriched.peran || "USER",
          statusPersetujuan: enriched.statusPersetujuan || "aktif",
          isOnline: true,
          forceFirestore: true,
        });

        // Update last login + device on Firestore
        try {
          await fetchWithTimeout(
            setDoc(
              userRef,
              {
                lastDevice: compactDev,
                lastDeviceLabel: formatDeviceShortLabel(compactDev),
                lastDeviceId: devRecord.deviceId,
                lastDeviceSeenAt: nowIso,
                loginTerakhir: serverTimestamp(),
                updatedAt: serverTimestamp(),
              },
              { merge: true }
            ).then(() => reportFirestoreWriteResult(true)),
            3500,
            null
          );
        } catch (timestampErr) {
          reportFirestoreWriteResult(false, timestampErr);
          console.warn("Gagal memperbarui login terakhir di awan:", timestampErr);
        }
      } else {
        // Also check realtime hub in case user registered recently while Firestore write quota was capped
        const hubUser = await fetchUserFromRealtimeHub(currentUser.email || "", currentUser.uid);
        if (hubUser) {
          updateSession(currentUser, hubUser);
        } else {
          // Normal user: Needs registration
          updateSession(currentUser, {
            uid: currentUser.uid,
            email: currentUser.email || "",
            isNotRegistered: true
          });
        }
      }
    } catch (err) {
      console.error("Error fetching user profile:", err);
      // Fallback local setting to prevent lock out in case of rules delay
      updateSession(currentUser, {
        uid: currentUser.uid,
        email: currentUser.email || "",
        isNotRegistered: true
      });
    }
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
          noHp: "0822271059251",
          sandi: "Odhe@1998",
          peran: "ADMIN",
          statusPersetujuan: "aktif",
          berakhirPada: "2030-12-31"
        });
        setLoading(false);
        return;
      }

      let customUid = localStorage.getItem("custom_logged_in_uid");

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

      if (customUid) {
        if (customUid === "admin_sukriyusuf82" || customUid === "admin_syukriyusuf82" || customUid === "syukriyusuf82_simulated_uid") {
          const isSukri = customUid === "admin_sukriyusuf82";
          const finalEmail = isSukri ? "sukriyusuf82@gmail.com" : "syukriyusuf82@gmail.com";
          updateSession({
            uid: customUid,
            email: finalEmail,
            displayName: "Syukri Yusuf (Admin)",
            emailVerified: true
          } as any, {
            uid: customUid,
            email: finalEmail,
            namaLengkap: "Syukri Yusuf (Admin)",
            profesi: "Administrator Utama",
            namaSPPG: "Pusat Gizi SPPG",
            noHp: "0822271059251",
            sandi: "Odhe@1998",
            peran: "ADMIN",
            statusPersetujuan: "aktif",
            berakhirPada: "2030-12-31"
          });
          setLoading(false);
          return;
        }

        // Fast load from saved local profile
        const savedProfile = localStorage.getItem("sisper_user_profile");
        let localData: any = null;
        if (savedProfile) {
          try {
            localData = JSON.parse(savedProfile);
            updateSession({
              uid: customUid,
              email: localData.email,
              displayName: localData.namaLengkap,
              emailVerified: true
            } as any, localData);
          } catch (e) {
            console.warn("Gagal parsing profil lokal:", e);
          }
        }

        // Instantly unblock loading screen so app opens without delay
        setLoading(false);

        // Background profile verification with Firestore validation
        fetchWithTimeout(getDoc(doc(db, "users", customUid)), 3500, null)
          .then((snap: any) => {
            if (snap && snap.exists()) {
              const remoteData = snap.data();
              
              // Validate approval status across Firestore and local cache
              const isApproved = remoteData.statusPersetujuan === "aktif" || localData?.statusPersetujuan === "aktif";
              const isBlocked = remoteData.statusPersetujuan === "diblokir" || localData?.statusPersetujuan === "diblokir";
              const finalStatus = isBlocked ? "diblokir" : isApproved ? "aktif" : "menunggu";

              const mergedData = {
                ...localData,
                ...remoteData,
                statusPersetujuan: finalStatus,
                uid: customUid
              };

              localStorage.setItem("sisper_user_profile", JSON.stringify(mergedData));
              localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(mergedData));

              if (remoteData.statusPersetujuan !== finalStatus) {
                setDoc(doc(db, "users", customUid), { statusPersetujuan: finalStatus }, { merge: true }).catch(() => {});
              }

              updateSession({
                uid: customUid,
                email: mergedData.email,
                displayName: mergedData.namaLengkap,
                emailVerified: true
              } as any, mergedData);
            }
          })
          .catch((e) => {
            console.warn("Gagal memuat profil awan secara asinkron:", e);
          });

        return;
      }

      const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
        if (getIsDriveConnecting()) {
          return;
        }
        if (currentUser) {
          setUser(currentUser);
          await fetchUserProfile(currentUser);
        } else {
          if (!localStorage.getItem("custom_logged_in_uid")) {
            updateSession(null, null);
          }
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

  // Real-time synchronization of user profile status and expiration date from Firestore
  useEffect(() => {
    const activeUid = userProfile?.uid || user?.uid || localStorage.getItem("custom_logged_in_uid");
    if (!activeUid) return;

    // Skip snapshot listener for demo accounts, local simulated admin override, or main admin
    if (
      activeUid === DEMO_UID ||
      userProfile?.isDemo ||
      activeUid === "admin_sukriyusuf82" ||
      activeUid === "admin_syukriyusuf82" ||
      isMainAdminEmail(userProfile?.email)
    ) {
      return;
    }

    try {
      const userRef = doc(db, "users", activeUid);
      const unsubscribe = onSnapshot(userRef, (snap) => {
        if (snap.exists()) {
          const freshData = snap.data();
          
          // Check expiration
          const now = new Date();
          const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
          const isExpired = freshData.berakhirPada ? todayStr > freshData.berakhirPada : false;

          const currentLocalStatus = userProfile?.statusPersetujuan;
          const isApproved = freshData.statusPersetujuan === "aktif" || currentLocalStatus === "aktif";
          const isBlocked = freshData.statusPersetujuan === "diblokir";

          if (isBlocked || (!isApproved && freshData.statusPersetujuan === "menunggu") || isExpired) {
            signOutUser();
            setAuthError(
              isExpired
                ? `Masa aktif akun Anda telah berakhir pada ${freshData.berakhirPada}. Silakan hubungi admin untuk memperpanjang.`
                : freshData.statusPersetujuan === "menunggu"
                ? `Status akun Anda berada dalam DAFTAR TUNGGU (menunggu persetujuan Admin).`
                : "Akses akun Anda telah diputuskan/diblokir oleh Administrator."
            );
          } else {
            const finalStatus = "aktif";
            if (freshData.statusPersetujuan !== finalStatus) {
              setDoc(userRef, { statusPersetujuan: finalStatus }, { merge: true }).catch(() => {});
            }

            setUserProfile((prev: any) => {
              if (!prev) return { ...freshData, statusPersetujuan: finalStatus };
              const updated = { ...prev, ...freshData, statusPersetujuan: finalStatus };
              localStorage.setItem("sisper_user_profile", JSON.stringify(updated));
              if (activeUid) {
                localStorage.setItem(`offline_user_${activeUid}`, JSON.stringify(updated));
              }
              return updated;
            });
          }
        } else {
          // User document removed or reset by Admin in Firestore -> Immediate real-time force logout
          if (!isMainAdminEmail(userProfile?.email)) {
            signOutUser();
            setAuthError("Data akun Anda telah direset/dihapus oleh Administrator. Silakan melakukan pendaftaran ulang dengan email Anda.");
          }
        }
      }, (err) => {
        console.warn("User profile onSnapshot listener notice:", err);
      });

      return () => unsubscribe();
    } catch (err) {
      console.warn("Could not attach user profile onSnapshot:", err);
    }
  }, [userProfile?.uid, user?.uid]);

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
      noHp: "0822271059251",
      sandi: "Odhe@1998",
      peran: "ADMIN",
      statusPersetujuan: "aktif",
      berakhirPada: "2030-12-31"
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
      setDemoNotice(null);
      setAuthError(null);
      await signInWithPopup(auth, googleAuthProvider);
    } catch (error: any) {
      console.error("Google Sign-In failed:", error);
      setLoading(false);
      let friendlyError = error?.code || error?.message || String(error);
      if (error?.code === "auth/network-request-failed" || String(error).includes("network-request-failed")) {
        friendlyError = "IFRAME_BLOCKED";
      }
      setAuthError(friendlyError);
      throw error;
    }
  };

  const signOutUser = async () => {
    try {
      setLoading(true);
      clearDriveSession();
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

  const registerUser = async (formData: { namaLengkap: string; profesi: string; namaSPPG: string; noHp: string }) => {
    if (!user) throw new Error("No authenticated user found");
    const userRef = doc(db, "users", user.uid);
    const path = `users/${user.uid}`;
    const nowIso = new Date().toISOString();
    
    const lowerEmail = user.email?.toLowerCase();
    const isMainAdmin = lowerEmail === "syukriyusuf82@gmail.com" || lowerEmail === "sukriyusuf82@gmail.com";

    const devRecord = await buildDeviceRecord({
      sessionState: "active",
      userUid: user.uid,
      userEmail: user.email || "",
      userName: formData.namaLengkap,
      namaSPPG: formData.namaSPPG,
      peran: isMainAdmin ? "ADMIN" : "USER",
      statusPersetujuan: "aktif",
      isOnline: true,
    });
    const compactDev = toCompactDeviceSummary(devRecord);
    
    const profileData = {
      uid: user.uid,
      email: user.email || "",
      namaLengkap: formData.namaLengkap,
      profesi: formData.profesi,
      namaSPPG: formData.namaSPPG,
      noHp: formData.noHp,
      peran: isMainAdmin ? ("ADMIN" as const) : ("USER" as const),
      statusPersetujuan: "aktif" as const,
      emailVerified: true,
      verificationStatus: "verified" as const,
      berakhirPada: "2030-12-31",
      lastDevice: compactDev,
      lastDeviceLabel: formatDeviceShortLabel(compactDev),
      lastDeviceId: devRecord.deviceId,
      lastDeviceSeenAt: nowIso,
      createdAt: nowIso,
      updatedAt: nowIso,
      loginTerakhir: nowIso
    };

    // Push immediately to real-time server hub so mobile Google registration shows on laptop Admin Panel in <1s
    await syncUserToRealtimeHub(profileData);
    await syncDeviceToFirestore({
      sessionState: "active",
      userUid: user.uid,
      userEmail: user.email || "",
      userName: formData.namaLengkap,
      namaSPPG: formData.namaSPPG,
      peran: profileData.peran,
      statusPersetujuan: "aktif",
      isOnline: true,
      forceFirestore: true,
    });

    try {
      await fetchWithTimeout(
        setDoc(userRef, {
          ...profileData,
          createdAt: serverTimestamp(),
          updatedAt: serverTimestamp(),
          loginTerakhir: serverTimestamp()
        }).then(() => reportFirestoreWriteResult(true)),
        4000,
        null
      );
      updateSession(user, profileData);
    } catch (error) {
      reportFirestoreWriteResult(false, error);
      handleFirestoreError(error, OperationType.WRITE, path);
      updateSession(user, profileData);
    }
  };

const removeDeletedTenantEmail = (email: string) => {
  if (!email) return;
  try {
    const raw = localStorage.getItem("deleted_tenant_emails");
    if (raw) {
      const arr: string[] = JSON.parse(raw);
      const filtered = arr.filter((e) => e !== email.toLowerCase().trim());
      localStorage.setItem("deleted_tenant_emails", JSON.stringify(filtered));
    }
  } catch (e) {}
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

  const refreshUserProfile = async () => {
    const customUid = localStorage.getItem("custom_logged_in_uid");
    if (customUid) {
      try {
        const userRef = doc(db, "users", customUid);
        const snap = await getDoc(userRef);
        if (snap.exists()) {
          updateSession({
            uid: customUid,
            email: snap.data().email,
            displayName: snap.data().namaLengkap,
            emailVerified: true
          } as any, snap.data());
        }
      } catch (e) {
        console.error(e);
      }
      return;
    }
    if (user) {
      await fetchUserProfile(user);
    }
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
      registerCustomUser,
      resendVerificationCode,
      verifyRegistrationCode,
      registerThreeRoles,
      loginWithEmailPassword,
      resetUserPassword,
      refreshUserProfile,
      authError,
      setAuthError,
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
