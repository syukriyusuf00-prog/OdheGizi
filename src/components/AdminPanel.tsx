import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import { onAuthStateChanged, User } from "firebase/auth";
import {
  collection,
  onSnapshot,
  getDocsFromServer,
  doc,
  setDoc,
  deleteDoc,
  serverTimestamp,
} from "firebase/firestore";
import {
  auth,
  db,
  firestoreRestDeleteDoc,
  firestoreRestListCollection,
  firestoreRestUpsertDoc,
  reportFirestoreWriteResult,
  stripUndefinedDeep,
} from "../lib/firebase.ts";
import {
  DEFAULT_WAITING_NOTE,
  isMainAdminEmail,
  syncUserToRealtimeHub,
  deleteUserFromRealtimeHub,
  useAuth,
} from "../context/AuthContext.tsx";
import {
  Search,
  Download,
  UserCheck,
  Ban,
  Trash2,
  RotateCcw,
  RefreshCw,
  Save,
  X,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Smartphone,
  Tablet,
  Monitor,
  Camera,
  ChevronDown,
  ChevronUp,
  ArrowLeft,
} from "lucide-react";
import { DeviceRecord, formatDeviceShortLabel } from "../lib/deviceTracker.ts";
import { downloadUsersCsv } from "../lib/googleSheetsSync.ts";
import SpreadsheetUserSyncCard from "./SpreadsheetUserSyncCard.tsx";

const isDeviceCurrentlyOnline = (d: DeviceRecord): boolean => {
  if (!d || d.isOnline === false) return false;
  const lastMs = d.lastSeenAt ? new Date(d.lastSeenAt).getTime() : 0;
  if (!lastMs || Number.isNaN(lastMs)) return Boolean(d.isOnline);
  return Date.now() - lastMs <= 2 * 60 * 1000;
};

const deleteDeviceFromRealtimeHub = async (deviceId: string): Promise<void> => {
  try {
    await fetch("/api/realtime/delete-device", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ deviceId }),
    });
  } catch {}
};

interface AdminPanelProps {
  onClose: () => void;
  globalLogoUrl?: string | null;
  onUploadGlobalLogo?: (e: React.ChangeEvent<HTMLInputElement>) => void;
  onResetGlobalLogo?: () => void;
  adminEmail?: string;
}

type StatusFilter = "semua" | "menunggu" | "aktif" | "kedaluwarsa" | "diblokir";

interface UserDraft {
  mulaiPada: string;
  berakhirPada: string;
  keterangan: string;
}

const fetchWithTimeout = <T,>(promise: Promise<T>, timeoutMs: number, defaultValue: T): Promise<T> => {
  return Promise.race([
    promise,
    new Promise<T>((resolve) => setTimeout(() => resolve(defaultValue), timeoutMs)),
  ]);
};

const DELETED_EMAIL_TTL_MS = 10 * 60 * 1000; // 10 menit — mencegah cache lokal menyembunyikan email yang login ulang

const getDeletedTenantEmailMap = (): Record<string, number> => {
  try {
    const raw = localStorage.getItem("deleted_tenant_emails");
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    const now = Date.now();
    // Migrate legacy permanent string[] format by expiring it so re-logging emails are never permanently hidden
    if (Array.isArray(parsed)) {
      localStorage.removeItem("deleted_tenant_emails");
      return {};
    }
    if (parsed && typeof parsed === "object") {
      const cleaned: Record<string, number> = {};
      let mutated = false;
      for (const [k, v] of Object.entries(parsed)) {
        const em = String(k).toLowerCase().trim();
        const ts = Number(v);
        if (em && Number.isFinite(ts) && now - ts < DELETED_EMAIL_TTL_MS) {
          cleaned[em] = ts;
        } else {
          mutated = true;
        }
      }
      if (mutated) {
        localStorage.setItem("deleted_tenant_emails", JSON.stringify(cleaned));
      }
      return cleaned;
    }
  } catch {}
  return {};
};

export const getDeletedTenantEmails = (): Set<string> => {
  return new Set(Object.keys(getDeletedTenantEmailMap()));
};

export const addDeletedTenantEmail = (email: string) => {
  if (!email) return;
  const map = getDeletedTenantEmailMap();
  map[email.toLowerCase().trim()] = Date.now();
  try {
    localStorage.setItem("deleted_tenant_emails", JSON.stringify(map));
  } catch {}
};

export const removeDeletedTenantEmail = (email: string) => {
  if (!email) return;
  const clean = email.toLowerCase().trim();
  const map = getDeletedTenantEmailMap();
  if (clean in map) {
    delete map[clean];
    try {
      localStorage.setItem("deleted_tenant_emails", JSON.stringify(map));
    } catch {}
  }
};

const getTodayIsoLocal = (): string => {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(
    now.getDate()
  ).padStart(2, "0")}`;
};

const addDurationToDateIso = (
  baseIso: string | undefined,
  delta: { days?: number; months?: number; years?: number }
): string => {
  const todayIso = getTodayIsoLocal();
  const baseStr = baseIso && /^\d{4}-\d{2}-\d{2}$/.test(baseIso) && baseIso >= todayIso ? baseIso : todayIso;
  const [y, m, d] = baseStr.split("-").map(Number);
  const dt = new Date(y, (m || 1) - 1, d || 1);
  if (delta.days) dt.setDate(dt.getDate() + delta.days);
  if (delta.months) dt.setMonth(dt.getMonth() + delta.months);
  if (delta.years) dt.setFullYear(dt.getFullYear() + delta.years);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(
    dt.getDate()
  ).padStart(2, "0")}`;
};

const getTimestampMs = (ts: any): number => {
  if (!ts) return 0;
  try {
    if (typeof ts?.seconds === "number") return ts.seconds * 1000;
    if (typeof ts?.toDate === "function") return ts.toDate().getTime();
    const d = new Date(ts);
    return Number.isNaN(d.getTime()) ? 0 : d.getTime();
  } catch {
    return 0;
  }
};

const formatDateTimeId = (ts: any): string => {
  const ms = getTimestampMs(ts);
  if (!ms) return "—";
  try {
    const dt = new Date(ms);
    const datePart = dt.toLocaleDateString("id-ID", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
    const timePart = dt
      .toLocaleTimeString("id-ID", {
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(":", ".");
    return `${datePart}, ${timePart}`;
  } catch {
    return "—";
  }
};

const getEffectiveCategory = (
  u: any,
  todayIso: string
): "menunggu" | "aktif" | "kedaluwarsa" | "diblokir" => {
  const rawStatus = String(u?.statusPersetujuan || "menunggu").toLowerCase();
  if (rawStatus === "diblokir") return "diblokir";
  if (rawStatus === "menunggu" || rawStatus === "pending") return "menunggu";
  const berakhir = typeof u?.berakhirPada === "string" ? u.berakhirPada.trim() : "";
  if (berakhir && u?.peran !== "ADMIN" && !isMainAdminEmail(u?.email) && todayIso > berakhir) {
    return "kedaluwarsa";
  }
  return "aktif";
};

const isScheduledInFuture = (u: any, todayIso: string): boolean => {
  if (u?.peran === "ADMIN" || isMainAdminEmail(u?.email)) return false;
  if (String(u?.statusPersetujuan || "").toLowerCase() !== "aktif") return false;
  const mulai = typeof u?.mulaiPada === "string" ? u.mulaiPada.trim() : "";
  return Boolean(mulai && todayIso < mulai);
};

const readLocalCachedUsers = (): any[] => {
  const localUsers: any[] = [];
  const deletedEmails = getDeletedTenantEmails();
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith("offline_user_")) {
        const val = localStorage.getItem(k);
        if (val) {
          const parsed = JSON.parse(val);
          if (parsed && parsed.email) {
            const em = parsed.email.toLowerCase().trim();
            if (!deletedEmails.has(em) && em !== "demo@sispermbg.id") {
              localUsers.push({
                id: parsed.uid || k.replace("offline_user_", ""),
                ...parsed,
              });
            }
          }
        }
      }
    }
  } catch {}
  return localUsers;
};

const synthesizeUsersFromDevices = (devicesList: DeviceRecord[]): any[] => {
  const byEmail = new Map<string, any>();
  devicesList.forEach((d) => {
    const em = String(d?.userEmail || "").toLowerCase().trim();
    if (!em || em === "demo@sispermbg.id" || !em.includes("@")) return;
    const seenMs = getTimestampMs(d.lastSeenAt);
    const prev = byEmail.get(em);
    if (prev && getTimestampMs(prev.loginTerakhir) >= seenMs) return;

    const uid =
      (d.userUid && d.userUid !== "demo_user_5min" ? d.userUid : "") ||
      `device_user_${em.replace(/[^a-z0-9]/g, "_")}`;
    byEmail.set(em, {
      id: uid,
      uid,
      email: em,
      namaLengkap: d.userName || em.split("@")[0] || "Pengguna Baru",
      profesi: "Ahli Gizi / Pengelola SPPG",
      namaSPPG: d.namaSPPG || "",
      noHp: "",
      peran: d.peran === "ADMIN" ? "ADMIN" : "USER",
      statusPersetujuan:
        d.statusPersetujuan === "aktif" || d.statusPersetujuan === "diblokir"
          ? d.statusPersetujuan
          : "menunggu",
      mulaiPada: "",
      berakhirPada: "",
      keterangan: d.statusPersetujuan === "aktif" ? "" : DEFAULT_WAITING_NOTE,
      loginPertama: d.firstSeenAt || d.lastSeenAt || new Date().toISOString(),
      loginTerakhir: d.lastSeenAt || new Date().toISOString(),
      updatedAt: d.lastSeenAt || new Date().toISOString(),
      createdAt: d.firstSeenAt || d.lastSeenAt || new Date().toISOString(),
      loginCount: 1,
      sumber: "Login Google",
      lastDeviceId: d.deviceId,
      lastDeviceLabel: formatDeviceShortLabel(d),
      _fromDeviceFallback: true,
    });
  });
  return Array.from(byEmail.values());
};

interface MergeUserSources {
  authoritative?: any[]; // Cloud Firestore users & Real-Time Hub users — NEVER hidden by local deleted_tenant_emails
  deviceFallback?: any[]; // Synthesized from devices — filtered by deleted_tenant_emails unless newer
  localCached?: any[]; // Local browser cache — filtered by deleted_tenant_emails
  pendingDeleted?: Map<string, number>; // Short-lived in-flight deletes (prevents 1-2s snapshot flicker)
}

const mergeDeduplicatedUsers = (sources: MergeUserSources): any[] => {
  const deletedMap = getDeletedTenantEmailMap();
  const byEmail = new Map<string, any>();
  const now = Date.now();

  const isPlaceholderId = (id: string) =>
    id.startsWith("custom_user_") ||
    id.startsWith("sheet_user_") ||
    id.startsWith("device_user_") ||
    id.startsWith("admin_");

  const taggedItems: Array<{ raw: any; isAuthoritative: boolean }> = [
    ...(sources.authoritative || []).map((raw) => ({ raw, isAuthoritative: true })),
    ...(sources.deviceFallback || []).map((raw) => ({ raw, isAuthoritative: false })),
    ...(sources.localCached || []).map((raw) => ({ raw, isAuthoritative: false })),
  ];

  taggedItems.forEach(({ raw, isAuthoritative }) => {
    if (!raw || !raw.email) return;
    const em = String(raw.email).toLowerCase().trim();
    if (!em || em === "demo@sispermbg.id") return;

    // Suppress only while an explicit Admin DELETE is currently propagating (< 6 seconds)
    const pendingDeleteTs = sources.pendingDeleted?.get(em);
    if (pendingDeleteTs && now - pendingDeleteTs < 6000) {
      return;
    }

    if (isAuthoritative) {
      // Authoritative live record from Cloud Firestore or Hub must NEVER be hidden by stale localStorage cache
      if (deletedMap[em]) {
        removeDeletedTenantEmail(em);
        delete deletedMap[em];
      }
    } else {
      const deletedAtMs = deletedMap[em];
      if (deletedAtMs) {
        const recordActivityMs = Math.max(
          getTimestampMs(raw.loginTerakhir),
          getTimestampMs(raw.updatedAt),
          getTimestampMs(raw.createdAt)
        );
        // If this user logged in AGAIN after Admin deleted them, automatically un-hide them!
        if (recordActivityMs > deletedAtMs + 1500) {
          removeDeletedTenantEmail(em);
          delete deletedMap[em];
        } else {
          return;
        }
      }
    }

    const uid = String(raw.uid || raw.id || `user_${em.replace(/[^a-z0-9]/g, "_")}`);
    const normalized = {
      ...raw,
      id: uid,
      uid,
      email: em,
    };

    const existing = byEmail.get(em);
    if (!existing) {
      byEmail.set(em, {
        ...normalized,
        _allIds: [uid],
      });
      return;
    }

    // Real user records always take precedence over synthesized device fallback records
    if (existing._fromDeviceFallback && !normalized._fromDeviceFallback) {
      byEmail.set(em, {
        ...normalized,
        _allIds: Array.from(new Set([...(existing._allIds || [existing.id]), uid])),
      });
      return;
    }
    if (!existing._fromDeviceFallback && normalized._fromDeviceFallback) {
      return;
    }

    const allIds = Array.from(new Set([...(existing._allIds || [existing.id]), uid]));
    const existingMs = Math.max(getTimestampMs(existing.updatedAt), getTimestampMs(existing.loginTerakhir));
    const incomingMs = Math.max(getTimestampMs(normalized.updatedAt), getTimestampMs(normalized.loginTerakhir));

    const preferredId =
      !isPlaceholderId(uid) && isPlaceholderId(existing.id)
        ? uid
        : !isPlaceholderId(existing.id) && isPlaceholderId(uid)
        ? existing.id
        : incomingMs >= existingMs
        ? uid
        : existing.id;

    const base = incomingMs >= existingMs ? { ...existing, ...normalized } : { ...normalized, ...existing };
    const firstMsExisting = getTimestampMs(existing.loginPertama || existing.createdAt);
    const firstMsIncoming = getTimestampMs(normalized.loginPertama || normalized.createdAt);
    const earliestFirst =
      firstMsExisting > 0 && (firstMsIncoming === 0 || firstMsExisting <= firstMsIncoming)
        ? existing.loginPertama || existing.createdAt
        : normalized.loginPertama || normalized.createdAt;

    byEmail.set(em, {
      ...base,
      id: preferredId,
      uid: preferredId,
      email: em,
      loginPertama: earliestFirst || base.createdAt,
      loginCount: Math.max(Number(existing.loginCount) || 1, Number(normalized.loginCount) || 1),
      mulaiPada:
        incomingMs >= existingMs
          ? normalized.mulaiPada !== undefined
            ? String(normalized.mulaiPada ?? "")
            : String(existing.mulaiPada ?? "")
          : existing.mulaiPada !== undefined
          ? String(existing.mulaiPada ?? "")
          : String(normalized.mulaiPada ?? ""),
      berakhirPada:
        incomingMs >= existingMs
          ? normalized.berakhirPada !== undefined
            ? String(normalized.berakhirPada ?? "")
            : String(existing.berakhirPada ?? "")
          : existing.berakhirPada !== undefined
          ? String(existing.berakhirPada ?? "")
          : String(normalized.berakhirPada ?? ""),
      keterangan:
        incomingMs >= existingMs
          ? normalized.keterangan !== undefined
            ? String(normalized.keterangan ?? "")
            : String(existing.keterangan ?? "")
          : existing.keterangan !== undefined
          ? String(existing.keterangan ?? "")
          : String(normalized.keterangan ?? ""),
      _allIds: allIds,
    });
  });

  const todayIso = getTodayIsoLocal();
  const priorityWeight = (u: any): number => {
    const cat = getEffectiveCategory(u, todayIso);
    if (cat === "menunggu") return 0;
    if (cat === "kedaluwarsa") return 1;
    if (cat === "diblokir") return 2;
    return 3;
  };

  return Array.from(byEmail.values()).sort((a, b) => {
    const diff = priorityWeight(a) - priorityWeight(b);
    if (diff !== 0) return diff;
    const tA = Math.max(getTimestampMs(a.loginTerakhir), getTimestampMs(a.updatedAt), getTimestampMs(a.createdAt));
    const tB = Math.max(getTimestampMs(b.loginTerakhir), getTimestampMs(b.updatedAt), getTimestampMs(b.createdAt));
    return tB - tA;
  });
};

export default function AdminPanel({
  onClose,
  globalLogoUrl,
  onUploadGlobalLogo,
  onResetGlobalLogo,
  adminEmail,
}: AdminPanelProps) {
  const { user: authUser } = useAuth();
  const [firebaseAuthUser, setFirebaseAuthUser] = useState<User | null>(() => auth.currentUser);
  const firestoreUsersRef = useRef<any[]>([]);
  const hubUsersRef = useRef<any[]>([]);
  const firestoreDevicesRef = useRef<DeviceRecord[]>([]);
  const hubDevicesRef = useRef<DeviceRecord[]>([]);
  const pendingDeletedEmailsRef = useRef<Map<string, number>>(new Map());
  const isFetchingCloudRef = useRef<boolean>(false);

  const [users, setUsers] = useState<any[]>(() =>
    mergeDeduplicatedUsers({ localCached: readLocalCachedUsers() })
  );
  const [devices, setDevices] = useState<DeviceRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [firestoreSyncError, setFirestoreSyncError] = useState<string | null>(null);

  const [searchInput, setSearchInput] = useState("");
  const [appliedSearch, setAppliedSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState<StatusFilter>("semua");
  const [highlightedEmail, setHighlightedEmail] = useState<string | null>(null);

  const [drafts, setDrafts] = useState<Record<string, UserDraft>>({});
  const [savingEmail, setSavingEmail] = useState<string | null>(null);
  const [recentlySavedEmail, setRecentlySavedEmail] = useState<string | null>(null);
  const [toast, setToast] = useState<{ type: "success" | "error" | "info"; text: string } | null>(null);

  // Modals & compact drawers
  const [deleteTargetUser, setDeleteTargetUser] = useState<any | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isLogoModalOpen, setIsLogoModalOpen] = useState(false);
  const [showDevicesDrawer, setShowDevicesDrawer] = useState(false);

  const searchInputRef = useRef<HTMLInputElement | null>(null);

  // Track the live Firebase Auth instance directly (independent of local simulated sessions)
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => {
      setFirebaseAuthUser(u);
    });
    return () => unsub();
  }, []);

  const showToast = useCallback((type: "success" | "error" | "info", text: string) => {
    setToast({ type, text });
  }, []);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(null), 5000);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const recomputeUsers = useCallback(() => {
    const deviceFallbackUsers = synthesizeUsersFromDevices([
      ...firestoreDevicesRef.current,
      ...hubDevicesRef.current,
    ]);
    const merged = mergeDeduplicatedUsers({
      authoritative: [...firestoreUsersRef.current, ...hubUsersRef.current],
      deviceFallback: deviceFallbackUsers,
      localCached: readLocalCachedUsers(),
      pendingDeleted: pendingDeletedEmailsRef.current,
    });
    setUsers(merged);
  }, []);

  const recomputeDevices = useCallback(() => {
    const byId = new Map<string, DeviceRecord>();
    [...firestoreDevicesRef.current, ...hubDevicesRef.current].forEach((d) => {
      if (!d?.deviceId) return;
      const prev = byId.get(d.deviceId);
      if (!prev || getTimestampMs(d.lastSeenAt) >= getTimestampMs(prev.lastSeenAt)) {
        byId.set(d.deviceId, d);
      }
    });
    const sorted = Array.from(byId.values()).sort(
      (a, b) => getTimestampMs(b.lastSeenAt) - getTimestampMs(a.lastSeenAt)
    );
    setDevices(sorted);
    // Also reconcile any signed-in email reported via device heartbeat into the user list
    recomputeUsers();
  }, [recomputeUsers]);

  // Authoritative Cloud Firestore fetch via Direct Authenticated REST API v1 (+ getDocsFromServer fallback)
  // Bypasses stalled WebChannels / stale IndexedDB caches so mobile & Incognito logins appear immediately!
  const fetchFirestoreAuthoritative = useCallback(
    async (fbUser: User | null, forceTokenRefresh = false): Promise<{ ok: boolean; count: number }> => {
      if (!fbUser || isFetchingCloudRef.current) {
        return { ok: false, count: firestoreUsersRef.current.length };
      }
      isFetchingCloudRef.current = true;
      try {
        let idToken = await fbUser.getIdToken(forceTokenRefresh).catch(() => null);
        if (!idToken) {
          idToken = await fbUser.getIdToken(true).catch(() => null);
        }

        let [usersRes, devicesRes] = await Promise.all([
          firestoreRestListCollection("users", { idToken, timeoutMs: 6500 }),
          firestoreRestListCollection("devices", { idToken, timeoutMs: 6500 }),
        ]);

        // If 403 occurred due to stale token, force-refresh ID token and retry once
        if (!usersRes.ok && usersRes.permissionDenied && !forceTokenRefresh) {
          const freshToken = await fbUser.getIdToken(true).catch(() => null);
          if (freshToken) {
            [usersRes, devicesRes] = await Promise.all([
              firestoreRestListCollection("users", { idToken: freshToken, timeoutMs: 6500 }),
              firestoreRestListCollection("devices", { idToken: freshToken, timeoutMs: 6500 }),
            ]);
          }
        }

        let gotUsers = false;
        if (usersRes.ok && Array.isArray(usersRes.data)) {
          gotUsers = true;
          setFirestoreSyncError(null);
          firestoreUsersRef.current = usersRes.data
            .filter((d) => d && d.email)
            .map((d) => ({
              ...d,
              id: d.id,
              uid: d.uid || d.id,
            }));
        } else if (usersRes.permissionDenied) {
          setFirestoreSyncError(
            "Akses baca database cloud ditolak untuk akun ini. Pastikan Anda masuk Panel Admin menggunakan email Google Admin utama (bukan sesi lokal) lalu klik Muat Ulang Cloud."
          );
        } else if (usersRes.quotaExceeded) {
          setFirestoreSyncError(
            "Kuota harian Firestore (Free Tier) sedang penuh sementara. Klik Muat Ulang Cloud lagi nanti atau tinjau kuota Firebase Console."
          );
        } else {
          // Fallback to Firestore SDK server fetch when direct REST encounters transient network hiccup
          try {
            const sdkUsersSnap = await fetchWithTimeout(
              getDocsFromServer(collection(db, "users")),
              5500,
              null
            );
            if (sdkUsersSnap) {
              gotUsers = true;
              setFirestoreSyncError(null);
              const list: any[] = [];
              sdkUsersSnap.forEach((docSnap) => {
                const data = docSnap.data();
                if (data && data.email) {
                  list.push({ id: docSnap.id, uid: data.uid || docSnap.id, ...data });
                }
              });
              firestoreUsersRef.current = list;
            }
          } catch (err: any) {
            setFirestoreSyncError(
              `Sinkronisasi Cloud Firestore sementara terganggu (${err?.code || err?.message || "offline"}). Klik Muat Ulang Cloud untuk mencoba lagi.`
            );
          }
        }

        if (devicesRes.ok && Array.isArray(devicesRes.data)) {
          firestoreDevicesRef.current = devicesRes.data.filter(
            (d): d is DeviceRecord => Boolean(d && d.deviceId)
          );
        } else if (!devicesRes.permissionDenied && !devicesRes.quotaExceeded) {
          try {
            const sdkDevicesSnap = await fetchWithTimeout(
              getDocsFromServer(collection(db, "devices")),
              4500,
              null
            );
            if (sdkDevicesSnap) {
              const list: DeviceRecord[] = [];
              sdkDevicesSnap.forEach((docSnap) => {
                const data = docSnap.data() as DeviceRecord;
                if (data && data.deviceId) list.push(data);
              });
              firestoreDevicesRef.current = list;
            }
          } catch {}
        }

        recomputeDevices();
        recomputeUsers();
        return { ok: gotUsers, count: firestoreUsersRef.current.length };
      } finally {
        isFetchingCloudRef.current = false;
      }
    },
    [recomputeDevices, recomputeUsers]
  );

  // Manual / on-demand refresh across BOTH Cloud Firestore (REST + SDK) and Real-Time Hub
  const handleManualRefresh = useCallback(
    async (silent = false) => {
      if (!silent) setIsRefreshing(true);
      try {
        const activeFbUser = firebaseAuthUser || auth.currentUser;
        const [hubSnap] = await Promise.all([
          fetch("/api/realtime/snapshot", { cache: "no-store" })
            .then((r) => (r.ok ? r.json() : null))
            .catch(() => null),
          activeFbUser ? fetchFirestoreAuthoritative(activeFbUser, !silent) : Promise.resolve(null),
        ]);

        if (hubSnap) {
          if (Array.isArray(hubSnap.users)) hubUsersRef.current = hubSnap.users;
          if (Array.isArray(hubSnap.devices)) hubDevicesRef.current = hubSnap.devices;
        }
        recomputeDevices();
        recomputeUsers();
        if (!silent) {
          const totalCount = mergeDeduplicatedUsers({
            authoritative: [...firestoreUsersRef.current, ...hubUsersRef.current],
            deviceFallback: synthesizeUsersFromDevices([
              ...firestoreDevicesRef.current,
              ...hubDevicesRef.current,
            ]),
            localCached: readLocalCachedUsers(),
            pendingDeleted: pendingDeletedEmailsRef.current,
          }).length;
          showToast(
            "success",
            `Daftar pengguna berhasil disegarkan dari Cloud (${totalCount} akun terdeteksi).`
          );
        }
      } finally {
        if (!silent) setIsRefreshing(false);
        setLoading(false);
      }
    },
    [firebaseAuthUser, fetchFirestoreAuthoritative, recomputeDevices, recomputeUsers, showToast]
  );

  // Triple-Channel Real-Time Sync:
  // 1. Direct Cloud Firestore REST authoritative fetch (on mount, window focus, tab return, and 15s poll)
  // 2. Cloud Firestore SDK onSnapshot real-time push listeners
  // 3. Express SSE Real-Time Hub (for same-origin instant notifications)
  useEffect(() => {
    let mounted = true;

    // Channel 1: Express Real-Time Snapshot + SSE Stream
    fetch("/api/realtime/snapshot", { cache: "no-store" })
      .then((res) => (res.ok ? res.json() : null))
      .then((snap) => {
        if (!mounted || !snap) return;
        if (Array.isArray(snap.users)) {
          hubUsersRef.current = snap.users;
          recomputeUsers();
        }
        if (Array.isArray(snap.devices)) {
          hubDevicesRef.current = snap.devices;
          recomputeDevices();
        }
        if (!firebaseAuthUser) setLoading(false);
      })
      .catch(() => {
        if (mounted && !firebaseAuthUser) setLoading(false);
      });

    let sse: EventSource | null = null;
    try {
      sse = new EventSource("/api/realtime/stream");
      sse.addEventListener("snapshot", (ev: MessageEvent) => {
        if (!mounted) return;
        try {
          const data = JSON.parse(ev.data);
          if (Array.isArray(data?.users)) {
            hubUsersRef.current = data.users;
            recomputeUsers();
          }
          if (Array.isArray(data?.devices)) {
            hubDevicesRef.current = data.devices;
            recomputeDevices();
          }
          setLoading(false);
        } catch {}
      });
      sse.addEventListener("user_upsert", (ev: MessageEvent) => {
        if (!mounted) return;
        try {
          const incoming = JSON.parse(ev.data);
          if (!incoming?.email) return;
          const em = String(incoming.email).toLowerCase().trim();
          removeDeletedTenantEmail(em);
          pendingDeletedEmailsRef.current.delete(em);
          const next = hubUsersRef.current.filter(
            (u) => String(u.email || "").toLowerCase().trim() !== em && u.uid !== incoming.uid
          );
          next.unshift(incoming);
          hubUsersRef.current = next;
          recomputeUsers();
        } catch {}
      });
      sse.addEventListener("user_delete", (ev: MessageEvent) => {
        if (!mounted) return;
        try {
          const payload = JSON.parse(ev.data);
          const delEmail = String(payload?.email || "").toLowerCase().trim();
          const delUid = String(payload?.uid || "");
          hubUsersRef.current = hubUsersRef.current.filter(
            (u) =>
              !(
                (delUid && (u.uid === delUid || u.id === delUid)) ||
                (delEmail && String(u.email || "").toLowerCase().trim() === delEmail)
              )
          );
          recomputeUsers();
        } catch {}
      });
      sse.addEventListener("device_upsert", (ev: MessageEvent) => {
        if (!mounted) return;
        try {
          const incoming = JSON.parse(ev.data);
          if (!incoming?.deviceId) return;
          hubDevicesRef.current = [
            incoming,
            ...hubDevicesRef.current.filter((d) => d.deviceId !== incoming.deviceId),
          ];
          recomputeDevices();
        } catch {}
      });
      sse.addEventListener("device_delete", (ev: MessageEvent) => {
        if (!mounted) return;
        try {
          const payload = JSON.parse(ev.data);
          if (payload?.deviceId) {
            hubDevicesRef.current = hubDevicesRef.current.filter(
              (d) => d.deviceId !== payload.deviceId
            );
            recomputeDevices();
          }
        } catch {}
      });
    } catch {}

    // Channel 2 & 3: Cloud Firestore Authoritative REST Fetch + SDK onSnapshot + Focus/Interval Safety Net
    let unsubUsers: (() => void) | null = null;
    let unsubDevices: (() => void) | null = null;
    let pollTimer: number | null = null;
    let handleWindowFocus: (() => void) | null = null;
    let handleVisibilityChange: (() => void) | null = null;

    if (firebaseAuthUser?.uid) {
      // Immediate authoritative fetch on mount so Admin never waits on WebChannel negotiation
      fetchFirestoreAuthoritative(firebaseAuthUser, false).finally(() => {
        if (mounted) setLoading(false);
      });

      // Live real-time push listeners via Firestore SDK
      try {
        unsubUsers = onSnapshot(
          collection(db, "users"),
          (snapshot) => {
            if (!mounted) return;
            setFirestoreSyncError(null);
            const list: any[] = [];
            snapshot.forEach((docSnap) => {
              const data = docSnap.data();
              if (data && data.email) {
                list.push({ id: docSnap.id, uid: data.uid || docSnap.id, ...data });
              }
            });
            firestoreUsersRef.current = list;
            recomputeUsers();
            setLoading(false);
          },
          (err) => {
            console.warn("Firestore users listener notice:", err);
            // Fall back smoothly to authoritative REST fetch when SDK stream hits hiccup
            if (mounted) {
              fetchFirestoreAuthoritative(firebaseAuthUser, false).finally(() => {
                if (mounted) setLoading(false);
              });
            }
          }
        );
      } catch {
        if (mounted) setLoading(false);
      }

      try {
        unsubDevices = onSnapshot(
          collection(db, "devices"),
          (snapshot) => {
            if (!mounted) return;
            const list: DeviceRecord[] = [];
            snapshot.forEach((docSnap) => {
              const data = docSnap.data() as DeviceRecord;
              if (data && data.deviceId) list.push(data);
            });
            firestoreDevicesRef.current = list;
            recomputeDevices();
          },
          (err) => {
            console.warn("Firestore devices listener notice:", err);
          }
        );
      } catch {}

      // Refresh immediately when Admin returns to laptop/tab after testing on HP or Incognito
      handleWindowFocus = () => {
        if (mounted) fetchFirestoreAuthoritative(firebaseAuthUser, false);
      };
      handleVisibilityChange = () => {
        if (mounted && document.visibilityState === "visible") {
          fetchFirestoreAuthoritative(firebaseAuthUser, false);
        }
      };
      window.addEventListener("focus", handleWindowFocus);
      document.addEventListener("visibilitychange", handleVisibilityChange);

      // Gentle 15-second visible background poll so even if a browser throttles WebChannel, new logins appear automatically
      pollTimer = window.setInterval(() => {
        if (mounted && document.visibilityState === "visible") {
          fetchFirestoreAuthoritative(firebaseAuthUser, false);
        }
      }, 15_000);
    } else if (authUser?.uid) {
      setLoading(false);
    }

    return () => {
      mounted = false;
      if (sse) sse.close();
      if (unsubUsers) unsubUsers();
      if (unsubDevices) unsubDevices();
      if (pollTimer !== null) window.clearInterval(pollTimer);
      if (handleWindowFocus) window.removeEventListener("focus", handleWindowFocus);
      if (handleVisibilityChange) document.removeEventListener("visibilitychange", handleVisibilityChange);
    };
  }, [firebaseAuthUser, authUser?.uid, fetchFirestoreAuthoritative, recomputeUsers, recomputeDevices]);

  const todayIso = getTodayIsoLocal();

  // Compute summary stats
  const stats = useMemo(() => {
    let menunggu = 0;
    let aktif = 0;
    let kedaluwarsa = 0;
    let diblokir = 0;
    users.forEach((u) => {
      const cat = getEffectiveCategory(u, todayIso);
      if (cat === "menunggu") menunggu++;
      else if (cat === "aktif") aktif++;
      else if (cat === "kedaluwarsa") kedaluwarsa++;
      else if (cat === "diblokir") diblokir++;
    });
    return {
      total: users.length,
      menunggu,
      aktif,
      kedaluwarsa,
      diblokir,
    };
  }, [users, todayIso]);

  // Get draft values for a user (falling back to persisted user fields)
  const getUserDraft = useCallback(
    (u: any): UserDraft => {
      const key = String(u.email || "").toLowerCase().trim();
      if (drafts[key]) return drafts[key];
      const defaultNote =
        getEffectiveCategory(u, todayIso) === "menunggu" && !u.keterangan
          ? DEFAULT_WAITING_NOTE
          : String(u.keterangan ?? "");
      return {
        mulaiPada: typeof u.mulaiPada === "string" ? u.mulaiPada.trim() : "",
        berakhirPada: typeof u.berakhirPada === "string" ? u.berakhirPada.trim() : "",
        keterangan: defaultNote,
      };
    },
    [drafts, todayIso]
  );

  const isDraftDirty = useCallback(
    (u: any): boolean => {
      const key = String(u.email || "").toLowerCase().trim();
      const d = drafts[key];
      if (!d) return false;
      const origMulai = typeof u.mulaiPada === "string" ? u.mulaiPada.trim() : "";
      const origBerakhir = typeof u.berakhirPada === "string" ? u.berakhirPada.trim() : "";
      const origKet =
        getEffectiveCategory(u, todayIso) === "menunggu" && !u.keterangan
          ? DEFAULT_WAITING_NOTE
          : String(u.keterangan ?? "");
      return (
        d.mulaiPada !== origMulai ||
        d.berakhirPada !== origBerakhir ||
        d.keterangan !== origKet
      );
    },
    [drafts, todayIso]
  );

  const updateUserDraft = (u: any, patch: Partial<UserDraft>) => {
    const key = String(u.email || "").toLowerCase().trim();
    const current = getUserDraft(u);
    setDrafts((prev) => ({
      ...prev,
      [key]: {
        ...current,
        ...patch,
      },
    }));
  };

  // Core persistence helper: saves user access changes simultaneously to Local, Real-Time Hub (+ auto Webhook), and Firestore
  const persistUserAccessUpdate = async (
    targetUser: any,
    updates: {
      statusPersetujuan?: "aktif" | "menunggu" | "diblokir";
      mulaiPada?: string;
      berakhirPada?: string;
      keterangan?: string;
    },
    successMessage: string
  ) => {
    const email = String(targetUser.email || "").toLowerCase().trim();
    if (!email) return;

    setSavingEmail(email);
    const nowIso = new Date().toISOString();
    const docIds: string[] = Array.from(
      new Set([targetUser.uid || targetUser.id, ...(targetUser._allIds || [])].filter(Boolean))
    );

    const nextMulai =
      updates.mulaiPada !== undefined
        ? updates.mulaiPada.trim()
        : typeof targetUser.mulaiPada === "string"
        ? targetUser.mulaiPada.trim()
        : "";
    const nextBerakhir =
      updates.berakhirPada !== undefined
        ? updates.berakhirPada.trim()
        : typeof targetUser.berakhirPada === "string"
        ? targetUser.berakhirPada.trim()
        : "";
    const nextStatus = updates.statusPersetujuan || targetUser.statusPersetujuan || "menunggu";
    const nextKeterangan =
      updates.keterangan !== undefined
        ? updates.keterangan.trim()
        : typeof targetUser.keterangan === "string"
        ? targetUser.keterangan.trim()
        : "";

    const updatedRecord: Record<string, any> = {
      ...targetUser,
      uid: targetUser.uid || targetUser.id,
      id: targetUser.uid || targetUser.id,
      email,
      statusPersetujuan: nextStatus,
      mulaiPada: nextMulai,
      berakhirPada: nextBerakhir,
      keterangan: nextKeterangan,
      updatedAt: nowIso,
    };
    delete updatedRecord._allIds;
    delete updatedRecord._fromDeviceFallback;
    removeDeletedTenantEmail(email);

    try {
      // Optimistically update local refs & state so UI feels instantaneous
      const applyToList = (list: any[]) => {
        const filtered = list.filter(
          (item) => String(item.email || "").toLowerCase().trim() !== email
        );
        return [updatedRecord, ...filtered];
      };
      hubUsersRef.current = applyToList(hubUsersRef.current);
      firestoreUsersRef.current = applyToList(firestoreUsersRef.current);
      try {
        localStorage.setItem(`offline_user_${updatedRecord.uid}`, JSON.stringify(updatedRecord));
      } catch {}
      recomputeUsers();

      // Clear draft so UI reflects authoritative saved state cleanly
      setDrafts((prev) => {
        const next = { ...prev };
        delete next[email];
        return next;
      });

      // Push to Express Real-Time Hub (+ optional Apps Script Webhook) AND Cloud Firestore (Direct REST + SDK fallback)
      const activeFbUser = firebaseAuthUser || auth.currentUser;
      const idToken = activeFbUser ? await activeFbUser.getIdToken(false).catch(() => null) : null;
      let anyCloudWriteFailed = false;
      let lastCloudErr: string | null = null;

      await Promise.all([
        syncUserToRealtimeHub(updatedRecord),
        ...docIds.map(async (docId) => {
          const cleanPayload = stripUndefinedDeep({
            ...updatedRecord,
            uid: docId,
            id: docId,
            updatedAt: nowIso,
          });
          // Primary fast path: Direct Authenticated Cloud Firestore REST PATCH (keepalive)
          if (idToken) {
            const restRes = await firestoreRestUpsertDoc("users", docId, cleanPayload, {
              idToken,
              keepalive: true,
              timeoutMs: 5500,
            });
            if (restRes.ok) return;
            lastCloudErr = restRes.error || null;
          }
          // Fallback path: Firestore SDK setDoc
          const sdkOk = await fetchWithTimeout(
            setDoc(
              doc(db, "users", docId),
              stripUndefinedDeep({
                ...cleanPayload,
                updatedAt: serverTimestamp(),
              }),
              { merge: true }
            )
              .then(() => {
                reportFirestoreWriteResult(true);
                return true;
              })
              .catch((err) => {
                reportFirestoreWriteResult(false, err);
                lastCloudErr = err?.message || String(err);
                return false;
              }),
            5000,
            false
          );
          if (!sdkOk) anyCloudWriteFailed = true;
        }),
      ]);

      setRecentlySavedEmail(email);
      window.setTimeout(() => {
        setRecentlySavedEmail((prev) => (prev === email ? null : prev));
      }, 2400);
      if (anyCloudWriteFailed && activeFbUser) {
        showToast(
          "error",
          `Perubahan lokal tersimpan, tetapi gagal sinkron ke Cloud Firestore (${lastCloudErr || "jaringan"}). Klik Muat Ulang Cloud lalu ulangi.`
        );
      } else {
        showToast("success", successMessage);
      }
    } catch (err: any) {
      showToast("error", err?.message || `Gagal menyimpan perubahan untuk ${email}.`);
    } finally {
      setSavingEmail(null);
    }
  };

  // Quick ACC action
  const handleAccUser = (u: any) => {
    const draft = getUserDraft(u);
    let nextBerakhir = draft.berakhirPada;
    // If user was expired and Admin clicked ACC without changing past expiration date, clear to Tanpa Batas
    if (nextBerakhir && todayIso > nextBerakhir) {
      nextBerakhir = "";
    }
    const nextKet =
      draft.keterangan === DEFAULT_WAITING_NOTE || !draft.keterangan.trim()
        ? "Akses disetujui (ACC) oleh Admin Odhe Gizi"
        : draft.keterangan;

    persistUserAccessUpdate(
      u,
      {
        statusPersetujuan: "aktif",
        mulaiPada: draft.mulaiPada,
        berakhirPada: nextBerakhir,
        keterangan: nextKet,
      },
      `Pengguna ${u.email} berhasil di-ACC dan langsung aktif.`
    );
  };

  // Return user to Daftar Tunggu
  const handleReturnToWaiting = (u: any) => {
    const draft = getUserDraft(u);
    const nextKet =
      !draft.keterangan.trim() ||
      draft.keterangan === "Akses disetujui (ACC) oleh Admin Odhe Gizi"
        ? DEFAULT_WAITING_NOTE
        : draft.keterangan;

    persistUserAccessUpdate(
      u,
      {
        statusPersetujuan: "menunggu",
        mulaiPada: draft.mulaiPada,
        berakhirPada: draft.berakhirPada,
        keterangan: nextKet,
      },
      `${u.email} dikembalikan ke Daftar Tunggu.`
    );
  };

  // Block user
  const handleBlockUser = (u: any) => {
    const draft = getUserDraft(u);
    const nextKet =
      !draft.keterangan.trim() ||
      draft.keterangan === DEFAULT_WAITING_NOTE ||
      draft.keterangan === "Akses disetujui (ACC) oleh Admin Odhe Gizi"
        ? "Akses diblokir oleh Admin Odhe Gizi"
        : draft.keterangan;

    persistUserAccessUpdate(
      u,
      {
        statusPersetujuan: "diblokir",
        mulaiPada: draft.mulaiPada,
        berakhirPada: draft.berakhirPada,
        keterangan: nextKet,
      },
      `Akses akun ${u.email} telah diblokir.`
    );
  };

  // Save draft changes (dates + keterangan)
  const handleSaveDraft = (u: any) => {
    const draft = getUserDraft(u);
    persistUserAccessUpdate(
      u,
      {
        mulaiPada: draft.mulaiPada,
        berakhirPada: draft.berakhirPada,
        keterangan: draft.keterangan,
      },
      `Pengaturan tanggal & keterangan untuk ${u.email} berhasil disimpan.`
    );
  };

  // Apply quick duration preset (+7 Hr, +30 Hr, +3 Bln, +1 Thn) or Tanpa Batas
  const handleApplyExpirationPreset = (
    u: any,
    preset: "unlimited" | "7d" | "30d" | "3m" | "1y"
  ) => {
    const draft = getUserDraft(u);
    const nextBerakhir =
      preset === "unlimited"
        ? ""
        : preset === "7d"
        ? addDurationToDateIso(draft.mulaiPada, { days: 7 })
        : preset === "30d"
        ? addDurationToDateIso(draft.mulaiPada, { days: 30 })
        : preset === "3m"
        ? addDurationToDateIso(draft.mulaiPada, { months: 3 })
        : addDurationToDateIso(draft.mulaiPada, { years: 1 });

    updateUserDraft(u, { berakhirPada: nextBerakhir });

    // If user is already active, save immediately so Admin doesn't need an extra click
    if (String(u.statusPersetujuan || "").toLowerCase() === "aktif") {
      const label =
        preset === "unlimited"
          ? "Tanpa Batas"
          : `berakhir ${nextBerakhir}`;
      persistUserAccessUpdate(
        u,
        {
          mulaiPada: draft.mulaiPada,
          berakhirPada: nextBerakhir,
          keterangan: draft.keterangan,
        },
        `Masa aktif ${u.email} diperbarui (${label}).`
      );
    }
  };

  // Delete user record permanently (confirmed via custom modal — no window.confirm)
  const handleConfirmDeleteUser = async () => {
    if (!deleteTargetUser) return;
    const target = deleteTargetUser;
    const email = String(target.email || "").toLowerCase().trim();
    const docIds: string[] = Array.from(
      new Set([target.uid || target.id, ...(target._allIds || [])].filter(Boolean))
    );

    setIsDeleting(true);
    try {
      const deleteTs = Date.now();
      pendingDeletedEmailsRef.current.set(email, deleteTs);
      addDeletedTenantEmail(email);

      // Identify all device logs tied to this email so synthesizeUsersFromDevices never resurrects the deleted user
      const matchingDeviceIds = Array.from(
        new Set(
          [...firestoreDevicesRef.current, ...hubDevicesRef.current]
            .filter((d) => String(d?.userEmail || "").toLowerCase().trim() === email)
            .map((d) => d.deviceId)
            .filter(Boolean)
        )
      );

      hubUsersRef.current = hubUsersRef.current.filter(
        (u) => String(u.email || "").toLowerCase().trim() !== email
      );
      firestoreUsersRef.current = firestoreUsersRef.current.filter(
        (u) => String(u.email || "").toLowerCase().trim() !== email
      );
      if (matchingDeviceIds.length > 0) {
        const delDevSet = new Set(matchingDeviceIds);
        hubDevicesRef.current = hubDevicesRef.current.filter((d) => !delDevSet.has(d.deviceId));
        firestoreDevicesRef.current = firestoreDevicesRef.current.filter(
          (d) => !delDevSet.has(d.deviceId)
        );
      }
      docIds.forEach((id) => {
        try {
          localStorage.removeItem(`offline_user_${id}`);
        } catch {}
      });
      recomputeDevices();
      recomputeUsers();

      const activeFbUser = firebaseAuthUser || auth.currentUser;
      const idToken = activeFbUser ? await activeFbUser.getIdToken(false).catch(() => null) : null;

      await Promise.all([
        deleteUserFromRealtimeHub(target.uid || target.id, email),
        ...docIds.map(async (id) => {
          if (idToken) {
            const restDel = await firestoreRestDeleteDoc("users", id, { idToken, timeoutMs: 4500 });
            if (restDel.ok) return;
          }
          await fetchWithTimeout(
            deleteDoc(doc(db, "users", id))
              .then(() => reportFirestoreWriteResult(true))
              .catch((err) => reportFirestoreWriteResult(false, err)),
            3500,
            null
          );
        }),
        ...matchingDeviceIds.map(async (devId) => {
          await deleteDeviceFromRealtimeHub(devId);
          if (idToken) {
            const restDevDel = await firestoreRestDeleteDoc("devices", devId, {
              idToken,
              timeoutMs: 4000,
            });
            if (restDevDel.ok) return;
          }
          await fetchWithTimeout(deleteDoc(doc(db, "devices", devId)).catch(() => null), 3000, null);
        }),
      ]);

      setDeleteTargetUser(null);
      showToast("success", `Catatan pengguna ${email} berhasil dihapus dari Panel Admin.`);
    } catch (err: any) {
      showToast("error", err?.message || `Gagal menghapus catatan ${email}.`);
    } finally {
      setIsDeleting(false);
    }
  };

  // Handle "Cari Email untuk ACC" button / form submission
  const handleSearchSubmit = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const q = searchInput.trim();
    setAppliedSearch(q);

    if (!q) {
      setHighlightedEmail(null);
      searchInputRef.current?.focus();
      return;
    }

    const lower = q.toLowerCase();
    const matched = users.find(
      (u) =>
        String(u.email || "").toLowerCase().includes(lower) ||
        String(u.namaLengkap || "").toLowerCase().includes(lower) ||
        String(u.namaSPPG || "").toLowerCase().includes(lower)
    );

    if (matched) {
      const cat = getEffectiveCategory(matched, todayIso);
      if (statusFilter !== "semua" && statusFilter !== cat) {
        setStatusFilter("semua");
      }
      setHighlightedEmail(String(matched.email).toLowerCase());
      window.setTimeout(() => setHighlightedEmail(null), 3200);
    } else {
      setStatusFilter("semua");
      showToast("info", `Email atau nama "${q}" belum terdeteksi di database pengguna.`);
    }
  };

  // Merge imported users from Google Sheets directly into view
  const handleUsersImportedFromSheet = useCallback(
    (importedUsers: any[]) => {
      if (!importedUsers.length) return;
      importedUsers.forEach((u) => {
        if (u.email) {
          const clean = String(u.email).toLowerCase().trim();
          removeDeletedTenantEmail(clean);
          pendingDeletedEmailsRef.current.delete(clean);
        }
      });
      hubUsersRef.current = mergeDeduplicatedUsers({
        authoritative: [...importedUsers, ...hubUsersRef.current],
      });
      recomputeUsers();
    },
    [recomputeUsers]
  );

  // Filtered user list (live filter on typing + applied search)
  const effectiveQuery = (searchInput || appliedSearch).trim().toLowerCase();
  const filteredUsers = useMemo(() => {
    return users.filter((u) => {
      const cat = getEffectiveCategory(u, todayIso);
      if (statusFilter !== "semua" && cat !== statusFilter) return false;
      if (!effectiveQuery) return true;
      return (
        String(u.email || "").toLowerCase().includes(effectiveQuery) ||
        String(u.namaLengkap || "").toLowerCase().includes(effectiveQuery) ||
        String(u.namaSPPG || "").toLowerCase().includes(effectiveQuery) ||
        String(u.profesi || "").toLowerCase().includes(effectiveQuery)
      );
    });
  }, [users, statusFilter, effectiveQuery, todayIso]);

  const onlineDevicesCount = useMemo(
    () => devices.filter((d) => isDeviceCurrentlyOnline(d)).length,
    [devices]
  );

  const handleDeleteDevice = async (deviceId: string) => {
    hubDevicesRef.current = hubDevicesRef.current.filter((d) => d.deviceId !== deviceId);
    firestoreDevicesRef.current = firestoreDevicesRef.current.filter((d) => d.deviceId !== deviceId);
    recomputeDevices();
    const activeFbUser = firebaseAuthUser || auth.currentUser;
    const idToken = activeFbUser ? await activeFbUser.getIdToken(false).catch(() => null) : null;
    await Promise.all([
      deleteDeviceFromRealtimeHub(deviceId),
      idToken
        ? firestoreRestDeleteDoc("devices", deviceId, { idToken, timeoutMs: 4000 }).then((r) => {
            if (!r.ok) {
              return fetchWithTimeout(deleteDoc(doc(db, "devices", deviceId)).catch(() => null), 2500, null);
            }
          })
        : fetchWithTimeout(deleteDoc(doc(db, "devices", deviceId)).catch(() => null), 2500, null),
    ]);
  };

  const displayAdminEmail = adminEmail?.trim() || "Admin Utama";

  return (
    <div className="min-h-full bg-[#F8FAFC] -m-6 p-4 sm:p-6 lg:p-8">
      <div className="max-w-[1380px] mx-auto space-y-5">
        {/* TOP HEADER CARD — Matches Reference Layout */}
        <header className="bg-white rounded-2xl border border-slate-200 px-5 py-5 sm:px-6 sm:py-6 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="min-w-0">
            <div className="flex flex-wrap items-center gap-x-2 gap-y-1 text-xs">
              <span className="font-bold text-emerald-700">Panel Kontrol Admin Odhe Gizi</span>
              <span className="text-slate-300">·</span>
              <span className="font-medium text-slate-500">{displayAdminEmail}</span>
            </div>
            <h1 className="text-xl sm:text-2xl font-extrabold text-slate-900 tracking-tight mt-1">
              Manajemen Akses Pengguna &amp; Database Spreadsheet
            </h1>
            <p className="text-sm text-slate-600 mt-1.5 max-w-3xl leading-relaxed">
              Setiap pengguna yang mencoba masuk dengan email pribadi Google otomatis terdeteksi di Daftar Tunggu. Cari email untuk melakukan ACC, atur tanggal login &amp; tanggal berakhir login, serta sinkronkan jumlah pengguna ke Google Spreadsheet.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              type="button"
              disabled={isRefreshing}
              onClick={() => handleManualRefresh(false)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 rounded-xl text-xs font-bold whitespace-nowrap transition cursor-pointer disabled:opacity-60"
            >
              <RefreshCw className={`w-3.5 h-3.5 text-emerald-600 ${isRefreshing ? "animate-spin" : ""}`} />
              <span>{isRefreshing ? "Menyegarkan..." : "Segarkan"}</span>
            </button>

            <button
              type="button"
              onClick={() => downloadUsersCsv(users)}
              className="inline-flex items-center gap-2 px-4 py-2.5 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 rounded-xl text-xs font-bold whitespace-nowrap transition cursor-pointer"
            >
              <Download className="w-4 h-4 text-emerald-600" />
              <span className="tabular-nums">Unduh Database CSV ({stats.total})</span>
            </button>

            {onUploadGlobalLogo && (
              <button
                type="button"
                onClick={() => setIsLogoModalOpen(true)}
                className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 rounded-xl text-xs font-bold whitespace-nowrap transition cursor-pointer"
              >
                <Camera className="w-3.5 h-3.5 text-slate-500" />
                <span>Logo Global</span>
              </button>
            )}

            <button
              type="button"
              onClick={onClose}
              className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold whitespace-nowrap transition cursor-pointer"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Dashboard</span>
            </button>
          </div>
        </header>

        {/* ADMIN CLOUD AUTH NOTICE (Shown only if Admin is on a local/offline session without Firebase Auth token) */}
        {!firebaseAuthUser && (
          <div className="rounded-xl px-4 py-3 text-xs font-medium flex flex-wrap items-center justify-between gap-3 border bg-amber-50 border-amber-300 text-amber-950">
            <div className="flex items-start gap-2">
              <AlertCircle className="w-4 h-4 text-amber-700 shrink-0 mt-0.5" />
              <span>
                Sesi Admin saat ini belum terhubung ke autentikasi Google Cloud. Untuk melihat pengguna yang masuk dari HP, mode samaran (Incognito), atau browser lain secara real-time, silakan keluar lalu masuk kembali menggunakan <strong>email Google Admin utama</strong>.
              </span>
            </div>
          </div>
        )}

        {/* FIRESTORE SYNC WARNING BANNER (Only shown if cloud listener reports an issue) */}
        {firestoreSyncError && (
          <div className="rounded-xl px-4 py-3 text-xs font-medium flex flex-wrap items-center justify-between gap-3 border bg-amber-50 border-amber-200 text-amber-900">
            <div className="flex items-center gap-2">
              <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
              <span>{firestoreSyncError}</span>
            </div>
            <button
              type="button"
              disabled={isRefreshing}
              onClick={() => handleManualRefresh(false)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-lg text-[11px] cursor-pointer disabled:opacity-60"
            >
              <RefreshCw className={`w-3 h-3 ${isRefreshing ? "animate-spin" : ""}`} />
              <span>Muat Ulang Cloud</span>
            </button>
          </div>
        )}

        {/* INLINE TOAST NOTIFICATION */}
        {toast && (
          <div
            className={`rounded-xl px-4 py-3 text-xs font-medium flex items-center justify-between gap-3 border ${
              toast.type === "success"
                ? "bg-emerald-50 border-emerald-200 text-emerald-900"
                : toast.type === "error"
                ? "bg-rose-50 border-rose-200 text-rose-900"
                : "bg-slate-100 border-slate-200 text-slate-800"
            }`}
          >
            <div className="flex items-center gap-2">
              {toast.type === "success" ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              ) : (
                <AlertCircle className="w-4 h-4 shrink-0" />
              )}
              <span>{toast.text}</span>
            </div>
            <button
              type="button"
              onClick={() => setToast(null)}
              className="opacity-60 hover:opacity-100 cursor-pointer shrink-0"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* 4 CLICKABLE KPI SUMMARY CARDS — Exact Reference Structure */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <button
            type="button"
            onClick={() => setStatusFilter("semua")}
            className={`text-left bg-white rounded-2xl p-5 border transition cursor-pointer ${
              statusFilter === "semua"
                ? "border-emerald-500 ring-1 ring-emerald-500/20"
                : "border-slate-200 hover:border-slate-300"
            }`}
          >
            <div className="text-xs font-semibold text-slate-600">Total Pengguna Terdeteksi</div>
            <div className="text-3xl font-extrabold text-slate-900 tabular-nums mt-2 leading-none">
              {stats.total}
            </div>
            <div className="text-xs text-slate-500 mt-2.5">Semua email pribadi masuk</div>
          </button>

          <button
            type="button"
            onClick={() => setStatusFilter("menunggu")}
            className={`text-left bg-white rounded-2xl p-5 border transition cursor-pointer ${
              statusFilter === "menunggu"
                ? "border-amber-500 ring-1 ring-amber-500/20"
                : "border-slate-200 hover:border-slate-300"
            }`}
          >
            <div className="text-xs font-semibold text-amber-700">Daftar Tunggu (Perlu ACC)</div>
            <div className="text-3xl font-extrabold text-amber-700 tabular-nums mt-2 leading-none">
              {stats.menunggu}
            </div>
            <div className="text-xs text-slate-500 mt-2.5">Menunggu persetujuan Odhe Gizi</div>
          </button>

          <button
            type="button"
            onClick={() => setStatusFilter("aktif")}
            className={`text-left bg-white rounded-2xl p-5 border transition cursor-pointer ${
              statusFilter === "aktif"
                ? "border-emerald-500 ring-1 ring-emerald-500/20"
                : "border-slate-200 hover:border-slate-300"
            }`}
          >
            <div className="text-xs font-semibold text-emerald-700">Pengguna Aktif (ACC)</div>
            <div className="text-3xl font-extrabold text-emerald-700 tabular-nums mt-2 leading-none">
              {stats.aktif}
            </div>
            <div className="text-xs text-slate-500 mt-2.5">Bisa akses semua menu non-admin</div>
          </button>

          <button
            type="button"
            onClick={() =>
              setStatusFilter(stats.kedaluwarsa > 0 ? "kedaluwarsa" : "diblokir")
            }
            className={`text-left bg-white rounded-2xl p-5 border transition cursor-pointer ${
              statusFilter === "kedaluwarsa" || statusFilter === "diblokir"
                ? "border-rose-500 ring-1 ring-rose-500/20"
                : "border-slate-200 hover:border-slate-300"
            }`}
          >
            <div className="text-xs font-semibold text-rose-700">Kedaluwarsa / Diblokir</div>
            <div className="text-3xl font-extrabold text-rose-700 tabular-nums mt-2 leading-none">
              {stats.kedaluwarsa + stats.diblokir}
            </div>
            <div className="text-xs text-slate-500 mt-2.5 tabular-nums">
              {stats.kedaluwarsa} habis masa · {stats.diblokir} diblokir
            </div>
          </button>
        </section>

        {/* MAIN USER ACCESS CONTROL WORKSPACE CARD */}
        <section className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-5 space-y-4">
          {/* Search Bar + Cari Email untuk ACC + Segmented Status Filter Tabs */}
          <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-3">
            <form onSubmit={handleSearchSubmit} className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 flex-1 max-w-2xl">
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                <input
                  ref={searchInputRef}
                  type="text"
                  value={searchInput}
                  onChange={(e) => setSearchInput(e.target.value)}
                  placeholder="Ketik email atau nama pengguna yang ingin di-ACC..."
                  className="w-full pl-10 pr-8 py-2.5 bg-white border border-slate-300 rounded-xl text-xs sm:text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
                />
                {searchInput && (
                  <button
                    type="button"
                    onClick={() => {
                      setSearchInput("");
                      setAppliedSearch("");
                    }}
                    aria-label="Hapus pencarian"
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>

              <button
                type="submit"
                className="inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs sm:text-sm rounded-xl whitespace-nowrap transition cursor-pointer shrink-0"
              >
                <Search className="w-4 h-4" />
                <span>Cari Email untuk ACC</span>
              </button>
            </form>

            {/* Segmented Status Tabs */}
            <div className="bg-slate-100 p-1 rounded-xl flex items-center gap-0.5 overflow-x-auto shrink-0">
              {(
                [
                  { id: "semua", label: "Semua", count: stats.total },
                  { id: "menunggu", label: "Daftar Tunggu", count: stats.menunggu },
                  { id: "aktif", label: "Aktif", count: stats.aktif },
                  { id: "kedaluwarsa", label: "Kedaluwarsa", count: stats.kedaluwarsa },
                  { id: "diblokir", label: "Diblokir", count: stats.diblokir },
                ] as const
              ).map((tab) => {
                const isActive = statusFilter === tab.id;
                return (
                  <button
                    key={tab.id}
                    type="button"
                    onClick={() => setStatusFilter(tab.id)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold whitespace-nowrap tabular-nums transition cursor-pointer ${
                      isActive
                        ? "bg-white text-slate-900 shadow-2xs"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    {tab.label} ({tab.count})
                  </button>
                );
              })}
            </div>
          </div>

          {/* USER ACCESS CONTROL CARDS */}
          {loading && users.length === 0 ? (
            <div className="space-y-3 py-2">
              {[1, 2, 3].map((idx) => (
                <div
                  key={idx}
                  className="border border-slate-200 rounded-2xl p-5 animate-pulse space-y-4"
                >
                  <div className="flex items-center justify-between">
                    <div className="space-y-2">
                      <div className="h-4 w-64 bg-slate-200 rounded" />
                      <div className="h-3 w-96 bg-slate-100 rounded" />
                    </div>
                    <div className="h-9 w-36 bg-slate-200 rounded-xl" />
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4 pt-3 border-t border-slate-100">
                    <div className="h-10 bg-slate-100 rounded-xl" />
                    <div className="h-10 bg-slate-100 rounded-xl" />
                    <div className="h-10 bg-slate-100 rounded-xl" />
                  </div>
                </div>
              ))}
            </div>
          ) : filteredUsers.length === 0 ? (
            <div className="border border-dashed border-slate-200 rounded-2xl py-12 px-6 text-center space-y-3">
              <p className="text-sm font-bold text-slate-800">
                {effectiveQuery
                  ? `Tidak ditemukan pengguna dengan kata kunci "${effectiveQuery}".`
                  : "Belum ada pengguna pada kategori filter ini."}
              </p>
              <p className="text-xs text-slate-500 max-w-md mx-auto">
                Setiap pengguna baru yang mencoba masuk dengan email Google akan otomatis muncul di Daftar Tunggu secara real-time.
              </p>
              {(effectiveQuery || statusFilter !== "semua") && (
                <button
                  type="button"
                  onClick={() => {
                    setSearchInput("");
                    setAppliedSearch("");
                    setStatusFilter("semua");
                  }}
                  className="inline-flex items-center gap-1.5 px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold rounded-xl cursor-pointer transition"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  <span>Tampilkan Semua Pengguna</span>
                </button>
              )}
            </div>
          ) : (
            <div className="space-y-3">
              {filteredUsers.map((u) => {
                const email = String(u.email || "").toLowerCase().trim();
                const cat = getEffectiveCategory(u, todayIso);
                const futureScheduled = isScheduledInFuture(u, todayIso);
                const isProtectedMainAdmin = isMainAdminEmail(email);
                const draft = getUserDraft(u);
                const dirty = isDraftDirty(u);
                const isSavingThis = savingEmail === email;
                const justSaved = recentlySavedEmail === email;
                const isHighlighted = highlightedEmail === email;

                const statusLabel =
                  cat === "menunggu"
                    ? "Daftar Tunggu (Menunggu ACC)"
                    : cat === "diblokir"
                    ? "Diblokir"
                    : cat === "kedaluwarsa"
                    ? "Kedaluwarsa"
                    : futureScheduled
                    ? "Aktif (Jadwal Belum Mulai)"
                    : "Aktif (ACC)";

                const statusColorClass =
                  cat === "menunggu"
                    ? "text-amber-700"
                    : cat === "diblokir" || cat === "kedaluwarsa"
                    ? "text-rose-700"
                    : futureScheduled
                    ? "text-indigo-700"
                    : "text-emerald-700";

                const loginCount = Math.max(1, Number(u.loginCount) || 1);

                return (
                  <article
                    key={email}
                    className={`rounded-2xl border p-4 sm:p-5 space-y-4 transition ${
                      isHighlighted
                        ? "border-emerald-500 ring-2 ring-emerald-500/20 bg-emerald-50/20"
                        : "border-slate-200 bg-white hover:border-slate-300"
                    }`}
                  >
                    {/* Top Row: Email · Status + Action Buttons */}
                    <div className="flex flex-col lg:flex-row lg:items-start justify-between gap-3">
                      <div className="min-w-0">
                        <div className="flex flex-wrap items-baseline gap-x-1.5 gap-y-0.5">
                          <span className="text-sm sm:text-[15px] font-extrabold text-slate-900 break-all">
                            {email}
                          </span>
                          <span className="text-slate-300">·</span>
                          <span className={`text-xs font-bold ${statusColorClass}`}>
                            {statusLabel}
                          </span>
                          {u.peran === "ADMIN" && (
                            <>
                              <span className="text-slate-300">·</span>
                              <span className="text-xs font-bold text-slate-600">Admin</span>
                            </>
                          )}
                        </div>

                        <p className="text-xs text-slate-500 mt-1 tabular-nums leading-relaxed">
                          <span className="font-medium text-slate-700">
                            {u.namaLengkap || email.split("@")[0]}
                          </span>
                          {u.namaSPPG ? (
                            <>
                              {" · "}
                              <span>{u.namaSPPG}</span>
                            </>
                          ) : null}
                          {" · "}
                          <span>Login Pertama: {formatDateTimeId(u.loginPertama || u.createdAt)}</span>
                          {" · "}
                          <span>
                            Login Terakhir: {formatDateTimeId(u.loginTerakhir)} ({loginCount}x)
                          </span>
                        </p>
                      </div>

                      {/* Row Action Controls */}
                      <div className="flex flex-wrap items-center gap-2 shrink-0">
                        {(cat === "menunggu" || cat === "kedaluwarsa" || cat === "diblokir") && (
                          <button
                            type="button"
                            disabled={isSavingThis}
                            onClick={() => handleAccUser(u)}
                            className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl whitespace-nowrap transition cursor-pointer disabled:opacity-60"
                          >
                            {isSavingThis ? (
                              <Loader2 className="w-3.5 h-3.5 animate-spin" />
                            ) : (
                              <UserCheck className="w-3.5 h-3.5" />
                            )}
                            <span>ACC Pengguna</span>
                          </button>
                        )}

                        {dirty && (
                          <button
                            type="button"
                            disabled={isSavingThis}
                            onClick={() => handleSaveDraft(u)}
                            className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs rounded-xl whitespace-nowrap transition cursor-pointer disabled:opacity-60"
                          >
                            {isSavingThis ? (
                              <Loader2 className="w-3.5 h-3.5 animate-spin" />
                            ) : (
                              <Save className="w-3.5 h-3.5" />
                            )}
                            <span>Simpan</span>
                          </button>
                        )}

                        {!isProtectedMainAdmin && cat !== "menunggu" && (
                          <button
                            type="button"
                            disabled={isSavingThis}
                            onClick={() => handleReturnToWaiting(u)}
                            className="inline-flex items-center gap-1.5 px-3 py-2 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 font-bold text-xs rounded-xl whitespace-nowrap transition cursor-pointer disabled:opacity-60"
                          >
                            <RotateCcw className="w-3.5 h-3.5" />
                            <span>Daftar Tunggu</span>
                          </button>
                        )}

                        {!isProtectedMainAdmin && cat !== "diblokir" && (
                          <button
                            type="button"
                            disabled={isSavingThis}
                            onClick={() => handleBlockUser(u)}
                            className="inline-flex items-center gap-1.5 px-3 py-2 bg-white hover:bg-rose-50 text-rose-600 border border-rose-200 font-bold text-xs rounded-xl whitespace-nowrap transition cursor-pointer disabled:opacity-60"
                          >
                            <Ban className="w-3.5 h-3.5" />
                            <span>Blokir</span>
                          </button>
                        )}

                        {!isProtectedMainAdmin && (
                          <button
                            type="button"
                            disabled={isSavingThis}
                            onClick={() => setDeleteTargetUser(u)}
                            title="Hapus catatan pengguna"
                            aria-label={`Hapus ${email}`}
                            className="inline-flex items-center justify-center p-2 bg-white hover:bg-rose-50 text-slate-500 hover:text-rose-600 border border-slate-200 hover:border-rose-200 rounded-xl transition cursor-pointer disabled:opacity-60"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        )}
                      </div>
                    </div>

                    {/* 3-Column Inline Access Controls — Exact Screenshot Grid */}
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 pt-3 border-t border-slate-100">
                      {/* Col 1: Tanggal Mulai Login */}
                      <div>
                        <div className="flex items-center justify-between mb-1.5">
                          <label
                            htmlFor={`start-date-${email}`}
                            className="text-xs font-semibold text-slate-700"
                          >
                            Tanggal Mulai Login
                          </label>
                          {draft.mulaiPada ? (
                            <button
                              type="button"
                              onClick={() => updateUserDraft(u, { mulaiPada: "" })}
                              className="text-[11px] font-semibold text-slate-400 hover:text-slate-700 cursor-pointer"
                            >
                              Hapus
                            </button>
                          ) : (
                            <button
                              type="button"
                              onClick={() => updateUserDraft(u, { mulaiPada: todayIso })}
                              className="text-[11px] font-semibold text-slate-400 hover:text-emerald-700 cursor-pointer"
                            >
                              Hari Ini
                            </button>
                          )}
                        </div>
                        <input
                          id={`start-date-${email}`}
                          type="date"
                          value={draft.mulaiPada}
                          onChange={(e) => updateUserDraft(u, { mulaiPada: e.target.value })}
                          className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs sm:text-sm text-slate-800 tabular-nums focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
                        />
                      </div>

                      {/* Col 2: Tanggal Berakhir Login (Atur Admin) + Tanpa Batas + Quick Presets */}
                      <div>
                        <div className="flex items-center justify-between mb-1.5">
                          <label
                            htmlFor={`end-date-${email}`}
                            className="text-xs font-semibold text-slate-700"
                          >
                            Tanggal Berakhir Login (Atur Admin)
                          </label>
                          <button
                            type="button"
                            onClick={() => handleApplyExpirationPreset(u, "unlimited")}
                            className={`text-xs font-semibold cursor-pointer transition ${
                              draft.berakhirPada === ""
                                ? "text-emerald-700"
                                : "text-slate-400 hover:text-emerald-700"
                            }`}
                          >
                            Tanpa Batas
                          </button>
                        </div>
                        <input
                          id={`end-date-${email}`}
                          type="date"
                          value={draft.berakhirPada}
                          onChange={(e) => updateUserDraft(u, { berakhirPada: e.target.value })}
                          className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs sm:text-sm text-slate-800 tabular-nums focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
                        />
                        <div className="flex flex-wrap items-center gap-1.5 mt-2">
                          {(
                            [
                              { id: "7d", label: "+7 Hr" },
                              { id: "30d", label: "+30 Hr" },
                              { id: "3m", label: "+3 Bln" },
                              { id: "1y", label: "+1 Thn" },
                            ] as const
                          ).map((preset) => (
                            <button
                              key={preset.id}
                              type="button"
                              disabled={isSavingThis}
                              onClick={() => handleApplyExpirationPreset(u, preset.id)}
                              className="px-2.5 py-1 bg-slate-50 hover:bg-emerald-50 text-slate-600 hover:text-emerald-800 border border-slate-200 hover:border-emerald-300 rounded-lg text-[11px] font-semibold tabular-nums transition cursor-pointer disabled:opacity-50"
                            >
                              {preset.label}
                            </button>
                          ))}
                        </div>
                      </div>

                      {/* Col 3: Keterangan (Ditampilkan ke Pengguna) */}
                      <div>
                        <div className="flex items-center justify-between mb-1.5">
                          <label
                            htmlFor={`note-${email}`}
                            className="text-xs font-semibold text-slate-700"
                          >
                            Keterangan (Ditampilkan ke Pengguna)
                          </label>
                          {justSaved ? (
                            <span className="text-[11px] font-semibold text-emerald-700">
                              Tersimpan
                            </span>
                          ) : dirty ? (
                            <span className="text-[11px] font-semibold text-amber-700">
                              Belum disimpan
                            </span>
                          ) : null}
                        </div>
                        <input
                          id={`note-${email}`}
                          type="text"
                          value={draft.keterangan}
                          onChange={(e) => updateUserDraft(u, { keterangan: e.target.value })}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") {
                              e.preventDefault();
                              handleSaveDraft(u);
                            }
                          }}
                          placeholder="Daftar Tunggu - Menunggu ACC Admin Odhe Gizi"
                          className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs sm:text-sm text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
                        />
                      </div>
                    </div>
                  </article>
                );
              })}
            </div>
          )}
        </section>

        {/* DATABASE SPREADSHEET ADMIN (CSV + Google Sheets Sync + Apps Script Webhook) */}
        <SpreadsheetUserSyncCard
          users={users}
          onUsersImportedFromSheet={handleUsersImportedFromSheet}
        />

        {/* COMPACT COLLAPSIBLE REAL-TIME DEVICE REGISTRY */}
        <section className="bg-white rounded-2xl border border-slate-200 overflow-hidden">
          <button
            type="button"
            onClick={() => setShowDevicesDrawer((prev) => !prev)}
            className="w-full px-5 py-4 flex items-center justify-between gap-3 hover:bg-slate-50/70 transition cursor-pointer text-left"
          >
            <div className="flex flex-wrap items-center gap-x-2 gap-y-1">
              <span className="text-sm font-bold text-slate-900">Perangkat Terdeteksi Otomatis</span>
              <span className="text-slate-300">·</span>
              <span className="text-xs font-semibold text-emerald-700 tabular-nums">
                {onlineDevicesCount} online
              </span>
              <span className="text-slate-300">·</span>
              <span className="text-xs text-slate-500 tabular-nums">
                {devices.length} total perangkat (Seluler &amp; Laptop)
              </span>
            </div>
            {showDevicesDrawer ? (
              <ChevronUp className="w-4 h-4 text-slate-500 shrink-0" />
            ) : (
              <ChevronDown className="w-4 h-4 text-slate-500 shrink-0" />
            )}
          </button>

          {showDevicesDrawer && (
            <div className="border-t border-slate-100 divide-y divide-slate-100">
              {devices.length === 0 ? (
                <div className="px-5 py-6 text-xs text-slate-500 text-center">
                  Belum ada perangkat yang tercatat pada sesi saat ini.
                </div>
              ) : (
                devices.slice(0, 30).map((d) => {
                  const online = isDeviceCurrentlyOnline(d);
                  const DevIcon =
                    d.deviceType === "Mobile"
                      ? Smartphone
                      : d.deviceType === "Tablet"
                      ? Tablet
                      : Monitor;
                  return (
                    <div
                      key={d.deviceId}
                      className="px-5 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs"
                    >
                      <div className="flex items-start gap-3 min-w-0">
                        <DevIcon className="w-4 h-4 text-slate-500 shrink-0 mt-0.5" />
                        <div className="min-w-0">
                          <div className="flex flex-wrap items-center gap-x-1.5">
                            <span className="font-bold text-slate-900">
                              {d.userEmail || "Pengunjung Login"}
                            </span>
                            <span className="text-slate-300">·</span>
                            <span
                              className={`font-semibold ${
                                online ? "text-emerald-700" : "text-slate-500"
                              }`}
                            >
                              {online ? "Online" : "Offline"}
                            </span>
                          </div>
                          <div className="text-slate-500 mt-0.5 tabular-nums">
                            {d.deviceModel || d.deviceType} · {d.osName} · {d.browserName} · IP:{" "}
                            {d.ipAddress || "-"} · Terakhir: {formatDateTimeId(d.lastSeenAt)}
                          </div>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleDeleteDevice(d.deviceId)}
                        title="Hapus log perangkat"
                        className="self-end sm:self-center p-1.5 text-slate-400 hover:text-rose-600 rounded-lg cursor-pointer"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  );
                })
              )}
            </div>
          )}
        </section>
      </div>

      {/* CUSTOM DELETE CONFIRMATION MODAL (No window.confirm) */}
      {deleteTargetUser && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-xs z-[9999] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full shadow-2xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <h4 className="text-sm font-extrabold text-slate-900">Hapus Catatan Pengguna</h4>
              <button
                type="button"
                disabled={isDeleting}
                onClick={() => setDeleteTargetUser(null)}
                className="p-1 text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-2 text-xs text-slate-600 leading-relaxed">
              <p>
                Apakah Anda yakin ingin menghapus catatan akses untuk{" "}
                <strong className="text-slate-900">{deleteTargetUser.email}</strong>?
              </p>
              <p>
                Pengguna akan otomatis dikeluarkan dari aplikasi dan perlu masuk ulang ke Daftar Tunggu jika ingin meminta ACC kembali.
              </p>
            </div>

            <div className="px-5 py-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                disabled={isDeleting}
                onClick={() => setDeleteTargetUser(null)}
                className="px-4 py-2 bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 font-bold text-xs rounded-xl cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                disabled={isDeleting}
                onClick={handleConfirmDeleteUser}
                className="inline-flex items-center gap-1.5 px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs rounded-xl cursor-pointer disabled:opacity-60"
              >
                {isDeleting ? (
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <Trash2 className="w-3.5 h-3.5" />
                )}
                <span>{isDeleting ? "Menghapus..." : "Ya, Hapus Catatan"}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* GLOBAL LOGO UPLOAD MODAL */}
      {isLogoModalOpen && onUploadGlobalLogo && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-xs z-[9999] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full shadow-2xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <h4 className="text-sm font-extrabold text-slate-900">Logo Global Aplikasi &amp; Halaman Login</h4>
              <button
                type="button"
                onClick={() => setIsLogoModalOpen(false)}
                className="p-1 text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-4 text-center">
              <div className="w-24 h-24 rounded-full overflow-hidden border-2 border-emerald-500 mx-auto bg-slate-50">
                <img
                  src={globalLogoUrl || "/assets/images/app_logo_sppg.png"}
                  alt="Logo Global SISPERMBG"
                  className="w-full h-full object-cover"
                  referrerPolicy="no-referrer"
                  onError={(e) => {
                    (e.target as HTMLImageElement).src = "/assets/images/app_logo_sppg.png";
                  }}
                />
              </div>
              <p className="text-xs text-slate-600 leading-relaxed">
                Khusus Admin: Logo ini hanya mengubah logo lingkaran di header atas aplikasi dan halaman depan Login (terpisah dari pengaturan Logo KOP Surat di Profil SPPG).
              </p>
              <div className="flex flex-wrap items-center justify-center gap-2">
                <label className="inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl cursor-pointer transition">
                  <Camera className="w-4 h-4" />
                  <span>Pilih Gambar Logo Baru</span>
                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e) => {
                      onUploadGlobalLogo(e);
                      setIsLogoModalOpen(false);
                    }}
                    className="hidden"
                  />
                </label>
                {onResetGlobalLogo && globalLogoUrl && globalLogoUrl !== "/assets/images/app_logo_sppg.png" && (
                  <button
                    type="button"
                    onClick={() => {
                      onResetGlobalLogo();
                      setIsLogoModalOpen(false);
                      showToast("success", "Logo header aplikasi dikembalikan ke standar default.");
                    }}
                    className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl cursor-pointer transition"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>Reset Default</span>
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
