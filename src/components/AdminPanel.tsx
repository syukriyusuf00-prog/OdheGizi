import React, { useState, useEffect, useRef } from "react";
import { 
  collection, 
  getDocs, 
  doc, 
  setDoc,
  updateDoc, 
  deleteDoc, 
  onSnapshot,
  serverTimestamp 
} from "firebase/firestore";
import {
  db,
  FIREBASE_PROJECT_ID,
  FIRESTORE_DATABASE_ID,
  FIRESTORE_UPGRADE_URL,
  FirestoreHealthState,
  getFirestoreHealth,
  subscribeFirestoreHealth,
  testFirestoreConnection,
  reportFirestoreWriteResult,
} from "../lib/firebase.ts";
import {
  isMainAdminEmail,
  syncUserToRealtimeHub,
  deleteUserFromRealtimeHub,
} from "../context/AuthContext.tsx";
import { 
  Users, 
  CheckCircle, 
  Clock, 
  Search, 
  Trash2, 
  Calendar, 
  Shield, 
  ShieldAlert, 
  RefreshCw, 
  UserCheck, 
  XCircle, 
  Database, 
  Cpu, 
  History,
  Mail,
  User,
  Briefcase,
  Building,
  Key,
  Camera,
  Smartphone,
  Tablet,
  Laptop,
  Monitor,
  Globe,
  Wifi,
  WifiOff,
  MailCheck,
  AlertTriangle,
  ExternalLink
} from "lucide-react";
import { DeviceRecord, getCachedLocalDevice, formatDeviceShortLabel } from "../lib/deviceTracker.ts";

interface AdminPanelProps {
  onClose: () => void;
  globalLogoUrl?: string;
  onUploadGlobalLogo?: (e: React.ChangeEvent<HTMLInputElement>) => void;
}

const fetchWithTimeout = <T,>(promise: Promise<T>, timeoutMs: number, defaultValue: T): Promise<T> => {
  return Promise.race([
    promise,
    new Promise<T>((resolve) => setTimeout(() => resolve(defaultValue), timeoutMs))
  ]);
};

const DEFAULT_DELETED_EMAILS = [
  "sofstory96@gmail.com",
  "desylarasati33@gmail.com",
  "saharhesa781@gmail.com",
  "raniramadani696@gmail.com",
  "erickhernando1@gmail.com",
  "imranjho373@gmail.com",
  "widyaainunpratiwi@gmail.com"
];

// Helper to manage deleted tenant emails persistence
export const getDeletedTenantEmails = (): Set<string> => {
  try {
    const raw = localStorage.getItem("deleted_tenant_emails");
    if (raw !== null) {
      const arr = JSON.parse(raw);
      return new Set(arr.map((e: string) => e.toLowerCase().trim()));
    }
  } catch (e) {}
  return new Set(DEFAULT_DELETED_EMAILS.map(e => e.toLowerCase().trim()));
};

export const addDeletedTenantEmail = (email: string) => {
  if (!email) return;
  const set = getDeletedTenantEmails();
  set.add(email.toLowerCase().trim());
  localStorage.setItem("deleted_tenant_emails", JSON.stringify(Array.from(set)));
};

export const removeDeletedTenantEmail = (email: string) => {
  if (!email) return;
  const set = getDeletedTenantEmails();
  set.delete(email.toLowerCase().trim());
  localStorage.setItem("deleted_tenant_emails", JSON.stringify(Array.from(set)));
};

// Scan localStorage for any offline registered user profiles
const getOfflineUsers = () => {
  const list: any[] = [];
  const deletedEmails = getDeletedTenantEmails();
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith("offline_user_")) {
        const val = localStorage.getItem(key);
        if (val) {
          try {
            const parsed = JSON.parse(val);
            const userEmail = (parsed.email || "").toLowerCase().trim();
            if (!deletedEmails.has(userEmail)) {
              list.push(parsed);
            }
          } catch (e) {}
        }
      }
    }
  } catch (e) {
    console.warn("Notice reading offline users:", e);
  }
  return list;
};

export default function AdminPanel({ onClose, globalLogoUrl, onUploadGlobalLogo }: AdminPanelProps) {
  const [users, setUsers] = useState<any[]>([]);
  const [devices, setDevices] = useState<DeviceRecord[]>(() => {
    const localDev = getCachedLocalDevice();
    return localDev ? [localDev] : [];
  });
  const [deviceFilter, setDeviceFilter] = useState<"all" | "online" | "Desktop / Laptop" | "Mobile" | "Tablet" | "guest" | "demo">("all");
  const [deviceSearchQuery, setDeviceSearchQuery] = useState("");
  const [loading, setLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState("");
  const [activeTab, setActiveTab] = useState<"manajemen" | "perangkat" | "riwayat" | "diagnostik">(() => {
    return (localStorage.getItem("sisper_admin_tab") as any) || "manajemen";
  });
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [selectedUserUids, setSelectedUserUids] = useState<string[]>([]);
  const [nowTick, setNowTick] = useState(() => Date.now());
  const [firestoreHealth, setFirestoreHealth] = useState<FirestoreHealthState>(() => getFirestoreHealth());
  const [sseStatus, setSseStatus] = useState<"connecting" | "connected" | "reconnecting">("connecting");
  const [lastLiveEventAt, setLastLiveEventAt] = useState<string | null>(null);
  const [testingDb, setTestingDb] = useState(false);

  // Keep dual-channel real-time stores (Firestore + Express SSE Hub) merged cleanly without clobbering
  const firestoreUsersRef = useRef<any[]>([]);
  const hubUsersRef = useRef<Map<string, any>>(new Map());
  const firestoreDevicesRef = useRef<DeviceRecord[]>([]);
  const hubDevicesRef = useRef<Map<string, DeviceRecord>>(new Map());

  useEffect(() => {
    const unsub = subscribeFirestoreHealth(setFirestoreHealth);
    testFirestoreConnection().catch(() => {});
    return unsub;
  }, []);

  useEffect(() => {
    const timer = window.setInterval(() => setNowTick(Date.now()), 30_000);
    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => {
    localStorage.setItem("sisper_admin_tab", activeTab);
  }, [activeTab]);

  // Load all users & devices combining Firestore real-time onSnapshot + Server SSE Real-Time Hub + local cache
  useEffect(() => {
    setLoading(true);

    const mainAdminDoc = {
      id: "admin_syukriyusuf82",
      uid: "admin_syukriyusuf82",
      email: "syukriyusuf82@gmail.com",
      namaLengkap: "LA ODE MUHAMMAD SUKRI YUSUF",
      profesi: "AHLI GIZI",
      namaSPPG: "SPPG MUNA BARAT SAWERIGADI ONDOKE",
      noHp: "082271095251",
      peran: "ADMIN",
      statusPersetujuan: "aktif",
      berakhirPada: "2035-12-31"
    };

    const getRecordTimeMs = (val: any): number => {
      if (!val) return 0;
      if (typeof val === "string") {
        const t = new Date(val).getTime();
        return Number.isNaN(t) ? 0 : t;
      }
      if (typeof val?.seconds === "number") return val.seconds * 1000;
      return 0;
    };

    const rebuildMergedUsers = () => {
      const map = new Map<string, any>();
      map.set("admin_syukriyusuf82", mainAdminDoc);
      const deletedEmails = getDeletedTenantEmails();

      const addToMap = (u: any, isRemote = false) => {
        if (!u) return;
        const userEmail = (u.email || "").toLowerCase().trim();
        if (
          u.uid === "admin_sukriyusuf82" ||
          u.uid === "admin_syukriyusuf82" ||
          (userEmail && isMainAdminEmail(userEmail))
        ) {
          return;
        }
        // Only apply local deleted-email filter to stale localStorage caches — NEVER drop live remote registrations from mobile/cloud!
        if (!isRemote && userEmail && deletedEmails.has(userEmail)) {
          return;
        }

        const primaryKey = u.uid || u.id || (userEmail ? `custom_user_${userEmail.replace(/[@.]/g, "_")}` : null);
        if (!primaryKey && !userEmail) return;

        let existingKey: string | null = null;
        if (primaryKey && map.has(primaryKey)) {
          existingKey = primaryKey;
        } else if (userEmail) {
          for (const [k, v] of map.entries()) {
            if ((v.email || "").toLowerCase().trim() === userEmail) {
              existingKey = k;
              break;
            }
          }
        }

        const existing = existingKey ? map.get(existingKey) : null;
        // Prefer real Firebase Auth UID over legacy custom_user_* key when merging same email
        const incomingIsRealUid = Boolean(primaryKey && !primaryKey.startsWith("custom_user_"));
        const existingIsLegacyUid = Boolean(existingKey && existingKey.startsWith("custom_user_"));
        const keyToUse =
          existingKey && existingIsLegacyUid && incomingIsRealUid && primaryKey
            ? primaryKey
            : existingKey || primaryKey || `custom_user_${userEmail.replace(/[@.]/g, "_")}`;

        if (!existing) {
          map.set(keyToUse, {
            ...u,
            uid: keyToUse,
            id: keyToUse,
            statusPersetujuan: u.statusPersetujuan || "menunggu",
          });
          return;
        }

        if (existingKey && existingKey !== keyToUse) {
          map.delete(existingKey);
        }

        const existingUpdated = Math.max(
          getRecordTimeMs(existing.updatedAt),
          getRecordTimeMs(existing.loginTerakhir),
          getRecordTimeMs(existing.lastDeviceSeenAt)
        );
        const incomingUpdated = Math.max(
          getRecordTimeMs(u.updatedAt),
          getRecordTimeMs(u.loginTerakhir),
          getRecordTimeMs(u.lastDeviceSeenAt)
        );

        const newer = incomingUpdated >= existingUpdated ? u : existing;
        const older = incomingUpdated >= existingUpdated ? existing : u;

        const finalStatus = newer.statusPersetujuan || older.statusPersetujuan || "menunggu";

        const merged = {
          ...older,
          ...newer,
          uid: keyToUse,
          id: keyToUse,
          statusPersetujuan: finalStatus,
          loginTerakhir:
            getRecordTimeMs(u.loginTerakhir) >= getRecordTimeMs(existing.loginTerakhir)
              ? u.loginTerakhir || existing.loginTerakhir
              : existing.loginTerakhir || u.loginTerakhir,
        };

        map.set(keyToUse, merged);
      };

      // 1. Local offline cache first
      const offlineList = getOfflineUsers();
      offlineList.forEach((u) => addToMap(u, false));

      // 2. Remote Firestore documents
      firestoreUsersRef.current.forEach((u) => addToMap(u, true));

      // 3. Live Server Real-Time Hub users (instant Mobile -> Laptop sync)
      hubUsersRef.current.forEach((u) => addToMap(u, true));

      const userList = Array.from(map.values());
      userList.sort((a, b) => {
        const timeA = Math.max(getRecordTimeMs(a.loginTerakhir), getRecordTimeMs(a.updatedAt), getRecordTimeMs(a.createdAt));
        const timeB = Math.max(getRecordTimeMs(b.loginTerakhir), getRecordTimeMs(b.updatedAt), getRecordTimeMs(b.createdAt));
        return timeB - timeA;
      });

      setUsers(userList);
      setLoading(false);
    };

    const rebuildMergedDevices = () => {
      const map = new Map<string, DeviceRecord>();
      const upsert = (d: DeviceRecord) => {
        if (!d?.deviceId) return;
        const prev = map.get(d.deviceId);
        if (!prev) {
          map.set(d.deviceId, d);
          return;
        }
        const prevTime = new Date(prev.lastSeenAt || 0).getTime();
        const nextTime = new Date(d.lastSeenAt || 0).getTime();
        map.set(d.deviceId, nextTime >= prevTime ? { ...prev, ...d } : { ...d, ...prev });
      };

      const localDev = getCachedLocalDevice();
      if (localDev?.deviceId) upsert(localDev);
      firestoreDevicesRef.current.forEach(upsert);
      hubDevicesRef.current.forEach(upsert);

      const sorted = Array.from(map.values()).sort(
        (a, b) => new Date(b.lastSeenAt || 0).getTime() - new Date(a.lastSeenAt || 0).getTime()
      );
      setDevices(sorted);
    };

    // Immediate initial render
    rebuildMergedUsers();
    rebuildMergedDevices();

    // Channel A: Firestore real-time listeners (/users & /devices)
    const usersCol = collection(db, "users");
    const unsubscribeUsers = onSnapshot(
      usersCol,
      (snapshot) => {
        const remoteList: any[] = [];
        snapshot.docs.forEach((docSnap) => {
          const data = docSnap.data();
          if (docSnap.id === "admin_sukriyusuf82" || docSnap.id === "admin_syukriyusuf82" || (data.email && isMainAdminEmail(data.email))) {
            return;
          }
          const userUid = data.uid || docSnap.id;
          remoteList.push({
            ...data,
            id: docSnap.id,
            uid: userUid,
          });
        });
        firestoreUsersRef.current = remoteList;
        setLastLiveEventAt(new Date().toISOString());
        rebuildMergedUsers();
      },
      (err) => {
        console.warn("Firestore user listener notice:", err);
        setLoading(false);
      }
    );

    const devicesCol = collection(db, "devices");
    const unsubscribeDevices = onSnapshot(
      devicesCol,
      (snapshot) => {
        firestoreDevicesRef.current = snapshot.docs.map((docSnap) => {
          const data = docSnap.data() as DeviceRecord;
          return {
            ...data,
            deviceId: data.deviceId || docSnap.id,
          };
        });
        setLastLiveEventAt(new Date().toISOString());
        rebuildMergedDevices();
      },
      (err) => {
        console.warn("Firestore devices listener notice:", err);
      }
    );

    // Channel B: Real-Time Server SSE Hub (/api/realtime/stream + /api/realtime/snapshot)
    // Guarantees instant (<1s) Mobile -> Laptop synchronization across networks
    const hydrateFromServerSnapshot = async () => {
      try {
        const resp = await fetch("/api/realtime/snapshot", { cache: "no-store" });
        if (!resp.ok) return;
        const json = await resp.json();
        if (Array.isArray(json?.users)) {
          hubUsersRef.current = new Map(json.users.filter((u: any) => u?.uid).map((u: any) => [u.uid, u]));
          rebuildMergedUsers();
        }
        if (Array.isArray(json?.devices)) {
          hubDevicesRef.current = new Map(json.devices.filter((d: any) => d?.deviceId).map((d: any) => [d.deviceId, d]));
          rebuildMergedDevices();
        }
      } catch {}
    };

    hydrateFromServerSnapshot();

    let eventSource: EventSource | null = null;
    try {
      setSseStatus("connecting");
      eventSource = new EventSource("/api/realtime/stream");

      eventSource.onopen = () => {
        setSseStatus("connected");
      };

      eventSource.addEventListener("snapshot", (ev: MessageEvent) => {
        try {
          const parsed = JSON.parse(ev.data);
          setSseStatus("connected");
          setLastLiveEventAt(new Date().toISOString());
          if (Array.isArray(parsed?.users)) {
            hubUsersRef.current = new Map(parsed.users.filter((u: any) => u?.uid).map((u: any) => [u.uid, u]));
            rebuildMergedUsers();
          }
          if (Array.isArray(parsed?.devices)) {
            hubDevicesRef.current = new Map(parsed.devices.filter((d: any) => d?.deviceId).map((d: any) => [d.deviceId, d]));
            rebuildMergedDevices();
          }
        } catch {}
      });

      eventSource.addEventListener("user_upsert", (ev: MessageEvent) => {
        try {
          const incoming = JSON.parse(ev.data);
          if (incoming?.uid) {
            hubUsersRef.current.set(incoming.uid, incoming);
            setLastLiveEventAt(new Date().toISOString());
            rebuildMergedUsers();
          }
        } catch {}
      });

      eventSource.addEventListener("user_delete", (ev: MessageEvent) => {
        try {
          const payload = JSON.parse(ev.data);
          const targetEmail = (payload?.email || "").toLowerCase().trim();
          const targetUid = payload?.uid || "";
          for (const [k, u] of hubUsersRef.current.entries()) {
            if ((targetUid && (k === targetUid || u.uid === targetUid)) || (targetEmail && (u.email || "").toLowerCase().trim() === targetEmail)) {
              hubUsersRef.current.delete(k);
            }
          }
          setLastLiveEventAt(new Date().toISOString());
          rebuildMergedUsers();
        } catch {}
      });

      eventSource.addEventListener("device_upsert", (ev: MessageEvent) => {
        try {
          const incoming = JSON.parse(ev.data) as DeviceRecord;
          if (incoming?.deviceId) {
            hubDevicesRef.current.set(incoming.deviceId, incoming);
            setLastLiveEventAt(new Date().toISOString());
            rebuildMergedDevices();
          }
        } catch {}
      });

      eventSource.addEventListener("device_delete", (ev: MessageEvent) => {
        try {
          const payload = JSON.parse(ev.data);
          if (payload?.deviceId) {
            hubDevicesRef.current.delete(payload.deviceId);
            setLastLiveEventAt(new Date().toISOString());
            rebuildMergedDevices();
          }
        } catch {}
      });

      eventSource.onerror = () => {
        setSseStatus("reconnecting");
      };
    } catch {
      setSseStatus("reconnecting");
    }

    // Periodic lightweight snapshot safety poll (every 12s) in case a corporate proxy buffers SSE
    const safetyPoll = window.setInterval(() => {
      if (document.visibilityState === "visible") {
        hydrateFromServerSnapshot();
      }
    }, 12_000);

    // Same-tab instant registration / device listeners
    const handleRegistered = (e: any) => {
      if (e?.detail) {
        const newUser = e.detail;
        const newUserEmail = (newUser.email || "").toLowerCase().trim();
        const emailSlug = newUserEmail.replace(/[@.]/g, "_");
        const customUid = newUser.uid || `custom_user_${emailSlug}`;
        localStorage.setItem(`offline_user_${customUid}`, JSON.stringify(newUser));
        hubUsersRef.current.set(customUid, { ...newUser, uid: customUid, id: customUid });
      }
      rebuildMergedUsers();
    };

    const handleLocalDeviceUpdate = (e: any) => {
      if (e?.detail?.deviceId) {
        hubDevicesRef.current.set(e.detail.deviceId, e.detail as DeviceRecord);
        rebuildMergedDevices();
      }
    };

    window.addEventListener("sisper_user_registered", handleRegistered);
    window.addEventListener("sisper_device_updated", handleLocalDeviceUpdate);

    return () => {
      unsubscribeUsers();
      unsubscribeDevices();
      if (eventSource) eventSource.close();
      window.clearInterval(safetyPoll);
      window.removeEventListener("sisper_user_registered", handleRegistered);
      window.removeEventListener("sisper_device_updated", handleLocalDeviceUpdate);
    };
  }, [refreshTrigger]);

  const isDeviceLiveOnline = (dev?: Partial<DeviceRecord> | null) => {
    if (!dev || dev.isOnline === false || !dev.lastSeenAt) return false;
    const seenMs = new Date(dev.lastSeenAt).getTime();
    if (Number.isNaN(seenMs)) return false;
    return nowTick - seenMs < 3 * 60 * 1000;
  };

  const formatRelativeTime = (iso?: string | null) => {
    if (!iso) return "-";
    const ms = new Date(iso).getTime();
    if (Number.isNaN(ms)) return "-";
    const diffSec = Math.max(0, Math.floor((nowTick - ms) / 1000));
    if (diffSec < 45) return "Baru saja";
    if (diffSec < 3600) return `${Math.floor(diffSec / 60)} mnt lalu`;
    if (diffSec < 86400) return `${Math.floor(diffSec / 3600)} jam lalu`;
    return `${Math.floor(diffSec / 86400)} hari lalu`;
  };

  const handleDeleteDeviceRecord = async (deviceId: string) => {
    if (!deviceId) return;
    try {
      hubDevicesRef.current.delete(deviceId);
      setDevices((prev) => prev.filter((d) => d.deviceId !== deviceId));
      await Promise.all([
        fetch("/api/realtime/delete-device", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ deviceId }),
        }).catch(() => {}),
        fetchWithTimeout(deleteDoc(doc(db, "devices", deviceId)), 2500, null).catch((err) => {
          reportFirestoreWriteResult(false, err);
        }),
      ]);
    } catch (err) {
      console.warn("Gagal menghapus catatan perangkat:", err);
    }
  };

  const getDevicesForUser = (userItem: any): DeviceRecord[] => {
    const emailLower = (userItem?.email || "").toLowerCase().trim();
    const matched = devices.filter(
      (d) =>
        (d.userUid && (d.userUid === userItem.uid || d.userUid === userItem.id)) ||
        (emailLower && d.userEmail && d.userEmail.toLowerCase().trim() === emailLower)
    );
    if (matched.length > 0) return matched;
    if (userItem?.lastDevice?.deviceId) {
      return [
        {
          deviceId: userItem.lastDevice.deviceId,
          deviceType: userItem.lastDevice.deviceType || "Desktop / Laptop",
          deviceModel: userItem.lastDevice.deviceModel || "Perangkat Web",
          osName: userItem.lastDevice.osName || "-",
          browserName: userItem.lastDevice.browserName || "-",
          screenResolution: userItem.lastDevice.screenResolution || "-",
          viewportSize: "-",
          language: "id-ID",
          timezone: userItem.lastDevice.timezone || "-",
          ipAddress: userItem.lastDevice.ipAddress || "-",
          userAgent: "",
          sessionState: "active",
          userUid: userItem.uid,
          userEmail: userItem.email,
          userName: userItem.namaLengkap,
          namaSPPG: userItem.namaSPPG,
          peran: userItem.peran,
          statusPersetujuan: userItem.statusPersetujuan,
          isOnline: Boolean(userItem.lastDevice.isOnline),
          firstSeenAt: userItem.lastDevice.lastSeenAt || userItem.createdAt || "",
          lastSeenAt: userItem.lastDevice.lastSeenAt || userItem.lastDeviceSeenAt || "",
        },
      ];
    }
    return [];
  };

  const renderDeviceIcon = (deviceType?: string, className = "w-3.5 h-3.5") => {
    if (deviceType === "Mobile") return <Smartphone className={className} />;
    if (deviceType === "Tablet") return <Tablet className={className} />;
    return <Laptop className={className} />;
  };

  // Handle setting explicit approval status (ACC / Menunggu / Blokir)
  const handleSetStatus = async (userId: string, newStatus: "aktif" | "menunggu" | "diblokir") => {
    try {
      const now = new Date();
      const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      const userObj = users.find(u => u.uid === userId || u.id === userId || (u.email && u.email.toLowerCase().trim() === userId.toLowerCase().trim()));

      const userEmail = (userObj?.email || "").toLowerCase().trim();
      const emailSlug = userEmail.replace(/[@.]/g, '_');
      const legacyCustomUid = emailSlug ? `custom_user_${emailSlug}` : "";
      const targetUid = userObj?.uid || userObj?.id || userId || legacyCustomUid;

      const updates: any = {
        statusPersetujuan: newStatus,
        updatedAt: new Date().toISOString()
      };

      if (newStatus === "aktif") {
        updates.activatedBy = "admin_acc";
        if (!userObj?.berakhirPada || userObj.berakhirPada < todayStr) {
          const nextYear = new Date();
          nextYear.setFullYear(now.getFullYear() + 1);
          const defaultExp = `${nextYear.getFullYear()}-${String(nextYear.getMonth() + 1).padStart(2, '0')}-${String(nextYear.getDate()).padStart(2, '0')}`;
          updates.berakhirPada = defaultExp;
        }
      }

      const fullUserData = {
        ...(userObj || {}),
        ...updates,
        uid: targetUid,
        id: targetUid,
        email: userEmail || userObj?.email || ""
      };

      // 1. Instant local state & live hub cache update (0ms UI latency)
      hubUsersRef.current.set(targetUid, fullUserData);
      setUsers(prev => prev.map(u => 
        (u.uid === userId || u.id === userId || u.uid === targetUid || (userEmail && (u.email || "").toLowerCase().trim() === userEmail)) 
          ? { ...u, ...fullUserData } 
          : u
      ));

      // 2. Instant offline storage persistence
      const keysToUpdate = Array.from(new Set([
        `offline_user_${targetUid}`,
        `offline_user_${userId}`,
        legacyCustomUid && targetUid === legacyCustomUid ? `offline_user_${legacyCustomUid}` : ''
      ].filter(Boolean)));

      keysToUpdate.forEach((k) => {
        const savedStr = localStorage.getItem(k);
        const parsed = savedStr ? (JSON.parse(savedStr) || {}) : (userObj || {});
        localStorage.setItem(k, JSON.stringify({ ...parsed, ...fullUserData }));
      });

      if (userEmail) {
        try {
          for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key && key.startsWith("offline_user_")) {
              const val = localStorage.getItem(key);
              if (val) {
                const parsed = JSON.parse(val);
                if ((parsed.email || "").toLowerCase().trim() === userEmail) {
                  localStorage.setItem(key, JSON.stringify({ ...parsed, ...fullUserData }));
                }
              }
            }
          }
        } catch (e) {}
      }

      // 3. Real-time hub broadcast (instant tenant unlock on mobile/laptop) + Firestore persistence
      syncUserToRealtimeHub(fullUserData).catch(() => {});
      const docIdsToUpdate = Array.from(new Set([targetUid, userId, userObj?.uid, userObj?.id].filter(Boolean)));
      for (const docId of docIdsToUpdate) {
        try {
          const userRef = doc(db, "users", docId);
          await fetchWithTimeout(
            setDoc(userRef, fullUserData, { merge: true }).then(() => reportFirestoreWriteResult(true)),
            2500,
            null
          );
        } catch (e) {
          reportFirestoreWriteResult(false, e);
          console.warn("Notice setDoc status:", e);
        }
      }
    } catch (err) {
      console.warn("Status persetujuan diperbarui secara lokal:", err);
    }
  };

  // Toggle role between USER and ADMIN (granting/revoking Panel Admin access), protecting Main Admin
  const handleToggleRole = async (userId: string, currentRole: string) => {
    const userObj = users.find(u => u.uid === userId || u.id === userId);
    if (
      userId === "admin_syukriyusuf82" ||
      userId === "admin_sukriyusuf82" ||
      isMainAdminEmail(userObj?.email)
    ) {
      return;
    }

    const nextRole = currentRole === "ADMIN" ? "USER" : "ADMIN";
    const nowIso = new Date().toISOString();
    const userEmail = (userObj?.email || "").toLowerCase().trim();
    const targetUid = userObj?.uid || userObj?.id || userId;
    const updatedUser = {
      ...(userObj || {}),
      uid: targetUid,
      id: targetUid,
      email: userEmail || userObj?.email || "",
      peran: nextRole,
      updatedAt: nowIso,
    };

    hubUsersRef.current.set(targetUid, updatedUser);
    setUsers(prev => prev.map(u => (u.uid === userId || u.id === userId || u.uid === targetUid) ? { ...u, peran: nextRole, updatedAt: nowIso } : u));

    const docIds = Array.from(new Set([targetUid, userId, userObj?.uid, userObj?.id].filter(Boolean)));

    docIds.forEach((id) => {
      const k = `offline_user_${id}`;
      const saved = localStorage.getItem(k);
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          localStorage.setItem(k, JSON.stringify({ ...parsed, peran: nextRole, updatedAt: nowIso }));
        } catch (e) {}
      }
    });

    if (userObj) {
      syncUserToRealtimeHub(updatedUser).catch(() => {});
    }
    for (const docId of docIds) {
      try {
        await fetchWithTimeout(
          setDoc(doc(db, "users", docId), {
            peran: nextRole,
            updatedAt: nowIso
          }, { merge: true }).then(() => reportFirestoreWriteResult(true)),
          2500,
          null
        );
      } catch (e) {
        reportFirestoreWriteResult(false, e);
        console.warn("Notice updating user role:", e);
      }
    }
  };

  // Handle update expiration date
  const handleUpdateExpiration = async (userId: string, dateStr: string) => {
    const userObj = users.find(u => u.uid === userId || u.id === userId);
    const nowIso = new Date().toISOString();
    const targetUid = userObj?.uid || userObj?.id || userId;
    const updatedUser = {
      ...(userObj || {}),
      uid: targetUid,
      id: targetUid,
      berakhirPada: dateStr,
      updatedAt: nowIso,
    };

    hubUsersRef.current.set(targetUid, updatedUser);
    setUsers(prev => prev.map(u => (u.uid === userId || u.id === userId || u.uid === targetUid) ? { ...u, berakhirPada: dateStr, updatedAt: nowIso } : u));

    const keysToUpdate = Array.from(new Set([
      `offline_user_${targetUid}`,
      `offline_user_${userId}`,
      userObj?.uid ? `offline_user_${userObj.uid}` : ''
    ].filter(Boolean)));

    keysToUpdate.forEach((k) => {
      const savedStr = localStorage.getItem(k);
      if (savedStr) {
        try {
          const parsed = JSON.parse(savedStr);
          localStorage.setItem(k, JSON.stringify({ ...parsed, berakhirPada: dateStr, updatedAt: nowIso }));
        } catch (e) {}
      }
    });

    if (userObj) {
      syncUserToRealtimeHub(updatedUser).catch(() => {});
    }
    const docIdsToUpdate = Array.from(new Set([targetUid, userId, userObj?.uid].filter(Boolean)));
    for (const docId of docIdsToUpdate) {
      try {
        const userRef = doc(db, "users", docId);
        await fetchWithTimeout(setDoc(userRef, {
          berakhirPada: dateStr,
          updatedAt: nowIso
        }, { merge: true }).then(() => reportFirestoreWriteResult(true)), 1500, null);
      } catch (err) {
        reportFirestoreWriteResult(false, err);
        console.warn("Tanggal kedaluwarsa diperbarui lokal:", err);
      }
    }
  };

  // Quick preset helpers for expiration date
  const handleAddDuration = async (userId: string, monthsToAdd: number) => {
    const d = new Date();
    d.setMonth(d.getMonth() + monthsToAdd);
    const dateStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    await handleSetStatus(userId, "aktif");
    await handleUpdateExpiration(userId, dateStr);
  };

  const handleRevokeAccessNow = async (userId: string) => {
    const userObj = users.find(u => u.uid === userId || u.id === userId);
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const pastStr = `${yesterday.getFullYear()}-${String(yesterday.getMonth() + 1).padStart(2, '0')}-${String(yesterday.getDate()).padStart(2, '0')}`;
    const nowIso = new Date().toISOString();
    const targetUid = userObj?.uid || userObj?.id || userId;
    const updatedUser = {
      ...(userObj || {}),
      uid: targetUid,
      id: targetUid,
      berakhirPada: pastStr,
      statusPersetujuan: "diblokir",
      updatedAt: nowIso,
    };

    hubUsersRef.current.set(targetUid, updatedUser);
    setUsers(prev => prev.map(u => (u.uid === userId || u.id === userId || u.uid === targetUid) ? { ...u, berakhirPada: pastStr, statusPersetujuan: "diblokir", updatedAt: nowIso } : u));

    const keysToUpdate = Array.from(new Set([
      `offline_user_${targetUid}`,
      `offline_user_${userId}`,
      userObj?.uid ? `offline_user_${userObj.uid}` : ''
    ].filter(Boolean)));

    keysToUpdate.forEach((k) => {
      const savedStr = localStorage.getItem(k);
      if (savedStr) {
        try {
          const parsed = JSON.parse(savedStr);
          localStorage.setItem(k, JSON.stringify({ ...parsed, berakhirPada: pastStr, statusPersetujuan: "diblokir", updatedAt: nowIso }));
        } catch (e) {}
      }
    });

    if (userObj) {
      syncUserToRealtimeHub(updatedUser).catch(() => {});
    }
    const docIdsToUpdate = Array.from(new Set([targetUid, userId, userObj?.uid].filter(Boolean)));
    for (const docId of docIdsToUpdate) {
      try {
        const userRef = doc(db, "users", docId);
        await fetchWithTimeout(
          setDoc(userRef, {
            berakhirPada: pastStr,
            statusPersetujuan: "diblokir",
            updatedAt: serverTimestamp()
          }, { merge: true }).then(() => reportFirestoreWriteResult(true)),
          2500,
          null
        );
      } catch (err) {
        reportFirestoreWriteResult(false, err);
        console.error("Gagal memutuskan akses:", err);
      }
    }
  };

  const handleAdminResetPassword = async (userId: string, email: string) => {
    const userObj = users.find(u => u.uid === userId || u.id === userId);
    const currentSandi = userObj?.sandi || "";
    const promptMsg = currentSandi 
      ? `Kata sandi saat ini untuk ${email} adalah: "${currentSandi}"\n\nMasukkan kata sandi baru (minimal 6 karakter):`
      : `Masukkan kata sandi baru untuk ${email} (minimal 6 karakter):`;
    const input = prompt(promptMsg, currentSandi);
    if (!input) return;
    if (input.trim().length < 6) {
      alert("Kata sandi minimal 6 karakter demi keamanan.");
      return;
    }
    const newSandi = input.trim();
    try {
      const updates = { sandi: newSandi, updatedAt: new Date().toISOString() };
      setUsers(prev => prev.map(u => (u.uid === userId || u.id === userId) ? { ...u, ...updates } : u));
      
      const userEmail = (email || "").toLowerCase().trim();
      const emailSlug = userEmail.replace(/[@.]/g, '_');
      const customUid = `custom_user_${emailSlug}`;

      const keysToUpdate = Array.from(new Set([
        `offline_user_${userId}`,
        `offline_user_${customUid}`,
        userObj?.uid ? `offline_user_${userObj.uid}` : '',
        emailSlug ? `offline_user_custom_user_${emailSlug}` : ''
      ].filter(Boolean)));

      keysToUpdate.forEach((k) => {
        const savedStr = localStorage.getItem(k);
        const parsed = savedStr ? (JSON.parse(savedStr) || {}) : (userObj || {});
        localStorage.setItem(k, JSON.stringify({ ...parsed, ...updates }));
      });

      if (userObj) {
        syncUserToRealtimeHub({ ...userObj, ...updates, uid: customUid }).catch(() => {});
      }
      const docIdsToUpdate = Array.from(new Set([userId, customUid, userObj?.uid].filter(Boolean)));
      for (const docId of docIdsToUpdate) {
        try {
          const userRef = doc(db, "users", docId);
          await fetchWithTimeout(setDoc(userRef, updates, { merge: true }), 2500, null);
        } catch (e) {}
      }
      alert(`Kata sandi untuk ${email} berhasil diperbarui menjadi: "${newSandi}"`);
    } catch (err: any) {
      alert("Gagal memperbarui kata sandi: " + err.message);
    }
  };

  // Handle delete user
  const handleDeleteUser = async (userId: string, email: string) => {
    const lowerEmail = (email || "").toLowerCase().trim();
    if (userId === "admin_syukriyusuf82" || isMainAdminEmail(lowerEmail)) {
      alert(`Admin Utama (${email}) tidak dapat dihapus!`);
      return;
    }

    if (confirm(`Apakah Anda yakin ingin menghapus data akun ${email}? Data pengguna akan terhapus permanen dari server.`)) {
      const emailSlug = lowerEmail.replace(/[@.]/g, "_");
      const customUid = `custom_user_${emailSlug}`;
      const idsToDelete = Array.from(new Set([userId, customUid].filter(Boolean)));

      // 1. Mark email as deleted & delete from Real-Time Hub + Firestore
      addDeletedTenantEmail(lowerEmail);
      hubUsersRef.current.delete(userId);
      hubUsersRef.current.delete(customUid);
      deleteUserFromRealtimeHub(userId, lowerEmail).catch(() => {});
      for (const id of idsToDelete) {
        try {
          await fetchWithTimeout(deleteDoc(doc(db, "users", id)), 2500, null);
        } catch (err) {
          console.warn(`Peringatan saat menghapus dokumen Firestore ${id}:`, err);
        }
      }

      // 2. Purge local storage offline caches
      try {
        idsToDelete.forEach((id) => {
          localStorage.removeItem(`offline_user_${id}`);
        });
        localStorage.removeItem(`offline_user_custom_user_${emailSlug}`);

        const keysToRemove: string[] = [];
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key && (key.includes(userId) || key.includes(emailSlug) || key.includes(lowerEmail))) {
            keysToRemove.push(key);
          }
        }
        keysToRemove.forEach((k) => localStorage.removeItem(k));
      } catch (e) {
        console.warn("Notice clearing local storage for deleted user:", e);
      }

      // 3. Purge active session if it was logged in
      const currentUid = localStorage.getItem("custom_logged_in_uid");
      if (currentUid && idsToDelete.includes(currentUid)) {
        localStorage.removeItem("custom_logged_in_uid");
        localStorage.removeItem("sisper_user_profile");
      }

      // 4. Update UI immediately
      setUsers(prev => prev.filter(u => !idsToDelete.includes(u.uid) && !idsToDelete.includes(u.id) && (u.email || "").toLowerCase().trim() !== lowerEmail));
      alert(`Data pengguna ${email} berhasil terhapus permanen dari server dan sistem.`);
    }
  };

  // Handle clearing marked/selected users
  const handleClearSelectedUsers = async () => {
    if (selectedUserUids.length === 0) {
      alert("Silakan tandai / centang pengguna yang ingin dibersihkan terlebih dahulu.");
      return;
    }

    const markedUsers = users.filter(u => selectedUserUids.includes(u.uid));
    const confirmMsg = `Peringatan: Anda akan menghapus & mereset SELURUH data (${markedUsers.length} akun) pengguna yang ditandai secara permanen dari server.\n\nApakah Anda yakin ingin melanjutkan?`;
    
    if (confirm(confirmMsg)) {
      setLoading(true);
      
      const deletePromises = markedUsers.map(async (u) => {
        const lowerEmail = (u.email || "").toLowerCase().trim();
        addDeletedTenantEmail(lowerEmail);
        const emailSlug = lowerEmail.replace(/[@.]/g, "_");
        const customUid = `custom_user_${emailSlug}`;
        const idsToDelete = Array.from(new Set([u.uid, u.id, customUid].filter(Boolean)));
        idsToDelete.forEach((id) => hubUsersRef.current.delete(id));
        deleteUserFromRealtimeHub(u.uid || customUid, lowerEmail).catch(() => {});

        for (const id of idsToDelete) {
          try {
            await fetchWithTimeout(deleteDoc(doc(db, "users", id)), 3000, null);
          } catch (err) {
            console.warn(`Catatan penghapusan dokumen ${id}:`, err);
          }
          localStorage.removeItem(`offline_user_${id}`);
        }
        localStorage.removeItem(`offline_user_custom_user_${emailSlug}`);
      });

      await Promise.all(deletePromises);

      // Clean local storage
      const clearedUids = new Set(markedUsers.map(m => m.uid));
      setUsers(prev => prev.filter(u => !clearedUids.has(u.uid)));
      setSelectedUserUids([]);
      setLoading(false);
      alert(`Pembersihan Berhasil!\n\nData ${markedUsers.length} pengguna yang ditandai telah dihapus secara permanen dari server.`);
    }
  };

  // Handle clearing all users with the "USER" role
  const handleClearAllUsers = async () => {
    // Exclude main admin
    const usersWithUserRole = users.filter(u => {
      const emailLower = (u.email || "").toLowerCase().trim();
      const isAdmin = u.uid === "admin_syukriyusuf82" || isMainAdminEmail(emailLower) || u.peran === "ADMIN";
      return !isAdmin;
    });

    if (usersWithUserRole.length === 0) {
      alert("Tidak ada data pendaftar pengguna (role USER) yang perlu dibersihkan.");
      return;
    }

    const confirmMsg = `Peringatan: Anda akan menghapus & mereset SELURUH data pendaftar (${usersWithUserRole.length} akun) secara permanen dari server.\n\nApakah Anda yakin ingin melanjutkan?`;
    if (confirm(confirmMsg)) {
      setLoading(true);

      const deletePromises = usersWithUserRole.map(async (u) => {
        const lowerEmail = (u.email || "").toLowerCase().trim();
        addDeletedTenantEmail(lowerEmail);
        const emailSlug = lowerEmail.replace(/[@.]/g, "_");
        const customUid = `custom_user_${emailSlug}`;
        const idsToDelete = Array.from(new Set([u.uid, u.id, customUid].filter(Boolean)));
        idsToDelete.forEach((id) => hubUsersRef.current.delete(id));
        deleteUserFromRealtimeHub(u.uid || customUid, lowerEmail).catch(() => {});

        for (const id of idsToDelete) {
          try {
            await fetchWithTimeout(deleteDoc(doc(db, "users", id)), 3000, null);
          } catch (err) {
            console.warn(`Notice deleting document ${id}:`, err);
          }
          localStorage.removeItem(`offline_user_${id}`);
        }
        localStorage.removeItem(`offline_user_custom_user_${emailSlug}`);
      });

      await Promise.all(deletePromises);

      // Clean all non-admin offline_user_ keys in localStorage
      try {
        const keysToRemove: string[] = [];
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key && key.startsWith("offline_user_") && !key.includes("syukriyusuf") && !key.includes("sukriyusuf")) {
            keysToRemove.push(key);
          }
        }
        keysToRemove.forEach(k => localStorage.removeItem(k));
      } catch (e) {
        console.warn("Notice clearing offline_user_ keys:", e);
      }

      setUsers(prev => prev.filter(u => u.uid === "admin_syukriyusuf82" || isMainAdminEmail(u.email)));
      setSelectedUserUids([]);
      setLoading(false);
      alert(`Pembersihan & Reset Data Berhasil!\n\nSeluruh data pendaftar (${usersWithUserRole.length} akun) telah terhapus permanen dari server.`);
    }
  };

  // Filter users by search query
  const filteredUsers = users.filter((u) => {
    const q = searchQuery.toLowerCase();
    return (
      (u.email || "").toLowerCase().includes(q) ||
      (u.namaLengkap || "").toLowerCase().includes(q) ||
      (u.profesi || "").toLowerCase().includes(q) ||
      (u.namaSPPG || "").toLowerCase().includes(q)
    );
  });

  // Selection helpers
  const selectableUsers = filteredUsers.filter(u => u.uid !== "admin_syukriyusuf82" && u.uid !== "admin_sukriyusuf82" && !isMainAdminEmail(u.email));
  const isAllSelected = selectableUsers.length > 0 && selectableUsers.every(u => selectedUserUids.includes(u.uid));

  const toggleSelectUser = (uid: string) => {
    setSelectedUserUids(prev =>
      prev.includes(uid) ? prev.filter(id => id !== uid) : [...prev, uid]
    );
  };

  const toggleSelectAll = () => {
    if (isAllSelected) {
      setSelectedUserUids([]);
    } else {
      setSelectedUserUids(selectableUsers.map(u => u.uid));
    }
  };

  // Calculate metrics based on actual non-admin tenants & live devices
  const tenantUsers = users.filter(u => u.uid !== "admin_syukriyusuf82" && u.uid !== "admin_sukriyusuf82" && !isMainAdminEmail(u.email));
  const totalUsersCount = tenantUsers.length;
  const activeUsersCount = tenantUsers.filter(u => u.statusPersetujuan === "aktif").length;
  const pendingUsersCount = tenantUsers.filter(u => u.statusPersetujuan === "menunggu" || u.statusPersetujuan === "pending").length;
  const onlineDevicesCount = devices.filter(d => isDeviceLiveOnline(d)).length;

  const filteredDevices = devices.filter((d) => {
    if (deviceFilter === "online" && !isDeviceLiveOnline(d)) return false;
    if (deviceFilter === "Desktop / Laptop" && d.deviceType !== "Desktop / Laptop") return false;
    if (deviceFilter === "Mobile" && d.deviceType !== "Mobile") return false;
    if (deviceFilter === "Tablet" && d.deviceType !== "Tablet") return false;
    if (deviceFilter === "guest" && d.sessionState !== "guest" && d.sessionState !== "registering") return false;
    if (deviceFilter === "demo" && d.sessionState !== "demo") return false;

    const q = deviceSearchQuery.toLowerCase().trim();
    if (!q) return true;
    return (
      (d.deviceModel || "").toLowerCase().includes(q) ||
      (d.osName || "").toLowerCase().includes(q) ||
      (d.browserName || "").toLowerCase().includes(q) ||
      (d.deviceId || "").toLowerCase().includes(q) ||
      (d.ipAddress || "").toLowerCase().includes(q) ||
      (d.userEmail || "").toLowerCase().includes(q) ||
      (d.userName || "").toLowerCase().includes(q) ||
      (d.namaSPPG || "").toLowerCase().includes(q)
    );
  });
  
  // Calculate users logged in last 24h
  const oneDayAgo = Date.now() - 24 * 60 * 60 * 1000;
  const login24hCount = tenantUsers.filter(u => {
    if (!u.loginTerakhir) return false;
    const loginTime = u.loginTerakhir?.seconds ? u.loginTerakhir.seconds * 1000 : new Date(u.loginTerakhir).getTime();
    return loginTime > oneDayAgo;
  }).length;

  // Format date display
  const formatTimestamp = (ts: any) => {
    if (!ts) return "-";
    let dateObj: Date;
    if (ts.seconds) {
      dateObj = new Date(ts.seconds * 1000);
    } else {
      dateObj = new Date(ts);
    }
    return dateObj.toLocaleString("id-ID", {
      year: "numeric",
      month: "numeric",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  };

  return (
    <div className="bg-slate-50 rounded-2xl border border-slate-200 shadow-sm overflow-hidden p-6 space-y-6">
      
      {/* Title Header with Close */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 pb-4">
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-indigo-100 text-indigo-600 rounded-xl">
            <Shield className="w-6 h-6" />
          </div>
          <div>
            <h2 className="text-xl font-extrabold text-slate-900 tracking-tight">Panel Administrasi SISPERMBG</h2>
            <div className="flex flex-wrap items-center gap-2 mt-0.5">
              <p className="text-xs text-slate-500">Multi-tenant Management Hub • Sinkronisasi Real-Time HP & Laptop</p>
              <span
                className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[10px] font-extrabold border ${
                  sseStatus === "connected"
                    ? "bg-emerald-50 text-emerald-700 border-emerald-200"
                    : "bg-amber-50 text-amber-700 border-amber-200"
                }`}
              >
                <span
                  className={`w-1.5 h-1.5 rounded-full ${
                    sseStatus === "connected" ? "bg-emerald-500 animate-pulse" : "bg-amber-500"
                  }`}
                ></span>
                {sseStatus === "connected" ? "Sinkronisasi Real-Time Aktif" : "Menghubungkan Real-Time..."}
              </span>
            </div>
          </div>
        </div>
        <button
          onClick={onClose}
          className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-xl transition cursor-pointer self-start sm:self-auto"
        >
          Tutup Panel
        </button>
      </div>

      {/* Honest Firestore Daily Write Quota Banner if Free Tier cap reached */}
      {firestoreHealth.writeQuotaExceeded && (
        <div className="bg-amber-50 border border-amber-300 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="flex items-start gap-3">
            <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
            <div className="text-xs text-amber-950 space-y-1">
              <div className="font-extrabold">
                Kuota Harian Tulis Firestore (Free Tier) Sedang Penuh — Sinkronisasi Real-Time Tetap Berjalan via Server Relay
              </div>
              <p className="text-[11px] text-amber-800 leading-relaxed">
                Aktivitas pendaftaran, login, dan deteksi perangkat seluler (HP) ke laptop tetap tampil <strong>secara real-time</strong> melalui jalur sinkronisasi server langsung. Kuota tulis Firestore akan otomatis reset keesokan hari, atau Anda dapat menaikkan paket database di Firebase Console.
              </p>
            </div>
          </div>
          <a
            href={FIRESTORE_UPGRADE_URL}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-amber-600 hover:bg-amber-700 text-white font-bold text-[11px] rounded-xl transition shrink-0 self-start sm:self-center"
          >
            <span>Buka Firebase Console</span>
            <ExternalLink className="w-3.5 h-3.5" />
          </a>
        </div>
      )}

      {/* Global App Logo Card for Admin / Panel Admin holders */}
      {onUploadGlobalLogo && (
        <div className="bg-white p-4 rounded-2xl border border-indigo-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="flex items-center gap-3.5">
            <div className="w-14 h-14 rounded-full border-2 border-emerald-400 overflow-hidden bg-white shrink-0 flex items-center justify-center shadow-xs">
              <img
                src={globalLogoUrl || "/assets/images/app_logo_sppg.png"}
                alt="Logo Aplikasi Global"
                className="w-full h-full object-cover"
                referrerPolicy="no-referrer"
                onError={(e) => {
                  (e.target as HTMLImageElement).src = "/assets/images/app_logo_sppg.png";
                }}
              />
            </div>
            <div>
              <h3 className="text-xs font-extrabold text-slate-900 uppercase tracking-wide">
                Logo Aplikasi Global (Login, Header & Kop Surat)
              </h3>
              <p className="text-[11px] text-slate-500 max-w-xl">
                Perubahan logo di sini otomatis diterapkan ke halaman masuk (login) dan seluruh akun pengguna aktif secara real-time.
              </p>
            </div>
          </div>
          <label className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl transition cursor-pointer inline-flex items-center gap-2 shrink-0 shadow-xs">
            <Camera className="w-4 h-4" />
            <span>Ubah Logo Aplikasi Global</span>
            <input
              type="file"
              accept="image/*"
              onChange={onUploadGlobalLogo}
              className="hidden"
            />
          </label>
        </div>
      )}

      {/* 4 Stats Cards matching Image 1 */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Total Pengguna */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-xs flex items-center gap-4">
          <div className="p-3.5 bg-blue-50 text-blue-600 rounded-xl">
            <Users className="w-6 h-6" />
          </div>
          <div>
            <p className="text-xs font-semibold text-slate-400">Total Pengguna</p>
            <p className="text-2xl font-extrabold text-slate-900">{totalUsersCount}</p>
          </div>
        </div>

        {/* Pengguna Aktif */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-xs flex items-center gap-4">
          <div className="p-3.5 bg-emerald-50 text-emerald-600 rounded-xl">
            <CheckCircle className="w-6 h-6" />
          </div>
          <div>
            <p className="text-xs font-semibold text-slate-400">Pengguna Aktif (ACC)</p>
            <div className="flex items-baseline gap-2">
              <p className="text-2xl font-extrabold text-slate-900 tabular-nums">{activeUsersCount}</p>
              {pendingUsersCount > 0 && (
                <span className="text-[10px] font-bold text-amber-800 bg-amber-100 px-1.5 py-0.5 rounded">
                  {pendingUsersCount} Menunggu ACC
                </span>
              )}
            </div>
          </div>
        </div>

        {/* Login 24 Jam Terakhir */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-xs flex items-center gap-4">
          <div className="p-3.5 bg-purple-50 text-purple-600 rounded-xl">
            <Clock className="w-6 h-6" />
          </div>
          <div>
            <p className="text-xs font-semibold text-slate-400">Login 24 Jam Terakhir</p>
            <p className="text-2xl font-extrabold text-slate-900">{login24hCount}</p>
          </div>
        </div>

        {/* Perangkat Terdeteksi Otomatis */}
        <button
          type="button"
          onClick={() => setActiveTab("perangkat")}
          className="bg-white p-5 rounded-2xl border border-slate-100 hover:border-teal-300 shadow-xs flex items-center gap-4 text-left transition cursor-pointer"
          title="Klik untuk melihat seluruh device yang terbaca otomatis"
        >
          <div className="p-3.5 bg-teal-50 text-teal-600 rounded-xl">
            <Monitor className="w-6 h-6" />
          </div>
          <div>
            <p className="text-xs font-semibold text-slate-400">Perangkat Terbaca</p>
            <div className="flex items-center gap-2">
              <p className="text-2xl font-extrabold text-slate-900">{devices.length}</p>
              <span className="inline-flex items-center gap-1 text-[10px] font-extrabold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                {onlineDevicesCount} Online
              </span>
            </div>
          </div>
        </button>
      </div>

      {/* Navigation Sub-Tabs */}
      <div className="flex flex-wrap border-b border-slate-200">
        <button
          onClick={() => setActiveTab("manajemen")}
          className={`flex items-center gap-2 px-5 py-3 text-xs font-extrabold border-b-2 transition cursor-pointer ${
            activeTab === "manajemen"
              ? "border-indigo-600 text-indigo-600"
              : "border-transparent text-slate-500 hover:text-slate-900"
          }`}
        >
          <Users className="w-4 h-4" />
          Manajemen Pengguna
        </button>
        <button
          onClick={() => setActiveTab("perangkat")}
          className={`flex items-center gap-2 px-5 py-3 text-xs font-extrabold border-b-2 transition cursor-pointer ${
            activeTab === "perangkat"
              ? "border-indigo-600 text-indigo-600"
              : "border-transparent text-slate-500 hover:text-slate-900"
          }`}
        >
          <Monitor className="w-4 h-4" />
          Perangkat (Device)
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-800">
            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
            {devices.length}
          </span>
        </button>
        <button
          onClick={() => setActiveTab("riwayat")}
          className={`flex items-center gap-2 px-5 py-3 text-xs font-extrabold border-b-2 transition cursor-pointer ${
            activeTab === "riwayat"
              ? "border-indigo-600 text-indigo-600"
              : "border-transparent text-slate-500 hover:text-slate-900"
          }`}
        >
          <History className="w-4 h-4" />
          Riwayat Login
        </button>
        <button
          onClick={() => setActiveTab("diagnostik")}
          className={`flex items-center gap-2 px-5 py-3 text-xs font-extrabold border-b-2 transition cursor-pointer ${
            activeTab === "diagnostik"
              ? "border-indigo-600 text-indigo-600"
              : "border-transparent text-slate-500 hover:text-slate-900"
          }`}
        >
          <Cpu className="w-4 h-4" />
          Sistem Diagnostik
        </button>
      </div>

      {/* Active Tab Content Area */}
      {activeTab === "manajemen" && (
        <div className="space-y-4">
          {/* Search bar & Refresh Button */}
          <div className="flex flex-col sm:flex-row gap-3 items-center justify-between bg-white p-4 rounded-xl border border-slate-100 shadow-2xs">
            <div className="relative w-full max-w-md">
              <Search className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Cari berdasarkan email, nama, atau instansi..."
                className="w-full pl-10 pr-4 py-2 text-xs border border-slate-200 rounded-xl bg-slate-50 focus:outline-none focus:ring-2 focus:ring-indigo-500/30 focus:bg-white"
              />
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {selectedUserUids.length > 0 && (
                <button
                  onClick={handleClearSelectedUsers}
                  className="flex items-center gap-1.5 px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-extrabold text-xs rounded-xl transition cursor-pointer shadow-sm animate-bounce"
                  title="Bersihkan data pengguna yang ditandai saja"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  Bersihkan Terpilih ({selectedUserUids.length})
                </button>
              )}
              <button
                onClick={handleClearAllUsers}
                className="flex items-center gap-1.5 px-4 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 font-bold text-xs rounded-xl transition cursor-pointer"
                title="Bersihkan seluruh pengguna dengan peran 'USER' dari sistem secara permanen"
              >
                <Trash2 className="w-3.5 h-3.5" />
                Bersihkan Semua USER
              </button>
              <button
                onClick={() => setRefreshTrigger(prev => prev + 1)}
                className="flex items-center gap-1.5 px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition cursor-pointer"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${loading ? "animate-spin" : ""}`} />
                Segarkan Data
              </button>
            </div>
          </div>

          {/* User List Table */}
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-2xs">
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-left text-xs font-sans">
                <thead className="bg-slate-50 text-slate-500 font-bold uppercase tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="p-4 text-center w-10">
                      <input
                        type="checkbox"
                        checked={isAllSelected}
                        onChange={toggleSelectAll}
                        className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer w-4 h-4"
                        title="Tandai / Centang Semua Pengguna"
                      />
                    </th>
                    <th className="p-4">Email / Tenant ID</th>
                    <th className="p-4">Data Diri & SPPG</th>
                    <th className="p-4">Perangkat (Device)</th>
                    <th className="p-4 text-center">Peran</th>
                    <th className="p-4 text-center">Status & Verifikasi</th>
                    <th className="p-4 text-center">Berakhir Pada</th>
                    <th className="p-4 text-center">Login Terakhir</th>
                    <th className="p-4 text-center">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-slate-700">
                  {loading ? (
                    <tr>
                      <td colSpan={9} className="p-8 text-center text-slate-400">
                        <RefreshCw className="w-6 h-6 animate-spin mx-auto mb-2 text-indigo-500" />
                        Sedang menyinkronkan data dari Firestore...
                      </td>
                    </tr>
                  ) : filteredUsers.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="p-8 text-center text-slate-400">
                        Tidak ada tenant atau pengguna ditemukan.
                      </td>
                    </tr>
                  ) : (
                    filteredUsers.map((item) => (
                      <tr key={item.uid} className={`hover:bg-slate-50/50 transition ${selectedUserUids.includes(item.uid) ? "bg-indigo-50/30" : ""}`}>
                        {/* Checkbox Tandai Pengguna */}
                        <td className="p-4 text-center">
                          {item.uid !== "admin_syukriyusuf82" && !isMainAdminEmail(item.email) ? (
                            <input
                              type="checkbox"
                              checked={selectedUserUids.includes(item.uid)}
                              onChange={() => toggleSelectUser(item.uid)}
                              className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer w-4 h-4"
                            />
                          ) : (
                            <Shield className="w-4 h-4 text-purple-500 mx-auto" title="Admin Utama" />
                          )}
                        </td>
                        {/* Email/Tenant UID */}
                        <td className="p-4">
                          <div className="font-semibold text-slate-900 flex items-center gap-1.5">
                            <Mail className="w-3.5 h-3.5 text-slate-400" />
                            {item.email}
                          </div>
                          <div className="text-[10px] text-slate-400 font-mono mt-0.5">{item.uid}</div>
                        </td>

                        {/* Data Diri & SPPG */}
                        <td className="p-4 space-y-1">
                          {item.namaLengkap ? (
                            <>
                              <div className="font-semibold text-slate-800 flex items-center gap-1">
                                <User className="w-3 h-3 text-slate-400" />
                                {item.namaLengkap}
                              </div>
                              <div className="text-[10px] text-slate-500 flex items-center gap-1">
                                <Briefcase className="w-3 h-3 text-slate-400" />
                                Profesi: {item.profesi}
                              </div>
                              <div className="text-[10px] text-indigo-600 font-bold flex items-center gap-1">
                                <Building className="w-3 h-3 text-indigo-400" />
                                {item.namaSPPG}
                              </div>
                              <div className="text-[10px] text-slate-400">WA: {item.noHp || "-"}</div>
                            </>
                          ) : (
                            <span className="text-slate-400 italic">Belum mengisi pendaftaran</span>
                          )}
                        </td>

                        {/* Perangkat (Device) Terdeteksi Otomatis */}
                        <td className="p-4">
                          {(() => {
                            const userDevs = getDevicesForUser(item);
                            if (userDevs.length === 0) {
                              return <span className="text-[10.5px] text-slate-400 italic">Belum terdeteksi</span>;
                            }
                            return (
                              <div className="space-y-1.5 max-w-[220px]">
                                {userDevs.slice(0, 2).map((dev) => {
                                  const live = isDeviceLiveOnline(dev);
                                  return (
                                    <div
                                      key={dev.deviceId}
                                      className="p-1.5 rounded-lg bg-slate-50 border border-slate-200/80 text-[10px] space-y-0.5"
                                    >
                                      <div className="flex items-center justify-between gap-1.5">
                                        <span className="font-bold text-slate-800 flex items-center gap-1 truncate">
                                          {renderDeviceIcon(dev.deviceType, "w-3 h-3 text-indigo-500 shrink-0")}
                                          <span className="truncate">{dev.deviceModel || dev.osName}</span>
                                        </span>
                                        <span
                                          className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded-full text-[9px] font-extrabold shrink-0 ${
                                            live
                                              ? "bg-emerald-100 text-emerald-700"
                                              : "bg-slate-200/80 text-slate-600"
                                          }`}
                                        >
                                          <span
                                            className={`w-1.5 h-1.5 rounded-full ${
                                              live ? "bg-emerald-500 animate-pulse" : "bg-slate-400"
                                            }`}
                                          ></span>
                                          {live ? "Online" : formatRelativeTime(dev.lastSeenAt)}
                                        </span>
                                      </div>
                                      <div className="text-slate-500 truncate">
                                        {dev.osName} · {dev.browserName}
                                      </div>
                                      {dev.ipAddress && dev.ipAddress !== "-" && (
                                        <div className="text-[9.5px] font-mono text-slate-400 truncate">
                                          IP: {dev.ipAddress}
                                        </div>
                                      )}
                                    </div>
                                  );
                                })}
                                {userDevs.length > 2 && (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setDeviceSearchQuery(item.email || "");
                                      setActiveTab("perangkat");
                                    }}
                                    className="text-[10px] font-bold text-indigo-600 hover:underline cursor-pointer"
                                  >
                                    +{userDevs.length - 2} perangkat lainnya
                                  </button>
                                )}
                              </div>
                            );
                          })()}
                        </td>

                        {/* Peran / Role */}
                        <td className="p-4 text-center">
                          <button
                            onClick={() => handleToggleRole(item.uid, item.peran)}
                            className={`px-2.5 py-1 text-[10px] font-extrabold rounded-md uppercase transition cursor-pointer ${
                              item.peran === "ADMIN"
                                ? "bg-purple-100 text-purple-700 hover:bg-purple-200"
                                : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                            }`}
                          >
                            {item.peran || "USER"}
                          </button>
                        </td>

                        {/* Status Persetujuan (ACC Admin) */}
                        <td className="p-4 text-center">
                          <div className="flex flex-col items-center gap-1.5">
                            {item.statusPersetujuan === "aktif" ? (
                              <div className="flex flex-col items-center gap-1">
                                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 text-[10.5px] font-extrabold rounded-md bg-emerald-100 text-emerald-800 border border-emerald-300">
                                  <CheckCircle className="w-3 h-3 text-emerald-600" />
                                  Disetujui (ACC)
                                </span>
                                {item.authProvider === "google" && (
                                  <span className="text-[9.5px] text-slate-500 font-semibold">
                                    Akun Email Google
                                  </span>
                                )}
                                {!isMainAdminEmail(item.email) && item.uid !== "admin_syukriyusuf82" && (
                                  <div className="flex items-center gap-2 pt-0.5">
                                    <button
                                      type="button"
                                      onClick={() => handleSetStatus(item.uid || item.id, "menunggu")}
                                      className="text-[10px] text-amber-700 hover:text-amber-900 font-bold underline cursor-pointer"
                                      title="Kembalikan status ke Menunggu ACC"
                                    >
                                      Tahan (Menunggu)
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => handleSetStatus(item.uid || item.id, "diblokir")}
                                      className="text-[10px] text-rose-600 hover:text-rose-800 font-bold underline cursor-pointer"
                                      title="Putuskan akses pengguna ini"
                                    >
                                      Putuskan
                                    </button>
                                  </div>
                                )}
                              </div>
                            ) : item.statusPersetujuan === "menunggu" || item.statusPersetujuan === "pending" ? (
                              <div className="flex flex-col items-center gap-1.5">
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-extrabold rounded-md bg-amber-100 text-amber-800 border border-amber-300">
                                  <Clock className="w-3 h-3 text-amber-600" />
                                  Menunggu ACC
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleSetStatus(item.uid || item.id, "aktif")}
                                  className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-black rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white transition shadow-sm cursor-pointer"
                                  title="Klik untuk menyetujui (ACC) dan membuka akses penuh dashboard pengguna ini"
                                >
                                  <CheckCircle className="w-3.5 h-3.5 text-white" />
                                  ACC / Setujui
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleSetStatus(item.uid || item.id, "diblokir")}
                                  className="text-[10px] text-rose-600 hover:text-rose-800 font-bold cursor-pointer"
                                >
                                  Tolak & Blokir
                                </button>
                              </div>
                            ) : (
                              <div className="flex flex-col items-center gap-1.5">
                                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 text-[10.5px] font-extrabold rounded-md bg-rose-100 text-rose-800 border border-rose-300">
                                  <XCircle className="w-3 h-3 text-rose-600" />
                                  Diblokir
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleSetStatus(item.uid || item.id, "aktif")}
                                  className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white font-extrabold text-[10px] rounded-md transition cursor-pointer shadow-xs"
                                  title="Aktifkan kembali akses akun"
                                >
                                  Buka Akses (ACC)
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleSetStatus(item.uid || item.id, "menunggu")}
                                  className="text-[10px] text-amber-700 hover:text-amber-900 font-bold underline cursor-pointer"
                                >
                                  Set Menunggu
                                </button>
                              </div>
                            )}
                          </div>
                        </td>

                        {/* Berakhir Pada */}
                        <td className="p-4 text-center">
                          <div className="flex flex-col items-center gap-1">
                            <input
                              type="date"
                              value={item.berakhirPada || ""}
                              onChange={(e) => handleUpdateExpiration(item.uid || item.id, e.target.value)}
                              className="p-1 border border-slate-200 rounded text-xs bg-slate-50 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono w-[130px] text-center"
                            />
                            
                            {/* Status Expiration Check */}
                            {(() => {
                              const now = new Date();
                              const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
                              const isExpired = item.berakhirPada ? todayStr > item.berakhirPada : false;
                              if (isExpired && item.statusPersetujuan === "aktif") {
                                return <span className="text-[9.5px] font-bold text-rose-600 bg-rose-50 px-1.5 py-0.5 rounded border border-rose-200">⚠️ Kedaluwarsa</span>;
                              }
                              return null;
                            })()}

                            {/* Quick Presets */}
                            <div className="flex items-center gap-1 text-[9px] font-bold text-slate-500 mt-0.5">
                              <button
                                onClick={() => handleAddDuration(item.uid || item.id, 1)}
                                className="px-1.5 py-0.5 bg-slate-100 hover:bg-indigo-100 hover:text-indigo-700 rounded transition cursor-pointer"
                                title="+1 Bulan masa aktif"
                              >
                                +1Bln
                              </button>
                              <button
                                onClick={() => handleAddDuration(item.uid || item.id, 6)}
                                className="px-1.5 py-0.5 bg-slate-100 hover:bg-indigo-100 hover:text-indigo-700 rounded transition cursor-pointer"
                                title="+6 Bulan masa aktif"
                              >
                                +6Bln
                              </button>
                              <button
                                onClick={() => handleAddDuration(item.uid || item.id, 12)}
                                className="px-1.5 py-0.5 bg-slate-100 hover:bg-indigo-100 hover:text-indigo-700 rounded transition cursor-pointer"
                                title="+1 Tahun masa aktif"
                              >
                                +1Thn
                              </button>
                              <button
                                onClick={() => handleRevokeAccessNow(item.uid || item.id)}
                                className="px-1.5 py-0.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded transition cursor-pointer"
                                title="Set tanggal kedaluwarsa ke kemarin (Putuskan Akses Langsung)"
                              >
                                Putuskan
                              </button>
                            </div>
                          </div>
                        </td>

                        {/* Login Terakhir */}
                        <td className="p-4 text-center font-mono text-slate-500">
                          {formatTimestamp(item.loginTerakhir)}
                        </td>

                        {/* Actions */}
                        <td className="p-4 text-center">
                          <div className="flex items-center justify-center gap-1.5">
                            {item.sandi && item.authProvider !== "google" && (
                              <button
                                onClick={() => handleAdminResetPassword(item.uid || item.id, item.email)}
                                className="p-1.5 bg-amber-50 hover:bg-amber-100 text-amber-700 rounded-lg transition cursor-pointer"
                                title="Lihat / Reset Kata Sandi Legacy"
                              >
                                <Key className="w-3.5 h-3.5" />
                              </button>
                            )}
                            {!isMainAdminEmail(item.email) && item.uid !== "admin_syukriyusuf82" && (
                              <button
                                onClick={() => handleDeleteUser(item.uid, item.email)}
                                className="p-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg transition cursor-pointer"
                                title="Hapus Pengguna Permanen"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {activeTab === "perangkat" && (
        <div className="space-y-4">
          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs space-y-3">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
              <div>
                <h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-2">
                  <Monitor className="w-4 h-4 text-indigo-600" />
                  Monitoring Perangkat (Device) Otomatis Real-Time
                </h3>
                <p className="text-[11px] text-slate-500 mt-0.5">
                  Setiap perangkat (HP, Tablet, Laptop/PC) yang mengakses, mendaftar, atau masuk ke aplikasi terbaca otomatis secara real-time.
                </p>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                <div className="relative w-full sm:w-64">
                  <Search className="absolute left-3 top-2.5 w-3.5 h-3.5 text-slate-400" />
                  <input
                    type="text"
                    value={deviceSearchQuery}
                    onChange={(e) => setDeviceSearchQuery(e.target.value)}
                    placeholder="Cari device, OS, browser, email, IP..."
                    className="w-full pl-9 pr-3 py-2 text-xs border border-slate-200 rounded-xl bg-slate-50 focus:outline-none focus:ring-2 focus:ring-indigo-500/30 focus:bg-white"
                  />
                </div>
                <button
                  onClick={() => setRefreshTrigger((prev) => prev + 1)}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition cursor-pointer"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  Segarkan
                </button>
              </div>
            </div>

            {/* Device Category Filters */}
            <div className="flex flex-wrap gap-1.5 pt-1">
              {[
                { id: "all", label: `Semua (${devices.length})` },
                { id: "online", label: `Online (${onlineDevicesCount})` },
                { id: "Desktop / Laptop", label: `Desktop / Laptop (${devices.filter((d) => d.deviceType === "Desktop / Laptop").length})` },
                { id: "Mobile", label: `Mobile / HP (${devices.filter((d) => d.deviceType === "Mobile").length})` },
                { id: "Tablet", label: `Tablet (${devices.filter((d) => d.deviceType === "Tablet").length})` },
                { id: "guest", label: `Halaman Masuk / Daftar (${devices.filter((d) => d.sessionState === "guest" || d.sessionState === "registering").length})` },
                { id: "demo", label: `Demo (${devices.filter((d) => d.sessionState === "demo").length})` },
              ].map((f) => (
                <button
                  key={f.id}
                  type="button"
                  onClick={() => setDeviceFilter(f.id as any)}
                  className={`px-3 py-1.5 rounded-lg text-[11px] font-bold transition cursor-pointer ${
                    deviceFilter === f.id
                      ? "bg-indigo-600 text-white shadow-2xs"
                      : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                  }`}
                >
                  {f.label}
                </button>
              ))}
            </div>
          </div>

          {/* Real-Time Devices Table */}
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-2xs">
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-left text-xs font-sans">
                <thead className="bg-slate-50 text-slate-500 font-bold uppercase tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="p-4">Perangkat & Sistem</th>
                    <th className="p-4">Pengguna / Sesi</th>
                    <th className="p-4 text-center">Status Koneksi</th>
                    <th className="p-4">Jaringan & Layar</th>
                    <th className="p-4">Terakhir Aktif</th>
                    <th className="p-4 text-center">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-slate-700">
                  {filteredDevices.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="p-8 text-center text-slate-400">
                        Belum ada perangkat yang sesuai dengan filter saat ini.
                      </td>
                    </tr>
                  ) : (
                    filteredDevices.map((dev) => {
                      const live = isDeviceLiveOnline(dev);
                      return (
                        <tr key={dev.deviceId} className="hover:bg-slate-50/60 transition">
                          <td className="p-4">
                            <div className="flex items-start gap-2.5">
                              <div
                                className={`p-2 rounded-xl shrink-0 ${
                                  live ? "bg-indigo-50 text-indigo-600" : "bg-slate-100 text-slate-500"
                                }`}
                              >
                                {renderDeviceIcon(dev.deviceType, "w-4 h-4")}
                              </div>
                              <div className="space-y-0.5">
                                <div className="font-extrabold text-slate-900 flex items-center gap-1.5">
                                  <span>{dev.deviceModel || dev.deviceType}</span>
                                  <span className="text-[9.5px] font-bold px-1.5 py-0.5 rounded bg-slate-100 text-slate-600">
                                    {dev.deviceType}
                                  </span>
                                </div>
                                <div className="text-[11px] text-slate-600 font-medium">
                                  {dev.osName} · {dev.browserName}
                                </div>
                                <div className="text-[10px] font-mono text-slate-400">
                                  ID: {dev.deviceId}
                                </div>
                              </div>
                            </div>
                          </td>

                          <td className="p-4 space-y-0.5">
                            {dev.userEmail ? (
                              <>
                                <div className="font-bold text-slate-900 flex items-center gap-1.5">
                                  <User className="w-3 h-3 text-slate-400" />
                                  <span>{dev.userName || dev.userEmail}</span>
                                </div>
                                <div className="text-[11px] text-indigo-700 font-semibold">{dev.userEmail}</div>
                                {dev.namaSPPG && (
                                  <div className="text-[10px] text-slate-500 flex items-center gap-1">
                                    <Building className="w-3 h-3 text-slate-400" />
                                    {dev.namaSPPG}
                                  </div>
                                )}
                                <div className="flex items-center gap-1.5 pt-0.5">
                                  <span className="text-[9.5px] font-extrabold uppercase px-1.5 py-0.5 rounded bg-purple-50 text-purple-700">
                                    {dev.peran || "USER"}
                                  </span>
                                  {dev.sessionState === "registering" && (
                                    <span className="text-[9.5px] font-bold px-1.5 py-0.5 rounded bg-amber-100 text-amber-800">
                                      Verifikasi OTP
                                    </span>
                                  )}
                                  {dev.sessionState === "demo" && (
                                    <span className="text-[9.5px] font-bold px-1.5 py-0.5 rounded bg-teal-100 text-teal-800">
                                      Demo 5 Menit
                                    </span>
                                  )}
                                </div>
                              </>
                            ) : (
                              <div className="space-y-1">
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-slate-100 text-slate-600 font-bold text-[10px]">
                                  Halaman Masuk / Tamu
                                </span>
                                <div className="text-[10px] text-slate-400">Belum masuk ke akun</div>
                              </div>
                            )}
                          </td>

                          <td className="p-4 text-center">
                            {live ? (
                              <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300 font-extrabold text-[10.5px]">
                                <Wifi className="w-3 h-3 text-emerald-600" />
                                Online
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-slate-100 text-slate-600 border border-slate-200 font-bold text-[10.5px]">
                                <WifiOff className="w-3 h-3 text-slate-400" />
                                Offline
                              </span>
                            )}
                          </td>

                          <td className="p-4 space-y-0.5 text-[11px]">
                            <div className="font-mono font-bold text-slate-700 flex items-center gap-1">
                              <Globe className="w-3 h-3 text-slate-400" />
                              {dev.ipAddress || "-"}
                            </div>
                            <div className="text-[10px] text-slate-500">
                              Layar: {dev.screenResolution || "-"} · {dev.timezone || "-"}
                            </div>
                          </td>

                          <td className="p-4 space-y-0.5">
                            <div className="font-bold text-slate-800 text-[11px]">
                              {formatRelativeTime(dev.lastSeenAt)}
                            </div>
                            <div className="text-[10px] font-mono text-slate-500">
                              {formatTimestamp(dev.lastSeenAt)}
                            </div>
                            {dev.firstSeenAt && (
                              <div className="text-[9.5px] text-slate-400">
                                Pertama: {formatTimestamp(dev.firstSeenAt)}
                              </div>
                            )}
                          </td>

                          <td className="p-4 text-center">
                            <button
                              type="button"
                              onClick={() => handleDeleteDeviceRecord(dev.deviceId)}
                              className="p-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg transition cursor-pointer"
                              title="Hapus catatan perangkat ini"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {activeTab === "riwayat" && (
        <div className="bg-white p-5 rounded-xl border border-slate-200 space-y-4">
          <h3 className="text-sm font-extrabold text-slate-800 flex items-center gap-2">
            <History className="w-4 h-4 text-indigo-500" />
            Riwayat Aktivitas Login Tenant
          </h3>
          <p className="text-xs text-slate-500">Aktivitas autentikasi masuk terbaru beserta perangkat yang tercatat secara real-time di Firestore.</p>
          
          <div className="overflow-x-auto border border-slate-100 rounded-lg">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-50 text-slate-500 font-bold border-b border-slate-100">
                <tr>
                  <th className="p-3">Email Pengguna</th>
                  <th className="p-3">SPPG / Instansi</th>
                  <th className="p-3">Perangkat & IP</th>
                  <th className="p-3">Peran</th>
                  <th className="p-3">Waktu Login</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50 font-mono text-slate-600">
                {users.filter(u => u.loginTerakhir).map((user, idx) => {
                  const userDevs = getDevicesForUser(user);
                  const latestDev = userDevs[0];
                  return (
                    <tr key={idx} className="hover:bg-slate-50/50">
                      <td className="p-3 font-sans font-medium text-slate-800">{user.email}</td>
                      <td className="p-3 font-sans text-xs">{user.namaSPPG || "SPPG Pusat"}</td>
                      <td className="p-3 font-sans text-[11px] text-slate-600">
                        {latestDev ? (
                          <div className="flex items-center gap-1.5">
                            {renderDeviceIcon(latestDev.deviceType, "w-3.5 h-3.5 text-indigo-500 shrink-0")}
                            <span>
                              {formatDeviceShortLabel(latestDev)}
                              {latestDev.ipAddress && latestDev.ipAddress !== "-" ? ` (${latestDev.ipAddress})` : ""}
                            </span>
                          </div>
                        ) : (
                          user.lastDeviceLabel || "-"
                        )}
                      </td>
                      <td className="p-3"><span className="bg-purple-50 text-purple-700 px-1.5 py-0.5 rounded text-[10px]">{user.peran || "USER"}</span></td>
                      <td className="p-3">{formatTimestamp(user.loginTerakhir)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {activeTab === "diagnostik" && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          {/* Firestore Connection Diagnostic */}
          <div className="bg-white p-5 rounded-xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between gap-2">
              <h3 className="text-sm font-extrabold text-slate-800 flex items-center gap-2">
                <Database className="w-4 h-4 text-emerald-500" />
                Koneksi Database Firestore & Real-Time
              </h3>
              <button
                type="button"
                disabled={testingDb}
                onClick={async () => {
                  setTestingDb(true);
                  await testFirestoreConnection();
                  setRefreshTrigger((prev) => prev + 1);
                  setTestingDb(false);
                }}
                className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-[11px] rounded-lg inline-flex items-center gap-1 cursor-pointer disabled:opacity-50"
              >
                <RefreshCw className={`w-3 h-3 ${testingDb ? "animate-spin" : ""}`} />
                Uji Koneksi
              </button>
            </div>
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-xs space-y-2 text-slate-700">
              <div className="flex items-center justify-between gap-2">
                <span>Project ID:</span>
                <span className="font-mono text-[11px] font-bold text-slate-700">{FIREBASE_PROJECT_ID}</span>
              </div>
              <div className="flex items-center justify-between gap-2">
                <span>Database ID Aktif:</span>
                <span className="font-mono text-[10.5px] font-bold text-indigo-700 break-all text-right">
                  {FIRESTORE_DATABASE_ID}
                </span>
              </div>
              <div className="flex items-center justify-between gap-2">
                <span>Status Firestore (Cloud):</span>
                {firestoreHealth.writeQuotaExceeded ? (
                  <span className="font-bold text-amber-700 flex items-center gap-1">
                    <AlertTriangle className="w-3.5 h-3.5 text-amber-600" />
                    Baca Aktif · Kuota Tulis Harian Penuh
                  </span>
                ) : firestoreHealth.status === "connected" ? (
                  <span className="font-bold text-emerald-600 flex items-center gap-1">
                    <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                    Terhubung Penuh (Baca & Tulis)
                  </span>
                ) : firestoreHealth.status === "checking" ? (
                  <span className="font-bold text-slate-500">Memeriksa koneksi...</span>
                ) : (
                  <span className="font-bold text-rose-600">
                    {firestoreHealth.status === "offline" ? "Offline" : "Perlu Pemeriksaan"}
                  </span>
                )}
              </div>
              <div className="flex items-center justify-between gap-2">
                <span>Jalur Real-Time HP ↔ Laptop (SSE):</span>
                <span
                  className={`font-bold flex items-center gap-1 ${
                    sseStatus === "connected" ? "text-emerald-600" : "text-amber-600"
                  }`}
                >
                  <span
                    className={`w-2 h-2 rounded-full ${
                      sseStatus === "connected" ? "bg-emerald-500 animate-pulse" : "bg-amber-500"
                    }`}
                  ></span>
                  {sseStatus === "connected" ? "Terhubung Real-Time (<1 dtk)" : "Menyambung Ulang..."}
                </span>
              </div>
              {lastLiveEventAt && (
                <div className="flex items-center justify-between gap-2 pt-1 border-t border-slate-200/80">
                  <span>Sinkronisasi Terakhir:</span>
                  <span className="font-mono text-[11px] text-slate-500">{formatTimestamp(lastLiveEventAt)}</span>
                </div>
              )}
            </div>
            {firestoreHealth.writeQuotaExceeded && (
              <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg text-[11px] text-amber-900 space-y-1.5">
                <p className="font-bold">Info Kuota Harian Firestore (Free Tier):</p>
                <p>
                  Batas harian unit tulis gratis pada proyek Firebase sedang tercapai dan akan reset otomatis keesokan hari.
                  Aplikasi telah dioptimalkan untuk menghemat &gt;95% kuota tulis, dan sinkronisasi langsung antar HP dan Laptop tetap berjalan melalui Server Real-Time Relay.
                </p>
                <a
                  href={FIRESTORE_UPGRADE_URL}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1 font-bold text-amber-800 underline hover:text-amber-950"
                >
                  Lihat Database & Opsi Upgrade di Firebase Console
                  <ExternalLink className="w-3 h-3" />
                </a>
              </div>
            )}
          </div>

          {/* Tenant Isolation & Multi-Device Diagnostic */}
          <div className="bg-white p-5 rounded-xl border border-slate-200 space-y-4">
            <h3 className="text-sm font-extrabold text-slate-800 flex items-center gap-2">
              <Cpu className="w-4 h-4 text-indigo-500" />
              Sistem Multi-Tenancy & Multi-Device
            </h3>
            <div className="p-3 bg-indigo-50/50 border border-indigo-100 rounded-lg text-xs space-y-2 text-slate-700">
              <div className="flex items-center justify-between">
                <span>Sistem Isolasi:</span>
                <span className="font-bold text-indigo-600">Subkoleksi States Berbasis UID</span>
              </div>
              <div className="flex items-center justify-between">
                <span>Jumlah Tenant Terdaftar:</span>
                <span className="font-bold">{totalUsersCount} Tenant</span>
              </div>
              <div className="flex items-center justify-between">
                <span>Perangkat Terdeteksi (HP & Laptop):</span>
                <span className="font-bold text-teal-700">
                  {devices.length} Perangkat ({onlineDevicesCount} Online)
                </span>
              </div>
              <div className="flex items-center justify-between">
                <span>TKPI Rujukan Bersama:</span>
                <span className="font-bold text-emerald-600 flex items-center gap-1">
                  <CheckCircle className="w-3.5 h-3.5" />
                  Aktif (Delta Hemat Kuota)
                </span>
              </div>
            </div>
            <p className="text-[11px] text-slate-500 leading-relaxed">
              Setiap aktivitas pendaftaran, verifikasi kode email, dan login dari perangkat seluler (Android/iPhone) maupun komputer langsung dikirim melalui jalur sinkronisasi ganda (Firestore Cloud + Real-Time Server Stream) agar langsung terlihat di Panel Admin.
            </p>
          </div>
        </div>
      )}

    </div>
  );
}
