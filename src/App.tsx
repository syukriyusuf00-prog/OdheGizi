/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from "react";
import { SPPGProfile, SekolahPM, TigaBPM, MasterMenu, FoodCostDay, TKPIItem, KelompokSasaranPM, HariPM } from "./types";
import { 
  INITIAL_SPPG_PROFILE, 
  INITIAL_SEKOLAH_PM, 
  INITIAL_TIGAB_PM, 
  INITIAL_MASTER_MENU, 
  INITIAL_FOOD_COST_DAYS, 
  INITIAL_HARIAN_PM,
  EMPTY_SPPG_PROFILE,
  EMPTY_SEKOLAH_PM,
  EMPTY_TIGAB_PM,
  EMPTY_HARIAN_PM,
  EMPTY_MASTER_MENU,
  EMPTY_FOOD_COST_DAYS,
  EMPTY_SASARAN_LIST
} from "./initialData";
import { INITIAL_TKPI_DATABASE } from "./tkpiData";
import { getCountsForDay } from "./utils/calc";
import { composeMenuName, ensureMasterMenuPlans, MPASI_SASARAN_IDS, SEKOLAH_SASARAN_IDS } from "./utils/menuSchedule";

// Components
import SPPGProfileTab from "./components/SPPGProfileTab";
import PenerimaManfaatTab from "./components/PenerimaManfaatTab";
import MasterMenuTab from "./components/MasterMenuTab";
import FoodCostTab from "./components/FoodCostTab";
import CekGiziTab from "./components/CekGiziTab";
import DashboardOutputs from "./components/DashboardOutputs";
import NotaPesananLogistikTab from "./components/NotaPesananLogistikTab";
import RujukanJuknisTab from "./components/RujukanJuknisTab";
import HandbookTab from "./components/HandbookTab";
import ArsipTab from "./components/ArsipTab";
import AdminPanel from "./components/AdminPanel";

// Icons & Animation
import {
  Heart,
  Book,
  TrendingUp,
  Settings,
  Users,
  Utensils,
  BookOpen,
  DollarSign,
  Search,
  Database,
  Printer,
  ChevronRight,
  Sparkles,
  Info,
  Calendar,
  Layers,
  ArrowRight,
  Plus,
  Trash2,
  Upload,
  Download,
  FileText,
  Camera,
  Cloud,
  CloudOff,
  RefreshCw,
  Clock,
  Lock,
  LogOut,
  AlertCircle,
  CheckCircle2,
  Zap,
  Eye,
  EyeOff,
  Tags,
  RotateCcw,
  MessageCircle,
  Key,
  X,
  Archive,
  PlayCircle,
  Mail,
  ShieldCheck,
  ExternalLink,
  UserCheck,
  Building2,
  Copy,
  Check
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  useAuth,
  isMainAdminEmail,
  clearPerUserWorksheetStorage
} from "./context/AuthContext.tsx";
import {
  db,
  FIREBASE_AUTH_SETTINGS_URL,
  getPreferredAuthorizedAppUrls
} from "./lib/firebase.ts";
import { doc, getDoc, onSnapshot, setDoc } from "firebase/firestore";

const DEFAULT_FALLBACK_LOGO = "/assets/images/app_logo_sppg.png";
const GLOBAL_LOGO_SCHEMA_VERSION = 2;
const GLOBAL_LOGO_VERSION_KEY = "sisper_global_app_logo_version";

const normalizeAppLogoUrl = (url?: string | null): string => {
  if (!url || !url.trim() || url === "/src/assets/images/logo_sppg_1782256222616.jpg") {
    return DEFAULT_FALLBACK_LOGO;
  }
  return url;
};

const readCachedGlobalLeftLogo = (): string => {
  try {
    if (localStorage.getItem(GLOBAL_LOGO_VERSION_KEY) === String(GLOBAL_LOGO_SCHEMA_VERSION)) {
      return normalizeAppLogoUrl(localStorage.getItem("sisper_global_app_logo"));
    }
  } catch {}
  return DEFAULT_FALLBACK_LOGO;
};

export default function App() {
  // Load initial state from LocalStorage only when logged in; otherwise start completely empty
  const [profile, setProfile] = useState<SPPGProfile>(() => {
    const saved = localStorage.getItem("sisper_profile");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return EMPTY_SPPG_PROFILE;
  });

  const [sekolahPM, setSekolahPM] = useState<SekolahPM[]>(() => {
    const saved = localStorage.getItem("sisper_sekolah");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return EMPTY_SEKOLAH_PM;
  });

  const [tigaBPM, setTigaBPM] = useState<TigaBPM[]>(() => {
    const saved = localStorage.getItem("sisper_tigab");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return EMPTY_TIGAB_PM;
  });

  const [harianPM, setHarianPM] = useState<HariPM[]>(() => {
    const saved = localStorage.getItem("sisper_harian_pm");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error(e);
      }
    }
    return EMPTY_HARIAN_PM;
  });

  const [masterMenu, setMasterMenu] = useState<MasterMenu>(() => {
    const saved = localStorage.getItem("sisper_menu");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return EMPTY_MASTER_MENU;
  });

  const [foodCostDays, setFoodCostDays] = useState<FoodCostDay[]>(() => {
    const saved = localStorage.getItem("sisper_food_cost");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return EMPTY_FOOD_COST_DAYS;
  });

  const mergeWithTkpi2020 = (incoming: TKPIItem[]): TKPIItem[] => {
    if (!Array.isArray(incoming) || incoming.length === 0) return INITIAL_TKPI_DATABASE;
    const hasTkpi2020 = incoming.some((t) => typeof t.no === "number" && t.no >= 100);
    if (hasTkpi2020) return incoming;
    const refIds = new Set(INITIAL_TKPI_DATABASE.map((i) => i.id));
    const customOnly = incoming.filter((i) => !refIds.has(i.id) || i.isCustom === true);
    return [...INITIAL_TKPI_DATABASE, ...customOnly];
  };

  const extractCustomTkpiDelta = (list: TKPIItem[]): TKPIItem[] => {
    const refIds = new Set(INITIAL_TKPI_DATABASE.map((i) => i.id));
    return list.filter((i) => i.isCustom === true || !refIds.has(i.id));
  };

  const [tkpiList, setTkpiList] = useState<TKPIItem[]>(() => {
    if (localStorage.getItem("sisper_tkpi_cleared") === "true") {
      return [];
    }
    const saved = localStorage.getItem("sisper_tkpi_list");
    if (saved) {
      try {
        return mergeWithTkpi2020(JSON.parse(saved));
      } catch (e) {}
    }
    return INITIAL_TKPI_DATABASE;
  });

  const [tkpiNotice, setTkpiNotice] = useState<{ type: "success" | "info" | "error"; text: string } | null>(null);
  const showTkpiNotice = (type: "success" | "info" | "error", text: string) => {
    setTkpiNotice({ type, text });
    setTimeout(() => setTkpiNotice(null), 5000);
  };

  const [activeTab, setActiveTab] = useState<string>(() => {
    return localStorage.getItem("sisper_active_tab") || "dashboard";
  });

  useEffect(() => {
    if (activeTab) {
      localStorage.setItem("sisper_active_tab", activeTab);
    }
  }, [activeTab]);
  const [tkpiSearchQuery, setTkpiSearchQuery] = useState<string>("");

  // Global App Logo (Left) controlled by Admin / Panel Admin holders and automatically followed by Login screen & all users
  const [globalAppLogo, setGlobalAppLogo] = useState<string>(() => readCachedGlobalLeftLogo());
  const [globalRightLogo, setGlobalRightLogo] = useState<string>(() => {
    return localStorage.getItem("sisper_global_app_logo_right") || "";
  });

  const [customLogo, setCustomLogo] = useState<string>(() => readCachedGlobalLeftLogo());

  // State untuk penyesuaian/pemotongan logo profil
  const [isCropModalOpen, setIsCropModalOpen] = useState(false);
  const [cropImageSrc, setCropImageSrc] = useState<string>("");
  const [cropImageRatio, setCropImageRatio] = useState<number>(1);
  const [cropZoom, setCropZoom] = useState<number>(1);
  const [cropOffsetX, setCropOffsetX] = useState<number>(0);
  const [cropOffsetY, setCropOffsetY] = useState<number>(0);
  const [cropDragging, setCropDragging] = useState(false);
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });

  // Synchronized KOP Surat, Logos, Logo Crop & Paper Size
  const [kopLine1, setKopLine1] = useState(() => localStorage.getItem("kop_line1") || "BADAN GIZI NASIONAL");
  const [kopLine2, setKopLine2] = useState(() => localStorage.getItem("kop_line2") || "SATUAN PELAYANAN PEMENUHAN GIZI");
  const [kopLine3, setKopLine3] = useState(() => localStorage.getItem("kop_line3") || "");
  const [kopLine4, setKopLine4] = useState(() => localStorage.getItem("kop_line4") || "");
  const [kopLeftLogo, setKopLeftLogo] = useState(() => readCachedGlobalLeftLogo());
  const [kopRightLogo, setKopRightLogo] = useState(() => localStorage.getItem("sisper_global_app_logo_right") || localStorage.getItem("kop_right_logo") || "");
  
  const [leftLogoCrop, setLeftLogoCrop] = useState(() => {
    const saved = localStorage.getItem("kop_left_logo_crop");
    return saved ? JSON.parse(saved) : { top: 0, bottom: 0, left: 0, right: 0 };
  });

  const [rightLogoCrop, setRightLogoCrop] = useState(() => {
    const saved = localStorage.getItem("kop_right_logo_crop");
    return saved ? JSON.parse(saved) : { top: 0, bottom: 0, left: 0, right: 0 };
  });

  const [paperSize, setPaperSize] = useState<"A4" | "F4">(() => {
    return (localStorage.getItem("kop_paper_size") as "A4" | "F4") || "A4";
  });

  // Synchronized Sandbox Gizi (Cek Gizi) items (empty by default for new users)
  const [cekGiziItems, setCekGiziItems] = useState<any[]>(() => {
    const saved = localStorage.getItem("sisper_cek_gizi_items");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return [];
  });

  // Sistem Manajemen Multi-Periode & Riwayat Perencanaan
  const [savedPeriods, setSavedPeriods] = useState<any[]>(() => {
    const saved = localStorage.getItem("sisper_saved_periods");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return [];
  });

  // Pengaturan Target Harga & Klasifikasi Sasaran PM
  const [pmSettings, setPmSettings] = useState<any>(() => {
    const saved = localStorage.getItem("sisper_pm_settings");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return {
      porsiKecilHarga: 8000,
      porsiBesarHarga: 10000,
      porsiKecilSasaranIds: ["tk_paud_lb", "sd_kelas_1_3", "anak_balita", "anak_balita_13_59", "balita_6_11"],
      porsiBesarSasaranIds: ["sd_kelas_4_6", "smp_mts_smplb", "sma_smk_ma", "pendidik", "tenaga_kependidikan", "ibu_hamil", "ibu_menyusui"]
    };
  });

  const { 
    user, 
    userProfile,
    loading,
    isCloudActive, 
    saveStateToCloud, 
    loadStateFromCloud, 
    signInWithGoogle, 
    signOutUser,
    updateOwnProfileDetails,
    refreshUserProfile,
    authError,
    setAuthError,
    cloudSyncStatus,
    cloudSyncError,
    retryCloudRegistration,
    domainAuthInfo,
    recheckDomainAndSignIn,
    isDemo,
    demoSecondsLeft,
    demoNotice,
    clearDemoNotice,
    loginDemoAccount
  } = useAuth();

  const isDemoUser = Boolean(
    isDemo ||
    userProfile?.isDemo ||
    userProfile?.peran === "DEMO" ||
    user?.uid === "demo_user_5min"
  );

  // Strict non-admin guarantee: new tenants (peran: "USER") and demo accounts NEVER get Panel Admin access
  const isAdminUser =
    !isDemoUser &&
    Boolean(user) &&
    (userProfile?.statusPersetujuan === "aktif" || isMainAdminEmail(user?.email)) &&
    (isMainAdminEmail(user?.email) ||
      userProfile?.peran === "ADMIN" ||
      user?.uid === "admin_syukriyusuf82" ||
      user?.uid === "admin_sukriyusuf82");

  const getDefaultTabForProfesi = (profesi?: string, peran?: string) => {
    if (peran === "ADMIN") return "dashboard";
    const p = (profesi || "").toLowerCase();
    if (p.includes("gizi") || p.includes("nutrisi")) return "dashboard";
    if (p.includes("keuangan") || p.includes("akuntan") || p.includes("bendahara") || p.includes("finance")) return "notalogistik";
    if (p.includes("chef") || p.includes("masak") || p.includes("koki") || p.includes("logistik")) return "menu";
    if (p.includes("kepala") || p.includes("pimpinan") || p.includes("manager") || p.includes("direktur")) return "sppg";
    if (p.includes("kesehatan") || p.includes("pengawas") || p.includes("lapangan")) return "penerima";
    return "dashboard";
  };

  useEffect(() => {
    if (userProfile?.profesi && userProfile?.statusPersetujuan === "aktif") {
      const hasChosen = sessionStorage.getItem("sisper_user_tab_chosen");
      if (!hasChosen) {
        const roleTab = getDefaultTabForProfesi(userProfile.profesi, userProfile.peran);
        setActiveTab(roleTab);
      }
    }
  }, [userProfile?.profesi, userProfile?.peran, userProfile?.statusPersetujuan]);

  // Google Email Sign-In & Waiting-for-ACC UI states
  const [isGoogleSigningIn, setIsGoogleSigningIn] = useState(false);
  const [isRecheckingDomain, setIsRecheckingDomain] = useState(false);
  const [domainCopied, setDomainCopied] = useState(false);
  const [isCheckingStatus, setIsCheckingStatus] = useState(false);
  const [isRetryingCloudSync, setIsRetryingCloudSync] = useState(false);
  const [statusCheckFeedback, setStatusCheckFeedback] = useState<{ type: "info" | "success" | "error"; text: string } | null>(null);
  const [isSavingWaitingProfile, setIsSavingWaitingProfile] = useState(false);
  const [waitingProfileSavedMsg, setWaitingProfileSavedMsg] = useState<string | null>(null);

  // Tenant identity fields (optional quick-fill on Waiting ACC screen so Admin immediately sees SPPG details)
  const [regNama, setRegNama] = useState("");
  const [regProfesi, setRegProfesi] = useState("Ahli Gizi");
  const [regNamaSPPG, setRegNamaSPPG] = useState("");
  const [regNoHp, setRegNoHp] = useState("");

  useEffect(() => {
    if (!userProfile) return;
    setRegNama((prev) => prev || userProfile.namaLengkap || user?.displayName || "");
    setRegProfesi((prev) => (prev && prev !== "Ahli Gizi" ? prev : userProfile.profesi || "Ahli Gizi"));
    setRegNamaSPPG((prev) => prev || userProfile.namaSPPG || "");
    setRegNoHp((prev) => prev || userProfile.noHp || "");
  }, [userProfile?.uid, userProfile?.namaLengkap, userProfile?.profesi, userProfile?.namaSPPG, userProfile?.noHp, user?.displayName]);

  // Admin Panel states (never permitted for demo accounts or non-admin tenants)
  const [isAdminPanelOpen, setIsAdminPanelOpen] = useState(false);

  useEffect(() => {
    if ((!isAdminUser || isDemoUser) && isAdminPanelOpen) {
      setIsAdminPanelOpen(false);
    }
  }, [isAdminUser, isDemoUser, isAdminPanelOpen]);

  const [isCloudLoading, setIsCloudLoading] = useState(false);
  const [loadedCloudUid, setLoadedCloudUid] = useState<string | null>(null);
  const [cloudStatusMessage, setCloudStatusMessage] = useState<string>("Mode Lokal");

  const isMainAdminAccount =
    !isDemoUser && user
      ? isMainAdminEmail(user.email) ||
        user.uid === "admin_syukriyusuf82" ||
        user.uid === "admin_sukriyusuf82"
      : false;

  // Real-time Global Logo & TKPI 2020 subscriptions (always active across Login screen and all user accounts)
  useEffect(() => {
    const logoDocRef = doc(db, "configs", "app_logo");
    const unsubscribeLogo = onSnapshot(logoDocRef, (snap) => {
      const data = snap.exists() ? snap.data() : null;
      const hasCurrentSchema = Boolean(data && (data.schemaVersion ?? 0) >= GLOBAL_LOGO_SCHEMA_VERSION);
      const primaryLogo = hasCurrentSchema
        ? normalizeAppLogoUrl(data?.url)
        : DEFAULT_FALLBACK_LOGO;
      const rightLogoVal = data?.rightUrl ?? "";

      setGlobalAppLogo(primaryLogo);
      setGlobalRightLogo(rightLogoVal);
      setCustomLogo(primaryLogo);
      setKopLeftLogo(primaryLogo);
      setKopRightLogo(rightLogoVal);
      localStorage.setItem("sisper_global_app_logo", primaryLogo);
      localStorage.setItem("sisper_global_app_logo_right", rightLogoVal);
      localStorage.setItem(GLOBAL_LOGO_VERSION_KEY, String(GLOBAL_LOGO_SCHEMA_VERSION));
      localStorage.setItem("sisper_custom_logo", primaryLogo);
      localStorage.setItem("kop_left_logo", primaryLogo);
      localStorage.setItem("kop_right_logo", rightLogoVal);
    }, (error) => {
      console.warn("Gagal berlangganan logo global dari awan:", error);
    });

    const tkpiDocRef = doc(db, "global_settings", "tkpi");
    const unsubscribeTkpi = onSnapshot(tkpiDocRef, (snap) => {
      if (snap.exists() && localStorage.getItem("sisper_tkpi_cleared") !== "true") {
        const data = snap.data();
        if (data && data.data) {
          try {
            const parsed = JSON.parse(data.data);
            if (Array.isArray(parsed) && parsed.length > 0) {
              const upgraded = mergeWithTkpi2020(parsed);
              setTkpiList(upgraded);
              localStorage.setItem("sisper_tkpi_list", JSON.stringify(upgraded));
              localStorage.setItem("sisper_global_admin_tkpi", JSON.stringify(upgraded));
            }
          } catch (e) {
            console.warn("Gagal parsing TKPI global:", e);
          }
        }
      }
    }, (error) => {
      console.warn("Gagal berlangganan TKPI global dari awan:", error);
    });

    return () => {
      unsubscribeLogo();
      unsubscribeTkpi();
    };
  }, []);

  // Keep global logo reflected for all non-admin users if globalAppLogo changes
  useEffect(() => {
    if (globalAppLogo) {
      setCustomLogo(globalAppLogo);
      setKopLeftLogo(globalAppLogo);
    }
    if (globalRightLogo !== undefined) {
      setKopRightLogo(globalRightLogo);
    }
  }, [globalAppLogo, globalRightLogo]);

  // Tenant State Loader: strict empty isolation for every new user & demo separation
  useEffect(() => {
    let cancelled = false;

    if (user && isDemoUser) {
      setLoadedCloudUid(null);
      setIsCloudLoading(false);
      setCloudStatusMessage("Mode Demo 5 Menit (Lokal)");
      if (localStorage.getItem("sisper_demo_just_logged_in") === "true") {
        localStorage.removeItem("sisper_demo_just_logged_in");
        localStorage.setItem("sisper_loaded_uid", user.uid);
        setProfile(INITIAL_SPPG_PROFILE);
        setSekolahPM(INITIAL_SEKOLAH_PM);
        setTigaBPM(INITIAL_TIGAB_PM);
        setHarianPM(INITIAL_HARIAN_PM);
        setMasterMenu(INITIAL_MASTER_MENU);
        setFoodCostDays(INITIAL_FOOD_COST_DAYS);
        setTkpiList(INITIAL_TKPI_DATABASE);
        setKopLine1("BADAN GIZI NASIONAL");
        setKopLine2("SATUAN PELAYANAN PEMENUHAN GIZI");
        setKopLine3("SPPG DEMO PERENCANAAN GIZI");
        setKopLine4("Alamat : Lingkungan Simulasi Sistem SISPERMBG");
        setCekGiziItems([
          { id: "item_1", namaMenu: "Menu Uji Coba", tkpiId: "beras_giling", berat: 80, urt: "1 piring" },
          { id: "item_2", namaMenu: "Menu Uji Coba", tkpiId: "daging_ayam_tanpa_kulit", berat: 60, urt: "1 potong" },
          { id: "item_3", namaMenu: "Menu Uji Coba", tkpiId: "wortel_segar", berat: 30, urt: "1/2 gelas" }
        ]);
        setSavedPeriods([]);
        setActiveTab("dashboard");
      }
      return;
    }

    if (user) {
      const canAccessWorkspace = userProfile?.statusPersetujuan === "aktif" || isMainAdminAccount;
      if (!canAccessWorkspace) {
        setLoadedCloudUid(null);
        return;
      }

      const prevLoadedUid = localStorage.getItem("sisper_loaded_uid");
      if (prevLoadedUid && prevLoadedUid !== user.uid) {
        clearPerUserWorksheetStorage();
      }

      const loadAllCloudData = async () => {
        setLoadedCloudUid(null);
        setIsCloudLoading(true);
        setCloudStatusMessage("Sinkronisasi Awan...");
        try {
          const [
            cloudProfile,
            cloudSekolah,
            cloudTigab,
            cloudHarianPM,
            cloudMenu,
            cloudFoodCost,
            cloudTkpi,
            cloudKopLine1,
            cloudKopLine2,
            cloudKopLine3,
            cloudKopLine4,
            cloudCekGiziItems,
            cloudSavedPeriods,
            cloudPmSettings,
            globalLogoSnap
          ] = await Promise.all([
            loadStateFromCloud("profile"),
            loadStateFromCloud("sekolah"),
            loadStateFromCloud("tigab"),
            loadStateFromCloud("harianPM"),
            loadStateFromCloud("menu"),
            loadStateFromCloud("foodCost"),
            loadStateFromCloud("tkpi"),
            loadStateFromCloud("kopLine1"),
            loadStateFromCloud("kopLine2"),
            loadStateFromCloud("kopLine3"),
            loadStateFromCloud("kopLine4"),
            loadStateFromCloud("cekGiziItems"),
            loadStateFromCloud("savedPeriods"),
            loadStateFromCloud("pmSettings"),
            getDoc(doc(db, "configs", "app_logo")).catch(() => null)
          ]);

          if (cancelled) return;

          // Apply latest Global Admin Left Logo to all users & upgrade legacy logo records to v2
          const logoData = globalLogoSnap && globalLogoSnap.exists() ? globalLogoSnap.data() : null;
          const hasCurrentLogoSchema = Boolean(
            logoData && (logoData.schemaVersion ?? 0) >= GLOBAL_LOGO_SCHEMA_VERSION
          );
          const primaryLogo = hasCurrentLogoSchema
            ? normalizeAppLogoUrl(logoData?.url)
            : DEFAULT_FALLBACK_LOGO;
          const rightLogoVal = logoData?.rightUrl ?? "";

          setGlobalAppLogo(primaryLogo);
          setGlobalRightLogo(rightLogoVal);
          setCustomLogo(primaryLogo);
          setKopLeftLogo(primaryLogo);
          setKopRightLogo(rightLogoVal);
          localStorage.setItem("sisper_global_app_logo", primaryLogo);
          localStorage.setItem("sisper_global_app_logo_right", rightLogoVal);
          localStorage.setItem(GLOBAL_LOGO_VERSION_KEY, String(GLOBAL_LOGO_SCHEMA_VERSION));
          localStorage.setItem("sisper_custom_logo", primaryLogo);
          localStorage.setItem("kop_left_logo", primaryLogo);

          if (!hasCurrentLogoSchema && isAdminUser) {
            setDoc(
              doc(db, "configs", "app_logo"),
              {
                url: DEFAULT_FALLBACK_LOGO,
                schemaVersion: GLOBAL_LOGO_SCHEMA_VERSION,
                updatedBy: user.email || "Admin SISPERMBG",
                updatedAt: new Date().toISOString(),
              },
              { merge: true }
            ).catch((err) => {
              console.warn("Gagal menyinkronkan standar logo aplikasi global:", err);
            });
          }

          // New users (non-main-admin) always begin with completely empty worksheets until they fill them.
          // Do NOT blindly write 14 default documents on login — data is written to Firestore only when modified.
          if (cloudProfile) {
            setProfile(cloudProfile);
          } else {
            const initialProfile = isMainAdminAccount ? INITIAL_SPPG_PROFILE : EMPTY_SPPG_PROFILE;
            setProfile(initialProfile);
          }

          if (cloudSekolah) {
            setSekolahPM(cloudSekolah);
          } else {
            const initialSekolah = isMainAdminAccount ? INITIAL_SEKOLAH_PM : EMPTY_SEKOLAH_PM;
            setSekolahPM(initialSekolah);
          }

          if (cloudTigab) {
            setTigaBPM(cloudTigab);
          } else {
            const initialTigab = isMainAdminAccount ? INITIAL_TIGAB_PM : EMPTY_TIGAB_PM;
            setTigaBPM(initialTigab);
          }

          if (cloudHarianPM) {
            setHarianPM(cloudHarianPM);
          } else {
            const initialHarian = isMainAdminAccount ? INITIAL_HARIAN_PM : EMPTY_HARIAN_PM;
            setHarianPM(initialHarian);
          }

          if (cloudMenu) {
            setMasterMenu(cloudMenu);
          } else {
            const initialMenu = isMainAdminAccount ? INITIAL_MASTER_MENU : EMPTY_MASTER_MENU;
            setMasterMenu(initialMenu);
          }

          if (cloudFoodCost) {
            setFoodCostDays(cloudFoodCost);
          } else {
            const initialFoodCost = isMainAdminAccount ? INITIAL_FOOD_COST_DAYS : EMPTY_FOOD_COST_DAYS;
            setFoodCostDays(initialFoodCost);
          }

          if (cloudTkpi) {
            const resolvedTkpi =
              localStorage.getItem("sisper_tkpi_cleared") === "true"
                ? cloudTkpi
                : mergeWithTkpi2020(cloudTkpi);
            setTkpiList(resolvedTkpi);
          } else {
            let defaultTkpi: TKPIItem[] = INITIAL_TKPI_DATABASE;
            try {
              const globalSnap = await getDoc(doc(db, "global_settings", "tkpi"));
              if (globalSnap && globalSnap.exists()) {
                const raw = globalSnap.data().data;
                if (raw) {
                  defaultTkpi = mergeWithTkpi2020(JSON.parse(raw));
                }
              }
            } catch (e) {}

            setTkpiList(defaultTkpi);
          }

          if (cloudKopLine1) {
            setKopLine1(cloudKopLine1);
          } else {
            setKopLine1("BADAN GIZI NASIONAL");
          }

          if (cloudKopLine2) {
            setKopLine2(cloudKopLine2);
          } else {
            setKopLine2("SATUAN PELAYANAN PEMENUHAN GIZI");
          }

          if (cloudKopLine3 !== null && cloudKopLine3 !== undefined) {
            setKopLine3(cloudKopLine3);
          } else {
            const defaultKop3 = isMainAdminAccount ? "SPPG MUNA BARAT SAWERIGADI ONDOKE" : "";
            setKopLine3(defaultKop3);
          }

          if (cloudKopLine4 !== null && cloudKopLine4 !== undefined) {
            setKopLine4(cloudKopLine4);
          } else {
            const defaultKop4 = isMainAdminAccount ? "Alamat : Jln. Poros Lagadi-Tondasi, Desa Ondoke, Kec. Sawerigadi, Kab. Muna Barat" : "";
            setKopLine4(defaultKop4);
          }

          if (cloudCekGiziItems) {
            setCekGiziItems(cloudCekGiziItems);
          } else {
            const defaultCekGizi = isMainAdminAccount
              ? [
                  { id: "item_1", namaMenu: "Menu Uji Coba", tkpiId: "beras_giling", berat: 80, urt: "1 piring" },
                  { id: "item_2", namaMenu: "Menu Uji Coba", tkpiId: "daging_ayam_tanpa_kulit", berat: 60, urt: "1 potong" },
                  { id: "item_3", namaMenu: "Menu Uji Coba", tkpiId: "wortel_segar", berat: 30, urt: "1/2 gelas" }
                ]
              : [];
            setCekGiziItems(defaultCekGizi);
          }

          if (cloudSavedPeriods) {
            setSavedPeriods(cloudSavedPeriods);
          } else {
            setSavedPeriods([]);
          }

          if (cloudPmSettings) {
            setPmSettings(cloudPmSettings);
          } else {
            const defaultPmSettings = {
              porsiKecilHarga: 8000,
              porsiBesarHarga: 10000,
              porsiKecilSasaranIds: ["tk_paud_lb", "sd_kelas_1_3", "anak_balita", "anak_balita_13_59", "balita_6_11"],
              porsiBesarSasaranIds: ["sd_kelas_4_6", "smp_mts_smplb", "sma_smk_ma", "pendidik", "tenaga_kependidikan", "ibu_hamil", "ibu_menyusui"]
            };
            setPmSettings(defaultPmSettings);
          }

          if (!cancelled) {
            localStorage.setItem("sisper_loaded_uid", user.uid);
            setLoadedCloudUid(user.uid);
            setCloudStatusMessage("Awan Aktif");
          }
        } catch (error) {
          console.error("Gagal sinkronisasi awan:", error);
          setCloudStatusMessage("Gagal Sinkron");
        } finally {
          if (!cancelled) {
            setIsCloudLoading(false);
          }
        }
      };
      loadAllCloudData();
    } else {
      setLoadedCloudUid(null);
      setCloudStatusMessage("Mode Lokal");
      // Reset all worksheet states to completely empty on logout while keeping the Admin Global Logo for Login screen
      setProfile(EMPTY_SPPG_PROFILE);
      setSekolahPM(EMPTY_SEKOLAH_PM);
      setTigaBPM(EMPTY_TIGAB_PM);
      setHarianPM(EMPTY_HARIAN_PM);
      setMasterMenu(EMPTY_MASTER_MENU);
      setFoodCostDays(EMPTY_FOOD_COST_DAYS);
      setTkpiList(INITIAL_TKPI_DATABASE);
      setKopLine1("BADAN GIZI NASIONAL");
      setKopLine2("SATUAN PELAYANAN PEMENUHAN GIZI");
      setKopLine3("");
      setKopLine4("");
      setCustomLogo(globalAppLogo || DEFAULT_FALLBACK_LOGO);
      setKopLeftLogo(globalAppLogo || DEFAULT_FALLBACK_LOGO);
      setKopRightLogo(globalRightLogo || "");
      setCekGiziItems([]);
      setSavedPeriods([]);
      setPmSettings({
        porsiKecilHarga: 8000,
        porsiBesarHarga: 10000,
        porsiKecilSasaranIds: ["tk_paud_lb", "sd_kelas_1_3", "anak_balita", "anak_balita_13_59", "balita_6_11"],
        porsiBesarSasaranIds: ["sd_kelas_4_6", "smp_mts_smplb", "sma_smk_ma", "pendidik", "tenaga_kependidikan", "ibu_hamil", "ibu_menyusui"]
      });
      clearPerUserWorksheetStorage();
    }

    return () => {
      cancelled = true;
    };
  }, [user?.uid, isDemoUser, userProfile?.statusPersetujuan, isMainAdminAccount]);

  const canAutoSyncCloud = Boolean(
    isCloudActive &&
    !isCloudLoading &&
    user?.uid &&
    loadedCloudUid === user.uid
  );

  // Guard against firing 14 redundant Firestore writes on the exact render tick when cloud hydration completes
  const syncedBaselineRef = React.useRef<Record<string, string>>({});
  const lastBaselineUidRef = React.useRef<string | null>(null);

  useEffect(() => {
    if (!canAutoSyncCloud || !user?.uid) {
      if (!user?.uid) {
        syncedBaselineRef.current = {};
        lastBaselineUidRef.current = null;
      }
      return;
    }
    if (lastBaselineUidRef.current !== user.uid) {
      lastBaselineUidRef.current = user.uid;
      syncedBaselineRef.current = {
        profile: JSON.stringify(profile),
        sekolah: JSON.stringify(sekolahPM),
        tigab: JSON.stringify(tigaBPM),
        menu: JSON.stringify(masterMenu),
        foodCost: JSON.stringify(foodCostDays),
        harianPM: JSON.stringify(harianPM),
        tkpi: JSON.stringify(extractCustomTkpiDelta(tkpiList)),
        kopLine1: kopLine1,
        kopLine2: kopLine2,
        kopLine3: kopLine3,
        kopLine4: kopLine4,
        cekGiziItems: JSON.stringify(cekGiziItems),
        savedPeriods: JSON.stringify(savedPeriods),
        pmSettings: JSON.stringify(pmSettings),
      };
    }
  }, [canAutoSyncCloud, user?.uid]);

  const scheduleCloudSaveIfChanged = (key: string, value: any) => {
    if (!canAutoSyncCloud || !user?.uid || lastBaselineUidRef.current !== user.uid) return null;
    const serialized = typeof value === "string" ? value : JSON.stringify(value);
    if (syncedBaselineRef.current[key] === serialized) return null;
    return setTimeout(() => {
      syncedBaselineRef.current[key] = serialized;
      saveStateToCloud(key, value);
    }, 1500);
  };

  // Save states to local storage (only when logged in) and Cloud (only when actually edited after hydration)
  useEffect(() => {
    if (!user) return;
    localStorage.setItem("kop_left_logo_crop", JSON.stringify(leftLogoCrop));
    localStorage.setItem("kop_right_logo_crop", JSON.stringify(rightLogoCrop));
    localStorage.setItem("kop_paper_size", paperSize);
  }, [user, leftLogoCrop, rightLogoCrop, paperSize]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_profile", JSON.stringify(profile));
    const timer = scheduleCloudSaveIfChanged("profile", profile);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, profile, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_sekolah", JSON.stringify(sekolahPM));
    const timer = scheduleCloudSaveIfChanged("sekolah", sekolahPM);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, sekolahPM, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_tigab", JSON.stringify(tigaBPM));
    const timer = scheduleCloudSaveIfChanged("tigab", tigaBPM);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, tigaBPM, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_menu", JSON.stringify(masterMenu));
    const timer = scheduleCloudSaveIfChanged("menu", masterMenu);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, masterMenu, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_food_cost", JSON.stringify(foodCostDays));
    const timer = scheduleCloudSaveIfChanged("foodCost", foodCostDays);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, foodCostDays, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_harian_pm", JSON.stringify(harianPM));
    const timer = scheduleCloudSaveIfChanged("harianPM", harianPM);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, harianPM, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_tkpi_list", JSON.stringify(tkpiList));
    if (isAdminUser) {
      localStorage.setItem("sisper_global_admin_tkpi", JSON.stringify(tkpiList));
    }
    if (canAutoSyncCloud && lastBaselineUidRef.current === user.uid) {
      // Only persist custom/user-added TKPI rows (small delta) to Firestore — never re-upload 1,146 static TKPI 2020 rows
      const customDelta =
        localStorage.getItem("sisper_tkpi_cleared") === "true" && tkpiList.length === 0
          ? []
          : extractCustomTkpiDelta(tkpiList);
      const serializedDelta = JSON.stringify(customDelta);
      if (syncedBaselineRef.current.tkpi === serializedDelta) return;
      const timer = setTimeout(() => {
        syncedBaselineRef.current.tkpi = serializedDelta;
        saveStateToCloud("tkpi", customDelta);
        if (isAdminUser) {
          try {
            setDoc(
              doc(db, "global_settings", "tkpi"),
              {
                data: serializedDelta,
                updatedAt: new Date().toISOString(),
              },
              { merge: true }
            ).catch(() => {});
          } catch (e) {}
        }
      }, 1500);
      return () => clearTimeout(timer);
    }
  }, [user, tkpiList, canAutoSyncCloud, isAdminUser]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("kop_line1", kopLine1);
    const timer = scheduleCloudSaveIfChanged("kopLine1", kopLine1);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, kopLine1, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("kop_line2", kopLine2);
    const timer = scheduleCloudSaveIfChanged("kopLine2", kopLine2);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, kopLine2, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("kop_line3", kopLine3);
    const timer = scheduleCloudSaveIfChanged("kopLine3", kopLine3);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, kopLine3, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("kop_line4", kopLine4);
    const timer = scheduleCloudSaveIfChanged("kopLine4", kopLine4);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, kopLine4, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_cek_gizi_items", JSON.stringify(cekGiziItems));
    const timer = scheduleCloudSaveIfChanged("cekGiziItems", cekGiziItems);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, cekGiziItems, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_saved_periods", JSON.stringify(savedPeriods));
    const timer = scheduleCloudSaveIfChanged("savedPeriods", savedPeriods);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, savedPeriods, canAutoSyncCloud]);

  useEffect(() => {
    if (!user) return;
    localStorage.setItem("sisper_pm_settings", JSON.stringify(pmSettings));
    const timer = scheduleCloudSaveIfChanged("pmSettings", pmSettings);
    return () => { if (timer) clearTimeout(timer); };
  }, [user, pmSettings, canAutoSyncCloud]);

  // Recipient totals (using Day 1 as default baseline display)
  const defaultDayCounts = getCountsForDay(harianPM, 1);
  const totalSekolahSiswa = defaultDayCounts.pmKecilSekolah + defaultDayCounts.pmBesarSekolah;
  const total3BOrang = defaultDayCounts.totalBalita + defaultDayCounts.totalBumil + defaultDayCounts.totalBusui;
  const grandTotalRecipients = totalSekolahSiswa + total3BOrang;

  // State for adding custom TKPI item manually (TKPI 2020 Table Format)
  const [newTkpiItem, setNewTkpiItem] = useState<Partial<TKPIItem>>({
    id: "",
    no: undefined,
    nama: "",
    sumber: "TKPI 2020 / Kustom",
    kategori: "Serealia & Olahannya",
    beratStandar: 100,
    bdd: 100,
    energi: 0,
    protein: 0,
    lemak: 0,
    kh: 0,
    serat: 0,
    abu: 0,
    ca: 0,
    p: 0,
    fe: 0,
    na: 0,
    k: 0,
    cu: 0,
    zn: 0,
    retinol: 0,
    b_karoten: 0,
    kartotal: 0,
    thiamin: 0,
    riboflavin: 0,
    niasin: 0,
    vit_c: 0,
    air: 0
  });

  const handleAddTkpiSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTkpiItem.nama || !newTkpiItem.nama.trim()) {
      showTkpiNotice("error", "Mohon masukkan NAMA BAHAN terlebih dahulu.");
      return;
    }
    localStorage.removeItem("sisper_tkpi_cleared");
    const nextNo = newTkpiItem.no ? Number(newTkpiItem.no) : tkpiList.reduce((max, item) => Math.max(max, item.no || 0), 0) + 1;
    const slug = (newTkpiItem.nama || "").toLowerCase().replace(/[^a-z0-9]/g, "_") + "_" + Date.now();
    const finalItem: TKPIItem = {
      id: slug,
      no: nextNo,
      nama: newTkpiItem.nama.trim(),
      sumber: newTkpiItem.sumber || "Kustom",
      kategori: newTkpiItem.kategori || "Makanan & Lainnya",
      beratStandar: Number(newTkpiItem.beratStandar) || 100,
      bdd: Number(newTkpiItem.bdd) || 100,
      energi: Number(newTkpiItem.energi) || 0,
      protein: Number(newTkpiItem.protein) || 0,
      lemak: Number(newTkpiItem.lemak) || 0,
      kh: Number(newTkpiItem.kh) || 0,
      serat: Number(newTkpiItem.serat) || 0,
      abu: Number(newTkpiItem.abu) || 0,
      ca: Number(newTkpiItem.ca) || 0,
      p: Number(newTkpiItem.p) || 0,
      fe: Number(newTkpiItem.fe) || 0,
      na: Number(newTkpiItem.na) || 0,
      k: Number(newTkpiItem.k) || 0,
      cu: Number(newTkpiItem.cu) || 0,
      zn: Number(newTkpiItem.zn) || 0,
      retinol: Number(newTkpiItem.retinol) || 0,
      b_karoten: Number(newTkpiItem.b_karoten) || 0,
      kartotal: Number(newTkpiItem.kartotal) || 0,
      thiamin: Number(newTkpiItem.thiamin) || 0,
      riboflavin: Number(newTkpiItem.riboflavin) || 0,
      niasin: Number(newTkpiItem.niasin) || 0,
      vit_c: Number(newTkpiItem.vit_c) || 0,
      air: Number(newTkpiItem.air) || 0,
      isCustom: true
    };

    setTkpiList((prev) => [finalItem, ...prev]);
    showTkpiNotice("success", `Bahan "${finalItem.nama}" berhasil ditambahkan sesuai format tabel TKPI 2020.`);
    
    // Reset form
    setNewTkpiItem({
      id: "",
      no: undefined,
      nama: "",
      sumber: "TKPI 2020 / Kustom",
      kategori: finalItem.kategori,
      beratStandar: 100,
      bdd: 100,
      energi: 0,
      protein: 0,
      lemak: 0,
      kh: 0,
      serat: 0,
      abu: 0,
      ca: 0,
      p: 0,
      fe: 0,
      na: 0,
      k: 0,
      cu: 0,
      zn: 0,
      retinol: 0,
      b_karoten: 0,
      kartotal: 0,
      thiamin: 0,
      riboflavin: 0,
      niasin: 0,
      vit_c: 0,
      air: 0
    });
  };

  const handleDownloadTemplate = () => {
    // Official TKPI 2020 spreadsheet column format
    const headers = [
      "NO",
      "NAMA BAHAN",
      "SUMBER",
      "BERAT STANDAR",
      "% BDD",
      "ENERGI",
      "PROTEIN",
      "LEMAK",
      "KH",
      "SERAT",
      "ABU",
      "KALSIUM",
      "FOSFOR",
      "BESI",
      "NATRIUM",
      "KALIUM",
      "TEMBAGA",
      "SENG",
      "RETINOL",
      "BKAR",
      "KARTOTAL",
      "THIAMIN",
      "RIBOFLAVIN",
      "NIASIN",
      "VIT_C",
      "AIR",
      "KATEGORI"
    ];

    const sampleRow = {
      "NO": 1191,
      "NAMA BAHAN": "Susu Kambing Segar",
      "SUMBER": "TKPI 2020 / Lokal",
      "BERAT STANDAR": 100,
      "% BDD": 100,
      "ENERGI": 64,
      "PROTEIN": 4.3,
      "LEMAK": 2.3,
      "KH": 6.6,
      "SERAT": 0,
      "ABU": 0.8,
      "KALSIUM": 98,
      "FOSFOR": 78,
      "BESI": 2.7,
      "NATRIUM": 35,
      "KALIUM": 145,
      "TEMBAGA": 0.05,
      "SENG": 0.3,
      "RETINOL": 37,
      "BKAR": 0,
      "KARTOTAL": 0,
      "THIAMIN": 0.05,
      "RIBOFLAVIN": 0.14,
      "NIASIN": 0.2,
      "VIT_C": 1,
      "AIR": 85.9,
      "KATEGORI": "Susu & Olahannya"
    };

    const ws = XLSX.utils.json_to_sheet([sampleRow], { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template_TKPI_2020");
    XLSX.writeFile(wb, "Template_Database_TKPI_2020.xlsx");
  };

  const resizeAndCompressImage = (file: File, maxWidth = 300, maxHeight = 300, quality = 0.75): Promise<string> => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (readerEvent) => {
        const image = new Image();
        image.onload = () => {
          const canvas = document.createElement("canvas");
          let width = image.width;
          let height = image.height;

          if (width > height) {
            if (width > maxWidth) {
              height = Math.round((height * maxWidth) / width);
              width = maxWidth;
            }
          } else {
            if (height > maxHeight) {
              width = Math.round((width * maxHeight) / height);
              height = maxHeight;
            }
          }

          canvas.width = width;
          canvas.height = height;

          const ctx = canvas.getContext("2d");
          if (ctx) {
            ctx.drawImage(image, 0, 0, width, height);
            const mimeType = file.type === "image/png" ? "image/png" : "image/jpeg";
            const dataUrl = canvas.toDataURL(mimeType, mimeType === "image/jpeg" ? quality : undefined);
            resolve(dataUrl);
          } else {
            resolve(readerEvent.target?.result as string);
          }
        };
        image.onerror = () => {
          resolve(readerEvent.target?.result as string);
        };
        image.src = readerEvent.target?.result as string;
      };
      reader.onerror = () => {
        resolve("");
      };
      reader.readAsDataURL(file);
    });
  };

  const handleLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!isAdminUser) {
      alert("Hanya Administrator atau pemilik akses Panel Admin yang memiliki wewenang untuk mengubah logo aplikasi!");
      return;
    }

    if (file.size > 10 * 1024 * 1024) {
      alert("Ukuran berkas terlalu besar. Silakan pilih berkas di bawah 10 MB.");
      return;
    }

    try {
      const reader = new FileReader();
      reader.onload = (readerEvent) => {
        const img = new Image();
        img.onload = () => {
          setCropImageRatio(img.width / img.height);
          setCropImageSrc(readerEvent.target?.result as string);
          setCropZoom(1.0);
          setCropOffsetX(0);
          setCropOffsetY(0);
          setIsCropModalOpen(true);
        };
        img.src = readerEvent.target?.result as string;
      };
      reader.readAsDataURL(file);
    } catch (err) {
      console.error("Gagal memuat gambar untuk dipotong:", err);
      alert("Gagal memuat gambar. Silakan coba gambar lain.");
    }
  };

  const handleCropMouseDown = (e: React.MouseEvent) => {
    setCropDragging(true);
    setDragStart({ x: e.clientX - cropOffsetX, y: e.clientY - cropOffsetY });
  };

  const handleCropMouseMove = (e: React.MouseEvent) => {
    if (!cropDragging) return;
    setCropOffsetX(e.clientX - dragStart.x);
    setCropOffsetY(e.clientY - dragStart.y);
  };

  const handleCropMouseUp = () => setCropDragging(false);
  const handleCropMouseLeave = () => setCropDragging(false);

  const handleCropTouchStart = (e: React.TouchEvent) => {
    if (e.touches.length === 1) {
      setCropDragging(true);
      setDragStart({
        x: e.touches[0].clientX - cropOffsetX,
        y: e.touches[0].clientY - cropOffsetY,
      });
    }
  };

  const handleCropTouchMove = (e: React.TouchEvent) => {
    if (!cropDragging) return;
    if (e.touches.length === 1) {
      setCropOffsetX(e.touches[0].clientX - dragStart.x);
      setCropOffsetY(e.touches[0].clientY - dragStart.y);
    }
  };

  const handleCropTouchEnd = () => setCropDragging(false);

  const handleSaveCroppedLogo = () => {
    if (!cropImageSrc) return;
    const img = new Image();
    img.onload = async () => {
      const canvas = document.createElement("canvas");
      canvas.width = 300;
      canvas.height = 300;
      const ctx = canvas.getContext("2d");
      if (ctx) {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, 300, 300);

        const containerSize = 240;
        const targetSize = 300;
        const scaleFactor = targetSize / containerSize;

        let baseWidth = containerSize;
        let baseHeight = containerSize;
        if (cropImageRatio > 1) {
          baseWidth = containerSize * cropImageRatio;
        } else {
          baseHeight = containerSize / cropImageRatio;
        }

        const wDraw = baseWidth * cropZoom * scaleFactor;
        const hDraw = baseHeight * cropZoom * scaleFactor;

        const cx = 150 + (cropOffsetX * scaleFactor);
        const cy = 150 + (cropOffsetY * scaleFactor);

        const xDraw = cx - wDraw / 2;
        const yDraw = cy - hDraw / 2;

        ctx.drawImage(img, xDraw, yDraw, wDraw, hDraw);

        const compressedBase64 = canvas.toDataURL("image/jpeg", 0.85);
        if (compressedBase64 && isAdminUser) {
          setGlobalAppLogo(compressedBase64);
          setCustomLogo(compressedBase64);
          setKopLeftLogo(compressedBase64);
          localStorage.setItem("sisper_global_app_logo", compressedBase64);
          localStorage.setItem(GLOBAL_LOGO_VERSION_KEY, String(GLOBAL_LOGO_SCHEMA_VERSION));
          localStorage.setItem("sisper_custom_logo", compressedBase64);
          localStorage.setItem("kop_left_logo", compressedBase64);

          try {
            // Write left/main app logo directly to configs/app_logo for automatic sync to Login screen & all users
            await setDoc(doc(db, "configs", "app_logo"), {
              url: compressedBase64,
              schemaVersion: GLOBAL_LOGO_SCHEMA_VERSION,
              updatedBy: user?.email || "Admin SISPERMBG",
              updatedAt: new Date().toISOString()
            }, { merge: true });
          } catch (err) {
            console.warn("Gagal memperbarui logo global di Firestore:", err);
          }
        }
      }
      setIsCropModalOpen(false);
    };
    img.src = cropImageSrc;
  };

  const handleUpdateLeftLogo = async (newLogo: string) => {
    if (!isAdminUser) return;
    const normalized = normalizeAppLogoUrl(newLogo);
    setGlobalAppLogo(normalized);
    setKopLeftLogo(normalized);
    setCustomLogo(normalized);
    localStorage.setItem("sisper_global_app_logo", normalized);
    localStorage.setItem(GLOBAL_LOGO_VERSION_KEY, String(GLOBAL_LOGO_SCHEMA_VERSION));
    localStorage.setItem("kop_left_logo", normalized);
    localStorage.setItem("sisper_custom_logo", normalized);

    try {
      await setDoc(doc(db, "configs", "app_logo"), {
        url: normalized,
        schemaVersion: GLOBAL_LOGO_SCHEMA_VERSION,
        updatedBy: user?.email || "Admin SISPERMBG",
        updatedAt: new Date().toISOString()
      }, { merge: true });
    } catch (err) {
      console.warn("Gagal menyinkronkan logo global ke cloud:", err);
    }
  };

  const handleUpdateRightLogo = async (newLogo: string) => {
    if (!isAdminUser) return;
    setGlobalRightLogo(newLogo);
    setKopRightLogo(newLogo);
    localStorage.setItem("sisper_global_app_logo_right", newLogo);
    localStorage.setItem("kop_right_logo", newLogo);

    try {
      await setDoc(doc(db, "configs", "app_logo"), {
        rightUrl: newLogo,
        updatedBy: user?.email || "Admin SISPERMBG",
        updatedAt: new Date().toISOString()
      }, { merge: true });
    } catch (err) {
      console.warn("Gagal menyinkronkan logo kanan global ke cloud:", err);
    }
  };

  const buildOdheGiziWhatsAppUrl = (opts?: {
    nama?: string;
    email?: string;
    namaSPPG?: string;
    keperluan?: string;
  }) => {
    const nama = (opts?.nama || userProfile?.namaLengkap || user?.displayName || "").trim();
    const email = (opts?.email || user?.email || userProfile?.email || "").trim();
    const sppg = (opts?.namaSPPG || regNamaSPPG || userProfile?.namaSPPG || "").trim();
    const keperluan = opts?.keperluan || "mohon ACC Akses dashboard SISPERMBG saya";
    const lines = [
      "Halo Odhe Gizi,",
      nama ? `Saya ${nama}` : "Saya pengguna SISPERMBG",
      email ? `Email Akun: ${email}` : "",
      sppg ? `Instansi / SPPG: ${sppg}` : "",
      `Izin ${keperluan}. Terima kasih.`
    ].filter(Boolean);
    return `https://wa.me/6282271095251?text=${encodeURIComponent(lines.join("\n"))}`;
  };

  const handleTkpiUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const parseTkpiNum = (v: any, fallback = 0): number => {
      if (v === undefined || v === null || v === "") return fallback;
      if (typeof v === "number") return isNaN(v) ? fallback : v;
      const s = String(v).replace(/,/g, "").replace(/\s+/g, "").trim();
      if (!s || /^#+$/.test(s) || s === "-") return fallback;
      const n = parseFloat(s);
      return isNaN(n) ? fallback : n;
    };

    const pickCol = (row: Record<string, any>, candidates: string[]): any => {
      for (const c of candidates) {
        if (row[c] !== undefined && row[c] !== "") return row[c];
      }
      const lowerMap: Record<string, any> = {};
      Object.keys(row).forEach((k) => {
        lowerMap[k.trim().toLowerCase()] = row[k];
      });
      for (const c of candidates) {
        const hit = lowerMap[c.trim().toLowerCase()];
        if (hit !== undefined && hit !== "") return hit;
      }
      return undefined;
    };

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rawData = XLSX.utils.sheet_to_json<any>(ws);

        if (rawData && rawData.length > 0) {
          localStorage.removeItem("sisper_tkpi_cleared");
          const parsed: TKPIItem[] = rawData
            .map((row: any, idx: number) => {
              const rawNama = pickCol(row, ["NAMA BAHAN", "Nama Bahan", "nama", "NamaBahan"]);
              if (!rawNama || String(rawNama).trim() === "") return null;
              const nama = String(rawNama).replace(/\s+/g, " ").trim();
              const sumber = String(pickCol(row, ["SUMBER", "Sumber", "sumber"]) || "TKPI 2020").trim();
              const kategori = String(pickCol(row, ["KATEGORI", "Kategori", "kategori"]) || "Makanan & Lainnya").trim();
              const noVal = parseTkpiNum(pickCol(row, ["NO", "No", "no"]), tkpiList.length + idx + 1);
              return {
                id: `tkpi_import_${idx}_${Date.now()}`,
                no: noVal,
                nama,
                sumber,
                kategori,
                beratStandar: parseTkpiNum(pickCol(row, ["BERAT STANDAR", "Berat Standar", "berat_standar"]), 100),
                bdd: parseTkpiNum(pickCol(row, ["% BDD", "BDD", "bdd"]), 100),
                energi: parseTkpiNum(pickCol(row, ["ENERGI", "Energi", "energi"]), 0),
                protein: parseTkpiNum(pickCol(row, ["PROTEIN", "Protein", "protein"]), 0),
                lemak: parseTkpiNum(pickCol(row, ["LEMAK", "Lemak", "lemak"]), 0),
                kh: parseTkpiNum(pickCol(row, ["KH", "Karbohidrat", "kh"]), 0),
                serat: parseTkpiNum(pickCol(row, ["SERAT", "Serat", "serat"]), 0),
                abu: parseTkpiNum(pickCol(row, ["ABU", "Abu", "abu"]), 0),
                ca: parseTkpiNum(pickCol(row, ["KALSIUM", "Kalsium", "Ca", "ca"]), 0),
                p: parseTkpiNum(pickCol(row, ["FOSFOR", "Fosfor", "P", "p"]), 0),
                fe: parseTkpiNum(pickCol(row, ["BESI", "Zat Besi", "Fe", "fe"]), 0),
                na: parseTkpiNum(pickCol(row, ["NATRIUM", "Natrium", "Na", "na"]), 0),
                k: parseTkpiNum(pickCol(row, ["KALIUM", "Kalium", "K", "k"]), 0),
                cu: parseTkpiNum(pickCol(row, ["TEMBAGA", "Tembaga", "Cu", "cu"]), 0),
                zn: parseTkpiNum(pickCol(row, ["SENG", "Seng", "Zn", "zn"]), 0),
                retinol: parseTkpiNum(pickCol(row, ["RETINOL", "Retinol", "retinol"]), 0),
                b_karoten: parseTkpiNum(pickCol(row, ["BKAR", "Beta Karoten", "b_karoten"]), 0),
                kartotal: parseTkpiNum(pickCol(row, ["KARTOTAL", "Kartotal", "kartotal"]), 0),
                thiamin: parseTkpiNum(pickCol(row, ["THIAMIN", "Thiamin", "thiamin"]), 0),
                riboflavin: parseTkpiNum(pickCol(row, ["RIBOFLAVIN", "Riboflavin", "riboflavin"]), 0),
                niasin: parseTkpiNum(pickCol(row, ["NIASIN", "Niasin", "niasin"]), 0),
                vit_c: parseTkpiNum(pickCol(row, ["VIT_C", "Vit C", "vit_c"]), 0),
                air: parseTkpiNum(pickCol(row, ["AIR", "Air", "air"]), 0),
                isCustom: true
              } as TKPIItem;
            })
            .filter((x): x is TKPIItem => x !== null);

          if (parsed.length > 0) {
            setTkpiList((prev) => [...parsed, ...prev]);
            showTkpiNotice("success", `Berhasil mengimpor ${parsed.length} bahan makanan sesuai format TKPI 2020.`);
          } else {
            showTkpiNotice("error", "Tidak ditemukan baris valid dengan kolom NAMA BAHAN / Nama Bahan.");
          }
        }
      } catch (err) {
        showTkpiNotice("error", "Gagal memproses berkas Excel/CSV. Pastikan menggunakan format kolom TKPI 2020.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const [selectedCategoryFilter, setSelectedCategoryFilter] = useState<string>("all");

  const deleteTkpiItem = (id: string) => {
    const item = tkpiList.find((t) => t.id === id);
    setTkpiList((prev) => prev.filter((t) => t.id !== id));
    if (item) {
      showTkpiNotice("info", `"${item.nama}" dihapus dari daftar TKPI aktif.`);
    }
  };

  const deleteCategory = (categoryName: string) => {
    const itemsToDelete = tkpiList.filter(
      (t) => (t.kategori || t.sumber || "Umum").trim() === categoryName
    );
    if (itemsToDelete.length === 0) return;
    setTkpiList((prev) =>
      prev.filter((t) => (t.kategori || t.sumber || "Umum").trim() !== categoryName)
    );
    if (selectedCategoryFilter === categoryName) {
      setSelectedCategoryFilter("all");
    }
    showTkpiNotice("info", `Kategori "${categoryName}" (${itemsToDelete.length} bahan) dihapus.`);
  };

  const clearAllTkpi = () => {
    localStorage.setItem("sisper_tkpi_cleared", "true");
    setTkpiList([]);
    setSelectedCategoryFilter("all");
    showTkpiNotice("info", "Database TKPI dikosongkan. Klik 'Sinkronkan TKPI 2020' kapan saja untuk memulihkan 1.189 bahan.");
  };

  const restoreDefaultTkpi = () => {
    localStorage.removeItem("sisper_tkpi_cleared");
    const customItems = tkpiList.filter((i) => i.isCustom === true);
    const refIds = new Set(INITIAL_TKPI_DATABASE.map((i) => i.id));
    const uniqueCustom = customItems.filter((c) => !refIds.has(c.id));
    setTkpiList([...INITIAL_TKPI_DATABASE, ...uniqueCustom]);
    setSelectedCategoryFilter("all");
    showTkpiNotice(
      "success",
      `Database TKPI 2020 berhasil disinkronkan (${INITIAL_TKPI_DATABASE.length} bahan rujukan${uniqueCustom.length > 0 ? ` + ${uniqueCustom.length} bahan kustom` : ""}).`
    );
  };

  const [tkpiPage, setTkpiPage] = useState<number>(1);
  const [tkpiItemsPerPage, setTkpiItemsPerPage] = useState<number>(50);

  const { availableCategories, categoryCountsMap } = React.useMemo(() => {
    const counts: Record<string, number> = {};
    tkpiList.forEach((item) => {
      const cat = (item.kategori || item.sumber || "Umum").trim();
      if (cat) {
        counts[cat] = (counts[cat] || 0) + 1;
      }
    });
    return {
      availableCategories: Object.keys(counts).sort(),
      categoryCountsMap: counts
    };
  }, [tkpiList]);

  const filteredTkpiList = React.useMemo(() => {
    const q = tkpiSearchQuery.toLowerCase().trim();
    return tkpiList.filter((item) => {
      const cat = (item.kategori || item.sumber || "Umum").trim();
      if (selectedCategoryFilter !== "all" && cat !== selectedCategoryFilter) {
        return false;
      }
      if (!q) return true;
      const namaMatch = (item.nama || "").toLowerCase().includes(q);
      const sumberMatch = (item.sumber || "").toLowerCase().includes(q);
      const kategoriMatch = (item.kategori || "").toLowerCase().includes(q);
      return namaMatch || sumberMatch || kategoriMatch;
    });
  }, [tkpiList, selectedCategoryFilter, tkpiSearchQuery]);

  const totalTkpiPages = Math.ceil(filteredTkpiList.length / tkpiItemsPerPage) || 1;

  const paginatedTkpiList = React.useMemo(() => {
    const start = (tkpiPage - 1) * tkpiItemsPerPage;
    return filteredTkpiList.slice(start, start + tkpiItemsPerPage);
  }, [filteredTkpiList, tkpiPage, tkpiItemsPerPage]);

  // Reset page to 1 when filter or items per page change
  React.useEffect(() => {
    setTkpiPage(1);
  }, [selectedCategoryFilter, tkpiSearchQuery, tkpiItemsPerPage]);

  if (loading) {
    return (
      <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col justify-center items-center p-6 relative font-sans">
        <div className="animate-spin rounded-full h-10 w-10 border-b-2 border-emerald-400"></div>
        <p className="mt-4 text-xs font-semibold text-slate-400">Memuat Sistem GiziSync...</p>
      </div>
    );
  }

  const handleGoogleSignInClick = async () => {
    setAuthError(null);
    setStatusCheckFeedback(null);
    setIsGoogleSigningIn(true);
    try {
      await signInWithGoogle();
    } finally {
      setIsGoogleSigningIn(false);
    }
  };

  const handleCopyCurrentDomain = async (domainText: string) => {
    if (!domainText) return;
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(domainText);
      } else {
        const ta = document.createElement("textarea");
        ta.value = domainText;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
      }
      setDomainCopied(true);
      setTimeout(() => setDomainCopied(false), 2500);
    } catch {
      setDomainCopied(false);
    }
  };

  const handleRecheckDomainClick = async () => {
    setIsRecheckingDomain(true);
    try {
      await recheckDomainAndSignIn();
    } finally {
      setIsRecheckingDomain(false);
    }
  };

  const handleManualStatusCheck = async () => {
    setIsCheckingStatus(true);
    setStatusCheckFeedback(null);
    try {
      const updated = await refreshUserProfile();
      if (updated?.statusPersetujuan === "aktif") {
        setStatusCheckFeedback({
          type: "success",
          text: "Akses Anda telah disetujui (ACC)! Membuka dashboard SISPERMBG..."
        });
      } else if (updated?.statusPersetujuan === "diblokir") {
        setStatusCheckFeedback({
          type: "error",
          text: "Status akun Anda saat ini sedang dinonaktifkan oleh Admin. Silakan hubungi Odhe Gizi."
        });
      } else {
        setStatusCheckFeedback({
          type: "info",
          text: "Status akun masih Menunggu ACC. Silakan hubungi Odhe Gizi via WhatsApp untuk persetujuan cepat."
        });
      }
    } finally {
      setIsCheckingStatus(false);
    }
  };

  const handleRetryCloudRegistration = async () => {
    setIsRetryingCloudSync(true);
    setStatusCheckFeedback(null);
    try {
      const ok = await retryCloudRegistration();
      if (ok) {
        setStatusCheckFeedback({
          type: "success",
          text: `Email ${user?.email || ""} berhasil dikirim dan tercatat otomatis di Panel Admin SISPERMBG.`,
        });
      } else {
        setStatusCheckFeedback({
          type: "error",
          text: "Pengiriman ke Panel Admin belum berhasil. Pastikan koneksi internet stabil lalu coba lagi.",
        });
      }
    } finally {
      setIsRetryingCloudSync(false);
    }
  };

  const handleSaveWaitingIdentity = async (e: React.FormEvent) => {
    e.preventDefault();
    setWaitingProfileSavedMsg(null);
    setIsSavingWaitingProfile(true);
    try {
      await updateOwnProfileDetails({
        namaLengkap: regNama.trim() || user?.displayName || user?.email?.split("@")[0] || "Pengguna SISPERMBG",
        profesi: regProfesi.trim() || "Ahli Gizi",
        namaSPPG: regNamaSPPG.trim(),
        noHp: regNoHp.trim(),
      });
      setWaitingProfileSavedMsg("Data identitas SPPG berhasil dikirim ke Panel Admin Odhe Gizi.");
    } catch (err: any) {
      setWaitingProfileSavedMsg(err?.message || "Gagal menyimpan identitas.");
    } finally {
      setIsSavingWaitingProfile(false);
    }
  };

  // If user is not logged in, show clean Email Google Login screen (no password / registration tabs)
  if (!user) {
    const isIframeBlocked = authError === "IFRAME_BLOCKED";
    const isUnauthorizedDomain =
      authError === "UNAUTHORIZED_DOMAIN" ||
      (domainAuthInfo.checked && !domainAuthInfo.isAuthorized);
    const currentDomainHost =
      domainAuthInfo.currentHostname ||
      (typeof window !== "undefined" ? window.location.hostname : "");
    const preferredAuthorizedUrls = getPreferredAuthorizedAppUrls(
      domainAuthInfo.authorizedDomains || []
    );

    return (
      <div className="min-h-screen bg-[#F1F3F9] text-slate-700 flex items-center justify-center p-4 md:p-8 font-sans">
        <div className="w-full max-w-5xl bg-white rounded-3xl shadow-xl flex flex-col md:flex-row overflow-hidden border border-slate-200/90">
          
          {/* LEFT COLUMN: Branded Navy Banner with Synced Global Admin Logo */}
          <div className="w-full md:w-[410px] shrink-0 bg-[#112461] bg-gradient-to-b from-[#112461] via-[#0E1E52] to-[#0A153A] text-white p-8 md:p-10 flex flex-col justify-between min-h-[420px] md:min-h-[600px] relative items-center text-center">
            <div className="flex flex-col items-center gap-3 w-full">
              <div className="relative shrink-0 w-24 h-24 rounded-full overflow-hidden border-2 border-emerald-400 shadow-lg bg-white flex items-center justify-center">
                <img
                  src={globalAppLogo || customLogo || kopLeftLogo || DEFAULT_FALLBACK_LOGO}
                  alt="Logo GiziSync"
                  className="w-full h-full object-cover"
                  referrerPolicy="no-referrer"
                  onError={(e) => {
                    (e.target as HTMLImageElement).src = DEFAULT_FALLBACK_LOGO;
                  }}
                />
              </div>

              <div>
                <div className="font-extrabold text-lg tracking-wider text-white">GIZISYNC</div>
                <div className="text-[10px] text-emerald-400 font-bold uppercase tracking-widest -mt-0.5">SISPERMBG</div>
              </div>
            </div>

            <div className="my-6 md:my-8 space-y-3.5">
              <h1 className="text-xl md:text-2xl font-extrabold text-white leading-tight tracking-tight">
                Sistem Perencanaan &amp; Pemantauan Gizi Nasional
              </h1>
              <p className="text-slate-300 text-xs leading-relaxed font-normal">
                Platform terpadu untuk perencanaan menu siklus, perhitungan kecukupan gizi TKPI 2020, food cost, nota pesanan logistik, dan arsip bulanan SPPG.
              </p>
            </div>

            <div className="space-y-5 w-full">
              <div className="bg-white/5 border border-white/10 rounded-2xl p-4 text-left">
                <p className="text-[10px] font-bold text-blue-300 tracking-wider uppercase mb-2">MULTI-PLATFORM READY</p>
                <div className="flex flex-wrap gap-x-4 gap-y-1 text-[11px] text-slate-200">
                  <span>iPhone / Android</span>
                  <span>·</span>
                  <span>iPad / Tablet</span>
                  <span>·</span>
                  <span>Laptop / PC</span>
                </div>
              </div>

              <p className="text-[10px] text-slate-400 font-medium">
                © 2026 GIZISYNC | SISPERMBG Berdasarkan TKPI 2020.
              </p>
            </div>
          </div>

          {/* RIGHT COLUMN: Clean Per-User Email Google Sign-In */}
          <div className="bg-white p-7 md:p-10 flex-grow flex flex-col justify-between gap-6">
            <div className="space-y-5">
              <div className="space-y-1.5">
                <h2 className="text-2xl font-extrabold text-slate-900 tracking-tight">
                  Masuk Dashboard SISPERMBG
                </h2>
                <p className="text-slate-500 text-xs leading-relaxed">
                  Masuk langsung menggunakan akun email Google masing-masing untuk mengakses ruang kerja dashboard pribadi Anda.
                </p>
              </div>

              {demoNotice && (
                <div className="bg-rose-50 border border-rose-200 rounded-2xl p-3.5 text-xs text-rose-900 flex items-start justify-between gap-3">
                  <div className="flex items-start gap-2">
                    <Clock className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                    <span className="font-semibold">{demoNotice}</span>
                  </div>
                  <button
                    type="button"
                    onClick={clearDemoNotice}
                    className="text-rose-700 font-bold hover:underline shrink-0 cursor-pointer"
                  >
                    Tutup
                  </button>
                </div>
              )}

              {isUnauthorizedDomain && (
                <div className="bg-amber-50 border border-amber-300 rounded-2xl p-4 text-xs text-amber-950 space-y-3 text-left">
                  <div className="flex items-start justify-between gap-3">
                    <div className="flex items-start gap-2">
                      <AlertCircle className="w-4 h-4 text-amber-700 shrink-0 mt-0.5" />
                      <div>
                        <p className="font-extrabold text-amber-950">
                          Domain Belum Terdaftar di Firebase Authentication
                        </p>
                        <p className="text-[11px] text-amber-900/90 mt-0.5 leading-relaxed">
                          Login Google ditolak (<code>auth/unauthorized-domain</code>) karena domain aplikasi yang Anda buka saat ini belum ditambahkan ke daftar <strong>Authorized domains</strong> Firebase.
                        </p>
                      </div>
                    </div>
                    {authError === "UNAUTHORIZED_DOMAIN" && (
                      <button
                        type="button"
                        onClick={() => setAuthError(null)}
                        className="text-amber-800 font-bold hover:underline shrink-0 cursor-pointer"
                      >
                        Tutup
                      </button>
                    )}
                  </div>

                  {currentDomainHost && (
                    <div className="bg-white border border-amber-200 rounded-xl p-2.5 flex items-center justify-between gap-2">
                      <div className="min-w-0">
                        <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400">
                          Domain Aktif Saat Ini
                        </div>
                        <code className="text-xs font-bold text-slate-900 break-all">
                          {currentDomainHost}
                        </code>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleCopyCurrentDomain(currentDomainHost)}
                        className="shrink-0 inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold rounded-lg text-[11px] transition cursor-pointer"
                      >
                        {domainCopied ? (
                          <>
                            <Check className="w-3.5 h-3.5 text-emerald-600" />
                            <span>Tersalin</span>
                          </>
                        ) : (
                          <>
                            <Copy className="w-3.5 h-3.5" />
                            <span>Salin Domain</span>
                          </>
                        )}
                      </button>
                    </div>
                  )}

                  <ol className="space-y-1 text-[11px] text-amber-950/90 list-decimal pl-4">
                    <li>Klik tombol <strong>Buka Authorized Domains Firebase</strong> di bawah.</li>
                    <li>Pilih tab <strong>Authorized domains</strong> lalu klik <strong>Add domain</strong>.</li>
                    <li>Tempel domain <strong>{currentDomainHost || "aplikasi Anda"}</strong>, klik <strong>Add</strong>, lalu klik <strong>Cek Ulang &amp; Masuk Google</strong>.</li>
                  </ol>

                  <div className="flex flex-wrap items-center gap-2 pt-0.5">
                    <a
                      href={FIREBASE_AUTH_SETTINGS_URL}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-[#112461] hover:bg-[#193382] text-white font-bold rounded-xl text-xs transition"
                    >
                      <ExternalLink className="w-3.5 h-3.5" />
                      <span>Buka Authorized Domains Firebase</span>
                    </a>
                    <button
                      type="button"
                      onClick={handleRecheckDomainClick}
                      disabled={isRecheckingDomain}
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs transition cursor-pointer disabled:opacity-60"
                    >
                      <RefreshCw className={`w-3.5 h-3.5 ${isRecheckingDomain ? "animate-spin" : ""}`} />
                      <span>{isRecheckingDomain ? "Memeriksa..." : "Cek Ulang & Masuk Google"}</span>
                    </button>
                  </div>

                  {preferredAuthorizedUrls.length > 0 && (
                    <div className="pt-2 border-t border-amber-200/80 space-y-1.5">
                      <p className="text-[11px] font-bold text-amber-900">
                        Atau buka langsung melalui domain SISPERMBG yang sudah terdaftar:
                      </p>
                      <div className="flex flex-wrap gap-2">
                        {preferredAuthorizedUrls.map((host) => (
                          <a
                            key={host}
                            href={`https://${host}`}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-flex items-center gap-1 px-2.5 py-1 bg-white hover:bg-amber-100 border border-amber-300 text-amber-950 font-semibold rounded-lg text-[11px] transition"
                          >
                            <ExternalLink className="w-3 h-3" />
                            <span>{host}</span>
                          </a>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}

              {authError && authError !== "UNAUTHORIZED_DOMAIN" && (
                <div className="bg-rose-50 border border-rose-200 rounded-2xl p-4 text-xs text-rose-900 space-y-2.5">
                  <div className="flex items-start justify-between gap-3">
                    <div className="flex items-start gap-2">
                      <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                      <span className="font-semibold leading-relaxed">
                        {isIframeBlocked
                          ? "Browser memblokir jendela login Google di tampilan tertanam. Silakan buka aplikasi di tab baru agar Anda dapat memilih email Google Anda."
                          : authError}
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() => setAuthError(null)}
                      className="text-rose-700 font-bold hover:underline shrink-0 cursor-pointer"
                    >
                      Tutup
                    </button>
                  </div>
                  {isIframeBlocked && (
                    <a
                      href={typeof window !== "undefined" ? window.location.href : "#"}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-[#112461] hover:bg-[#193382] text-white font-bold rounded-xl text-xs transition"
                    >
                      <ExternalLink className="w-3.5 h-3.5" />
                      <span>Buka Aplikasi di Tab Baru</span>
                    </a>
                  )}
                </div>
              )}

              {/* PRIMARY ACTION: Sign In With Own Google Email */}
              <div className="space-y-3 pt-1">
                <button
                  type="button"
                  id="btn-google-email-login"
                  onClick={handleGoogleSignInClick}
                  disabled={isGoogleSigningIn}
                  className="w-full py-3.5 px-5 bg-[#112461] hover:bg-[#193382] active:scale-[0.99] text-white font-extrabold rounded-2xl transition cursor-pointer shadow-md shadow-indigo-950/15 disabled:opacity-60 text-sm flex items-center justify-center gap-3"
                >
                  {isGoogleSigningIn ? (
                    <>
                      <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                      <span>Membuka Pilihan Email Google...</span>
                    </>
                  ) : (
                    <>
                      <span className="w-7 h-7 rounded-lg bg-white flex items-center justify-center shrink-0 shadow-xs">
                        <svg className="w-4 h-4" viewBox="0 0 24 24" aria-hidden="true">
                          <path fill="#4285F4" d="M23.49 12.27c0-.79-.07-1.54-.2-2.27H12v4.3h6.44a5.5 5.5 0 0 1-2.39 3.61v3h3.87c2.26-2.08 3.57-5.15 3.57-8.64z" />
                          <path fill="#34A853" d="M12 24c3.24 0 5.95-1.07 7.93-2.91l-3.87-3c-1.07.72-2.44 1.15-4.06 1.15-3.12 0-5.77-2.11-6.72-4.95H1.28v3.1A12 12 0 0 0 12 24z" />
                          <path fill="#FBBC05" d="M5.28 14.29A7.2 7.2 0 0 1 4.9 12c0-.79.14-1.56.38-2.29v-3.1H1.28A12 12 0 0 0 0 12c0 1.94.46 3.77 1.28 5.39l4-3.1z" />
                          <path fill="#EA4335" d="M12 4.76c1.76 0 3.34.61 4.59 1.8l3.44-3.44C17.95 1.19 15.24 0 12 0A12 12 0 0 0 1.28 6.61l4 3.1C6.23 6.87 8.88 4.76 12 4.76z" />
                        </svg>
                      </span>
                      <span>Masuk dengan Email Google Anda</span>
                    </>
                  )}
                </button>
                <p className="text-[11px] text-slate-500 text-center">
                  Pilih akun Gmail pribadi Anda — tanpa perlu membuat atau mengingat kata sandi.
                </p>
              </div>

              {/* Concise Access Flow Guide */}
              <div className="bg-slate-50 border border-slate-200/80 rounded-2xl p-4 space-y-3 text-left">
                <div className="text-[11px] font-extrabold text-slate-800 uppercase tracking-wide flex items-center gap-1.5">
                  <ShieldCheck className="w-4 h-4 text-[#112461] shrink-0" />
                  <span>Alur Akses Pengguna SISPERMBG</span>
                </div>
                <ol className="space-y-2 text-xs text-slate-600">
                  <li className="flex items-start gap-2.5">
                    <span className="w-5 h-5 rounded-md bg-indigo-100 text-[#112461] font-extrabold text-[11px] flex items-center justify-center shrink-0 mt-0.5 tabular-nums">1</span>
                    <span>
                      <strong className="text-slate-900">Masuk dengan Email Sendiri:</strong> Klik tombol di atas lalu pilih email Google aktif Anda.
                    </span>
                  </li>
                  <li className="flex items-start gap-2.5">
                    <span className="w-5 h-5 rounded-md bg-amber-100 text-amber-800 font-extrabold text-[11px] flex items-center justify-center shrink-0 mt-0.5 tabular-nums">2</span>
                    <span>
                      <strong className="text-slate-900">Menunggu ACC Odhe Gizi:</strong> Pengguna baru otomatis diarahkan menghubungi <strong className="text-slate-900">Odhe Gizi</strong> agar akses disetujui melalui Panel Admin.
                    </span>
                  </li>
                  <li className="flex items-start gap-2.5">
                    <span className="w-5 h-5 rounded-md bg-emerald-100 text-emerald-800 font-extrabold text-[11px] flex items-center justify-center shrink-0 mt-0.5 tabular-nums">3</span>
                    <span>
                      <strong className="text-slate-900">Akses Penuh &amp; Tersimpan Otomatis:</strong> Setelah di-ACC, Anda langsung mendapatkan dashboard pribadi sendiri secara permanen (terkelola oleh Admin, tanpa Panel Admin).
                    </span>
                  </li>
                </ol>
              </div>

              {/* Direct WhatsApp Contact to Odhe Gizi */}
              <a
                href={buildOdheGiziWhatsAppUrl({ keperluan: "konsultasi & ACC Akses Dashboard SISPERMBG" })}
                target="_blank"
                rel="noopener noreferrer"
                className="block bg-amber-50/90 border border-amber-300/90 hover:border-emerald-500 rounded-2xl p-4 text-left transition-all hover:shadow-sm group"
              >
                <div className="flex items-center justify-between gap-2 mb-1">
                  <span className="text-xs font-extrabold text-amber-950 flex items-center gap-1.5">
                    <Info className="w-4 h-4 text-amber-600 shrink-0" />
                    Hubungi Odhe Gizi Untuk ACC Akses
                  </span>
                  <span className="text-[11px] font-extrabold text-emerald-800 bg-emerald-100 border border-emerald-300 px-2.5 py-0.5 rounded-lg inline-flex items-center gap-1 group-hover:bg-emerald-600 group-hover:text-white transition-colors shrink-0">
                    <MessageCircle className="w-3.5 h-3.5" />
                    Chat WA
                  </span>
                </div>
                <p className="text-xs text-slate-700 font-semibold">
                  Syukri Odhe | Ahli Gizi — <span className="font-mono text-emerald-700 font-extrabold">WA 082271095251</span>
                </p>
              </a>

              {/* PANEL MASUK DEMO 5 MENIT */}
              <div className="bg-emerald-50/70 border border-emerald-200 rounded-2xl p-4 space-y-2.5 text-left">
                <div className="flex items-center gap-2.5">
                  <div className="w-8 h-8 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0">
                    <Clock className="w-4 h-4" />
                  </div>
                  <div>
                    <h4 className="text-xs font-extrabold text-emerald-950 uppercase tracking-wide">
                      Panel Masuk Demo 5 Menit
                    </h4>
                    <p className="text-[11px] text-emerald-800 font-medium">
                      Coba fitur secara langsung · Keluar otomatis setelah 5 menit · Tanpa Panel Admin
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  id="btn-demo-5min-login"
                  onClick={loginDemoAccount}
                  className="w-full py-2.5 px-4 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.99] text-white font-extrabold rounded-xl text-xs transition flex items-center justify-center gap-2 shadow-xs cursor-pointer"
                >
                  <PlayCircle className="w-4 h-4" />
                  <span>Masuk Akun Demo (Batas Waktu 5 Menit)</span>
                </button>
              </div>
            </div>

            <div className="border-t border-slate-100 pt-4 flex flex-wrap justify-between gap-x-6 gap-y-2 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
              <span>AUTENTIKASI EMAIL GOOGLE</span>
              <span>DATA TENANT TERISOLASI</span>
              <span>TKPI 2020 · JUKNIS BGN 2025</span>
            </div>
          </div>

        </div>
      </div>
    );
  }

  // Check active login window: Tanggal Mulai Login (mulaiPada) & Tanggal Berakhir Login (berakhirPada; "" = Tanpa Batas)
  const localTodayStr = (() => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(
      now.getDate()
    ).padStart(2, "0")}`;
  })();

  const isNotYetStarted = (() => {
    if (isAdminUser || isDemoUser) return false;
    if (userProfile?.statusPersetujuan !== "aktif") return false;
    const mulai = typeof userProfile?.mulaiPada === "string" ? userProfile.mulaiPada.trim() : "";
    if (!mulai) return false;
    return localTodayStr < mulai;
  })();

  const isExpired = (() => {
    if (isAdminUser || isDemoUser) return false;
    const berakhir = typeof userProfile?.berakhirPada === "string" ? userProfile.berakhirPada.trim() : "";
    if (!berakhir) return false; // Empty string = Tanpa Batas (never expires)
    return localTodayStr > berakhir;
  })();

  const adminNoteText =
    typeof userProfile?.keterangan === "string" ? userProfile.keterangan.trim() : "";

  // If non-admin tenant is waiting for Admin ACC -> Direct immediately to Hubungi Odhe Gizi Untuk ACC Akses Anda
  if (
    !isDemoUser &&
    !isMainAdminAccount &&
    (userProfile?.statusPersetujuan === "menunggu" ||
      userProfile?.statusPersetujuan === "pending" ||
      !userProfile?.statusPersetujuan)
  ) {
    const waAccUrl = buildOdheGiziWhatsAppUrl({
      nama: regNama || userProfile?.namaLengkap || user.displayName || "",
      email: user.email || userProfile?.email || "",
      namaSPPG: regNamaSPPG || userProfile?.namaSPPG || "",
      keperluan: "mohon ACC Akses dashboard SISPERMBG saya agar akun langsung aktif",
    });

    return (
      <div className="min-h-screen bg-[#F1F3F9] text-slate-800 flex items-center justify-center p-4 md:p-8 font-sans">
        <div className="w-full max-w-2xl bg-white rounded-3xl shadow-xl border border-slate-200 overflow-hidden">
          {/* Top Branded Header Band */}
          <div className="bg-[#112461] bg-gradient-to-r from-[#112461] to-[#193382] text-white px-6 py-6 md:px-8 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-4">
              <div className="w-16 h-16 rounded-full overflow-hidden border-2 border-emerald-400 bg-white shrink-0 shadow-md">
                <img
                  src={globalAppLogo || customLogo || kopLeftLogo || DEFAULT_FALLBACK_LOGO}
                  alt="Logo SISPERMBG"
                  className="w-full h-full object-cover"
                  referrerPolicy="no-referrer"
                  onError={(e) => {
                    (e.target as HTMLImageElement).src = DEFAULT_FALLBACK_LOGO;
                  }}
                />
              </div>
              <div>
                <div className="text-[10px] font-bold uppercase tracking-widest text-emerald-300">
                  GIZISYNC · SISPERMBG
                </div>
                <h1 className="text-lg md:text-xl font-extrabold tracking-tight text-white">
                  Menunggu ACC Akses Dashboard
                </h1>
                <p className="text-xs text-slate-300 mt-0.5">
                  Email masuk: <strong className="text-white font-mono">{user.email}</strong>
                </p>
              </div>
            </div>
            <div className="self-start sm:self-center px-3 py-1.5 rounded-lg bg-amber-400/20 border border-amber-300/40 text-amber-200 text-xs font-extrabold flex items-center gap-1.5 shrink-0">
              <Clock className="w-3.5 h-3.5" />
              <span>Menunggu ACC Admin</span>
            </div>
          </div>

          <div className="p-6 md:p-8 space-y-6">
            {/* Primary Callout: Direct to Odhe Gizi for ACC */}
            <div className="bg-amber-50 border border-amber-200 rounded-2xl p-5 space-y-4">
              <div className="space-y-1.5">
                <h2 className="text-base font-extrabold text-slate-900">
                  Hubungi Odhe Gizi Untuk ACC Akses Anda
                </h2>
                {cloudSyncStatus === "syncing" ? (
                  <p className="text-xs text-slate-700 leading-relaxed flex items-center gap-2">
                    <RefreshCw className="w-3.5 h-3.5 text-amber-600 animate-spin shrink-0" />
                    <span>
                      Sedang mencatat akun email <strong className="text-slate-900">{user.email}</strong> ke Panel Admin SISPERMBG...
                    </span>
                  </p>
                ) : cloudSyncStatus === "error" ? (
                  <div className="mt-1 bg-rose-50 border border-rose-200 rounded-xl p-3.5 space-y-2.5 text-xs text-rose-900">
                    <div className="flex items-start gap-2">
                      <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                      <div className="space-y-1">
                        <p className="font-bold text-rose-950">
                          Pendaftaran email belum masuk ke Panel Admin Cloud
                        </p>
                        <p className="text-[11px] text-rose-800 leading-relaxed">
                          {cloudSyncError ||
                            "Koneksi ke server database sementara terganggu sehingga email Anda belum tercatat di Panel Admin."}
                        </p>
                      </div>
                    </div>
                    <button
                      type="button"
                      disabled={isRetryingCloudSync}
                      onClick={handleRetryCloudRegistration}
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-lg text-xs transition cursor-pointer disabled:opacity-60"
                    >
                      <RefreshCw className={`w-3.5 h-3.5 ${isRetryingCloudSync ? "animate-spin" : ""}`} />
                      <span>
                        {isRetryingCloudSync ? "Mengirim Ulang..." : "Kirim Ulang ke Panel Admin"}
                      </span>
                    </button>
                  </div>
                ) : (
                  <p className="text-xs text-slate-700 leading-relaxed">
                    Akun email <strong className="text-slate-900">{user.email}</strong> telah tercatat otomatis di Panel Admin SISPERMBG. Silakan hubungi <strong className="text-slate-900">Odhe Gizi</strong> melalui WhatsApp di bawah ini agar akses penuh dashboard pribadi Anda segera disetujui (ACC).
                  </p>
                )}
                {adminNoteText && (
                  <div className="mt-2 px-3 py-2 bg-white/90 border border-amber-200 rounded-xl text-xs text-slate-700">
                    <span className="font-bold text-amber-800">Keterangan Admin: </span>
                    <span>{adminNoteText}</span>
                  </div>
                )}
              </div>

              <a
                href={waAccUrl}
                target="_blank"
                rel="noopener noreferrer"
                id="btn-hubungi-odhe-gizi-acc"
                className="w-full py-3.5 px-5 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.99] text-white font-extrabold rounded-xl text-sm transition shadow-md shadow-emerald-600/20 flex items-center justify-center gap-2.5"
              >
                <MessageCircle className="w-5 h-5 shrink-0" />
                <span>Hubungi Odhe Gizi Untuk ACC Akses Anda (WA: 082271095251)</span>
              </a>

              <div className="flex items-start gap-2 text-[11px] text-emerald-900 bg-emerald-50 border border-emerald-200/80 rounded-xl p-3">
                <UserCheck className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <span>
                  <strong>Terbuka Otomatis Setelah ACC:</strong> Setelah Odhe Gizi menekan tombol <strong>ACC / Setujui</strong> di Panel Admin, halaman ini akan langsung membuka dashboard penuh milik Anda secara otomatis dan tersimpan permanen.
                </span>
              </div>
            </div>

            {statusCheckFeedback && (
              <div
                className={`rounded-xl p-3.5 text-xs font-semibold flex items-start gap-2 border ${
                  statusCheckFeedback.type === "success"
                    ? "bg-emerald-50 border-emerald-200 text-emerald-900"
                    : statusCheckFeedback.type === "error"
                    ? "bg-rose-50 border-rose-200 text-rose-900"
                    : "bg-indigo-50 border-indigo-200 text-indigo-900"
                }`}
              >
                {statusCheckFeedback.type === "success" ? (
                  <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                ) : (
                  <Info className="w-4 h-4 text-indigo-600 shrink-0 mt-0.5" />
                )}
                <span>{statusCheckFeedback.text}</span>
              </div>
            )}

            {/* Optional Quick Profile Form so Admin sees Tenant SPPG details immediately */}
            <form onSubmit={handleSaveWaitingIdentity} className="bg-slate-50 border border-slate-200 rounded-2xl p-5 space-y-4">
              <div className="flex items-center justify-between gap-2">
                <div>
                  <h3 className="text-xs font-extrabold text-slate-900 uppercase tracking-wide flex items-center gap-1.5">
                    <Building2 className="w-4 h-4 text-[#112461]" />
                    <span>Identitas SPPG Anda (Tampil di Panel Admin)</span>
                  </h3>
                  <p className="text-[11px] text-slate-500 mt-0.5">
                    Lengkapi nama &amp; SPPG Anda agar Odhe Gizi mudah mengenali akun Anda saat ACC.
                  </p>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div className="space-y-1">
                  <label className="font-bold text-slate-700 block">Nama Lengkap</label>
                  <input
                    type="text"
                    required
                    value={regNama}
                    onChange={(e) => setRegNama(e.target.value)}
                    placeholder="Nama Lengkap Anda"
                    className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-xl text-slate-900 font-semibold focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                  />
                </div>
                <div className="space-y-1">
                  <label className="font-bold text-slate-700 block">Profesi / Jabatan</label>
                  <select
                    value={regProfesi}
                    onChange={(e) => setRegProfesi(e.target.value)}
                    className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-xl text-slate-900 font-semibold focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 cursor-pointer"
                  >
                    <option value="Ahli Gizi">Ahli Gizi / Nutritionist</option>
                    <option value="Pengelola SPPG">Pengelola SPPG</option>
                    <option value="Kepala SPPG">Kepala Satuan Pelayanan</option>
                    <option value="Tenaga Akuntansi">Akuntan / Keuangan</option>
                    <option value="Chef / Juru Masak">Chef / Juru Masak</option>
                    <option value="Staf Operasional">Staf Operasional</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="font-bold text-slate-700 block">Nama SPPG / Instansi</label>
                  <input
                    type="text"
                    value={regNamaSPPG}
                    onChange={(e) => setRegNamaSPPG(e.target.value)}
                    placeholder="Contoh: SPPG Muna Barat"
                    className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-xl text-slate-900 font-semibold focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                  />
                </div>
                <div className="space-y-1">
                  <label className="font-bold text-slate-700 block">No. WhatsApp Aktif</label>
                  <input
                    type="text"
                    value={regNoHp}
                    onChange={(e) => setRegNoHp(e.target.value)}
                    placeholder="Contoh: 081234567890"
                    className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-xl text-slate-900 font-semibold focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                  />
                </div>
              </div>

              <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
                <div className="text-[11px] font-semibold text-emerald-700">
                  {waitingProfileSavedMsg && (
                    <span className="inline-flex items-center gap-1">
                      <CheckCircle2 className="w-3.5 h-3.5" />
                      {waitingProfileSavedMsg}
                    </span>
                  )}
                </div>
                <button
                  type="submit"
                  disabled={isSavingWaitingProfile}
                  className="px-4 py-2 bg-[#112461] hover:bg-[#193382] text-white font-bold text-xs rounded-xl transition cursor-pointer disabled:opacity-60"
                >
                  {isSavingWaitingProfile ? "Menyimpan..." : "Simpan Identitas ke Admin"}
                </button>
              </div>
            </form>

            {/* Footer Actions */}
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={signOutUser}
                className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl transition cursor-pointer text-xs flex items-center justify-center gap-1.5"
              >
                <LogOut className="w-3.5 h-3.5" />
                <span>Keluar &amp; Ganti Email</span>
              </button>
              <button
                type="button"
                onClick={handleManualStatusCheck}
                disabled={isCheckingStatus}
                className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold rounded-xl transition cursor-pointer text-xs flex items-center justify-center gap-1.5 disabled:opacity-60"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isCheckingStatus ? "animate-spin" : ""}`} />
                <span>{isCheckingStatus ? "Memeriksa Status ACC..." : "Periksa Status ACC Sekarang"}</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // If user is blocked / diblokir by Admin
  if (!isDemoUser && !isMainAdminAccount && userProfile?.statusPersetujuan === "diblokir") {
    const waBlockedUrl = buildOdheGiziWhatsAppUrl({
      nama: userProfile?.namaLengkap || user.displayName || "",
      email: user.email || userProfile?.email || "",
      namaSPPG: userProfile?.namaSPPG || "",
      keperluan: "memohon pengaktifan kembali akses dashboard SISPERMBG saya",
    });

    return (
      <div className="min-h-screen bg-[#F1F3F9] text-slate-800 flex items-center justify-center p-4 md:p-8 font-sans">
        <div className="w-full max-w-md bg-white border border-slate-200 rounded-3xl shadow-xl overflow-hidden text-center">
          <div className="bg-rose-600 text-white px-6 py-6 flex flex-col items-center gap-3">
            <div className="w-14 h-14 rounded-2xl bg-white/15 flex items-center justify-center">
              <Lock className="w-7 h-7 text-white" />
            </div>
            <div>
              <h2 className="text-lg font-extrabold">Akses Akun Dinonaktifkan</h2>
              <p className="text-xs text-rose-100 font-mono mt-0.5">{user.email}</p>
            </div>
          </div>

          <div className="p-6 space-y-5">
            <p className="text-xs text-slate-600 leading-relaxed">
              Akses dashboard Anda saat ini sedang dinonaktifkan atau ditangguhkan oleh Administrator melalui Panel Admin. Silakan hubungi <strong className="text-slate-900">Odhe Gizi</strong> untuk membuka kembali akses akun Anda.
            </p>
            {adminNoteText && (
              <div className="px-3.5 py-2.5 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-900 text-left">
                <span className="font-bold">Keterangan Admin: </span>
                <span>{adminNoteText}</span>
              </div>
            )}

            <a
              href={waBlockedUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="w-full py-3 px-4 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold rounded-xl text-xs transition flex items-center justify-center gap-2 shadow-sm"
            >
              <MessageCircle className="w-4 h-4" />
              <span>Hubungi Odhe Gizi (WA: 082271095251)</span>
            </a>

            <div className="flex gap-2.5 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={signOutUser}
                className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl transition cursor-pointer text-xs"
              >
                Keluar &amp; Ganti Email
              </button>
              <button
                type="button"
                onClick={handleManualStatusCheck}
                disabled={isCheckingStatus}
                className="flex-1 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl transition cursor-pointer text-xs flex items-center justify-center gap-1.5"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isCheckingStatus ? "animate-spin" : ""}`} />
                <span>Cek Status</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // If user profile access period has expired
  if (isExpired) {
    const waExpiredUrl = buildOdheGiziWhatsAppUrl({
      nama: userProfile?.namaLengkap || user.displayName || "",
      email: user.email || userProfile?.email || "",
      namaSPPG: userProfile?.namaSPPG || "",
      keperluan: `memperpanjang masa aktif akun SISPERMBG (berakhir ${userProfile?.berakhirPada || ""})`,
    });

    return (
      <div className="min-h-screen bg-[#F1F3F9] text-slate-800 flex items-center justify-center p-4 md:p-8 font-sans">
        <div className="w-full max-w-md bg-white border border-slate-200 rounded-3xl shadow-xl overflow-hidden text-center">
          <div className="bg-amber-600 text-white px-6 py-6 flex flex-col items-center gap-3">
            <div className="w-14 h-14 rounded-2xl bg-white/15 flex items-center justify-center">
              <Clock className="w-7 h-7 text-white" />
            </div>
            <div>
              <h2 className="text-lg font-extrabold">Masa Aktif Akun Berakhir</h2>
              <p className="text-xs text-amber-100 font-mono mt-0.5">{user.email}</p>
            </div>
          </div>

          <div className="p-6 space-y-5">
            <p className="text-xs text-slate-600 leading-relaxed">
              Masa aktif akun Anda berakhir pada <strong className="text-rose-600 font-mono">{userProfile?.berakhirPada}</strong>. Hubungi <strong className="text-slate-900">Odhe Gizi</strong> agar masa aktif dashboard Anda diperpanjang di Panel Admin.
            </p>
            {adminNoteText && (
              <div className="px-3.5 py-2.5 bg-amber-50 border border-amber-200 rounded-xl text-xs text-amber-900 text-left">
                <span className="font-bold">Keterangan Admin: </span>
                <span>{adminNoteText}</span>
              </div>
            )}

            <a
              href={waExpiredUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="w-full py-3 px-4 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold rounded-xl text-xs transition flex items-center justify-center gap-2 shadow-sm"
            >
              <MessageCircle className="w-4 h-4" />
              <span>Hubungi Odhe Gizi Untuk Perpanjangan</span>
            </a>

            <div className="flex gap-2.5 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={signOutUser}
                className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl transition cursor-pointer text-xs"
              >
                Keluar &amp; Ganti Email
              </button>
              <button
                type="button"
                onClick={handleManualStatusCheck}
                disabled={isCheckingStatus}
                className="flex-1 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl transition cursor-pointer text-xs flex items-center justify-center gap-1.5"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isCheckingStatus ? "animate-spin" : ""}`} />
                <span>Cek Status</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // If user is approved (ACC) but Tanggal Mulai Login is scheduled in the future
  if (isNotYetStarted) {
    const waScheduleUrl = buildOdheGiziWhatsAppUrl({
      nama: userProfile?.namaLengkap || user.displayName || "",
      email: user.email || userProfile?.email || "",
      namaSPPG: userProfile?.namaSPPG || "",
      keperluan: `konfirmasi jadwal Tanggal Mulai Login SISPERMBG (${userProfile?.mulaiPada || ""})`,
    });

    return (
      <div className="min-h-screen bg-[#F1F3F9] text-slate-800 flex items-center justify-center p-4 md:p-8 font-sans">
        <div className="w-full max-w-md bg-white border border-slate-200 rounded-3xl shadow-xl overflow-hidden text-center">
          <div className="bg-indigo-700 text-white px-6 py-6 flex flex-col items-center gap-3">
            <div className="w-14 h-14 rounded-2xl bg-white/15 flex items-center justify-center">
              <Calendar className="w-7 h-7 text-white" />
            </div>
            <div>
              <h2 className="text-lg font-extrabold">Jadwal Akses Belum Dimulai</h2>
              <p className="text-xs text-indigo-100 font-mono mt-0.5">{user.email}</p>
            </div>
          </div>

          <div className="p-6 space-y-5">
            <p className="text-xs text-slate-600 leading-relaxed">
              Akun Anda telah disetujui (ACC) oleh Admin dengan jadwal mulai login pada{" "}
              <strong className="text-indigo-700 font-mono">{userProfile?.mulaiPada}</strong>
              {userProfile?.berakhirPada ? (
                <>
                  {" "}sampai{" "}
                  <strong className="text-slate-900 font-mono">{userProfile.berakhirPada}</strong>
                </>
              ) : (
                <> (Tanpa Batas)</>
              )}
              .
            </p>
            {adminNoteText && (
              <div className="px-3.5 py-2.5 bg-indigo-50 border border-indigo-200 rounded-xl text-xs text-indigo-900 text-left">
                <span className="font-bold">Keterangan Admin: </span>
                <span>{adminNoteText}</span>
              </div>
            )}

            <a
              href={waScheduleUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="w-full py-3 px-4 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold rounded-xl text-xs transition flex items-center justify-center gap-2 shadow-sm"
            >
              <MessageCircle className="w-4 h-4" />
              <span>Hubungi Odhe Gizi (WA: 082271095251)</span>
            </a>

            <div className="flex gap-2.5 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={signOutUser}
                className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl transition cursor-pointer text-xs"
              >
                Keluar &amp; Ganti Email
              </button>
              <button
                type="button"
                onClick={handleManualStatusCheck}
                disabled={isCheckingStatus}
                className="flex-1 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl transition cursor-pointer text-xs flex items-center justify-center gap-1.5"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isCheckingStatus ? "animate-spin" : ""}`} />
                <span>Cek Status</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div id="sisper-app-root" className="min-h-screen bg-slate-50 text-slate-800 font-sans flex flex-col">
      {/* HEADER BANNER */}
      <header className="bg-gradient-to-r from-indigo-900 via-indigo-950 to-slate-900 text-white shadow-md py-5 px-6 shrink-0 relative overflow-hidden no-print">
        {/* Abstract design elements to express high craftsmanship */}
        <div className="absolute right-0 top-0 w-80 h-80 bg-indigo-500/10 rounded-full blur-3xl pointer-events-none"></div>
        <div className="absolute left-1/3 bottom-0 w-64 h-64 bg-emerald-500/5 rounded-full blur-2xl pointer-events-none"></div>
        
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center justify-between gap-5 relative z-10">
          <div className="flex items-center gap-3.5">
            <div className={`relative shrink-0 w-24 h-24 rounded-full overflow-hidden border-2 border-emerald-400 shadow-lg bg-white ${isAdminUser ? "group cursor-pointer" : ""}`}>
              <img
                src={globalAppLogo || customLogo || kopLeftLogo || DEFAULT_FALLBACK_LOGO}
                alt="Logo SPPG"
                className="w-full h-full object-cover"
                referrerPolicy="no-referrer"
                onError={(e) => {
                  (e.target as HTMLImageElement).src = DEFAULT_FALLBACK_LOGO;
                }}
              />
              {isAdminUser && (
                <>
                  <label
                    htmlFor="logo-uploader"
                    className="absolute inset-0 bg-slate-950/60 opacity-0 group-hover:opacity-100 transition-opacity flex flex-col items-center justify-center cursor-pointer text-[10px] text-white font-bold"
                  >
                    <Camera className="w-5 h-5 mb-1 text-emerald-400" />
                    Ubah Logo
                  </label>
                  <input
                    id="logo-uploader"
                    type="file"
                    accept="image/*"
                    onChange={handleLogoUpload}
                    className="hidden"
                  />
                </>
              )}
            </div>
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <div className="bg-white px-3.5 py-1.5 rounded-full inline-flex items-center shadow-md font-sans text-[15px] font-black tracking-wide leading-none select-none">
                  <span className="text-[#0D6D41] font-extrabold">Gizi</span>
                  <span className="text-[#2088C2] font-extrabold">Sync</span>
                  <span className="text-[#A2B4FF] mx-1 font-normal text-sm">|</span>
                  <span className="text-[#0D6D41] text-[11px] font-extrabold">Nutrition</span>
                  <span className="text-[#626FFF] text-[11px] ml-0.5 font-extrabold">Synchronized</span>
                </div>
              </div>
              <h1 className="text-2xl font-extrabold text-white tracking-tight mt-0.5">
                SISPERMBG <span className="font-light text-slate-300">| Sistem Perencanaan Makan Bergizi Gratis</span>
              </h1>
              <p className="text-xs text-slate-400 mt-0.5">
                {profile.namaLembaga} • {profile.alamat}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            {isDemoUser ? (
              <div className="flex flex-col items-end relative select-none">
                <span className="text-[9px] text-emerald-300 font-extrabold tracking-wider uppercase leading-none mb-1">
                  MODE DEMO 5 MENIT
                </span>
                <div className="flex items-center gap-2.5 bg-emerald-500/15 border border-emerald-400/40 px-3 py-1.5 rounded-2xl shadow-sm">
                  <div className="flex items-center gap-1.5 text-emerald-200 font-black text-xs font-mono">
                    <Clock className="w-3.5 h-3.5 text-emerald-400 animate-pulse" />
                    <span>
                      {String(Math.floor((demoSecondsLeft ?? 0) / 60)).padStart(2, "0")}:
                      {String((demoSecondsLeft ?? 0) % 60).padStart(2, "0")}
                    </span>
                  </div>
                  <span className="text-white/80 font-bold text-xs hidden sm:inline">Akun Demo</span>
                  <button
                    type="button"
                    onClick={signOutUser}
                    title="Keluar Mode Demo"
                    className="px-2 py-1 bg-rose-500/20 hover:bg-rose-500/40 rounded-lg text-rose-200 text-[11px] font-bold transition inline-flex items-center gap-1 cursor-pointer"
                  >
                    <LogOut className="w-3 h-3" />
                    Keluar
                  </button>
                </div>
              </div>
            ) : isCloudActive ? (
              <div className="flex flex-col items-end relative select-none">
                <span className="text-[9px] text-slate-400 font-extrabold tracking-wider uppercase leading-none mb-1">
                  PENGGUNA AKTIF
                </span>
                <div className="flex items-center gap-2 bg-rose-50/90 border border-rose-100/80 px-3 py-1.5 rounded-2xl shadow-sm">
                  {/* Avatar with Dentist/Nutritionist emoji */}
                  <div className="w-6 h-6 rounded-full bg-[#E5F6F0] border border-[#CBEFE3] flex items-center justify-center text-xs shrink-0 select-none shadow-inner">
                    👩‍⚕️
                  </div>
                  <span className="text-[#E6004C] font-extrabold text-xs tracking-tight">
                    {user?.email?.split("@")[0] || "syukriyusuf82"}
                  </span>
                  <button
                    type="button"
                    onClick={signOutUser}
                    title="Keluar Akun"
                    className="p-1 hover:bg-rose-100/60 rounded-lg text-[#E6004C] transition duration-150 ml-0.5 inline-flex items-center justify-center"
                  >
                    <LogOut className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
            ) : (
              <div className="flex flex-col items-end relative select-none">
                <span className="text-[9px] text-slate-400 font-extrabold tracking-wider uppercase leading-none mb-1">
                  SINKRONISASI
                </span>
                <button
                  type="button"
                  disabled={isCloudLoading}
                  onClick={signInWithGoogle}
                  className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 border border-slate-200 px-3.5 py-1.5 rounded-xl shadow-sm text-xs text-slate-700 font-extrabold transition disabled:opacity-50"
                >
                  <Cloud className="w-3.5 h-3.5 text-[#3B4FEB] animate-pulse" />
                  <span>Masuk dengan Google</span>
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* CORE STATS BAR */}
      <section className="bg-white border-b border-slate-200 py-3.5 px-6 sticky top-0 z-40 shadow-sm shrink-0 no-print">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row items-stretch md:items-center justify-between gap-4">
          
          {/* Main Navigation Tabs */}
          <nav className="flex flex-wrap items-center gap-1.5 bg-slate-100 p-1 rounded-2xl">
            {[
              { id: "dashboard", label: "Dashboard and Output", icon: Layers, color: "text-indigo-600" },
              { id: "sppg", label: "Profil SPPG", icon: Settings, color: "text-slate-500" },
              { id: "juknis", label: "Rujukan Juknis 2025", icon: BookOpen, color: "text-teal-600" },
              { id: "penerima", label: "Penerima Manfaat", icon: Users, color: "text-cyan-500" },
              { id: "menu", label: "Master Menu", icon: Utensils, color: "text-emerald-500" },
              { id: "foodcost", label: "Food Cost", icon: DollarSign, color: "text-amber-500" },
              { id: "notalogistik", label: "Gabungan Semua Food Cost", icon: FileText, color: "text-orange-500" },
              { id: "arsip", label: "Arsip Bulanan", icon: Archive, color: "text-indigo-600" },
              { id: "cekgizi", label: "Sandbox Gizi", icon: Search, color: "text-purple-500" },
              { id: "tkpi", label: "Database TKPI 2020", icon: Database, color: "text-rose-500" },
              { id: "handbook", label: "Handbook", icon: Book, color: "text-indigo-500" }
            ].map((tab) => {
              const TabIcon = tab.icon;
              return (
                <button
                  key={tab.id}
                  id={`tab-nav-${tab.id}`}
                  onClick={() => {
                    setActiveTab(tab.id);
                    sessionStorage.setItem("sisper_user_tab_chosen", "true");
                    setIsAdminPanelOpen(false);
                  }}
                  className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-bold transition duration-150 ${
                    activeTab === tab.id && !isAdminPanelOpen
                      ? "bg-white text-slate-900 shadow-sm"
                      : "text-slate-600 hover:text-slate-950 hover:bg-white/50"
                  }`}
                >
                  <TabIcon className={`w-4 h-4 ${tab.color}`} />
                  {tab.label}
                </button>
              );
            })}

            {isAdminUser && !isDemoUser && (
              <button
                id="tab-nav-admin"
                onClick={() => setIsAdminPanelOpen(!isAdminPanelOpen)}
                className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-bold transition duration-150 border ${
                  isAdminPanelOpen
                    ? "bg-rose-600 text-white border-rose-600 shadow-sm"
                    : "bg-rose-50 hover:bg-rose-100 text-rose-700 border-rose-100 hover:border-rose-200"
                }`}
              >
                <Lock className={`w-4 h-4 ${isAdminPanelOpen ? "text-white" : "text-rose-500 animate-pulse"}`} />
                Panel Admin
              </button>
            )}
          </nav>

          {/* Minimal info */}
          <div className="hidden lg:flex items-center gap-4 text-xs font-medium text-slate-500">
            <div className="flex items-center gap-1.5">
              <Calendar className="w-4 h-4 text-slate-400" />
              <span>Siklus {Math.max(profile.periodeDates?.length || 10, 1)} Hari Kerja Aktif</span>
            </div>
            <span className="text-slate-300">|</span>
            <div className="flex items-center gap-1.5">
              <Heart className="w-4 h-4 text-rose-500 animate-pulse" />
              <span>Sesuai Juknis BGN No. 401.1/2025</span>
            </div>
          </div>

        </div>
      </section>

      {/* ACTIVE SCREEN CONTENT */}
      <main key={user?.uid || "guest"} className={`flex-grow p-6 ${isAdminPanelOpen && isAdminUser ? "max-w-full" : activeTab === "notalogistik" ? "max-w-full px-4 md:px-8" : activeTab === "foodcost" ? "max-w-[1600px]" : "max-w-7xl"} w-full mx-auto overflow-y-auto`}>
        {isAdminPanelOpen && isAdminUser ? (
          <AdminPanel
            onClose={() => setIsAdminPanelOpen(false)}
            globalLogoUrl={globalAppLogo || customLogo}
            onUploadGlobalLogo={handleLogoUpload}
            adminEmail={user?.email || userProfile?.email || ""}
          />
        ) : (
          <>
            {/* PERSONALIZED ROLE-BASED DASHBOARD BANNER FOR USERS */}
            {userProfile && (
              <div className="mb-6 bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 rounded-3xl p-5 md:p-6 text-white shadow-md relative overflow-hidden no-print">
                <div className="absolute right-0 top-0 w-80 h-80 bg-blue-500/10 rounded-full blur-3xl pointer-events-none"></div>
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-5 relative z-10">
                  <div className="space-y-1.5">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 px-3 py-1 rounded-full text-[11px] font-black uppercase tracking-wider flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                        DASBOR {userProfile.profesi?.toUpperCase() || "PENGAWAS GIZI"}
                      </span>
                      <span className="bg-white/10 text-slate-200 px-3 py-1 rounded-full text-[11px] font-bold">
                        {userProfile.namaSPPG || profile.namaLembaga}
                      </span>
                      <span className="bg-blue-500/20 text-blue-300 px-2.5 py-0.5 rounded-full text-[10px] font-semibold border border-blue-400/30">
                        Status: {userProfile.statusPersetujuan === "aktif" ? "Akses Disetujui (Aktif)" : "Menunggu Persetujuan Admin"}
                      </span>
                    </div>
                    <h2 className="text-xl md:text-2xl font-black text-white tracking-tight">
                      Selamat Datang, {userProfile.namaLengkap || user?.email}
                    </h2>
                    <p className="text-xs text-slate-300 leading-relaxed max-w-3xl">
                      {userProfile.profesi?.toLowerCase().includes("keuangan") || userProfile.profesi?.toLowerCase().includes("akuntan")
                        ? "Dasbor Anda dikustomisasi khusus untuk Manajemen Keuangan, Rekapitulasi Food Cost, Anggaran Belanja, dan Verifikasi Nota Logistik SPPG."
                        : userProfile.profesi?.toLowerCase().includes("chef") || userProfile.profesi?.toLowerCase().includes("masak") || userProfile.profesi?.toLowerCase().includes("koki")
                        ? "Dasbor Anda dikustomisasi khusus untuk Operasional Dapur, Master Resep Siklus 10 Hari, Pembagian Gramatur Porsi, dan Nota Pesanan Bahan Baku."
                        : userProfile.profesi?.toLowerCase().includes("kepala") || userProfile.profesi?.toLowerCase().includes("pimpinan") || userProfile.profesi?.toLowerCase().includes("manager")
                        ? "Dasbor Pengawasan Executive SPPG: Monitoring Realisasi Anggaran, Kepatuhan Juknis BGN 2025, dan Status Penerima Manfaat."
                        : userProfile.profesi?.toLowerCase().includes("kesehatan") || userProfile.profesi?.toLowerCase().includes("pengawas")
                        ? "Dasbor Pengawasan Lapangan & Kesehatan: Tracking kecukupan nutrisi penerima manfaat (Siswa SD/SMP/SMA, Ibu Hamil & Balita)."
                        : "Dasbor Utama Perencanaan Gizi: Analisis Kecukupan Standar AKG 2025, Penyusunan Menu Siklus, dan Sandbox Formulasi Nutrisi."}
                    </p>
                  </div>

                  <div className="flex flex-wrap items-center gap-2 shrink-0">
                    <button
                      onClick={() => {
                        const targetTab = getDefaultTabForProfesi(userProfile.profesi, userProfile.peran);
                        setActiveTab(targetTab);
                        sessionStorage.setItem("sisper_user_tab_chosen", "true");
                      }}
                      className="px-4 py-2.5 bg-blue-500 hover:bg-blue-400 active:scale-98 text-white font-extrabold text-xs rounded-xl shadow-sm transition-all flex items-center gap-2 cursor-pointer"
                    >
                      <Layers className="w-4 h-4" />
                      <span>Fokus Dasbor Peranku ({userProfile.profesi || "Gizi"})</span>
                    </button>
                  </div>
                </div>
              </div>
            )}

            {activeTab === "dashboard" && (
              <div className="space-y-6">
                <DashboardOutputs
                  foodCostDays={foodCostDays}
                  tkpiList={tkpiList}
                  masterMenu={masterMenu}
                  harianPM={harianPM}
                  profile={profile}
                  kopLine1={kopLine1}
                  setKopLine1={setKopLine1}
                  kopLine2={kopLine2}
                  setKopLine2={setKopLine2}
                  kopLine3={kopLine3}
                  setKopLine3={setKopLine3}
                  kopLine4={kopLine4}
                  setKopLine4={setKopLine4}
                  leftLogo={kopLeftLogo}
                  setLeftLogo={handleUpdateLeftLogo}
                  rightLogo={kopRightLogo}
                  setRightLogo={handleUpdateRightLogo}
                  leftLogoCrop={leftLogoCrop}
                  setLeftLogoCrop={setLeftLogoCrop}
                  rightLogoCrop={rightLogoCrop}
                  setRightLogoCrop={setRightLogoCrop}
                  paperSize={paperSize}
                  setPaperSize={setPaperSize}
                  isAdmin={isAdminUser}
                />
              </div>
            )}

        {activeTab === "notalogistik" && (
          <NotaPesananLogistikTab
            profile={profile}
            foodCostDays={foodCostDays}
            tkpiList={tkpiList}
            harianPM={harianPM}
            masterMenu={masterMenu}
            kopLine1={kopLine1}
            setKopLine1={setKopLine1}
            kopLine2={kopLine2}
            setKopLine2={setKopLine2}
            kopLine3={kopLine3}
            setKopLine3={setKopLine3}
            kopLine4={kopLine4}
            setKopLine4={setKopLine4}
            leftLogo={kopLeftLogo}
            setLeftLogo={handleUpdateLeftLogo}
            rightLogo={kopRightLogo}
            setRightLogo={handleUpdateRightLogo}
            leftLogoCrop={leftLogoCrop}
            setLeftLogoCrop={setLeftLogoCrop}
            rightLogoCrop={rightLogoCrop}
            setRightLogoCrop={setRightLogoCrop}
            paperSize={paperSize}
            setPaperSize={setPaperSize}
            isAdmin={isAdminUser}
          />
        )}

        {activeTab === "arsip" && (
          <ArsipTab
            profile={profile}
            harianPM={harianPM}
            masterMenu={masterMenu}
            foodCostDays={foodCostDays}
            tkpiList={tkpiList}
            pmSettings={pmSettings}
            savedPeriods={savedPeriods}
            setSavedPeriods={setSavedPeriods}
            activeSekolahPM={sekolahPM}
            activeTigaBPM={tigaBPM}
            kopLines={[kopLine1, kopLine2, kopLine3, kopLine4]}
            userEmail={user?.email}
            isDemo={isDemoUser}
          />
        )}

        {activeTab === "sppg" && (
          <SPPGProfileTab
            profile={profile}
            onChange={setProfile}
            savedPeriods={savedPeriods}
            setSavedPeriods={setSavedPeriods}
            activeState={{
              sekolahPM,
              tigaBPM,
              harianPM,
              masterMenu,
              foodCostDays
            }}
            loadPeriod={(period: any) => {
              setProfile(period.profile);
              setSekolahPM(period.sekolahPM);
              setTigaBPM(period.tigaBPM);
              setHarianPM(period.harianPM);
              setMasterMenu(period.masterMenu);
              setFoodCostDays(period.foodCostDays);
            }}
            startNewPeriod={() => {
              // Create clean blank state
              setProfile({
                ...profile,
                periodeDates: Array.from({ length: 10 }, () => ""),
                awalPeriodeBerikutnya: ""
              });
              setSekolahPM([]);
              setTigaBPM([]);
              setHarianPM(EMPTY_HARIAN_PM);
              setMasterMenu(EMPTY_MASTER_MENU);
              setFoodCostDays(EMPTY_FOOD_COST_DAYS);
            }}
          />
        )}

        {activeTab === "penerima" && (
          <PenerimaManfaatTab
            harianPM={harianPM}
            onChange={setHarianPM}
            profile={profile}
            pmSettings={pmSettings}
            setPmSettings={setPmSettings}
            kopLine1={kopLine1}
            setKopLine1={setKopLine1}
            kopLine2={kopLine2}
            setKopLine2={setKopLine2}
            kopLine3={kopLine3}
            setKopLine3={setKopLine3}
            kopLine4={kopLine4}
            setKopLine4={setKopLine4}
            leftLogo={kopLeftLogo}
            setLeftLogo={handleUpdateLeftLogo}
            rightLogo={kopRightLogo}
            setRightLogo={handleUpdateRightLogo}
            leftLogoCrop={leftLogoCrop}
            setLeftLogoCrop={setLeftLogoCrop}
            rightLogoCrop={rightLogoCrop}
            setRightLogoCrop={setRightLogoCrop}
            paperSize={paperSize}
            setPaperSize={setPaperSize}
            isAdmin={isAdminUser}
          />
        )}

        {activeTab === "menu" && (
          <MasterMenuTab 
            menu={masterMenu} 
            onChange={setMasterMenu} 
            profile={profile}
            onProfileChange={(updatedProfile) => {
              setProfile(updatedProfile);
              const targetDays = updatedProfile.periodeDates?.length || 0;
              if (targetDays > harianPM.length) {
                const templateSasaran = harianPM[0]?.sasaran || EMPTY_SASARAN_LIST;
                const nextHarian = [...harianPM];
                for (let d = harianPM.length + 1; d <= targetDays; d++) {
                  nextHarian.push({
                    hariKe: d,
                    sasaran: templateSasaran.map((s) => ({ ...s })),
                  });
                }
                setHarianPM(nextHarian);
              }
            }}
            onSendToFoodCost={(updatedMenu, dates) => {
              const normMenu = ensureMasterMenuPlans(updatedMenu, dates);
              const targetDays = Math.max(dates.length, 1);

              // Ensure harianPM covers all planned days
              if (targetDays > harianPM.length) {
                const templateSasaran = harianPM[0]?.sasaran || EMPTY_SASARAN_LIST;
                const nextHarian = [...harianPM];
                for (let d = harianPM.length + 1; d <= targetDays; d++) {
                  nextHarian.push({
                    hariKe: d,
                    sasaran: templateSasaran.map((s) => ({ ...s })),
                  });
                }
                setHarianPM(nextHarian);
              }

              // Map Master Menu plans & dates into FoodCostDay entries automatically
              const nextFoodCostDays = [...foodCostDays];
              const plans = normMenu.plans || [];

              plans.forEach((plan) => {
                const targets: FoodCostDay["jenisMenu"][] = [];
                const hasSekolah = plan.sasaranIds.some((id) => SEKOLAH_SASARAN_IDS.includes(id));
                const hasMpAsi = plan.sasaranIds.some((id) => MPASI_SASARAN_IDS.includes(id));
                const has3B = plan.sasaranIds.some((id) =>
                  ["anak_balita", "anak_balita_13_59", "ibu_hamil", "ibu_menyusui"].includes(id)
                );

                if (hasSekolah) {
                  if (plan.dietType === "Alergi") targets.push("Alergi");
                  else if (plan.dietType === "Kering") targets.push("Kering");
                  else targets.push("Basah");
                }
                if (hasMpAsi || plan.dietType === "MP-ASI") {
                  targets.push("MP-ASI");
                }
                if (has3B) {
                  if (plan.dietType === "Alergi") targets.push("3B-Alergi");
                  else if (plan.dietType === "Kering") targets.push("3B-Kering");
                  else if (plan.dietType !== "MP-ASI") targets.push("3B-Basah");
                }

                targets.forEach((jenisMenu) => {
                  for (let idx = 0; idx < targetDays; idx++) {
                    const dayNum = idx + 1;
                    const menuName = composeMenuName(plan.days?.[idx]);
                    const tanggal = dates[idx] || "";
                    const existingIdx = nextFoodCostDays.findIndex(
                      (fc) => fc.hariKe === dayNum && fc.jenisMenu === jenisMenu
                    );
                    if (existingIdx >= 0) {
                      nextFoodCostDays[existingIdx] = {
                        ...nextFoodCostDays[existingIdx],
                        namaMenu: menuName || nextFoodCostDays[existingIdx].namaMenu,
                        tanggal,
                        sasaranIds: plan.sasaranIds,
                        masterMenuPlanId: plan.id,
                      };
                    } else {
                      nextFoodCostDays.push({
                        hariKe: dayNum,
                        jenisMenu,
                        namaMenu: menuName,
                        tanggal,
                        sasaranIds: plan.sasaranIds,
                        masterMenuPlanId: plan.id,
                        porsiBesarBahan: [],
                        porsiKecilBahan: [],
                        bufferPct: 5,
                      });
                    }
                  }
                });
              });

              setFoodCostDays(nextFoodCostDays);
            }}
            onNavigateToFoodCost={() => setActiveTab("foodcost")}
            customLogo={customLogo} 
            kopLine1={kopLine1}
            setKopLine1={setKopLine1}
            kopLine2={kopLine2}
            setKopLine2={setKopLine2}
            kopLine3={kopLine3}
            setKopLine3={setKopLine3}
            kopLine4={kopLine4}
            setKopLine4={setKopLine4}
            leftLogo={kopLeftLogo}
            setLeftLogo={handleUpdateLeftLogo}
            rightLogo={kopRightLogo}
            setRightLogo={handleUpdateRightLogo}
            leftLogoCrop={leftLogoCrop}
            setLeftLogoCrop={setLeftLogoCrop}
            rightLogoCrop={rightLogoCrop}
            setRightLogoCrop={setRightLogoCrop}
            paperSize={paperSize}
            setPaperSize={setPaperSize}
            isAdmin={isAdminUser}
          />
        )}

        {activeTab === "foodcost" && (
          <FoodCostTab
            foodCostDays={foodCostDays}
            tkpiList={tkpiList}
            masterMenu={masterMenu}
            harianPM={harianPM}
            onFoodCostDaysChange={setFoodCostDays}
            profile={profile}
            customLogo={customLogo}
            kopLine1={kopLine1}
            setKopLine1={setKopLine1}
            kopLine2={kopLine2}
            setKopLine2={setKopLine2}
            kopLine3={kopLine3}
            setKopLine3={setKopLine3}
            kopLine4={kopLine4}
            setKopLine4={setKopLine4}
            leftLogo={kopLeftLogo}
            setLeftLogo={handleUpdateLeftLogo}
            rightLogo={kopRightLogo}
            setRightLogo={handleUpdateRightLogo}
            leftLogoCrop={leftLogoCrop}
            setLeftLogoCrop={setLeftLogoCrop}
            rightLogoCrop={rightLogoCrop}
            setRightLogoCrop={setRightLogoCrop}
            paperSize={paperSize}
            setPaperSize={setPaperSize}
            pmSettings={pmSettings}
            setPmSettings={setPmSettings}
            onNavigateToMasterMenu={() => setActiveTab("menu")}
            onNavigateToPM={() => setActiveTab("penerima")}
            isAdmin={isAdminUser}
          />
        )}

        {activeTab === "cekgizi" && (
          <CekGiziTab 
            tkpiList={tkpiList} 
            items={cekGiziItems}
            setItems={setCekGiziItems}
            kopLine1={kopLine1}
            setKopLine1={setKopLine1}
            kopLine2={kopLine2}
            setKopLine2={setKopLine2}
            kopLine3={kopLine3}
            setKopLine3={setKopLine3}
            kopLine4={kopLine4}
            setKopLine4={setKopLine4}
            leftLogo={kopLeftLogo}
            setLeftLogo={handleUpdateLeftLogo}
            rightLogo={kopRightLogo}
            setRightLogo={handleUpdateRightLogo}
            leftLogoCrop={leftLogoCrop}
            setLeftLogoCrop={setLeftLogoCrop}
            rightLogoCrop={rightLogoCrop}
            setRightLogoCrop={setRightLogoCrop}
            paperSize={paperSize}
            setPaperSize={setPaperSize}
            isAdmin={isAdminUser}
          />
        )}

        {activeTab === "juknis" && (
          <RujukanJuknisTab />
        )}

        {activeTab === "tkpi" && (
          <div className="space-y-6" id="tkpi-database-tab">
            {/* TKPI top summary and excel upload */}
            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-5">
              <div className="space-y-1">
                <h3 className="text-base font-semibold text-slate-800 flex items-center gap-2">
                  <Database className="w-5 h-5 text-rose-500" />
                  Database TKPI 2020 — Tersinkronisasi Spreadsheet TKPI 2020
                </h3>
                <p className="text-xs text-slate-500">
                  Data komposisi gizi per 100 gram sesuai spreadsheet resmi TKPI 2020 ({INITIAL_TKPI_DATABASE.length} bahan). Tambah bahan kustom atau unggah CSV/Excel mengikuti kolom standar TKPI.
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={restoreDefaultTkpi}
                  type="button"
                  className="flex items-center gap-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold cursor-pointer transition"
                  title="Sinkronkan ulang seluruh data standar spreadsheet TKPI 2020"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  Sinkronkan TKPI 2020
                </button>
                <button
                  onClick={clearAllTkpi}
                  type="button"
                  className="flex items-center gap-1.5 px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-xl text-xs font-semibold cursor-pointer transition"
                  title="Kosongkan seluruh data TKPI untuk unggah kustom"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  Kosongkan TKPI
                </button>
                <button
                  onClick={handleDownloadTemplate}
                  type="button"
                  className="flex items-center gap-1.5 px-3 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-600 rounded-xl text-xs font-semibold cursor-pointer transition"
                >
                  <Download className="w-3.5 h-3.5" />
                  Unduh Format TKPI 2020
                </button>
                <label className="flex items-center gap-1.5 px-3.5 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold cursor-pointer transition shadow-xs">
                  <Upload className="w-3.5 h-3.5" />
                  Unggah Excel / CSV
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={handleTkpiUpload} className="hidden" />
                </label>
              </div>
            </div>

            {tkpiNotice && (
              <div
                className={`p-3.5 rounded-xl border text-xs font-semibold flex items-center justify-between gap-2 ${
                  tkpiNotice.type === "success"
                    ? "bg-emerald-50 border-emerald-200 text-emerald-900"
                    : tkpiNotice.type === "error"
                      ? "bg-rose-50 border-rose-200 text-rose-900"
                      : "bg-indigo-50 border-indigo-200 text-indigo-900"
                }`}
              >
                <span>{tkpiNotice.text}</span>
                <button
                  type="button"
                  onClick={() => setTkpiNotice(null)}
                  className="px-2 py-0.5 text-xs font-bold hover:underline cursor-pointer"
                >
                  Tutup
                </button>
              </div>
            )}

            {/* Category Management Panel */}
            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm space-y-3">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-slate-100">
                <div className="flex items-center gap-2">
                  <Tags className="w-4 h-4 text-indigo-600" />
                  <h4 className="font-bold text-slate-800 text-sm">Kelola Kategori Bahan Pangan ({availableCategories.length} Kategori)</h4>
                </div>
                <span className="text-[11px] text-slate-400">Pilih kategori untuk memfilter, atau klik tombol hapus (<Trash2 className="w-3 h-3 inline text-rose-500" />) pada kategori untuk menghapus seluruh bahannya.</span>
              </div>

              <div className="flex flex-wrap gap-2 pt-1 max-h-36 overflow-y-auto pr-1">
                <button
                  type="button"
                  onClick={() => setSelectedCategoryFilter("all")}
                  className={`px-3 py-1.5 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                    selectedCategoryFilter === "all"
                      ? "bg-slate-900 text-white shadow-xs"
                      : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                  }`}
                >
                  Semua Kategori ({tkpiList.length})
                </button>
                {availableCategories.map((cat) => {
                  const count = categoryCountsMap[cat] || 0;
                  const isSelected = selectedCategoryFilter === cat;
                  return (
                    <div
                      key={cat}
                      className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold border transition ${
                        isSelected
                          ? "bg-indigo-50 border-indigo-200 text-indigo-700 shadow-2xs"
                          : "bg-white border-slate-200 text-slate-700 hover:bg-slate-50"
                      }`}
                    >
                      <button
                        type="button"
                        onClick={() => setSelectedCategoryFilter(isSelected ? "all" : cat)}
                        className="flex items-center gap-1.5 cursor-pointer"
                      >
                        <span>{cat}</span>
                        <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-mono ${isSelected ? "bg-indigo-200 text-indigo-900" : "bg-slate-100 text-slate-600"}`}>
                          {count}
                        </span>
                      </button>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          deleteCategory(cat);
                        }}
                        className="text-slate-400 hover:text-rose-600 p-0.5 rounded transition cursor-pointer ml-1"
                        title={`Hapus kategori "${cat}" beserta ${count} bahannya`}
                      >
                        <Trash2 className="w-3 h-3" />
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="grid grid-cols-1 xl:grid-cols-12 gap-6">
              {/* Form to add item manually */}
              <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm h-fit xl:col-span-4 space-y-4">
                <h4 className="font-bold text-slate-800 text-sm pb-2 border-b border-slate-100 flex items-center gap-1.5">
                  <Plus className="w-4 h-4 text-rose-500" />
                  Tambah Bahan Makanan Kustom
                </h4>

                <form onSubmit={handleAddTkpiSubmit} className="space-y-4 text-xs">
                  {/* GROUP 1: Informasi Utama */}
                  <div className="space-y-3 bg-slate-50/50 p-3 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-rose-700 uppercase tracking-wider block">1. Informasi Utama & Klasifikasi</span>
                    
                    <div className="space-y-1">
                      <label className="font-bold text-slate-500 uppercase tracking-wider block">Nama Bahan Pangan <span className="text-red-500">*</span></label>
                      <input
                        id="form-tkpi-nama"
                        type="text"
                        required
                        value={newTkpiItem.nama || ""}
                        onChange={(e) => setNewTkpiItem({ ...newTkpiItem, nama: e.target.value })}
                        className="w-full px-3 py-1.5 border border-slate-200 rounded-lg text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-rose-500 bg-white"
                        placeholder="e.g. Ikan Mas Segar"
                      />
                    </div>

                    <div className="grid grid-cols-2 gap-2">
                      <div className="space-y-1">
                        <label className="font-bold text-slate-500 uppercase tracking-wider block">Kategori</label>
                        <input
                          id="form-tkpi-kategori"
                          type="text"
                          value={newTkpiItem.kategori || ""}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, kategori: e.target.value })}
                          className="w-full px-2 py-1.5 border border-slate-200 rounded-lg text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-rose-500 bg-white"
                          placeholder="e.g. Ikan & Kerang"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="font-bold text-slate-500 uppercase tracking-wider block">(SUMBER)</label>
                        <input
                          id="form-tkpi-sumber"
                          type="text"
                          value={newTkpiItem.sumber || ""}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, sumber: e.target.value })}
                          className="w-full px-2 py-1.5 border border-slate-200 rounded-lg text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-rose-500 bg-white"
                          placeholder="e.g. Hewani"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-2">
                      <div className="space-y-1">
                        <label className="font-bold text-slate-500 uppercase tracking-wider block">Berat Standar (g)</label>
                        <input
                          id="form-tkpi-beratStandar"
                          type="number"
                          min="1"
                          value={newTkpiItem.beratStandar || ""}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, beratStandar: Number(e.target.value) })}
                          className="w-full px-2 py-1.5 border border-slate-200 rounded-lg text-xs text-slate-800 font-mono text-right bg-white"
                          placeholder="100"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="font-bold text-slate-500 uppercase tracking-wider block">BDD (%)</label>
                        <input
                          id="form-tkpi-bdd"
                          type="number"
                          min="1"
                          max="100"
                          value={newTkpiItem.bdd || ""}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, bdd: Number(e.target.value) })}
                          className="w-full px-2 py-1.5 border border-slate-200 rounded-lg text-xs text-slate-800 font-mono text-right bg-white"
                          placeholder="100"
                        />
                      </div>
                    </div>
                  </div>

                  {/* GROUP 2: Proksimat / Makro Gizi */}
                  <div className="space-y-3 bg-amber-50/20 p-3 rounded-xl border border-amber-100/30">
                    <span className="text-[10px] font-black text-amber-700 uppercase tracking-wider block">2. Kandungan Makro & Proksimat</span>
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Energi (kkal)</label>
                        <input
                          id="form-tkpi-energi"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.energi || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, energi: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Protein (g)</label>
                        <input
                          id="form-tkpi-protein"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.protein || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, protein: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Lemak (g)</label>
                        <input
                          id="form-tkpi-lemak"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.lemak || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, lemak: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Karbohidrat (g)</label>
                        <input
                          id="form-tkpi-kh"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.kh || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, kh: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Serat (g)</label>
                        <input
                          id="form-tkpi-serat"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.serat || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, serat: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Abu (g)</label>
                        <input
                          id="form-tkpi-abu"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.abu || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, abu: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1 col-span-2 sm:col-span-1">
                        <label className="text-slate-500 block text-[10px]">AIR (g)</label>
                        <input
                          id="form-tkpi-air"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.air || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, air: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                    </div>
                  </div>

                  {/* GROUP 3: Mineral */}
                  <div className="space-y-3 bg-indigo-50/20 p-3 rounded-xl border border-indigo-100/30">
                    <span className="text-[10px] font-black text-indigo-700 uppercase tracking-wider block">3. Kandungan Mineral</span>
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Kalsium (mg)</label>
                        <input
                          id="form-tkpi-ca"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.ca || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, ca: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Fosfor (mg)</label>
                        <input
                          id="form-tkpi-p"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.p || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, p: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Zat Besi (mg)</label>
                        <input
                          id="form-tkpi-fe"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.fe || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, fe: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Natrium (mg)</label>
                        <input
                          id="form-tkpi-na"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.na || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, na: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Kalium (mg)</label>
                        <input
                          id="form-tkpi-k"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.k || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, k: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Tembaga (mg)</label>
                        <input
                          id="form-tkpi-cu"
                          type="number"
                          step="0.01"
                          value={newTkpiItem.cu || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, cu: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Seng (mg)</label>
                        <input
                          id="form-tkpi-zn"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.zn || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, zn: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                    </div>
                  </div>

                  {/* GROUP 4: Vitamin */}
                  <div className="space-y-3 bg-rose-50/20 p-3 rounded-xl border border-rose-100/30">
                    <span className="text-[10px] font-black text-rose-700 uppercase tracking-wider block">4. Kandungan Vitamin</span>
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">Retinol (mcg)</label>
                        <input
                          id="form-tkpi-retinol"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.retinol || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, retinol: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">BKAR (mcg)</label>
                        <input
                          id="form-tkpi-b-karoten"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.b_karoten || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, b_karoten: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">KARTOTAL (mcg)</label>
                        <input
                          id="form-tkpi-kartotal"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.kartotal || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, kartotal: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">THIAMIN (mg)</label>
                        <input
                          id="form-tkpi-thiamin"
                          type="number"
                          step="0.01"
                          value={newTkpiItem.thiamin || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, thiamin: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">RIBOFLAVIN (mg)</label>
                        <input
                          id="form-tkpi-riboflavin"
                          type="number"
                          step="0.01"
                          value={newTkpiItem.riboflavin || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, riboflavin: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">NIASIN (mg)</label>
                        <input
                          id="form-tkpi-niasin"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.niasin || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, niasin: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-slate-500 block text-[10px]">VIT_C (mg)</label>
                        <input
                          id="form-tkpi-vit_c"
                          type="number"
                          step="0.1"
                          value={newTkpiItem.vit_c || 0}
                          onChange={(e) => setNewTkpiItem({ ...newTkpiItem, vit_c: Number(e.target.value) })}
                          className="w-full px-1.5 py-1 border border-slate-200 rounded bg-white font-mono text-right"
                        />
                      </div>
                    </div>
                  </div>

                  <button
                    id="btn-form-tkpi-submit"
                    type="submit"
                    className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold transition text-xs shadow-md mt-6"
                  >
                    Tambah Bahan (Format Tabel TKPI)
                  </button>
                </form>
              </div>

              {/* List of items */}
              <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden xl:col-span-8 flex flex-col h-[750px]">
                <div className="p-4 bg-slate-50/50 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <div className="space-y-1">
                    <h4 className="font-bold text-slate-800 text-sm">
                      Daftar Bahan Terdaftar ({filteredTkpiList.length} Bahan)
                    </h4>
                    <span className="text-[10px] text-slate-400 italic font-medium">
                      Bahan kustom memiliki tombol hapus. Geser tabel ke kanan untuk melihat nutrisi lengkap.
                    </span>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <div className="relative w-full sm:w-60">
                      <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                      <input
                        id="search-tkpi-input"
                        type="text"
                        placeholder="Cari nama, kategori, sumber..."
                        value={tkpiSearchQuery}
                        onChange={(e) => setTkpiSearchQuery(e.target.value)}
                        className="w-full pl-8 pr-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs text-slate-700 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-rose-500"
                      />
                    </div>
                  </div>
                </div>

                <div className="overflow-x-auto flex-1">
                  <table className="w-full text-left border-collapse text-xs table-fixed min-w-[2160px]">
                    <thead className="bg-slate-100 text-slate-700 font-bold uppercase tracking-wider text-[10px] sticky top-0 z-10 shadow-xs">
                      <tr>
                        <th className="p-3 w-14 text-center">NO</th>
                        <th className="p-3 w-40">KATEGORI</th>
                        <th className="p-3 w-56">NAMA BAHAN</th>
                        <th className="p-3 w-36">SUMBER</th>
                        <th className="p-3 w-28 text-center">BERAT STANDAR</th>
                        <th className="p-3 w-24 text-center">% BDD</th>
                        <th className="p-3 w-24 text-center">ENERGI</th>
                        <th className="p-3 w-24 text-center">PROTEIN</th>
                        <th className="p-3 w-24 text-center">LEMAK</th>
                        <th className="p-3 w-20 text-center">KH</th>
                        <th className="p-3 w-24 text-center">SERAT</th>
                        <th className="p-3 w-20 text-center">ABU</th>
                        <th className="p-3 w-24 text-center">KALSIUM</th>
                        <th className="p-3 w-24 text-center">FOSFOR</th>
                        <th className="p-3 w-20 text-center">BESI</th>
                        <th className="p-3 w-24 text-center">NATRIUM</th>
                        <th className="p-3 w-24 text-center">KALIUM</th>
                        <th className="p-3 w-24 text-center">TEMBAGA</th>
                        <th className="p-3 w-20 text-center">SENG</th>
                        <th className="p-3 w-24 text-center">RETINOL</th>
                        <th className="p-3 w-24 text-center">BKAR</th>
                        <th className="p-3 w-24 text-center">KARTOTAL</th>
                        <th className="p-3 w-24 text-center">THIAMIN</th>
                        <th className="p-3 w-24 text-center">RIBOFLAVIN</th>
                        <th className="p-3 w-24 text-center">NIASIN</th>
                        <th className="p-3 w-24 text-center">VIT_C</th>
                        <th className="p-3 w-24 text-center">AIR</th>
                        <th className="p-3 w-20 text-center">AKSI</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-700">
                      {paginatedTkpiList.length > 0 ? (
                        paginatedTkpiList.map((t, idx) => {
                          const globalIdx = (tkpiPage - 1) * tkpiItemsPerPage + idx + 1;
                          return (
                            <tr key={t.id} className="hover:bg-slate-50/50">
                              <td className="p-2.5 text-center text-slate-500 font-mono tabular-nums">{t.no ?? globalIdx}</td>
                              <td className="p-2.5 truncate font-semibold text-rose-800 bg-rose-50/10">
                                {t.kategori || t.sumber || "Umum"}
                              </td>
                              <td className="p-2.5 font-bold text-slate-800 truncate" title={t.nama}>{t.nama}</td>
                              <td className="p-2.5 truncate text-slate-500 font-medium">{t.sumber}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-indigo-700 font-medium">
                                {t.beratStandar !== undefined ? t.beratStandar : 100} g
                              </td>
                              <td className="p-2.5 text-center font-mono tabular-nums font-bold text-slate-600">{t.bdd}%</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-amber-700 font-extrabold">{(t.energi || 0).toFixed(0)}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-indigo-700 font-bold">{(t.protein || 0).toFixed(1)}g</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-rose-700 font-bold">{(t.lemak || 0).toFixed(1)}g</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-cyan-700 font-bold">{(t.kh || 0).toFixed(1)}g</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-emerald-700 font-bold">{(t.serat || 0).toFixed(1)}g</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-slate-500">{t.abu !== undefined ? t.abu.toFixed(1) : "0.0"}g</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-teal-700 font-semibold">{t.ca ?? 0}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-purple-700">{t.p ?? 0}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-pink-700 font-semibold">{t.fe !== undefined ? t.fe.toFixed(1) : "0.0"}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-blue-700">{t.na ?? 0}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-orange-700">{t.k ?? 0}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-yellow-800">{t.cu !== undefined ? t.cu.toFixed(2) : "0.00"}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-violet-800">{t.zn !== undefined ? t.zn.toFixed(1) : "0.0"}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-red-600">{t.retinol ?? 0}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-orange-600">{t.b_karoten ?? 0}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-amber-700">{t.kartotal ?? 0}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-amber-600 font-medium">{t.thiamin !== undefined ? t.thiamin.toFixed(2) : "0.00"}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-lime-600 font-medium">{t.riboflavin !== undefined ? t.riboflavin.toFixed(2) : "0.00"}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-indigo-500">{t.niasin !== undefined ? t.niasin.toFixed(1) : "0.0"}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-red-500 font-semibold">{t.vit_c !== undefined ? t.vit_c.toFixed(1) : "0.0"}</td>
                              <td className="p-2.5 text-center font-mono tabular-nums text-sky-600 font-semibold bg-sky-50/10">
                                {t.air !== undefined ? t.air.toFixed(1) : "0.0"}g
                              </td>
                              <td className="p-2.5 text-center">
                                <button
                                  id={`btn-del-tkpi-${t.id}`}
                                  type="button"
                                  onClick={() => deleteTkpiItem(t.id)}
                                  className="p-1 hover:bg-rose-50 text-slate-400 hover:text-rose-600 rounded transition inline-flex items-center justify-center cursor-pointer"
                                  title="Hapus bahan makanan ini"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={28} className="p-8 text-center text-slate-400 italic">
                            Tidak ada bahan makanan yang sesuai dengan kata kunci atau filter.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Pagination Controls */}
                <div className="p-3 bg-slate-50 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-600 shrink-0 select-none">
                  <div className="flex items-center gap-2">
                    <span>Tampilkan per halaman:</span>
                    <select
                      value={tkpiItemsPerPage}
                      onChange={(e) => setTkpiItemsPerPage(Number(e.target.value))}
                      className="bg-white border border-slate-200 rounded px-2 py-1 text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-rose-500"
                    >
                      <option value={25}>25</option>
                      <option value={50}>50</option>
                      <option value={100}>100</option>
                      <option value={250}>250</option>
                    </select>
                    <span className="text-slate-400">
                      • Menampilkan {filteredTkpiList.length > 0 ? (tkpiPage - 1) * tkpiItemsPerPage + 1 : 0} - {Math.min(tkpiPage * tkpiItemsPerPage, filteredTkpiList.length)} dari {filteredTkpiList.length} bahan
                    </span>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      disabled={tkpiPage <= 1}
                      onClick={() => setTkpiPage((p) => Math.max(1, p - 1))}
                      className="px-3 py-1 bg-white border border-slate-200 rounded text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100 font-semibold transition"
                    >
                      &laquo; Sebelumnya
                    </button>
                    <span className="font-bold text-slate-800">
                      Halaman {tkpiPage} dari {totalTkpiPages}
                    </span>
                    <button
                      type="button"
                      disabled={tkpiPage >= totalTkpiPages}
                      onClick={() => setTkpiPage((p) => Math.min(totalTkpiPages, p + 1))}
                      className="px-3 py-1 bg-white border border-slate-200 rounded text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100 font-semibold transition"
                    >
                      Berikutnya &raquo;
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

            {activeTab === "handbook" && (
              <HandbookTab />
            )}
          </>
        )}
      </main>

      {/* FOOTER */}
      <footer className="bg-white border-t border-slate-200 text-slate-600 text-xs py-4 px-6 mt-auto shrink-0 flex flex-col md:flex-row items-center justify-between gap-4 max-w-full shadow-[0_-4px_12px_-4px_rgba(0,0,0,0.05)] select-none">
        <div className="flex flex-col items-center md:items-start gap-1">
          <div className="flex items-center gap-1.5 text-slate-800 text-xs font-bold">
            <span className="font-extrabold text-slate-950">Syukri Odhe | Ahli Gizi</span>
            <span className="text-slate-300">|</span>
            <span className="text-slate-600 font-medium">Sistem Perencanaan Menu Makan Bergizi Gratis</span>
          </div>
          <span className="italic text-[10px] text-slate-400">
            Dirancang untuk {profile.namaLembaga} • Juknis BGN No. 401.1/2025
          </span>
        </div>
      </footer>

      {/* AUTH ERROR WARNING MODAL */}
      {authError && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-[9999] p-4 font-sans select-none">
          <div className="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-rose-100/60 animate-in fade-in zoom-in duration-200 text-left">
            <div className="flex items-center gap-3 text-rose-600 mb-4">
              <div className="p-3 bg-rose-50 rounded-2xl text-rose-500 shrink-0">
                <AlertCircle className="w-6 h-6" />
              </div>
              <div>
                <h3 className="font-extrabold text-slate-900 text-base leading-snug">
                  {authError === "UNAUTHORIZED_DOMAIN"
                    ? "Domain Belum Terdaftar di Firebase"
                    : "Otentikasi Google Terkendala"}
                </h3>
                <p className="text-[10px] text-rose-500 font-extrabold tracking-wider uppercase">
                  {authError === "UNAUTHORIZED_DOMAIN"
                    ? "auth/unauthorized-domain"
                    : "Kebijakan Proteksi Iframe Browser"}
                </p>
              </div>
            </div>

            {authError === "UNAUTHORIZED_DOMAIN" ? (
              <div className="space-y-4 text-slate-600 text-xs leading-relaxed">
                <p>
                  Domain <strong>{domainAuthInfo.currentHostname || (typeof window !== "undefined" ? window.location.hostname : "")}</strong> belum terdaftar pada <strong>Authorized domains</strong> Firebase Authentication.
                </p>
                <div className="bg-amber-50 border border-amber-200 p-3.5 rounded-2xl text-amber-950 space-y-2.5">
                  <p className="text-[11px] leading-snug">
                    Salin domain di bawah lalu tambahkan di Firebase Console → Authentication → Settings → <strong>Authorized domains</strong>.
                  </p>
                  <div className="flex items-center justify-between gap-2 bg-white border border-amber-200 rounded-xl px-2.5 py-2">
                    <code className="text-[11px] font-bold text-slate-900 break-all">
                      {domainAuthInfo.currentHostname || (typeof window !== "undefined" ? window.location.hostname : "")}
                    </code>
                    <button
                      type="button"
                      onClick={() =>
                        handleCopyCurrentDomain(
                          domainAuthInfo.currentHostname ||
                            (typeof window !== "undefined" ? window.location.hostname : "")
                        )
                      }
                      className="shrink-0 px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-lg text-[11px] font-bold cursor-pointer"
                    >
                      {domainCopied ? "Tersalin" : "Salin"}
                    </button>
                  </div>
                  <a
                    href={FIREBASE_AUTH_SETTINGS_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="w-full py-2.5 px-3 bg-[#112461] hover:bg-[#193382] text-white text-xs font-bold rounded-xl transition shadow-sm flex items-center justify-center gap-2"
                  >
                    <ExternalLink className="w-3.5 h-3.5" />
                    <span>Buka Authorized Domains Firebase</span>
                  </a>
                </div>
              </div>
            ) : (
              <div className="space-y-4 text-slate-600 text-xs leading-relaxed">
                <p>
                  {authError === "IFRAME_BLOCKED"
                    ? "Browser Anda memblokir Google Sign-In pop-up karena aplikasi berjalan di dalam jendela pratinjau (iframe). Buka di tab baru agar masuk dengan email Google berjalan lancar."
                    : authError}
                </p>

                <div className="bg-blue-50/80 border border-blue-100 p-3.5 rounded-2xl text-slate-700 space-y-2.5">
                  <p className="font-bold text-blue-800 flex items-center gap-1.5">
                    <span className="inline-block w-1.5 h-1.5 rounded-full bg-blue-600"></span>
                    Solusi Utama (Direkomendasikan):
                  </p>
                  <p className="text-[11px] text-slate-600 pl-3 leading-snug">
                    Buka aplikasi SISPERMBG di tab browser terpisah lalu masuk menggunakan akun email Google masing-masing.
                  </p>
                  <a
                    href={typeof window !== "undefined" ? window.location.href : "#"}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="w-full py-2.5 px-3 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl transition shadow-sm flex items-center justify-center gap-2"
                  >
                    <span>Buka Aplikasi di Tab Baru</span>
                  </a>
                </div>
              </div>
            )}

            <div className="mt-5 pt-3 border-t border-slate-100 flex justify-end gap-2.5">
              <button
                type="button"
                onClick={() => {
                  if (authError === "UNAUTHORIZED_DOMAIN") {
                    handleRecheckDomainClick();
                  } else {
                    setAuthError(null);
                    signInWithGoogle();
                  }
                }}
                className="py-2 px-4 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl transition cursor-pointer"
              >
                {authError === "UNAUTHORIZED_DOMAIN" ? "Cek Ulang & Masuk" : "Coba Masuk Google Lagi"}
              </button>
              <button
                type="button"
                onClick={() => setAuthError(null)}
                className="py-2 px-4 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-xl transition cursor-pointer"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL PENYESUAIAN & PEMOTONGAN LOGO / FOTO PROFIL */}
      {isCropModalOpen && cropImageSrc && (
        <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm flex items-center justify-center z-[9999] p-4 font-sans select-none animate-in fade-in duration-200">
          <div className="bg-white rounded-3xl max-w-md w-full shadow-2xl border border-slate-100 overflow-hidden flex flex-col animate-in zoom-in-95 duration-200">
            {/* Header */}
            <div className="px-6 py-4 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Camera className="w-5 h-5 text-emerald-600" />
                <div>
                  <h3 className="font-extrabold text-slate-800 text-sm">Sesuaikan Foto Profil</h3>
                  <p className="text-[10px] text-slate-400">Atur ukuran & letak logo SPPG</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsCropModalOpen(false)}
                className="p-1.5 hover:bg-slate-200/60 rounded-full text-slate-400 hover:text-slate-600 transition cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Content */}
            <div className="p-6 flex flex-col items-center gap-6">
              {/* Circular Viewport for Cropping */}
              <div
                className="relative w-[240px] h-[240px] rounded-full border-2 border-dashed border-emerald-500 overflow-hidden bg-slate-950 flex items-center justify-center cursor-move shadow-inner"
                onMouseDown={handleCropMouseDown}
                onMouseMove={handleCropMouseMove}
                onMouseUp={handleCropMouseUp}
                onMouseLeave={handleCropMouseLeave}
                onTouchStart={handleCropTouchStart}
                onTouchMove={handleCropTouchMove}
                onTouchEnd={handleCropTouchEnd}
                title="Seret gambar untuk memposisikan"
              >
                <img
                  src={cropImageSrc}
                  alt="Crop preview"
                  className="max-w-none absolute select-none pointer-events-none"
                  style={{
                    width: cropImageRatio > 1 ? `${240 * cropImageRatio}px` : "240px",
                    height: cropImageRatio > 1 ? "240px" : `${240 / cropImageRatio}px`,
                    transform: `translate(${cropOffsetX}px, ${cropOffsetY}px) scale(${cropZoom})`,
                    transformOrigin: "center center",
                  }}
                />
                
                {/* Visual grid guide lines for professional look */}
                <div className="absolute inset-0 pointer-events-none border border-white/10 rounded-full flex items-center justify-center">
                  <div className="w-1/3 h-full border-l border-r border-white/5"></div>
                  <div className="h-1/3 w-full border-t border-b border-white/5 absolute"></div>
                </div>

                {/* Dark circular overlay for non-selected region */}
                <div className="absolute inset-0 pointer-events-none rounded-full shadow-[0_0_0_9999px_rgba(15,23,42,0.45)]"></div>
              </div>

              {/* Instructions */}
              <div className="text-center space-y-1">
                <p className="text-xs font-bold text-slate-700">Gunakan Mouse / Sentuhan untuk Menggeser</p>
                <p className="text-[10px] text-slate-400">Seret gambar tepat ke posisi tengah lingkaran crop.</p>
              </div>

              {/* Interactive Sliders */}
              <div className="w-full space-y-4 pt-2 border-t border-slate-100">
                {/* Zoom control */}
                <div className="space-y-1">
                  <div className="flex justify-between items-center text-xs">
                    <span className="font-bold text-slate-600">Perbesaran (Zoom):</span>
                    <span className="font-mono bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded text-[10px] font-bold">
                      {cropZoom.toFixed(2)}x
                    </span>
                  </div>
                  <div className="flex items-center gap-3">
                    <button
                      type="button"
                      onClick={() => setCropZoom(prev => Math.max(1, prev - 0.1))}
                      className="text-xs font-bold w-6 h-6 rounded-full bg-slate-100 hover:bg-slate-200 flex items-center justify-center text-slate-600 transition cursor-pointer"
                    >
                      -
                    </button>
                    <input
                      type="range"
                      min="1"
                      max="3"
                      step="0.01"
                      value={cropZoom}
                      onChange={(e) => setCropZoom(parseFloat(e.target.value))}
                      className="flex-grow h-1.5 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-emerald-600"
                    />
                    <button
                      type="button"
                      onClick={() => setCropZoom(prev => Math.min(3, prev + 0.1))}
                      className="text-xs font-bold w-6 h-6 rounded-full bg-slate-100 hover:bg-slate-200 flex items-center justify-center text-slate-600 transition cursor-pointer"
                    >
                      +
                    </button>
                  </div>
                </div>

                {/* Fine-tune X-axis alignment */}
                <div className="space-y-1">
                  <div className="flex justify-between items-center text-xs">
                    <span className="font-bold text-slate-600">Geser Horizontal (X):</span>
                    <span className="font-mono bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded text-[10px] font-bold">
                      {cropOffsetX > 0 ? `+${cropOffsetX}` : cropOffsetX}px
                    </span>
                  </div>
                  <div className="flex items-center gap-3">
                    <button
                      type="button"
                      onClick={() => setCropOffsetX(prev => prev - 5)}
                      className="text-xs font-bold w-6 h-6 rounded-full bg-slate-100 hover:bg-slate-200 flex items-center justify-center text-slate-600 transition cursor-pointer"
                    >
                      ←
                    </button>
                    <input
                      type="range"
                      min="-200"
                      max="200"
                      value={cropOffsetX}
                      onChange={(e) => setCropOffsetX(parseInt(e.target.value))}
                      className="flex-grow h-1.5 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-slate-600"
                    />
                    <button
                      type="button"
                      onClick={() => setCropOffsetX(prev => prev + 5)}
                      className="text-xs font-bold w-6 h-6 rounded-full bg-slate-100 hover:bg-slate-200 flex items-center justify-center text-slate-600 transition cursor-pointer"
                    >
                      →
                    </button>
                  </div>
                </div>

                {/* Fine-tune Y-axis alignment */}
                <div className="space-y-1">
                  <div className="flex justify-between items-center text-xs">
                    <span className="font-bold text-slate-600">Geser Vertikal (Y):</span>
                    <span className="font-mono bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded text-[10px] font-bold">
                      {cropOffsetY > 0 ? `+${cropOffsetY}` : cropOffsetY}px
                    </span>
                  </div>
                  <div className="flex items-center gap-3">
                    <button
                      type="button"
                      onClick={() => setCropOffsetY(prev => prev - 5)}
                      className="text-xs font-bold w-6 h-6 rounded-full bg-slate-100 hover:bg-slate-200 flex items-center justify-center text-slate-600 transition cursor-pointer"
                    >
                      ↑
                    </button>
                    <input
                      type="range"
                      min="-200"
                      max="200"
                      value={cropOffsetY}
                      onChange={(e) => setCropOffsetY(parseInt(e.target.value))}
                      className="flex-grow h-1.5 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-slate-600"
                    />
                    <button
                      type="button"
                      onClick={() => setCropOffsetY(prev => prev + 5)}
                      className="text-xs font-bold w-6 h-6 rounded-full bg-slate-100 hover:bg-slate-200 flex items-center justify-center text-slate-600 transition cursor-pointer"
                    >
                      ↓
                    </button>
                  </div>
                </div>
              </div>
            </div>

            {/* Footer */}
            <div className="px-6 py-4 border-t border-slate-100 bg-slate-50 flex justify-end gap-3 shrink-0">
              <button
                type="button"
                onClick={() => setIsCropModalOpen(false)}
                className="px-4 py-2 bg-white hover:bg-slate-100 text-slate-600 text-xs font-bold rounded-xl border border-slate-200 transition cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleSaveCroppedLogo}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black rounded-xl shadow-md transition cursor-pointer flex items-center gap-1.5"
              >
                <Camera className="w-4 h-4" />
                <span>Simpan & Terapkan</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
