import React, { useState, useEffect, useRef, useCallback } from "react";
import { doc, setDoc, serverTimestamp } from "firebase/firestore";
import {
  FileSpreadsheet,
  Link2,
  RefreshCw,
  UploadCloud,
  DownloadCloud,
  ExternalLink,
  CheckCircle2,
  AlertCircle,
  Loader2,
  LogOut,
  Zap,
  ShieldCheck,
  X,
} from "lucide-react";
import { db, reportFirestoreWriteResult } from "../lib/firebase.ts";
import { syncUserToRealtimeHub } from "../context/AuthContext.tsx";
import {
  CANONICAL_SHEET_HEADERS,
  SheetSyncConfig,
  SheetTabMeta,
  SheetSyncPlan,
  SheetsSessionInfo,
  ParsedSheetUserRow,
  getSheetsSession,
  clearSheetsSession,
  connectGoogleSheets,
  parseGoogleSheetUrl,
  buildGoogleSheetUrl,
  fetchSpreadsheetMetadata,
  readUsersFromGoogleSheet,
  prepareSheetSyncPlan,
  syncUsersToGoogleSheet,
  loadGlobalSheetSyncConfig,
  saveGlobalSheetSyncConfig,
  cacheSheetUsersOnServer,
  readPublicSpreadsheetViaServer,
} from "../lib/googleSheetsSync.ts";

interface SpreadsheetUserSyncCardProps {
  users: any[];
  onUsersImportedFromSheet: (importedUsers: any[]) => void;
}

const fetchWithTimeout = <T,>(promise: Promise<T>, timeoutMs: number, defaultValue: T): Promise<T> => {
  return Promise.race([
    promise,
    new Promise<T>((resolve) => setTimeout(() => resolve(defaultValue), timeoutMs)),
  ]);
};

export default function SpreadsheetUserSyncCard({
  users,
  onUsersImportedFromSheet,
}: SpreadsheetUserSyncCardProps) {
  const [config, setConfig] = useState<SheetSyncConfig | null>(null);
  const [urlInput, setUrlInput] = useState("");
  const [session, setSession] = useState<SheetsSessionInfo | null>(() => getSheetsSession());
  const [tabs, setTabs] = useState<SheetTabMeta[]>([]);
  const [selectedTab, setSelectedTab] = useState<string>("");
  const [spreadsheetTitle, setSpreadsheetTitle] = useState<string>("");

  const [isConnecting, setIsConnecting] = useState(false);
  const [isSavingLink, setIsSavingLink] = useState(false);
  const [isPulling, setIsPulling] = useState(false);
  const [isPreparingPush, setIsPreparingPush] = useState(false);
  const [isPushing, setIsPushing] = useState(false);

  // Explicit confirmation modals for mutating Google Sheets operations
  const [pendingSyncPlan, setPendingSyncPlan] = useState<SheetSyncPlan | null>(null);
  const [isAutoSyncConfirmOpen, setIsAutoSyncConfirmOpen] = useState(false);
  const [autoSyncEnabled, setAutoSyncEnabled] = useState(false);

  const [statusFeedback, setStatusFeedback] = useState<{
    type: "success" | "info" | "error";
    text: string;
  } | null>(null);
  const [lastSyncAt, setLastSyncAt] = useState<string | null>(null);

  const usersRef = useRef<any[]>(users);
  usersRef.current = users;
  const lastPushedSignatureRef = useRef<string>("");
  const autoPushInFlightRef = useRef<boolean>(false);

  // Load saved global Spreadsheet link on mount
  useEffect(() => {
    let mounted = true;
    loadGlobalSheetSyncConfig()
      .then((saved) => {
        if (!mounted || !saved) return;
        setConfig(saved);
        setUrlInput(saved.sheetUrl || buildGoogleSheetUrl(saved.spreadsheetId, saved.gid));
        if (saved.sheetTitle) setSelectedTab(saved.sheetTitle);
        if (saved.spreadsheetTitle) setSpreadsheetTitle(saved.spreadsheetTitle);
      })
      .catch(() => {});
    return () => {
      mounted = false;
    };
  }, []);

  // Load sheet tab metadata when both session & spreadsheetId exist
  const loadTabsForSpreadsheet = useCallback(
    async (activeToken: string, spreadsheetId: string, preferredGid?: string, preferredTitle?: string) => {
      const meta = await fetchSpreadsheetMetadata(activeToken, spreadsheetId);
      setSpreadsheetTitle(meta.title);
      setTabs(meta.sheets);

      let chosen = "";
      if (preferredTitle && meta.sheets.some((s) => s.title === preferredTitle)) {
        chosen = preferredTitle;
      } else if (preferredGid) {
        const gidMatch = meta.sheets.find((s) => String(s.sheetId) === String(preferredGid));
        if (gidMatch) chosen = gidMatch.title;
      }
      if (!chosen && meta.sheets.length > 0) {
        chosen = meta.sheets[0].title;
      }
      setSelectedTab(chosen);
      return { title: meta.title, chosenTab: chosen, sheets: meta.sheets };
    },
    []
  );

  useEffect(() => {
    if (!session?.accessToken || !config?.spreadsheetId) return;
    loadTabsForSpreadsheet(session.accessToken, config.spreadsheetId, config.gid, config.sheetTitle).catch(
      (err) => {
        console.warn("Notice loading sheet metadata:", err);
      }
    );
  }, [session?.accessToken, config?.spreadsheetId, config?.gid, config?.sheetTitle, loadTabsForSpreadsheet]);

  // Persist imported Spreadsheet rows into Panel Admin + Express Real-Time Hub + Firestore
  const applyImportedSheetUsers = useCallback(
    async (sheetRows: ParsedSheetUserRow[], targetSpreadsheetId: string, targetTabTitle: string) => {
      const nowIso = new Date().toISOString();
      const existingByEmail = new Map<string, any>();
      usersRef.current.forEach((u) => {
        const em = (u.email || "").toLowerCase().trim();
        if (em) existingByEmail.set(em, u);
      });

      const importedObjects: any[] = [];
      let newlyRegisteredCount = 0;

      for (const row of sheetRows) {
        const email = row.email.toLowerCase().trim();
        const existing = existingByEmail.get(email);
        const sheetUid = `sheet_user_${email.replace(/[^a-z0-9]/g, "_")}`;
        const targetUid = existing?.uid || existing?.id || sheetUid;

        if (!existing) newlyRegisteredCount++;

        const mergedUser: Record<string, any> = {
          ...(existing || {}),
          uid: targetUid,
          id: targetUid,
          email,
          namaLengkap: existing?.namaLengkap || row.namaLengkap,
          profesi: existing?.profesi || row.profesi,
          namaSPPG: existing?.namaSPPG || row.namaSPPG,
          noHp: existing?.noHp || row.noHp,
          peran: existing?.peran === "ADMIN" ? "ADMIN" : row.peran,
          statusPersetujuan:
            !existing
              ? row.statusPersetujuan
              : existing.statusPersetujuan === "menunggu" && row.statusPersetujuan === "aktif"
              ? "aktif"
              : existing.statusPersetujuan || row.statusPersetujuan,
          berakhirPada: existing?.berakhirPada || row.berakhirPada,
          authProvider: existing?.authProvider || "google",
          emailVerified: true,
          registeredViaSpreadsheet: true,
          sumber: existing?.sumber || "Spreadsheet",
          createdAt: existing?.createdAt || nowIso,
          updatedAt: nowIso,
          loginTerakhir: existing?.loginTerakhir || row.loginTerakhir || null,
        };

        importedObjects.push(mergedUser);

        // Save locally + sync to Real-Time Hub + Firestore
        try {
          localStorage.setItem(`offline_user_${targetUid}`, JSON.stringify(mergedUser));
        } catch {}

        syncUserToRealtimeHub(mergedUser).catch(() => {});

        fetchWithTimeout(
          setDoc(
            doc(db, "users", targetUid),
            {
              ...mergedUser,
              createdAt: existing?.createdAt || serverTimestamp(),
              updatedAt: serverTimestamp(),
            },
            { merge: true }
          )
            .then(() => reportFirestoreWriteResult(true))
            .catch((err) => reportFirestoreWriteResult(false, err)),
          2500,
          null
        );
      }

      await cacheSheetUsersOnServer(sheetRows, targetSpreadsheetId, targetTabTitle);
      if (importedObjects.length > 0) {
        onUsersImportedFromSheet(importedObjects);
      }

      return { totalRead: sheetRows.length, newlyRegisteredCount };
    },
    [onUsersImportedFromSheet]
  );

  // Save Spreadsheet URL globally
  const handleSaveSheetLink = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    setStatusFeedback(null);
    const parsed = parseGoogleSheetUrl(urlInput);
    if (!parsed) {
      setStatusFeedback({
        type: "error",
        text: "Format link Google Spreadsheet tidak valid. Tempel URL lengkap (contoh: https://docs.google.com/spreadsheets/d/xxxx/edit).",
      });
      return;
    }

    setIsSavingLink(true);
    try {
      let detectedTitle = spreadsheetTitle || "";
      let detectedTab = selectedTab || "";

      if (session?.accessToken) {
        try {
          const meta = await loadTabsForSpreadsheet(
            session.accessToken,
            parsed.spreadsheetId,
            parsed.gid
          );
          detectedTitle = meta.title;
          detectedTab = meta.chosenTab;
        } catch (metaErr: any) {
          console.warn("Metadata load notice on save:", metaErr);
        }
      }

      const nextCfg: SheetSyncConfig = {
        sheetUrl: urlInput.trim(),
        spreadsheetId: parsed.spreadsheetId,
        gid: parsed.gid,
        sheetTitle: detectedTab,
        spreadsheetTitle: detectedTitle,
        updatedAt: new Date().toISOString(),
      };

      await saveGlobalSheetSyncConfig(nextCfg);
      setConfig(nextCfg);

      // Immediately try a read-only sync if sheet is accessible (OAuth or shared link)
      if (session?.accessToken && detectedTab) {
        const readRes = await readUsersFromGoogleSheet(
          session.accessToken,
          parsed.spreadsheetId,
          detectedTab
        );
        const imported = await applyImportedSheetUsers(
          readRes.users,
          parsed.spreadsheetId,
          detectedTab
        );
        setLastSyncAt(new Date().toISOString());
        setStatusFeedback({
          type: "success",
          text: `Link Spreadsheet "${detectedTitle || parsed.spreadsheetId}" tersimpan & ${imported.totalRead} baris pengguna berhasil ditarik (${imported.newlyRegisteredCount} baru terdaftar di Panel Admin).`,
        });
      } else {
        const pubRes = await readPublicSpreadsheetViaServer({
          spreadsheetId: parsed.spreadsheetId,
          gid: parsed.gid,
          sheetTitle: detectedTab,
        });
        if (pubRes.ok) {
          const imported = await applyImportedSheetUsers(
            pubRes.users,
            parsed.spreadsheetId,
            detectedTab || "Spreadsheet"
          );
          setLastSyncAt(new Date().toISOString());
          setStatusFeedback({
            type: "success",
            text: `Link Spreadsheet tersimpan! ${imported.totalRead} email terbaca dari spreadsheet (${imported.newlyRegisteredCount} baru otomatis terdaftar). Klik "Hubungkan Google Sheets" agar login baru juga otomatis ditulis ke spreadsheet.`,
          });
        } else {
          setStatusFeedback({
            type: "info",
            text: `Link Spreadsheet tersimpan. Klik "Hubungkan Google Sheets" untuk membaca spreadsheet privat dan menulis data pengguna yang login secara otomatis.`,
          });
        }
      }
    } catch (err: any) {
      setStatusFeedback({
        type: "error",
        text: err?.message || "Gagal menyimpan link Google Spreadsheet.",
      });
    } finally {
      setIsSavingLink(false);
    }
  };

  // Connect Admin Google account with Google Sheets permission
  const handleConnectSheets = async () => {
    setStatusFeedback(null);
    setIsConnecting(true);
    try {
      const nextSession = await connectGoogleSheets();
      setSession(nextSession);

      const parsedFromInput = parseGoogleSheetUrl(urlInput);
      const activeSpreadsheetId = config?.spreadsheetId || parsedFromInput?.spreadsheetId || "";
      const activeGid = config?.gid || parsedFromInput?.gid || "";

      if (activeSpreadsheetId) {
        const meta = await loadTabsForSpreadsheet(
          nextSession.accessToken,
          activeSpreadsheetId,
          activeGid,
          config?.sheetTitle
        );
        const updatedCfg: SheetSyncConfig = {
          sheetUrl:
            config?.sheetUrl || urlInput.trim() || buildGoogleSheetUrl(activeSpreadsheetId, activeGid),
          spreadsheetId: activeSpreadsheetId,
          gid: activeGid,
          sheetTitle: meta.chosenTab,
          spreadsheetTitle: meta.title,
          updatedAt: new Date().toISOString(),
          updatedBy: nextSession.email,
        };
        await saveGlobalSheetSyncConfig(updatedCfg);
        setConfig(updatedCfg);
        setStatusFeedback({
          type: "success",
          text: `Terhubung ke Google Sheets (${nextSession.email}) — Spreadsheet "${meta.title}" tab "${meta.chosenTab}" siap disinkronkan.`,
        });
      } else {
        setStatusFeedback({
          type: "success",
          text: `Google Sheets terhubung (${nextSession.email}). Silakan tempel link Google Spreadsheet Anda di kolom di atas lalu klik Simpan Link.`,
        });
      }
    } catch (err: any) {
      setStatusFeedback({
        type: "error",
        text: err?.message || "Gagal menghubungkan Google Sheets.",
      });
    } finally {
      setIsConnecting(false);
    }
  };

  const handleDisconnectSheets = () => {
    clearSheetsSession();
    setSession(null);
    setAutoSyncEnabled(false);
    setStatusFeedback({
      type: "info",
      text: "Koneksi sesi Google Sheets diputus dari memori browser. Tautan spreadsheet tetap tersimpan.",
    });
  };

  const handleTabChange = async (nextTabTitle: string) => {
    setSelectedTab(nextTabTitle);
    if (!config) return;
    const matchedTab = tabs.find((t) => t.title === nextTabTitle);
    const updatedCfg: SheetSyncConfig = {
      ...config,
      sheetTitle: nextTabTitle,
      gid: matchedTab ? String(matchedTab.sheetId) : config.gid,
      spreadsheetTitle: spreadsheetTitle || config.spreadsheetTitle,
      updatedAt: new Date().toISOString(),
    };
    setConfig(updatedCfg);
    await saveGlobalSheetSyncConfig(updatedCfg);
  };

  // Pull users added directly in Google Sheets -> Panel Admin + Firestore + Real-Time Hub
  const handlePullFromSheet = useCallback(
    async (silent = false) => {
      const activeSpreadsheetId = config?.spreadsheetId || parseGoogleSheetUrl(urlInput)?.spreadsheetId;
      if (!activeSpreadsheetId) {
        if (!silent) {
          setStatusFeedback({
            type: "error",
            text: "Tempel dan simpan link Google Spreadsheet terlebih dahulu.",
          });
        }
        return;
      }

      if (!silent) {
        setIsPulling(true);
        setStatusFeedback(null);
      }

      try {
        const activeSession = getSheetsSession();
        let activeTab = selectedTab || config?.sheetTitle || "";

        if (activeSession?.accessToken) {
          if (!activeTab) {
            const meta = await loadTabsForSpreadsheet(
              activeSession.accessToken,
              activeSpreadsheetId,
              config?.gid
            );
            activeTab = meta.chosenTab;
          }
          const readRes = await readUsersFromGoogleSheet(
            activeSession.accessToken,
            activeSpreadsheetId,
            activeTab
          );
          const imported = await applyImportedSheetUsers(
            readRes.users,
            activeSpreadsheetId,
            activeTab
          );
          setLastSyncAt(new Date().toISOString());
          if (!silent) {
            setStatusFeedback({
              type: "success",
              text: `Berhasil menarik ${imported.totalRead} pengguna dari tab "${activeTab}" (${imported.newlyRegisteredCount} email baru otomatis terdaftar di Panel Admin).`,
            });
          }
          return;
        }

        // Fallback for publicly shared spreadsheet without active OAuth token
        const pubRes = await readPublicSpreadsheetViaServer({
          spreadsheetId: activeSpreadsheetId,
          gid: config?.gid,
          sheetTitle: activeTab,
        });
        if (pubRes.ok) {
          const imported = await applyImportedSheetUsers(
            pubRes.users,
            activeSpreadsheetId,
            activeTab || "Spreadsheet"
          );
          setLastSyncAt(new Date().toISOString());
          if (!silent) {
            setStatusFeedback({
              type: "success",
              text: `Berhasil membaca ${imported.totalRead} pengguna dari spreadsheet (${imported.newlyRegisteredCount} baru otomatis terdaftar).`,
            });
          }
        } else if (!silent) {
          setStatusFeedback({
            type: "error",
            text:
              pubRes.error ||
              'Spreadsheet belum dapat dibaca. Klik "Hubungkan Google Sheets" di sebelah kanan untuk akses penuh.',
          });
        }
      } catch (err: any) {
        if (!silent) {
          setStatusFeedback({
            type: "error",
            text: err?.message || "Gagal menarik data dari Google Sheets.",
          });
        }
      } finally {
        if (!silent) setIsPulling(false);
      }
    },
    [applyImportedSheetUsers, config, loadTabsForSpreadsheet, selectedTab, urlInput]
  );

  // Step 1 of Push: Prepare exact diff preview & open mandatory confirmation modal
  const handlePreparePushToSheet = async () => {
    setStatusFeedback(null);
    const activeSession = getSheetsSession();
    if (!activeSession?.accessToken) {
      setStatusFeedback({
        type: "error",
        text: 'Klik "Hubungkan Google Sheets" terlebih dahulu agar aplikasi memiliki izin menulis data pengguna ke spreadsheet.',
      });
      return;
    }

    const activeSpreadsheetId = config?.spreadsheetId || parseGoogleSheetUrl(urlInput)?.spreadsheetId;
    if (!activeSpreadsheetId) {
      setStatusFeedback({
        type: "error",
        text: "Tempel dan simpan link Google Spreadsheet terlebih dahulu.",
      });
      return;
    }

    setIsPreparingPush(true);
    try {
      let activeTab = selectedTab || config?.sheetTitle || "";
      let activeTitle = spreadsheetTitle || config?.spreadsheetTitle || "Google Spreadsheet";

      if (!activeTab) {
        const meta = await loadTabsForSpreadsheet(
          activeSession.accessToken,
          activeSpreadsheetId,
          config?.gid
        );
        activeTab = meta.chosenTab;
        activeTitle = meta.title;
      }

      const plan = await prepareSheetSyncPlan(
        activeSession.accessToken,
        activeSpreadsheetId,
        activeTab,
        activeTitle,
        usersRef.current
      );
      setPendingSyncPlan(plan);
    } catch (err: any) {
      setStatusFeedback({
        type: "error",
        text: err?.message || "Gagal memeriksa perubahan spreadsheet.",
      });
    } finally {
      setIsPreparingPush(false);
    }
  };

  // Step 2 of Push: Execute write to Google Sheets ONLY after explicit user confirmation
  const handleConfirmPushToSheet = async () => {
    if (!pendingSyncPlan) return;
    const activeSession = getSheetsSession();
    const activeSpreadsheetId = config?.spreadsheetId || parseGoogleSheetUrl(urlInput)?.spreadsheetId;
    if (!activeSession?.accessToken || !activeSpreadsheetId) {
      setPendingSyncPlan(null);
      return;
    }

    setIsPushing(true);
    try {
      const result = await syncUsersToGoogleSheet(
        activeSession.accessToken,
        activeSpreadsheetId,
        pendingSyncPlan.sheetTitle,
        usersRef.current
      );
      setPendingSyncPlan(null);
      setLastSyncAt(result.syncedAt);
      setStatusFeedback({
        type: "success",
        text: `Sinkronisasi ke Google Sheets selesai: ${result.appendedCount} pengguna baru ditambahkan, ${result.updatedCount} baris diperbarui (total ${result.totalInSheet} baris di tab "${pendingSyncPlan.sheetTitle}").`,
      });
    } catch (err: any) {
      setPendingSyncPlan(null);
      setStatusFeedback({
        type: "error",
        text: err?.message || "Gagal menulis data ke Google Sheets.",
      });
    } finally {
      setIsPushing(false);
    }
  };

  // Toggle Two-Way Automatic Sync (requires explicit one-time confirmation modal before mutating Sheets)
  const handleToggleAutoSyncClick = () => {
    if (autoSyncEnabled) {
      setAutoSyncEnabled(false);
      setStatusFeedback({
        type: "info",
        text: "Sinkronisasi otomatis 2 arah ke Google Sheets dinonaktifkan.",
      });
      return;
    }

    if (!session?.accessToken) {
      setStatusFeedback({
        type: "error",
        text: 'Hubungkan akun Google Sheets terlebih dahulu sebelum mengaktifkan sinkronisasi otomatis.',
      });
      return;
    }
    if (!config?.spreadsheetId || !selectedTab) {
      setStatusFeedback({
        type: "error",
        text: "Simpan link Google Spreadsheet dan pilih tab terlebih dahulu.",
      });
      return;
    }

    setIsAutoSyncConfirmOpen(true);
  };

  const handleConfirmEnableAutoSync = () => {
    setIsAutoSyncConfirmOpen(false);
    setAutoSyncEnabled(true);
    setStatusFeedback({
      type: "success",
      text: `Sinkronisasi otomatis 2 arah AKTIF pada tab "${selectedTab}". Setiap email yang login, perubahan status ACC, maupun email baru yang ditambah langsung di spreadsheet akan otomatis tersinkronisasi.`,
    });
  };

  // Automatic Background Two-Way Sync Loop (runs ONLY after Admin explicitly confirms Auto-Sync)
  const usersSyncSignature = users
    .map(
      (u) =>
        `${(u.email || "").toLowerCase()}|${u.statusPersetujuan || ""}|${u.peran || ""}|${
          u.namaSPPG || ""
        }|${u.berakhirPada || ""}|${u.namaLengkap || ""}`
    )
    .sort()
    .join(";");

  useEffect(() => {
    if (!autoSyncEnabled || !session?.accessToken || !config?.spreadsheetId || !selectedTab) return;
    if (usersSyncSignature === lastPushedSignatureRef.current) return;

    const timer = window.setTimeout(async () => {
      if (autoPushInFlightRef.current) return;
      const activeSession = getSheetsSession();
      if (!activeSession?.accessToken) {
        setAutoSyncEnabled(false);
        return;
      }
      autoPushInFlightRef.current = true;
      try {
        const res = await syncUsersToGoogleSheet(
          activeSession.accessToken,
          config.spreadsheetId,
          selectedTab,
          usersRef.current
        );
        lastPushedSignatureRef.current = usersSyncSignature;
        setLastSyncAt(res.syncedAt);
      } catch (err: any) {
        console.warn("Auto-sync to Google Sheets notice:", err);
      } finally {
        autoPushInFlightRef.current = false;
      }
    }, 2500);

    return () => window.clearTimeout(timer);
  }, [autoSyncEnabled, session?.accessToken, config?.spreadsheetId, selectedTab, usersSyncSignature]);

  // Background 25-second pull loop when spreadsheet is linked (keeps rows added directly in Sheets live in Panel Admin!)
  useEffect(() => {
    if (!config?.spreadsheetId) return;
    if (!autoSyncEnabled && !session?.accessToken) return;

    const interval = window.setInterval(() => {
      if (document.visibilityState === "visible") {
        handlePullFromSheet(true).catch(() => {});
      }
    }, 25_000);

    return () => window.clearInterval(interval);
  }, [autoSyncEnabled, session?.accessToken, config?.spreadsheetId, handlePullFromSheet]);

  return (
    <div className="bg-white rounded-2xl border border-emerald-200/90 shadow-xs p-5 space-y-4">
      {/* Header Row */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
        <div className="flex items-start gap-3">
          <div className="p-2.5 bg-emerald-50 text-emerald-700 rounded-xl border border-emerald-200 shrink-0">
            <FileSpreadsheet className="w-5 h-5" />
          </div>
          <div>
            <div className="flex flex-wrap items-center gap-2">
              <h3 className="text-sm font-extrabold text-slate-900">
                Sinkronisasi Data Pengguna ↔ Google Spreadsheet
              </h3>
              {session ? (
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-700 border border-emerald-200 text-[10px] font-bold">
                  <CheckCircle2 className="w-3 h-3" />
                  Terhubung: {session.email}
                </span>
              ) : (
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-slate-100 text-slate-600 border border-slate-200 text-[10px] font-bold">
                  Belum Terhubung Google Sheets
                </span>
              )}
              {autoSyncEnabled && (
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-indigo-50 text-indigo-700 border border-indigo-200 text-[10px] font-bold">
                  <Zap className="w-3 h-3" />
                  Sinkronisasi Otomatis 2 Arah Aktif
                </span>
              )}
            </div>
            <p className="text-[11px] text-slate-500 mt-0.5 leading-relaxed">
              Setiap pengguna yang login menggunakan email otomatis disinkronkan ke spreadsheet. Email baru yang Anda tambahkan langsung di spreadsheet juga otomatis terdaftar di Panel Admin dan langsung dapat masuk dashboard.
            </p>
          </div>
        </div>

        {/* Official Google Connect / Disconnect Button */}
        <div className="flex flex-wrap items-center gap-2 shrink-0">
          {!session ? (
            <button
              type="button"
              onClick={handleConnectSheets}
              disabled={isConnecting}
              className="inline-flex items-center gap-2.5 px-4 py-2 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 rounded-xl text-xs font-bold shadow-xs transition cursor-pointer disabled:opacity-60"
            >
              {isConnecting ? (
                <Loader2 className="w-4 h-4 animate-spin text-emerald-600" />
              ) : (
                <svg className="w-4 h-4 shrink-0" viewBox="0 0 48 48" aria-hidden="true">
                  <path
                    fill="#EA4335"
                    d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"
                  />
                  <path
                    fill="#4285F4"
                    d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"
                  />
                  <path
                    fill="#FBBC05"
                    d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"
                  />
                  <path
                    fill="#34A853"
                    d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"
                  />
                </svg>
              )}
              <span>{isConnecting ? "Menghubungkan..." : "Hubungkan Google Sheets"}</span>
            </button>
          ) : (
            <button
              type="button"
              onClick={handleDisconnectSheets}
              className="inline-flex items-center gap-1.5 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition cursor-pointer"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span>Putus Sesi Sheets</span>
            </button>
          )}
        </div>
      </div>

      {/* Spreadsheet URL Input Form */}
      <form onSubmit={handleSaveSheetLink} className="grid grid-cols-1 lg:grid-cols-12 gap-2.5 items-center">
        <div className="lg:col-span-7 relative">
          <Link2 className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={urlInput}
            onChange={(e) => setUrlInput(e.target.value)}
            placeholder="Tempel link Google Spreadsheet Anda (https://docs.google.com/spreadsheets/d/.../edit)"
            className="w-full pl-9 pr-3 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/30 focus:border-emerald-500"
          />
        </div>

        <div className="lg:col-span-5 flex flex-wrap items-center gap-2">
          <button
            type="submit"
            disabled={isSavingLink || !urlInput.trim()}
            className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-[#112461] hover:bg-[#193382] text-white font-bold text-xs rounded-xl transition cursor-pointer disabled:opacity-50"
          >
            {isSavingLink ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <ShieldCheck className="w-3.5 h-3.5" />}
            <span>{isSavingLink ? "Menyimpan..." : "Simpan Link"}</span>
          </button>

          {config?.spreadsheetId && (
            <a
              href={config.sheetUrl || buildGoogleSheetUrl(config.spreadsheetId, config.gid)}
              target="_blank"
              rel="noopener noreferrer"
              className="inline-flex items-center gap-1.5 px-3 py-2.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 font-bold text-xs rounded-xl transition"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              <span>Buka Sheet</span>
            </a>
          )}
        </div>
      </form>

      {/* Active Sheet Tab Selector & Bidirectional Sync Controls */}
      {config?.spreadsheetId && (
        <div className="bg-slate-50 border border-slate-200/80 rounded-xl p-3.5 flex flex-col xl:flex-row xl:items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-3">
            {spreadsheetTitle && (
              <div className="text-xs">
                <span className="text-slate-400 font-semibold">File: </span>
                <span className="font-bold text-slate-800">{spreadsheetTitle}</span>
              </div>
            )}

            {tabs.length > 0 && (
              <div className="flex items-center gap-2 text-xs">
                <label htmlFor="sheet-tab-select" className="text-slate-500 font-semibold">
                  Tab Sheet:
                </label>
                <select
                  id="sheet-tab-select"
                  value={selectedTab}
                  onChange={(e) => handleTabChange(e.target.value)}
                  className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-500"
                >
                  {tabs.map((t) => (
                    <option key={t.sheetId} value={t.title}>
                      {t.title}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {lastSyncAt && (
              <span className="text-[11px] text-slate-500">
                Sinkron terakhir:{" "}
                <strong className="text-slate-700">
                  {new Date(lastSyncAt).toLocaleTimeString("id-ID")}
                </strong>
              </span>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => handlePullFromSheet(false)}
              disabled={isPulling}
              className="inline-flex items-center gap-1.5 px-3 py-2 bg-white hover:bg-slate-100 text-slate-800 border border-slate-300 font-bold text-xs rounded-xl transition cursor-pointer disabled:opacity-50"
            >
              {isPulling ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-600" />
              ) : (
                <DownloadCloud className="w-3.5 h-3.5 text-emerald-600" />
              )}
              <span>{isPulling ? "Menarik..." : "Tarik dari Spreadsheet"}</span>
            </button>

            <button
              type="button"
              onClick={handlePreparePushToSheet}
              disabled={isPreparingPush || !session}
              title={!session ? "Hubungkan Google Sheets terlebih dahulu" : "Tulis & perbarui data pengguna ke Spreadsheet"}
              className="inline-flex items-center gap-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl transition cursor-pointer disabled:opacity-50"
            >
              {isPreparingPush ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <UploadCloud className="w-3.5 h-3.5" />
              )}
              <span>{isPreparingPush ? "Memeriksa..." : "Sinkronkan ke Spreadsheet"}</span>
            </button>

            <button
              type="button"
              onClick={handleToggleAutoSyncClick}
              disabled={!session}
              className={`inline-flex items-center gap-1.5 px-3 py-2 font-bold text-xs rounded-xl border transition cursor-pointer disabled:opacity-50 ${
                autoSyncEnabled
                  ? "bg-indigo-600 hover:bg-indigo-700 text-white border-indigo-600"
                  : "bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border-indigo-200"
              }`}
            >
              <RefreshCw className={`w-3.5 h-3.5 ${autoSyncEnabled ? "animate-spin" : ""}`} />
              <span>{autoSyncEnabled ? "Auto-Sync: AKTIF" : "Aktifkan Auto-Sync"}</span>
            </button>
          </div>
        </div>
      )}

      {/* Compact Column Format Reference */}
      <div className="text-[11px] text-slate-500 bg-slate-50/70 border border-slate-200/60 rounded-xl px-3.5 py-2.5 leading-relaxed">
        <span className="font-bold text-slate-700">Format Kolom Otomatis:</span>{" "}
        <code className="text-[10px] bg-white px-1.5 py-0.5 rounded border border-slate-200 text-slate-700">
          {CANONICAL_SHEET_HEADERS.join(" | ")}
        </code>{" "}
        — Anda juga cukup menambahkan alamat <strong>Email</strong> di kolom Email spreadsheet (status kosong otomatis dianggap <strong>aktif</strong> sebagai pengguna panel/tenant, non-admin).
      </div>

      {/* Feedback Banner */}
      {statusFeedback && (
        <div
          className={`rounded-xl p-3 text-xs flex items-start justify-between gap-2 border ${
            statusFeedback.type === "success"
              ? "bg-emerald-50 border-emerald-200 text-emerald-900"
              : statusFeedback.type === "error"
              ? "bg-rose-50 border-rose-200 text-rose-900"
              : "bg-blue-50 border-blue-200 text-blue-900"
          }`}
        >
          <div className="flex items-start gap-2">
            {statusFeedback.type === "success" ? (
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
            ) : (
              <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
            )}
            <span className="font-medium leading-relaxed">{statusFeedback.text}</span>
          </div>
          <button
            type="button"
            onClick={() => setStatusFeedback(null)}
            className="text-current opacity-70 hover:opacity-100 cursor-pointer shrink-0"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* MANDATORY CONFIRMATION MODAL: Push / Update Spreadsheet */}
      {pendingSyncPlan && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-[9999] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-lg w-full shadow-2xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
                <div>
                  <h4 className="text-sm font-extrabold text-slate-900">
                    Konfirmasi Pembaruan Google Spreadsheet
                  </h4>
                  <p className="text-[11px] text-slate-500">
                    {pendingSyncPlan.spreadsheetTitle} • Tab "{pendingSyncPlan.sheetTitle}"
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setPendingSyncPlan(null)}
                className="p-1 text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-4 text-xs text-slate-700">
              <p className="leading-relaxed">
                Sistem akan memperbarui data pengguna pada spreadsheet Google Anda dengan rincian perubahan berikut:
              </p>

              <div className="grid grid-cols-3 gap-2.5 text-center">
                <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-2.5">
                  <div className="text-lg font-extrabold text-emerald-700 tabular-nums">
                    {pendingSyncPlan.toAppend.length}
                  </div>
                  <div className="text-[10px] font-bold text-emerald-800 uppercase">Baris Baru</div>
                </div>
                <div className="bg-blue-50 border border-blue-200 rounded-xl p-2.5">
                  <div className="text-lg font-extrabold text-blue-700 tabular-nums">
                    {pendingSyncPlan.toUpdate.length}
                  </div>
                  <div className="text-[10px] font-bold text-blue-800 uppercase">Baris Diperbarui</div>
                </div>
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-2.5">
                  <div className="text-lg font-extrabold text-slate-700 tabular-nums">
                    {pendingSyncPlan.unchangedCount}
                  </div>
                  <div className="text-[10px] font-bold text-slate-600 uppercase">Sudah Sama</div>
                </div>
              </div>

              {(pendingSyncPlan.toAppend.length > 0 || pendingSyncPlan.toUpdate.length > 0) ? (
                <div className="max-h-40 overflow-y-auto bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-1.5 text-[11px]">
                  {pendingSyncPlan.toAppend.slice(0, 8).map((item) => (
                    <div key={`add-${item.email}`} className="flex items-center justify-between gap-2">
                      <span className="truncate font-semibold text-emerald-800">+ {item.email}</span>
                      <span className="shrink-0 text-[10px] font-bold uppercase text-slate-500">
                        {item.statusPersetujuan}
                      </span>
                    </div>
                  ))}
                  {pendingSyncPlan.toUpdate.slice(0, 8).map((item) => (
                    <div key={`upd-${item.email}`} className="flex items-center justify-between gap-2">
                      <span className="truncate font-semibold text-blue-800">
                        ↻ Baris {item.rowNumber}: {item.email}
                      </span>
                      <span className="shrink-0 text-[10px] font-bold uppercase text-slate-500">
                        {item.statusPersetujuan}
                      </span>
                    </div>
                  ))}
                  {pendingSyncPlan.toAppend.length + pendingSyncPlan.toUpdate.length > 16 && (
                    <div className="text-slate-400 italic pt-1">
                      ...dan {pendingSyncPlan.toAppend.length + pendingSyncPlan.toUpdate.length - 16} baris lainnya.
                    </div>
                  )}
                </div>
              ) : (
                <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 text-emerald-800 font-semibold">
                  Semua data email pengguna saat ini sudah sama dengan isi spreadsheet.
                </div>
              )}
            </div>

            <div className="px-5 py-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setPendingSyncPlan(null)}
                disabled={isPushing}
                className="px-4 py-2 bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 font-bold text-xs rounded-xl cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmPushToSheet}
                disabled={isPushing}
                className="inline-flex items-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl cursor-pointer disabled:opacity-60"
              >
                {isPushing ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <UploadCloud className="w-3.5 h-3.5" />}
                <span>{isPushing ? "Menulis ke Sheets..." : "Konfirmasi & Sinkronkan"}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MANDATORY CONFIRMATION MODAL: Enable Automatic Two-Way Sync */}
      {isAutoSyncConfirmOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-[9999] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full shadow-2xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Zap className="w-5 h-5 text-indigo-600" />
                <h4 className="text-sm font-extrabold text-slate-900">
                  Konfirmasi Sinkronisasi Otomatis 2 Arah
                </h4>
              </div>
              <button
                type="button"
                onClick={() => setIsAutoSyncConfirmOpen(false)}
                className="p-1 text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-3 text-xs text-slate-700 leading-relaxed">
              <p>
                Dengan mengaktifkan mode ini pada <strong>{spreadsheetTitle || "Spreadsheet Anda"}</strong> (tab{" "}
                <strong>"{selectedTab}"</strong>):
              </p>
              <ul className="list-disc pl-4 space-y-1.5 text-[11px] text-slate-600">
                <li>Setiap pengguna baru yang masuk dengan email dan perubahan status ACC di Panel Admin akan otomatis ditulis dan diperbarui di Google Spreadsheet.</li>
                <li>Email baru yang ditambahkan langsung di Google Spreadsheet akan ditarik otomatis ke Panel Admin setiap 25 detik dan langsung terdaftar saat pengguna tersebut login.</li>
              </ul>
            </div>

            <div className="px-5 py-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setIsAutoSyncConfirmOpen(false)}
                className="px-4 py-2 bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 font-bold text-xs rounded-xl cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmEnableAutoSync}
                className="inline-flex items-center gap-1.5 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl cursor-pointer"
              >
                <CheckCircle2 className="w-3.5 h-3.5" />
                <span>Ya, Aktifkan Auto-Sync</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
