import React from "react";

interface AppErrorBoundaryProps {
  children: React.ReactNode;
}

interface AppErrorBoundaryState {
  hasError: boolean;
  errorMessage: string;
}

const SAFE_RESET_PREFIXES = [
  "sisper_",
  "kop_",
  "offline_user_",
  "custom_logged_in_uid",
  "simulated_user_email",
  "sisper_demo_",
];

export class AppErrorBoundary extends React.Component<
  AppErrorBoundaryProps,
  AppErrorBoundaryState
> {
  declare props: Readonly<AppErrorBoundaryProps>;
  state: AppErrorBoundaryState = {
    hasError: false,
    errorMessage: "",
  };

  constructor(props: AppErrorBoundaryProps) {
    super(props);
  }

  static getDerivedStateFromError(error: unknown): AppErrorBoundaryState {
    const msg = error instanceof Error ? error.message : String(error || "Terjadi kendala saat memuat aplikasi.");
    return {
      hasError: true,
      errorMessage: msg,
    };
  }

  componentDidCatch(error: unknown, info: React.ErrorInfo) {
    console.error("[GiziSync ErrorBoundary]:", error, info?.componentStack);
  }

  handleReload = () => {
    window.location.reload();
  };

  handleResetLocalSession = () => {
    try {
      const keysToRemove: string[] = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (!key) continue;
        if (SAFE_RESET_PREFIXES.some((prefix) => key.startsWith(prefix))) {
          keysToRemove.push(key);
        }
      }
      keysToRemove.forEach((k) => {
        try {
          localStorage.removeItem(k);
        } catch {}
      });
    } catch {}
    try {
      sessionStorage.clear();
    } catch {}
    window.location.reload();
  };

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-slate-950 text-slate-100 flex items-center justify-center p-6 font-sans">
          <div className="w-full max-w-md bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-2xl space-y-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-amber-500/15 border border-amber-400/30 flex items-center justify-center text-amber-400 font-bold text-lg shrink-0">
                !
              </div>
              <div>
                <h1 className="text-base font-bold text-white leading-tight">
                  Pemulihan Sistem GiziSync
                </h1>
                <p className="text-xs text-slate-400 mt-0.5">
                  Terdeteksi gangguan sementara saat memuat tampilan aplikasi.
                </p>
              </div>
            </div>

            {this.state.errorMessage && (
              <div className="bg-slate-950/80 border border-slate-800 rounded-xl p-3 text-[11px] text-slate-300 font-mono break-words max-h-24 overflow-auto">
                {this.state.errorMessage}
              </div>
            )}

            <p className="text-xs text-slate-400 leading-relaxed">
              Data lembar kerja Anda yang sudah tersimpan di cloud tetap aman. Silakan muat ulang halaman atau bersihkan cache sesi lokal di bawah ini.
            </p>

            <div className="flex flex-col sm:flex-row gap-2.5 pt-1">
              <button
                type="button"
                onClick={this.handleReload}
                className="flex-1 px-4 py-2.5 bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold rounded-xl text-xs transition cursor-pointer"
              >
                Muat Ulang Aplikasi
              </button>
              <button
                type="button"
                onClick={this.handleResetLocalSession}
                className="flex-1 px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-200 font-semibold rounded-xl text-xs border border-slate-700 transition cursor-pointer"
              >
                Bersihkan Cache Sesi
              </button>
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}
