"use client";

import { ChangeEvent, FormEvent, useEffect, useState } from "react";

type PreviewResponse = {
  sheet_used: string;
  available_sheets: string[];
  source_columns: string[];
  source_total_rows: number;
  source_preview: Record<string, string | number | null>[];
};

type SharepointDiagnosticResponse = {
  graph: {
    configured: boolean;
    token_ok: boolean;
    token_error?: string;
  };
  download_ok: boolean;
  filename?: string;
  bytes?: number;
  download_error?: string;
};

const API_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";

function DataTable({ title, rows, darkMode }: { title: string; rows: Record<string, string | number | null>[]; darkMode: boolean }) {
  if (!rows.length) {
    return (
      <section className="card p-6">
        <h2 className={`text-lg font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{title}</h2>
        <p className={`mt-3 text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Sin datos para mostrar.</p>
      </section>
    );
  }

  const headers = Object.keys(rows[0]);
  return (
    <section className="card p-6">
      <h2 className={`text-lg font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{title}</h2>
      <div className="mt-4 overflow-x-auto rounded-2xl border border-slate-200/10">
        <table className="min-w-full text-sm">
          <thead className={darkMode ? "bg-slate-900/80" : "bg-brand-50"}>
            <tr>
              {headers.map((header) => (
                <th key={header} className={`whitespace-nowrap px-3 py-2 text-left font-semibold ${darkMode ? "text-brand-200" : "text-brand-900"}`}>
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, index) => (
              <tr key={index} className={`${darkMode ? "border-slate-800/90 odd:bg-slate-900/40 even:bg-slate-900/70" : "border-slate-100 odd:bg-white even:bg-slate-50/50"} border-t`}>
                {headers.map((header) => (
                  <td key={header} className={`whitespace-nowrap px-3 py-2 ${darkMode ? "text-slate-200" : "text-slate-700"}`}>
                    {row[header] ?? "-"}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

export default function HomePage() {
  const [file, setFile] = useState<File | null>(null);
  const [sharepointUrl, setSharepointUrl] = useState("");
  const [sheetName, setSheetName] = useState("Procesos Adminis_Penal");
  const [loading, setLoading] = useState(false);
  const [loadingProgress, setLoadingProgress] = useState(0);
  const [loadingStage, setLoadingStage] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [data, setData] = useState<PreviewResponse | null>(null);
  const [showSettings, setShowSettings] = useState(false);
  const [darkMode, setDarkMode] = useState(false);
  const [diagnosingSharepoint, setDiagnosingSharepoint] = useState(false);
  const [diagnosticError, setDiagnosticError] = useState<string | null>(null);
  const [diagnosticData, setDiagnosticData] = useState<SharepointDiagnosticResponse | null>(null);

  useEffect(() => {
    const enabled = window.localStorage.getItem("dark_mode") === "1";
    setDarkMode(enabled);
    document.documentElement.classList.toggle("dark-theme", enabled);
  }, []);

  useEffect(() => {
    window.localStorage.setItem("dark_mode", darkMode ? "1" : "0");
    document.documentElement.classList.toggle("dark-theme", darkMode);
  }, [darkMode]);

  const requestPreview = (formData: FormData) =>
    new Promise<PreviewResponse>((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      let processingTimer: number | null = null;

      xhr.open("POST", `${API_URL}/api/alerts/preview`);
      xhr.responseType = "json";

      xhr.upload.onloadstart = () => {
        setLoadingProgress(8);
        setLoadingStage("Preparando archivo...");
      };

      xhr.upload.onprogress = (event) => {
        if (!event.lengthComputable) return;
        const uploadPercent = Math.round((event.loaded / event.total) * 68);
        setLoadingProgress(Math.max(10, Math.min(68, uploadPercent)));
        setLoadingStage("Subiendo archivo...");
      };

      xhr.upload.onload = () => {
        setLoadingStage("Leyendo hoja y cargando vista previa...");
        setLoadingProgress((prev) => Math.max(prev, 70));
        processingTimer = window.setInterval(() => {
          setLoadingProgress((prev) => {
            if (prev >= 94) return prev;
            return prev < 84 ? prev + 3 : prev + 1;
          });
        }, 450);
      };

      xhr.onerror = () => {
        if (processingTimer) window.clearInterval(processingTimer);
        reject(new Error("No se pudo conectar con el servidor."));
      };

      xhr.onload = () => {
        if (processingTimer) window.clearInterval(processingTimer);
        setLoadingProgress(100);
        setLoadingStage("Finalizando respuesta...");
        const payload =
          xhr.response && typeof xhr.response === "object"
            ? xhr.response
            : JSON.parse(xhr.responseText || "{}");
        if (xhr.status >= 200 && xhr.status < 300) {
          resolve(payload as PreviewResponse);
          return;
        }
        reject(new Error((payload as { detail?: string }).detail ?? "Error procesando."));
      };

      xhr.send(formData);
    });

  const onSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const hasFile = !!file;
    const hasSharepoint = sharepointUrl.trim().length > 0;
    if (!hasFile && !hasSharepoint) {
      setError("Selecciona archivo Excel o pega link de SharePoint.");
      return;
    }

    setLoading(true);
    setLoadingProgress(5);
    setLoadingStage("Iniciando carga...");
    setError(null);
    setDiagnosticError(null);

    const formData = new FormData();
    if (file) formData.append("file", file);
    if (hasSharepoint) formData.append("sharepoint_url", sharepointUrl.trim());
    formData.append("sheet_name", sheetName);

    try {
      const parsed = await requestPreview(formData);
      setData(parsed);
    } catch (err) {
      setError(err instanceof Error ? err.message : "No se pudo procesar.");
    } finally {
      setLoading(false);
      setLoadingProgress(0);
      setLoadingStage("");
    }
  };

  const onDiagnoseSharepoint = async () => {
    const url = sharepointUrl.trim();
    if (!url) {
      setDiagnosticError("Pega un link de SharePoint para diagnostico.");
      setDiagnosticData(null);
      return;
    }
    setDiagnosingSharepoint(true);
    setDiagnosticError(null);
    setDiagnosticData(null);
    try {
      const formData = new FormData();
      formData.append("sharepoint_url", url);
      const response = await fetch(`${API_URL}/api/sharepoint/diagnostic`, { method: "POST", body: formData });
      const payload = (await response.json()) as SharepointDiagnosticResponse | { detail?: string };
      if (!response.ok) throw new Error((payload as { detail?: string }).detail ?? "Error diagnostico.");
      setDiagnosticData(payload as SharepointDiagnosticResponse);
    } catch (err) {
      setDiagnosticError(err instanceof Error ? err.message : "No se pudo diagnosticar.");
    } finally {
      setDiagnosingSharepoint(false);
    }
  };

  const onFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    setFile(event.target.files?.[0] ?? null);
    setError(null);
  };

  return (
    <main className="relative mx-auto min-h-screen max-w-7xl px-4 py-10 sm:px-6 lg:px-8">
      <div className="pointer-events-none fixed inset-0 -z-10">
        <div className={`absolute inset-0 transition-opacity duration-700 ease-[cubic-bezier(0.22,1,0.36,1)] ${darkMode ? "opacity-0" : "opacity-100"}`} style={{ backgroundImage: "radial-gradient(circle at 10% 20%, rgba(66,182,175,0.14), transparent 36%), radial-gradient(circle at 80% 0%, rgba(122,212,202,0.20), transparent 28%), linear-gradient(180deg, #f3fbfa 0%, #f7faf9 48%, #eef7f5 100%)" }} />
        <div className={`absolute inset-0 transition-opacity duration-700 ease-[cubic-bezier(0.22,1,0.36,1)] ${darkMode ? "opacity-100" : "opacity-0"}`} style={{ backgroundImage: "radial-gradient(circle at 10% 20%, rgba(30,41,59,0.45), transparent 36%), radial-gradient(circle at 80% 0%, rgba(15,23,42,0.5), transparent 28%), linear-gradient(180deg, #0b1220 0%, #0f172a 48%, #111827 100%)" }} />
      </div>

      <aside className="fixed left-4 top-1/2 z-40 -translate-y-1/2">
        <div className="flex flex-col items-start gap-2">
          <button type="button" onClick={() => setShowSettings((prev) => !prev)} className={`rounded-full border p-3 shadow-lg transition ${darkMode ? "border-slate-700 bg-slate-900/90 text-slate-200 hover:bg-slate-800" : "border-slate-300 bg-white/90 text-slate-700 hover:bg-white"}`} aria-label="Abrir configuracion">
            ⚙
          </button>
          {showSettings && (
            <section className="card w-80 p-4">
              <h3 className={`text-base font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Configuracion</h3>
              <div className={`mt-4 flex items-center justify-between rounded-xl border px-3 py-2 ${darkMode ? "border-slate-700 bg-slate-900/80" : "border-slate-200 bg-slate-50"}`}>
                <div>
                  <p className={`text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Modo oscuro</p>
                </div>
                <button type="button" onClick={() => setDarkMode((prev) => !prev)} className={`h-7 w-12 rounded-full p-1 transition ${darkMode ? "bg-brand-600" : "bg-slate-300"}`}>
                  <span className={`block h-5 w-5 rounded-full bg-white transition ${darkMode ? "translate-x-5" : "translate-x-0"}`} />
                </button>
              </div>
              <p className={`mt-4 text-xs ${darkMode ? "text-slate-300" : "text-slate-600"}`}>Desarrollado por el ingeniero William Rodriguez</p>
            </section>
          )}
        </div>
      </aside>

      <section className={`mb-8 overflow-hidden rounded-3xl p-8 text-white shadow-glow ${darkMode ? "bg-gradient-to-r from-slate-900 via-slate-800 to-slate-700" : "bg-gradient-to-r from-brand-700 via-brand-600 to-brand-500"}`}>
        <p className={`text-sm font-medium uppercase tracking-[0.2em] ${darkMode ? "text-slate-300" : "text-brand-100"}`}>BASE ALERTAS</p>
        <h1 className="mt-3 text-3xl font-bold sm:text-4xl">Lectura limpia de Excel</h1>
        <p className={`mt-3 max-w-3xl text-sm ${darkMode ? "text-slate-300" : "text-brand-50"}`}>
          Modo base para reconstruir el proyecto desde cero. Solo se carga el archivo, se valida la hoja y se muestra una vista previa directa de los datos leidos.
        </p>
      </section>

      <section className="card mb-8 p-6">
        <form onSubmit={onSubmit} className="grid gap-4 sm:grid-cols-12">
          <div className="sm:col-span-4">
            <label className={`mb-2 block text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Archivo Excel</label>
            <input type="file" accept=".xlsx,.xlsm,.xltx,.xltm" onChange={onFileChange} className={`block w-full rounded-xl border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/70 text-slate-200" : "border-slate-300 bg-white"}`} />
          </div>
          <div className="sm:col-span-5">
            <label className={`mb-2 block text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Si es archivo de SharePoint, pega el link aqui</label>
            <input type="url" value={sharepointUrl} onChange={(e) => setSharepointUrl(e.target.value)} placeholder="https://tuempresa.sharepoint.com/.../archivo.xlsx" className={`block w-full rounded-xl border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/70 text-slate-100" : "border-slate-300 text-slate-800"}`} />
          </div>
          <div className="sm:col-span-3">
            <label className={`mb-2 block text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Nombre de hoja</label>
            <input type="text" value={sheetName} onChange={(e) => setSheetName(e.target.value)} className={`block w-full rounded-xl border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/70 text-slate-100" : "border-slate-300 text-slate-800"}`} />
          </div>
          <div className="sm:col-span-12 sm:flex sm:items-end sm:justify-end sm:gap-3">
            <button type="button" onClick={onDiagnoseSharepoint} disabled={diagnosingSharepoint} className={`w-full sm:w-64 rounded-xl border px-4 py-2.5 text-sm font-semibold transition disabled:opacity-60 ${darkMode ? "border-slate-600 bg-slate-900/70 text-slate-100 hover:bg-slate-800" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-50"}`}>
              {diagnosingSharepoint ? "Probando..." : "Probar link SharePoint"}
            </button>
            <button type="submit" disabled={loading} className={`w-full sm:w-64 rounded-xl px-4 py-2.5 text-sm font-semibold text-white transition disabled:opacity-60 ${darkMode ? "bg-brand-600 hover:bg-brand-500" : "bg-slate-900 hover:bg-slate-800"}`}>
              {loading ? "Procesando..." : "Leer Excel"}
            </button>
          </div>
          {loading && (
            <div className="sm:col-span-12">
              <div className={`overflow-hidden rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <div className="flex items-center justify-between gap-3">
                  <div>
                    <p className={`text-sm font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{loadingStage || "Procesando archivo..."}</p>
                    <p className={`mt-1 text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Se esta validando la hoja y cargando una vista previa simple del archivo.</p>
                  </div>
                  <p className={`text-sm font-semibold tabular-nums ${darkMode ? "text-brand-200" : "text-brand-700"}`}>{loadingProgress}%</p>
                </div>
                <div className={`mt-3 h-3 overflow-hidden rounded-full ${darkMode ? "bg-slate-800" : "bg-slate-200"}`}>
                  <div className="h-full rounded-full bg-gradient-to-r from-brand-500 via-cyan-400 to-emerald-400 transition-[width] duration-300 ease-out" style={{ width: `${Math.max(6, loadingProgress)}%` }} />
                </div>
              </div>
            </div>
          )}
        </form>

        {error && <p className={`mt-4 rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-red-900/50 bg-red-950/40 text-red-300" : "border-red-200 bg-red-50 text-red-700"}`}>{error}</p>}
        {diagnosticError && <p className={`mt-4 rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-amber-900/50 bg-amber-950/40 text-amber-300" : "border-amber-200 bg-amber-50 text-amber-700"}`}>{diagnosticError}</p>}

        {diagnosticData && (
          <div className={`mt-4 rounded-2xl border p-4 text-sm ${darkMode ? "border-slate-700 bg-slate-900/60 text-slate-200" : "border-slate-200 bg-slate-50 text-slate-700"}`}>
            <p><span className="font-semibold">Graph configurado:</span> {diagnosticData.graph.configured ? "Si" : "No"}</p>
            <p><span className="font-semibold">Token Graph:</span> {diagnosticData.graph.token_ok ? "OK" : diagnosticData.graph.token_error || "Sin token"}</p>
            <p><span className="font-semibold">Descarga:</span> {diagnosticData.download_ok ? `OK (${diagnosticData.filename || "archivo"}, ${diagnosticData.bytes || 0} bytes)` : diagnosticData.download_error || "Fallo"}</p>
          </div>
        )}
      </section>

      {data && (
        <section className="mb-8 grid gap-4 md:grid-cols-3">
          <div className={`rounded-2xl border p-5 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-white"}`}>
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Hoja usada</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{data.sheet_used}</p>
          </div>
          <div className={`rounded-2xl border p-5 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-white"}`}>
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Filas leidas</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{data.source_total_rows}</p>
          </div>
          <div className={`rounded-2xl border p-5 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-white"}`}>
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Columnas detectadas</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{data.source_columns.length}</p>
          </div>
        </section>
      )}

      {data && (
        <section className="card mb-8 p-6">
          <h2 className={`text-lg font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Hojas disponibles</h2>
          <div className="mt-4 flex flex-wrap gap-2">
            {data.available_sheets.map((sheet) => (
              <span key={sheet} className={`rounded-full border px-3 py-1 text-sm ${sheet === data.sheet_used ? (darkMode ? "border-brand-500 bg-brand-500/20 text-brand-100" : "border-brand-300 bg-brand-50 text-brand-800") : (darkMode ? "border-slate-700 bg-slate-900/70 text-slate-300" : "border-slate-200 bg-slate-50 text-slate-700")}`}>
                {sheet}
              </span>
            ))}
          </div>
        </section>
      )}

      {data && (
        <section className="card mb-8 p-6">
          <h2 className={`text-lg font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Columnas detectadas</h2>
          <div className="mt-4 flex flex-wrap gap-2">
            {data.source_columns.map((column) => (
              <span key={column} className={`rounded-full border px-3 py-1 text-sm ${darkMode ? "border-slate-700 bg-slate-900/70 text-slate-200" : "border-slate-200 bg-slate-50 text-slate-700"}`}>
                {column}
              </span>
            ))}
          </div>
        </section>
      )}

      {data && <DataTable title="Vista previa fuente (20 filas)" rows={data.source_preview} darkMode={darkMode} />}
    </main>
  );
}
