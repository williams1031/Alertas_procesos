"use client";

import { ChangeEvent, FormEvent, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  Bar,
  BarChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis
} from "recharts";

type BoardRow = {
  responsable: string;
  vencidos: number;
  total_general: number;
  counts: Record<string, number>;
};

type BoardData = {
  key: string;
  title: string;
  description: string;
  day_columns: number[];
  rows: BoardRow[];
  totals: {
    vencidos: number;
    total_general: number;
    counts: Record<string, number>;
  };
};

type ChartPoint = {
  label: string;
  count: number;
};

type PreviewResponse = {
  sheet_used: string;
  source_total_rows: number;
  alerts_total_rows: number;
  source_preview: Record<string, string | number | null>[];
  alerts_preview: Record<string, string | number | null>[];
  tableros: BoardData[];
  status_analysis: {
    estatus_top: ChartPoint[];
    estado_top: ChartPoint[];
    analisis_top: ChartPoint[];
    pendientes_status_totals: ChartPoint[];
  };
  control_dashboard: {
    totals: {
      alertas_total: number;
      vencidas: number;
      por_vencer_0_10: number;
      rango_11_30: number;
      rango_31_60: number;
      rango_61_150: number;
    };
    tipo_counts: ChartPoint[];
    regla_counts: ChartPoint[];
    responsable_top: ChartPoint[];
    ciudad_top: ChartPoint[];
    dias_distribution: ChartPoint[];
    trigger_counts: ChartPoint[];
  };
  analysis_records: {
    Tipo: string;
    Regla: string;
    Responsable: string;
    Ciudad: string;
    DiasInt: number;
    EmailTrigger: string;
    Aviso: string;
    "Cuenta Contrato": string;
    Estatus: string;
    Quien_Liquida: string;
    Fecha_Vencimiento: string;
  }[];
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

function BoardTable({ board, darkMode }: { board: BoardData; darkMode: boolean }) {
  const [showHelp, setShowHelp] = useState(false);
  const daysMin = board.day_columns.length ? Math.min(...board.day_columns) : 0;
  const daysMax = board.day_columns.length ? Math.max(...board.day_columns) : 0;

  const boardSpecificExplanation = (key: string): string => {
    if (key === "pendientes_procedencia") {
      return "Este tablero solo muestra casos pendientes por determinar procedencia (liquidacion/analisis), dentro del horizonte de 5 meses.";
    }
    if (key === "penales_5m") {
      return "Este tablero concentra solo alertas del flujo penal con vencimiento en 5 meses o menos.";
    }
    if (key === "administrativos_5m") {
      return "Este tablero concentra solo alertas del flujo administrativo con vencimiento en 5 meses o menos.";
    }
    if (key === "combinado_5m") {
      return "Este tablero combina penal + administrativo para tener vista unificada del horizonte de 5 meses.";
    }
    return "Tablero de alertas por responsable y dias de vencimiento.";
  };

  return (
    <section className="card relative p-6">
      <div className="mb-3 flex items-start justify-between gap-3">
        <div>
          <h2 className={`text-xl font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>{board.title}</h2>
          <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>{board.description}</p>
        </div>
        <button
          type="button"
          onClick={() => setShowHelp((prev) => !prev)}
          className={`rounded-full border px-3 py-1 text-xs font-semibold transition ${
            darkMode
              ? "border-slate-600 bg-slate-900/80 text-slate-200 hover:bg-slate-800"
              : "border-slate-300 bg-white text-slate-700 hover:bg-slate-100"
          }`}
          aria-label={`Explicacion de ${board.title}`}
        >
          ? Ayuda
        </button>
      </div>

      {showHelp && (
        <div className={`absolute right-6 top-20 z-30 w-[22rem] rounded-2xl border p-4 shadow-2xl ${darkMode ? "border-slate-700 bg-slate-950/95 text-slate-200" : "border-slate-200 bg-white text-slate-700"}`}>
          <p className={`text-sm font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Como leer este tablero</p>
          <p className="mt-2 text-xs leading-relaxed">{boardSpecificExplanation(board.key)}</p>
          <p className="mt-2 text-xs leading-relaxed">
            Los numeros en cada celda indican cuantas alertas tiene ese responsable para ese dia exacto de vencimiento.
            Columnas visibles: dias entre {daysMin} y {daysMax}.
          </p>
          <p className="mt-2 text-xs leading-relaxed">
            Celdas resaltadas en ambar: alertas cercanas (0 a 10 dias). Columna <b>Vencidos</b> en rojo: casos con dias negativos.
            <b> Total general</b>: total acumulado por responsable.
          </p>
          <p className="mt-2 text-xs leading-relaxed">
            Si ves diferencias entre responsables, puede ser por casos con dobles responsables, que se contabilizan para cada responsable asignado.
          </p>
        </div>
      )}

      <div className={`overflow-auto rounded-xl border ${darkMode ? "border-slate-700/80" : "border-slate-200"}`}>
        <table className="min-w-full text-xs">
          <thead className={darkMode ? "bg-slate-900/80" : "bg-brand-50"}>
            <tr>
              <th className={`sticky left-0 z-20 px-3 py-2 text-left font-semibold ${darkMode ? "bg-slate-900 text-brand-200" : "bg-brand-50 text-brand-900"}`}>
                {board.key === "pendientes_procedencia" ? "Asignación (Liquidación)" : "Responsable"}
              </th>
              {board.day_columns.map((day) => (
                <th key={day} className={`px-2 py-2 text-center font-semibold ${darkMode ? "text-brand-200" : "text-brand-900"}`}>
                  {day}
                </th>
              ))}
              <th className={`px-2 py-2 text-center font-semibold ${darkMode ? "text-rose-300" : "text-rose-700"}`}>Vencidos</th>
              <th className={`px-2 py-2 text-center font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Total general</th>
            </tr>
          </thead>
          <tbody>
            {board.rows.map((row) => (
              <tr
                key={row.responsable}
                className={`${darkMode ? "border-slate-800/90 odd:bg-slate-900/40 even:bg-slate-900/70" : "border-slate-100 odd:bg-white even:bg-slate-50/50"} border-t`}
              >
                <td className={`sticky left-0 z-10 px-3 py-2 font-semibold ${darkMode ? "bg-slate-900/95 text-slate-200" : "bg-white text-slate-800"}`}>
                  {row.responsable}
                </td>
                {board.day_columns.map((day) => {
                  const value = row.counts[String(day)] ?? 0;
                  const urgent = day <= 10 && value > 0;
                  return (
                    <td
                      key={`${row.responsable}-${day}`}
                      className={`px-2 py-2 text-center ${darkMode ? "text-slate-300" : "text-slate-700"} ${urgent ? (darkMode ? "bg-amber-900/40 text-amber-200 font-semibold" : "bg-amber-100 text-amber-900 font-semibold") : ""}`}
                    >
                      {value || ""}
                    </td>
                  );
                })}
                <td className={`px-2 py-2 text-center font-semibold ${row.vencidos > 0 ? (darkMode ? "bg-rose-900/50 text-rose-200" : "bg-rose-100 text-rose-800") : (darkMode ? "text-slate-300" : "text-slate-700")}`}>
                  {row.vencidos || ""}
                </td>
                <td className={`px-2 py-2 text-center font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{row.total_general}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className={darkMode ? "border-t border-slate-600 bg-slate-950/80" : "border-t border-slate-300 bg-slate-100"}>
              <td className={`sticky left-0 z-10 px-3 py-2 font-bold ${darkMode ? "bg-slate-950 text-slate-100" : "bg-slate-100 text-slate-900"}`}>Total general</td>
              {board.day_columns.map((day) => (
                <td key={`tot-${board.key}-${day}`} className={`px-2 py-2 text-center font-bold ${darkMode ? "text-slate-200" : "text-slate-900"}`}>
                  {board.totals.counts[String(day)] || ""}
                </td>
              ))}
              <td className={`px-2 py-2 text-center font-bold ${darkMode ? "text-rose-200" : "text-rose-800"}`}>{board.totals.vencidos || ""}</td>
              <td className={`px-2 py-2 text-center font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{board.totals.total_general}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </section>
  );
}

function MiniBarChart({ title, data, darkMode }: { title: string; data: ChartPoint[]; darkMode: boolean }) {
  const maxValue = useMemo(() => Math.max(...data.map((x) => x.count), 1), [data]);
  return (
    <section className="card p-5">
      <h3 className={`mb-3 text-base font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>{title}</h3>
      <div className="space-y-2">
        {data.length === 0 && <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Sin datos.</p>}
        {data.map((item) => (
          <div key={`${title}-${item.label}`} className="space-y-1">
            <div className="flex items-center justify-between text-xs">
              <span className={`truncate pr-2 ${darkMode ? "text-slate-300" : "text-slate-700"}`}>{item.label}</span>
              <span className={`font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{item.count}</span>
            </div>
            <div className={`h-2 rounded-full ${darkMode ? "bg-slate-800" : "bg-slate-200"}`}>
              <div
                className={`h-2 rounded-full ${darkMode ? "bg-brand-500" : "bg-brand-600"}`}
                style={{ width: `${(item.count / maxValue) * 100}%` }}
              />
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

function InteractiveBarChart({
  title,
  data,
  darkMode,
  barColor = "#239790"
}: {
  title: string;
  data: ChartPoint[];
  darkMode: boolean;
  barColor?: string;
}) {
  const chartData = data.slice(0, 12);
  return (
    <section className="card p-5">
      <h3 className={`mb-3 text-base font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>{title}</h3>
      <div className="h-64 w-full">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={chartData} margin={{ top: 10, right: 10, left: -10, bottom: 30 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={darkMode ? "#334155" : "#e2e8f0"} />
            <XAxis
              dataKey="label"
              angle={-20}
              textAnchor="end"
              height={60}
              interval={0}
              tick={{ fill: darkMode ? "#cbd5e1" : "#334155", fontSize: 11 }}
            />
            <YAxis tick={{ fill: darkMode ? "#cbd5e1" : "#334155", fontSize: 11 }} allowDecimals={false} />
            <Tooltip
              contentStyle={{
                background: darkMode ? "#0f172a" : "#ffffff",
                border: darkMode ? "1px solid #334155" : "1px solid #cbd5e1",
                borderRadius: "10px",
                color: darkMode ? "#e2e8f0" : "#0f172a"
              }}
            />
            <Bar dataKey="count" fill={barColor} radius={[6, 6, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </section>
  );
}

function DataTable({
  title,
  rows,
  darkMode
}: {
  title: string;
  rows: Record<string, string | number | null>[];
  darkMode: boolean;
}) {
  const headers = useMemo(() => (rows.length ? Object.keys(rows[0]) : []), [rows]);
  return (
    <section className="card p-5">
      <h3 className={`mb-4 text-lg font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>{title}</h3>
      {!rows.length ? (
        <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Sin datos para mostrar.</p>
      ) : (
        <div className={`overflow-auto rounded-xl border ${darkMode ? "border-slate-700/80" : "border-slate-200"}`}>
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
      )}
    </section>
  );
}

export default function HomePage() {
  const [file, setFile] = useState<File | null>(null);
  const [sharepointUrl, setSharepointUrl] = useState("");
  const [sheetName, setSheetName] = useState("Procesos Adminis_Penal");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [data, setData] = useState<PreviewResponse | null>(null);
  const [showSettings, setShowSettings] = useState(false);
  const [darkMode, setDarkMode] = useState(false);
  const [diagnosingSharepoint, setDiagnosingSharepoint] = useState(false);
  const [diagnosticError, setDiagnosticError] = useState<string | null>(null);
  const [diagnosticData, setDiagnosticData] = useState<SharepointDiagnosticResponse | null>(null);
  const [exportMode, setExportMode] = useState<"BASICA" | "COMPLETA">("BASICA");
  const [filterTipo, setFilterTipo] = useState("TODOS");
  const [filterResponsable, setFilterResponsable] = useState("TODOS");
  const [filterCiudad, setFilterCiudad] = useState("TODOS");
  const [filterRangoDias, setFilterRangoDias] = useState("TODOS");

  useEffect(() => {
    const enabled = window.localStorage.getItem("dark_mode") === "1";
    setDarkMode(enabled);
    document.documentElement.classList.toggle("dark-theme", enabled);
  }, []);

  useEffect(() => {
    window.localStorage.setItem("dark_mode", darkMode ? "1" : "0");
    document.documentElement.classList.toggle("dark-theme", darkMode);
  }, [darkMode]);

  useEffect(() => {
    if (data) {
      setFilterTipo("TODOS");
      setFilterResponsable("TODOS");
      setFilterCiudad("TODOS");
      setFilterRangoDias("TODOS");
    }
  }, [data]);

  const analysisRecords = data?.analysis_records ?? [];

  const tipoOptions = useMemo(() => {
    const vals = Array.from(new Set(analysisRecords.map((r) => r.Tipo).filter(Boolean)));
    return vals.sort();
  }, [analysisRecords]);

  const responsableOptions = useMemo(() => {
    const vals = Array.from(new Set(analysisRecords.map((r) => r.Responsable).filter(Boolean)));
    return vals.sort();
  }, [analysisRecords]);

  const ciudadOptions = useMemo(() => {
    const vals = Array.from(new Set(analysisRecords.map((r) => r.Ciudad).filter(Boolean)));
    return vals.sort();
  }, [analysisRecords]);

  const filteredRecords = useMemo(() => {
    return analysisRecords.filter((r) => {
      if (filterTipo !== "TODOS" && r.Tipo !== filterTipo) return false;
      if (filterResponsable !== "TODOS" && r.Responsable !== filterResponsable) return false;
      if (filterCiudad !== "TODOS" && r.Ciudad !== filterCiudad) return false;
      if (filterRangoDias !== "TODOS") {
        const d = Number(r.DiasInt);
        if (filterRangoDias === "VENCIDAS" && d >= 0) return false;
        if (filterRangoDias === "D0_10" && !(d >= 0 && d <= 10)) return false;
        if (filterRangoDias === "D11_30" && !(d >= 11 && d <= 30)) return false;
        if (filterRangoDias === "D31_60" && !(d >= 31 && d <= 60)) return false;
        if (filterRangoDias === "D61_150" && !(d >= 61 && d <= 150)) return false;
      }
      return true;
    });
  }, [analysisRecords, filterTipo, filterResponsable, filterCiudad, filterRangoDias]);

  const aggregateBy = (rows: typeof filteredRecords, key: keyof (typeof filteredRecords)[number], top = 12): ChartPoint[] => {
    const counts = new Map<string, number>();
    for (const row of rows) {
      const label = String(row[key] ?? "").trim();
      if (!label) continue;
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return Array.from(counts.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, top);
  };

  const chartTipo = useMemo(() => aggregateBy(filteredRecords, "Tipo", 12), [filteredRecords]);
  const chartRegla = useMemo(() => aggregateBy(filteredRecords, "Regla", 12), [filteredRecords]);
  const chartResponsable = useMemo(() => aggregateBy(filteredRecords, "Responsable", 12), [filteredRecords]);
  const chartCiudad = useMemo(() => aggregateBy(filteredRecords, "Ciudad", 12), [filteredRecords]);

  const chartDias = useMemo(() => {
    const bucket = new Map<string, number>();
    for (const row of filteredRecords) {
      const d = Number(row.DiasInt);
      const label = d < 0 ? "Vencidas" : String(d);
      bucket.set(label, (bucket.get(label) ?? 0) + 1);
    }
    return Array.from(bucket.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => {
        if (a.label === "Vencidas") return -1;
        if (b.label === "Vencidas") return 1;
        return Number(a.label) - Number(b.label);
      })
      .slice(0, 20);
  }, [filteredRecords]);

  const filteredTotals = useMemo(() => {
    const base = {
      alertas_total: filteredRecords.length,
      vencidas: 0,
      por_vencer_0_10: 0,
      rango_11_30: 0,
      rango_31_60: 0,
      rango_61_150: 0
    };
    for (const row of filteredRecords) {
      const d = Number(row.DiasInt);
      if (d < 0) base.vencidas += 1;
      else if (d <= 10) base.por_vencer_0_10 += 1;
      else if (d <= 30) base.rango_11_30 += 1;
      else if (d <= 60) base.rango_31_60 += 1;
      else if (d <= 150) base.rango_61_150 += 1;
    }
    return base;
  }, [filteredRecords]);

  const exportFilteredAnalysis = (format: "csv" | "xlsx") => {
    if (!filteredRecords.length) return;
    const exportRows =
      exportMode === "BASICA"
        ? filteredRecords.map((row) => ({
            AVISO: row.Aviso ?? "",
            "CUENTA CONTRATO": row["Cuenta Contrato"] ?? "",
            ESTATUS: row.Estatus ?? "",
            "QUIEN LIQUIDA": row.Quien_Liquida ?? "",
            "FECHA DE VENCIMIENTO": row.Fecha_Vencimiento ?? ""
          }))
        : filteredRecords.map((row) => ({
            AVISO: row.Aviso ?? "",
            "CUENTA CONTRATO": row["Cuenta Contrato"] ?? "",
            ESTATUS: row.Estatus ?? "",
            "QUIEN LIQUIDA": row.Quien_Liquida ?? "",
            "FECHA DE VENCIMIENTO": row.Fecha_Vencimiento ?? "",
            "DIAS PARA VENCIMIENTO": row.DiasInt,
            TIPO: row.Tipo ?? "",
            REGLA: row.Regla ?? "",
            RESPONSABLE: row.Responsable ?? "",
            CIUDAD: row.Ciudad ?? "",
            "TRIGGER CORREO": row.EmailTrigger || ""
          }));
    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "AnalisisFiltrado");

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
    const modeTag = exportMode === "BASICA" ? "basica" : "completa";
    const filename = `analisis_filtrado_${modeTag}_${stamp}.${format}`;
    XLSX.writeFile(wb, filename, { bookType: format });
  };

  const onFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    setFile(event.target.files?.[0] ?? null);
    setError(null);
  };

  const onSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const hasFile = !!file;
    const hasSharepoint = sharepointUrl.trim().length > 0;
    if (!hasFile && !hasSharepoint) {
      setError("Selecciona archivo Excel o pega link de SharePoint.");
      return;
    }
    setLoading(true);
    setError(null);
    const formData = new FormData();
    if (file) formData.append("file", file);
    if (hasSharepoint) formData.append("sharepoint_url", sharepointUrl.trim());
    formData.append("sheet_name", sheetName);
    try {
      const response = await fetch(`${API_URL}/api/alerts/preview`, { method: "POST", body: formData });
      const payload = (await response.json()) as PreviewResponse | { detail?: string };
      if (!response.ok) throw new Error((payload as { detail?: string }).detail ?? "Error procesando.");
      const parsed = payload as PreviewResponse;
      setData(parsed);
    } catch (err) {
      setError(err instanceof Error ? err.message : "No se pudo procesar.");
    } finally {
      setLoading(false);
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
              <h3 className={`text-base font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>Configuracion</h3>
              <div className={`mt-4 flex items-center justify-between rounded-xl border px-3 py-2 ${darkMode ? "border-slate-700 bg-slate-900/80" : "border-slate-200 bg-slate-50"}`}>
                <div>
                  <p className={`text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Modo oscuro</p>
                </div>
                <button type="button" onClick={() => setDarkMode((prev) => !prev)} className={`h-7 w-12 rounded-full p-1 transition ${darkMode ? "bg-brand-600" : "bg-slate-300"}`}>
                  <span className={`block h-5 w-5 rounded-full bg-white transition ${darkMode ? "translate-x-5" : "translate-x-0"}`} />
                </button>
              </div>
              <p className={`mt-4 text-xs ${darkMode ? "text-slate-300" : "text-slate-600"}`}>
                Desarrollado por el ingeniero William Rodriguez
              </p>
            </section>
          )}
        </div>
      </aside>

      <section className={`mb-8 overflow-hidden rounded-3xl p-8 text-white shadow-glow ${darkMode ? "bg-gradient-to-r from-slate-900 via-slate-800 to-slate-700" : "bg-gradient-to-r from-brand-700 via-brand-600 to-brand-500"}`}>
        <p className={`text-sm font-medium uppercase tracking-[0.2em] ${darkMode ? "text-slate-300" : "text-brand-100"}`}>TABLERO ALERTAS</p>
        <h1 className="mt-3 text-3xl font-bold sm:text-4xl">Tableros por acto administrativo</h1>
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
            <button type="submit" disabled={loading} className={`w-full sm:w-64 rounded-xl px-4 py-2.5 text-sm font-semibold text-white transition disabled:opacity-60 ${darkMode ? "bg-brand-600 hover:bg-brand-500" : "bg-ink hover:bg-slate-900"}`}>
              {loading ? "Procesando..." : "Leer y construir tableros"}
            </button>
          </div>
        </form>

        {diagnosticError && <p className={`mt-4 rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-red-900/50 bg-red-950/40 text-red-300" : "border-red-200 bg-red-50 text-red-700"}`}>{diagnosticError}</p>}
        {diagnosticData && (
          <div className={`mt-4 rounded-xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/70" : "border-slate-200 bg-slate-50/80"}`}>
            <p className={`text-sm font-semibold ${darkMode ? "text-slate-100" : "text-slate-800"}`}>Diagnostico SharePoint</p>
            <p className={`text-xs ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Graph: {diagnosticData.graph.configured ? "Configurado" : "No configurado"} | Token: {diagnosticData.graph.token_ok ? "OK" : "FALLO"} | Descarga: {diagnosticData.download_ok ? "OK" : "FALLO"}</p>
          </div>
        )}
        {error && <p className={`mt-4 rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-red-900/50 bg-red-950/40 text-red-300" : "border-red-200 bg-red-50 text-red-700"}`}>{error}</p>}
      </section>

      {data && (
        <section className="mb-8 grid gap-4 sm:grid-cols-3">
          <div className="card p-4"><p className={`text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Hoja leida</p><p className={`mt-1 text-lg font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>{data.sheet_used}</p></div>
          <div className="card p-4"><p className={`text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Registros fuente</p><p className={`mt-1 text-lg font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>{data.source_total_rows}</p></div>
          <div className="card p-4"><p className={`text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Alertas</p><p className={`mt-1 text-lg font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>{data.alerts_total_rows}</p></div>
        </section>
      )}

      {data && (
        <section className="space-y-6">
          {data.tableros.map((board) => (
            <BoardTable key={board.key} board={board} darkMode={darkMode} />
          ))}
        </section>
      )}

      {data && (
        <section className="mt-8 grid gap-6 lg:grid-cols-3">
          <MiniBarChart title="Estatus (Top)" data={data.status_analysis.estatus_top} darkMode={darkMode} />
          <MiniBarChart title="Estado (Top)" data={data.status_analysis.estado_top} darkMode={darkMode} />
          <MiniBarChart title="Pendientes Clave (Totales)" data={data.status_analysis.pendientes_status_totals} darkMode={darkMode} />
        </section>
      )}

      {data && (
        <section className="mt-8 space-y-6">
          <section className="card p-6">
            <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
              <div>
                <h2 className={`text-xl font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>
                  Tablero de Control Analitico
                </h2>
                <p className={`mt-1 text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>
                  Analisis completo del documento con indicadores dinamicos para toma de decisiones.
                </p>
              </div>
              <div className="flex gap-2">
                <select
                  value={exportMode}
                  onChange={(e) => setExportMode(e.target.value as "BASICA" | "COMPLETA")}
                  className={`rounded-lg border px-3 py-2 text-xs font-semibold ${
                    darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 bg-white text-slate-800"
                  }`}
                >
                  <option value="BASICA">Exportación básica</option>
                  <option value="COMPLETA">Exportación completa</option>
                </select>
                <button
                  type="button"
                  onClick={() => exportFilteredAnalysis("csv")}
                  disabled={!filteredRecords.length}
                  className={`rounded-lg px-3 py-2 text-xs font-semibold transition disabled:opacity-50 ${
                    darkMode ? "bg-slate-800 text-slate-100 hover:bg-slate-700" : "bg-slate-200 text-slate-800 hover:bg-slate-300"
                  }`}
                >
                  Exportar CSV
                </button>
                <button
                  type="button"
                  onClick={() => exportFilteredAnalysis("xlsx")}
                  disabled={!filteredRecords.length}
                  className={`rounded-lg px-3 py-2 text-xs font-semibold text-white transition disabled:opacity-50 ${
                    darkMode ? "bg-brand-600 hover:bg-brand-500" : "bg-brand-700 hover:bg-brand-800"
                  }`}
                >
                  Exportar Excel
                </button>
              </div>
            </div>
            <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Tipo</label>
                <select value={filterTipo} onChange={(e) => setFilterTipo(e.target.value)} className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}>
                  <option value="TODOS">Todos</option>
                  {tipoOptions.map((x) => <option key={x} value={x}>{x}</option>)}
                </select>
              </div>
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Responsable</label>
                <select value={filterResponsable} onChange={(e) => setFilterResponsable(e.target.value)} className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}>
                  <option value="TODOS">Todos</option>
                  {responsableOptions.map((x) => <option key={x} value={x}>{x}</option>)}
                </select>
              </div>
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Ciudad</label>
                <select value={filterCiudad} onChange={(e) => setFilterCiudad(e.target.value)} className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}>
                  <option value="TODOS">Todas</option>
                  {ciudadOptions.map((x) => <option key={x} value={x}>{x}</option>)}
                </select>
              </div>
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Rango dias</label>
                <select value={filterRangoDias} onChange={(e) => setFilterRangoDias(e.target.value)} className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}>
                  <option value="TODOS">Todos</option>
                  <option value="VENCIDAS">Vencidas</option>
                  <option value="D0_10">0 a 10</option>
                  <option value="D11_30">11 a 30</option>
                  <option value="D31_60">31 a 60</option>
                  <option value="D61_150">61 a 150</option>
                </select>
              </div>
            </div>
            <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6">
              <div className={`rounded-xl border p-3 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <p className={`text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Alertas totales</p>
                <p className={`text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>
                  {filteredTotals.alertas_total}
                </p>
              </div>
              <div className={`rounded-xl border p-3 ${darkMode ? "border-rose-900/60 bg-rose-950/40" : "border-rose-200 bg-rose-50"}`}>
                <p className={`text-xs ${darkMode ? "text-rose-300" : "text-rose-600"}`}>Vencidas</p>
                <p className={`text-2xl font-bold ${darkMode ? "text-rose-200" : "text-rose-700"}`}>
                  {filteredTotals.vencidas}
                </p>
              </div>
              <div className={`rounded-xl border p-3 ${darkMode ? "border-amber-900/60 bg-amber-950/40" : "border-amber-200 bg-amber-50"}`}>
                <p className={`text-xs ${darkMode ? "text-amber-300" : "text-amber-700"}`}>0 a 10 dias</p>
                <p className={`text-2xl font-bold ${darkMode ? "text-amber-200" : "text-amber-700"}`}>
                  {filteredTotals.por_vencer_0_10}
                </p>
              </div>
              <div className={`rounded-xl border p-3 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <p className={`text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>11 a 30 dias</p>
                <p className={`text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>
                  {filteredTotals.rango_11_30}
                </p>
              </div>
              <div className={`rounded-xl border p-3 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <p className={`text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>31 a 60 dias</p>
                <p className={`text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>
                  {filteredTotals.rango_31_60}
                </p>
              </div>
              <div className={`rounded-xl border p-3 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <p className={`text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>61 a 150 dias</p>
                <p className={`text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>
                  {filteredTotals.rango_61_150}
                </p>
              </div>
            </div>
          </section>

          <section className="grid gap-6 lg:grid-cols-3">
            <InteractiveBarChart title="Alertas por Tipo" data={chartTipo} darkMode={darkMode} barColor="#0ea5e9" />
            <InteractiveBarChart title="Alertas por Regla" data={chartRegla} darkMode={darkMode} barColor="#10b981" />
            <InteractiveBarChart title="Top Responsables" data={chartResponsable} darkMode={darkMode} barColor="#8b5cf6" />
          </section>

          <section className="grid gap-6 lg:grid-cols-3">
            <InteractiveBarChart title="Top Ciudades" data={chartCiudad} darkMode={darkMode} barColor="#f59e0b" />
            <InteractiveBarChart title="Distribucion por Dias" data={chartDias} darkMode={darkMode} barColor="#ef4444" />
          </section>
        </section>
      )}

      {data && (
        <section className="mt-8 grid gap-6 lg:grid-cols-2">
          <DataTable title="Vista previa fuente (20)" rows={data.source_preview} darkMode={darkMode} />
          <DataTable title="Vista previa alertas (60)" rows={data.alerts_preview} darkMode={darkMode} />
        </section>
      )}
    </main>
  );
}
