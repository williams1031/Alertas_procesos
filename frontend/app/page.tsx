"use client";

import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";
import html2canvas from "html2canvas";
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
  source_columns: string[];
  source_total_rows: number;
  alerts_total_rows: number;
  source_preview: Record<string, string | number | null>[];
  alerts_preview: Record<string, string | number | null>[];
  status_records: {
    Anio: number | null;
    Estatus: string;
    Estado: string;
    Responsable: string;
    Ciudad: string;
    "Cuenta Contrato": string;
  }[];
  general_board_records: {
    Tipo: string;
    Responsable: string;
    Estatus: string;
    Estado: string;
    DiasInt: number;
    Ciudad: string;
    "Cuenta Contrato": string;
  }[];
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

type ChatMessage = {
  id: number;
  role: "assistant" | "user";
  text: string;
};

type ChatContext = {
  rows: Array<Record<string, unknown>>;
  scope: string;
  summary: string;
};

const API_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";

function normalizeForSearch(value: string | number | null | undefined) {
  return String(value ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function BoardTable({ board, darkMode }: { board: BoardData; darkMode: boolean }) {
  const [showHelp, setShowHelp] = useState(false);
  const [capturing, setCapturing] = useState(false);
  const [captureMessage, setCaptureMessage] = useState<string | null>(null);
  const boardRef = useRef<HTMLElement | null>(null);
  const daysMin = board.day_columns.length ? Math.min(...board.day_columns) : 0;
  const daysMax = board.day_columns.length ? Math.max(...board.day_columns) : 0;

  const captureBoard = async (): Promise<Blob> => {
    if (!boardRef.current) throw new Error("No se encontro el tablero.");
    const canvas = await html2canvas(boardRef.current, {
      scale: 2,
      useCORS: true,
      backgroundColor: darkMode ? "#0f172a" : "#ffffff"
    });
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
    if (!blob) throw new Error("No fue posible generar la imagen.");
    return blob;
  };

  const onCopyBoard = async () => {
    try {
      setCapturing(true);
      setCaptureMessage(null);
      const blob = await captureBoard();
      if (!window.ClipboardItem || !navigator.clipboard?.write) {
        throw new Error("Tu navegador no permite copiar imagenes al portapapeles.");
      }
      await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
      setCaptureMessage("Imagen copiada.");
    } catch (err) {
      setCaptureMessage(err instanceof Error ? err.message : "No se pudo copiar la imagen.");
    } finally {
      setCapturing(false);
    }
  };

  const onSaveBoard = async () => {
    try {
      setCapturing(true);
      setCaptureMessage(null);
      const blob = await captureBoard();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `tablero_${board.key}.png`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
      setCaptureMessage("Imagen guardada.");
    } catch (err) {
      setCaptureMessage(err instanceof Error ? err.message : "No se pudo guardar la imagen.");
    } finally {
      setCapturing(false);
    }
  };

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
    <section ref={boardRef} className="card relative p-6">
      <div className="mb-3 flex items-start justify-between gap-3">
        <div>
          <h2 className={`text-xl font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>{board.title}</h2>
          <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>{board.description}</p>
        </div>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={onCopyBoard}
            disabled={capturing}
            className={`rounded-full border px-3 py-1 text-xs font-semibold transition disabled:opacity-60 ${
              darkMode
                ? "border-slate-600 bg-slate-900/80 text-slate-200 hover:bg-slate-800"
                : "border-slate-300 bg-white text-slate-700 hover:bg-slate-100"
            }`}
          >
            Copiar
          </button>
          <button
            type="button"
            onClick={onSaveBoard}
            disabled={capturing}
            className={`rounded-full border px-3 py-1 text-xs font-semibold transition disabled:opacity-60 ${
              darkMode
                ? "border-slate-600 bg-slate-900/80 text-slate-200 hover:bg-slate-800"
                : "border-slate-300 bg-white text-slate-700 hover:bg-slate-100"
            }`}
          >
            Guardar
          </button>
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
      </div>
      {captureMessage && (
        <p className={`mb-3 text-xs ${darkMode ? "text-slate-300" : "text-slate-600"}`}>{captureMessage}</p>
      )}

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
  const [loadingProgress, setLoadingProgress] = useState(0);
  const [loadingStage, setLoadingStage] = useState("");
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
  const [generalFilterTipo, setGeneralFilterTipo] = useState("TODOS");
  const [generalFilterEstatus, setGeneralFilterEstatus] = useState("TODOS");
  const [generalFilterEstado, setGeneralFilterEstado] = useState("TODOS");
  const [generalFilterResponsable, setGeneralFilterResponsable] = useState("TODOS");
  const [generalSearch, setGeneralSearch] = useState("");
  const [statusYearFilter, setStatusYearFilter] = useState("TODOS");
  const [chatOpen, setChatOpen] = useState(false);
  const [chatInput, setChatInput] = useState("");
  const [chatMessages, setChatMessages] = useState<ChatMessage[]>([
    {
      id: Date.now(),
      role: "assistant",
      text: "Hola. Soy tu asistente del tablero. Puedes preguntarme por el programa o por el analisis actual."
    }
  ]);
  const chatScrollRef = useRef<HTMLDivElement | null>(null);
  const chatContextRef = useRef<ChatContext | null>(null);

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
      setGeneralFilterTipo("TODOS");
      setGeneralFilterEstatus("TODOS");
      setGeneralFilterEstado("TODOS");
      setGeneralFilterResponsable("TODOS");
      setGeneralSearch("");
      setStatusYearFilter("TODOS");
      chatContextRef.current = null;
    }
  }, [data]);

  const analysisRecords = data?.analysis_records ?? [];
  const statusRecords = data?.status_records ?? [];
  const generalBoardRecords = data?.general_board_records ?? [];

  const normalizedGeneralRecords = useMemo(
    () =>
      statusRecords.map((row) => {
        const values = Object.values(row);
        const responsables = [normalizeForSearch(row.Responsable)].filter(Boolean);
        return {
          ...row,
          Responsable: String(row.Responsable ?? "").trim() || "Sin responsable",
          Ciudad: String(row.Ciudad ?? "").trim(),
          _search: values.map((value) => normalizeForSearch(value)).join(" "),
          _estatus: normalizeForSearch(row.Estatus),
          _estado: normalizeForSearch(row.Estado),
          _responsables: responsables,
          _responsable: responsables.join(" | "),
          _ciudad: normalizeForSearch(row.Ciudad),
          _cuenta: normalizeForSearch(row["Cuenta Contrato"])
        };
      }),
    [statusRecords]
  );

  const statusYearOptions = useMemo(() => {
    const years = new Set<number>();
    for (const row of statusRecords) {
      if (typeof row.Anio === "number" && Number.isFinite(row.Anio)) years.add(row.Anio);
    }
    return Array.from(years).sort((a, b) => a - b);
  }, [statusRecords]);

  const statusRowsByYear = useMemo(() => {
    if (statusYearFilter === "TODOS") return statusRecords;
    const targetYear = Number(statusYearFilter);
    return statusRecords.filter((row) => row.Anio === targetYear);
  }, [statusRecords, statusYearFilter]);

  const buildTopCountsFromRows = (
    rows: Record<string, string | number | null>[],
    column: string,
    topN = 12
  ): ChartPoint[] => {
    const counts = new Map<string, number>();
    for (const row of rows) {
      const label = String(row[column] ?? "").trim();
      if (!label) continue;
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return Array.from(counts.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, topN);
  };

  const filteredStatusAnalysis = useMemo(() => {
    const estatusTop = buildTopCountsFromRows(statusRowsByYear, "Estatus", 12);
    const estadoTop = buildTopCountsFromRows(statusRowsByYear, "Estado", 12);
    const paraAdministrativo = statusRowsByYear.filter((row) =>
      normalizeForSearch(row.Estatus).includes("para administrativo")
    ).length;
    const paraExpediente = statusRowsByYear.filter((row) =>
      normalizeForSearch(row.Estatus).includes("para expediente")
    ).length;
    return {
      estatus_top: estatusTop,
      estado_top: estadoTop,
      pendientes_status_totals: [
        { label: "Para administrativo (incluye mixtos)", count: paraAdministrativo },
        { label: "Para expediente (incluye mixtos)", count: paraExpediente }
      ]
    };
  }, [statusRowsByYear]);

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

  const generalEstatusOptions = useMemo(() => {
    const vals = Array.from(
      new Set(statusRecords.map((r) => String(r.Estatus ?? "").trim()).filter(Boolean))
    );
    return vals.sort((a, b) => a.localeCompare(b, "es"));
  }, [statusRecords]);

  const generalEstadoOptions = useMemo(() => {
    const vals = Array.from(
      new Set(statusRecords.map((r) => String(r.Estado ?? "").trim()).filter(Boolean))
    );
    return vals.sort((a, b) => a.localeCompare(b, "es"));
  }, [statusRecords]);

  const generalTipoOptions = useMemo(() => {
    const vals = Array.from(new Set(generalBoardRecords.map((r) => String(r.Tipo ?? "").trim()).filter(Boolean)));
    return vals.sort((a, b) => a.localeCompare(b, "es"));
  }, [generalBoardRecords]);

  const generalResponsableOptions = useMemo(() => {
    const raw = generalBoardRecords.map((r) => String(r.Responsable ?? "").trim());
    return Array.from(new Set(raw.filter(Boolean))).sort((a, b) => a.localeCompare(b, "es"));
  }, [generalBoardRecords]);

  const filteredGeneralBoardRecords = useMemo(() => {
    const search = generalSearch
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();

    return generalBoardRecords.filter((row) => {
      const tipo = String(row.Tipo ?? "").trim();
      const estatus = String(row.Estatus ?? "").trim();
      const estado = String(row.Estado ?? "").trim();
      const responsable = String(row.Responsable ?? "").trim();

      if (generalFilterTipo !== "TODOS" && tipo !== generalFilterTipo) return false;
      if (generalFilterEstatus !== "TODOS" && estatus !== generalFilterEstatus) return false;
      if (generalFilterEstado !== "TODOS" && estado !== generalFilterEstado) return false;
      if (generalFilterResponsable !== "TODOS" && responsable !== generalFilterResponsable) {
        return false;
      }

      if (!search) return true;

      const haystack = Object.values(row)
        .map((value) =>
          String(value ?? "")
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
        )
        .join(" ");
      return haystack.includes(search);
    });
  }, [generalBoardRecords, generalFilterTipo, generalFilterEstatus, generalFilterEstado, generalFilterResponsable, generalSearch]);

  const generalBoard = useMemo<BoardData>(() => {
    if (!filteredGeneralBoardRecords.length) {
      return {
        key: "general_board",
        title: "Tablero General",
        description: "Vista consolidada con filtros por tipo, estatus y estado.",
        day_columns: [],
        rows: [],
        totals: { vencidos: 0, total_general: 0, counts: {} }
      };
    }

    const dayColumns = Array.from(
      new Set(
        filteredGeneralBoardRecords
          .map((row) => Number(row.DiasInt))
          .filter((value) => Number.isFinite(value) && value >= 0 && value <= 150)
      )
    ).sort((a, b) => a - b);

    const orderedDayColumns = dayColumns.length <= 32 ? dayColumns : [...dayColumns.slice(0, 24), 30, 45, 60, 90, 120, 150]
      .filter((value, index, arr) => arr.indexOf(value) === index)
      .sort((a, b) => a - b);

    const byResponsable = new Map<string, typeof filteredGeneralBoardRecords>();
    for (const row of filteredGeneralBoardRecords) {
      const key = String(row.Responsable || "").trim() || "Sin responsable";
      const current = byResponsable.get(key) ?? [];
      current.push(row);
      byResponsable.set(key, current);
    }

    const rows = Array.from(byResponsable.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([responsable, entries]) => {
        const counts: Record<string, number> = {};
        for (const day of orderedDayColumns) counts[String(day)] = 0;
        let vencidos = 0;
        for (const entry of entries) {
          const d = Number(entry.DiasInt);
          if (d < 0) vencidos += 1;
          if (orderedDayColumns.includes(d)) counts[String(d)] = (counts[String(d)] ?? 0) + 1;
        }
        return {
          responsable,
          vencidos,
          total_general: entries.length,
          counts
        };
      });

    const totalsCounts: Record<string, number> = {};
    for (const day of orderedDayColumns) totalsCounts[String(day)] = 0;
    let totalVencidos = 0;
    let totalGeneral = 0;
    for (const row of rows) {
      totalVencidos += row.vencidos;
      totalGeneral += row.total_general;
      for (const day of orderedDayColumns) {
        totalsCounts[String(day)] += row.counts[String(day)] ?? 0;
      }
    }

    return {
      key: "general_board",
      title: "Tablero General",
      description: "Vista consolidada filtrable por tipo, estatus y estado.",
      day_columns: orderedDayColumns,
      rows,
      totals: {
        vencidos: totalVencidos,
        total_general: totalGeneral,
        counts: totalsCounts
      }
    };
  }, [filteredGeneralBoardRecords]);

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

  useEffect(() => {
    if (chatScrollRef.current) {
      chatScrollRef.current.scrollTop = chatScrollRef.current.scrollHeight;
    }
  }, [chatMessages, chatOpen]);

  const normalizedAnalysis = useMemo(
    () =>
      analysisRecords.map((row) => ({
        ...row,
        _responsable: normalizeForSearch(row.Responsable),
        _tipo: normalizeForSearch(row.Tipo),
        _estatus: normalizeForSearch(row.Estatus),
        _ciudad: normalizeForSearch(row.Ciudad),
        _cuenta: normalizeForSearch(row["Cuenta Contrato"])
      })),
    [analysisRecords]
  );

  const buildChatReply = (question: string): { text: string; context?: ChatContext } => {
    const q = normalizeForSearch(question);

    if (!data) {
      return {
        text: "Aun no hay analisis cargado. Primero sube el Excel o pega el link de SharePoint y luego te respondo con datos del tablero."
      };
    }

    const daysFound = Array.from(q.matchAll(/-?\d+/g)).map((m) => Number(m[0])).filter((n) => !Number.isNaN(n));
    const mentionedDay = daysFound.length ? daysFound[0] : null;
    const asksWho =
      q.includes("quien") ||
      q.includes("quienes") ||
      q.includes("persona") ||
      q.includes("personas") ||
      q.includes("responsable") ||
      q.includes("responsables");
    const asksHowMany =
      q.includes("cuantos") ||
      q.includes("cuantas") ||
      q.includes("cantidad") ||
      q.includes("total");
    const asksPending =
      q.includes("pendiente") ||
      q.includes("pendientes") ||
      q.includes("por asignar") ||
      q.includes("proyeccion");
    const askedTipo = q.includes("penal")
      ? "penal"
      : q.includes("administrativo")
        ? "administrativo"
        : q.includes("procedencia")
          ? "pendiente determinar procedencia"
          : null;
    const askedStatusKeyword = q.includes("estatus") ? q.replace(/.*estatus\s+/, "").trim() : "";
    const askedEstadoKeyword = q.includes("estado") ? q.replace(/.*estado\s+/, "").trim() : "";
    const isFollowUp =
      q.includes("de esos") ||
      q.includes("de esas") ||
      q.includes("de ellos") ||
      q.includes("y cuales") ||
      q.includes("y quienes") ||
      q.includes("y de esos") ||
      q.includes("y de esas");

    const knownResponsables = Array.from(
      new Set(
        [
          ...normalizedAnalysis.map((row) => row._responsable),
          ...normalizedGeneralRecords.flatMap((row) => row._responsables)
        ].filter(Boolean)
      )
    );
    const knownCities = Array.from(new Set(normalizedAnalysis.map((row) => row._ciudad).filter(Boolean)));
    const knownEstatus = Array.from(new Set(normalizedGeneralRecords.map((row) => row._estatus).filter(Boolean)));
    const knownEstado = Array.from(new Set(normalizedGeneralRecords.map((row) => row._estado).filter(Boolean)));
    const matchedResponsable = knownResponsables.find((name) => name && q.includes(name));
    const matchedCity = knownCities.find((name) => name && q.includes(name));
    const matchedEstatus = knownEstatus.find((name) => name && q.includes(name));
    const matchedEstado = knownEstado.find((name) => name && q.includes(name));

    let scopedRows = isFollowUp && chatContextRef.current?.rows?.length
      ? (chatContextRef.current.rows as unknown as typeof normalizedAnalysis)
      : normalizedAnalysis;
    let sourceScope = isFollowUp && chatContextRef.current?.scope ? chatContextRef.current.scope : "analisis";

    if (askedTipo) {
      scopedRows = scopedRows.filter((row) => row._tipo.includes(askedTipo));
    }
    if (matchedResponsable) {
      scopedRows = scopedRows.filter((row) => row._responsable.includes(matchedResponsable));
    }
    if (matchedCity) {
      scopedRows = scopedRows.filter((row) => row._ciudad.includes(matchedCity));
    }
    if (q.includes("vencid")) {
      scopedRows = scopedRows.filter((row) => Number(row.DiasInt) < 0);
    } else if (mentionedDay !== null) {
      if (q.includes("mas de")) {
        scopedRows = scopedRows.filter((row) => Number(row.DiasInt) > mentionedDay);
      } else if (q.includes("menos de")) {
        scopedRows = scopedRows.filter((row) => Number(row.DiasInt) < mentionedDay);
      } else if (q.includes("entre") && daysFound.length >= 2) {
        const minDay = Math.min(daysFound[0], daysFound[1]);
        const maxDay = Math.max(daysFound[0], daysFound[1]);
        scopedRows = scopedRows.filter((row) => Number(row.DiasInt) >= minDay && Number(row.DiasInt) <= maxDay);
      } else {
        scopedRows = scopedRows.filter((row) => Number(row.DiasInt) === mentionedDay);
      }
    }

    if (q.includes("para asignacion")) {
      scopedRows = scopedRows.filter((row) => row._responsable.includes("pendiente por asignar"));
    }
    if (q.includes("proyeccion")) {
      scopedRows = scopedRows.filter((row) => row._responsable.includes("(proyeccion)"));
    }
    if (matchedEstatus || askedStatusKeyword) {
      scopedRows = scopedRows.filter((row) => row._estatus.includes(askedStatusKeyword));
    }
    if (matchedEstado || askedEstadoKeyword) {
      const stateNeedle = matchedEstado || askedEstadoKeyword;
      const rawScoped = isFollowUp && chatContextRef.current?.rows?.length
        ? (chatContextRef.current.rows as unknown as typeof normalizedGeneralRecords)
        : normalizedGeneralRecords;
      const estadoRows = rawScoped.filter((row) => row._estado.includes(stateNeedle));
      if (estadoRows.length) {
        sourceScope = "general";
        if (matchedResponsable) {
          const narrowed = estadoRows.filter((row) => row._responsables.some((name) => name.includes(matchedResponsable)));
          if (narrowed.length) {
            return {
              text: `En el Excel encontre ${narrowed.length} filas con estado ${stateNeedle}. Ejemplos: ${narrowed.slice(0, 6).map((row) => `${row["Cuenta Contrato"] || "Sin cuenta"} | ${row.Estado} | ${row.Estatus || "Sin estatus"}`).join(" | ")}.`,
              context: { rows: narrowed, scope: sourceScope, summary: `Estado ${stateNeedle}` }
            };
          }
        }
      }
    }

    if (q.includes("que hace") || q.includes("como funciona") || q.includes("programa")) {
      return {
        text: "Este programa lee el Excel, construye tableros por tipo de proceso, calcula vencimientos, muestra indicadores, permite filtrar por estatus y estado, exportar resultados y consultar el analisis desde este asistente."
      };
    }

    if (q.includes("leiste el excel") || q.includes("leer el excel") || q.includes("leiste el archivo") || q.includes("archivo cargado")) {
      return {
        text: `Si. El analisis actual sale del Excel cargado en la hoja ${data.sheet_used}. Tengo ${data.source_total_rows} registros fuente, ${data.alerts_total_rows} alertas procesadas y ${statusRecords.length} registros livianos para filtros y consultas.`
      };
    }

    if (q.includes("hoja") || q.includes("sheet")) {
      return {
        text: `La hoja analizada es: ${data.sheet_used}. Registros fuente: ${data.source_total_rows}. Alertas procesadas: ${data.alerts_total_rows}.`
      };
    }

    if (
      q === "tabla general" ||
      q === "tablero general" ||
      q.includes("vista general") ||
      q.includes("resumen del tablero general") ||
      q.includes("que muestra el tablero general")
    ) {
      const lecturaActiva = [
        generalFilterTipo === "TODOS" ? "Todos los tipos" : generalFilterTipo,
        generalFilterEstatus === "TODOS" ? "Todos los estatus" : generalFilterEstatus,
        generalFilterEstado === "TODOS" ? "Todos los estados" : generalFilterEstado,
        generalFilterResponsable === "TODOS" ? "Todos los responsables" : generalFilterResponsable
      ].join(" / ");
      const topResponsables = generalBoard.rows
        .slice(0, 5)
        .map((row) => `${row.responsable}: ${row.total_general}`)
        .join(" | ");
      return {
        text: `Tablero General activo. Lectura: ${lecturaActiva}. Responsables visibles: ${generalBoard.rows.length}. Alertas visibles: ${filteredGeneralBoardRecords.length}. ${topResponsables ? `Responsables mostrados: ${topResponsables}.` : "No hay filas visibles con esos filtros."}`,
        context: {
          rows: filteredGeneralBoardRecords as unknown as Array<Record<string, unknown>>,
          scope: "general_board",
          summary: "Tablero General activo"
        }
      };
    }

    if (q.includes("tableros") || q.includes("tablero")) {
      const names = ["Tablero General", ...data.tableros.map((b) => b.title)].join(" | ");
      return { text: `Tableros disponibles: ${names}.` };
    }

    if (q.includes("resumen") || q.includes("analisis") || q.includes("metricas") || q.includes("estado actual")) {
      return {
        text: `Resumen actual: total ${filteredTotals.alertas_total}, vencidas ${filteredTotals.vencidas}, 0-10 dias ${filteredTotals.por_vencer_0_10}, 11-30 dias ${filteredTotals.rango_11_30}, 31-60 dias ${filteredTotals.rango_31_60}, 61-150 dias ${filteredTotals.rango_61_150}.`
      };
    }

    if (q.includes("pendiente por asignar") || q.includes("asignacion") || q.includes("proyeccion")) {
      let pendientes = 0;
      let proyeccion = 0;
      for (const board of data.tableros) {
        for (const row of board.rows) {
          const label = String(row.responsable || "").toLowerCase();
          if (label.includes("pendiente por asignar")) pendientes += Number(row.total_general || 0);
          if (label.includes("(proyeccion)")) proyeccion += Number(row.total_general || 0);
        }
      }
      return {
        text: `En los tableros actuales hay ${pendientes} casos como "Pendiente por asignar" y ${proyeccion} casos marcados en "Proyeccion".`
      };
    }

    if (q.includes("vencidas")) {
      return { text: `Alertas vencidas en el analisis filtrado: ${filteredTotals.vencidas}.` };
    }

    if (q.includes("0 a 10") || q.includes("por vencer")) {
      return { text: `Alertas en rango 0 a 10 dias: ${filteredTotals.por_vencer_0_10}.` };
    }

    if (matchedResponsable && !asksWho && !asksHowMany) {
      const personRows = normalizedAnalysis.filter((row) => row._responsable.includes(matchedResponsable));
      const vencidas = personRows.filter((row) => Number(row.DiasInt) < 0).length;
      const proximas = personRows.filter((row) => Number(row.DiasInt) >= 0 && Number(row.DiasInt) <= 10).length;
      const tipos = Array.from(new Set(personRows.map((row) => row.Tipo).filter(Boolean))).join(", ");
      return {
        text: `${matchedResponsable} tiene ${personRows.length} alertas en total. Vencidas: ${vencidas}. Entre 0 y 10 dias: ${proximas}. Tipos detectados: ${tipos || "sin tipo"}.`,
        context: { rows: personRows, scope: "analisis", summary: `Responsable ${matchedResponsable}` }
      };
    }

    if (mentionedDay !== null && asksWho) {
      if (!scopedRows.length) {
        return { text: `No encontre responsables con alertas para ${mentionedDay} dias${askedTipo ? ` en ${askedTipo}` : ""}.` };
      }
      const byResp = new Map<string, number>();
      for (const row of scopedRows) {
        byResp.set(row.Responsable, (byResp.get(row.Responsable) ?? 0) + 1);
      }
      const ordered = Array.from(byResp.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8)
        .map(([name, count]) => `${name}: ${count}`)
        .join(" | ");
      return {
        text: `Responsables con alertas para ${mentionedDay} dias${askedTipo ? ` en ${askedTipo}` : ""}: ${ordered}.`,
        context: { rows: scopedRows, scope: sourceScope, summary: `Alertas para ${mentionedDay} dias` }
      };
    }

    if (mentionedDay !== null && asksHowMany) {
      return {
        text: `Hay ${scopedRows.length} alertas${askedTipo ? ` de tipo ${askedTipo}` : ""} para ${mentionedDay} dias.`,
        context: { rows: scopedRows, scope: sourceScope, summary: `Conteo ${mentionedDay} dias` }
      };
    }

    if (asksWho && asksPending) {
      if (!scopedRows.length) {
        return { text: "No encontre responsables para esa condicion." };
      }
      const grouped = new Map<string, number>();
      for (const row of scopedRows) {
        grouped.set(row.Responsable, (grouped.get(row.Responsable) ?? 0) + 1);
      }
      const summary = Array.from(grouped.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([name, count]) => `${name}: ${count}`)
        .join(" | ");
      return {
        text: `Responsables encontrados: ${summary}.`,
        context: { rows: scopedRows, scope: sourceScope, summary: "Responsables pendientes" }
      };
    }

    if ((q.includes("cuenta") || q.includes("contrato")) && daysFound.length) {
      const cuentas = scopedRows
        .slice(0, 10)
        .map((row) => `${row["Cuenta Contrato"]} (${row.Responsable}, ${row.DiasInt} dias)`)
        .join(" | ");
      return cuentas
        ? {
            text: `Cuentas encontradas: ${cuentas}.`,
            context: { rows: scopedRows, scope: sourceScope, summary: "Cuentas filtradas" }
          }
        : { text: "No encontre cuentas contrato con esa condicion." };
    }

    if (q.includes("ciudad")) {
      if (!scopedRows.length) {
        return { text: "No encontre ciudades para esa condicion." };
      }
      const byCity = new Map<string, number>();
      for (const row of scopedRows) {
        byCity.set(row.Ciudad || "Sin ciudad", (byCity.get(row.Ciudad || "Sin ciudad") ?? 0) + 1);
      }
      const summary = Array.from(byCity.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8)
        .map(([name, count]) => `${name}: ${count}`)
        .join(" | ");
      return {
        text: `Distribucion por ciudad: ${summary}.`,
        context: { rows: scopedRows, scope: sourceScope, summary: "Distribucion por ciudad" }
      };
    }

    if (q.includes("top") || q.includes("mas") || q.includes("mayor")) {
      const grouped = new Map<string, number>();
      for (const row of scopedRows) {
        grouped.set(row.Responsable, (grouped.get(row.Responsable) ?? 0) + 1);
      }
      const top = Array.from(grouped.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 6)
        .map(([name, count]) => `${name}: ${count}`)
        .join(" | ");
      if (top) {
        return {
          text: `Top encontrados: ${top}.`,
          context: { rows: scopedRows, scope: sourceScope, summary: "Top responsables" }
        };
      }
    }

    if (q.includes("estatus") || q.includes("estado")) {
      const baseRows = isFollowUp && chatContextRef.current?.rows?.length
        ? (chatContextRef.current.rows as unknown as typeof normalizedGeneralRecords)
        : normalizedGeneralRecords;
      const keyword = matchedEstatus || matchedEstado || askedStatusKeyword || askedEstadoKeyword;
      const found = keyword
        ? baseRows.filter((row) => row._estatus.includes(keyword) || row._estado.includes(keyword))
        : baseRows;
      if (!found.length) {
        return { text: "No encontre registros con ese estatus o estado." };
      }
      const sample = found
        .slice(0, 8)
        .map((row) => `${row["Cuenta Contrato"] || "Sin cuenta"} | ${row.Estatus || "Sin estatus"} | ${row.Estado || "Sin estado"}`)
        .join(" | ");
      return {
        text: `Registros encontrados: ${sample}.`,
        context: { rows: found, scope: "general", summary: `Filtro por estatus/estado ${keyword || ""}`.trim() }
      };
    }

    if (q.includes("columna") || q.includes("columnas") || q.includes("campos")) {
      return { text: `Columnas disponibles en la lectura fuente: ${data.source_columns.join(" | ")}.` };
    }

    if (q.includes("filtro") || q.includes("filtros")) {
      return {
        text: "Puedes filtrar en el programa por tipo, responsable, ciudad, rango de dias, estatus y estado. El Tablero General concentra filtros por tipo, estatus, estado, responsable y busqueda libre."
      };
    }

    if (q.includes("ejemplo") || q.includes("que te puedo preguntar")) {
      return {
        text: "Puedes preguntar: quien tiene 15 dias, cuantos penales hay vencidos, que tiene Lady pendiente, cuales cuentas estan en Bogota, que estados aparecen, cuales son los responsables con mas carga, o y de esos cuales son administrativos."
      };
    }

    return {
      text: "Puedo responder casi cualquier consulta sobre el programa y el analisis cargado. Prueba preguntas sobre responsables, dias, tipos, estatus, estado, cuentas, ciudades, columnas, filtros o usa seguimientos como: y de esos cuales son penales."
    };
  };

  const onSendChat = () => {
    const text = chatInput.trim();
    if (!text) return;
    const reply = buildChatReply(text);
    if (reply.context) {
      chatContextRef.current = reply.context;
    }
    setChatMessages((prev) => [
      ...prev,
      { id: Date.now(), role: "user", text },
      { id: Date.now() + 1, role: "assistant", text: reply.text }
    ]);
    setChatInput("");
  };

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
        setLoadingStage("Analizando hoja y construyendo tableros...");
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
          {loading && (
            <div className="sm:col-span-12">
              <div className={`overflow-hidden rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <div className="flex items-center justify-between gap-3">
                  <div>
                    <p className={`text-sm font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{loadingStage || "Procesando archivo..."}</p>
                    <p className={`mt-1 text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>
                      En archivos grandes puede tardar algunos segundos. El proceso sigue activo aunque el backend este analizando.
                    </p>
                  </div>
                  <p className={`text-sm font-semibold tabular-nums ${darkMode ? "text-brand-200" : "text-brand-700"}`}>
                    {loadingProgress}%
                  </p>
                </div>
                <div className={`mt-3 h-3 overflow-hidden rounded-full ${darkMode ? "bg-slate-800" : "bg-slate-200"}`}>
                  <div
                    className="h-full rounded-full bg-gradient-to-r from-brand-500 via-cyan-400 to-emerald-400 transition-[width] duration-300 ease-out"
                    style={{ width: `${Math.max(6, loadingProgress)}%` }}
                  />
                </div>
              </div>
            </div>
          )}
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
        <section className="card mb-8 p-6">
          <div className="flex flex-col gap-3 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <h2 className={`text-xl font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>
                Tablero General
              </h2>
              <p className={`mt-1 text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>
                Mismo formato de tablero, con filtros para mostrar solo los responsables que cumplen el tipo, estatus y estado seleccionados.
              </p>
            </div>
            <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Tipo</label>
                <select
                  value={generalFilterTipo}
                  onChange={(e) => setGeneralFilterTipo(e.target.value)}
                  className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}
                >
                  <option value="TODOS">Todos</option>
                  {generalTipoOptions.map((item) => <option key={item} value={item}>{item}</option>)}
                </select>
              </div>
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Estatus</label>
                <select
                  value={generalFilterEstatus}
                  onChange={(e) => setGeneralFilterEstatus(e.target.value)}
                  className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}
                >
                  <option value="TODOS">Todos</option>
                  {generalEstatusOptions.map((item) => <option key={item} value={item}>{item}</option>)}
                </select>
              </div>
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Estado</label>
                <select
                  value={generalFilterEstado}
                  onChange={(e) => setGeneralFilterEstado(e.target.value)}
                  className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}
                >
                  <option value="TODOS">Todos</option>
                  {generalEstadoOptions.map((item) => <option key={item} value={item}>{item}</option>)}
                </select>
              </div>
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Responsable</label>
                <select
                  value={generalFilterResponsable}
                  onChange={(e) => setGeneralFilterResponsable(e.target.value)}
                  className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}
                >
                  <option value="TODOS">Todos</option>
                  {generalResponsableOptions.map((item) => <option key={item} value={item}>{item}</option>)}
                </select>
              </div>
              <div>
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Busqueda</label>
                <input
                  type="text"
                  value={generalSearch}
                  onChange={(e) => setGeneralSearch(e.target.value)}
                  placeholder="Cuenta, ciudad, estado..."
                  className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}
                />
              </div>
            </div>
          </div>

          <div className="mt-4 grid gap-3 sm:grid-cols-3">
            <div className={`rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
              <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Responsables visibles</p>
              <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{generalBoard.rows.length}</p>
            </div>
            <div className={`rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
              <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Alertas visibles</p>
              <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{generalBoard.totals.total_general}</p>
            </div>
            <div className={`rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
              <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Lectura activa</p>
              <p className={`mt-2 text-sm font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>
                {generalFilterTipo === "TODOS" ? "Todos los tipos" : generalFilterTipo}
                {" / "}
                {generalFilterEstatus === "TODOS" ? "Todos los estatus" : generalFilterEstatus}
                {" / "}
                {generalFilterEstado === "TODOS" ? "Todos los estados" : generalFilterEstado}
              </p>
            </div>
          </div>
        </section>
      )}

      {data && (
        <section className="mb-8">
          <BoardTable board={generalBoard} darkMode={darkMode} />
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
        <section className="mt-8 space-y-5">
          <section className="card p-5">
            <div className="flex flex-col gap-3 lg:flex-row lg:items-end lg:justify-between">
              <div>
                <h2 className={`text-lg font-semibold ${darkMode ? "text-slate-100" : "text-ink"}`}>
                  Resumen General por Ano
                </h2>
                <p className={`mt-1 text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>
                  Filtra por ano del documento para recalcular estatus, estado y pendientes clave.
                </p>
              </div>
              <div className="w-full max-w-xs">
                <label className={`mb-1 block text-xs font-semibold ${darkMode ? "text-slate-300" : "text-slate-700"}`}>Ano</label>
                <select
                  value={statusYearFilter}
                  onChange={(e) => setStatusYearFilter(e.target.value)}
                  className={`w-full rounded-lg border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 text-slate-800"}`}
                >
                  <option value="TODOS">Todos</option>
                  {statusYearOptions.map((year) => <option key={year} value={String(year)}>{year}</option>)}
                </select>
              </div>
            </div>

            <div className="mt-4 grid gap-3 sm:grid-cols-3">
              <div className={`rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Total de casos</p>
                <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{statusRowsByYear.length}</p>
              </div>
              <div className={`rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Estatus distintos</p>
                <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{filteredStatusAnalysis.estatus_top.length}</p>
              </div>
              <div className={`rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Estados distintos</p>
                <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{filteredStatusAnalysis.estado_top.length}</p>
              </div>
            </div>
          </section>

          <section className="grid gap-6 lg:grid-cols-3">
            <MiniBarChart title="Estatus (Top)" data={filteredStatusAnalysis.estatus_top} darkMode={darkMode} />
            <MiniBarChart title="Estado (Top)" data={filteredStatusAnalysis.estado_top} darkMode={darkMode} />
            <MiniBarChart title="Pendientes Clave (Totales)" data={filteredStatusAnalysis.pendientes_status_totals} darkMode={darkMode} />
          </section>
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
          <DataTable title="Vista previa alertas (30)" rows={data.alerts_preview} darkMode={darkMode} />
        </section>
      )}

      <button
        type="button"
        onClick={() => setChatOpen((prev) => !prev)}
        className={`fixed bottom-5 right-5 z-50 h-14 w-14 rounded-full border text-2xl shadow-xl transition ${
          darkMode
            ? "border-slate-600 bg-slate-900/95 hover:bg-slate-800 text-slate-100"
            : "border-slate-300 bg-white hover:bg-slate-100 text-slate-800"
        }`}
        aria-label="Abrir asistente IA"
      >
        🤖
      </button>

      {chatOpen && (
        <section
          className={`fixed bottom-24 right-5 z-50 flex h-[28rem] w-[22rem] flex-col overflow-hidden rounded-2xl border shadow-2xl ${
            darkMode ? "border-slate-700 bg-slate-950 text-slate-100" : "border-slate-300 bg-white text-slate-800"
          }`}
        >
          <header className={`flex items-center justify-between border-b px-4 py-3 ${darkMode ? "border-slate-700" : "border-slate-200"}`}>
            <p className="text-sm font-semibold">Asistente IA</p>
            <button
              type="button"
              onClick={() => {
                chatContextRef.current = null;
                setChatMessages([{ id: Date.now(), role: "assistant", text: "Chat reiniciado. Preguntame lo que necesites sobre el tablero." }]);
              }}
              className={`rounded-lg px-2 py-1 text-xs ${darkMode ? "bg-slate-800 hover:bg-slate-700" : "bg-slate-100 hover:bg-slate-200"}`}
            >
              Limpiar
            </button>
          </header>
          <div ref={chatScrollRef} className={`flex-1 space-y-3 overflow-y-auto px-3 py-3 ${darkMode ? "bg-slate-950" : "bg-slate-50"}`}>
            {chatMessages.map((m) => (
              <div key={m.id} className={`max-w-[90%] rounded-xl px-3 py-2 text-xs leading-relaxed ${
                m.role === "assistant"
                  ? darkMode
                    ? "bg-slate-800 text-slate-100"
                    : "bg-white text-slate-800 border border-slate-200"
                  : darkMode
                    ? "ml-auto bg-brand-700 text-white"
                    : "ml-auto bg-brand-600 text-white"
              }`}>
                {m.text}
              </div>
            ))}
          </div>
          <div className={`flex gap-2 border-t p-3 ${darkMode ? "border-slate-700" : "border-slate-200"}`}>
            <input
              type="text"
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") onSendChat();
              }}
              placeholder="Pregunta sobre el programa o analisis..."
              className={`flex-1 rounded-lg border px-3 py-2 text-xs ${
                darkMode ? "border-slate-700 bg-slate-900 text-slate-100" : "border-slate-300 bg-white text-slate-800"
              }`}
            />
            <button
              type="button"
              onClick={onSendChat}
              className={`rounded-lg px-3 py-2 text-xs font-semibold text-white ${darkMode ? "bg-brand-600 hover:bg-brand-500" : "bg-brand-700 hover:bg-brand-800"}`}
            >
              Enviar
            </button>
          </div>
        </section>
      )}
    </main>
  );
}
