"use client";

import { toBlob } from "html-to-image";
import { ChangeEvent, FormEvent, ReactNode, useEffect, useId, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type ControlRecord = {
  Responsable: string;
  Fecha_Vencimiento: string;
  DiasInt: number;
  Estatus: string;
  Estado: string;
  Aviso_T2: string;
  Fecha_Aviso: string;
  "Cuenta Contrato": string;
  Anomalia_Visitada: string;
};

type PreviewResponse = {
  sheet_used: string;
  available_sheets: string[];
  source_columns: string[];
  source_total_rows: number;
  all_estatus_options: string[];
  all_estado_options: string[];
  report_records: {
    Aviso_T2: string;
    Fecha_Aviso: string;
    "Cuenta Contrato": string;
    Anomalia_Visita: string;
    Estatus: string;
    Estado: string;
    Observaciones: string;
  }[];
  chart_records: {
    Estatus: string;
    Estado: string;
    Observaciones: string;
  }[];
  source_preview: Record<string, string | number | null>[];
  admin_control_records: ControlRecord[];
  penal_control_records: ControlRecord[];
  no_procedente_control_records: ControlRecord[];
  medidores_total: number;
  medidores_retirado_por: {
    label: string;
    count: number;
    percentage: number;
  }[];
  medidores_concepto: {
    label: string;
    count: number;
    percentage: number;
  }[];
  medidores_pendientes: {
    "Cuenta Contrato": string;
    Medidor: string;
    Diametro: string;
    Lectura: string;
    Retiro: string;
    Concepto: string;
  }[];
  medidores_liquidacion_m3: {
    label: string;
    count: number;
    sum: number;
    pending_count: number;
  }[];
  medidores_pendientes_liquidar: number;
  casos_especiales_records: {
    "N°": string;
    "Intervención": string;
    Zona: string;
    "Porción": string;
    "Dirección": string;
    Localidad: string;
    Barrio: string;
    "Cuenta contrato": string;
    "Hallazgo encontrado": string;
    Interlocutor: string;
    Equipo: string;
    Total_seguimientos: number;
    Ultimo_resultado: string;
    Ultima_observacion: string;
    Tiene_seguimiento: string;
  }[];
  casos_especiales_seguimientos: {
    Zona: string;
    Localidad: string;
    Barrio: string;
    "Hallazgo encontrado": string;
    "Intervención": string;
    "Cuenta contrato": string;
    Resultado: string;
    "Observación": string;
  }[];
  casos_especiales_filter_options: {
    zona: string[];
    localidad: string[];
    barrio: string[];
    hallazgo: string[];
    resultado: string[];
  };
  casos_inicial_recuperacion_records: {
    N: string;
    "Actividad economica": string;
    Nombre: string;
    Direccion: string;
    Localidad: string;
    Barrio: string;
    "Cuenta contrato": string;
    Aviso: string;
    Fecha: string;
    "Lectura intervencion m3": string;
    "Personal visita": string;
    "Hallazgos encontrados": string;
    "Deuda 2025": string;
    "Ultima lectura": string;
    "Situacion predio": string;
    "Surtido predio": string;
    "Volumen recuperado": string;
    "Valor recuperado": string;
    "Estado deuda": string;
    "Liquidacion m3": string;
    "Liquidacion $": string;
    "Accion operativa": string;
    "Accion administrativa": string;
    "Accion penal": string;
    Observaciones: string;
    "Acto de suspension": string;
    "Avisos reincidencia": string;
    "Observaciones reincidencia": string;
  }[];
  casos_inicial_recuperacion_filter_options: {
    localidad: string[];
    barrio: string[];
    hallazgo: string[];
    situacion: string[];
    accion_administrativa: string[];
    estado_deuda: string[];
  };
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

type BoardRow = {
  responsable: string;
  vencidos: number;
  total_general: number;
  counts: Record<string, number>;
};

type BoardData = {
  title: string;
  description: string;
  row_label: string;
  accent: "teal" | "amber" | "rose";
  column_type: "days" | "dates";
  warning_after: number | null;
  overdue_after: number | null;
  day_columns: Array<number | string>;
  rows: BoardRow[];
  totals: {
    vencidos: number;
    total_general: number;
    counts: Record<string, number>;
  };
};

type MultiSelectFilterProps = {
  label: string;
  options: string[];
  selected: string[];
  onChange: (next: string[]) => void;
  darkMode: boolean;
};

type ChartPoint = {
  label: string;
  count: number;
};

type BaseMode = "administrativa" | "medidores" | "casos_especiales";

const API_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";
const DEFAULT_ADMIN_ESTATUS = [
  "Para administrativo",
  "Para administrativo/Para expediente",
  "Para administrativo/Penal"
];
const DEFAULT_PENAL_ESTATUS = [
  "Para administrativo/Para expediente",
  "Para expediente"
];
const DEFAULT_SHEET_BY_MODE: Record<BaseMode, string> = {
  administrativa: "Procesos Adminis_Penal",
  medidores: "Acueducto",
  casos_especiales: "3. Visitas seguimiento"
};

function normalizeForSearch(value: string | number | null | undefined) {
  return String(value ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function parseLocaleNumber(value: string | number | null | undefined) {
  const text = String(value ?? "").trim();
  const normalized = normalizeForSearch(text);
  if (!text || ["nan", "none", "null", "n/a", "na", "-", "pendiente"].includes(normalized)) return null;
  const cleaned = text
    .replace(/\$/g, "")
    .replace(/cop/gi, "")
    .replace(/m3/gi, "")
    .replace(/\s+/g, "");
  let parsed = cleaned;
  if (parsed.includes(",") && parsed.includes(".")) {
    if (parsed.lastIndexOf(",") > parsed.lastIndexOf(".")) {
      parsed = parsed.replace(/\./g, "").replace(",", ".");
    } else {
      parsed = parsed.replace(/,/g, "");
    }
  } else if (parsed.includes(",")) {
    parsed = parsed.replace(",", ".");
  }
  const valueNumber = Number(parsed);
  return Number.isFinite(valueNumber) ? valueNumber : null;
}

function hasMeaningfulValue(value: string | number | null | undefined) {
  const normalized = normalizeForSearch(value);
  return Boolean(normalized) && !["nan", "none", "null", "n/a", "na", "-", "undefined"].includes(normalized);
}

function isNoApplyValue(value: string | number | null | undefined) {
  const normalized = normalizeForSearch(value);
  return normalized.includes("no aplica") || normalized === "pendiente";
}

function compactDayColumns(days: number[]) {
  const values = Array.from(new Set(days.filter((day) => Number.isFinite(day) && day >= 0))).sort((a, b) => a - b);
  if (values.length <= 32) return values;
  const head = values.slice(0, 24);
  const milestones = [30, 45, 60, 90, 120, 150];
  const tail = milestones.filter((day) => values.includes(day));
  return Array.from(new Set([...head, ...tail])).sort((a, b) => a - b);
}

function buildBoardFromRecords(
  records: { Responsable: string; DiasInt: number }[],
  title: string,
  description: string,
  rowLabel = "Responsable",
  accent: BoardData["accent"] = "teal",
  warningAfter: number | null = null,
  overdueAfter: number | null = null
): BoardData {
  if (!records.length) {
    return {
      title,
      description,
      row_label: rowLabel,
      accent,
      column_type: "days",
      warning_after: warningAfter,
      overdue_after: overdueAfter,
      day_columns: [],
      rows: [],
      totals: { vencidos: 0, total_general: 0, counts: {} }
    };
  }

  const visibleDayValues = records
    .map((row) => Number(row.DiasInt))
    .filter((day) => Number.isFinite(day))
    .filter((day) => day >= 0)
    .filter((day) => overdueAfter === null || day <= overdueAfter);
  const dayColumns = compactDayColumns(visibleDayValues);
  const grouped = new Map<string, { Responsable: string; DiasInt: number }[]>();
  for (const row of records) {
    const key = row.Responsable || "Sin responsable";
    const current = grouped.get(key) ?? [];
    current.push(row);
    grouped.set(key, current);
  }

  const responsibles = Array.from(grouped.keys()).sort((a, b) => {
    if (a === "Pendiente por asignar") return 1;
    if (b === "Pendiente por asignar") return -1;
    if (a === "Sin responsable") return 1;
    if (b === "Sin responsable") return -1;
    return a.localeCompare(b, "es");
  });

  const totalCounts: Record<string, number> = {};
  for (const day of dayColumns) totalCounts[String(day)] = 0;

  let totalVencidos = 0;
  let totalGeneral = 0;
  const rows: BoardRow[] = responsibles.map((responsable) => {
    const entries = grouped.get(responsable) ?? [];
    const counts: Record<string, number> = {};
    for (const day of dayColumns) counts[String(day)] = 0;
    let vencidos = 0;
    for (const entry of entries) {
      const isOverdueByThreshold = overdueAfter !== null && entry.DiasInt > overdueAfter;
      if (entry.DiasInt < 0 || isOverdueByThreshold) {
        vencidos += 1;
      } else if (dayColumns.includes(entry.DiasInt)) {
        counts[String(entry.DiasInt)] += 1;
        totalCounts[String(entry.DiasInt)] += 1;
      }
    }
    totalVencidos += vencidos;
    totalGeneral += entries.length;
    return {
      responsable,
      vencidos,
      total_general: entries.length,
      counts
    };
  });

  return {
    title,
    description,
    row_label: rowLabel,
    accent,
    column_type: "days",
    warning_after: warningAfter,
    overdue_after: overdueAfter,
    day_columns: dayColumns,
    rows,
    totals: {
      vencidos: totalVencidos,
      total_general: totalGeneral,
      counts: totalCounts
    }
  };
}

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

function BoardTable({
  board,
  darkMode,
  onDownloadSummary,
  downloadingSummary = false,
  extraControls
}: {
  board: BoardData;
  darkMode: boolean;
  onDownloadSummary?: () => void;
  downloadingSummary?: boolean;
  extraControls?: ReactNode;
}) {
  const [showHelp, setShowHelp] = useState(false);
  const [captureMessage, setCaptureMessage] = useState<string | null>(null);
  const [capturing, setCapturing] = useState(false);
  const boardRef = useRef<HTMLElement | null>(null);
  const accentStyles = {
    teal: {
      line: darkMode ? "via-cyan-300/60" : "via-cyan-500/60",
      eyebrow: darkMode ? "text-cyan-200/70" : "text-cyan-800/70",
      actionCopy: darkMode
        ? "border-cyan-900/60 bg-slate-900/85 text-slate-100 hover:border-cyan-700 hover:bg-slate-800"
        : "border-cyan-200 bg-white/90 text-slate-700 hover:border-cyan-300 hover:bg-cyan-50/70",
      actionSave: darkMode
        ? "border-emerald-900/60 bg-slate-900/85 text-slate-100 hover:border-emerald-700 hover:bg-slate-800"
        : "border-emerald-200 bg-white/90 text-slate-700 hover:border-emerald-300 hover:bg-emerald-50/70",
      actionHelp: darkMode
        ? "border-sky-900/60 bg-slate-900/85 text-slate-100 hover:border-sky-700 hover:bg-slate-800"
        : "border-sky-200 bg-white/90 text-slate-700 hover:border-sky-300 hover:bg-sky-50/70",
      shell: darkMode ? "border-slate-700/80 bg-slate-950/45" : "border-slate-200 bg-white/80",
      headWrap: darkMode ? "bg-slate-950/85" : "bg-cyan-50/80",
      headCell: darkMode ? "bg-slate-950 text-cyan-200" : "bg-cyan-50/95 text-cyan-950",
      headText: darkMode ? "text-cyan-200" : "text-cyan-950"
    },
    amber: {
      line: darkMode ? "via-amber-300/60" : "via-amber-500/60",
      eyebrow: darkMode ? "text-amber-200/75" : "text-amber-800/75",
      actionCopy: darkMode
        ? "border-amber-900/60 bg-slate-900/85 text-slate-100 hover:border-amber-700 hover:bg-slate-800"
        : "border-amber-200 bg-white/90 text-slate-700 hover:border-amber-300 hover:bg-amber-50/70",
      actionSave: darkMode
        ? "border-orange-900/60 bg-slate-900/85 text-slate-100 hover:border-orange-700 hover:bg-slate-800"
        : "border-orange-200 bg-white/90 text-slate-700 hover:border-orange-300 hover:bg-orange-50/70",
      actionHelp: darkMode
        ? "border-yellow-900/60 bg-slate-900/85 text-slate-100 hover:border-yellow-700 hover:bg-slate-800"
        : "border-yellow-200 bg-white/90 text-slate-700 hover:border-yellow-300 hover:bg-yellow-50/70",
      shell: darkMode ? "border-slate-700/80 bg-slate-950/45" : "border-amber-100 bg-white/85",
      headWrap: darkMode ? "bg-slate-950/85" : "bg-amber-50/80",
      headCell: darkMode ? "bg-slate-950 text-amber-200" : "bg-amber-50/95 text-amber-950",
      headText: darkMode ? "text-amber-200" : "text-amber-950"
    },
    rose: {
      line: darkMode ? "via-rose-300/60" : "via-rose-500/60",
      eyebrow: darkMode ? "text-rose-200/75" : "text-rose-800/75",
      actionCopy: darkMode
        ? "border-rose-900/60 bg-slate-900/85 text-slate-100 hover:border-rose-700 hover:bg-slate-800"
        : "border-rose-200 bg-white/90 text-slate-700 hover:border-rose-300 hover:bg-rose-50/70",
      actionSave: darkMode
        ? "border-fuchsia-900/60 bg-slate-900/85 text-slate-100 hover:border-fuchsia-700 hover:bg-slate-800"
        : "border-fuchsia-200 bg-white/90 text-slate-700 hover:border-fuchsia-300 hover:bg-fuchsia-50/70",
      actionHelp: darkMode
        ? "border-pink-900/60 bg-slate-900/85 text-slate-100 hover:border-pink-700 hover:bg-slate-800"
        : "border-pink-200 bg-white/90 text-slate-700 hover:border-pink-300 hover:bg-pink-50/70",
      shell: darkMode ? "border-slate-700/80 bg-slate-950/45" : "border-rose-100 bg-white/85",
      headWrap: darkMode ? "bg-slate-950/85" : "bg-rose-50/80",
      headCell: darkMode ? "bg-slate-950 text-rose-200" : "bg-rose-50/95 text-rose-950",
      headText: darkMode ? "text-rose-200" : "text-rose-950"
    }
  }[board.accent];

  const captureBoard = async (): Promise<Blob> => {
    if (!boardRef.current) throw new Error("No se encontr? el tablero.");
    const blob = await toBlob(boardRef.current, {
      cacheBust: true,
      pixelRatio: 2,
      backgroundColor: darkMode ? "#0f172a" : "#ffffff",
      filter: (node) => {
        if (!(node instanceof HTMLElement)) return true;
        if (node.dataset.html2canvasIgnore === "true") return false;
        return true;
      },
      style: {
        backdropFilter: "none"
      }
    });
    if (!blob) throw new Error("No fue posible generar la imagen.");
    return blob;
  };

  const onCopyBoard = async () => {
    try {
      setCapturing(true);
      setCaptureMessage(null);
      const blob = await captureBoard();
      if (!window.ClipboardItem || !navigator.clipboard?.write) {
        throw new Error("Tu navegador no permite copiar im?genes al portapapeles.");
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
      link.download = `${board.title.toLowerCase().replace(/[^a-z0-9]+/gi, "_")}.png`;
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

  const boardHelp = `${board.description} Los n?meros dentro de cada celda indican cu?ntos casos tiene ese responsable para ese d?a exacto. La columna Vencidos agrupa d?as negativos y Total general suma todos los casos visibles del responsable.`;

  return (
    <section ref={boardRef} className="card panel-grid relative mb-8 overflow-hidden p-6 sm:p-7">
      <div className={`absolute inset-x-0 top-0 h-px bg-gradient-to-r from-transparent ${accentStyles.line} to-transparent`} />
      <div className="mb-5 flex items-start justify-between gap-3">
        <div>
          <p className={`text-[11px] font-semibold uppercase tracking-[0.28em] ${accentStyles.eyebrow}`}>Tablero operativo</p>
          <h2 className={`mt-2 text-xl font-semibold tracking-tight ${darkMode ? "text-slate-50" : "text-slate-900"}`}>{board.title}</h2>
          <p className={`mt-2 max-w-3xl text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>{board.description}</p>
          {extraControls && <div className="mt-3">{extraControls}</div>}
        </div>
        <div className="flex items-center gap-2">
          {onDownloadSummary && (
            <button type="button" onClick={onDownloadSummary} disabled={downloadingSummary} className={`rounded-full border px-3.5 py-1.5 text-xs font-semibold transition disabled:opacity-60 ${accentStyles.actionCopy}`}>
              {downloadingSummary ? "Generando..." : "Descargar resumen"}
            </button>
          )}
          <button type="button" onClick={onCopyBoard} disabled={capturing} className={`rounded-full border px-3.5 py-1.5 text-xs font-semibold transition disabled:opacity-60 ${accentStyles.actionCopy}`}>
            Copiar
          </button>
          <button type="button" onClick={onSaveBoard} disabled={capturing} className={`rounded-full border px-3.5 py-1.5 text-xs font-semibold transition disabled:opacity-60 ${accentStyles.actionSave}`}>
            Guardar
          </button>
          <button type="button" onClick={() => setShowHelp((prev) => !prev)} className={`rounded-full border px-3.5 py-1.5 text-xs font-semibold transition ${accentStyles.actionHelp}`}>
            ? Ayuda
          </button>
        </div>
      </div>
      {captureMessage && <p className={`mb-3 text-xs ${darkMode ? "text-slate-300" : "text-slate-600"}`}>{captureMessage}</p>}
      {showHelp && (
        <div className={`absolute right-6 top-24 z-30 w-[24rem] rounded-[24px] border p-4 shadow-2xl ${darkMode ? "border-slate-700 bg-slate-950/95 text-slate-200" : "border-slate-200 bg-white/95 text-slate-700"}`}>
          <p className={`text-sm font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>C?mo leer este tablero</p>
          <p className="mt-2 text-xs leading-relaxed">{boardHelp}</p>
        </div>
      )}
      <div className={`overflow-auto rounded-[22px] border ${accentStyles.shell}`}>
        <table className="min-w-full text-xs">
          <thead className={accentStyles.headWrap}>
            <tr>
              <th className={`sticky left-0 z-20 px-3 py-3 text-left font-semibold ${accentStyles.headCell}`}>
                {board.row_label}
              </th>
              {board.day_columns.map((day) => {
                const isWarningDay =
                  board.column_type === "days" &&
                  typeof day === "number" &&
                  board.warning_after !== null &&
                  day > board.warning_after;
                return (
                  <th
                    key={String(day)}
                    className={`px-2 py-3 text-center font-semibold ${
                      isWarningDay
                        ? darkMode
                          ? "bg-amber-950/65 text-amber-200"
                          : "bg-amber-100/95 text-amber-900"
                        : accentStyles.headText
                    }`}
                  >
                    {day}
                  </th>
                );
              })}
              <th className={`px-2 py-3 text-center font-semibold ${darkMode ? "text-rose-300" : "text-rose-700"}`}>Vencidos</th>
              <th className={`px-2 py-3 text-center font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Total general</th>
            </tr>
          </thead>
          <tbody>
            {board.rows.map((row) => (
              <tr key={row.responsable} className={`${darkMode ? "border-slate-800/90 odd:bg-slate-900/20 even:bg-slate-900/55" : "border-slate-100 odd:bg-white/95 even:bg-slate-50/65"} border-t`}>
                <td className={`sticky left-0 z-10 px-3 py-2.5 font-semibold ${darkMode ? "bg-slate-900/95 text-slate-200" : "bg-white/95 text-slate-800"}`}>
                  {row.responsable}
                </td>
                {board.day_columns.map((day) => {
                  const value = row.counts[String(day)] ?? 0;
                  const isWarningDay =
                    board.column_type === "days" &&
                    typeof day === "number" &&
                    board.warning_after !== null &&
                    day > board.warning_after;
                  const isUrgentDay =
                    board.column_type === "days" &&
                    typeof day === "number" &&
                    day <= 10;
                  return (
                    <td
                      key={`${row.responsable}-${String(day)}`}
                      className={`px-2 py-2.5 text-center ${darkMode ? "text-slate-300" : "text-slate-700"} ${
                        isWarningDay && value > 0
                          ? darkMode
                            ? "bg-amber-950/60 text-amber-200 font-semibold"
                            : "bg-yellow-100/95 text-amber-900 font-semibold"
                          : isUrgentDay && value > 0
                            ? darkMode
                              ? "bg-amber-900/35 text-amber-200 font-semibold"
                              : "bg-amber-100/90 text-amber-900 font-semibold"
                            : ""
                      }`}
                    >
                      {value || ""}
                    </td>
                  );
                })}
                <td className={`px-2 py-2.5 text-center font-semibold ${row.vencidos > 0 ? (darkMode ? "bg-rose-900/45 text-rose-200" : "bg-rose-100/90 text-rose-800") : (darkMode ? "text-slate-300" : "text-slate-700")}`}>
                  {row.vencidos || ""}
                </td>
                <td className={`px-2 py-2.5 text-center font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{row.total_general}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className={darkMode ? "border-t border-slate-600 bg-slate-950/90" : "border-t border-slate-300 bg-slate-100/90"}>
              <td className={`sticky left-0 z-10 px-3 py-3 font-bold ${darkMode ? "bg-slate-950 text-slate-100" : "bg-slate-100 text-slate-900"}`}>Total general</td>
              {board.day_columns.map((day) => {
                const totalValue = board.totals.counts[String(day)] || "";
                const isWarningDay =
                  board.column_type === "days" &&
                  typeof day === "number" &&
                  board.warning_after !== null &&
                  day > board.warning_after;
                return (
                  <td
                    key={`tot-${String(day)}`}
                    className={`px-2 py-3 text-center font-bold ${
                      isWarningDay && totalValue
                        ? darkMode
                          ? "bg-amber-950/55 text-amber-200"
                          : "bg-yellow-100/95 text-amber-900"
                        : darkMode
                          ? "text-slate-200"
                          : "text-slate-900"
                    }`}
                  >
                    {totalValue}
                  </td>
                );
              })}
              <td className={`px-2 py-3 text-center font-bold ${darkMode ? "text-rose-200" : "text-rose-800"}`}>{board.totals.vencidos || ""}</td>
              <td className={`px-2 py-3 text-center font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{board.totals.total_general}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </section>
  );
}

function MiniBarChart({
  title,
  data,
  darkMode,
  filterValue,
  onFilterChange,
  filterOptions,
  filterLabel = "Filtro",
  allowWrapLabels = false
}: {
  title: string;
  data: ChartPoint[];
  darkMode: boolean;
  filterValue?: string;
  onFilterChange?: (value: string) => void;
  filterOptions?: string[];
  filterLabel?: string;
  allowWrapLabels?: boolean;
}) {
  const COLLAPSED_LIMIT = 12;
  const [expanded, setExpanded] = useState(false);
  useEffect(() => {
    setExpanded(false);
  }, [data, filterValue]);
  const maxValue = Math.max(...data.map((item) => item.count), 1);
  const visibleData = expanded ? data : data.slice(0, COLLAPSED_LIMIT);
  const canExpand = data.length > COLLAPSED_LIMIT;
  const chartAccent =
    title.includes("Estatus")
      ? {
          line: darkMode ? "via-cyan-300/60" : "via-cyan-500/60",
          bar: darkMode ? "from-cyan-400 via-sky-400 to-brand-500" : "from-cyan-500 via-sky-500 to-brand-600",
          glow: darkMode ? "shadow-[0_12px_30px_-16px_rgba(34,211,238,0.55)]" : "shadow-[0_12px_30px_-16px_rgba(14,165,233,0.35)]"
        }
      : title.includes("Estado")
        ? {
            line: darkMode ? "via-amber-300/60" : "via-amber-500/60",
            bar: darkMode ? "from-amber-300 via-orange-400 to-yellow-500" : "from-amber-400 via-orange-500 to-yellow-500",
            glow: darkMode ? "shadow-[0_12px_30px_-16px_rgba(251,191,36,0.4)]" : "shadow-[0_12px_30px_-16px_rgba(245,158,11,0.28)]"
          }
        : {
            line: darkMode ? "via-rose-300/60" : "via-rose-500/60",
            bar: darkMode ? "from-rose-400 via-fuchsia-400 to-pink-500" : "from-rose-500 via-pink-500 to-fuchsia-500",
            glow: darkMode ? "shadow-[0_12px_30px_-16px_rgba(244,114,182,0.38)]" : "shadow-[0_12px_30px_-16px_rgba(244,63,94,0.28)]"
          };
  return (
    <section className="card relative overflow-hidden p-5">
      <div className={`absolute inset-x-6 top-0 h-px bg-gradient-to-r from-transparent ${chartAccent.line} to-transparent`} />
      <div className="mb-4 flex items-end justify-between gap-3">
        <h3 className={`text-base font-semibold tracking-tight ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{title}</h3>
        {onFilterChange && filterOptions && (
          <div className="w-44">
            <label className={`mb-1.5 block text-[11px] font-semibold uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>{filterLabel}</label>
            <select
              value={filterValue ?? ""}
              onChange={(event) => onFilterChange(event.target.value)}
              className={`w-full rounded-2xl border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 bg-white/90 text-slate-800"}`}
            >
              <option value="">Todos</option>
              {filterOptions.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </div>
        )}
      </div>
      <div className="space-y-2">
        {data.length === 0 && <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Sin datos.</p>}
        {visibleData.map((item) => (
          <div key={`${title}-${item.label}`} className="space-y-1.5">
            <div className="flex items-center justify-between text-xs">
              <span className={`${allowWrapLabels ? "line-clamp-2 whitespace-normal break-words" : "truncate"} pr-2 ${darkMode ? "text-slate-300" : "text-slate-700"}`}>{item.label}</span>
              <span className={`font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{item.count}</span>
            </div>
            <div className={`h-2.5 rounded-full ${darkMode ? "bg-slate-800" : "bg-slate-200"}`}>
              <div
                className={`h-2.5 rounded-full bg-gradient-to-r ${chartAccent.bar} ${chartAccent.glow}`}
                style={{ width: `${(item.count / maxValue) * 100}%` }}
              />
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

function MultiSelectFilter({ label, options, selected, onChange, darkMode }: MultiSelectFilterProps) {
  const instanceId = useId();
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const selectedLabel =
    selected.length === 0 ? "Todos" : selected.length === 1 ? selected[0] : `${selected.length} seleccionados`;

  useEffect(() => {
    const handleDocumentClick = (event: MouseEvent) => {
      if (!rootRef.current) return;
      if (!rootRef.current.contains(event.target as Node)) {
        setOpen(false);
      }
    };

    const handleOpenEvent = (event: Event) => {
      const customEvent = event as CustomEvent<string>;
      if (customEvent.detail !== instanceId) {
        setOpen(false);
      }
    };

    document.addEventListener("mousedown", handleDocumentClick);
    window.addEventListener("multiselect-open", handleOpenEvent as EventListener);
    return () => {
      document.removeEventListener("mousedown", handleDocumentClick);
      window.removeEventListener("multiselect-open", handleOpenEvent as EventListener);
    };
  }, [instanceId]);

  const toggleOption = (option: string) => {
    const exists = selected.includes(option);
    if (exists) {
      onChange(selected.filter((item) => item !== option));
      return;
    }
    onChange([...selected, option]);
  };

  return (
    <div ref={rootRef} className={`relative ${open ? "z-[200]" : "z-10"}`}>
      <label className={`mb-1.5 block text-[11px] font-semibold uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>{label}</label>
      <button
        type="button"
        onClick={() => {
          setOpen((prev) => {
            const next = !prev;
            if (next) {
              window.dispatchEvent(new CustomEvent("multiselect-open", { detail: instanceId }));
            }
            return next;
          });
        }}
        className={`flex w-full items-center justify-between rounded-2xl border px-4 py-3 text-sm shadow-sm transition ${open ? (darkMode ? "border-amber-300/70 shadow-[0_0_0_3px_rgba(251,191,36,0.10)]" : "border-amber-400/80 shadow-[0_0_0_3px_rgba(245,158,11,0.08)]") : ""} ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100 hover:border-slate-500" : "border-slate-300 bg-white/90 text-slate-800 hover:border-slate-400"}`}
      >
        <span className="truncate">{selectedLabel}</span>
        <span className={`ml-3 text-[10px] transition ${open ? "rotate-180" : ""}`}>▼</span>
      </button>
      {open && (
        <div className={`absolute left-0 top-full z-[220] mt-3 max-h-80 w-full overflow-hidden rounded-[22px] border shadow-2xl ${darkMode ? "border-slate-700 bg-slate-950" : "border-slate-200 bg-white"}`}>
          <button
            type="button"
            onClick={() => onChange([])}
            className={`flex w-full items-center justify-between px-4 py-3 text-left text-sm font-semibold ${darkMode ? "border-b border-slate-800 text-slate-100 hover:bg-slate-900" : "border-b border-slate-200 text-slate-800 hover:bg-slate-50"}`}
          >
            <span>Todos</span>
            {selected.length === 0 && <span>✓</span>}
          </button>
          <div className="max-h-64 overflow-auto py-1">
            {options.map((option) => {
              const checked = selected.includes(option);
              return (
                <label
                  key={option}
                  className={`flex cursor-pointer items-center gap-3 px-4 py-2.5 text-sm ${darkMode ? "text-slate-200 hover:bg-slate-900" : "text-slate-700 hover:bg-slate-50"}`}
                >
                  <input type="checkbox" checked={checked} onChange={() => toggleOption(option)} className="h-4 w-4 rounded border-slate-400" />
                  <span className="min-w-0 flex-1 break-words">{option}</span>
                </label>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

export default function HomePage() {
  const [baseMode, setBaseMode] = useState<BaseMode | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [localSheetOptions, setLocalSheetOptions] = useState<string[]>([]);
  const [sharepointUrl, setSharepointUrl] = useState("");
  const [sheetName, setSheetName] = useState(DEFAULT_SHEET_BY_MODE.administrativa);
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
  const [filterEstatus, setFilterEstatus] = useState<string[]>([]);
  const [filterEstado, setFilterEstado] = useState<string[]>([]);
  const [estatusPanelFilter, setEstatusPanelFilter] = useState("");
  const [exportingReport, setExportingReport] = useState(false);
  const [casosZonaFilter, setCasosZonaFilter] = useState("Todos");
  const [casosLocalidadFilter, setCasosLocalidadFilter] = useState("Todos");
  const [casosHallazgoFilter, setCasosHallazgoFilter] = useState("Todos");
  const [casosResultadoFilter, setCasosResultadoFilter] = useState("Todos");
  const [casosInicialLocalidadFilter, setCasosInicialLocalidadFilter] = useState<string[]>([]);
  const [casosInicialBarrioFilter, setCasosInicialBarrioFilter] = useState<string[]>([]);
  const [casosInicialHallazgoFilter, setCasosInicialHallazgoFilter] = useState<string[]>([]);
  const [casosInicialSituacionFilter, setCasosInicialSituacionFilter] = useState<string[]>([]);
  const [casosInicialAccionAdminFilter, setCasosInicialAccionAdminFilter] = useState<string[]>([]);
  const [casosInicialEstadoDeudaFilter, setCasosInicialEstadoDeudaFilter] = useState<string[]>([]);
  const [noProcedenteThreshold, setNoProcedenteThreshold] = useState(60);

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
    setFilterEstatus([]);
    setFilterEstado([]);
    setEstatusPanelFilter("");
  }, [data]);

  useEffect(() => {
    setCasosZonaFilter("Todos");
    setCasosLocalidadFilter("Todos");
    setCasosHallazgoFilter("Todos");
    setCasosResultadoFilter("Todos");
    setCasosInicialLocalidadFilter([]);
    setCasosInicialBarrioFilter([]);
    setCasosInicialHallazgoFilter([]);
    setCasosInicialSituacionFilter([]);
    setCasosInicialAccionAdminFilter([]);
    setCasosInicialEstadoDeudaFilter([]);
  }, [data, baseMode]);

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
        setLoadingStage("Leyendo hoja y construyendo el análisis...");
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
        const payload = xhr.response && typeof xhr.response === "object" ? xhr.response : JSON.parse(xhr.responseText || "{}");
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
    if (!baseMode) {
      setError("Selecciona primero el tipo de base.");
      return;
    }
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
    formData.append("base_mode", baseMode);

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
    if (!baseMode) {
      setDiagnosticError("Selecciona primero el tipo de base.");
      setDiagnosticData(null);
      return;
    }
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

  const downloadBoardSummary = async (
    rows: ControlRecord[],
    responsableHeader: string,
    filename: string
  ) => {
    try {
      setExportingReport(true);
      const exportRows = rows.map((row) => ({
        Aviso_T2: row.Aviso_T2,
        Fecha_Aviso: row.Fecha_Aviso,
        "Cuenta Contrato": row["Cuenta Contrato"],
        Anomalia_Visitada: row.Anomalia_Visitada,
        [responsableHeader]: row.Responsable,
        Dias: String(row.DiasInt),
        Estatus: row.Estatus,
        Estado: row.Estado,
      }));
      const response = await fetch(`${API_URL}/api/report/download`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ records: exportRows, filename }),
      });
      if (!response.ok) {
        const payload = (await response.json().catch(() => ({}))) as { detail?: string };
        throw new Error(payload.detail ?? "No se pudo generar el resumen.");
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : "No se pudo descargar el resumen.");
    } finally {
      setExportingReport(false);
    }
  };

  const downloadMedidoresPendientes = async () => {
    try {
      setExportingReport(true);
      const response = await fetch(`${API_URL}/api/report/download`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          records: medidoresPendientes,
          filename: "medidores_pendientes.xlsx",
        }),
      });
      if (!response.ok) {
        const payload = (await response.json().catch(() => ({}))) as { detail?: string };
        throw new Error(payload.detail ?? "No se pudo generar el archivo.");
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "medidores_pendientes.xlsx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : "No se pudo descargar el resumen.");
    } finally {
      setExportingReport(false);
    }
  };

  const onFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const nextFile = event.target.files?.[0] ?? null;
    setFile(nextFile);
    setError(null);

    if (!nextFile) {
      setLocalSheetOptions([]);
      return;
    }

    try {
      const buffer = await nextFile.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetNames = workbook.SheetNames ?? [];
      setLocalSheetOptions(sheetNames);
      if (sheetNames.length > 0) {
        const preferred = sheetNames.find((item) => normalizeForSearch(item) === normalizeForSearch(sheetName));
        setSheetName(preferred ?? sheetNames[0]);
      }
    } catch {
      setLocalSheetOptions([]);
    }
  };

  const onSelectBaseMode = (mode: BaseMode) => {
    setBaseMode(mode);
    setSheetName(DEFAULT_SHEET_BY_MODE[mode]);
    setFile(null);
    setLocalSheetOptions([]);
    setSharepointUrl("");
    setData(null);
    setError(null);
    setDiagnosticData(null);
    setDiagnosticError(null);
  };

  const adminRecords = data?.admin_control_records ?? [];
  const penalRecords = data?.penal_control_records ?? [];
  const noProcedenteRecords = data?.no_procedente_control_records ?? [];
  const chartRecords = data?.chart_records ?? [];
  const medidoresRetiradoPor = data?.medidores_retirado_por ?? [];
  const medidoresConcepto = data?.medidores_concepto ?? [];
  const medidoresPendientes = data?.medidores_pendientes ?? [];
  const medidoresLiquidacionM3 = data?.medidores_liquidacion_m3 ?? [];
  const medidoresPendientesLiquidar = data?.medidores_pendientes_liquidar ?? 0;
  const medidoresTotal = data?.medidores_total ?? 0;
  const casosEspecialesRecords = data?.casos_especiales_records ?? [];
  const casosEspecialesSeguimientos = data?.casos_especiales_seguimientos ?? [];
  const casosInicialRecuperacionRecords = useMemo(() => {
    const rawRows = data?.casos_inicial_recuperacion_records ?? [];
    const getValue = (row: Record<string, string>, candidates: string[]) => {
      for (const candidate of candidates) {
        const value = row[candidate];
        if (value !== undefined && value !== null && String(value).trim() !== "") {
          return String(value).trim();
        }
      }
      return "";
    };
    return rawRows.map((row) => ({
      "N°": getValue(row as unknown as Record<string, string>, ["N°", "N", "N?", "No"]),
      "Actividad económica": getValue(row as unknown as Record<string, string>, ["Actividad económica", "Actividad economica", "Actividad econ?mica"]),
      Nombre: getValue(row as unknown as Record<string, string>, ["Nombre"]),
      Dirección: getValue(row as unknown as Record<string, string>, ["Dirección", "Direccion", "Direcci?n"]),
      Localidad: getValue(row as unknown as Record<string, string>, ["Localidad"]),
      Barrio: getValue(row as unknown as Record<string, string>, ["Barrio"]),
      "Cuenta contrato": getValue(row as unknown as Record<string, string>, ["Cuenta contrato"]),
      Aviso: getValue(row as unknown as Record<string, string>, ["Aviso"]),
      Fecha: getValue(row as unknown as Record<string, string>, ["Fecha"]),
      "Lectura intervención m3": getValue(row as unknown as Record<string, string>, ["Lectura intervención m3", "Lectura intervencion m3", "Lectura intervenci?n m3"]),
      "Personal visita": getValue(row as unknown as Record<string, string>, ["Personal visita"]),
      "Hallazgos encontrados": getValue(row as unknown as Record<string, string>, ["Hallazgos encontrados"]),
      "Deuda 2025": getValue(row as unknown as Record<string, string>, ["Deuda 2025"]),
      "Última lectura": getValue(row as unknown as Record<string, string>, ["Última lectura", "Ultima lectura", "?ltima lectura"]),
      "Situación predio": getValue(row as unknown as Record<string, string>, ["Situación predio", "Situacion predio", "Situaci?n predio"]),
      "Surtido predio": getValue(row as unknown as Record<string, string>, ["Surtido predio"]),
      "Volumen recuperado": getValue(row as unknown as Record<string, string>, ["Volumen recuperado"]),
      "Valor recuperado": getValue(row as unknown as Record<string, string>, ["Valor recuperado"]),
      "Estado deuda": getValue(row as unknown as Record<string, string>, ["Estado deuda"]),
      "Liquidación m3": getValue(row as unknown as Record<string, string>, ["Liquidación m3", "Liquidacion m3", "Liquidaci?n m3"]),
      "Liquidación $": getValue(row as unknown as Record<string, string>, ["Liquidación $", "Liquidacion $", "Liquidaci?n $"]),
      "Acción operativa": getValue(row as unknown as Record<string, string>, ["Acción operativa", "Accion operativa", "Acci?n operativa"]),
      "Acción administrativa": getValue(row as unknown as Record<string, string>, ["Acción administrativa", "Accion administrativa", "Acci?n administrativa"]),
      "Acción penal": getValue(row as unknown as Record<string, string>, ["Acción penal", "Accion penal", "Acci?n penal"]),
      Observaciones: getValue(row as unknown as Record<string, string>, ["Observaciones"]),
      "Acto de suspensión": getValue(row as unknown as Record<string, string>, ["Acto de suspensión", "Acto de suspension", "Acto de suspensi?n"]),
      "Avisos reincidencia": getValue(row as unknown as Record<string, string>, ["Avisos reincidencia"]),
      "Observaciones reincidencia": getValue(row as unknown as Record<string, string>, ["Observaciones reincidencia"]),
    }));
  }, [data?.casos_inicial_recuperacion_records]);
  const casosEspecialesFilterOptions = data?.casos_especiales_filter_options ?? {
    zona: [],
    localidad: [],
    barrio: [],
    hallazgo: [],
    resultado: [],
  };
  const casosInicialRecuperacionFilterOptions = data?.casos_inicial_recuperacion_filter_options ?? {
    localidad: [],
    barrio: [],
    hallazgo: [],
    situacion: [],
    accion_administrativa: [],
    estado_deuda: [],
  };
  const isMedidoresMode = baseMode === "medidores";
  const isCasosEspecialesMode = baseMode === "casos_especiales";
  const isCasosInicialRecuperacionMode = useMemo(() => {
    if (!isCasosEspecialesMode) return false;
    const normalizedSheet = normalizeForSearch(data?.sheet_used ?? "").replace("1.", "1 ").replace(/-/g, " ");
    return normalizedSheet.startsWith("1 visita inicial");
  }, [data?.sheet_used, isCasosEspecialesMode]);
  const sheetOptions = useMemo(
    () => (localSheetOptions.length > 0 ? localSheetOptions : data?.available_sheets ?? []),
    [localSheetOptions, data?.available_sheets]
  );

  const estatusOptions = useMemo(() => {
    return data?.all_estatus_options ?? [];
  }, [data?.all_estatus_options]);

  const estadoOptions = useMemo(() => {
    return data?.all_estado_options ?? [];
  }, [data?.all_estado_options]);

  const filteredAdminRecords = useMemo(() => {
    return adminRecords.filter((row) => {
      const matchesDefaultAdmin = DEFAULT_ADMIN_ESTATUS.some(
        (item) => normalizeForSearch(item) === normalizeForSearch(row.Estatus)
      );
      if (!matchesDefaultAdmin) return false;
      if (
        filterEstatus.length > 0 &&
        !filterEstatus.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estatus))
      ) return false;
      if (
        filterEstado.length > 0 &&
        !filterEstado.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estado))
      ) return false;
      return true;
    });
  }, [adminRecords, filterEstatus, filterEstado]);

  const filteredPenalRecords = useMemo(() => {
    return penalRecords.filter((row) => {
      const matchesDefaultPenal = DEFAULT_PENAL_ESTATUS.some(
        (item) => normalizeForSearch(item) === normalizeForSearch(row.Estatus)
      );
      if (!matchesDefaultPenal) return false;
      if (
        filterEstatus.length > 0 &&
        !filterEstatus.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estatus))
      ) return false;
      if (
        filterEstado.length > 0 &&
        !filterEstado.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estado))
      ) return false;
      return true;
    });
  }, [penalRecords, filterEstatus, filterEstado]);

  const adminBoard = useMemo(
    () =>
      buildBoardFromRecords(
        filteredAdminRecords,
        "Tablero de Responsables Administrativos",
        "Solo muestra pendientes con Estatus que contengan para expediente o para administrativo, incluyendo mixtos. Si el responsable viene vacio, se marca como Pendiente por asignar.",
        "Responsable Administrativo",
        "teal",
        45,
        null
      ),
    [filteredAdminRecords]
  );

  const penalBoard = useMemo(
    () =>
      buildBoardFromRecords(
        filteredPenalRecords,
        "Tablero de Responsables Penales",
        "Solo muestra pendientes con Estatus que contengan para expediente o para administrativo, incluyendo mixtos. Si el responsable viene vacio, se marca como Pendiente por asignar.",
        "Responsable Penal",
        "amber",
        45,
        null
      ),
    [filteredPenalRecords]
  );

  const filteredNoProcedenteRecords = useMemo(() => {
    return noProcedenteRecords.filter((row) => {
      if (
        filterEstado.length > 0 &&
        !filterEstado.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estado))
      ) return false;
      return true;
    });
  }, [noProcedenteRecords, filterEstado]);

  const noProcedenteBoard = useMemo(
    () =>
      buildBoardFromRecords(
        filteredNoProcedenteRecords,
        "Tablero de Pendiente Determinar Procedencia",
        `Solo muestra registros con Estatus Pendiente determinar procedencia. La asignación sale de la columna Liquidación y los días se calculan desde la fecha actual contra Fecha de Vencimiento. Los vencidos se cuentan después de ${noProcedenteThreshold} días.`,
        "Liquidación",
        "rose",
        noProcedenteThreshold,
        noProcedenteThreshold
      ),
    [filteredNoProcedenteRecords, noProcedenteThreshold]
  );

  const filteredChartRecords = useMemo(() => {
    return chartRecords.filter((row) => {
      if (
        filterEstatus.length > 0 &&
        !filterEstatus.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estatus))
      ) return false;
      if (
        filterEstado.length > 0 &&
        !filterEstado.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estado))
      ) return false;
      return true;
    });
  }, [chartRecords, filterEstatus, filterEstado]);

  const buildCountsFromRows = (
    rows: { Estatus: string; Estado: string }[],
    key: "Estatus" | "Estado"
  ): ChartPoint[] => {
    const counts = new Map<string, number>();
    for (const row of rows) {
      const label = String(row[key] ?? "").trim();
      if (!label) continue;
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return Array.from(counts.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "es"));
  };

  const estatusChart = useMemo(() => {
    const allCounts = buildCountsFromRows(filteredChartRecords, "Estatus");
    if (!estatusPanelFilter) return allCounts;
    return allCounts.filter((item) => normalizeForSearch(item.label) === normalizeForSearch(estatusPanelFilter));
  }, [filteredChartRecords, estatusPanelFilter]);
  const estadoChart = useMemo(() => {
    const rows = filteredChartRecords.filter((row) => {
      if (!estatusPanelFilter) return true;
      return normalizeForSearch(row.Estatus) === normalizeForSearch(estatusPanelFilter);
    });
    return buildCountsFromRows(rows, "Estado");
  }, [filteredChartRecords, estatusPanelFilter]);
  const pendientesClaveChart = useMemo(() => {
    const normalizeContains = (text: string, target: string) => normalizeForSearch(text).includes(normalizeForSearch(target));
    const pendingItems = estatusOptions.filter((item) => {
      const normalized = normalizeForSearch(item);
      return (
        normalized.includes("para administrativo") ||
        normalized.includes("para expediente")
      );
    });
    return pendingItems
      .map((item) => ({
        label: item,
        count: filteredChartRecords.filter((row) => normalizeForSearch(row.Estatus) === normalizeForSearch(item)).length
      }))
      .filter((item) => item.count > 0)
      .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "es"));
  }, [estatusOptions, filteredChartRecords]);

  const observationChart = useMemo(() => {
    const rows = filteredChartRecords.filter((row) => {
      if (!estatusPanelFilter) return true;
      return normalizeForSearch(row.Estatus) === normalizeForSearch(estatusPanelFilter);
    });
    const counts = new Map<string, number>();
    for (const row of rows) {
      const label = String(row.Observaciones ?? "").trim();
      if (!label) continue;
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return Array.from(counts.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "es"));
  }, [filteredChartRecords, estatusPanelFilter]);

  const filteredCasosSeguimientos = useMemo(() => {
    return casosEspecialesSeguimientos.filter((row) => {
      if (casosZonaFilter !== "Todos" && normalizeForSearch(row.Zona) !== normalizeForSearch(casosZonaFilter)) return false;
      if (casosLocalidadFilter !== "Todos" && normalizeForSearch(row.Localidad) !== normalizeForSearch(casosLocalidadFilter)) return false;
      if (casosHallazgoFilter !== "Todos" && normalizeForSearch(row["Hallazgo encontrado"]) !== normalizeForSearch(casosHallazgoFilter)) return false;
      if (casosResultadoFilter !== "Todos" && normalizeForSearch(row.Resultado) !== normalizeForSearch(casosResultadoFilter)) return false;
      return true;
    });
  }, [casosEspecialesSeguimientos, casosZonaFilter, casosLocalidadFilter, casosHallazgoFilter, casosResultadoFilter]);

  const filteredCasosRecords = useMemo(() => {
    return casosEspecialesRecords.filter((row) => {
      if (casosZonaFilter !== "Todos" && normalizeForSearch(row.Zona) !== normalizeForSearch(casosZonaFilter)) return false;
      if (casosLocalidadFilter !== "Todos" && normalizeForSearch(row.Localidad) !== normalizeForSearch(casosLocalidadFilter)) return false;
      if (casosHallazgoFilter !== "Todos" && normalizeForSearch(row["Hallazgo encontrado"]) !== normalizeForSearch(casosHallazgoFilter)) return false;
      if (casosResultadoFilter !== "Todos") {
        if (normalizeForSearch(row.Ultimo_resultado) === normalizeForSearch(casosResultadoFilter)) return true;
        const hasAny = filteredCasosSeguimientos.some(
          (item) =>
            normalizeForSearch(item["Cuenta contrato"]) === normalizeForSearch(row["Cuenta contrato"]) &&
            normalizeForSearch(item.Resultado) === normalizeForSearch(casosResultadoFilter)
        );
        if (!hasAny) return false;
      }
      return true;
    });
  }, [casosEspecialesRecords, casosZonaFilter, casosLocalidadFilter, casosHallazgoFilter, casosResultadoFilter, filteredCasosSeguimientos]);

  const buildCaseCounts = (rows: Record<string, string>[], field: string): ChartPoint[] => {
    const counts = new Map<string, number>();
    for (const row of rows) {
      const label = String(row[field] ?? "").trim();
      if (!label) continue;
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return Array.from(counts.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "es"));
  };

  const casosHallazgoChart = useMemo(
    () => buildCaseCounts(filteredCasosRecords as unknown as Record<string, string>[], "Hallazgo encontrado"),
    [filteredCasosRecords]
  );
  const casosLocalidadChart = useMemo(
    () => buildCaseCounts(filteredCasosRecords as unknown as Record<string, string>[], "Localidad"),
    [filteredCasosRecords]
  );
  const casosResultadoChart = useMemo(
    () => buildCaseCounts(filteredCasosSeguimientos as unknown as Record<string, string>[], "Resultado"),
    [filteredCasosSeguimientos]
  );
  const casosObservacionChart = useMemo(
    () => buildCaseCounts(filteredCasosSeguimientos as unknown as Record<string, string>[], "Observación"),
    [filteredCasosSeguimientos]
  );
  const casosBarrioChart = useMemo(
    () => buildCaseCounts(filteredCasosRecords as unknown as Record<string, string>[], "Barrio"),
    [filteredCasosRecords]
  );
  const casosIntervencionChart = useMemo(
    () => buildCaseCounts(filteredCasosRecords as unknown as Record<string, string>[], "Intervención"),
    [filteredCasosRecords]
  );
  const casosSeguimientoDepthChart = useMemo(() => {
    const buckets = new Map<string, number>();
    for (const row of filteredCasosRecords) {
      const total = Number(row.Total_seguimientos || 0);
      const label =
        total <= 0 ? "Sin seguimiento" :
        total === 1 ? "1 seguimiento" :
        total === 2 ? "2 seguimientos" :
        total <= 4 ? "3 a 4 seguimientos" :
        "5 o más";
      buckets.set(label, (buckets.get(label) ?? 0) + 1);
    }
    return Array.from(buckets.entries()).map(([label, count]) => ({ label, count }));
  }, [filteredCasosRecords]);
  const casosEstadoOperativoChart = useMemo(() => {
    const counts = new Map<string, number>();
    for (const row of filteredCasosRecords) {
      const result = normalizeForSearch(row.Ultimo_resultado);
      const obs = normalizeForSearch(row.Ultima_observacion);
      let label = "Sin clasificar";
      if (Number(row.Total_seguimientos || 0) <= 0) label = "Sin seguimiento";
      else if (obs.includes("reprogramar")) label = "Reprogramado";
      else if (obs.includes("seguimiento") || result.includes("seguimiento")) label = "En seguimiento";
      else if (result.includes("servicio normal")) label = "Servicio normal";
      else if (result.includes("otras") || result.includes("otra")) label = "Otras";
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return Array.from(counts.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "es"));
  }, [filteredCasosRecords]);

  const casosSinSeguimientoCount = useMemo(
    () => filteredCasosRecords.filter((row) => Number(row.Total_seguimientos || 0) <= 0).length,
    [filteredCasosRecords]
  );
  const casosReprogramadosCount = useMemo(
    () => filteredCasosRecords.filter((row) => normalizeForSearch(row.Ultima_observacion).includes("reprogramar")).length,
    [filteredCasosRecords]
  );
  const casosReincidentesCount = useMemo(
    () => filteredCasosRecords.filter((row) => Number(row.Total_seguimientos || 0) > 1).length,
    [filteredCasosRecords]
  );
  const casosHallazgosSensiblesCount = useMemo(
    () =>
      filteredCasosRecords.filter((row) => {
        const hallazgo = normalizeForSearch(row["Hallazgo encontrado"]);
        return (
          hallazgo.includes("bypass") ||
          hallazgo.includes("uso no autorizado") ||
          hallazgo.includes("servicio directo") ||
          hallazgo.includes("conexion no autorizada") ||
          hallazgo.includes("clandestina") ||
          hallazgo.includes("fraudulenta")
        );
      }).length,
    [filteredCasosRecords]
  );

  const casosResumenRows = useMemo(
    () =>
      filteredCasosRecords
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          "Intervención": row["Intervención"],
          Zona: row.Zona,
          Localidad: row.Localidad,
          "Hallazgo encontrado": row["Hallazgo encontrado"],
          Seguimientos: row.Total_seguimientos,
          "Último resultado": row.Ultimo_resultado || "Sin resultado",
          "Última observación": row.Ultima_observacion || "Sin observación",
        }))
        .sort((a, b) => Number(b.Seguimientos) - Number(a.Seguimientos)),
    [filteredCasosRecords]
  );
  const casosReincidenciaRows = useMemo(
    () =>
      filteredCasosRecords
        .filter((row) => Number(row.Total_seguimientos || 0) > 1)
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          "Intervención": row["Intervención"],
          Localidad: row.Localidad,
          Barrio: row.Barrio,
          "Hallazgo": row["Hallazgo encontrado"],
          "Seguimientos": row.Total_seguimientos,
          "Último resultado": row.Ultimo_resultado || "Sin resultado",
        }))
        .sort((a, b) => Number(b.Seguimientos) - Number(a.Seguimientos)),
    [filteredCasosRecords]
  );
  const casosCriticosRows = useMemo(
    () =>
      filteredCasosRecords
        .filter((row) => {
          const hallazgo = normalizeForSearch(row["Hallazgo encontrado"]);
          const resultado = normalizeForSearch(row.Ultimo_resultado);
          const observacion = normalizeForSearch(row.Ultima_observacion);
          return (
            hallazgo.includes("bypass") ||
            hallazgo.includes("uso no autorizado") ||
            hallazgo.includes("servicio directo") ||
            hallazgo.includes("conexion no autorizada") ||
            observacion.includes("reprogramar") ||
            observacion.includes("medidor robado") ||
            resultado.includes("posible anomalia")
          );
        })
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          "Intervención": row["Intervención"],
          Localidad: row.Localidad,
          "Hallazgo": row["Hallazgo encontrado"],
          "Último resultado": row.Ultimo_resultado || "Sin resultado",
          "Última observación": row.Ultima_observacion || "Sin observación",
        })),
    [filteredCasosRecords]
  );

  const filteredCasosInicialRecuperacion = useMemo(() => {
    const matchesMulti = (value: string | number | null | undefined, selected: string[]) => {
      if (selected.length === 0) return true;
      const normalizedValue = normalizeForSearch(String(value ?? ""));
      return selected.some((item) => normalizeForSearch(item) === normalizedValue);
    };

    return casosInicialRecuperacionRecords.filter((row) => {
      if (!matchesMulti(row.Localidad, casosInicialLocalidadFilter)) return false;
      if (!matchesMulti(row.Barrio, casosInicialBarrioFilter)) return false;
      if (!matchesMulti(row["Hallazgos encontrados"], casosInicialHallazgoFilter)) return false;
      if (!matchesMulti(row["Situación predio"], casosInicialSituacionFilter)) return false;
      if (!matchesMulti(row["Acción administrativa"], casosInicialAccionAdminFilter)) return false;
      if (!matchesMulti(row["Estado deuda"], casosInicialEstadoDeudaFilter)) return false;
      return true;
    });
  }, [
    casosInicialRecuperacionRecords,
    casosInicialLocalidadFilter,
    casosInicialBarrioFilter,
    casosInicialHallazgoFilter,
    casosInicialSituacionFilter,
    casosInicialAccionAdminFilter,
    casosInicialEstadoDeudaFilter
  ]);

  const casosInicialHallazgoChart = useMemo(
    () => buildCaseCounts(filteredCasosInicialRecuperacion as unknown as Record<string, string>[], "Hallazgos encontrados"),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialLocalidadChart = useMemo(
    () => buildCaseCounts(filteredCasosInicialRecuperacion as unknown as Record<string, string>[], "Localidad"),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialEstadoDeudaChart = useMemo(
    () => buildCaseCounts(filteredCasosInicialRecuperacion as unknown as Record<string, string>[], "Estado deuda"),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialActividadChart = useMemo(
    () => buildCaseCounts(filteredCasosInicialRecuperacion as unknown as Record<string, string>[], "Actividad económica"),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialAccionAdminChart = useMemo(
    () => buildCaseCounts(filteredCasosInicialRecuperacion as unknown as Record<string, string>[], "Acción administrativa"),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialSurtidoChart = useMemo(
    () => buildCaseCounts(filteredCasosInicialRecuperacion as unknown as Record<string, string>[], "Surtido predio"),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialRutaChart = useMemo(() => {
    const summary = new Map<string, number>();
    filteredCasosInicialRecuperacion.forEach((row) => {
      const admin = normalizeForSearch(row["Acción administrativa"]);
      const penal = normalizeForSearch(row["Acción penal"]);
      let label = "Sin ruta definida";
      if (admin.includes("procede") || admin.includes("g&u") || admin.includes("acto")) {
        label = "Ruta administrativa activa";
      } else if (hasMeaningfulValue(row["Acción penal"]) && !isNoApplyValue(row["Acción penal"])) {
        label = "Ruta penal activa";
      } else if (hasMeaningfulValue(row["Acción operativa"]) && !isNoApplyValue(row["Acción operativa"])) {
        label = "Ruta operativa";
      } else if (penal.includes("fiscalia") || penal.includes("audiencia")) {
        label = "Ruta penal activa";
      }
      summary.set(label, (summary.get(label) ?? 0) + 1);
    });
    return Array.from(summary.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count);
  }, [filteredCasosInicialRecuperacion]);
  const casosInicialSituacionChart = useMemo(
    () => buildCaseCounts(filteredCasosInicialRecuperacion as unknown as Record<string, string>[], "Situación predio"),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialAccionOperativaChart = useMemo(
    () => buildCaseCounts(filteredCasosInicialRecuperacion as unknown as Record<string, string>[], "Acción operativa"),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialVolumenRecuperado = useMemo(
    () => filteredCasosInicialRecuperacion.reduce((acc, row) => acc + (parseLocaleNumber(row["Volumen recuperado"]) ?? 0), 0),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialValorRecuperado = useMemo(
    () => filteredCasosInicialRecuperacion.reduce((acc, row) => acc + (parseLocaleNumber(row["Valor recuperado"]) ?? 0), 0),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialLiquidacionPesos = useMemo(
    () => filteredCasosInicialRecuperacion.reduce((acc, row) => acc + (parseLocaleNumber(row["Liquidación $"]) ?? 0), 0),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialLiquidacionM3 = useMemo(
    () => filteredCasosInicialRecuperacion.reduce((acc, row) => acc + (parseLocaleNumber(row["Liquidación m3"]) ?? 0), 0),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialRutaPenalActivaCount = useMemo(
    () =>
      filteredCasosInicialRecuperacion.filter((row) => {
        const penal = normalizeForSearch(row["Acción penal"]);
        return hasMeaningfulValue(row["Acción penal"]) && !isNoApplyValue(penal);
      }).length,
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialActoSuspensionCount = useMemo(
    () =>
      filteredCasosInicialRecuperacion.filter((row) => {
        const acto = normalizeForSearch(row["Acto de suspensión"]);
        return hasMeaningfulValue(row["Acto de suspensión"]) && !acto.startsWith("no");
      }).length,
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialConRecuperacionCount = useMemo(
    () =>
      filteredCasosInicialRecuperacion.filter(
        (row) =>
          (parseLocaleNumber(row["Volumen recuperado"]) ?? 0) > 0 ||
          (parseLocaleNumber(row["Valor recuperado"]) ?? 0) > 0 ||
          (parseLocaleNumber(row["Liquidación $"]) ?? 0) > 0
      ).length,
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialRutaJuridicaCount = useMemo(
    () =>
      filteredCasosInicialRecuperacion.filter((row) => {
        const admin = normalizeForSearch(row["Acción administrativa"]);
        const penal = normalizeForSearch(row["Acción penal"]);
        return (
          admin.includes("procede") ||
          admin.includes("g&u") ||
          admin.includes("acto") ||
          penal.includes("fiscalia") ||
          penal.includes("audiencia") ||
          (hasMeaningfulValue(row["Acción penal"]) && !isNoApplyValue(row["Acción penal"]))
        );
      }).length,
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialSinRecuperarCount = useMemo(
    () => filteredCasosInicialRecuperacion.filter((row) => normalizeForSearch(row["Estado deuda"]).includes("sin recuperar")).length,
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialReincidentesCount = useMemo(
    () =>
      filteredCasosInicialRecuperacion.filter(
        (row) => hasMeaningfulValue(row["Avisos reincidencia"]) || hasMeaningfulValue(row["Observaciones reincidencia"])
      ).length,
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialCriticosCount = useMemo(
    () =>
      filteredCasosInicialRecuperacion.filter((row) => {
        const hallazgo = normalizeForSearch(row["Hallazgos encontrados"]);
        const observaciones = normalizeForSearch(row.Observaciones);
        const reincidencia = normalizeForSearch(row["Observaciones reincidencia"]);
        return (
          hallazgo.includes("bypass") ||
          hallazgo.includes("clandestina") ||
          hallazgo.includes("fraudulenta") ||
          hallazgo.includes("uso no autorizado") ||
          observaciones.includes("policia") ||
          observaciones.includes("inspeccion") ||
          reincidencia.includes("reincid") ||
          reincidencia.includes("ampliara") ||
          reincidencia.includes("denuncia")
        );
      }).length,
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialPenalChart = useMemo(() => {
    const labels = new Map<string, number>();
    filteredCasosInicialRecuperacion.forEach((row) => {
      const penal = normalizeForSearch(row["Acción penal"]);
      let label = "Sin gestión penal";
      if (penal.includes("fiscalia") || penal.includes("audiencia") || (hasMeaningfulValue(row["Acción penal"]) && !isNoApplyValue(penal))) {
        label = "Gestión penal activa";
      } else if (isNoApplyValue(penal)) {
        label = "No aplica";
      }
      labels.set(label, (labels.get(label) ?? 0) + 1);
    });
    return Array.from(labels.entries()).map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count);
  }, [filteredCasosInicialRecuperacion]);
  const casosInicialSuspensionChart = useMemo(() => {
    const labels = new Map<string, number>();
    filteredCasosInicialRecuperacion.forEach((row) => {
      const acto = normalizeForSearch(row["Acto de suspensión"]);
      const label =
        hasMeaningfulValue(row["Acto de suspensión"]) && !acto.startsWith("no")
          ? "Con acto de suspensión"
          : "Sin acto de suspensión";
      labels.set(label, (labels.get(label) ?? 0) + 1);
    });
    return Array.from(labels.entries()).map(([label, count]) => ({ label, count }));
  }, [filteredCasosInicialRecuperacion]);
  const casosInicialRecuperacionBucketChart = useMemo(() => {
    const labels = new Map<string, number>();
    filteredCasosInicialRecuperacion.forEach((row) => {
      const value = parseLocaleNumber(row["Liquidación $"]) ?? parseLocaleNumber(row["Valor recuperado"]) ?? 0;
      let label = "Sin valor económico";
      if (value > 250000000) label = "Mayor a 250M";
      else if (value > 100000000) label = "100M a 250M";
      else if (value > 25000000) label = "25M a 100M";
      else if (value > 0) label = "1 a 25M";
      labels.set(label, (labels.get(label) ?? 0) + 1);
    });
    return Array.from(labels.entries()).map(([label, count]) => ({ label, count }));
  }, [filteredCasosInicialRecuperacion]);
  const casosInicialPendientesAccionRows = useMemo(
    () =>
      filteredCasosInicialRecuperacion
        .filter((row) => {
          const accionAdmin = normalizeForSearch(row["Acción administrativa"]);
          const accionPenal = normalizeForSearch(row["Acción penal"]);
          const accionOperativa = normalizeForSearch(row["Acción operativa"]);
          const observaciones = normalizeForSearch(row.Observaciones);
          return (
            !hasMeaningfulValue(row["Acción administrativa"]) ||
            !hasMeaningfulValue(row["Acción penal"]) ||
            !hasMeaningfulValue(row["Acción operativa"]) ||
            isNoApplyValue(accionAdmin) ||
            isNoApplyValue(accionPenal) ||
            isNoApplyValue(accionOperativa) ||
            observaciones.includes("coordinar") ||
            observaciones.includes("solicitar") ||
            observaciones.includes("programar") ||
            observaciones.includes("nuevo operativo")
          );
        })
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          Nombre: row.Nombre,
          Localidad: row.Localidad,
          Barrio: row.Barrio,
          Hallazgo: row["Hallazgos encontrados"],
          "Acción operativa": row["Acción operativa"] || "Pendiente",
          "Acción administrativa": row["Acción administrativa"] || "Pendiente",
          "Acción penal": row["Acción penal"] || "Pendiente",
          Observaciones: row.Observaciones || "-",
        })),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialTopRecuperacionRows = useMemo(
    () =>
      filteredCasosInicialRecuperacion
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          Nombre: row.Nombre,
          Localidad: row.Localidad,
          Hallazgo: row["Hallazgos encontrados"],
          "Volumen recuperado (m3)": parseLocaleNumber(row["Volumen recuperado"]) ?? 0,
          "Valor recuperado (COP)": parseLocaleNumber(row["Valor recuperado"]) ?? 0,
          "Liquidación ($)": parseLocaleNumber(row["Liquidación $"]) ?? 0,
        }))
        .sort((a, b) => Number(b["Valor recuperado (COP)"]) - Number(a["Valor recuperado (COP)"]))
        .slice(0, 12),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialMesaEconomicaRows = useMemo(
    () =>
      filteredCasosInicialRecuperacion
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          Nombre: row.Nombre,
          Localidad: row.Localidad,
          Hallazgo: row["Hallazgos encontrados"],
          "Estado deuda": row["Estado deuda"] || "-",
          "Liquidación ($)": parseLocaleNumber(row["Liquidación $"]) ?? 0,
          "Volumen recuperado": parseLocaleNumber(row["Volumen recuperado"]) ?? 0,
        }))
        .sort((a, b) => Number(b["Liquidación ($)"]) - Number(a["Liquidación ($)"]))
        .slice(0, 12),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialRutaJuridicaRows = useMemo(
    () =>
      filteredCasosInicialRecuperacion
        .filter((row) => {
          const admin = normalizeForSearch(row["Acción administrativa"]);
          const penal = normalizeForSearch(row["Acción penal"]);
          return (
            admin.includes("procede") ||
            admin.includes("g&u") ||
            admin.includes("acto") ||
            penal.includes("fiscalia") ||
            penal.includes("audiencia") ||
            (hasMeaningfulValue(row["Acción penal"]) && !isNoApplyValue(row["Acción penal"]))
          );
        })
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          Nombre: row.Nombre,
          Hallazgo: row["Hallazgos encontrados"],
          "Acción administrativa": row["Acción administrativa"] || "-",
          "Acción penal": row["Acción penal"] || "-",
          Observaciones: row.Observaciones || "-",
        }))
        .slice(0, 12),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialReincidenciaRows = useMemo(
    () =>
      filteredCasosInicialRecuperacion
        .filter((row) => hasMeaningfulValue(row["Avisos reincidencia"]) || hasMeaningfulValue(row["Observaciones reincidencia"]))
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          Nombre: row.Nombre,
          Hallazgo: row["Hallazgos encontrados"],
          "Avisos reincidencia": row["Avisos reincidencia"] || "-",
          "Observaciones reincidencia": row["Observaciones reincidencia"] || "-",
          "Acto de suspensión": row["Acto de suspensión"] || "-",
        }))
        .slice(0, 12),
    [filteredCasosInicialRecuperacion]
  );
  const casosInicialCriticosRows = useMemo(
    () =>
      filteredCasosInicialRecuperacion
        .filter((row) => {
          const hallazgo = normalizeForSearch(row["Hallazgos encontrados"]);
          const observaciones = normalizeForSearch(row.Observaciones);
          const reincidencia = normalizeForSearch(row["Observaciones reincidencia"]);
          return (
            hallazgo.includes("bypass") ||
            hallazgo.includes("clandestina") ||
            hallazgo.includes("fraudulenta") ||
            hallazgo.includes("uso no autorizado") ||
            observaciones.includes("policia") ||
            observaciones.includes("inspeccion") ||
            reincidencia.includes("reincid") ||
            reincidencia.includes("denuncia")
          );
        })
        .map((row) => ({
          "Cuenta contrato": row["Cuenta contrato"],
          Nombre: row.Nombre,
          Localidad: row.Localidad,
          Hallazgo: row["Hallazgos encontrados"],
          "Estado deuda": row["Estado deuda"] || "-",
          Observaciones: row.Observaciones || "-",
        }))
        .slice(0, 12),
    [filteredCasosInicialRecuperacion]
  );

  return (
    <main className="relative mx-auto min-h-screen max-w-7xl px-4 py-10 sm:px-6 lg:px-8">
      <div className="pointer-events-none fixed inset-0 -z-10">
        <div className="app-orbs" />
        <div className={`absolute inset-0 transition-opacity duration-700 ease-[cubic-bezier(0.22,1,0.36,1)] ${darkMode ? "opacity-0" : "opacity-100"}`} style={{ backgroundImage: "radial-gradient(circle at 10% 20%, rgba(66,182,175,0.14), transparent 36%), radial-gradient(circle at 80% 0%, rgba(122,212,202,0.20), transparent 28%), linear-gradient(180deg, #f3fbfa 0%, #f7faf9 48%, #eef7f5 100%)" }} />
        <div className={`absolute inset-0 transition-opacity duration-700 ease-[cubic-bezier(0.22,1,0.36,1)] ${darkMode ? "opacity-100" : "opacity-0"}`} style={{ backgroundImage: "radial-gradient(circle at 10% 20%, rgba(30,41,59,0.45), transparent 36%), radial-gradient(circle at 80% 0%, rgba(15,23,42,0.5), transparent 28%), linear-gradient(180deg, #0b1220 0%, #0f172a 48%, #111827 100%)" }} />
      </div>

      <header className={`sticky top-3 z-30 mb-6 rounded-[26px] border px-5 py-3 backdrop-blur-xl ${darkMode ? "border-slate-700/70 bg-slate-950/60 text-slate-100" : "border-white/70 bg-white/70 text-slate-900"}`}>
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <p className={`text-[11px] font-semibold uppercase tracking-[0.3em] ${darkMode ? "text-cyan-200/70" : "text-cyan-800/70"}`}>William Rodriguez</p>
            <h2 className="mt-1 text-lg font-semibold tracking-tight sm:text-xl">Sistema de alertas y control de pendientes</h2>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <div className={`rounded-full border px-3 py-1.5 text-xs font-medium ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-300" : "border-slate-200 bg-white/80 text-slate-600"}`}>
              Hoja: {data?.sheet_used ?? "Sin lectura"}
            </div>
            <div className={`rounded-full border px-3 py-1.5 text-xs font-medium ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-300" : "border-slate-200 bg-white/80 text-slate-600"}`}>
              Registros: {data?.source_total_rows ?? 0}
            </div>
          </div>
        </div>
      </header>

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

      <section className={`shadow-glow reveal-up relative mb-8 overflow-hidden rounded-[34px] border px-8 py-9 text-white ${darkMode ? "border-slate-700/60 bg-[radial-gradient(circle_at_top_left,rgba(56,189,248,0.16),transparent_24%),radial-gradient(circle_at_80%_20%,rgba(16,185,129,0.10),transparent_22%),linear-gradient(135deg,#0c1424_0%,#111827_44%,#1f2937_100%)]" : "border-cyan-200/50 bg-[radial-gradient(circle_at_top_left,rgba(255,255,255,0.24),transparent_28%),radial-gradient(circle_at_80%_20%,rgba(45,212,191,0.18),transparent_25%),linear-gradient(135deg,#0f766e_0%,#155e75_42%,#1d4ed8_100%)]"}`}>
        <div className="absolute inset-y-0 right-0 w-1/2 bg-[radial-gradient(circle_at_center,rgba(255,255,255,0.16),transparent_62%)]" />
        <div className="absolute -right-10 top-10 h-48 w-48 rounded-full border border-white/10 bg-white/5 blur-2xl" />
        <div className="relative z-10 max-w-4xl">
          <p className={`text-sm font-medium uppercase tracking-[0.32em] ${darkMode ? "text-cyan-200/75" : "text-cyan-50/85"}`}>Centro de Control</p>
          <h1 className="mt-4 max-w-3xl text-4xl font-bold tracking-tight sm:text-5xl">Tablero de control y análisis</h1>
          <p className={`mt-4 max-w-3xl text-sm leading-7 sm:text-base ${darkMode ? "text-slate-300" : "text-cyan-50/90"}`}>
            El aplicativo toma la hoja activa, detecta pendientes administrativos, penales y de procedencia, y los organiza en tableros más limpios para seguimiento operativo. La idea es que la lectura sea inmediata y que el tablero se sienta institucional, no improvisado.
          </p>
        </div>
      </section>

      <section className="card reveal-up reveal-delay-1 relative mb-8 overflow-hidden p-6 sm:p-7">
        <div className={`absolute inset-x-6 top-0 h-px ${darkMode ? "bg-gradient-to-r from-transparent via-cyan-300/60 to-transparent" : "bg-gradient-to-r from-transparent via-cyan-500/60 to-transparent"}`} />
        <form onSubmit={onSubmit} className="grid gap-4 sm:grid-cols-12">
          <div className="sm:col-span-12">
            <label className={`mb-3 block text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Tipo de base</label>
            <div className="grid gap-3 md:grid-cols-3">
              {([
                {
                  key: "administrativa" as const,
                  title: "Base administrativa",
                  description: "Usa la lectura de procesos administrativos, penales y procedencia."
                },
                {
                  key: "medidores" as const,
                  title: "Base de medidores",
                  description: "Prepara la carga para la nueva base de medidores y su analisis independiente."
                },
                {
                  key: "casos_especiales" as const,
                  title: "Casos especiales",
                  description: "Activa la carga para un excel independiente de casos especiales."
                }
              ]).map((option) => {
                const active = baseMode === option.key;
                return (
                  <button
                    key={option.key}
                    type="button"
                    onClick={() => onSelectBaseMode(option.key)}
                    className={`rounded-[24px] border px-5 py-4 text-left transition ${
                      active
                        ? darkMode
                          ? "border-cyan-500/70 bg-cyan-500/10 shadow-[0_0_0_1px_rgba(34,211,238,0.25)]"
                          : "border-cyan-500/60 bg-cyan-50 shadow-[0_0_0_1px_rgba(6,182,212,0.18)]"
                        : darkMode
                          ? "border-slate-700 bg-slate-900/50 hover:border-slate-600 hover:bg-slate-900/75"
                          : "border-slate-200 bg-white/80 hover:border-slate-300 hover:bg-slate-50"
                    }`}
                  >
                    <p className={`text-base font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{option.title}</p>
                    <p className={`mt-1 text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>{option.description}</p>
                  </button>
                );
              })}
            </div>
          </div>

          {baseMode && (
            <>
              <div className="sm:col-span-4">
                <label className={`mb-2 block text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Archivo Excel</label>
                <div className={`group relative overflow-hidden rounded-[22px] border p-2 shadow-sm ${darkMode ? "border-slate-700 bg-slate-900/70" : "border-slate-300 bg-white/90"}`}>
                  <input id="excel-file" type="file" accept=".xlsx,.xlsm,.xltx,.xltm" onChange={onFileChange} className="absolute inset-0 cursor-pointer opacity-0" />
                  <div className="flex min-h-[52px] items-center gap-3">
                    <label htmlFor="excel-file" className={`inline-flex shrink-0 items-center rounded-2xl px-4 py-2 text-sm font-semibold shadow-sm transition ${darkMode ? "bg-gradient-to-r from-cyan-500 to-emerald-500 text-slate-950 group-hover:brightness-110" : "bg-gradient-to-r from-cyan-600 to-emerald-500 text-white group-hover:brightness-105"}`}>
                      Seleccionar archivo
                    </label>
                    <div className="min-w-0">
                      <p className={`truncate text-sm font-medium ${darkMode ? "text-slate-100" : "text-slate-800"}`}>{file?.name ?? "Ningún archivo seleccionado"}</p>
                      <p className={`mt-0.5 text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Formatos permitidos: .xlsx, .xlsm, .xltx, .xltm</p>
                    </div>
                  </div>
                </div>
              </div>
              <div className="sm:col-span-5">
                <label className={`mb-2 block text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Si es archivo de SharePoint, pega el link aqui</label>
                <input type="url" value={sharepointUrl} onChange={(e) => setSharepointUrl(e.target.value)} placeholder="https://tuempresa.sharepoint.com/.../archivo.xlsx" className={`block w-full rounded-xl border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/70 text-slate-100" : "border-slate-300 text-slate-800"}`} />
              </div>
              <div className="sm:col-span-3">
                <label className={`mb-2 block text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>Nombre de hoja</label>
                {sheetOptions.length > 0 ? (
                  <select value={sheetName} onChange={(e) => setSheetName(e.target.value)} className={`block w-full rounded-xl border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/70 text-slate-100" : "border-slate-300 text-slate-800"}`}>
                    {sheetOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                ) : (
                  <input type="text" value={sheetName} onChange={(e) => setSheetName(e.target.value)} className={`block w-full rounded-xl border px-3 py-2 text-sm ${darkMode ? "border-slate-700 bg-slate-900/70 text-slate-100" : "border-slate-300 text-slate-800"}`} />
                )}
              </div>
              <div className="sm:col-span-12 sm:flex sm:items-end sm:justify-end sm:gap-3">
                <button type="button" onClick={onDiagnoseSharepoint} disabled={diagnosingSharepoint} className={`w-full sm:w-64 rounded-xl border px-4 py-2.5 text-sm font-semibold transition disabled:opacity-60 ${darkMode ? "border-slate-600 bg-slate-900/70 text-slate-100 hover:bg-slate-800" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-50"}`}>
                  {diagnosingSharepoint ? "Probando..." : "Probar link SharePoint"}
                </button>
                <button type="submit" disabled={loading} className={`w-full sm:w-64 rounded-xl px-4 py-2.5 text-sm font-semibold text-white transition disabled:opacity-60 ${darkMode ? "bg-brand-600 hover:bg-brand-500" : "bg-slate-900 hover:bg-slate-800"}`}>
                  {loading ? "Procesando..." : "Leer Excel"}
                </button>
              </div>
            </>
          )}
          {loading && (
            <div className="sm:col-span-12">
              <div className={`overflow-hidden rounded-2xl border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50"}`}>
                <div className="flex items-center justify-between gap-3">
                  <div>
                    <p className={`text-sm font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{loadingStage || "Procesando archivo..."}</p>
                    <p className={`mt-1 text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Se esta validando la hoja y construyendo los tableros de pendientes administrativos, penales y procedencia.</p>
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

      {data && !isMedidoresMode && !isCasosEspecialesMode && (
        <section className="mb-8 grid gap-4 md:grid-cols-2 xl:grid-cols-5">
          <div className={`metric-card reveal-up reveal-delay-1 rounded-[26px] border p-5 ${darkMode ? "border-cyan-950/40 bg-slate-900/60" : "border-cyan-100 bg-white/90"}`}>
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Hoja usada</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{data.sheet_used}</p>
          </div>
          <div className={`metric-card reveal-up reveal-delay-2 rounded-[26px] border p-5 ${darkMode ? "border-sky-950/40 bg-slate-900/60" : "border-sky-100 bg-white/90"}`}>
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Filas leidas</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{data.source_total_rows}</p>
          </div>
          <div className={`metric-card reveal-up reveal-delay-3 rounded-[26px] border p-5 ${darkMode ? "border-teal-950/40 bg-slate-900/60" : "border-teal-100 bg-white/90"}`}>
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Pendientes administrativos</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{adminRecords.length}</p>
          </div>
          <div className={`metric-card reveal-up reveal-delay-4 rounded-[26px] border p-5 ${darkMode ? "border-amber-950/40 bg-slate-900/60" : "border-amber-100 bg-white/90"}`}>
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Pendientes penales</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{penalRecords.length}</p>
          </div>
          <div className={`metric-card reveal-up reveal-delay-5 rounded-[26px] border p-5 ${darkMode ? "border-rose-950/40 bg-slate-900/60" : "border-rose-100 bg-white/90"}`}>
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Pendiente procedencia</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{noProcedenteRecords.length}</p>
          </div>
        </section>
      )}

      {data && isMedidoresMode && (
        <DashboardShell
          title="Analizador de medidores"
          eyebrow="Modulo especializado"
          description="Vista compacta para revisar retiros, conceptos, liquidación en m3 y pendientes críticos con una lectura más ejecutiva."
          navItems={["Dashboard", "Retiros", "Conceptos", "Liquidación", "Pendientes"]}
          darkMode={darkMode}
        >
          <section className="grid gap-4 md:grid-cols-2 xl:grid-cols-4 2xl:grid-cols-5">
            <DashboardMetricCard label="Hoja usada" value={data.sheet_used} accent="cyan" detail="Fuente activa" darkMode={darkMode} />
            <DashboardMetricCard label="Filas leídas" value={data.source_total_rows} accent="violet" detail="Registros detectados" darkMode={darkMode} />
            <DashboardMetricCard label="Retiros contabilizados" value={medidoresTotal} accent="emerald" detail="Columna Retirado por" darkMode={darkMode} />
            <DashboardMetricCard label="Pendientes por concepto" value={medidoresPendientes.length} accent="amber" detail="Concepto = Pendiente" darkMode={darkMode} />
            <DashboardMetricCard label="Pendientes por liquidar" value={medidoresPendientesLiquidar} accent="rose" detail="Liquidación en m3 = Pendiente" darkMode={darkMode} />
          </section>

          <section className="grid gap-5 xl:grid-cols-2 2xl:grid-cols-[1.05fr_1.05fr_1.2fr]">
            <div className="dashboard-panel"><MedidoresRetiradoChart data={medidoresRetiradoPor} darkMode={darkMode} /></div>
            <div className="dashboard-panel"><MedidoresConceptoChart data={medidoresConcepto} darkMode={darkMode} /></div>
            <div className="dashboard-panel"><MedidoresLiquidacionCard data={medidoresLiquidacionM3} darkMode={darkMode} /></div>
          </section>

          <div className="dashboard-panel">
            <DataTableCard
              title="Cuentas contrato con concepto Pendiente"
              rows={medidoresPendientes}
              darkMode={darkMode}
              actionLabel="Descargar pendientes"
              onAction={downloadMedidoresPendientes}
              actionLoading={exportingReport}
            />
          </div>
        </DashboardShell>
      )}

      {data && !isMedidoresMode && !isCasosEspecialesMode && (
        <section className="card panel-grid reveal-up reveal-delay-2 relative z-20 mb-8 overflow-visible p-6 sm:p-7">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <p className={`text-[11px] font-semibold uppercase tracking-[0.28em] ${darkMode ? "text-cyan-200/70" : "text-cyan-800/70"}`}>Filtro maestro</p>
              <h2 className={`mt-2 text-xl font-semibold tracking-tight ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Filtros de pendientes</h2>
              <p className={`mt-2 max-w-2xl text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>Solo se muestran los estatus pendientes que contengan para expediente o para administrativo, incluyendo mixtos. Estos filtros afinan esa vista por Estatus y Estado.</p>
            </div>
            <div className="grid w-full gap-3 sm:grid-cols-2 lg:max-w-2xl">
              <MultiSelectFilter
                label="Estatus"
                options={estatusOptions}
                selected={filterEstatus}
                onChange={setFilterEstatus}
                darkMode={darkMode}
              />
              <MultiSelectFilter
                label="Estado"
                options={estadoOptions}
                selected={filterEstado}
                onChange={setFilterEstado}
                darkMode={darkMode}
              />
            </div>
          </div>
          <div className="mt-4 grid gap-3 sm:grid-cols-3">
            <div className={`metric-card rounded-[24px] border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50/90"}`}>
              <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Responsables visibles</p>
              <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{adminBoard.rows.length}</p>
            </div>
            <div className={`metric-card rounded-[24px] border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50/90"}`}>
              <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Alertas visibles</p>
              <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{adminBoard.totals.total_general}</p>
            </div>
            <div className={`metric-card rounded-[24px] border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50/90"}`}>
              <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Vencidas</p>
              <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-rose-200" : "text-rose-700"}`}>{adminBoard.totals.vencidos}</p>
            </div>
          </div>
        </section>
      )}

      {data && !isMedidoresMode && !isCasosEspecialesMode && (
        <section className="mb-8">
          <section className="grid gap-6 xl:grid-cols-2 2xl:grid-cols-4">
          <div className="reveal-up reveal-delay-2">
            <MiniBarChart
              title="Estatus"
              data={estatusChart}
              darkMode={darkMode}
              filterValue={estatusPanelFilter}
              onFilterChange={setEstatusPanelFilter}
              filterOptions={estatusOptions}
              filterLabel="Estatus"
            />
          </div>
          <div className="reveal-up reveal-delay-3"><MiniBarChart title="Estado" data={estadoChart} darkMode={darkMode} /></div>
          <div className="reveal-up reveal-delay-4"><MiniBarChart title="Pendientes Clave" data={pendientesClaveChart} darkMode={darkMode} /></div>
          <div className="reveal-up reveal-delay-5">
            <ObservationChart
              title="Observaciones"
              data={observationChart}
              darkMode={darkMode}
              activeEstatus={estatusPanelFilter}
            />
          </div>
          </section>
        </section>
      )}

      {data && !isMedidoresMode && !isCasosEspecialesMode && (
        <div className="reveal-up reveal-delay-2">
          <BoardTable
            board={adminBoard}
            darkMode={darkMode}
            onDownloadSummary={() => downloadBoardSummary(filteredAdminRecords, "Responsable Administrativo", "resumen_administrativos.xlsx")}
            downloadingSummary={exportingReport}
          />
        </div>
      )}

      {data && !isMedidoresMode && !isCasosEspecialesMode && (
        <div className="reveal-up reveal-delay-3">
          <BoardTable
            board={penalBoard}
            darkMode={darkMode}
            onDownloadSummary={() => downloadBoardSummary(filteredPenalRecords, "Responsable Penal", "resumen_penales.xlsx")}
            downloadingSummary={exportingReport}
          />
        </div>
      )}

      {data && !isMedidoresMode && !isCasosEspecialesMode && (
        <div className="reveal-up reveal-delay-4">
          <BoardTable
            board={noProcedenteBoard}
            darkMode={darkMode}
            onDownloadSummary={() => downloadBoardSummary(filteredNoProcedenteRecords, "Liquidación", "resumen_pendiente_determinar_procedencia.xlsx")}
            downloadingSummary={exportingReport}
            extraControls={
              <div className="flex flex-wrap items-center gap-3">
                <span className={`text-xs font-semibold uppercase tracking-[0.22em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>
                  Comparar vencidos a
                </span>
                <select
                  value={String(noProcedenteThreshold)}
                  onChange={(event) => setNoProcedenteThreshold(Number(event.target.value))}
                  className={`rounded-xl border px-3 py-2 text-sm font-medium ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 bg-white/90 text-slate-800"}`}
                >
                  {[30, 35, 40, 45, 50, 55, 60, 65, 70].map((day) => (
                    <option key={day} value={day}>
                      {day} días
                    </option>
                  ))}
                </select>
              </div>
            }
          />
        </div>
      )}

      {data && isCasosEspecialesMode && (
        isCasosInicialRecuperacionMode ? (
          <section className="space-y-6 reveal-up reveal-delay-2">
            <section className={`shadow-glow relative overflow-hidden rounded-[34px] border px-8 py-8 ${darkMode ? "border-slate-700/60 bg-[radial-gradient(circle_at_top_left,rgba(251,191,36,0.16),transparent_24%),radial-gradient(circle_at_80%_20%,rgba(14,165,233,0.12),transparent_20%),linear-gradient(135deg,#0f172a_0%,#172033_42%,#1f2937_100%)] text-white" : "border-amber-200/60 bg-[radial-gradient(circle_at_top_left,rgba(255,255,255,0.25),transparent_28%),radial-gradient(circle_at_80%_20%,rgba(251,191,36,0.20),transparent_25%),linear-gradient(135deg,#7c2d12_0%,#92400e_32%,#0f766e_100%)] text-white"}`}>
              <div className="pointer-events-none absolute -right-16 top-10 h-52 w-52 rounded-full border border-white/10 bg-white/10 blur-3xl" />
              <div className="pointer-events-none absolute -left-10 bottom-0 h-40 w-40 rounded-full bg-amber-300/10 blur-3xl" />
              <div className="relative z-10 grid gap-6 xl:grid-cols-[minmax(0,1.25fr)_minmax(320px,0.75fr)]">
                <div className="flex flex-col justify-between gap-6">
                  <div>
                    <p className="text-sm font-medium uppercase tracking-[0.32em] text-amber-100/85">Mesa inicial de recuperación</p>
                    <h2 className="mt-4 max-w-3xl text-3xl font-bold tracking-tight sm:text-5xl">Recuperación económica, deuda y decisión jurídica en una sola vista</h2>
                    <p className={`mt-5 max-w-3xl text-sm leading-7 sm:text-base ${darkMode ? "text-slate-300" : "text-amber-50/90"}`}>
                      Esta hoja se trata como una mesa de recuperación. Resume dinero, volumen, criticidad y ruta de salida para separar lo rentable, lo no recuperado y lo que ya exige escalamiento administrativo o penal.
                    </p>
                  </div>
                  <div className="grid gap-3 sm:grid-cols-3">
                    <div className={`rounded-[24px] border px-4 py-4 ${darkMode ? "border-white/10 bg-slate-950/30" : "border-white/20 bg-white/10"}`}>
                      <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-amber-100/70">Casos visibles</p>
                      <p className="mt-2 text-2xl font-bold">{filteredCasosInicialRecuperacion.length}</p>
                    </div>
                    <div className={`rounded-[24px] border px-4 py-4 ${darkMode ? "border-white/10 bg-slate-950/30" : "border-white/20 bg-white/10"}`}>
                      <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-amber-100/70">Sin recuperar</p>
                      <p className="mt-2 text-2xl font-bold">{casosInicialSinRecuperarCount}</p>
                    </div>
                    <div className={`rounded-[24px] border px-4 py-4 ${darkMode ? "border-white/10 bg-slate-950/30" : "border-white/20 bg-white/10"}`}>
                      <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-amber-100/70">Con recuperación</p>
                      <p className="mt-2 text-2xl font-bold">{casosInicialConRecuperacionCount}</p>
                    </div>
                  </div>
                </div>

                <div className="grid gap-3 sm:grid-cols-2">
                  <div className={`rounded-[28px] border p-5 backdrop-blur-xl sm:col-span-2 ${darkMode ? "border-emerald-300/15 bg-slate-950/35" : "border-white/20 bg-white/10"}`}>
                    <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-amber-100/75">Valor económico total</p>
                    <p className="mt-3 text-[clamp(2rem,3vw,3rem)] font-black leading-none">${casosInicialLiquidacionPesos.toLocaleString("es-CO", { maximumFractionDigits: 0 })}</p>
                    <div className="mt-4 flex items-center justify-between gap-4 text-sm text-white/80">
                      <span>Valor recuperado</span>
                      <span className="font-semibold">${casosInicialValorRecuperado.toLocaleString("es-CO", { maximumFractionDigits: 0 })}</span>
                    </div>
                  </div>
                  <div className={`rounded-[24px] border p-4 backdrop-blur-xl ${darkMode ? "border-cyan-300/15 bg-slate-950/35" : "border-white/20 bg-white/10"}`}>
                    <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-amber-100/75">Volumen / liquidación</p>
                    <p className="mt-3 text-3xl font-bold">{casosInicialVolumenRecuperado.toLocaleString("es-CO", { maximumFractionDigits: 2 })} m3</p>
                    <p className="mt-2 text-sm text-white/80">Liquidación en m3: {casosInicialLiquidacionM3.toLocaleString("es-CO", { maximumFractionDigits: 2 })}</p>
                  </div>
                  <div className={`rounded-[24px] border p-4 backdrop-blur-xl ${darkMode ? "border-fuchsia-300/15 bg-slate-950/35" : "border-white/20 bg-white/10"}`}>
                    <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-amber-100/75">Ruta / riesgo</p>
                    <p className="mt-3 text-3xl font-bold">{casosInicialRutaJuridicaCount}</p>
                    <p className="mt-2 text-sm text-white/80">{casosInicialCriticosCount} críticos, {casosInicialRutaPenalActivaCount} penales y {casosInicialActoSuspensionCount} con acto</p>
                  </div>
                </div>
              </div>
            </section>

            <section className="card panel-grid relative z-[80] overflow-visible p-6 sm:p-7">
              <div className={`absolute inset-x-6 top-0 h-px ${darkMode ? "bg-gradient-to-r from-transparent via-amber-300/60 to-transparent" : "bg-gradient-to-r from-transparent via-amber-500/60 to-transparent"}`} />
              <div className="grid gap-6 xl:grid-cols-[minmax(0,0.92fr)_minmax(0,1.08fr)]">
                <div className="space-y-4">
                  <div>
                    <p className={`text-[11px] font-semibold uppercase tracking-[0.28em] ${darkMode ? "text-amber-200/70" : "text-amber-800/70"}`}>Filtro táctico</p>
                    <h3 className={`mt-2 text-xl font-semibold tracking-tight ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Lectura inicial de recuperación</h3>
                    <p className={`mt-2 max-w-2xl text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>
                      Cruza territorio, hallazgo, situación y deuda para ubicar rápido dónde hay recuperación económica, qué casos siguen abiertos y cuáles deben pasar a decisión administrativa o penal.
                    </p>
                  </div>

                  <div className="grid gap-3 sm:grid-cols-3">
                    <div className={`rounded-[20px] border px-4 py-3 ${darkMode ? "border-slate-700/70 bg-slate-950/35" : "border-slate-200 bg-white/85"}`}>
                      <p className={`text-[11px] font-semibold uppercase tracking-[0.24em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Casos visibles</p>
                      <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-50" : "text-slate-900"}`}>{filteredCasosInicialRecuperacion.length}</p>
                    </div>
                    <div className={`rounded-[20px] border px-4 py-3 ${darkMode ? "border-rose-400/20 bg-rose-950/20" : "border-rose-200 bg-rose-50/80"}`}>
                      <p className={`text-[11px] font-semibold uppercase tracking-[0.24em] ${darkMode ? "text-rose-200/75" : "text-rose-700/75"}`}>Sin recuperar</p>
                      <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-rose-100" : "text-rose-900"}`}>{casosInicialSinRecuperarCount}</p>
                    </div>
                    <div className={`rounded-[20px] border px-4 py-3 ${darkMode ? "border-cyan-400/20 bg-cyan-950/20" : "border-cyan-200 bg-cyan-50/80"}`}>
                      <p className={`text-[11px] font-semibold uppercase tracking-[0.24em] ${darkMode ? "text-cyan-200/75" : "text-cyan-700/75"}`}>Ruta crítica</p>
                      <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-cyan-100" : "text-cyan-900"}`}>{casosInicialCriticosCount + casosInicialRutaPenalActivaCount}</p>
                    </div>
                  </div>
                </div>

                <div className={`rounded-[26px] border p-4 sm:p-5 ${darkMode ? "border-slate-700/70 bg-slate-950/35" : "border-slate-200 bg-white/80"}`}>
                  <div className="mb-4 flex items-center justify-between gap-3">
                    <p className={`text-[11px] font-semibold uppercase tracking-[0.24em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Selección múltiple</p>
                    <button
                      type="button"
                      onClick={() => {
                        setCasosInicialLocalidadFilter([]);
                        setCasosInicialBarrioFilter([]);
                        setCasosInicialHallazgoFilter([]);
                        setCasosInicialSituacionFilter([]);
                        setCasosInicialAccionAdminFilter([]);
                        setCasosInicialEstadoDeudaFilter([]);
                      }}
                      className={`rounded-xl border px-3 py-2 text-xs font-semibold transition ${darkMode ? "border-slate-600 bg-slate-900/70 text-slate-100 hover:bg-slate-800" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-50"}`}
                    >
                      Limpiar
                    </button>
                  </div>
                  <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                    <MultiSelectFilter
                      label="Localidad"
                      options={casosInicialRecuperacionFilterOptions.localidad}
                      selected={casosInicialLocalidadFilter}
                      onChange={setCasosInicialLocalidadFilter}
                      darkMode={darkMode}
                    />
                    <MultiSelectFilter
                      label="Barrio"
                      options={casosInicialRecuperacionFilterOptions.barrio}
                      selected={casosInicialBarrioFilter}
                      onChange={setCasosInicialBarrioFilter}
                      darkMode={darkMode}
                    />
                    <MultiSelectFilter
                      label="Hallazgo"
                      options={casosInicialRecuperacionFilterOptions.hallazgo}
                      selected={casosInicialHallazgoFilter}
                      onChange={setCasosInicialHallazgoFilter}
                      darkMode={darkMode}
                    />
                    <MultiSelectFilter
                      label="Situación"
                      options={casosInicialRecuperacionFilterOptions.situacion}
                      selected={casosInicialSituacionFilter}
                      onChange={setCasosInicialSituacionFilter}
                      darkMode={darkMode}
                    />
                    <MultiSelectFilter
                      label="Acción administrativa"
                      options={casosInicialRecuperacionFilterOptions.accion_administrativa}
                      selected={casosInicialAccionAdminFilter}
                      onChange={setCasosInicialAccionAdminFilter}
                      darkMode={darkMode}
                    />
                    <MultiSelectFilter
                      label="Estado de deuda"
                      options={casosInicialRecuperacionFilterOptions.estado_deuda}
                      selected={casosInicialEstadoDeudaFilter}
                      onChange={setCasosInicialEstadoDeudaFilter}
                      darkMode={darkMode}
                    />
                  </div>
                </div>
              </div>
            </section>

            <section className="space-y-5">
              <div className="flex items-end justify-between gap-4">
                <div>
                  <p className={`text-[11px] font-semibold uppercase tracking-[0.26em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Mapa ejecutivo</p>
                  <h3 className={`mt-1 text-lg font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Impacto económico, territorio y salida operativa</h3>
                </div>
                <div className={`hidden rounded-full border px-4 py-2 text-xs font-semibold md:block ${darkMode ? "border-slate-700 bg-slate-950/40 text-slate-300" : "border-slate-200 bg-white/85 text-slate-600"}`}>
                  {casosInicialRutaPenalActivaCount} con ruta penal activa · {casosInicialActoSuspensionCount} con acto
                </div>
              </div>

              <section className="grid gap-5 2xl:grid-cols-[1.18fr_0.92fr]">
                <div className="dashboard-panel">
                  <MiniBarChart title="Hallazgos con mayor impacto" data={casosInicialHallazgoChart} darkMode={darkMode} allowWrapLabels />
                </div>
                <div className="grid gap-5 md:grid-cols-2 2xl:grid-cols-1">
                  <div className="dashboard-panel">
                    <MiniBarChart title="Estado de deuda y recuperación" data={casosInicialEstadoDeudaChart} darkMode={darkMode} allowWrapLabels />
                  </div>
                  <div className="dashboard-panel">
                    <MiniBarChart title="Bandas económicas" data={casosInicialRecuperacionBucketChart} darkMode={darkMode} allowWrapLabels />
                  </div>
                </div>
              </section>

              <section className="grid gap-5 xl:grid-cols-[1.08fr_0.92fr]">
                <div className="dashboard-panel">
                  <MiniBarChart title="Territorio con mayor presión" data={casosInicialLocalidadChart} darkMode={darkMode} allowWrapLabels />
                </div>
                <div className="grid gap-5">
                  <div className="dashboard-panel">
                    <MiniBarChart title="Decisión administrativa" data={casosInicialAccionAdminChart} darkMode={darkMode} allowWrapLabels />
                  </div>
                  <div className="dashboard-panel">
                    <MiniBarChart title="Gestión penal y suspensión" data={casosInicialPenalChart} darkMode={darkMode} allowWrapLabels />
                  </div>
                </div>
              </section>

              <section className="grid gap-5 xl:grid-cols-[1fr_1fr_0.8fr] xl:items-start">
                <div className="dashboard-panel"><MiniBarChart title="Cómo se surte el predio" data={casosInicialSurtidoChart} darkMode={darkMode} allowWrapLabels /></div>
                <div className="dashboard-panel"><MiniBarChart title="Acción operativa ejecutada" data={casosInicialAccionOperativaChart} darkMode={darkMode} allowWrapLabels /></div>
                <div className="dashboard-panel"><MiniBarChart title="Acto de suspensión" data={casosInicialSuspensionChart} darkMode={darkMode} allowWrapLabels /></div>
              </section>
            </section>

            <section className="grid gap-5 xl:grid-cols-[1.1fr_0.9fr]">
              <div className="dashboard-panel">
                <DataTableCard title="Mesa económica: mayor recuperación y liquidación" rows={casosInicialMesaEconomicaRows} darkMode={darkMode} wrapCells variant="executive" />
              </div>
              <div className="dashboard-panel">
                <DataTableCard title="Casos con recuperación estimada más alta" rows={casosInicialTopRecuperacionRows} darkMode={darkMode} wrapCells variant="executive" />
              </div>
            </section>

            <section className="grid gap-5 xl:grid-cols-2">
              <div className="dashboard-panel">
                <DataTableCard title="Ruta jurídica y administrativa prioritaria" rows={casosInicialRutaJuridicaRows} darkMode={darkMode} wrapCells variant="executive" />
              </div>
              <div className="dashboard-panel">
                <DataTableCard title="Pendientes de acción operativa, administrativa o penal" rows={casosInicialPendientesAccionRows} darkMode={darkMode} wrapCells variant="executive" />
              </div>
            </section>

            <section className="grid gap-5 xl:grid-cols-2">
              <div className="dashboard-panel">
                <DataTableCard title="Reincidencias y suspensión sin cierre" rows={casosInicialReincidenciaRows} darkMode={darkMode} wrapCells variant="executive" />
              </div>
              <div className="dashboard-panel">
                <DataTableCard title="Casos críticos para mesa operativa" rows={casosInicialCriticosRows} darkMode={darkMode} wrapCells variant="executive" />
              </div>
            </section>
          </section>
        ) : (
          <DashboardShell
            title="Analizador de casos especiales"
            eyebrow="Centro de seguimiento"
            description="Vista de vigilancia territorial con filtros, distribución de hallazgos, profundidad de seguimiento y mesas de priorización para revisión operativa."
            navItems={["Dashboard", "Hallazgos", "Territorio", "Seguimientos", "Criticos"]}
            darkMode={darkMode}
            variant="cases"
          >
            <section className="grid gap-4 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-5">
              <DashboardMetricCard label="Casos críticos" value={casosCriticosRows.length} accent="rose" detail="Hallazgos u observaciones sensibles" compact darkMode={darkMode} />
              <DashboardMetricCard label="Sin seguimiento" value={casosSinSeguimientoCount} accent="amber" detail="Casos sin gestión registrada" compact darkMode={darkMode} />
              <DashboardMetricCard label="Reincidentes" value={casosReincidentesCount} accent="violet" detail="Con más de un seguimiento" compact darkMode={darkMode} />
              <DashboardMetricCard label="Reprogramados" value={casosReprogramadosCount} accent="fuchsia" detail="Última observación" compact darkMode={darkMode} />
              <DashboardMetricCard label="Hallazgos sensibles" value={casosHallazgosSensiblesCount} accent="cyan" detail="Bypass, no autorizados y similares" compact darkMode={darkMode} />
            </section>

            <section className={`dashboard-filterbar ${darkMode ? "dashboard-filterbar-dark" : "dashboard-filterbar-light"}`}>
              <div>
                <p className="dashboard-brand-eyebrow">Exploración operativa</p>
                <h3 className={`mt-2 text-lg font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Filtros de casos especiales</h3>
                <p className={`mt-2 max-w-3xl text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>
                  Cruza hallazgos, territorio y resultados para identificar focos, reincidencias y casos que siguen abiertos o reprogramados.
                </p>
              </div>
              <div className="grid w-full gap-3 md:grid-cols-2 xl:grid-cols-4">
                <select value={casosZonaFilter} onChange={(e) => setCasosZonaFilter(e.target.value)} className={`dashboard-select ${darkMode ? "dashboard-select-dark" : "dashboard-select-light"}`}>
                  <option>Todos</option>
                  {casosEspecialesFilterOptions.zona.map((option) => <option key={option} value={option}>{option}</option>)}
                </select>
                <select value={casosLocalidadFilter} onChange={(e) => setCasosLocalidadFilter(e.target.value)} className={`dashboard-select ${darkMode ? "dashboard-select-dark" : "dashboard-select-light"}`}>
                  <option>Todos</option>
                  {casosEspecialesFilterOptions.localidad.map((option) => <option key={option} value={option}>{option}</option>)}
                </select>
                <select value={casosHallazgoFilter} onChange={(e) => setCasosHallazgoFilter(e.target.value)} className={`dashboard-select ${darkMode ? "dashboard-select-dark" : "dashboard-select-light"}`}>
                  <option>Todos</option>
                  {casosEspecialesFilterOptions.hallazgo.map((option) => <option key={option} value={option}>{option}</option>)}
                </select>
                <select value={casosResultadoFilter} onChange={(e) => setCasosResultadoFilter(e.target.value)} className={`dashboard-select ${darkMode ? "dashboard-select-dark" : "dashboard-select-light"}`}>
                  <option>Todos</option>
                  {casosEspecialesFilterOptions.resultado.map((option) => <option key={option} value={option}>{option}</option>)}
                </select>
              </div>
            </section>

            <section className="grid gap-5 xl:grid-cols-2">
              <div className="dashboard-panel"><MiniBarChart title="Hallazgos" data={casosHallazgoChart} darkMode={darkMode} allowWrapLabels /></div>
              <div className="dashboard-panel"><MiniBarChart title="Localidades" data={casosLocalidadChart} darkMode={darkMode} allowWrapLabels /></div>
              <div className="dashboard-panel"><MiniBarChart title="Resultados" data={casosResultadoChart} darkMode={darkMode} allowWrapLabels /></div>
              <div className="dashboard-panel"><ObservationChart title="Observaciones" data={casosObservacionChart} darkMode={darkMode} activeEstatus="" allowWrapLabels /></div>
            </section>

            <section className="grid gap-5 xl:grid-cols-2">
              <div className="dashboard-panel"><MiniBarChart title="Barrios" data={casosBarrioChart} darkMode={darkMode} allowWrapLabels /></div>
              <div className="dashboard-panel"><MiniBarChart title="Intervenciones" data={casosIntervencionChart} darkMode={darkMode} allowWrapLabels /></div>
              <div className="dashboard-panel"><MiniBarChart title="Profundidad seguimiento" data={casosSeguimientoDepthChart} darkMode={darkMode} allowWrapLabels /></div>
              <div className="dashboard-panel"><MiniBarChart title="Estado operativo" data={casosEstadoOperativoChart} darkMode={darkMode} allowWrapLabels /></div>
            </section>

            <div className="dashboard-panel">
              <DataTableCard title="Resumen ejecutivo de casos especiales" rows={casosResumenRows} darkMode={darkMode} wrapCells variant="executive" />
            </div>

            <section className="grid gap-5 xl:grid-cols-2">
              <div className="dashboard-panel">
                <DataTableCard title="Reincidencias y seguimientos recurrentes" rows={casosReincidenciaRows} darkMode={darkMode} wrapCells variant="executive" />
              </div>
              <div className="dashboard-panel">
                <DataTableCard title="Casos críticos y observaciones sensibles" rows={casosCriticosRows} darkMode={darkMode} wrapCells variant="executive" />
              </div>
            </section>
          </DashboardShell>
        )
      )}

    </main>
  );
}

function DataTableCard({
  title,
  rows,
  darkMode,
  actionLabel,
  onAction,
  actionLoading = false,
  wrapCells = false,
  variant = "default"
}: {
  title: string;
  rows: Record<string, string | number | null>[];
  darkMode: boolean;
  actionLabel?: string;
  onAction?: () => void;
  actionLoading?: boolean;
  wrapCells?: boolean;
  variant?: "default" | "executive";
}) {
  const COLLAPSED_LIMIT = 12;
  const [expanded, setExpanded] = useState(false);
  useEffect(() => {
    setExpanded(false);
  }, [rows]);

  const formatCellValue = (value: string | number | null | undefined) => {
    const text = String(value ?? "").trim();
    const normalized = normalizeForSearch(text);
    if (!text || normalized === "nan" || normalized === "nat" || normalized === "none" || normalized === "undefined" || normalized === "null") {
      return "-";
    }
    return text;
  };

  const getExecutiveColumnClass = (header: string) => {
    const key = normalizeForSearch(header);
    if (key.includes("cuenta contrato")) return "w-[9rem]";
    if (key.includes("intervencion")) return "w-[16rem]";
    if (key === "zona") return "w-[5rem]";
    if (key.includes("localidad")) return "w-[10rem]";
    if (key.includes("hallazgo")) return "w-[20rem]";
    if (key.includes("seguimientos")) return "w-[7rem]";
    if (key.includes("ultimo resultado")) return "w-[10rem]";
    if (key.includes("ultima observacion")) return "w-[14rem]";
    return "w-auto";
  };

  if (!rows.length) {
    return (
      <section className={`card overflow-hidden p-6 ${variant === "executive" ? "dashboard-table-card" : ""}`}>
        <div className="flex items-center justify-between gap-3">
          <h2 className={`text-lg font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{title}</h2>
          {onAction && actionLabel && (
            <button
              type="button"
              onClick={onAction}
              disabled={actionLoading}
              className={`rounded-xl border px-4 py-2 text-sm font-semibold transition disabled:opacity-60 ${darkMode ? "border-slate-600 bg-slate-900/70 text-slate-100 hover:bg-slate-800" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-50"}`}
            >
              {actionLoading ? "Generando..." : actionLabel}
            </button>
          )}
        </div>
        <p className={`mt-3 text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Sin datos para mostrar.</p>
      </section>
    );
  }

  const headers = Object.keys(rows[0]);
  const visibleRows = expanded ? rows : rows.slice(0, COLLAPSED_LIMIT);
  const canExpand = rows.length > COLLAPSED_LIMIT;
  return (
    <section className={`card overflow-hidden p-6 ${variant === "executive" ? "dashboard-table-card" : ""}`}>
      <div className="flex items-center justify-between gap-3">
        <h2 className={`text-lg font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{title}</h2>
        {onAction && actionLabel && (
          <button
            type="button"
            onClick={onAction}
            disabled={actionLoading}
            className={`rounded-xl border px-4 py-2 text-sm font-semibold transition disabled:opacity-60 ${darkMode ? "border-slate-600 bg-slate-900/70 text-slate-100 hover:bg-slate-800" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-50"}`}
          >
            {actionLoading ? "Generando..." : actionLabel}
          </button>
        )}
      </div>
      <div className="mt-4 overflow-x-auto rounded-2xl border border-slate-200/10">
        <table className={`${variant === "executive" ? "min-w-[1180px] table-fixed text-sm" : "min-w-full text-sm"}`}>
          <thead className={darkMode ? "bg-slate-900/80" : "bg-brand-50"}>
            <tr>
              {headers.map((header) => (
                <th key={header} className={`${variant === "executive" ? getExecutiveColumnClass(header) : ""} ${wrapCells ? "whitespace-normal" : "whitespace-nowrap"} px-3 py-2 text-left font-semibold align-top ${darkMode ? "text-brand-200" : "text-brand-900"}`}>
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {visibleRows.map((row, index) => (
              <tr key={index} className={`${darkMode ? "border-slate-800/90 odd:bg-slate-900/40 even:bg-slate-900/70" : "border-slate-100 odd:bg-white even:bg-slate-50/50"} border-t`}>
                {headers.map((header) => (
                  <td key={header} className={`${variant === "executive" ? getExecutiveColumnClass(header) : ""} ${wrapCells ? "max-w-[22rem] whitespace-normal break-words" : "whitespace-nowrap"} px-3 py-2 align-top ${darkMode ? "text-slate-200" : "text-slate-700"}`}>
                    {formatCellValue(row[header])}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {canExpand && (
        <div className="mt-4 flex justify-end">
          <button
            type="button"
            onClick={() => setExpanded((prev) => !prev)}
            className={`rounded-xl border px-4 py-2 text-sm font-semibold transition ${darkMode ? "border-slate-600 bg-slate-900/70 text-slate-100 hover:bg-slate-800" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-50"}`}
          >
            {expanded ? "Ver menos" : `Ver más (${rows.length - COLLAPSED_LIMIT})`}
          </button>
        </div>
      )}
    </section>
  );
}

function DashboardShell({
  title,
  eyebrow,
  description,
  navItems,
  darkMode,
  children,
  variant = "default"
}: {
  title: string;
  eyebrow: string;
  description: string;
  navItems: string[];
  darkMode: boolean;
  children: ReactNode;
  variant?: "default" | "cases";
}) {
  return (
    <section className={`dashboard-shell ${darkMode ? "dashboard-shell-dark" : "dashboard-shell-light"} ${variant === "cases" ? "dashboard-shell-cases" : ""} reveal-up reveal-delay-2 relative mb-10 overflow-hidden rounded-[34px] border p-4 sm:p-5`}>
      <div className="grid gap-5 xl:grid-cols-[248px_minmax(0,1fr)] 2xl:grid-cols-[264px_minmax(0,1fr)]">
        <aside className={`dashboard-sidebar ${darkMode ? "dashboard-sidebar-dark" : "dashboard-sidebar-light"} hidden xl:flex`}>
          <div className="dashboard-brand">
            <div className="dashboard-brand-mark">A</div>
            <div>
              <p className="dashboard-brand-eyebrow">{eyebrow}</p>
              <h3 className={`dashboard-brand-title ${darkMode ? "text-slate-50" : "text-slate-900"}`}>{title}</h3>
            </div>
          </div>
          <nav className="mt-8 space-y-2">
            {navItems.map((item, index) => (
              <div
                key={item}
                className={`dashboard-nav-item ${darkMode ? "dashboard-nav-item-dark" : "dashboard-nav-item-light"} ${index === 0 ? "dashboard-nav-item-active" : ""}`}
              >
                <span className="dashboard-nav-dot" />
                {item}
              </div>
            ))}
          </nav>
          <div className="dashboard-sidebar-glow" />
        </aside>

        <div className="min-w-0 space-y-5">
          <header className={`dashboard-topbar ${darkMode ? "dashboard-topbar-dark" : "dashboard-topbar-light"}`}>
            <div>
              <p className="dashboard-brand-eyebrow">{eyebrow}</p>
              <h2 className={`mt-2 text-2xl font-semibold tracking-tight ${darkMode ? "text-slate-50" : "text-slate-900"}`}>{title}</h2>
              <p className={`mt-2 max-w-3xl text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>{description}</p>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <div className={`dashboard-top-chip ${darkMode ? "dashboard-top-chip-dark" : "dashboard-top-chip-light"}`}>Vista operativa</div>
              <div className={`dashboard-top-chip ${darkMode ? "dashboard-top-chip-dark" : "dashboard-top-chip-light"}`}>{darkMode ? "Tema oscuro" : "Tema claro"}</div>
            </div>
          </header>
          {children}
        </div>
      </div>
    </section>
  );
}

function DashboardMetricCard({
  label,
  value,
  accent = "cyan",
  detail,
  compact = false,
  darkMode
}: {
  label: string;
  value: string | number;
  accent?: "cyan" | "violet" | "amber" | "emerald" | "rose" | "fuchsia";
  detail?: string;
  compact?: boolean;
  darkMode: boolean;
}) {
  const accentClass = {
    cyan: "dashboard-metric-cyan",
    violet: "dashboard-metric-violet",
    amber: "dashboard-metric-amber",
    emerald: "dashboard-metric-emerald",
    rose: "dashboard-metric-rose",
    fuchsia: "dashboard-metric-fuchsia"
  }[accent];

  return (
    <article className={`dashboard-metric ${darkMode ? "dashboard-metric-dark" : "dashboard-metric-light"} ${compact ? "dashboard-metric-compact" : ""} ${accentClass}`}>
      <p className={`dashboard-metric-label ${darkMode ? "text-slate-400" : "text-slate-500"}`}>{label}</p>
      <p className={`dashboard-metric-value ${darkMode ? "text-slate-50" : "text-slate-900"}`}>{value}</p>
      {detail && <p className={`dashboard-metric-detail ${darkMode ? "text-slate-400" : "text-slate-600"}`}>{detail}</p>}
      <div className="dashboard-metric-wave" />
    </article>
  );
}

function ObservationChart({
  title,
  data,
  darkMode,
  activeEstatus,
  allowWrapLabels = false
}: {
  title: string;
  data: ChartPoint[];
  darkMode: boolean;
  activeEstatus: string;
  allowWrapLabels?: boolean;
}) {
  const COLLAPSED_LIMIT = 12;
  const [expanded, setExpanded] = useState(false);
  useEffect(() => {
    setExpanded(false);
  }, [data, activeEstatus]);
  const maxValue = Math.max(...data.map((item) => item.count), 1);
  const visibleData = expanded ? data : data.slice(0, COLLAPSED_LIMIT);
  const canExpand = data.length > COLLAPSED_LIMIT;
  return (
    <section className="card relative overflow-hidden p-5">
      <div className={`absolute inset-x-6 top-0 h-px ${darkMode ? "bg-gradient-to-r from-transparent via-fuchsia-300/60 to-transparent" : "bg-gradient-to-r from-transparent via-fuchsia-500/60 to-transparent"}`} />
      <div className="mb-4 flex items-end justify-between gap-3">
        <div>
          <h3 className={`text-base font-semibold tracking-tight ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{title}</h3>
          <p className={`mt-1 text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>
            {activeEstatus ? `Observaciones ligadas a ${activeEstatus}.` : "Observaciones ligadas al estatus visible en el panel de estatus."}
          </p>
        </div>
      </div>
      <div className="space-y-2">
        {data.length === 0 && <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Sin observaciones para ese filtro.</p>}
        {visibleData.map((item) => (
          <div key={`${title}-${item.label}`} className="space-y-1.5">
            <div className="flex items-center justify-between text-xs">
              <span className={`${allowWrapLabels ? "line-clamp-2 whitespace-normal break-words" : "truncate"} pr-2 ${darkMode ? "text-slate-300" : "text-slate-700"}`}>{item.label}</span>
              <span className={`font-semibold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{item.count}</span>
            </div>
            <div className={`h-2.5 rounded-full ${darkMode ? "bg-slate-800" : "bg-slate-200"}`}>
              <div
                className={`h-2.5 rounded-full bg-gradient-to-r ${darkMode ? "from-fuchsia-400 via-pink-400 to-rose-500 shadow-[0_12px_30px_-16px_rgba(236,72,153,0.4)]" : "from-fuchsia-500 via-pink-500 to-rose-500 shadow-[0_12px_30px_-16px_rgba(219,39,119,0.28)]"}`}
                style={{ width: `${(item.count / maxValue) * 100}%` }}
              />
            </div>
          </div>
        ))}
        {canExpand && (
          <button
            type="button"
            onClick={() => setExpanded((prev) => !prev)}
            className={`mt-2 inline-flex rounded-full border px-3 py-1.5 text-xs font-semibold transition ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-200 hover:bg-slate-800" : "border-slate-300 bg-white text-slate-700 hover:bg-slate-50"}`}
          >
            {expanded ? "Ver menos" : "Ver más"}
          </button>
        )}
      </div>
    </section>
  );
}

function MedidoresRetiradoChart({
  data,
  darkMode
}: {
  data: { label: string; count: number; percentage: number }[];
  darkMode: boolean;
}) {
  const maxValue = Math.max(...data.map((item) => item.count), 1);
  return (
    <section className="card relative overflow-hidden p-6">
      <div className={`absolute inset-x-6 top-0 h-px ${darkMode ? "bg-gradient-to-r from-transparent via-cyan-300/60 to-transparent" : "bg-gradient-to-r from-transparent via-cyan-500/60 to-transparent"}`} />
      <div className="mb-5">
        <p className={`text-[11px] font-semibold uppercase tracking-[0.28em] ${darkMode ? "text-cyan-200/70" : "text-cyan-800/70"}`}>Analisis de medidores</p>
        <h2 className={`mt-2 text-xl font-semibold tracking-tight ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Retirado por</h2>
        <p className={`mt-2 text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>Conteo total y porcentaje de los retiros segun la columna <span className="font-semibold">Retirado por</span>.</p>
      </div>
      <div className="space-y-3">
        {data.length === 0 && <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>No hay registros validos para mostrar.</p>}
        {data.map((item) => (
          <div key={item.label} className="space-y-1.5">
            <div className="flex items-center justify-between gap-3">
              <p className={`truncate text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>{item.label}</p>
              <div className="flex items-center gap-3 text-right">
                <span className={`text-sm font-semibold tabular-nums ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{item.count}</span>
                <span className={`min-w-[4.5rem] text-xs font-semibold tabular-nums ${darkMode ? "text-cyan-200" : "text-cyan-700"}`}>{item.percentage.toFixed(2)}%</span>
              </div>
            </div>
            <div className={`h-2.5 overflow-hidden rounded-full ${darkMode ? "bg-slate-800" : "bg-slate-200"}`}>
              <div
                className={`h-full rounded-full bg-gradient-to-r ${darkMode ? "from-cyan-400 via-sky-400 to-emerald-400" : "from-cyan-500 via-sky-500 to-emerald-500"}`}
                style={{ width: `${Math.max((item.count / maxValue) * 100, 6)}%` }}
              />
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

function MedidoresConceptoChart({
  data,
  darkMode
}: {
  data: { label: string; count: number; percentage: number }[];
  darkMode: boolean;
}) {
  const maxValue = Math.max(...data.map((item) => item.count), 1);
  return (
    <section className="card relative overflow-hidden p-6">
      <div className={`absolute inset-x-6 top-0 h-px ${darkMode ? "bg-gradient-to-r from-transparent via-amber-300/60 to-transparent" : "bg-gradient-to-r from-transparent via-amber-500/60 to-transparent"}`} />
      <div className="mb-5">
        <p className={`text-[11px] font-semibold uppercase tracking-[0.28em] ${darkMode ? "text-amber-200/70" : "text-amber-800/70"}`}>Estado del medidor</p>
        <h2 className={`mt-2 text-xl font-semibold tracking-tight ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Concepto</h2>
        <p className={`mt-2 text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>Conteo total y porcentaje segun la columna <span className="font-semibold">Concepto</span>.</p>
      </div>
      <div className="space-y-3">
        {data.length === 0 && <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>No hay conceptos validos para mostrar.</p>}
        {data.map((item) => (
          <div key={item.label} className="space-y-1.5">
            <div className="flex items-center justify-between gap-3">
              <p className={`truncate text-sm font-medium ${darkMode ? "text-slate-200" : "text-slate-700"}`}>{item.label}</p>
              <div className="flex items-center gap-3 text-right">
                <span className={`text-sm font-semibold tabular-nums ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{item.count}</span>
                <span className={`min-w-[4.5rem] text-xs font-semibold tabular-nums ${darkMode ? "text-amber-200" : "text-amber-700"}`}>{item.percentage.toFixed(2)}%</span>
              </div>
            </div>
            <div className={`h-2.5 overflow-hidden rounded-full ${darkMode ? "bg-slate-800" : "bg-slate-200"}`}>
              <div
                className={`h-full rounded-full bg-gradient-to-r ${darkMode ? "from-amber-300 via-orange-400 to-yellow-500" : "from-amber-400 via-orange-500 to-yellow-500"}`}
                style={{ width: `${Math.max((item.count / maxValue) * 100, 6)}%` }}
              />
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

function MedidoresLiquidacionCard({
  data,
  darkMode
}: {
  data: { label: string; count: number; sum: number; pending_count: number }[];
  darkMode: boolean;
}) {
  const totalRegistros = data.reduce((acc, item) => acc + Number(item.count || 0), 0);
  const totalM3 = data.reduce((acc, item) => acc + Number(item.sum || 0), 0);

  return (
    <section className="card relative overflow-hidden p-6">
      <div className={`absolute inset-x-6 top-0 h-px ${darkMode ? "bg-gradient-to-r from-transparent via-emerald-300/60 to-transparent" : "bg-gradient-to-r from-transparent via-emerald-500/60 to-transparent"}`} />
      <div className="mb-5">
        <div>
          <p className={`text-[11px] font-semibold uppercase tracking-[0.28em] ${darkMode ? "text-emerald-200/70" : "text-emerald-800/70"}`}>Liquidacion</p>
          <h2 className={`mt-2 text-xl font-semibold tracking-tight ${darkMode ? "text-slate-100" : "text-slate-900"}`}>Liquidación en m3</h2>
          <p className={`mt-2 text-sm leading-6 ${darkMode ? "text-slate-400" : "text-slate-600"}`}>Suma el total de m3 de la columna <span className="font-semibold">Liquidación en m3</span> y lo distribuye por <span className="font-semibold">Concepto</span>.</p>
        </div>
      </div>
      <div className={`mb-5 space-y-4 rounded-[24px] border p-4 ${darkMode ? "border-slate-700 bg-slate-900/60" : "border-slate-200 bg-slate-50/90"}`}>
        <div className={`rounded-[22px] border px-5 py-5 ${darkMode ? "border-emerald-900/50 bg-slate-950/60" : "border-emerald-100 bg-white/90"}`}>
          <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Total m3</p>
          <p className={`mt-3 whitespace-nowrap text-[clamp(1.7rem,2.2vw,2.35rem)] font-bold leading-none tracking-[-0.04em] ${darkMode ? "text-emerald-200" : "text-emerald-800"}`}>
            {totalM3.toLocaleString("es-CO", { minimumFractionDigits: 0, maximumFractionDigits: 2 })}
          </p>
        </div>
        <div className={`rounded-[20px] border px-4 py-4 ${darkMode ? "border-slate-800 bg-slate-950/50" : "border-slate-200 bg-white/85"}`}>
          <p className={`text-xs uppercase tracking-[0.18em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Registros con liquidacion</p>
          <p className={`mt-2 text-3xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{totalRegistros}</p>
        </div>
      </div>
      <div className="space-y-3">
        {data.length === 0 && <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>No hay valores de liquidación para mostrar.</p>}
        {data.map((item) => (
          <div key={item.label} className={`rounded-[18px] border p-4 ${darkMode ? "border-slate-800 bg-slate-900/40" : "border-slate-200 bg-slate-50/70"}`}>
            <div className="flex items-start justify-between gap-3">
              <div className="min-w-0">
                <p className={`truncate text-sm font-semibold ${darkMode ? "text-slate-200" : "text-slate-800"}`}>{item.label}</p>
                <p className={`mt-1 text-xs ${darkMode ? "text-slate-500" : "text-slate-500"}`}>{item.count} registros con valor de liquidación</p>
              </div>
              <span className={`shrink-0 rounded-full px-2.5 py-1 text-xs font-semibold ${darkMode ? "bg-amber-950/60 text-amber-200" : "bg-amber-100 text-amber-800"}`}>
                {item.pending_count} pendientes
              </span>
            </div>
            <div className="mt-4 grid gap-3 sm:grid-cols-2">
              <div className={`rounded-2xl border px-3 py-3 ${darkMode ? "border-slate-800 bg-slate-950/55" : "border-slate-200 bg-white/80"}`}>
                <p className={`text-[11px] uppercase tracking-[0.18em] ${darkMode ? "text-slate-500" : "text-slate-500"}`}>Total m3</p>
                <p className={`mt-2 whitespace-nowrap text-[clamp(1rem,1.2vw,1.25rem)] font-bold leading-tight tracking-[-0.03em] ${darkMode ? "text-emerald-200" : "text-emerald-700"}`}>
                  {Number(item.sum).toLocaleString("es-CO", { minimumFractionDigits: 0, maximumFractionDigits: 2 })}
                </p>
              </div>
              <div className={`rounded-2xl border px-3 py-3 ${darkMode ? "border-slate-800 bg-slate-950/55" : "border-slate-200 bg-white/80"}`}>
                <p className={`text-[11px] uppercase tracking-[0.18em] ${darkMode ? "text-slate-500" : "text-slate-500"}`}>Pendientes</p>
                <p className={`mt-2 text-xl font-bold leading-tight ${darkMode ? "text-amber-200" : "text-amber-700"}`}>{item.pending_count}</p>
              </div>
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

