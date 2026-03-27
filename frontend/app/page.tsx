"use client";

import html2canvas from "html2canvas";
import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";

type PreviewResponse = {
  sheet_used: string;
  available_sheets: string[];
  source_columns: string[];
  source_total_rows: number;
  source_preview: Record<string, string | number | null>[];
  admin_control_records: {
    Responsable: string;
    Fecha_Vencimiento: string;
    DiasInt: number;
    Estatus: string;
    Estado: string;
  }[];
  penal_control_records: {
    Responsable: string;
    Fecha_Vencimiento: string;
    DiasInt: number;
    Estatus: string;
    Estado: string;
  }[];
  procedencia_control_records: {
    Responsable: string;
    Fecha_Vencimiento: string;
    DiasInt: number;
    Estatus: string;
    Estado: string;
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
  day_columns: number[];
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

const API_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";

function normalizeForSearch(value: string | number | null | undefined) {
  return String(value ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
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
  accent: BoardData["accent"] = "teal"
): BoardData {
  if (!records.length) {
    return {
      title,
      description,
      row_label: rowLabel,
      accent,
      day_columns: [],
      rows: [],
      totals: { vencidos: 0, total_general: 0, counts: {} }
    };
  }

  const dayColumns = compactDayColumns(records.map((row) => Number(row.DiasInt)));
  const grouped = new Map<string, { Responsable: string; DiasInt: number }[]>();
  for (const row of records) {
    const key = row.Responsable || "Sin responsable";
    const current = grouped.get(key) ?? [];
    current.push(row);
    grouped.set(key, current);
  }

  const responsibles = Array.from(grouped.keys()).sort((a, b) => {
    if (a === "Sin responsable") return -1;
    if (b === "Sin responsable") return 1;
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
      if (entry.DiasInt < 0) {
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

function BoardTable({ board, darkMode }: { board: BoardData; darkMode: boolean }) {
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
        </div>
        <div className="flex items-center gap-2">
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
              {board.day_columns.map((day) => (
                <th key={day} className={`px-2 py-3 text-center font-semibold ${accentStyles.headText}`}>
                  {day}
                </th>
              ))}
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
                  return (
                    <td key={`${row.responsable}-${day}`} className={`px-2 py-2.5 text-center ${darkMode ? "text-slate-300" : "text-slate-700"} ${day <= 10 && value > 0 ? (darkMode ? "bg-amber-900/35 text-amber-200 font-semibold" : "bg-amber-100/90 text-amber-900 font-semibold") : ""}`}>
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
              {board.day_columns.map((day) => (
                <td key={`tot-${day}`} className={`px-2 py-3 text-center font-bold ${darkMode ? "text-slate-200" : "text-slate-900"}`}>
                  {board.totals.counts[String(day)] || ""}
                </td>
              ))}
              <td className={`px-2 py-3 text-center font-bold ${darkMode ? "text-rose-200" : "text-rose-800"}`}>{board.totals.vencidos || ""}</td>
              <td className={`px-2 py-3 text-center font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{board.totals.total_general}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </section>
  );
}

function MiniBarChart({ title, data, darkMode }: { title: string; data: ChartPoint[]; darkMode: boolean }) {
  const maxValue = Math.max(...data.map((item) => item.count), 1);
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
      <h3 className={`mb-4 text-base font-semibold tracking-tight ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{title}</h3>
      <div className="space-y-2">
        {data.length === 0 && <p className={`text-sm ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Sin datos.</p>}
        {data.map((item) => (
          <div key={`${title}-${item.label}`} className="space-y-1.5">
            <div className="flex items-center justify-between text-xs">
              <span className={`truncate pr-2 ${darkMode ? "text-slate-300" : "text-slate-700"}`}>{item.label}</span>
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
  const [open, setOpen] = useState(false);
  const selectedLabel =
    selected.length === 0 ? "Todos" : selected.length === 1 ? selected[0] : `${selected.length} seleccionados`;

  const toggleOption = (option: string) => {
    const exists = selected.includes(option);
    if (exists) {
      onChange(selected.filter((item) => item !== option));
      return;
    }
    onChange([...selected, option]);
  };

  return (
    <div className={`relative ${open ? "z-[120]" : "z-10"}`}>
      <label className={`mb-1.5 block text-[11px] font-semibold uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>{label}</label>
      <button
        type="button"
        onClick={() => setOpen((prev) => !prev)}
        className={`flex w-full items-center justify-between rounded-2xl border px-4 py-3 text-sm shadow-sm ${darkMode ? "border-slate-700 bg-slate-900/80 text-slate-100" : "border-slate-300 bg-white/90 text-slate-800"}`}
      >
        <span className="truncate">{selectedLabel}</span>
        <span className="ml-3 text-xs">{open ? "▲" : "▼"}</span>
      </button>
      {selected.length > 0 && (
        <div className="mt-3 flex flex-wrap gap-2">
          {selected.map((item) => (
            <button
              key={item}
              type="button"
              onClick={() => toggleOption(item)}
              className={`rounded-full border px-3 py-1 text-xs ${darkMode ? "border-brand-500 bg-brand-500/20 text-brand-100" : "border-brand-300 bg-brand-50 text-brand-800"}`}
            >
              {item} ×
            </button>
          ))}
          <button
            type="button"
            onClick={() => onChange([])}
            className={`rounded-full border px-3 py-1 text-xs ${darkMode ? "border-slate-600 bg-slate-900/80 text-slate-300" : "border-slate-300 bg-white text-slate-700"}`}
          >
            Limpiar
          </button>
        </div>
      )}
      {open && (
        <div className={`absolute left-0 top-full z-[140] mt-2 max-h-72 w-full overflow-auto rounded-[22px] border shadow-2xl ${darkMode ? "border-slate-700 bg-slate-950" : "border-slate-200 bg-white/95"}`}>
          <button
            type="button"
            onClick={() => onChange([])}
            className={`flex w-full items-center justify-between px-3 py-2 text-left text-sm font-semibold ${darkMode ? "border-b border-slate-800 text-slate-100 hover:bg-slate-900" : "border-b border-slate-200 text-slate-800 hover:bg-slate-50"}`}
          >
            <span>Todos</span>
            {selected.length === 0 && <span>✓</span>}
          </button>
          {options.map((option) => {
            const checked = selected.includes(option);
            return (
              <label
                key={option}
                className={`flex cursor-pointer items-center gap-3 px-3 py-2 text-sm ${darkMode ? "text-slate-200 hover:bg-slate-900" : "text-slate-700 hover:bg-slate-50"}`}
              >
                <input type="checkbox" checked={checked} onChange={() => toggleOption(option)} className="h-4 w-4" />
                <span className="truncate">{option}</span>
              </label>
            );
          })}
        </div>
      )}
    </div>
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
  const [filterEstatus, setFilterEstatus] = useState<string[]>([]);
  const [filterEstado, setFilterEstado] = useState<string[]>([]);

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
  }, [data]);

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
        setLoadingStage("Leyendo hoja y construyendo tableros de pendientes...");
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

  const adminRecords = data?.admin_control_records ?? [];
  const penalRecords = data?.penal_control_records ?? [];
  const procedenciaRecords = data?.procedencia_control_records ?? [];

  const estatusOptions = useMemo(() => {
    return Array.from(new Set(adminRecords.map((row) => row.Estatus).filter(Boolean))).sort((a, b) => a.localeCompare(b, "es"));
  }, [adminRecords]);

  const estadoOptions = useMemo(() => {
    return Array.from(new Set(adminRecords.map((row) => row.Estado).filter(Boolean))).sort((a, b) => a.localeCompare(b, "es"));
  }, [adminRecords]);

  const filteredAdminRecords = useMemo(() => {
    return adminRecords.filter((row) => {
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

  const filteredProcedenciaRecords = useMemo(() => {
    return procedenciaRecords.filter((row) => {
      if (
        filterEstado.length > 0 &&
        !filterEstado.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estado))
      ) return false;
      if (
        filterEstatus.length > 0 &&
        !filterEstatus.some((item) => normalizeForSearch(item) === normalizeForSearch(row.Estatus))
      ) return false;
      return true;
    });
  }, [procedenciaRecords, filterEstatus, filterEstado]);

  const adminBoard = useMemo(
    () =>
      buildBoardFromRecords(
        filteredAdminRecords,
        "Tablero de Responsables Administrativos",
        "Solo muestra pendientes con Estatus que contengan para expediente o para administrativo, incluyendo mixtos. Si el responsable viene vacio, se marca como Pendiente por asignar.",
        "Responsable Administrativo",
        "teal"
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
        "amber"
      ),
    [filteredPenalRecords]
  );

  const procedenciaBoard = useMemo(
    () =>
      buildBoardFromRecords(
        filteredProcedenciaRecords,
        "Tablero de Pendiente Determinar Procedencia (45 días)",
        "Solo muestra estatus Pendiente determinar procedencia. La asignación sale de la columna Liquidación y el horizonte es de 45 días.",
        "Liquidación",
        "rose"
      ),
    [filteredProcedenciaRecords]
  );

  const allFilteredRecords = useMemo(
    () => [...filteredAdminRecords, ...filteredPenalRecords, ...filteredProcedenciaRecords],
    [filteredAdminRecords, filteredPenalRecords, filteredProcedenciaRecords]
  );

  const buildTopCounts = (rows: typeof allFilteredRecords, key: "Estatus" | "Estado", topN = 12): ChartPoint[] => {
    const counts = new Map<string, number>();
    for (const row of rows) {
      const label = String(row[key] ?? "").trim();
      if (!label) continue;
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return Array.from(counts.entries())
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, topN);
  };

  const estatusChart = useMemo(() => buildTopCounts(allFilteredRecords, "Estatus", 12), [allFilteredRecords]);
  const estadoChart = useMemo(() => buildTopCounts(allFilteredRecords, "Estado", 12), [allFilteredRecords]);
  const pendientesClaveChart = useMemo(() => {
    const normalizeContains = (text: string, target: string) => normalizeForSearch(text).includes(normalizeForSearch(target));
    const paraAdministrativo = allFilteredRecords.filter((row) => normalizeContains(row.Estatus, "para administrativo")).length;
    const paraExpediente = allFilteredRecords.filter((row) => normalizeContains(row.Estatus, "para expediente")).length;
    const procedencia = filteredProcedenciaRecords.length;
    return [
      { label: "Para administrativo", count: paraAdministrativo },
      { label: "Para expediente", count: paraExpediente },
      { label: "Pendiente procedencia", count: procedencia }
    ].filter((item) => item.count > 0);
  }, [allFilteredRecords, filteredProcedenciaRecords]);

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
          <h1 className="mt-4 max-w-3xl text-4xl font-bold tracking-tight sm:text-5xl">Tableros de pendientes con lectura clara y control visual serio</h1>
          <p className={`mt-4 max-w-3xl text-sm leading-7 sm:text-base ${darkMode ? "text-slate-300" : "text-cyan-50/90"}`}>
            El aplicativo toma la hoja activa, detecta pendientes administrativos, penales y de procedencia, y los organiza en tableros más limpios para seguimiento operativo. La idea es que la lectura sea inmediata y que el tablero se sienta institucional, no improvisado.
          </p>
        </div>
      </section>

      <section className="card reveal-up reveal-delay-1 relative mb-8 overflow-hidden p-6 sm:p-7">
        <div className={`absolute inset-x-6 top-0 h-px ${darkMode ? "bg-gradient-to-r from-transparent via-cyan-300/60 to-transparent" : "bg-gradient-to-r from-transparent via-cyan-500/60 to-transparent"}`} />
        <form onSubmit={onSubmit} className="grid gap-4 sm:grid-cols-12">
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

      {data && (
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
            <p className={`text-xs uppercase tracking-[0.2em] ${darkMode ? "text-slate-400" : "text-slate-500"}`}>Pendientes procedencia</p>
            <p className={`mt-2 text-2xl font-bold ${darkMode ? "text-slate-100" : "text-slate-900"}`}>{procedenciaRecords.length}</p>
          </div>
        </section>
      )}

      {data && (
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

      {data && (
        <section className="mb-8 grid gap-6 lg:grid-cols-3">
          <div className="reveal-up reveal-delay-2"><MiniBarChart title="Estatus (Top)" data={estatusChart} darkMode={darkMode} /></div>
          <div className="reveal-up reveal-delay-3"><MiniBarChart title="Estado (Top)" data={estadoChart} darkMode={darkMode} /></div>
          <div className="reveal-up reveal-delay-4"><MiniBarChart title="Pendientes Clave (Totales)" data={pendientesClaveChart} darkMode={darkMode} /></div>
        </section>
      )}

      {data && <div className="reveal-up reveal-delay-2"><BoardTable board={adminBoard} darkMode={darkMode} /></div>}

      {data && <div className="reveal-up reveal-delay-3"><BoardTable board={penalBoard} darkMode={darkMode} /></div>}

      {data && <div className="reveal-up reveal-delay-4"><BoardTable board={procedenciaBoard} darkMode={darkMode} /></div>}

    </main>
  );
}
