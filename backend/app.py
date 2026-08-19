from __future__ import annotations

import base64
import os
import time
import unicodedata
from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import httpx
import pandas as pd
from fastapi import Body, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse


FIVE_MONTH_DAYS = 150
DEFAULT_SHEET = "Procesos Adminis_Penal"
PROJECT_ROOT = Path(__file__).resolve().parent.parent
GRAPH_TOKEN_CACHE: dict[str, Any] = {"access_token": None, "expires_at": 0}


app = FastAPI(
    title="Alertas Procesos API",
    version="2.0.0",
    description="API para generar tableros de alertas administrativos/penales/procedencia.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def normalize_text(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.split())


def find_column(columns: list[str], target: str) -> str | None:
    target_norm = normalize_text(target)
    for col in columns:
        if normalize_text(col) == target_norm:
            return str(col)
    return None


def col(df: pd.DataFrame, name: str) -> str:
    found = find_column([str(c) for c in df.columns], name)
    if found is None:
        raise ValueError(f"No se encontro la columna requerida: {name}")
    return found


def map_columns(df: pd.DataFrame, required: list[tuple[str, str]]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    missing: list[str] = []
    for original_name, final_name in required:
        found = find_column([str(c) for c in df.columns], original_name)
        if found is None:
            missing.append(original_name)
        else:
            mapping[found] = final_name
    if missing:
        raise ValueError(f"No se encontraron columnas requeridas: {', '.join(missing)}")
    return mapping


def series_empty(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.lower()
    return series.isna() | s.isin({"", "nan", "none", "null", "sin dato", "-"})


def is_unassigned_value(value: Any) -> bool:
    normalized = normalize_text(value)
    return normalized in {
        "",
        "nan",
        "none",
        "null",
        "-",
        "sin dato",
        "n/a",
        "na",
        "n / a",
        "#n/a",
        "#n / a",
        "s/a",
        "s / a",
    }


def split_responsables(value: Any) -> list[str]:
    text = str(value or "").strip()
    if not text or text.lower() in {"nan", "none"}:
        return []
    parts = [text]
    separators = [";", "/", ",", "&", " y "]
    for sep in separators:
        new_parts: list[str] = []
        for part in parts:
            if sep in part:
                new_parts.extend(part.split(sep))
            else:
                new_parts.append(part)
        parts = new_parts
    cleaned = [p.strip() for p in parts if p.strip()]
    return list(dict.fromkeys(cleaned))


def explode_by_responsable(df: pd.DataFrame, responsable_col: str = "Responsable") -> pd.DataFrame:
    if df.empty:
        return df.copy()
    work = df.copy()
    work[responsable_col] = work[responsable_col].apply(split_responsables)
    work = work.explode(responsable_col)
    work[responsable_col] = work[responsable_col].astype(str).str.strip()
    work = work[work[responsable_col] != ""]
    return work


def build_responsable_label(responsable: Any, estado: Any) -> str:
    resp = str(responsable or "").strip()
    estado_norm = normalize_text(estado)
    if not resp or resp.lower() in {"nan", "none"}:
        return "Pendiente por asignar"
    if "para asignacion" in estado_norm or "pendiente por asignar" in estado_norm:
        return "Pendiente por asignar"
    return f"{resp} (Proyeccion)"


def build_block(df: pd.DataFrame, required_columns: list[tuple[str, str]], tipo: str, regla: str) -> pd.DataFrame:
    col_map = map_columns(df, required_columns)
    block = df[list(col_map.keys())].rename(columns=col_map).copy()
    block["Tipo"] = tipo
    block["Regla"] = regla
    return block


def build_alerts_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    administrativo_required = [
        ("Cuenta Contrato", "Cuenta Contrato"),
        ("Interlocutor", "Interlocutor"),
        ("DirecciÃ³n", "DirecciÃ³n"),
        ("Ciudad", "Ciudad"),
        ("Responsable Administrativo", "Responsable"),
        ("Fecha de Vencimiento", "Fecha_Vencimiento"),
        ("DÃAS", "Dias"),
    ]
    penal_required = [
        ("Cuenta Contrato", "Cuenta Contrato"),
        ("Interlocutor", "Interlocutor"),
        ("DirecciÃ³n", "DirecciÃ³n"),
        ("Ciudad", "Ciudad"),
        ("Responsable Penal", "Responsable"),
        ("Fecha de Vencimiento.1", "Fecha_Vencimiento"),
        ("DÃAS.1", "Dias"),
    ]

    admin_df = build_block(
        df, administrativo_required, "Administrativo", f"Vencimiento <= {FIVE_MONTH_DAYS} dias"
    )
    penal_df = build_block(df, penal_required, "Penal", f"Vencimiento <= {FIVE_MONTH_DAYS} dias")
    admin_df["Dias"] = pd.to_numeric(admin_df["Dias"], errors="coerce")
    penal_df["Dias"] = pd.to_numeric(penal_df["Dias"], errors="coerce")

    admin_5m = admin_df[admin_df["Dias"] <= FIVE_MONTH_DAYS].copy()
    penal_5m = penal_df[penal_df["Dias"] <= FIVE_MONTH_DAYS].copy()

    cuenta_col = col(df, "Cuenta Contrato")
    interlocutor_col = col(df, "Interlocutor")
    direccion_col = col(df, "DirecciÃ³n")
    ciudad_col = col(df, "Ciudad")
    fecha_admin_col = col(df, "Fecha de Vencimiento")
    dias_admin_col = col(df, "DÃAS")
    liquidacion_col = col(df, "LiquidaciÃ³n")
    aviso_col = find_column([str(c) for c in df.columns], "Aviso_T2") or find_column(
        [str(c) for c in df.columns], "Aviso"
    )
    analisis_col = col(df, "ANALISIS")
    responsable_adm_col = col(df, "Responsable Administrativo")
    responsable_pen_col = col(df, "Responsable Penal")

    dias_admin = pd.to_numeric(df[dias_admin_col], errors="coerce")
    responsable_base = df[responsable_adm_col].copy()
    mask_empty = series_empty(responsable_base)
    responsable_base.loc[mask_empty] = df[responsable_pen_col].loc[mask_empty]

    # Tablero de "Pendiente determinar procedencia":
    # responsable sale de la columna Liquidación y aplica ventana de 60 días.
    estatus_col = col(df, "Estatus")
    estado_col = col(df, "Estado")
    estatus_norm = df[estatus_col].astype(str).str.strip().str.lower()
    proc_status_mask = estatus_norm.str.contains("pendiente determinar procedencia", na=False)

    proc_df = pd.DataFrame(
        {
            "Aviso": df[aviso_col] if aviso_col else "",
            "Cuenta Contrato": df[cuenta_col],
            "Interlocutor": df[interlocutor_col],
            "DirecciÃ³n": df[direccion_col],
            "Ciudad": df[ciudad_col],
            "Responsable": df[liquidacion_col].astype(str).str.strip(),
            "Quien_Liquida": df[liquidacion_col].astype(str).str.strip(),
            "Fecha_Vencimiento": df[fecha_admin_col],
            "Dias": dias_admin,
            "Tipo": "Pendiente determinar procedencia",
            "Regla": "Estatus pendiente determinar procedencia (<=60 dias)",
            "Estatus": df[estatus_col].astype(str).str.strip(),
            "Estado": df[estado_col].astype(str).str.strip(),
        }
    )
    proc_df = proc_df[proc_status_mask & (proc_df["Dias"] <= 60)].copy()
    proc_df["Responsable"] = proc_df.apply(
        lambda row: build_responsable_label(row.get("Responsable"), row.get("Estado")), axis=1
    )
    proc_df["EmailTrigger"] = proc_df["Dias"].apply(
        lambda d: "PENDIENTE_LIQUIDACION_30" if pd.notna(d) and int(d) == 30 else ""
    )
    pending_status_mask = estatus_norm.str.contains("para administrativo|para expediente", na=False)

    pending_status_df = pd.DataFrame(
        {
            "Aviso": df[aviso_col] if aviso_col else "",
            "Cuenta Contrato": df[cuenta_col],
            "Interlocutor": df[interlocutor_col],
            "DirecciÃ³n": df[direccion_col],
            "Ciudad": df[ciudad_col],
            "Responsable": responsable_base,
            "Quien_Liquida": df[liquidacion_col].astype(str).str.strip(),
            "Fecha_Vencimiento": df[fecha_admin_col],
            "Dias": dias_admin,
            "Tipo": "Pendiente Control",
            "Regla": "Estatus para expediente/para administrativo (incluye mixtos)",
            "Estatus": df[estatus_col].astype(str).str.strip(),
            "Estado": df[estado_col].astype(str).str.strip(),
        }
    )
    pending_status_df = pending_status_df[pending_status_mask & (pending_status_df["Dias"] <= FIVE_MONTH_DAYS)].copy()
    pending_status_df["EmailTrigger"] = pending_status_df["Dias"].apply(
        lambda d: "PENDIENTE_LIQUIDACION_30" if pd.notna(d) and int(d) == 30 else ""
    )

    pending_status_df["Responsable"] = pending_status_df.apply(
        lambda row: build_responsable_label(row.get("Responsable"), row.get("Estado")), axis=1
    )

    admin_5m["Estatus"] = df[estatus_col]
    admin_5m["Estado"] = df[estado_col]
    penal_5m["Estatus"] = df[estatus_col]
    penal_5m["Estado"] = df[estado_col]
    admin_5m["Responsable"] = admin_5m.apply(
        lambda row: build_responsable_label(row.get("Responsable"), row.get("Estado")), axis=1
    )
    penal_5m["Responsable"] = penal_5m.apply(
        lambda row: build_responsable_label(row.get("Responsable"), row.get("Estado")), axis=1
    )

    combinado_5m = pd.concat([admin_5m, penal_5m], ignore_index=True)
    combinado_5m["EmailTrigger"] = combinado_5m["Dias"].apply(
        lambda d: "VENCIMIENTO_10_CC" if pd.notna(d) and int(d) == 10 else ""
    )

    admin_5m["EmailTrigger"] = admin_5m["Dias"].apply(
        lambda d: "VENCIMIENTO_10_CC" if pd.notna(d) and int(d) == 10 else ""
    )
    penal_5m["EmailTrigger"] = penal_5m["Dias"].apply(
        lambda d: "VENCIMIENTO_10_CC" if pd.notna(d) and int(d) == 10 else ""
    )

    all_alerts = pd.concat([proc_df, admin_5m, penal_5m], ignore_index=True)
    all_alerts = all_alerts.sort_values(by=["Dias", "Tipo"], ascending=[True, True], na_position="last")
    return all_alerts, proc_df, pending_status_df


def compact_day_columns(series: pd.Series, min_day: int = 0, max_day: int = FIVE_MONTH_DAYS) -> list[int]:
    values = sorted({int(v) for v in series.dropna().tolist() if min_day <= int(v) <= max_day})
    if len(values) <= 32:
        return values
    head = values[:24]
    milestones = [30, 45, 60, 90, 120, 150]
    tail = [m for m in milestones if m in values]
    merged = sorted(set(head + tail))
    return merged


def build_day_board(alertas: pd.DataFrame, key: str, title: str, description: str) -> dict[str, Any]:
    base = explode_by_responsable(alertas, "Responsable")
    if base.empty:
        return {
            "key": key,
            "title": title,
            "description": description,
            "day_columns": [],
            "rows": [],
            "totals": {"vencidos": 0, "total_general": 0, "counts": {}},
        }

    base["Dias"] = pd.to_numeric(base["Dias"], errors="coerce")
    base = base[base["Dias"].notna()].copy()
    base["DiasInt"] = base["Dias"].astype(int)

    day_columns = compact_day_columns(base["DiasInt"], min_day=0, max_day=FIVE_MONTH_DAYS)
    rows: list[dict[str, Any]] = []
    total_counts = {str(day): 0 for day in day_columns}
    total_vencidos = 0
    total_general = 0

    board_responsables = sorted(set(base["Responsable"].astype(str).str.strip().tolist()))
    if "Pendiente por asignar" in board_responsables:
        board_responsables = ["Pendiente por asignar"] + [n for n in board_responsables if n != "Pendiente por asignar"]

    for name in board_responsables:
        resp_df = base[base["Responsable"] == name]
        counts = {str(day): 0 for day in day_columns}
        if not resp_df.empty:
            vc = resp_df["DiasInt"].value_counts()
            for day in day_columns:
                c = int(vc.get(day, 0))
                counts[str(day)] = c
                total_counts[str(day)] += c
            vencidos = int((resp_df["DiasInt"] < 0).sum())
            total_resp = int(len(resp_df))
        else:
            vencidos = 0
            total_resp = 0
        total_vencidos += vencidos
        total_general += total_resp
        rows.append(
            {
                "responsable": name,
                "vencidos": vencidos,
                "total_general": total_resp,
                "counts": counts,
            }
        )

    return {
        "key": key,
        "title": title,
        "description": description,
        "day_columns": day_columns,
        "rows": rows,
        "totals": {
            "vencidos": total_vencidos,
            "total_general": total_general,
            "counts": total_counts,
        },
    }


def build_pending_control_records(
    df: pd.DataFrame,
    responsable_column: str,
    fecha_column: str,
    dias_column: str,
) -> list[dict[str, Any]]:
    if df.empty:
        return []

    responsable_col = col(df, responsable_column)
    fecha_col = col(df, fecha_column)
    dias_col = col(df, dias_column)
    estatus_col = col(df, "Estatus")
    estado_col = col(df, "Estado")
    aviso_col = find_column([str(c) for c in df.columns], "Aviso_T2")
    fecha_aviso_col = find_column([str(c) for c in df.columns], "Fecha_Aviso")
    cuenta_col = find_column([str(c) for c in df.columns], "Cuenta Contrato")
    anomalia_col = find_column([str(c) for c in df.columns], "Anomalia_Visita")

    work = pd.DataFrame(
        {
            "Responsable": df[responsable_col],
            "Fecha_Vencimiento": df[fecha_col],
            "Dias": pd.to_numeric(df[dias_col], errors="coerce"),
            "Estatus": df[estatus_col],
            "Estado": df[estado_col],
            "Aviso_T2": df[aviso_col] if aviso_col else "",
            "Fecha_Aviso": df[fecha_aviso_col] if fecha_aviso_col else "",
            "Cuenta Contrato": df[cuenta_col] if cuenta_col else "",
            "Anomalia_Visitada": df[anomalia_col] if anomalia_col else "",
        }
    )
    estatus_norm = work["Estatus"].fillna("").astype(str).apply(normalize_text)
    pending_mask = estatus_norm.str.contains("para expediente", na=False) | estatus_norm.str.contains("para administrativo", na=False)
    work = work[pending_mask].copy()
    work = work[work["Dias"].notna()].copy()
    if work.empty:
        return []

    work["Responsable"] = work["Responsable"].fillna("").astype(str).str.strip()
    work.loc[work["Responsable"].apply(is_unassigned_value), "Responsable"] = "Pendiente por asignar"
    work["Estatus"] = work["Estatus"].fillna("").astype(str).str.strip()
    work["Estado"] = work["Estado"].fillna("").astype(str).str.strip()
    work["DiasInt"] = work["Dias"].astype(int)
    work["Fecha_Vencimiento"] = pd.to_datetime(work["Fecha_Vencimiento"], errors="coerce").dt.strftime("%Y-%m-%d")
    work["Fecha_Vencimiento"] = work["Fecha_Vencimiento"].fillna("")
    work["Fecha_Aviso"] = pd.to_datetime(work["Fecha_Aviso"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    work["Aviso_T2"] = work["Aviso_T2"].fillna("").astype(str).str.strip()
    work["Cuenta Contrato"] = work["Cuenta Contrato"].fillna("").astype(str).str.strip()
    work["Anomalia_Visitada"] = work["Anomalia_Visitada"].fillna("").astype(str).str.strip()
    work = explode_by_responsable(work, "Responsable")
    work = work.sort_values(by=["Responsable", "DiasInt", "Fecha_Vencimiento"], ascending=[True, True, True])

    return work[["Responsable", "Fecha_Vencimiento", "DiasInt", "Estatus", "Estado", "Aviso_T2", "Fecha_Aviso", "Cuenta Contrato", "Anomalia_Visitada"]].to_dict(orient="records")


def build_admin_control_records(df: pd.DataFrame) -> list[dict[str, Any]]:
    return build_pending_control_records(
        df,
        responsable_column="Responsable Administrativo",
        fecha_column="Fecha de Vencimiento",
        dias_column="D\u00cdAS",
    )


def build_penal_control_records(df: pd.DataFrame) -> list[dict[str, Any]]:
    return build_pending_control_records(
        df,
        responsable_column="Responsable Penal",
        fecha_column="Fecha de Vencimiento.1",
        dias_column="D\u00cdAS.1",
    )


def build_no_procedente_control_records(df: pd.DataFrame) -> list[dict[str, Any]]:
    if df.empty:
        return []

    responsable_col = (
        find_column([str(c) for c in df.columns], "Liquidación")
        or find_column([str(c) for c in df.columns], "Liquidacion")
        or find_column([str(c) for c in df.columns], "LiquidaciÃ³n")
        or find_column([str(c) for c in df.columns], "LiquidaciÃƒÂ³n")
        or find_column([str(c) for c in df.columns], "Liquidaci?n")
    )
    if responsable_col is None:
        raise ValueError("No se encontro la columna requerida: Liquidación")
    fecha_col = (
        find_column([str(c) for c in df.columns], "F. Vencimiento")
        or find_column([str(c) for c in df.columns], "F.Vencimiento")
        or find_column([str(c) for c in df.columns], "Fecha de Vencimiento")
    )
    if fecha_col is None:
        raise ValueError("No se encontro la columna requerida: F. Vencimiento")
    estatus_col = col(df, "Estatus")
    estado_col = col(df, "Estado")
    aviso_col = find_column([str(c) for c in df.columns], "Aviso_T2")
    fecha_aviso_col = find_column([str(c) for c in df.columns], "Fecha_Aviso")
    cuenta_col = find_column([str(c) for c in df.columns], "Cuenta Contrato")
    anomalia_col = find_column([str(c) for c in df.columns], "Anomalia_Visita")

    work = pd.DataFrame(
        {
            "Responsable": df[responsable_col],
            "Fecha_Vencimiento": pd.to_datetime(df[fecha_col], errors="coerce"),
            "Estatus": df[estatus_col],
            "Estado": df[estado_col],
            "Aviso_T2": df[aviso_col] if aviso_col else "",
            "Fecha_Aviso": df[fecha_aviso_col] if fecha_aviso_col else "",
            "Cuenta Contrato": df[cuenta_col] if cuenta_col else "",
            "Anomalia_Visitada": df[anomalia_col] if anomalia_col else "",
        }
    )
    estatus_norm = work["Estatus"].fillna("").astype(str).apply(normalize_text)
    work = work[estatus_norm == "pendiente determinar procedencia"].copy()
    work = work[work["Fecha_Vencimiento"].notna()].copy()
    if work.empty:
        return []

    today = pd.Timestamp(date.today())
    work["DiasInt"] = (work["Fecha_Vencimiento"] - today).dt.days.astype(int)
    work["Responsable"] = work["Responsable"].fillna("").astype(str).str.strip()
    work.loc[work["Responsable"].apply(is_unassigned_value), "Responsable"] = "Pendiente por asignar"
    work["Estatus"] = work["Estatus"].fillna("").astype(str).str.strip()
    work["Estado"] = work["Estado"].fillna("").astype(str).str.strip()
    work["Fecha_Vencimiento"] = work["Fecha_Vencimiento"].dt.strftime("%Y-%m-%d").fillna("")
    work["Fecha_Aviso"] = pd.to_datetime(work["Fecha_Aviso"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    work["Aviso_T2"] = work["Aviso_T2"].fillna("").astype(str).str.strip()
    work["Cuenta Contrato"] = work["Cuenta Contrato"].fillna("").astype(str).str.strip()
    work["Anomalia_Visitada"] = work["Anomalia_Visitada"].fillna("").astype(str).str.strip()
    work = work.sort_values(by=["Responsable", "DiasInt", "Fecha_Vencimiento"], ascending=[True, True, True])

    return work[["Responsable", "Fecha_Vencimiento", "DiasInt", "Estatus", "Estado", "Aviso_T2", "Fecha_Aviso", "Cuenta Contrato", "Anomalia_Visitada"]].to_dict(orient="records")


def build_status_analysis(df: pd.DataFrame) -> dict[str, Any]:
    def top_counts(column_name: str, top_n: int = 10) -> list[dict[str, Any]]:
        c = col(df, column_name)
        s = df[c].astype(str).str.strip()
        s = s[~s.str.lower().isin({"", "nan", "none"})]
        vc = s.value_counts().head(top_n)
        return [{"label": idx, "count": int(val)} for idx, val in vc.items()]

    estatus_col = col(df, "Estatus")
    estatus = df[estatus_col].astype(str).str.strip().str.lower()
    pending_admin = int(estatus.str.contains("para administrativo", na=False).sum())
    pending_expediente = int(estatus.str.contains("para expediente", na=False).sum())

    return {
        "estatus_top": top_counts("Estatus", 12),
        "estado_top": top_counts("Estado", 12),
        "analisis_top": top_counts("ANALISIS", 12),
        "pendientes_status_totals": [
            {"label": "Para administrativo (incluye mixtos)", "count": pending_admin},
            {"label": "Para expediente (incluye mixtos)", "count": pending_expediente},
        ],
    }


def build_control_dashboard(all_alerts: pd.DataFrame) -> dict[str, Any]:
    if all_alerts.empty:
        return {
            "totals": {
                "alertas_total": 0,
                "vencidas": 0,
                "por_vencer_0_10": 0,
                "rango_11_30": 0,
                "rango_31_60": 0,
                "rango_61_150": 0,
            },
            "tipo_counts": [],
            "regla_counts": [],
            "responsable_top": [],
            "ciudad_top": [],
            "dias_distribution": [],
            "trigger_counts": [],
        }

    work = all_alerts.copy()
    work["Dias"] = pd.to_numeric(work["Dias"], errors="coerce")
    work = work[work["Dias"].notna()].copy()
    work["DiasInt"] = work["Dias"].astype(int)

    def vc_to_list(series: pd.Series, top_n: int = 10) -> list[dict[str, Any]]:
        vc = series.value_counts().head(top_n)
        return [{"label": str(idx), "count": int(val)} for idx, val in vc.items()]

    totals = {
        "alertas_total": int(len(work)),
        "vencidas": int((work["DiasInt"] < 0).sum()),
        "por_vencer_0_10": int(((work["DiasInt"] >= 0) & (work["DiasInt"] <= 10)).sum()),
        "rango_11_30": int(((work["DiasInt"] >= 11) & (work["DiasInt"] <= 30)).sum()),
        "rango_31_60": int(((work["DiasInt"] >= 31) & (work["DiasInt"] <= 60)).sum()),
        "rango_61_150": int(((work["DiasInt"] >= 61) & (work["DiasInt"] <= 150)).sum()),
    }

    # Explode responsables to handle multi-assignment.
    resp_df = explode_by_responsable(work, "Responsable")
    responsable_top = vc_to_list(resp_df["Responsable"], top_n=12) if not resp_df.empty else []
    ciudad_top = vc_to_list(work["Ciudad"].astype(str).str.strip(), top_n=10)
    tipo_counts = vc_to_list(work["Tipo"].astype(str).str.strip(), top_n=10)
    regla_counts = vc_to_list(work["Regla"].astype(str).str.strip(), top_n=10)

    day_candidates = sorted({int(x) for x in work["DiasInt"].tolist() if 0 <= int(x) <= 60})
    day_columns = day_candidates[:20]
    dias_distribution = [{"label": str(day), "count": int((work["DiasInt"] == day).sum())} for day in day_columns]
    vencidas_extra = int((work["DiasInt"] < 0).sum())
    if vencidas_extra > 0:
        dias_distribution = [{"label": "Vencidas", "count": vencidas_extra}] + dias_distribution

    trigger_series = work["EmailTrigger"].astype(str).str.strip()
    trigger_series = trigger_series[trigger_series != ""]
    trigger_counts = vc_to_list(trigger_series, top_n=10)

    return {
        "totals": totals,
        "tipo_counts": tipo_counts,
        "regla_counts": regla_counts,
        "responsable_top": responsable_top,
        "ciudad_top": ciudad_top,
        "dias_distribution": dias_distribution,
        "trigger_counts": trigger_counts,
    }


def build_analysis_records(all_alerts: pd.DataFrame) -> list[dict[str, Any]]:
    if all_alerts.empty:
        return []
    work = all_alerts.copy()
    work["Dias"] = pd.to_numeric(work["Dias"], errors="coerce")
    work = work[work["Dias"].notna()].copy()
    work["DiasInt"] = work["Dias"].astype(int)
    work = explode_by_responsable(work, "Responsable")
    if work.empty:
        return []
    for optional_col, default_value in [
        ("Aviso", ""),
        ("Cuenta Contrato", ""),
        ("Estatus", ""),
        ("Quien_Liquida", ""),
        ("Fecha_Vencimiento", None),
    ]:
        if optional_col not in work.columns:
            work[optional_col] = default_value

    out = work[
        [
            "Tipo",
            "Regla",
            "Responsable",
            "Ciudad",
            "DiasInt",
            "EmailTrigger",
            "Aviso",
            "Cuenta Contrato",
            "Estatus",
            "Quien_Liquida",
            "Fecha_Vencimiento",
        ]
    ].copy()
    out["Tipo"] = out["Tipo"].astype(str).str.strip()
    out["Regla"] = out["Regla"].astype(str).str.strip()
    out["Responsable"] = out["Responsable"].astype(str).str.strip()
    out["Ciudad"] = out["Ciudad"].astype(str).str.strip()
    out["EmailTrigger"] = out["EmailTrigger"].astype(str).str.strip()
    out["Aviso"] = out["Aviso"].astype(str).str.strip()
    out["Cuenta Contrato"] = out["Cuenta Contrato"].astype(str).str.strip()
    out["Estatus"] = out["Estatus"].astype(str).str.strip()
    out["Quien_Liquida"] = out["Quien_Liquida"].astype(str).str.strip()
    out["Fecha_Vencimiento"] = pd.to_datetime(out["Fecha_Vencimiento"], errors="coerce").dt.strftime("%Y-%m-%d")
    return out.to_dict(orient="records")


def extract_row_years(row: pd.Series, date_columns: list[str]) -> list[int]:
    years: set[int] = set()
    for col_name in date_columns:
        if col_name not in row.index:
            continue
        raw_value = row.get(col_name)
        if pd.isna(raw_value):
            continue
        parsed = pd.to_datetime(raw_value, errors="coerce")
        if pd.notna(parsed):
            years.add(int(parsed.year))
            continue
        text = str(raw_value).strip()
        parts = text.replace("/", "-").split("-")
        year_part = next((part for part in parts if len(part) == 4 and part.startswith("20")), "")
        if year_part:
            years.add(int(year_part))
    return sorted(years)


def build_status_records(df: pd.DataFrame) -> list[dict[str, Any]]:
    if df.empty:
        return []

    optional_columns = [
        ("Cuenta Contrato", "Cuenta Contrato"),
        ("Ciudad", "Ciudad"),
        ("Estatus", "Estatus"),
        ("Estado", "Estado"),
        ("Responsable Administrativo", "Responsable_Administrativo"),
        ("Responsable Penal", "Responsable_Penal"),
        ("LiquidaciÃ³n", "Liquidacion"),
        ("LiquidaciÃƒÂ³n", "Liquidacion"),
        ("Fecha de Vencimiento", "Fecha_Vencimiento_Admin"),
        ("Fecha de Vencimiento.1", "Fecha_Vencimiento_Penal"),
    ]

    selected: dict[str, str] = {}
    for source_name, target_name in optional_columns:
        found = find_column([str(c) for c in df.columns], source_name)
        if found is not None and target_name not in selected.values():
            selected[found] = target_name

    if not selected:
        return []

    out = df[list(selected.keys())].rename(columns=selected).copy()
    date_columns = [name for name in ["Fecha_Vencimiento_Admin", "Fecha_Vencimiento_Penal"] if name in out.columns]
    records: list[dict[str, Any]] = []

    for _, row in out.iterrows():
        responsable = (
            str(row.get("Responsable_Administrativo") or "").strip()
            or str(row.get("Responsable_Penal") or "").strip()
            or str(row.get("Liquidacion") or "").strip()
            or "Sin responsable"
        )
        years = extract_row_years(row, date_columns) or [None]
        base_record = {
            "Cuenta Contrato": str(row.get("Cuenta Contrato") or "").strip(),
            "Ciudad": str(row.get("Ciudad") or "").strip(),
            "Estatus": str(row.get("Estatus") or "").strip(),
            "Estado": str(row.get("Estado") or "").strip(),
            "Responsable": responsable,
        }
        for year in years:
            record = dict(base_record)
            record["Anio"] = year
            records.append(record)

    return records


def build_general_board_records(all_alerts: pd.DataFrame) -> list[dict[str, Any]]:
    if all_alerts.empty:
        return []
    work = all_alerts.copy()
    work["Dias"] = pd.to_numeric(work["Dias"], errors="coerce")
    work = work[work["Dias"].notna()].copy()
    work["DiasInt"] = work["Dias"].astype(int)
    work = explode_by_responsable(work, "Responsable")
    if work.empty:
        return []

    for optional_col, default_value in [
        ("Tipo", ""),
        ("Responsable", ""),
        ("Estatus", ""),
        ("Estado", ""),
        ("Ciudad", ""),
        ("Cuenta Contrato", ""),
    ]:
        if optional_col not in work.columns:
            work[optional_col] = default_value

    out = work[["Tipo", "Responsable", "Estatus", "Estado", "DiasInt", "Ciudad", "Cuenta Contrato"]].copy()
    out["Tipo"] = out["Tipo"].astype(str).str.strip()
    out["Responsable"] = out["Responsable"].astype(str).str.strip()
    out["Estatus"] = out["Estatus"].astype(str).str.strip()
    out["Estado"] = out["Estado"].astype(str).str.strip()
    out["Ciudad"] = out["Ciudad"].astype(str).str.strip()
    out["Cuenta Contrato"] = out["Cuenta Contrato"].astype(str).str.strip()
    return out.to_dict(orient="records")


def serialize_for_json(df: pd.DataFrame, limit: int = 30) -> list[dict[str, Any]]:
    sample = df.head(limit).copy()
    for col_name in sample.columns:
        if pd.api.types.is_datetime64_any_dtype(sample[col_name]):
            sample[col_name] = sample[col_name].dt.strftime("%Y-%m-%d")
    sample = sample.where(pd.notna(sample), None)
    return sample.to_dict(orient="records")


def extract_filter_options(df: pd.DataFrame, column_name: str) -> list[str]:
    source_col = col(df, column_name)
    values = (
        df[source_col]
        .fillna("")
        .astype(str)
        .str.strip()
    )
    values = values[~values.str.lower().isin({"", "nan", "none", "null", "-"})]
    return sorted(values.drop_duplicates().tolist(), key=lambda item: item.lower())


def build_chart_records(df: pd.DataFrame) -> list[dict[str, str]]:
    estatus_col = col(df, "Estatus")
    estado_col = col(df, "Estado")
    observaciones_col = col(df, "Observaciones")
    work = pd.DataFrame(
        {
            "Estatus": df[estatus_col].fillna("").astype(str).str.strip(),
            "Estado": df[estado_col].fillna("").astype(str).str.strip(),
            "Observaciones": df[observaciones_col].fillna("").astype(str).str.strip(),
        }
    )
    return work.to_dict(orient="records")


def build_medidores_overview(df: pd.DataFrame) -> dict[str, Any]:
    concepto_col = col(df, "Concepto")
    retirado_por_col = (
        find_column([str(c) for c in df.columns], "Retirado por")
        or find_column([str(c) for c in df.columns], "Retirado_por")
        or find_column([str(c) for c in df.columns], "Retiro")
    )
    cuenta_col = col(df, "Cuenta Contrato")
    medidor_col = col(df, "Medidor")
    diametro_col = find_column([str(c) for c in df.columns], "Diametro") or find_column(
        [str(c) for c in df.columns], "DiÃ¡metro"
    )
    lectura_col = find_column([str(c) for c in df.columns], "Lectura")
    retiro_col = find_column([str(c) for c in df.columns], "Retiro")
    liquidacion_m3_col = find_column([str(c) for c in df.columns], "LiquidaciÃ³n en m3") or find_column(
        [str(c) for c in df.columns], "Liquidacion en m3"
    )

    def parse_liquidacion_m3(value: Any) -> float | None:
        text = str(value or "").strip()
        if not text or normalize_text(text) in {"nan", "none", "null", "-", "sin dato"}:
            return None
        text = text.replace("m3", "").replace("M3", "").replace(" ", "")
        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        elif "," in text:
            text = text.replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return None

    work = pd.DataFrame(
        {
            "Concepto": df[concepto_col].fillna("").astype(str).str.strip(),
            "Retirado_por": df[retirado_por_col].fillna("").astype(str).str.strip() if retirado_por_col else "",
            "Cuenta_Contrato": df[cuenta_col].fillna("").astype(str).str.strip(),
            "Medidor": df[medidor_col].fillna("").astype(str).str.strip(),
            "Diametro": df[diametro_col].fillna("").astype(str).str.strip() if diametro_col else "",
            "Lectura": df[lectura_col].fillna("").astype(str).str.strip() if lectura_col else "",
            "Retiro": df[retirado_por_col].fillna("").astype(str).str.strip() if retirado_por_col else "",
            "Retiro_detalle": df[retiro_col].fillna("").astype(str).str.strip() if retiro_col else "",
            "Liquidacion_raw": df[liquidacion_m3_col].fillna("").astype(str).str.strip() if liquidacion_m3_col else "",
            "Liquidacion_m3": df[liquidacion_m3_col].apply(parse_liquidacion_m3) if liquidacion_m3_col else None,
        }
    )

    concept_work = work[
        (work["Concepto"] != "")
        & ~work["Concepto"].str.lower().isin({"nan", "none", "null"})
    ].copy()
    retiro_work = work[
        (work["Retirado_por"] != "")
        & ~work["Retirado_por"].str.lower().isin({"nan", "none", "null", "n/a", "na"})
    ].copy()
    pendientes_work = concept_work[
        (concept_work["Cuenta_Contrato"] != "")
        & ~concept_work["Cuenta_Contrato"].str.lower().isin({"nan", "none", "null"})
    ].copy()

    total = int(len(concept_work))
    if total == 0:
        return {
            "medidores_total": 0,
            "medidores_retirado_por": [],
            "medidores_concepto": [],
            "medidores_pendientes": [],
            "medidores_liquidacion_m3": [],
            "medidores_pendientes_liquidar": 0,
        }

    counts = (
        retiro_work.groupby("Retirado_por", dropna=False)
        .size()
        .sort_values(ascending=False)
        .reset_index(name="count")
    )
    retiro_total = max(int(len(retiro_work)), 1)
    counts["percentage"] = counts["count"].apply(lambda value: round((float(value) / retiro_total) * 100, 2))

    concept_counts = (
        concept_work.groupby("Concepto", dropna=False)
        .size()
        .sort_values(ascending=False)
        .reset_index(name="count")
    )
    concept_counts["percentage"] = concept_counts["count"].apply(lambda value: round((float(value) / total) * 100, 2))

    pendientes = pendientes_work[pendientes_work["Concepto"].apply(normalize_text) == "pendiente"].copy()
    pendientes = pendientes.rename(columns={"Cuenta_Contrato": "Cuenta Contrato"})
    pendientes = pendientes[["Cuenta Contrato", "Medidor", "Diametro", "Lectura", "Retiro", "Concepto"]].sort_values(
        by=["Cuenta Contrato", "Medidor"], ascending=[True, True]
    )

    liquidacion_numeric = (
        concept_work[concept_work["Liquidacion_m3"].notna()]
        .groupby("Concepto", dropna=False)["Liquidacion_m3"]
        .agg(["count", "sum"])
        .reset_index()
    )
    liquidacion_pending = (
        concept_work[concept_work["Liquidacion_raw"].apply(normalize_text) == "pendiente"]
        .groupby("Concepto", dropna=False)
        .size()
        .reset_index(name="pending_count")
    )
    liquidacion_summary = liquidacion_numeric.merge(
        liquidacion_pending,
        on="Concepto",
        how="outer",
    ).fillna({"count": 0, "sum": 0, "pending_count": 0})
    liquidacion_summary["count"] = liquidacion_summary["count"].astype(int)
    liquidacion_summary["pending_count"] = liquidacion_summary["pending_count"].astype(int)
    liquidacion_summary["sum"] = liquidacion_summary["sum"].astype(float).round(2)
    liquidacion_summary = liquidacion_summary.sort_values(
        by=["sum", "pending_count", "Concepto"], ascending=[False, False, True]
    )
    pendientes_liquidar = int((concept_work["Liquidacion_raw"].apply(normalize_text) == "pendiente").sum())

    return {
        "medidores_total": total,
        "medidores_retirado_por": counts.rename(columns={"Retirado_por": "label"}).to_dict(orient="records"),
        "medidores_concepto": concept_counts.rename(columns={"Concepto": "label"}).to_dict(orient="records"),
        "medidores_pendientes": pendientes.to_dict(orient="records"),
        "medidores_liquidacion_m3": liquidacion_summary.rename(columns={"Concepto": "label"}).to_dict(orient="records"),
        "medidores_pendientes_liquidar": pendientes_liquidar,
    }


def parse_decimal_value(value: Any) -> float | None:
    text = str(value or "").strip()
    if not text:
        return None
    normalized = normalize_text(text)
    if normalized in {"nan", "none", "null", "n/a", "na", "-", "sin dato", "pendiente"}:
        return None
    text = (
        text.replace("$", "")
        .replace("cop", "")
        .replace("COP", "")
        .replace("m3", "")
        .replace("M3", "")
        .replace(" ", "")
    )
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def build_casos_inicial_recuperacion_payload(df: pd.DataFrame, target_sheet: str, available_sheets: list[str]) -> dict[str, Any]:
    working_df = df.copy()
    raw_columns = [str(c) for c in working_df.columns]
    normalized_raw = [normalize_text(col) for col in raw_columns]

    if "casos especiales" in normalized_raw:
        header_values = working_df.iloc[0].tolist()
        rebuilt_columns: list[str] = []
        for index, value in enumerate(header_values):
            text = str(value).strip() if value is not None and not pd.isna(value) else ""
            rebuilt_columns.append(text or f"Unnamed: {index}")
        working_df = working_df.iloc[1:].copy()
        working_df.columns = rebuilt_columns

    working_df = working_df.dropna(how="all").reset_index(drop=True)

    leading_fields = []
    for candidate in ["NÂ°", "Actividad econÃ³mica", "Nombre", "Localidad", "Barrio"]:
        found = find_column(columns=[str(c) for c in working_df.columns], target=candidate)
        if found:
            leading_fields.append(found)
    for field in leading_fields:
        working_df[field] = working_df[field].ffill()

    columns = [str(c) for c in working_df.columns]
    column_map = {
        "N": find_column(columns, "NÂ°") or find_column(columns, "No") or columns[0],
        "Actividad economica": find_column(columns, "Actividad econÃ³mica") or find_column(columns, "Actividad economica"),
        "Nombre": find_column(columns, "Nombre"),
        "Direccion": find_column(columns, "Direccion") or find_column(columns, "Direccion"),
        "Localidad": find_column(columns, "Localidad"),
        "Barrio": find_column(columns, "Barrio"),
        "Cuenta contrato": find_column(columns, "Cuenta contrato") or find_column(columns, "Cuenta Contrato"),
        "Aviso": find_column(columns, "Aviso"),
        "Fecha": find_column(columns, "Fecha"),
        "Lectura intervencion m3": find_column(columns, "Lectura dÃ­a de la intervenciÃ³n (m3)") or find_column(columns, "Lectura dia de la intervencion (m3)"),
        "Personal visita": find_column(columns, "Personal de la visita"),
        "Hallazgos encontrados": find_column(columns, "Hallazgos encontrados"),
        "Deuda 2025": find_column(columns, "Deuda 2025") or find_column(columns, "Deuda  2025"),
        "Ultima lectura": find_column(columns, "Ãšltima lectura despuÃ©s del procedimiento") or find_column(columns, "Ultima lectura despues del procedimiento"),
        "Situacion predio": find_column(columns, "SituaciÃ³n actual del predio") or find_column(columns, "Situacion actual del predio"),
        "Surtido predio": find_column(columns, "Â¿CÃ³mo se surte de agua el predio? 24 enero 2025") or find_column(columns, "Como se surte de agua el predio? 24 enero 2025"),
        "Volumen recuperado": find_column(columns, "Volumen recuperado (m3)") or find_column(columns, "Volumen recuperado o (m3)"),
        "Valor recuperado": find_column(columns, "Valor recuperado (COP)"),
        "Estado deuda": find_column(columns, "Estado de la deuda (en mora, pagada, financiada)"),
        "Liquidacion m3": find_column(columns, "LiquidaciÃ³n en m3") or find_column(columns, "Liquidacion en m3"),
        "Liquidacion $": find_column(columns, "LiquidaciÃ³n en $") or find_column(columns, "Liquidacion en $"),
        "Accion operativa": find_column(columns, "AcciÃ³n operativa") or find_column(columns, "Accion operativa"),
        "Accion administrativa": find_column(columns, "AcciÃ³n administrativa") or find_column(columns, "Accion administrativa"),
        "Accion penal": find_column(columns, "AcciÃ³n penal") or find_column(columns, "Accion penal"),
        "Observaciones": find_column(columns, "Observaciones"),
        "Acto de suspension": find_column(columns, "Acto de suspensiÃ³n") or find_column(columns, "Acto de suspension"),
        "Avisos reincidencia": find_column(columns, "Aviso(s) reincidencia") or find_column(columns, "Avisos reincidencia"),
        "Observaciones reincidencia": find_column(columns, "Observaciones reincidencia"),
    }

    def safe_text(source_name: str | None, row: pd.Series) -> str:
        if not source_name:
            return ""
        value = row.get(source_name)
        if isinstance(value, pd.Series):
            for item in value.tolist():
                if pd.isna(item):
                    continue
                text = str(item).strip()
                if text:
                    return text
            return ""
        if pd.isna(value):
            return ""
        return str(value).strip()

    records: list[dict[str, str]] = []
    for _, row in working_df.iterrows():
        item = {key: safe_text(source_col, row) for key, source_col in column_map.items()}
        if not any(item.values()):
            continue
        records.append(item)

    def build_option_values(items: list[dict[str, str]], field: str) -> list[str]:
        return sorted(
            {
                str(item.get(field) or "").strip()
                for item in items
                if str(item.get(field) or "").strip()
                and normalize_text(item.get(field)) not in {"nan", "none", "null", "n/a", "na", "-"}
            },
            key=lambda value: value.lower(),
        )

    return {
        "sheet_used": target_sheet,
        "available_sheets": available_sheets,
        "source_columns": [str(c) for c in working_df.columns],
        "source_total_rows": len(working_df.index),
        "all_estatus_options": [],
        "all_estado_options": [],
        "chart_records": [],
        "report_records": [],
        "source_preview": serialize_for_json(working_df, limit=20),
        "admin_control_records": [],
        "penal_control_records": [],
        "no_procedente_control_records": [],
        "medidores_total": 0,
        "medidores_retirado_por": [],
        "medidores_concepto": [],
        "medidores_pendientes": [],
        "medidores_liquidacion_m3": [],
        "medidores_pendientes_liquidar": 0,
        "casos_especiales_records": [],
        "casos_especiales_seguimientos": [],
        "casos_especiales_filter_options": {
            "zona": [],
            "localidad": [],
            "barrio": [],
            "hallazgo": [],
            "resultado": [],
        },
        "casos_inicial_recuperacion_records": records,
        "casos_inicial_recuperacion_filter_options": {
            "localidad": build_option_values(records, "Localidad"),
            "barrio": build_option_values(records, "Barrio"),
            "hallazgo": build_option_values(records, "Hallazgos encontrados"),
            "situacion": build_option_values(records, "Situacion predio"),
            "accion_administrativa": build_option_values(records, "Accion administrativa"),
            "estado_deuda": build_option_values(records, "Estado deuda"),
        },
    }


def build_administrativa_payload(df: pd.DataFrame, target_sheet: str, available_sheets: list[str]) -> dict[str, Any]:
    return {
        "sheet_used": target_sheet,
        "available_sheets": available_sheets,
        "source_columns": [str(c) for c in df.columns],
        "source_total_rows": len(df.index),
        "all_estatus_options": extract_filter_options(df, "Estatus"),
        "all_estado_options": extract_filter_options(df, "Estado"),
        "chart_records": build_chart_records(df),
        "report_records": build_report_records(df),
        "source_preview": serialize_for_json(df, limit=20),
        "admin_control_records": build_admin_control_records(df),
        "penal_control_records": build_penal_control_records(df),
        "no_procedente_control_records": build_no_procedente_control_records(df),
        "medidores_total": 0,
        "medidores_retirado_por": [],
        "medidores_concepto": [],
        "medidores_pendientes": [],
        "medidores_liquidacion_m3": [],
        "medidores_pendientes_liquidar": 0,
        "casos_especiales_records": [],
        "casos_especiales_seguimientos": [],
        "casos_especiales_filter_options": {
            "zona": [],
            "localidad": [],
            "barrio": [],
            "hallazgo": [],
            "resultado": [],
        },
        "casos_inicial_recuperacion_records": [],
        "casos_inicial_recuperacion_filter_options": {
            "localidad": [],
            "barrio": [],
            "hallazgo": [],
            "situacion": [],
            "accion_administrativa": [],
            "estado_deuda": [],
        },
    }


def build_medidores_payload(df: pd.DataFrame, target_sheet: str, available_sheets: list[str]) -> dict[str, Any]:
    overview = build_medidores_overview(df)
    return {
        "sheet_used": target_sheet,
        "available_sheets": available_sheets,
        "source_columns": [str(c) for c in df.columns],
        "source_total_rows": len(df.index),
        "all_estatus_options": [],
        "all_estado_options": [],
        "chart_records": [],
        "report_records": [],
        "source_preview": serialize_for_json(df, limit=20),
        "admin_control_records": [],
        "penal_control_records": [],
        "no_procedente_control_records": [],
        "medidores_total": overview["medidores_total"],
        "medidores_retirado_por": overview["medidores_retirado_por"],
        "medidores_concepto": overview["medidores_concepto"],
        "medidores_pendientes": overview["medidores_pendientes"],
        "medidores_liquidacion_m3": overview["medidores_liquidacion_m3"],
        "medidores_pendientes_liquidar": overview["medidores_pendientes_liquidar"],
        "casos_especiales_records": [],
        "casos_especiales_seguimientos": [],
        "casos_especiales_filter_options": {
            "zona": [],
            "localidad": [],
            "barrio": [],
            "hallazgo": [],
            "resultado": [],
        },
        "casos_inicial_recuperacion_records": [],
        "casos_inicial_recuperacion_filter_options": {
            "localidad": [],
            "barrio": [],
            "hallazgo": [],
            "situacion": [],
            "accion_administrativa": [],
            "estado_deuda": [],
        },
    }


def build_casos_especiales_payload(df: pd.DataFrame, target_sheet: str, available_sheets: list[str]) -> dict[str, Any]:
    normalized_sheet = normalize_text(target_sheet).replace("1.", "1 ").replace("-", " ")
    if normalized_sheet.startswith("1 visita inicial"):
        return build_casos_inicial_recuperacion_payload(df, target_sheet, available_sheets)

    columns = [str(c) for c in df.columns]

    base_map = {
        "NÂ°": find_column(columns, "NÂ°") or find_column(columns, "No") or find_column(columns, "N"),
        "IntervenciÃ³n": find_column(columns, "IntervenciÃ³n") or find_column(columns, "Intervencion"),
        "Zona": find_column(columns, "Zona"),
        "PorciÃ³n": find_column(columns, "PorciÃ³n") or find_column(columns, "Porcion"),
        "DirecciÃ³n": find_column(columns, "DirecciÃ³n") or find_column(columns, "Direccion"),
        "Localidad": find_column(columns, "Localidad"),
        "Barrio": find_column(columns, "Barrio"),
        "Cuenta contrato": find_column(columns, "Cuenta contrato") or find_column(columns, "Cuenta Contrato"),
        "Hallazgo encontrado": find_column(columns, "Hallazgo encontrado"),
        "Interlocutor": find_column(columns, "Interlocutor"),
        "Equipo": find_column(columns, "Equipo"),
    }

    normalized_columns = [(col_name, normalize_text(col_name)) for col_name in columns]
    seguimiento_cols = [name for name, norm in normalized_columns if norm.startswith("seguimiento no.")]
    if not seguimiento_cols:
        seguimiento_cols = [name for name, norm in normalized_columns if norm.startswith("seguimiento no")]
    resultado_cols = [name for name, norm in normalized_columns if norm == "resultado" or norm.startswith("resultado.")]
    observacion_cols = [name for name, norm in normalized_columns if norm == "observacion" or norm.startswith("observacion.")]

    max_groups = max(len(seguimiento_cols), len(resultado_cols), len(observacion_cols))

    def safe_text(source_name: str | None, row: pd.Series) -> str:
        if not source_name:
            return ""
        return str(row.get(source_name) or "").strip()

    case_records: list[dict[str, Any]] = []
    seguimiento_records: list[dict[str, str]] = []

    for _, row in df.iterrows():
        base_record = {key: safe_text(column_name, row) for key, column_name in base_map.items()}
        if not any(base_record.values()):
            continue

        seguimientos: list[dict[str, str]] = []
        for idx in range(max_groups):
            seguimiento_value = safe_text(seguimiento_cols[idx] if idx < len(seguimiento_cols) else None, row)
            resultado_value = safe_text(resultado_cols[idx] if idx < len(resultado_cols) else None, row)
            observacion_value = safe_text(observacion_cols[idx] if idx < len(observacion_cols) else None, row)
            if not seguimiento_value and not resultado_value and not observacion_value:
                continue
            track = {
                "Seguimiento": seguimiento_value,
                "Resultado": resultado_value,
                "ObservaciÃ³n": observacion_value,
            }
            seguimientos.append(track)
            seguimiento_records.append(
                {
                    "Zona": base_record["Zona"],
                    "Localidad": base_record["Localidad"],
                    "Barrio": base_record["Barrio"],
                    "Hallazgo encontrado": base_record["Hallazgo encontrado"],
                    "IntervenciÃ³n": base_record["IntervenciÃ³n"],
                    "Cuenta contrato": base_record["Cuenta contrato"],
                    "Resultado": resultado_value,
                    "ObservaciÃ³n": observacion_value,
                }
            )

        ultimo_resultado = next((item["Resultado"] for item in reversed(seguimientos) if item["Resultado"]), "")
        ultima_observacion = next((item["ObservaciÃ³n"] for item in reversed(seguimientos) if item["ObservaciÃ³n"]), "")

        case_records.append(
            {
                **base_record,
                "Total_seguimientos": len(seguimientos),
                "Ultimo_resultado": ultimo_resultado,
                "Ultima_observacion": ultima_observacion,
                "Tiene_seguimiento": "Si" if seguimientos else "No",
            }
        )

    def build_option_values(items: list[dict[str, Any]], field: str) -> list[str]:
        values = sorted(
            {
                str(item.get(field) or "").strip()
                for item in items
                if str(item.get(field) or "").strip()
            },
            key=lambda item: item.lower(),
        )
        return values

    return {
        "sheet_used": target_sheet,
        "available_sheets": available_sheets,
        "source_columns": [str(c) for c in df.columns],
        "source_total_rows": len(df.index),
        "all_estatus_options": [],
        "all_estado_options": [],
        "chart_records": [],
        "report_records": [],
        "source_preview": serialize_for_json(df, limit=20),
        "admin_control_records": [],
        "penal_control_records": [],
        "no_procedente_control_records": [],
        "medidores_total": 0,
        "medidores_retirado_por": [],
        "medidores_concepto": [],
        "medidores_pendientes": [],
        "medidores_liquidacion_m3": [],
        "medidores_pendientes_liquidar": 0,
        "casos_especiales_records": case_records,
        "casos_especiales_seguimientos": seguimiento_records,
        "casos_especiales_filter_options": {
            "zona": build_option_values(case_records, "Zona"),
            "localidad": build_option_values(case_records, "Localidad"),
            "barrio": build_option_values(case_records, "Barrio"),
            "hallazgo": build_option_values(case_records, "Hallazgo encontrado"),
            "resultado": build_option_values(seguimiento_records, "Resultado"),
        },
        "casos_inicial_recuperacion_records": [],
        "casos_inicial_recuperacion_filter_options": {
            "localidad": [],
            "barrio": [],
            "hallazgo": [],
            "situacion": [],
            "accion_administrativa": [],
            "estado_deuda": [],
        },
    }


def build_report_records(df: pd.DataFrame) -> list[dict[str, str]]:
    column_candidates = {
        "Aviso_T2": ["Aviso_T2"],
        "Fecha_Aviso": ["Fecha_Aviso"],
        "Cuenta Contrato": ["Cuenta Contrato", "Cuenta Contrato "],
        "Anomalia_Visita": ["Anomalia_Visita", "Anomalia_Visita "],
        "Estatus": ["Estatus"],
        "Estado": ["Estado"],
        "Observaciones": ["Observaciones"],
    }

    resolved: dict[str, str | None] = {}
    columns = [str(c) for c in df.columns]
    for output_name, candidates in column_candidates.items():
        resolved[output_name] = None
        for candidate in candidates:
            found = find_column(columns, candidate)
            if found:
                resolved[output_name] = found
                break

    work = pd.DataFrame()
    for output_name, source_name in resolved.items():
        if source_name:
            work[output_name] = df[source_name]
        else:
            work[output_name] = ""

    for col_name in work.columns:
        if pd.api.types.is_datetime64_any_dtype(work[col_name]):
            work[col_name] = work[col_name].dt.strftime("%Y-%m-%d")
        else:
            work[col_name] = work[col_name].fillna("").astype(str).str.strip()

    return work.to_dict(orient="records")


def process_excel_bytes(file_bytes: bytes, sheet_name: str | None, base_mode: str | None = None) -> dict[str, Any]:
    excel_file = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    available_sheets = list(excel_file.sheet_names)
    normalized_mode = normalize_text(base_mode)

    selected_sheet = sheet_name.strip() if sheet_name else DEFAULT_SHEET
    target_sheet = find_column(available_sheets, selected_sheet)
    if target_sheet is None:
        normalized_selected_sheet = normalize_text(selected_sheet).replace("1.", "1 ").replace("-", " ")
        for candidate in available_sheets:
            normalized_candidate = normalize_text(candidate).replace("1.", "1 ").replace("-", " ")
            if normalized_candidate == normalized_selected_sheet:
                target_sheet = candidate
                break
    if target_sheet is None and normalized_mode in {"casos especiales", "casos_especiales"}:
        normalized_selected_sheet = normalize_text(selected_sheet).replace("?", "").replace("1.", "1 ").replace("-", " ")
        for candidate in available_sheets:
            normalized_candidate = normalize_text(candidate).replace("?", "").replace("1.", "1 ").replace("-", " ")
            if "visita inicial" in normalized_selected_sheet and "visita inicial" in normalized_candidate:
                target_sheet = candidate
                break
            if "visitas seguimiento" in normalized_selected_sheet and "visitas seguimiento" in normalized_candidate:
                target_sheet = candidate
                break
    if target_sheet is None:
        if not sheet_name or normalized_mode == "medidores":
            target_sheet = available_sheets[0]
        else:
            raise ValueError(
                f"No se encontro la hoja '{selected_sheet}'. Hojas disponibles: {available_sheets}"
            )

    df = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    if normalized_mode == "medidores":
        return build_medidores_payload(df, target_sheet, available_sheets)
    if normalized_mode == "casos especiales" or normalized_mode == "casos_especiales":
        return build_casos_especiales_payload(df, target_sheet, available_sheets)
    return build_administrativa_payload(df, target_sheet, available_sheets)



def get_graph_access_token() -> str:
    tenant_id = (os.getenv("MS_TENANT_ID") or "").strip()
    client_id = (os.getenv("MS_CLIENT_ID") or "").strip()
    client_secret = (os.getenv("MS_CLIENT_SECRET") or "").strip()
    if not tenant_id or not client_id or not client_secret:
        raise ValueError("Faltan credenciales de Microsoft Graph (MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET).")

    now = time.time()
    token = GRAPH_TOKEN_CACHE.get("access_token")
    expires_at = float(GRAPH_TOKEN_CACHE.get("expires_at") or 0)
    if token and now < (expires_at - 60):
        return str(token)

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    with httpx.Client(timeout=30.0) as client:
        response = client.post(token_url, data=payload)
        if response.status_code >= 400:
            raise ValueError(f"No se pudo obtener token de Graph (HTTP {response.status_code}).")
        token_data = response.json()
        access_token = token_data.get("access_token")
        expires_in = int(token_data.get("expires_in", 3600))
        if not access_token:
            raise ValueError("La respuesta de Graph no incluyo access_token.")
        GRAPH_TOKEN_CACHE["access_token"] = access_token
        GRAPH_TOKEN_CACHE["expires_at"] = now + expires_in
        return str(access_token)


def fetch_excel_from_sharepoint_graph(sharepoint_url: str) -> tuple[bytes, str]:
    encoded = base64.urlsafe_b64encode(sharepoint_url.encode("utf-8")).decode("utf-8")
    sharing_token = f"u!{encoded.rstrip('=')}"
    endpoint = f"https://graph.microsoft.com/v1.0/shares/{sharing_token}/driveItem/content"
    headers = {"Authorization": f"Bearer {get_graph_access_token()}"}

    with httpx.Client(follow_redirects=True, timeout=60.0, headers=headers) as client:
        response = client.get(endpoint)
        if response.status_code >= 400:
            raise ValueError(f"Graph no pudo leer el archivo (HTTP {response.status_code}).")
        payload = response.content
        if not payload:
            raise ValueError("Graph devolvio un archivo vacio.")
        filename = urlparse(sharepoint_url).path.split("/")[-1] or "sharepoint.xlsx"
        return payload, filename


def fetch_excel_from_url(sharepoint_url: str) -> tuple[bytes, str]:
    url = sharepoint_url.strip()
    if not url:
        raise ValueError("El link de SharePoint esta vacio.")
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        raise ValueError("El link debe iniciar con http:// o https://")

    host = parsed.netloc.lower()
    if "sharepoint.com" in host or "onedrive.live.com" in host:
        try:
            return fetch_excel_from_sharepoint_graph(url)
        except Exception:
            pass

    headers = {"User-Agent": "AlertasProcesos/1.0"}
    with httpx.Client(follow_redirects=True, timeout=60.0, headers=headers) as client:
        response = client.get(url)
        if response.status_code >= 400:
            raise ValueError(f"No se pudo descargar el archivo desde SharePoint (HTTP {response.status_code}).")
        content_type = response.headers.get("content-type", "").lower()
        if "html" in content_type and "excel" not in content_type:
            raise ValueError("La URL devolvio HTML. Usa link de descarga directa o Graph configurado.")
        data = response.content
        if not data:
            raise ValueError("La descarga devolvio archivo vacio.")
        filename = response.headers.get("content-disposition", "") or parsed.path.split("/")[-1]
        return data, filename


def graph_config_status() -> dict[str, Any]:
    tenant_id = (os.getenv("MS_TENANT_ID") or "").strip()
    client_id = (os.getenv("MS_CLIENT_ID") or "").strip()
    client_secret = (os.getenv("MS_CLIENT_SECRET") or "").strip()
    return {
        "configured": bool(tenant_id and client_id and client_secret),
        "tenant_id_present": bool(tenant_id),
        "client_id_present": bool(client_id),
        "client_secret_present": bool(client_secret),
    }


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/alerts/preview")
async def preview_alerts(
    file: UploadFile | None = File(default=None),
    sharepoint_url: str | None = Form(default=None),
    sheet_name: str | None = Form(default=None),
    base_mode: str | None = Form(default=None),
) -> dict[str, Any]:
    try:
        payload: bytes | None = None
        if sharepoint_url and sharepoint_url.strip():
            payload, _ = fetch_excel_from_url(sharepoint_url)
        elif file is not None:
            if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
                raise ValueError("Sube un archivo Excel valido (.xlsx).")
            payload = await file.read()
            if not payload:
                raise ValueError("El archivo esta vacio.")
        else:
            raise ValueError("Debes subir un archivo o pegar un link de SharePoint.")
        return process_excel_bytes(payload, sheet_name, base_mode)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Error procesando archivo: {exc}") from exc


@app.post("/api/report/download")
async def download_filtered_report(payload: dict[str, Any] = Body(...)):
    records = payload.get("records") or []
    filename = str(payload.get("filename") or "informe_filtrado.xlsx").strip() or "informe_filtrado.xlsx"
    if not isinstance(records, list):
        raise HTTPException(status_code=400, detail="El cuerpo debe incluir una lista 'records'.")

    rows: list[dict[str, Any]] = []
    for record in records:
        if not isinstance(record, dict):
            continue
        rows.append({str(key): ("" if value is None else str(value).strip()) for key, value in record.items()})

    output = BytesIO()
    pd.DataFrame(rows).to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    safe_filename = filename if filename.lower().endswith(".xlsx") else f"{filename}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{safe_filename}"'}
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.post("/api/sharepoint/diagnostic")
def sharepoint_diagnostic(sharepoint_url: str = Form(...)) -> dict[str, Any]:
    url = (sharepoint_url or "").strip()
    if not url:
        raise HTTPException(status_code=400, detail="Debes enviar sharepoint_url.")

    config = graph_config_status()
    result: dict[str, Any] = {"graph": config, "url": url}
    if config["configured"]:
        try:
            token = get_graph_access_token()
            result["graph"]["token_ok"] = bool(token)
        except Exception as exc:
            result["graph"]["token_ok"] = False
            result["graph"]["token_error"] = str(exc)
    else:
        result["graph"]["token_ok"] = False
        result["graph"]["token_error"] = "Graph no configurado."

    try:
        payload, filename = fetch_excel_from_url(url)
        result["download_ok"] = True
        result["filename"] = filename
        result["bytes"] = len(payload)
    except Exception as exc:
        result["download_ok"] = False
        result["download_error"] = str(exc)
    return result



