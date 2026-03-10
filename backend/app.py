from __future__ import annotations

import base64
import os
import smtplib
import time
import unicodedata
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import httpx
import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel


FIVE_MONTH_DAYS = 150
DEFAULT_SHEET = "Procesos Adminis_Penal"
PROJECT_ROOT = Path(__file__).resolve().parent.parent
MAPPING_FILE = PROJECT_ROOT / "responsables_correos.csv"
GRAPH_TOKEN_CACHE: dict[str, Any] = {"access_token": None, "expires_at": 0}
REPORT_TO_EMAIL = "wrodriguezg@acueducto.com.co"


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


def load_email_mapping() -> dict[str, str]:
    if not MAPPING_FILE.exists():
        return {}
    map_df = pd.read_csv(MAPPING_FILE, encoding="utf-8-sig")
    map_df.columns = [str(c).strip() for c in map_df.columns]
    required = {"Responsable", "Correo"}
    if not required.issubset(set(map_df.columns)):
        return {}
    mapping: dict[str, str] = {}
    for _, row in map_df.dropna(subset=["Responsable"]).iterrows():
        responsable = str(row["Responsable"]).strip()
        correo = str(row["Correo"]).strip() if pd.notna(row["Correo"]) else ""
        if responsable:
            mapping[responsable] = correo
    return mapping


def save_email_mapping(entries: list[dict[str, str]]) -> None:
    rows = []
    for item in entries:
        responsable = str(item.get("Responsable", "")).strip()
        correo = str(item.get("Correo", "")).strip()
        if responsable:
            rows.append({"Responsable": responsable, "Correo": correo})
    out_df = pd.DataFrame(rows).drop_duplicates(subset=["Responsable"], keep="last")
    out_df = out_df.sort_values(by=["Responsable"], ascending=True)
    out_df.to_csv(MAPPING_FILE, index=False, encoding="utf-8-sig")


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
        ("Dirección", "Dirección"),
        ("Ciudad", "Ciudad"),
        ("Responsable Administrativo", "Responsable"),
        ("Fecha de Vencimiento", "Fecha_Vencimiento"),
        ("DÍAS", "Dias"),
    ]
    penal_required = [
        ("Cuenta Contrato", "Cuenta Contrato"),
        ("Interlocutor", "Interlocutor"),
        ("Dirección", "Dirección"),
        ("Ciudad", "Ciudad"),
        ("Responsable Penal", "Responsable"),
        ("Fecha de Vencimiento.1", "Fecha_Vencimiento"),
        ("DÍAS.1", "Dias"),
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
    direccion_col = col(df, "Dirección")
    ciudad_col = col(df, "Ciudad")
    fecha_admin_col = col(df, "Fecha de Vencimiento")
    dias_admin_col = col(df, "DÍAS")
    liquidacion_col = col(df, "Liquidación")
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
    # responsable sale de la columna Liquidación y aplica ventana de 45 dias.
    estatus_col = col(df, "Estatus")
    estatus_norm = df[estatus_col].astype(str).str.strip().str.lower()
    proc_status_mask = estatus_norm.str.contains("pendiente determinar procedencia", na=False)

    proc_df = pd.DataFrame(
        {
            "Aviso": df[aviso_col] if aviso_col else "",
            "Cuenta Contrato": df[cuenta_col],
            "Interlocutor": df[interlocutor_col],
            "Dirección": df[direccion_col],
            "Ciudad": df[ciudad_col],
            "Responsable": df[liquidacion_col].astype(str).str.strip(),
            "Quien_Liquida": df[liquidacion_col].astype(str).str.strip(),
            "Fecha_Vencimiento": df[fecha_admin_col],
            "Dias": dias_admin,
            "Tipo": "Pendiente determinar procedencia",
            "Regla": "Estatus pendiente determinar procedencia (<=45 dias)",
            "Estatus": df[estatus_col].astype(str).str.strip(),
        }
    )
    proc_df = proc_df[proc_status_mask & (proc_df["Dias"] <= 45)].copy()
    proc_df["Responsable"] = proc_df["Responsable"].replace(
        {"": "Pendiente por asignar", "nan": "Pendiente por asignar", "None": "Pendiente por asignar"}
    )
    proc_df.loc[proc_df["Responsable"].isna(), "Responsable"] = "Pendiente por asignar"
    proc_df["EmailTrigger"] = proc_df["Dias"].apply(
        lambda d: "PENDIENTE_LIQUIDACION_30" if pd.notna(d) and int(d) == 30 else ""
    )
    pending_status_mask = estatus_norm.str.contains("para administrativo|para expediente", na=False)

    pending_status_df = pd.DataFrame(
        {
            "Aviso": df[aviso_col] if aviso_col else "",
            "Cuenta Contrato": df[cuenta_col],
            "Interlocutor": df[interlocutor_col],
            "Dirección": df[direccion_col],
            "Ciudad": df[ciudad_col],
            "Responsable": responsable_base,
            "Quien_Liquida": df[liquidacion_col].astype(str).str.strip(),
            "Fecha_Vencimiento": df[fecha_admin_col],
            "Dias": dias_admin,
            "Tipo": "Pendiente Control",
            "Regla": "Estatus para expediente/para administrativo (incluye mixtos)",
            "Estatus": df[estatus_col].astype(str).str.strip(),
        }
    )
    pending_status_df = pending_status_df[pending_status_mask & (pending_status_df["Dias"] <= FIVE_MONTH_DAYS)].copy()
    pending_status_df["EmailTrigger"] = pending_status_df["Dias"].apply(
        lambda d: "PENDIENTE_LIQUIDACION_30" if pd.notna(d) and int(d) == 30 else ""
    )

    # If both responsables are empty, classify as pending assignment.
    resp_adm = df[responsable_adm_col]
    resp_pen = df[responsable_pen_col]
    no_resp_mask = series_empty(resp_adm) & series_empty(resp_pen)
    pending_status_df.loc[no_resp_mask.reindex(pending_status_df.index, fill_value=False), "Responsable"] = "Pendiente por asignar"
    pending_status_df.loc[
        no_resp_mask.reindex(pending_status_df.index, fill_value=False), "Quien_Liquida"
    ] = "Pendiente por asignar"

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


def extract_responsables(df: pd.DataFrame) -> list[str]:
    admin = df[col(df, "Responsable Administrativo")]
    penal = df[col(df, "Responsable Penal")]
    merged = pd.concat([admin, penal], ignore_index=True).dropna().astype(str)
    all_names: list[str] = []
    for value in merged.tolist():
        all_names.extend(split_responsables(value))
    return sorted(set(all_names))


def compact_day_columns(series: pd.Series, min_day: int = 0, max_day: int = FIVE_MONTH_DAYS) -> list[int]:
    values = sorted({int(v) for v in series.dropna().tolist() if min_day <= int(v) <= max_day})
    if len(values) <= 32:
        return values
    head = values[:24]
    milestones = [30, 45, 60, 90, 120, 150]
    tail = [m for m in milestones if m in values]
    merged = sorted(set(head + tail))
    return merged


def build_day_board(
    alertas: pd.DataFrame,
    responsables: list[str],
    key: str,
    title: str,
    description: str,
) -> dict[str, Any]:
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

    for name in responsables:
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


def build_report_records(all_alerts: pd.DataFrame) -> list[dict[str, Any]]:
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
        ("Cuenta Contrato", ""),
        ("Ciudad", ""),
        ("Fecha_Vencimiento", None),
        ("Regla", ""),
        ("Tipo", ""),
    ]:
        if optional_col not in work.columns:
            work[optional_col] = default_value

    out = work[["Responsable", "Tipo", "Regla", "Cuenta Contrato", "Ciudad", "Fecha_Vencimiento", "DiasInt"]].copy()
    out["Fecha_Vencimiento"] = pd.to_datetime(out["Fecha_Vencimiento"], errors="coerce").dt.strftime("%Y-%m-%d")
    out["Responsable"] = out["Responsable"].astype(str).str.strip()
    out["Tipo"] = out["Tipo"].astype(str).str.strip()
    out["Regla"] = out["Regla"].astype(str).str.strip()
    out["Cuenta Contrato"] = out["Cuenta Contrato"].astype(str).str.strip()
    out["Ciudad"] = out["Ciudad"].astype(str).str.strip()
    out = out.sort_values(by=["Responsable", "DiasInt", "Tipo"], ascending=[True, True, True])
    return out.to_dict(orient="records")


def send_report_email(records: list[dict[str, Any]]) -> dict[str, Any]:
    smtp_host = (os.getenv("ALERT_SMTP_HOST") or "").strip()
    smtp_port = int((os.getenv("ALERT_SMTP_PORT") or "587").strip())
    smtp_user = (os.getenv("ALERT_SMTP_USER") or "").strip()
    smtp_password = (os.getenv("ALERT_SMTP_PASSWORD") or "").strip()
    smtp_from = (os.getenv("ALERT_EMAIL_FROM") or "").strip() or smtp_user
    use_tls = (os.getenv("ALERT_SMTP_USE_TLS") or "1").strip().lower() in {"1", "true", "yes", "si", "on"}

    if not smtp_host or not smtp_user or not smtp_password or not smtp_from:
        raise ValueError(
            "Credenciales SMTP incompletas. Configure ALERT_SMTP_HOST, ALERT_SMTP_USER, "
            "ALERT_SMTP_PASSWORD y ALERT_EMAIL_FROM."
        )

    if not records:
        raise ValueError("No hay registros para enviar en el informe.")

    df = pd.DataFrame(records)
    for required in ["Responsable", "Tipo", "Cuenta Contrato", "Ciudad", "Fecha_Vencimiento", "DiasInt"]:
        if required not in df.columns:
            raise ValueError(f"Falta columna en registros del informe: {required}")

    lines = [
        "Cordial saludo,",
        "",
        "Por medio del presente se remite el informe de alertas con las acciones pendientes identificadas.",
        "A continuacion se detalla el consolidado por responsable y tipo de alerta:",
        "",
    ]

    grouped = df.groupby("Responsable", dropna=False)
    for responsable, group in grouped:
        responsable_name = str(responsable).strip() or "Sin responsable"
        tipo_counts = group["Tipo"].value_counts()
        tipo_summary = ", ".join([f"{tipo}: {int(count)}" for tipo, count in tipo_counts.items()])
        lines.append(f"{responsable_name} -> Total alertas: {len(group)} ({tipo_summary})")
        for _, row in group.head(12).iterrows():
            lines.append(
                f"  - Cuenta {row['Cuenta Contrato']} | {row['Tipo']} | {row['Ciudad']} | "
                f"Vence: {row['Fecha_Vencimiento']} | Dias: {row['DiasInt']}"
            )
        if len(group) > 12:
            lines.append(f"  - ... {len(group) - 12} alertas adicionales")
        lines.append("")

    lines.append("Este informe se genera automaticamente desde el tablero de control de alertas.")
    body = "\n".join(lines)

    # Adjunta el consolidado general en formato Excel.
    export_df = df.copy()
    export_df = export_df.rename(columns={"DiasInt": "Dias"})
    attachment_name = f"informe_alertas_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    excel_buffer = BytesIO()
    export_df.to_excel(excel_buffer, index=False, sheet_name="Informe", engine="openpyxl")
    excel_bytes = excel_buffer.getvalue()

    msg = EmailMessage()
    msg["Subject"] = "Informe de alertas"
    msg["From"] = smtp_from
    msg["To"] = REPORT_TO_EMAIL
    msg.set_content(body)
    msg.add_attachment(
        excel_bytes,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_name,
    )

    sent_personal = 0
    skipped_personal = 0
    no_email_personal = 0
    invalid_email_personal = 0
    personal_recipients: list[str] = []

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        if use_tls:
            server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

        # Correos personalizados por responsable (solo sus propias alertas).
        mapping = load_email_mapping()
        normalized_mapping: dict[str, str] = {
            normalize_text(k): str(v).strip() for k, v in mapping.items() if str(k).strip()
        }

        for responsable, group in df.groupby("Responsable", dropna=False):
            responsable_name = str(responsable).strip() or "Sin responsable"
            target_email = normalized_mapping.get(normalize_text(responsable_name), "").strip()
            if not target_email:
                no_email_personal += 1
                skipped_personal += 1
                continue
            if "@" not in target_email or "." not in target_email.split("@")[-1]:
                invalid_email_personal += 1
                skipped_personal += 1
                continue

            personal_lines = [
                f"Cordial saludo, {responsable_name}.",
                "",
                "Se remite su informe personalizado de alertas pendientes.",
                f"Total de alertas asignadas: {len(group)}.",
                "",
                "Detalle:",
            ]
            for _, row in group.sort_values(by=["DiasInt", "Tipo"], ascending=[True, True]).head(25).iterrows():
                personal_lines.append(
                    f"  - Cuenta {row['Cuenta Contrato']} | {row['Tipo']} | {row['Ciudad']} | "
                    f"Vence: {row['Fecha_Vencimiento']} | Dias: {row['DiasInt']}"
                )
            if len(group) > 25:
                personal_lines.append(f"  - ... {len(group) - 25} alertas adicionales (ver adjunto).")
            personal_lines.extend(
                [
                    "",
                    "Adjunto se incluye su consolidado en Excel.",
                    "Mensaje generado automaticamente por el sistema de alertas.",
                ]
            )

            personal_df = group.copy().rename(columns={"DiasInt": "Dias"})
            personal_buffer = BytesIO()
            personal_df.to_excel(personal_buffer, index=False, sheet_name="MisAlertas", engine="openpyxl")
            safe_name = (
                normalize_text(responsable_name)
                .replace(" ", "_")
                .replace("/", "_")
                .replace("\\", "_")
                .replace(":", "_")
            )
            personal_attachment = f"alertas_{safe_name}_{time.strftime('%Y%m%d')}.xlsx"

            personal_msg = EmailMessage()
            personal_msg["Subject"] = f"Alertas asignadas - {responsable_name}"
            personal_msg["From"] = smtp_from
            personal_msg["To"] = target_email
            personal_msg.set_content("\n".join(personal_lines))
            personal_msg.add_attachment(
                personal_buffer.getvalue(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=personal_attachment,
            )
            server.send_message(personal_msg)
            sent_personal += 1
            personal_recipients.append(f"{responsable_name} <{target_email}>")

    return {
        "to": REPORT_TO_EMAIL,
        "subject": "Informe de alertas",
        "rows": len(df),
        "attachment": attachment_name,
        "personal_sent": sent_personal,
        "personal_skipped": skipped_personal,
        "personal_missing_email": no_email_personal,
        "personal_invalid_email": invalid_email_personal,
        "personal_recipients": personal_recipients,
    }


def serialize_for_json(df: pd.DataFrame, limit: int = 30) -> list[dict[str, Any]]:
    sample = df.head(limit).copy()
    for col_name in sample.columns:
        if pd.api.types.is_datetime64_any_dtype(sample[col_name]):
            sample[col_name] = sample[col_name].dt.strftime("%Y-%m-%d")
    sample = sample.where(pd.notna(sample), None)
    return sample.to_dict(orient="records")


def process_excel_bytes(file_bytes: bytes, sheet_name: str | None) -> dict[str, Any]:
    excel_file = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    available_sheets = list(excel_file.sheet_names)

    selected_sheet = sheet_name.strip() if sheet_name else DEFAULT_SHEET
    target_sheet = find_column(available_sheets, selected_sheet)
    if target_sheet is None:
        if not sheet_name:
            target_sheet = available_sheets[0]
        else:
            raise ValueError(
                f"No se encontro la hoja '{selected_sheet}'. Hojas disponibles: {available_sheets}"
            )

    df = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    all_alerts, proc_df, pending_status_df = build_alerts_dataframe(df)
    responsables = extract_responsables(df)
    extra_resp: set[str] = set()
    for frame in [proc_df, pending_status_df]:
        if not frame.empty and "Responsable" in frame.columns:
            vals = frame["Responsable"].dropna().astype(str).str.strip()
            vals = vals[(vals != "") & (vals.str.lower() != "nan")]
            extra_resp.update(vals.tolist())
    responsables = sorted(set(responsables).union(extra_resp))
    email_mapping = load_email_mapping()

    admin_5m = all_alerts[all_alerts["Tipo"] == "Administrativo"].copy()
    penal_5m = all_alerts[all_alerts["Tipo"] == "Penal"].copy()
    combinado_5m = pd.concat([admin_5m, penal_5m], ignore_index=True)

    boards = [
        build_day_board(
            proc_df,
            responsables,
            key="pendientes_procedencia",
            title="Tablero de Pendiente determinar procedencia (45 dias)",
            description="Estatus 'Pendiente determinar procedencia'. Responsable tomado de columna Liquidación.",
        ),
        build_day_board(
            penal_5m,
            responsables,
            key="penales_5m",
            title="Tablero de Penales 5 Meses",
            description="Alertas penales con vencimiento en 5 meses o menos.",
        ),
        build_day_board(
            admin_5m,
            responsables,
            key="administrativos_5m",
            title="Tablero de Administrativos 5 Meses",
            description="Alertas administrativas con vencimiento en 5 meses o menos.",
        ),
        build_day_board(
            combinado_5m,
            responsables,
            key="combinado_5m",
            title="Penal Administrativo Combinado 5 Meses",
            description="Vista consolidada penal + administrativo a 5 meses.",
        ),
    ]

    return {
        "policy": {
            "five_month_days": FIVE_MONTH_DAYS,
            "pending_liquidacion_alert_day": 30,
            "vencimiento_alert_day": 10,
            "cc_recipients_rule": "En dia 10 incluir copia a Andrea y jefe (configurable en correo).",
        },
        "sheet_used": target_sheet,
        "available_sheets": available_sheets,
        "source_columns": [str(c) for c in df.columns],
        "source_total_rows": int(len(df)),
        "source_preview": serialize_for_json(df, limit=20),
        "alerts_total_rows": int(len(all_alerts)),
        "alerts_preview": serialize_for_json(all_alerts, limit=60),
        "responsables": responsables,
        "email_mapping": email_mapping,
        "tableros": boards,
        "status_analysis": build_status_analysis(df),
        "control_dashboard": build_control_dashboard(pending_status_df),
        "analysis_records": build_analysis_records(pending_status_df),
        "report_records": build_report_records(all_alerts),
    }


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


class ResponsableCorreoEntry(BaseModel):
    Responsable: str
    Correo: str = ""


class SaveResponsablesRequest(BaseModel):
    entries: list[ResponsableCorreoEntry]


class SendReportRequest(BaseModel):
    records: list[dict[str, Any]]


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/alerts/preview")
async def preview_alerts(
    file: UploadFile | None = File(default=None),
    sharepoint_url: str | None = Form(default=None),
    sheet_name: str | None = Form(default=None),
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
        return process_excel_bytes(payload, sheet_name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Error procesando archivo: {exc}") from exc


@app.get("/api/responsables")
def get_responsables_mapping() -> dict[str, Any]:
    return {
        "mapping_file": str(MAPPING_FILE),
        "entries": [{"Responsable": k, "Correo": v} for k, v in load_email_mapping().items()],
    }


@app.post("/api/responsables/save")
def save_responsables_mapping(payload: SaveResponsablesRequest) -> dict[str, Any]:
    try:
        save_email_mapping([entry.model_dump() for entry in payload.entries])
        return {"ok": True, "saved_count": len(payload.entries), "mapping_file": str(MAPPING_FILE)}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"No se pudo guardar el mapeo: {exc}") from exc


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


@app.post("/api/report/send")
def send_report(payload: SendReportRequest) -> dict[str, Any]:
    try:
        result = send_report_email(payload.records)
        return {"ok": True, **result}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"No se pudo enviar informe: {exc}") from exc
