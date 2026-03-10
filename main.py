import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

import pandas as pd


ADMIN_PENAL_THRESHOLD_DAYS = 150
LIQ_PROCED_THRESHOLD_DAYS = 60


def find_column(columns, target):
    """Find a column by normalized name (trim + lowercase)."""
    target_norm = target.strip().lower()
    for col in columns:
        if str(col).strip().lower() == target_norm:
            return col
    return None


def map_columns(df, required):
    mapping = {}
    missing = []
    for original_name, final_name in required:
        found = find_column(df.columns, original_name)
        if found is None:
            missing.append(original_name)
        else:
            mapping[found] = final_name
    if missing:
        raise ValueError(f"No se encontraron columnas requeridas: {', '.join(missing)}")
    return mapping


def build_block(df, required_columns, tipo, regla):
    col_map = map_columns(df, required_columns)
    block = df[list(col_map.keys())].rename(columns=col_map).copy()
    block["Tipo"] = tipo
    block["Regla"] = regla
    return block


def col(df, name):
    found = find_column(df.columns, name)
    if found is None:
        raise ValueError(f"No se encontro la columna requerida: {name}")
    return found


def series_empty(series):
    s = series.astype(str).str.strip().str.lower()
    return series.isna() | s.isin({"", "nan", "none", "null", "sin dato", "-"})


def coalesce_series(df, preferred, fallback):
    p = df[col(df, preferred)]
    f = df[col(df, fallback)]
    out = p.copy()
    mask = series_empty(out)
    out.loc[mask] = f.loc[mask]
    return out


def build_alerts_dataframe(df):
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
        df,
        administrativo_required,
        "Administrativo",
        f"Vencimiento <= {ADMIN_PENAL_THRESHOLD_DAYS} dias",
    )
    penal_df = build_block(
        df,
        penal_required,
        "Penal",
        f"Vencimiento <= {ADMIN_PENAL_THRESHOLD_DAYS} dias",
    )

    admin_df["Dias"] = pd.to_numeric(admin_df["Dias"], errors="coerce")
    penal_df["Dias"] = pd.to_numeric(penal_df["Dias"], errors="coerce")

    admin_alerts = admin_df[admin_df["Dias"] <= ADMIN_PENAL_THRESHOLD_DAYS].copy()
    penal_alerts = penal_df[penal_df["Dias"] <= ADMIN_PENAL_THRESHOLD_DAYS].copy()

    cuenta_col = col(df, "Cuenta Contrato")
    interlocutor_col = col(df, "Interlocutor")
    direccion_col = col(df, "Dirección")
    ciudad_col = col(df, "Ciudad")
    fecha_admin_col = col(df, "Fecha de Vencimiento")
    dias_admin_col = col(df, "DÍAS")
    liquidacion_col = col(df, "Liquidación")
    analisis_col = col(df, "ANALISIS")
    doc_sd_col = col(df, "Documento_SD")
    memo_juridica_col = col(df, "Memorando_Juridica")
    salida_doc_col = col(df, "Documento_Factura_salida")

    dias_admin = pd.to_numeric(df[dias_admin_col], errors="coerce")
    responsable_base = coalesce_series(df, "Responsable Administrativo", "Responsable Penal")

    liq_proc_pending = series_empty(df[liquidacion_col]) | series_empty(df[analisis_col])
    liq_proc_mask = (dias_admin <= LIQ_PROCED_THRESHOLD_DAYS) & liq_proc_pending

    liq_proc_alerts = pd.DataFrame(
        {
            "Cuenta Contrato": df[cuenta_col],
            "Interlocutor": df[interlocutor_col],
            "Dirección": df[direccion_col],
            "Ciudad": df[ciudad_col],
            "Responsable": responsable_base,
            "Fecha_Vencimiento": df[fecha_admin_col],
            "Dias": dias_admin,
            "Tipo": "Liquidacion/Procedencia",
            "Regla": f"Pendiente y <= {LIQ_PROCED_THRESHOLD_DAYS} dias",
        }
    )[liq_proc_mask].copy()

    enviado_mask = (~series_empty(df[doc_sd_col])) | (~series_empty(df[memo_juridica_col]))
    sin_salida_mask = series_empty(df[salida_doc_col])
    control_mask = enviado_mask & sin_salida_mask

    control_alerts = pd.DataFrame(
        {
            "Cuenta Contrato": df[cuenta_col],
            "Interlocutor": df[interlocutor_col],
            "Dirección": df[direccion_col],
            "Ciudad": df[ciudad_col],
            "Responsable": responsable_base,
            "Fecha_Vencimiento": df[fecha_admin_col],
            "Dias": dias_admin,
            "Tipo": "Control",
            "Regla": "Enviado sin salida",
        }
    )[control_mask].copy()

    alertas = pd.concat(
        [admin_alerts, penal_alerts, liq_proc_alerts, control_alerts],
        ignore_index=True,
    )
    alertas = alertas.sort_values(by=["Dias", "Tipo"], ascending=[True, True], na_position="last")

    return alertas


def parse_bool_env(name, default):
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "si", "on"}


def load_email_mapping(base_dir):
    """Load Responsable->Correo mapping from responsables_correos.csv."""
    mapping = {}
    mapping_file = base_dir / "responsables_correos.csv"

    if not mapping_file.exists():
        return mapping

    map_df = pd.read_csv(mapping_file, encoding="utf-8-sig")
    map_df.columns = [str(c).strip() for c in map_df.columns]
    required_cols = {"Responsable", "Correo"}
    if not required_cols.issubset(set(map_df.columns)):
        raise ValueError(
            "El archivo responsables_correos.csv debe tener columnas Responsable y Correo"
        )

    for _, row in map_df.dropna(subset=["Responsable", "Correo"]).iterrows():
        responsable = str(row["Responsable"]).strip()
        correo = str(row["Correo"]).strip()
        if responsable and correo:
            mapping[responsable.lower()] = correo

    return mapping


def resolve_recipient(row, email_mapping):
    responsable = str(row.get("Responsable", "")).strip()
    if not responsable:
        return None
    return email_mapping.get(responsable.lower())


def send_email_alerts(alertas, base_dir):
    enabled = parse_bool_env("ALERT_EMAIL_ENABLED", True)
    if not enabled:
        print("Envio de correos deshabilitado por ALERT_EMAIL_ENABLED.")
        return 0, 0

    smtp_host = os.getenv("ALERT_SMTP_HOST", "").strip()
    smtp_port = int(os.getenv("ALERT_SMTP_PORT", "587").strip())
    smtp_user = os.getenv("ALERT_SMTP_USER", "").strip()
    smtp_password = os.getenv("ALERT_SMTP_PASSWORD", "").strip()
    smtp_from = os.getenv("ALERT_EMAIL_FROM", "").strip() or smtp_user
    use_tls = parse_bool_env("ALERT_SMTP_USE_TLS", True)

    if not smtp_host or not smtp_user or not smtp_password or not smtp_from:
        print(
            "Credenciales SMTP incompletas. Defina ALERT_SMTP_HOST, ALERT_SMTP_USER, "
            "ALERT_SMTP_PASSWORD y ALERT_EMAIL_FROM para enviar correos."
        )
        return 0, 0

    email_mapping = load_email_mapping(base_dir)
    if not email_mapping:
        print(
            "No hay mapeo de correos. Cree responsables_correos.csv con columnas "
            "Responsable,Correo."
        )
        return 0, 0

    alertas_email = alertas.copy()
    alertas_email["Correo"] = alertas_email.apply(
        lambda row: resolve_recipient(row, email_mapping), axis=1
    )

    sin_correo = alertas_email[alertas_email["Correo"].isna()].copy()
    con_correo = alertas_email[alertas_email["Correo"].notna()].copy()

    if not sin_correo.empty:
        pendientes_path = base_dir / "salida" / "alertas_sin_correo.csv"
        sin_correo.to_csv(pendientes_path, index=False, encoding="utf-8-sig")
        print(
            f"Alertas sin correo asignado: {len(sin_correo)}. "
            f"Revisar {pendientes_path.name}."
        )

    if con_correo.empty:
        print("No hay alertas con correo destinatario.")
        return 0, len(sin_correo)

    enviados = 0
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        if use_tls:
            server.starttls()
        server.login(smtp_user, smtp_password)

        for correo, group in con_correo.groupby("Correo"):
            responsable = str(group["Responsable"].iloc[0]).strip() or "responsable"
            subject = f"Alerta de visitas pendientes - {responsable}"
            body_lines = [
                f"Hola {responsable},",
                "",
                "Estas son tus alertas de visitas:",
                "",
                group[
                    [
                        "Regla",
                        "Tipo",
                        "Cuenta Contrato",
                        "Interlocutor",
                        "Dirección",
                        "Ciudad",
                        "Fecha_Vencimiento",
                        "Dias",
                    ]
                ].to_string(index=False),
                "",
                f"Total alertas: {len(group)}",
            ]

            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = smtp_from
            msg["To"] = correo
            msg.set_content("\n".join(body_lines))

            server.send_message(msg)
            enviados += 1

    return enviados, len(sin_correo)


def main():
    print("ENTRO AL MAIN")
    try:
        print("Leyendo Excel...")

        base_dir = Path(__file__).resolve().parent
        input_file = base_dir / "penal_24_25.xlsx"
        output_dir = base_dir / "salida"
        output_xlsx = output_dir / "alertas_hoy.xlsx"
        output_csv = output_dir / "alertas_hoy.csv"

        if not input_file.exists():
            raise FileNotFoundError(f"No existe el archivo: {input_file}")

        excel_file = pd.ExcelFile(input_file, engine="openpyxl")
        target_sheet = find_column(excel_file.sheet_names, "Procesos Adminis_Penal")
        if target_sheet is None:
            raise ValueError(
                "No se encontro la hoja 'Procesos Adminis_Penal'. "
                f"Hojas disponibles: {excel_file.sheet_names}"
            )

        df = pd.read_excel(input_file, sheet_name=target_sheet, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]

        print("Construyendo alertas...")

        alertas = build_alerts_dataframe(df)

        output_dir.mkdir(parents=True, exist_ok=True)

        alertas.to_excel(output_xlsx, index=False)
        alertas.to_csv(output_csv, index=False, encoding="utf-8-sig")

        enviados, sin_correo = send_email_alerts(alertas, base_dir)

        print("LISTO.")
        print(f"Alertas encontradas: {len(alertas)}")
        print(f"Correos enviados: {enviados}")
        print(f"Alertas sin correo: {sin_correo}")
        print(alertas.head(10).to_string(index=False))

    except Exception as e:
        print(f"ERROR: {e}")
        raise


if __name__ == "__main__":
    main()
