from __future__ import annotations

import base64
import calendar
import html
import io
import os
import re
import smtplib
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
ASSETS_DIR = APP_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "logo.png"
HOME_IMAGE_PATH = ASSETS_DIR / "inicio.png"
DEFAULT_EXCEL = DATA_DIR / "Matriz Acciones Fase II.xlsx"
CONTACTS_EXCEL = DATA_DIR / "Contactos.xlsx"
DEFAULT_SHEET = "MATRIZ Evaluación"
LOCAL_SQLITE = f"sqlite:///{APP_DIR / 'evaluacion.db'}"

BASE_COLUMNS = [
    "Ámbito", "Título", "Tipo", "Estado", "Medida", "Agente promotor",
    "Descripción", "Personas destinatarias", "Temporalidad", "Indicadores", "Recursos",
]

CONTACT_COLUMNS = [
    "Categoria", "Subcategoria", "Comisión", "Perfil", "Entidad", "Teléfono", "Persona", "Mail",
]

STATUS_OPTIONS = ["Sin evaluar", "Iniciado", "Completado", "No procede"]
INDICATOR_OPTIONS = ["Sin evaluar", "Parcial", "Cumplido", "No procede"]
PRIORITY_OPTIONS = ["Baja", "Media", "Alta", "Crítica"]
TASK_STATUS_OPTIONS = ["No iniciado", "En curso", "Completado", "Descartado"]
EVENT_TYPE_OPTIONS = ["Evento", "Reunión", "Actividad", "Hito", "Otro"]
CALENDAR_ITEM_TYPES = ["Todos", "Eventos/Reuniones", "Tareas"]
QUESTIONNAIRE_TYPES = {"promotora": "Persona Promotora", "participante": "Persona participante"}
LIKERT_OPTIONS = [1, 2, 3, 4, 5]
LIKERT_SELECT_OPTIONS = ["Sin marcar", 1, 2, 3, 4, 5]
YES_NO_OPTIONS = ["Sin responder", "Sí", "No", "No procede"]

# Paleta corporativa Erandio Mugitzen ari da!
COLOR_NAVY = "#1C3054"
COLOR_SKY = "#32A4CF"
COLOR_CORAL = "#E95C47"
COLOR_MAUVE = "#AE7CAA"
COLOR_GOLD = "#F9C14F"
BRAND_COLORS = [COLOR_NAVY, COLOR_SKY, COLOR_CORAL, COLOR_MAUVE, COLOR_GOLD]

# Perfiles de acceso. Las claves reales se pueden cambiar desde Streamlit Secrets.
# ADMIN_PASSWORD mantiene la clave de administración solicitada.
PROFILE_CONFIG = {
    "Administradora": {
        "secret": "ADMIN_PASSWORD",
        "default_password": "Temporal001*2026",
        "role": "admin",
        "scope": None,
    },
    "Alimentación": {
        "secret": "PASSWORD_ALIMENTACION",
        "default_password": "Alimentacion2026!",
        "role": "eje",
        "scope": "Alimentación",
    },
    "Deporte": {
        "secret": "PASSWORD_DEPORTE",
        "default_password": "Deporte2026!",
        "role": "eje",
        "scope": "Deporte",
    },
    "Educación": {
        "secret": "PASSWORD_EDUCACION",
        "default_password": "Educacion2026!",
        "role": "eje",
        "scope": "Educación",
    },
    "Ocio y tiempo libre": {
        "secret": "PASSWORD_OCIO",
        "default_password": "Ocio2026!",
        "role": "eje",
        "scope": "Ocio y tiempo libre",
    },
    "Bienestar emocional": {
        "secret": "PASSWORD_BIENESTAR",
        "default_password": "Bienestar2026!",
        "role": "eje",
        "scope": "Bienestar emocional",
    },
}


def apply_brand_theme() -> None:
    """Aplica una interfaz clara: blanco dominante y paleta corporativa como acento."""
    st.markdown(
        f"""
        <style>
        :root {{
            --brand-navy: {COLOR_NAVY};
            --brand-sky: {COLOR_SKY};
            --brand-coral: {COLOR_CORAL};
            --brand-mauve: {COLOR_MAUVE};
            --brand-gold: {COLOR_GOLD};
            --brand-white: #ffffff;
            --brand-soft: #f7f9fc;
            --brand-line: rgba(28, 48, 84, 0.14);
        }}

        /* Base: el blanco manda; la paleta se usa como acento. */
        .stApp {{
            background: var(--brand-white);
            color: var(--brand-navy);
        }}

        .main .block-container {{
            padding-top: 1.35rem;
            padding-bottom: 2.5rem;
            max-width: 1280px;
        }}

        h1, h2, h3, h4, h5, h6, .stMarkdown, label, p, span {{
            color: var(--brand-navy);
        }}

        h1 {{
            font-weight: 750;
            letter-spacing: -0.02em;
        }}

        h2, h3 {{
            font-weight: 700;
        }}

        /* Cabecera interna: clara, blanca y con acento inferior. */
        .brand-header {{
            background: var(--brand-white);
            border: 1px solid var(--brand-line);
            border-bottom: 5px solid var(--brand-sky);
            border-radius: 1rem;
            padding: 0.85rem 1.05rem;
            margin-bottom: 1.1rem;
            box-shadow: 0 1px 8px rgba(28, 48, 84, 0.05);
        }}

        .brand-header [data-testid="stImage"] img {{
            object-fit: contain;
        }}

        /* Sidebar: blanco, limpio y con una banda sutil de marca. */
        section[data-testid="stSidebar"] {{
            background: var(--brand-white);
            border-right: 1px solid var(--brand-line);
        }}

        section[data-testid="stSidebar"] > div:first-child {{
            padding-top: 1rem;
        }}

        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3,
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] p {{
            color: var(--brand-navy);
        }}

        section[data-testid="stSidebar"] [data-testid="stImage"] {{
            background: var(--brand-white);
            border-bottom: 4px solid var(--brand-gold);
            padding: 0.25rem 0 0.9rem 0;
            margin-bottom: 0.8rem;
        }}

        /* Contenedores y expanders. */
        div[data-testid="stExpander"] {{
            background: var(--brand-white);
            border: 1px solid var(--brand-line);
            border-radius: 0.9rem;
        }}

        /* Botones: principal en azul marino; texto siempre blanco; hover celeste. */
        .stButton > button,
        .stDownloadButton > button,
        button[kind="primary"] {{
            background-color: var(--brand-navy) !important;
            color: var(--brand-white) !important;
            border: 1px solid var(--brand-navy) !important;
            border-radius: 0.55rem !important;
            font-weight: 650 !important;
        }}

        .stButton > button *,
        .stDownloadButton > button *,
        button[kind="primary"] * {{
            color: var(--brand-white) !important;
        }}

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        button[kind="primary"]:hover {{
            background-color: var(--brand-sky) !important;
            border-color: var(--brand-sky) !important;
            color: var(--brand-white) !important;
        }}

        .stButton > button:hover *,
        .stDownloadButton > button:hover *,
        button[kind="primary"]:hover * {{
            color: var(--brand-white) !important;
        }}

        /* Inputs: blancos, bordes suaves y foco celeste. */
        input, textarea, [data-baseweb="select"] > div {{
            background-color: var(--brand-white) !important;
            color: var(--brand-navy) !important;
            border-color: var(--brand-line) !important;
        }}

        input:focus, textarea:focus, [data-baseweb="select"]:focus-within {{
            border-color: var(--brand-sky) !important;
            box-shadow: 0 0 0 1px var(--brand-sky) !important;
        }}

        /* Tabs: blanco dominante, más aire entre pestañas y marcador coral. */
        div[data-baseweb="tab-list"] {{
            gap: 0.85rem !important;
            border-bottom: 1px solid var(--brand-line);
            padding-top: 0.25rem;
            padding-bottom: 0.25rem;
            flex-wrap: wrap;
        }}

        button[data-baseweb="tab"] {{
            color: var(--brand-navy) !important;
            font-weight: 650;
            background: var(--brand-white) !important;
            border-radius: 0.65rem 0.65rem 0 0;
            padding: 0.65rem 0.95rem !important;
            min-height: 2.55rem;
        }}

        button[data-baseweb="tab"] * {{
            color: var(--brand-navy) !important;
        }}

        button[data-baseweb="tab"][aria-selected="true"] {{
            color: var(--brand-coral) !important;
            border-bottom: 4px solid var(--brand-coral) !important;
            background: rgba(233, 92, 71, 0.05) !important;
        }}

        button[data-baseweb="tab"][aria-selected="true"] * {{
            color: var(--brand-coral) !important;
        }}

        /* Métricas: tarjetas blancas, acento dorado. */
        div[data-testid="stMetric"] {{
            background: var(--brand-white);
            border: 1px solid var(--brand-line);
            border-left: 6px solid var(--brand-gold);
            border-radius: 0.8rem;
            padding: 0.8rem 1rem;
            box-shadow: 0 1px 8px rgba(28, 48, 84, 0.04);
        }}

        div[data-testid="stMetric"] label,
        div[data-testid="stMetric"] [data-testid="stMetricValue"] {{
            color: var(--brand-navy) !important;
        }}

        /* Avisos: suaves, sin saturar la interfaz. */
        div[data-testid="stAlert"] {{
            border-radius: 0.75rem;
            border: 1px solid var(--brand-line);
        }}

        /* Enlaces. */
        a {{
            color: var(--brand-sky) !important;
            text-decoration: none;
            font-weight: 600;
        }}

        a:hover {{
            color: var(--brand-coral) !important;
            text-decoration: underline;
        }}

        /* Dataframes: blancos y legibles. */
        div[data-testid="stDataFrame"] {{
            border: 1px solid var(--brand-line);
            border-radius: 0.75rem;
            background: var(--brand-white);
            box-shadow: 0 1px 8px rgba(28, 48, 84, 0.035);
        }}

        /* Separadores visuales discretos. */
        hr {{
            border-color: var(--brand-line);
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

def apply_plotly_brand_layout(fig):
    """Normaliza colores y estilo de gráficos Plotly con la paleta corporativa."""
    fig.update_layout(
        colorway=BRAND_COLORS,
        font=dict(color=COLOR_NAVY),
        paper_bgcolor="white",
        plot_bgcolor="white",
        title_font=dict(color=COLOR_NAVY),
        legend_title_font_color=COLOR_NAVY,
        legend_font_color=COLOR_NAVY,
        margin=dict(l=20, r=20, t=40, b=20),
    )
    fig.update_xaxes(color=COLOR_NAVY, gridcolor="rgba(28,48,84,0.12)")
    fig.update_yaxes(color=COLOR_NAVY, gridcolor="rgba(28,48,84,0.12)")
    return fig


def get_secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, default)
        return str(value) if value is not None else default
    except Exception:
        return os.environ.get(name, default)


def clear_data_caches() -> None:
    """Limpia caches de datos tras guardar cambios para que la interfaz vea datos actualizados."""
    for fn in [
        get_evaluations, get_history, get_action_contacts, get_activity_documents,
        get_contacts, get_access_log, get_questionnaire_responses, get_calendar_events, load_contacts_and_assignments
    ]:
        try:
            fn.clear()
        except Exception:
            pass


def safe_text(value: Any) -> str:
    """Devuelve texto limpio para la interfaz.

    Streamlit / pandas / JavaScript pueden devolver valores técnicos como
    undefined, null o nan. En pantalla no deben mostrarse; se sustituyen por
    vacío para mantener claridad visual.
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text_value = str(value).strip()
    if text_value.lower() in {"undefined", "none", "null", "nan", "nat"}:
        return ""
    return text_value


def clean_interface_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Limpia valores técnicos antes de mostrar tablas o exportarlas."""
    cleaned = df.copy()
    for col in cleaned.columns:
        if cleaned[col].dtype == "object":
            cleaned[col] = cleaned[col].apply(safe_text)
    return cleaned


def normalize_column_name(value: Any) -> str:
    return str(value).replace("\n", " ").strip()


def likert_value(value: Any) -> int | None:
    """Convierte una selección de escala a entero; deja vacío si no se ha marcado."""
    try:
        return int(value)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def read_excel_bytes(file_bytes: bytes) -> pd.DataFrame:
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_name = DEFAULT_SHEET if DEFAULT_SHEET in excel_file.sheet_names else excel_file.sheet_names[0]
    df = pd.read_excel(excel_file, sheet_name=sheet_name).dropna(how="all").copy()
    df.columns = [normalize_column_name(c) for c in df.columns]
    df = df.drop(columns=[c for c in df.columns if c.lower().startswith("unnamed")], errors="ignore")
    for col in BASE_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[BASE_COLUMNS + [c for c in df.columns if c not in BASE_COLUMNS]].copy()
    df.insert(0, "id_accion", range(1, len(df) + 1))
    df["id_accion"] = df["id_accion"].astype(int)
    for col in df.columns:
        if col != "id_accion":
            df[col] = df[col].apply(safe_text)
    return df


@st.cache_data(show_spinner=False)
def load_default_matrix() -> pd.DataFrame:
    return read_excel_bytes(DEFAULT_EXCEL.read_bytes())


@st.cache_data(show_spinner=False)
def read_contacts_bytes(file_bytes: bytes) -> pd.DataFrame:
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_name = "Hoja1" if "Hoja1" in excel_file.sheet_names else excel_file.sheet_names[0]
    df = pd.read_excel(excel_file, sheet_name=sheet_name).dropna(how="all").copy()
    df.columns = [normalize_column_name(c) for c in df.columns]
    df = df.drop(columns=[c for c in df.columns if c.lower().startswith("unnamed")], errors="ignore")
    for col in CONTACT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[CONTACT_COLUMNS + [c for c in df.columns if c not in CONTACT_COLUMNS]].copy()
    df.insert(0, "contacto_id", range(1, len(df) + 1))
    df["contacto_id"] = df["contacto_id"].astype(int)
    for col in df.columns:
        if col != "contacto_id":
            df[col] = df[col].apply(safe_text)
    df["contacto_label"] = df.apply(make_contact_label, axis=1)
    return df


@st.cache_data(show_spinner=False)
def load_default_contacts() -> pd.DataFrame:
    if not CONTACTS_EXCEL.exists():
        return pd.DataFrame(columns=["contacto_id", *CONTACT_COLUMNS, "contacto_label"])
    return read_contacts_bytes(CONTACTS_EXCEL.read_bytes())


def make_contact_label(row: pd.Series) -> str:
    persona = safe_text(row.get("Persona")) or "Sin persona"
    entidad = safe_text(row.get("Entidad"))
    comision = safe_text(row.get("Comisión"))
    perfil = safe_text(row.get("Perfil"))
    parts = [persona]
    if entidad:
        parts.append(entidad)
    if comision:
        parts.append(comision)
    if perfil:
        parts.append(perfil)
    return " | ".join(parts)


@st.cache_resource(show_spinner=False)
def get_engine() -> Engine:
    database_url = get_secret("DATABASE_URL", "").strip() or LOCAL_SQLITE
    connect_args = {"check_same_thread": False} if database_url.startswith("sqlite") else {}
    return create_engine(database_url, connect_args=connect_args, pool_pre_ping=True)


def is_postgres(engine: Engine) -> bool:
    return engine.dialect.name.startswith("postgres")


def column_exists(engine: Engine, table_name: str, column_name: str) -> bool:
    with engine.begin() as conn:
        if is_postgres(engine):
            result = conn.execute(
                text("""
                SELECT 1 FROM information_schema.columns
                WHERE table_name = :table_name AND column_name = :column_name
                """),
                {"table_name": table_name, "column_name": column_name},
            ).first()
            return result is not None
        rows = conn.execute(text(f"PRAGMA table_info({table_name})")).fetchall()
        return any(row[1] == column_name for row in rows)


@st.cache_resource(show_spinner=False)
def init_db(_engine: Engine) -> None:
    engine = _engine
    history_id = "BIGSERIAL PRIMARY KEY" if is_postgres(engine) else "INTEGER PRIMARY KEY AUTOINCREMENT"
    with engine.begin() as conn:
        conn.execute(text(
            """
            CREATE TABLE IF NOT EXISTS evaluations (
                id_accion INTEGER PRIMARY KEY,
                avance INTEGER NOT NULL DEFAULT 0,
                estado_evaluacion TEXT NOT NULL DEFAULT 'Sin evaluar',
                cumplimiento_indicadores TEXT NOT NULL DEFAULT 'Sin evaluar',
                valoracion_tecnica TEXT,
                observaciones TEXT,
                responsable_seguimiento TEXT,
                fecha_actualizacion TEXT,
                evidencias TEXT,
                riesgos TEXT,
                proximos_pasos TEXT,
                prioridad TEXT,
                link_promotora TEXT,
                link_promotora_enviado INTEGER NOT NULL DEFAULT 0,
                link_promotora_fecha TEXT,
                link_participante TEXT,
                link_participante_enviado INTEGER NOT NULL DEFAULT 0,
                link_participante_fecha TEXT,
                updated_by TEXT,
                updated_at TEXT NOT NULL
            )
            """
        ))
        conn.execute(text(
            f"""
            CREATE TABLE IF NOT EXISTS evaluation_history (
                id {history_id},
                id_accion INTEGER NOT NULL,
                avance INTEGER NOT NULL DEFAULT 0,
                estado_evaluacion TEXT NOT NULL DEFAULT 'Sin evaluar',
                cumplimiento_indicadores TEXT NOT NULL DEFAULT 'Sin evaluar',
                valoracion_tecnica TEXT,
                observaciones TEXT,
                responsable_seguimiento TEXT,
                fecha_actualizacion TEXT,
                evidencias TEXT,
                riesgos TEXT,
                proximos_pasos TEXT,
                prioridad TEXT,
                link_promotora TEXT,
                link_promotora_enviado INTEGER NOT NULL DEFAULT 0,
                link_promotora_fecha TEXT,
                link_participante TEXT,
                link_participante_enviado INTEGER NOT NULL DEFAULT 0,
                link_participante_fecha TEXT,
                updated_by TEXT,
                updated_at TEXT NOT NULL
            )
            """
        ))
        conn.execute(text(
            """
            CREATE TABLE IF NOT EXISTS action_contacts (
                id_accion INTEGER NOT NULL,
                contacto_id INTEGER NOT NULL,
                assigned_by TEXT,
                assigned_at TEXT NOT NULL,
                PRIMARY KEY (id_accion, contacto_id)
            )
            """
        ))
        contact_id_type = "SERIAL PRIMARY KEY" if is_postgres(engine) else "INTEGER PRIMARY KEY AUTOINCREMENT"
        conn.execute(text(
            f"""
            CREATE TABLE IF NOT EXISTS contacts (
                contacto_id {contact_id_type},
                Categoria TEXT,
                Subcategoria TEXT,
                Comisión TEXT,
                Perfil TEXT,
                Entidad TEXT,
                Teléfono TEXT,
                Persona TEXT,
                Mail TEXT,
                updated_by TEXT,
                updated_at TEXT
            )
            """
        ))
        document_id_type = "BIGSERIAL PRIMARY KEY" if is_postgres(engine) else "INTEGER PRIMARY KEY AUTOINCREMENT"
        conn.execute(text(
            f"""
            CREATE TABLE IF NOT EXISTS activity_documents (
                document_id {document_id_type},
                id_accion INTEGER NOT NULL,
                filename TEXT NOT NULL,
                mime_type TEXT,
                size_bytes INTEGER,
                content_base64 TEXT NOT NULL,
                notes TEXT,
                uploaded_by TEXT,
                uploaded_at TEXT NOT NULL
            )
            """
        ))
        access_id_type = "BIGSERIAL PRIMARY KEY" if is_postgres(engine) else "INTEGER PRIMARY KEY AUTOINCREMENT"
        conn.execute(text(
            f"""
            CREATE TABLE IF NOT EXISTS access_log (
                id {access_id_type},
                perfil TEXT NOT NULL,
                role TEXT NOT NULL,
                scope TEXT,
                logged_at TEXT NOT NULL
            )
            """
        ))
        task_id_type = "BIGSERIAL PRIMARY KEY" if is_postgres(engine) else "INTEGER PRIMARY KEY AUTOINCREMENT"
        conn.execute(text(
            f"""
            CREATE TABLE IF NOT EXISTS manual_tasks (
                task_id {task_id_type},
                id_accion INTEGER,
                accion_titulo TEXT,
                proximos_pasos TEXT,
                descripcion TEXT,
                responsable_seguimiento TEXT,
                fecha_peticion TEXT,
                fecha_vencimiento TEXT,
                estado TEXT NOT NULL DEFAULT 'No iniciado',
                created_by TEXT,
                created_at TEXT NOT NULL,
                updated_by TEXT,
                updated_at TEXT NOT NULL
            )
            """
        ))
        conn.execute(text(
            """
            CREATE TABLE IF NOT EXISTS task_overrides (
                task_key TEXT PRIMARY KEY,
                descripcion TEXT,
                fecha_vencimiento TEXT,
                estado TEXT NOT NULL DEFAULT 'No iniciado',
                updated_by TEXT,
                updated_at TEXT NOT NULL
            )
            """
        ))
        conn.execute(text(
            """
            CREATE TABLE IF NOT EXISTS task_email_reminders (
                reminder_key TEXT PRIMARY KEY,
                task_source TEXT NOT NULL,
                task_identifier TEXT NOT NULL,
                id_accion INTEGER,
                accion_titulo TEXT,
                proximos_pasos TEXT,
                responsable_seguimiento TEXT,
                fecha_vencimiento TEXT,
                recipient TEXT NOT NULL,
                sent_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'sent',
                error_message TEXT
            )
            """
        ))
        event_id_type = "BIGSERIAL PRIMARY KEY" if is_postgres(engine) else "INTEGER PRIMARY KEY AUTOINCREMENT"
        conn.execute(text(
            f"""
            CREATE TABLE IF NOT EXISTS calendar_events (
                event_id {event_id_type},
                title TEXT NOT NULL,
                event_type TEXT NOT NULL DEFAULT 'Evento',
                description TEXT,
                location TEXT,
                start_date TEXT NOT NULL,
                start_time TEXT,
                end_date TEXT,
                end_time TEXT,
                id_accion INTEGER,
                accion_titulo TEXT,
                responsable TEXT,
                created_by TEXT,
                created_at TEXT NOT NULL,
                updated_by TEXT,
                updated_at TEXT NOT NULL
            )
            """
        ))
        response_id_type = "BIGSERIAL PRIMARY KEY" if is_postgres(engine) else "INTEGER PRIMARY KEY AUTOINCREMENT"
        conn.execute(text(
            f"""
            CREATE TABLE IF NOT EXISTS questionnaire_responses (
                response_id {response_id_type},
                id_accion INTEGER NOT NULL,
                questionnaire_type TEXT NOT NULL,
                action_title TEXT,
                respondent_name TEXT,
                respondent_entity TEXT,
                respondent_email TEXT,
                rating_general INTEGER,
                rating_usefulness INTEGER,
                rating_clarity INTEGER,
                objectives_met TEXT,
                coordination TEXT,
                participation TEXT,
                learning TEXT,
                positives TEXT,
                difficulties TEXT,
                improvements TEXT,
                would_repeat TEXT,
                would_recommend TEXT,
                comments TEXT,
                submitted_at TEXT NOT NULL
            )
            """
        ))
    evaluation_extra_columns = {
        "updated_by": "TEXT",
        "link_promotora": "TEXT",
        "link_promotora_enviado": "INTEGER NOT NULL DEFAULT 0",
        "link_promotora_fecha": "TEXT",
        "link_participante": "TEXT",
        "link_participante_enviado": "INTEGER NOT NULL DEFAULT 0",
        "link_participante_fecha": "TEXT",
    }
    for table_name in ["evaluations", "evaluation_history"]:
        for column_name, column_type in evaluation_extra_columns.items():
            if not column_exists(engine, table_name, column_name):
                with engine.begin() as conn:
                    conn.execute(text(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}"))


@st.cache_data(ttl=300, show_spinner=False)
def get_evaluations(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    with engine.begin() as conn:
        df = pd.read_sql_query(text("SELECT * FROM evaluations"), conn)
    if df.empty:
        return pd.DataFrame(columns=[
            "id_accion", "avance", "estado_evaluacion", "cumplimiento_indicadores", "valoracion_tecnica",
            "observaciones", "responsable_seguimiento", "fecha_actualizacion", "evidencias", "riesgos",
            "proximos_pasos", "prioridad", "link_promotora", "link_promotora_enviado", "link_promotora_fecha",
            "link_participante", "link_participante_enviado", "link_participante_fecha", "updated_by", "updated_at",
        ])
    df["id_accion"] = df["id_accion"].astype(int)
    return df


@st.cache_data(ttl=300, show_spinner=False)
def get_history(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    with engine.begin() as conn:
        df = pd.read_sql_query(text("SELECT * FROM evaluation_history ORDER BY updated_at ASC, id ASC"), conn)
    if df.empty:
        return pd.DataFrame(columns=[
            "id", "id_accion", "avance", "estado_evaluacion", "cumplimiento_indicadores", "valoracion_tecnica",
            "observaciones", "responsable_seguimiento", "fecha_actualizacion", "evidencias", "riesgos",
            "proximos_pasos", "prioridad", "updated_by", "updated_at",
        ])
    df["id_accion"] = df["id_accion"].astype(int)
    df["avance"] = df["avance"].fillna(0).astype(int)
    df["fecha_actualizacion_dt"] = pd.to_datetime(df["fecha_actualizacion"], errors="coerce")
    df["updated_at_dt"] = pd.to_datetime(df["updated_at"], errors="coerce")
    return df


@st.cache_data(ttl=300, show_spinner=False)
def get_action_contacts(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    with engine.begin() as conn:
        df = pd.read_sql_query(text("SELECT * FROM action_contacts"), conn)
    if df.empty:
        return pd.DataFrame(columns=["id_accion", "contacto_id", "assigned_by", "assigned_at"])
    df["id_accion"] = df["id_accion"].astype(int)
    df["contacto_id"] = df["contacto_id"].astype(int)
    return df


def save_action_contacts(engine: Engine, id_accion: int, contacto_ids: list[int], user_name: str) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM action_contacts WHERE id_accion = :id_accion"), {"id_accion": id_accion})
        for contacto_id in sorted(set(int(x) for x in contacto_ids)):
            conn.execute(
                text("""
                    INSERT INTO action_contacts (id_accion, contacto_id, assigned_by, assigned_at)
                    VALUES (:id_accion, :contacto_id, :assigned_by, :assigned_at)
                """),
                {"id_accion": id_accion, "contacto_id": contacto_id, "assigned_by": user_name, "assigned_at": now},
            )
    clear_data_caches()




@st.cache_data(ttl=300, show_spinner=False)
def get_activity_documents(_engine: Engine, id_accion: int | None = None) -> pd.DataFrame:
    engine = _engine
    query = "SELECT * FROM activity_documents"
    params: dict[str, Any] = {}
    if id_accion is not None:
        query += " WHERE id_accion = :id_accion"
        params["id_accion"] = int(id_accion)
    query += " ORDER BY uploaded_at DESC, document_id DESC"
    with engine.begin() as conn:
        df = pd.read_sql_query(text(query), conn, params=params)
    if df.empty:
        return pd.DataFrame(columns=[
            "document_id", "id_accion", "filename", "mime_type", "size_bytes", "content_base64",
            "notes", "uploaded_by", "uploaded_at",
        ])
    df["document_id"] = df["document_id"].astype(int)
    df["id_accion"] = df["id_accion"].astype(int)
    return df


def save_activity_documents(engine: Engine, id_accion: int, uploaded_files: list[Any], notes: str, user_name: str) -> int:
    now = datetime.now().isoformat(timespec="seconds")
    inserted = 0
    with engine.begin() as conn:
        for uploaded_file in uploaded_files:
            raw = uploaded_file.getvalue()
            conn.execute(
                text("""
                    INSERT INTO activity_documents (
                        id_accion, filename, mime_type, size_bytes, content_base64, notes, uploaded_by, uploaded_at
                    ) VALUES (
                        :id_accion, :filename, :mime_type, :size_bytes, :content_base64, :notes, :uploaded_by, :uploaded_at
                    )
                """),
                {
                    "id_accion": int(id_accion),
                    "filename": uploaded_file.name,
                    "mime_type": getattr(uploaded_file, "type", "") or "application/octet-stream",
                    "size_bytes": len(raw),
                    "content_base64": base64.b64encode(raw).decode("ascii"),
                    "notes": notes,
                    "uploaded_by": user_name,
                    "uploaded_at": now,
                },
            )
            inserted += 1
    clear_data_caches()
    return inserted


def delete_activity_document(engine: Engine, document_id: int) -> None:
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM activity_documents WHERE document_id = :document_id"), {"document_id": int(document_id)})
    clear_data_caches()



@st.cache_data(ttl=300, show_spinner=False)
def get_questionnaire_responses(_engine: Engine, id_accion: int | None = None) -> pd.DataFrame:
    engine = _engine
    query = "SELECT * FROM questionnaire_responses"
    params: dict[str, Any] = {}
    if id_accion is not None:
        query += " WHERE id_accion = :id_accion"
        params["id_accion"] = int(id_accion)
    query += " ORDER BY submitted_at DESC, response_id DESC"
    try:
        with engine.begin() as conn:
            df = pd.read_sql_query(text(query), conn, params=params)
    except Exception:
        return pd.DataFrame(columns=[
            "response_id", "id_accion", "questionnaire_type", "action_title", "respondent_name",
            "respondent_entity", "respondent_email", "rating_general", "rating_usefulness",
            "rating_clarity", "objectives_met", "coordination", "participation", "learning",
            "positives", "difficulties", "improvements", "would_repeat", "would_recommend",
            "comments", "submitted_at",
        ])
    return clean_interface_dataframe(df)


def save_questionnaire_response(engine: Engine, payload: dict[str, Any]) -> None:
    data = {**payload, "submitted_at": datetime.now().isoformat(timespec="seconds")}
    fields = list(data.keys())
    with engine.begin() as conn:
        conn.execute(
            text(f"""
                INSERT INTO questionnaire_responses ({", ".join(fields)})
                VALUES ({", ".join(":" + field for field in fields)})
            """),
            data,
        )
    clear_data_caches()


def get_app_base_url() -> str:
    return get_secret("APP_BASE_URL", "").strip().rstrip("/")


def build_questionnaire_link(id_accion: int, questionnaire_type: str) -> str:
    query = f"?eval={questionnaire_type}&id_accion={int(id_accion)}"
    base_url = get_app_base_url()
    return f"{base_url}{query}" if base_url else query


def get_query_param(name: str, default: str = "") -> str:
    try:
        value = st.query_params.get(name, default)
        if isinstance(value, list):
            return str(value[0]) if value else default
        return str(value) if value is not None else default
    except Exception:
        return default


def render_public_questionnaire(engine: Engine) -> bool:
    questionnaire_type = get_query_param("eval", "").strip().lower()
    id_value = get_query_param("id_accion", "").strip()
    if questionnaire_type not in QUESTIONNAIRE_TYPES or not id_value:
        return False
    try:
        id_accion = int(id_value)
    except ValueError:
        st.error("El enlace de evaluación no es válido.")
        return True

    init_db(engine)
    matrix = load_default_matrix()
    action_rows = matrix[matrix["id_accion"] == id_accion]
    if action_rows.empty:
        st.error("No se ha encontrado la actividad asociada a este enlace.")
        return True

    action = action_rows.iloc[0]
    title = safe_text(action.get("Título"))
    scope = safe_text(action.get("Ámbito"))
    q_label = QUESTIONNAIRE_TYPES[questionnaire_type]

    render_login_title()
    st.markdown(f"### Cuestionario de evaluación: {q_label}")
    st.info(f"Actividad: {title}\n\nÁmbito: {scope}")

    if st.session_state.get(f"questionnaire_submitted_{questionnaire_type}_{id_accion}"):
        st.success("Respuesta registrada. Gracias por completar la evaluación.")
        return True

    with st.form(f"public_questionnaire_{questionnaire_type}_{id_accion}"):
        if questionnaire_type == "promotora":
            respondent_entity = st.text_input("Entidad o persona promotora")
            respondent_name = st.text_input("Persona de contacto")
            respondent_email = st.text_input("Email de contacto, opcional")
            rating_general = st.selectbox("Valoración general de la actividad", LIKERT_SELECT_OPTIONS, index=0, help="1 = muy baja · 5 = muy alta")
            rating_usefulness = st.selectbox("Utilidad de la actividad", LIKERT_SELECT_OPTIONS, index=0, help="1 = muy baja · 5 = muy alta")
            rating_clarity = st.selectbox("Claridad de la información recibida", LIKERT_SELECT_OPTIONS, index=0, help="1 = muy baja · 5 = muy alta")
            objectives_met = st.selectbox("¿Se han cumplido los objetivos previstos?", YES_NO_OPTIONS)
            coordination = st.selectbox("Coordinación con la Red Local de Salud", LIKERT_SELECT_OPTIONS, index=0)
            participation = st.text_area("Participación conseguida", placeholder="Número aproximado, perfil de participantes o valoración de la participación")
            positives = st.text_area("Aspectos positivos")
            difficulties = st.text_area("Dificultades encontradas")
            improvements = st.text_area("Propuestas de mejora")
            would_repeat = st.selectbox("¿Repetirías o recomendarías impulsar esta actividad de nuevo?", YES_NO_OPTIONS)
            comments = st.text_area("Comentarios finales")
            payload = {
                "id_accion": id_accion,
                "questionnaire_type": questionnaire_type,
                "action_title": title,
                "respondent_name": respondent_name,
                "respondent_entity": respondent_entity,
                "respondent_email": respondent_email,
                "rating_general": likert_value(rating_general),
                "rating_usefulness": likert_value(rating_usefulness),
                "rating_clarity": likert_value(rating_clarity),
                "objectives_met": objectives_met,
                "coordination": "" if coordination == "Sin marcar" else str(coordination),
                "participation": participation,
                "learning": "",
                "positives": positives,
                "difficulties": difficulties,
                "improvements": improvements,
                "would_repeat": would_repeat,
                "would_recommend": "",
                "comments": comments,
            }
        else:
            respondent_name = st.text_input("Nombre, opcional")
            respondent_entity = st.text_input("Entidad, centro o grupo, opcional")
            respondent_email = st.text_input("Email, opcional")
            rating_general = st.selectbox("Valoración general de la actividad", LIKERT_SELECT_OPTIONS, index=0, help="1 = muy baja · 5 = muy alta")
            rating_usefulness = st.selectbox("Utilidad de la actividad", LIKERT_SELECT_OPTIONS, index=0, help="1 = muy baja · 5 = muy alta")
            rating_clarity = st.selectbox("Claridad de la información recibida", LIKERT_SELECT_OPTIONS, index=0, help="1 = muy baja · 5 = muy alta")
            learning = st.text_area("¿Qué te llevas o qué has aprendido?")
            improvements = st.text_area("¿Qué mejorarías?")
            would_recommend = st.selectbox("¿Recomendarías esta actividad?", YES_NO_OPTIONS)
            comments = st.text_area("Comentarios finales")
            payload = {
                "id_accion": id_accion,
                "questionnaire_type": questionnaire_type,
                "action_title": title,
                "respondent_name": respondent_name,
                "respondent_entity": respondent_entity,
                "respondent_email": respondent_email,
                "rating_general": likert_value(rating_general),
                "rating_usefulness": likert_value(rating_usefulness),
                "rating_clarity": likert_value(rating_clarity),
                "objectives_met": "",
                "coordination": "",
                "participation": "",
                "learning": learning,
                "positives": "",
                "difficulties": "",
                "improvements": improvements,
                "would_repeat": "",
                "would_recommend": would_recommend,
                "comments": comments,
            }
        submitted = st.form_submit_button("Enviar evaluación")

    if submitted:
        save_questionnaire_response(engine, payload)
        st.session_state[f"questionnaire_submitted_{questionnaire_type}_{id_accion}"] = True
        st.success("Respuesta registrada. Gracias por completar la evaluación.")
        st.rerun()
    return True


def render_questionnaire_results(engine: Engine, id_accion: int | None = None) -> None:
    responses = get_questionnaire_responses(engine, id_accion)
    st.markdown("### Respuestas de cuestionarios")
    if responses.empty:
        st.info("Todavía no hay respuestas de cuestionarios.")
        return
    responses = clean_interface_dataframe(responses)
    label_map = {"promotora": "Persona Promotora", "participante": "Persona participante"}
    responses["Tipo de cuestionario"] = responses["questionnaire_type"].map(label_map).fillna(responses["questionnaire_type"])
    summary = responses.groupby("Tipo de cuestionario", dropna=False).size().reset_index(name="Respuestas")
    st.dataframe(summary, use_container_width=True, hide_index=True)
    visible = [
        "submitted_at", "Tipo de cuestionario", "action_title", "respondent_entity", "respondent_name",
        "rating_general", "rating_usefulness", "rating_clarity", "objectives_met", "would_repeat",
        "would_recommend", "comments",
    ]
    visible = [col for col in visible if col in responses.columns]
    st.dataframe(responses[visible], use_container_width=True, hide_index=True)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        responses.to_excel(writer, index=False, sheet_name="Cuestionarios")
    st.download_button(
        "Descargar respuestas de cuestionarios",
        buffer.getvalue(),
        "respuestas_cuestionarios.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_questionnaire_responses_{id_accion or 'all'}",
    )


def state_to_legacy_progress(state: str) -> int:
    mapping = {"Sin evaluar": 0, "Iniciado": 50, "Completado": 100, "No procede": 0}
    return mapping.get(state, 0)


def normalize_contact_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza columnas de contactos procedentes de PostgreSQL/SQLite.

    En PostgreSQL los nombres no entrecomillados se devuelven en minúsculas
    (por ejemplo, Categoria -> categoria, Teléfono -> teléfono). Esta función
    los vuelve a poner con el nombre visible que usa la aplicación.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["contacto_id", *CONTACT_COLUMNS, "updated_by", "updated_at", "contacto_label"])

    def key(value: Any) -> str:
        text_value = str(value).strip().lower()
        return (
            text_value
            .replace("á", "a")
            .replace("é", "e")
            .replace("í", "i")
            .replace("ó", "o")
            .replace("ú", "u")
            .replace("ü", "u")
            .replace("ñ", "n")
        )

    aliases = {
        "contacto_id": "contacto_id",
        "id": "contacto_id",
        "categoria": "Categoria",
        "subcategoria": "Subcategoria",
        "comision": "Comisión",
        "perfil": "Perfil",
        "entidad": "Entidad",
        "telefono": "Teléfono",
        "persona": "Persona",
        "mail": "Mail",
        "email": "Mail",
        "correo": "Mail",
        "updated_by": "updated_by",
        "updated_at": "updated_at",
    }

    renamed = {}
    for col in df.columns:
        mapped = aliases.get(key(col))
        if mapped and mapped not in renamed.values():
            renamed[col] = mapped
    df = df.rename(columns=renamed).copy()

    for col in ["contacto_id", *CONTACT_COLUMNS, "updated_by", "updated_at"]:
        if col not in df.columns:
            df[col] = "" if col != "contacto_id" else range(1, len(df) + 1)

    try:
        df["contacto_id"] = pd.to_numeric(df["contacto_id"], errors="coerce").fillna(0).astype(int)
    except Exception:
        df["contacto_id"] = range(1, len(df) + 1)

    for col in CONTACT_COLUMNS:
        df[col] = df[col].apply(safe_text)
    df["updated_by"] = df["updated_by"].apply(safe_text)
    df["updated_at"] = df["updated_at"].apply(safe_text)
    df["contacto_label"] = df.apply(make_contact_label, axis=1)
    return df[["contacto_id", *CONTACT_COLUMNS, "updated_by", "updated_at", "contacto_label"]]


def seed_contacts_if_empty(engine: Engine) -> None:
    """Carga Contactos.xlsx si la tabla está vacía o solo contiene registros sin datos visibles."""
    if not CONTACTS_EXCEL.exists():
        return
    try:
        current = get_contacts(engine)
        meaningful_rows = current[CONTACT_COLUMNS].fillna("").astype(str).apply(
            lambda row: any(cell.strip() for cell in row), axis=1
        ).sum() if not current.empty else 0
    except Exception:
        meaningful_rows = 0
    if meaningful_rows > 0:
        return
    seed = load_default_contacts()
    if seed.empty:
        return
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM action_contacts"))
        conn.execute(text("DELETE FROM contacts"))
        for _, row in seed.iterrows():
            payload = {col: safe_text(row.get(col)) for col in CONTACT_COLUMNS}
            if not any(payload.values()):
                continue
            payload.update({"updated_by": "carga inicial", "updated_at": now})
            fields = list(payload.keys())
            conn.execute(
                text(f"""
                    INSERT INTO contacts ({", ".join(fields)})
                    VALUES ({", ".join([":" + f for f in fields])})
                """),
                payload,
            )


def import_contacts_dataframe(engine: Engine, contacts_df: pd.DataFrame, user_name: str, replace: bool = False) -> int:
    """Importa contactos desde un DataFrame normalizado. Si replace=True, sustituye el listado completo."""
    if contacts_df.empty:
        return 0
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as conn:
        if replace:
            conn.execute(text("DELETE FROM action_contacts"))
            conn.execute(text("DELETE FROM contacts"))
        inserted = 0
        for _, row in contacts_df.iterrows():
            payload = {col: safe_text(row.get(col)) for col in CONTACT_COLUMNS}
            if not any(payload.values()):
                continue
            payload.update({"updated_by": user_name or "carga desde panel", "updated_at": now})
            fields = list(payload.keys())
            conn.execute(
                text(f"""
                    INSERT INTO contacts ({", ".join(fields)})
                    VALUES ({", ".join([":" + f for f in fields])})
                """),
                payload,
            )
            inserted += 1
    clear_data_caches()
    return inserted


@st.cache_data(ttl=300, show_spinner=False)
def get_contacts(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    with engine.begin() as conn:
        raw = pd.read_sql_query(text("SELECT * FROM contacts ORDER BY contacto_id ASC"), conn)
    return normalize_contact_dataframe(raw)


def add_contact(engine: Engine, data: dict[str, Any], user_name: str) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    payload = {col: safe_text(data.get(col)) for col in CONTACT_COLUMNS}
    payload.update({"updated_by": user_name, "updated_at": now})
    fields = list(payload.keys())
    with engine.begin() as conn:
        conn.execute(
            text(f"""
                INSERT INTO contacts ({", ".join(fields)})
                VALUES ({", ".join([":" + f for f in fields])})
            """),
            payload,
        )
    clear_data_caches()


def update_contacts(engine: Engine, contacts_df: pd.DataFrame, user_name: str) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as conn:
        for _, row in contacts_df.iterrows():
            contacto_id = int(row["contacto_id"])
            payload = {col: safe_text(row.get(col)) for col in CONTACT_COLUMNS}
            payload.update({"updated_by": user_name, "updated_at": now, "contacto_id": contacto_id})
            conn.execute(
                text("""
                    UPDATE contacts
                    SET Categoria = :Categoria,
                        Subcategoria = :Subcategoria,
                        Comisión = :Comisión,
                        Perfil = :Perfil,
                        Entidad = :Entidad,
                        Teléfono = :Teléfono,
                        Persona = :Persona,
                        Mail = :Mail,
                        updated_by = :updated_by,
                        updated_at = :updated_at
                    WHERE contacto_id = :contacto_id
                """),
                payload,
            )
    clear_data_caches()


def delete_contact(engine: Engine, contacto_id: int) -> None:
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM action_contacts WHERE contacto_id = :contacto_id"), {"contacto_id": contacto_id})
        conn.execute(text("DELETE FROM contacts WHERE contacto_id = :contacto_id"), {"contacto_id": contacto_id})
    clear_data_caches()


def build_assignment_summary(assignments: pd.DataFrame, contacts: pd.DataFrame) -> pd.DataFrame:
    if assignments.empty or contacts.empty:
        return pd.DataFrame(columns=["id_accion", "personas_red_asignadas", "num_personas_red"])
    merged = assignments.merge(
        contacts[["contacto_id", "Persona", "Entidad", "Comisión", "Perfil", "contacto_label"]],
        on="contacto_id",
        how="left",
    )
    summary = merged.groupby("id_accion").agg(
        personas_red_asignadas=("contacto_label", lambda values: "; ".join([safe_text(v) for v in values if safe_text(v)])),
        num_personas_red=("contacto_id", "nunique"),
    ).reset_index()
    summary["id_accion"] = summary["id_accion"].astype(int)
    return summary


def merge_assignments(data: pd.DataFrame, assignments: pd.DataFrame, contacts: pd.DataFrame) -> pd.DataFrame:
    summary = build_assignment_summary(assignments, contacts)
    if summary.empty:
        data = data.copy()
        data["personas_red_asignadas"] = ""
        data["num_personas_red"] = 0
        return data
    merged = data.merge(summary, on="id_accion", how="left")
    merged["personas_red_asignadas"] = merged["personas_red_asignadas"].fillna("")
    merged["num_personas_red"] = merged["num_personas_red"].fillna(0).astype(int)
    return merged


def save_evaluation(engine: Engine, data: dict[str, Any]) -> None:
    payload = {**data, "updated_at": datetime.now().isoformat(timespec="seconds")}
    fields = list(payload.keys())
    insert_fields = ", ".join(fields)
    placeholders = ", ".join([":" + f for f in fields])
    updates = ", ".join([f"{f}=excluded.{f}" for f in fields if f != "id_accion"])
    sql_current = text(f"""
        INSERT INTO evaluations ({insert_fields})
        VALUES ({placeholders})
        ON CONFLICT(id_accion) DO UPDATE SET {updates}
    """)
    sql_history = text(f"""
        INSERT INTO evaluation_history ({insert_fields})
        VALUES ({placeholders})
    """)
    with engine.begin() as conn:
        conn.execute(sql_current, payload)
        conn.execute(sql_history, payload)
    clear_data_caches()


def merge_matrix_evaluations(matrix: pd.DataFrame, evaluations: pd.DataFrame) -> pd.DataFrame:
    if evaluations.empty:
        merged = matrix.copy()
        defaults = {
            "avance": 0, "estado_evaluacion": "Sin evaluar", "cumplimiento_indicadores": "Sin evaluar",
            "valoracion_tecnica": "", "observaciones": "", "responsable_seguimiento": "",
            "fecha_actualizacion": "", "evidencias": "", "riesgos": "", "proximos_pasos": "",
            "prioridad": "Media", "link_promotora": "", "link_promotora_enviado": 0, "link_promotora_fecha": "",
            "link_participante": "", "link_participante_enviado": 0, "link_participante_fecha": "",
            "updated_by": "", "updated_at": "",
        }
        for col, value in defaults.items():
            merged[col] = value
        return merged
    return matrix.merge(evaluations, on="id_accion", how="left")



def render_home_image() -> None:
    """Imagen de portada en la pantalla de acceso."""
    if HOME_IMAGE_PATH.exists():
        left, center, right = st.columns([1, 3, 1])
        with center:
            st.image(str(HOME_IMAGE_PATH), use_container_width=True)


def render_login_title() -> None:
    render_home_image()
    st.markdown('<div class="brand-header">', unsafe_allow_html=True)
    st.title(get_secret("APP_TITLE", "Erandio Mugitzen ari da! - Evaluación y seguimiento"))
    st.caption("Acceso restringido a la herramienta compartida de evaluación y seguimiento")
    st.markdown('</div>', unsafe_allow_html=True)


def render_internal_header() -> None:
    """Cabecera interna con logo proporcionado y título."""
    st.markdown('<div class="brand-header">', unsafe_allow_html=True)
    if LOGO_PATH.exists():
        col_logo, col_title = st.columns([1, 7], vertical_alignment="center")
        with col_logo:
            st.image(str(LOGO_PATH), width=96)
        with col_title:
            st.title(get_secret("APP_TITLE", "Evaluacion y seguimiento Emad!"))
            st.caption("Herramienta compartida de evaluación, seguimiento histórico y actualización de acciones")
    else:
        st.title(get_secret("APP_TITLE", "Evaluacion y seguimiento Emad!"))
        st.caption("Herramienta compartida de evaluación, seguimiento histórico y actualización de acciones")
    st.markdown('</div>', unsafe_allow_html=True)


def render_sidebar_logo() -> None:
    """Logo en la parte superior izquierda, integrado con el menú lateral."""
    if LOGO_PATH.exists():
        st.sidebar.image(str(LOGO_PATH), width=145)


def log_access(engine: Engine, perfil: str, role: str, scope: str | None) -> None:
    try:
        with engine.begin() as conn:
            conn.execute(
                text("""
                INSERT INTO access_log (perfil, role, scope, logged_at)
                VALUES (:perfil, :role, :scope, :logged_at)
                """),
                {
                    "perfil": perfil,
                    "role": role,
                    "scope": scope or "",
                    "logged_at": datetime.now().isoformat(timespec="seconds"),
                },
            )
    except Exception:
        # El acceso no debe bloquearse si falla el registro técnico.
        pass


@st.cache_data(ttl=300, show_spinner=False)
def get_access_log(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    try:
        with engine.begin() as conn:
            return pd.read_sql_query(text("SELECT * FROM access_log ORDER BY logged_at DESC"), conn)
    except Exception:
        return pd.DataFrame(columns=["id", "perfil", "role", "scope", "logged_at"])


def get_profile_password(profile_name: str, config: dict[str, Any]) -> str:
    # Para administración se mantiene Temporal001*2026 por defecto.
    # Si existe APP_PASSWORD en Secrets, también se acepta como compatibilidad con la versión anterior.
    password = get_secret(config["secret"], config["default_password"]).strip()
    if config.get("role") == "admin":
        return password or get_secret("APP_PASSWORD", "Temporal001*2026").strip()
    return password


def authenticate(engine: Engine) -> bool:
    if st.session_state.get("authenticated"):
        return True
    render_login_title()
    st.markdown("### Selecciona perfil de acceso")
    profile_name = st.selectbox("Perfil", list(PROFILE_CONFIG.keys()))
    password = st.text_input("Contraseña", type="password")
    if st.button("Entrar"):
        config = PROFILE_CONFIG[profile_name]
        expected_password = get_profile_password(profile_name, config)
        admin_legacy_password = get_secret("APP_PASSWORD", "").strip() if config.get("role") == "admin" else ""
        valid_passwords = {expected_password}
        if admin_legacy_password:
            valid_passwords.add(admin_legacy_password)
        if password in valid_passwords:
            st.session_state["authenticated"] = True
            st.session_state["profile_name"] = profile_name
            st.session_state["role"] = config["role"]
            st.session_state["scope"] = config.get("scope")
            st.session_state["user_name"] = profile_name
            init_db(engine)
            log_access(engine, profile_name, config["role"], config.get("scope"))
            st.rerun()
        else:
            st.error("Contraseña incorrecta para este perfil.")
    return False


def sidebar_user() -> str:
    with st.sidebar:
        st.header("Sesión")
        profile_name = st.session_state.get("profile_name", "Sin perfil")
        role = st.session_state.get("role", "")
        scope = st.session_state.get("scope", "")
        st.write(f"**Perfil:** {profile_name}")
        if role == "admin":
            st.caption("Acceso de administración: todas las actividades, evolución y administración técnica.")
        else:
            st.caption("Perfil de eje: puede ver todas las actividades; las modificaciones quedan registradas con este perfil.")
        if st.button("Cerrar sesión"):
            st.session_state.clear()
            st.rerun()
    return str(st.session_state.get("user_name", profile_name)).strip()


def filter_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    with st.sidebar:
        st.header("Filtros")

        def multi_filter(column: str, label: str) -> list[str]:
            values = sorted([v for v in df[column].dropna().unique().tolist() if str(v).strip()])
            return st.multiselect(label, values, default=values)

        ambitos = multi_filter("Ámbito", "Ámbito")
        estados = multi_filter("Estado", "Estado original")
        tipos = multi_filter("Tipo", "Tipo")
        promotores = multi_filter("Agente promotor", "Agente promotor")
        estado_eval = st.multiselect("Estado de evaluación", STATUS_OPTIONS, default=STATUS_OPTIONS)
        texto = st.text_input("Buscar texto", placeholder="Título, descripción, indicadores...")

    filtered = df[
        df["Ámbito"].isin(ambitos)
        & df["Estado"].isin(estados)
        & df["Tipo"].isin(tipos)
        & df["Agente promotor"].isin(promotores)
        & df["estado_evaluacion"].fillna("Sin evaluar").isin(estado_eval)
    ].copy()
    if texto:
        haystack = filtered[["Título", "Descripción", "Indicadores", "Medida"]].fillna("").agg(" ".join, axis=1).str.lower()
        filtered = filtered[haystack.str.contains(texto.lower(), regex=False)]
    return filtered


def apply_profile_scope(data: pd.DataFrame) -> pd.DataFrame:
    """Todos los perfiles ven todas las actividades.

    Los perfiles ya no limitan por eje; solo identifican quién accede y quién
    realiza cada modificación. Se mantiene la función para no tocar el flujo
    principal de la aplicación.
    """
    return data.copy()



@st.cache_data(ttl=300, show_spinner=False)
def get_manual_tasks(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    with engine.begin() as conn:
        df = pd.read_sql_query(text("SELECT * FROM manual_tasks ORDER BY updated_at DESC, task_id DESC"), conn)
    if df.empty:
        return pd.DataFrame(columns=[
            "task_id", "id_accion", "accion_titulo", "proximos_pasos", "descripcion", "responsable_seguimiento",
            "fecha_peticion", "fecha_vencimiento", "estado", "created_by", "created_at", "updated_by", "updated_at",
        ])
    df["task_id"] = df["task_id"].astype(int)
    df["id_accion"] = pd.to_numeric(df.get("id_accion"), errors="coerce")
    return clean_interface_dataframe(df)


@st.cache_data(ttl=300, show_spinner=False)
def get_task_overrides(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    with engine.begin() as conn:
        df = pd.read_sql_query(text("SELECT * FROM task_overrides"), conn)
    if df.empty:
        return pd.DataFrame(columns=["task_key", "descripcion", "fecha_vencimiento", "estado", "updated_by", "updated_at"])
    return clean_interface_dataframe(df)


def save_task_overrides(engine: Engine, tasks_df: pd.DataFrame, user_name: str) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    if tasks_df.empty:
        return
    with engine.begin() as conn:
        for _, row in tasks_df.iterrows():
            task_key = safe_text(row.get("task_key"))
            if not task_key:
                continue
            payload = {
                "task_key": task_key,
                "descripcion": safe_text(row.get("Descripción")),
                "fecha_vencimiento": safe_text(row.get("Fecha de vencimiento")),
                "estado": safe_text(row.get("Estado")) or "No iniciado",
                "updated_by": user_name,
                "updated_at": now,
            }
            if is_postgres(engine):
                conn.execute(text("""
                    INSERT INTO task_overrides (task_key, descripcion, fecha_vencimiento, estado, updated_by, updated_at)
                    VALUES (:task_key, :descripcion, :fecha_vencimiento, :estado, :updated_by, :updated_at)
                    ON CONFLICT(task_key) DO UPDATE SET
                        descripcion = excluded.descripcion,
                        fecha_vencimiento = excluded.fecha_vencimiento,
                        estado = excluded.estado,
                        updated_by = excluded.updated_by,
                        updated_at = excluded.updated_at
                """), payload)
            else:
                conn.execute(text("""
                    INSERT INTO task_overrides (task_key, descripcion, fecha_vencimiento, estado, updated_by, updated_at)
                    VALUES (:task_key, :descripcion, :fecha_vencimiento, :estado, :updated_by, :updated_at)
                    ON CONFLICT(task_key) DO UPDATE SET
                        descripcion = excluded.descripcion,
                        fecha_vencimiento = excluded.fecha_vencimiento,
                        estado = excluded.estado,
                        updated_by = excluded.updated_by,
                        updated_at = excluded.updated_at
                """), payload)
    clear_data_caches()


def add_manual_task(engine: Engine, payload: dict[str, Any], user_name: str) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    clean_payload = {
        "id_accion": payload.get("id_accion"),
        "accion_titulo": safe_text(payload.get("accion_titulo")),
        "proximos_pasos": safe_text(payload.get("proximos_pasos")),
        "descripcion": safe_text(payload.get("descripcion")),
        "responsable_seguimiento": safe_text(payload.get("responsable_seguimiento")),
        "fecha_peticion": safe_text(payload.get("fecha_peticion")),
        "fecha_vencimiento": safe_text(payload.get("fecha_vencimiento")),
        "estado": safe_text(payload.get("estado")) or "No iniciado",
        "created_by": user_name,
        "created_at": now,
        "updated_by": user_name,
        "updated_at": now,
    }
    fields = list(clean_payload.keys())
    with engine.begin() as conn:
        conn.execute(
            text(f"""
                INSERT INTO manual_tasks ({", ".join(fields)})
                VALUES ({", ".join([":" + f for f in fields])})
            """),
            clean_payload,
        )
    clear_data_caches()


def update_manual_tasks(engine: Engine, tasks_df: pd.DataFrame, user_name: str) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    if tasks_df.empty:
        return
    with engine.begin() as conn:
        for _, row in tasks_df.iterrows():
            task_id = int(row.get("task_id"))
            id_value = row.get("id_accion")
            try:
                id_accion = int(id_value) if pd.notna(id_value) and str(id_value).strip() else None
            except Exception:
                id_accion = None
            payload = {
                "task_id": task_id,
                "id_accion": id_accion,
                "accion_titulo": safe_text(row.get("accion_titulo")),
                "proximos_pasos": safe_text(row.get("proximos_pasos")),
                "descripcion": safe_text(row.get("descripcion")),
                "responsable_seguimiento": safe_text(row.get("responsable_seguimiento")),
                "fecha_peticion": safe_text(row.get("fecha_peticion")),
                "fecha_vencimiento": safe_text(row.get("fecha_vencimiento")),
                "estado": safe_text(row.get("estado")) or "No iniciado",
                "updated_by": user_name,
                "updated_at": now,
            }
            conn.execute(text("""
                UPDATE manual_tasks
                SET id_accion = :id_accion,
                    accion_titulo = :accion_titulo,
                    proximos_pasos = :proximos_pasos,
                    descripcion = :descripcion,
                    responsable_seguimiento = :responsable_seguimiento,
                    fecha_peticion = :fecha_peticion,
                    fecha_vencimiento = :fecha_vencimiento,
                    estado = :estado,
                    updated_by = :updated_by,
                    updated_at = :updated_at
                WHERE task_id = :task_id
            """), payload)
    clear_data_caches()


def delete_manual_task(engine: Engine, task_id: int) -> None:
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM manual_tasks WHERE task_id = :task_id"), {"task_id": int(task_id)})
    clear_data_caches()


def build_action_task_rows(data: pd.DataFrame, overrides: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame(columns=[
            "task_key", "Id_acción + Título", "Próximos pasos", "Descripción", "Responsable de seguimiento",
            "Fecha de petición", "Fecha de vencimiento", "Estado",
        ])
    rows = []
    overrides_by_key = {}
    if not overrides.empty:
        overrides_by_key = {safe_text(row.get("task_key")): row for _, row in overrides.iterrows()}
    for _, row in data.iterrows():
        next_steps = safe_text(row.get("proximos_pasos"))
        if not next_steps:
            continue
        id_accion = int(row.get("id_accion"))
        task_key = f"accion_{id_accion}"
        override = overrides_by_key.get(task_key)
        default_request_date = safe_text(row.get("fecha_actualizacion")) or safe_text(row.get("updated_at")) or date.today().isoformat()
        rows.append({
            "task_key": task_key,
            "Id_acción + Título": f"{id_accion:03d} - {safe_text(row.get('Título'))}",
            "Próximos pasos": next_steps,
            "Descripción": safe_text(override.get("descripcion")) if override is not None else "",
            "Responsable de seguimiento": safe_text(row.get("responsable_seguimiento")),
            "Fecha de petición": default_request_date[:10],
            "Fecha de vencimiento": safe_text(override.get("fecha_vencimiento")) if override is not None else "",
            "Estado": safe_text(override.get("estado")) if override is not None else "No iniciado",
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=[
            "task_key", "Id_acción + Título", "Próximos pasos", "Descripción", "Responsable de seguimiento",
            "Fecha de petición", "Fecha de vencimiento", "Estado",
        ])
    df["Estado"] = df["Estado"].apply(lambda x: x if x in TASK_STATUS_OPTIONS else "No iniciado")
    return clean_interface_dataframe(df)



def parse_iso_date(value: Any) -> date | None:
    raw = safe_text(value)
    if not raw:
        return None
    for candidate in [raw[:10], raw]:
        try:
            return datetime.fromisoformat(candidate).date()
        except Exception:
            pass
    try:
        parsed = pd.to_datetime(raw, errors="coerce")
        if pd.notna(parsed):
            return parsed.date()
    except Exception:
        pass
    return None


def get_task_mail_config() -> dict[str, Any]:
    """SMTP configuration for overdue-task reminders.

    Required Streamlit Secrets:
    - SMTP_HOST
    - SMTP_USER
    - SMTP_PASSWORD

    Optional Streamlit Secrets:
    - SMTP_PORT, default 587
    - SMTP_USE_SSL, default false
    - MAIL_FROM, default SMTP_USER
    - TASK_REMINDER_TO, default osasunsarea@erandio.eus
    - TASK_REMINDER_ENABLED, default true
    """
    enabled = get_secret("TASK_REMINDER_ENABLED", "true").strip().lower() not in {"0", "false", "no", "off"}
    host = get_secret("SMTP_HOST", "").strip()
    user = get_secret("SMTP_USER", "").strip()
    password = get_secret("SMTP_PASSWORD", "").strip()
    try:
        port = int(get_secret("SMTP_PORT", "587").strip() or "587")
    except Exception:
        port = 587
    use_ssl = get_secret("SMTP_USE_SSL", "false").strip().lower() in {"1", "true", "yes", "on"}
    return {
        "enabled": enabled,
        "host": host,
        "port": port,
        "user": user,
        "password": password,
        "use_ssl": use_ssl,
        "mail_from": get_secret("MAIL_FROM", user or "").strip(),
        "recipient": get_secret("TASK_REMINDER_TO", "osasunsarea@erandio.eus").strip(),
    }


def smtp_is_configured() -> bool:
    cfg = get_task_mail_config()
    return bool(cfg["enabled"] and cfg["host"] and cfg["user"] and cfg["password"] and cfg["mail_from"] and cfg["recipient"])


def send_task_email(subject: str, body: str, recipient: str) -> None:
    cfg = get_task_mail_config()
    if not smtp_is_configured():
        raise RuntimeError("SMTP no configurado. Faltan SMTP_HOST, SMTP_USER, SMTP_PASSWORD, MAIL_FROM o TASK_REMINDER_TO.")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = cfg["mail_from"]
    msg["To"] = recipient
    msg.set_content(body)
    if cfg["use_ssl"]:
        with smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=20) as smtp:
            smtp.login(cfg["user"], cfg["password"])
            smtp.send_message(msg)
    else:
        with smtplib.SMTP(cfg["host"], cfg["port"], timeout=20) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(cfg["user"], cfg["password"])
            smtp.send_message(msg)


def get_sent_task_reminder_keys(engine: Engine) -> set[str]:
    try:
        with engine.begin() as conn:
            rows = conn.execute(text("SELECT reminder_key FROM task_email_reminders WHERE status = 'sent'")).fetchall()
        return {safe_text(row[0]) for row in rows}
    except Exception:
        return set()


def record_task_reminder(engine: Engine, payload: dict[str, Any]) -> None:
    fields = list(payload.keys())
    with engine.begin() as conn:
        conn.execute(
            text(f"""
                INSERT INTO task_email_reminders ({", ".join(fields)})
                VALUES ({", ".join([":" + f for f in fields])})
                ON CONFLICT(reminder_key) DO UPDATE SET
                    sent_at = excluded.sent_at,
                    status = excluded.status,
                    error_message = excluded.error_message
            """),
            payload,
        )


def build_overdue_task_candidates(data: pd.DataFrame, engine: Engine) -> list[dict[str, Any]]:
    """Return overdue, active tasks from both derived and manual task sources."""
    today = date.today()
    candidates: list[dict[str, Any]] = []

    try:
        overrides = get_task_overrides(engine)
        action_tasks = build_action_task_rows(data, overrides)
    except Exception:
        action_tasks = pd.DataFrame()

    if not action_tasks.empty:
        for _, row in action_tasks.iterrows():
            status = safe_text(row.get("Estado")) or "No iniciado"
            if status in {"Completado", "Descartado"}:
                continue
            due_date = parse_iso_date(row.get("Fecha de vencimiento"))
            if not due_date or due_date >= today:
                continue
            task_key = safe_text(row.get("task_key"))
            action_label = safe_text(row.get("Id_acción + Título"))
            id_accion = None
            try:
                id_accion = int(action_label.split(" - ", 1)[0])
            except Exception:
                pass
            candidates.append({
                "reminder_key": f"derived:{task_key}:{due_date.isoformat()}",
                "task_source": "Ficha de evaluación",
                "task_identifier": task_key,
                "id_accion": id_accion,
                "accion_titulo": action_label,
                "proximos_pasos": safe_text(row.get("Próximos pasos")),
                "descripcion": safe_text(row.get("Descripción")),
                "responsable_seguimiento": safe_text(row.get("Responsable de seguimiento")),
                "fecha_vencimiento": due_date.isoformat(),
                "estado": status,
            })

    try:
        manual_tasks = get_manual_tasks(engine)
    except Exception:
        manual_tasks = pd.DataFrame()
    if not manual_tasks.empty:
        for _, row in manual_tasks.iterrows():
            status = safe_text(row.get("estado")) or "No iniciado"
            if status in {"Completado", "Descartado"}:
                continue
            due_date = parse_iso_date(row.get("fecha_vencimiento"))
            if not due_date or due_date >= today:
                continue
            task_id = safe_text(row.get("task_id"))
            id_accion = None
            try:
                id_value = row.get("id_accion")
                id_accion = int(id_value) if pd.notna(id_value) and safe_text(id_value) else None
            except Exception:
                id_accion = None
            accion_titulo = safe_text(row.get("accion_titulo"))
            if id_accion:
                action_label = f"{id_accion:03d} - {accion_titulo}"
            else:
                action_label = "Sin acción asociada"
            candidates.append({
                "reminder_key": f"manual:{task_id}:{due_date.isoformat()}",
                "task_source": "Tarea manual",
                "task_identifier": task_id,
                "id_accion": id_accion,
                "accion_titulo": action_label,
                "proximos_pasos": safe_text(row.get("proximos_pasos")),
                "descripcion": safe_text(row.get("descripcion")),
                "responsable_seguimiento": safe_text(row.get("responsable_seguimiento")),
                "fecha_vencimiento": due_date.isoformat(),
                "estado": status,
            })
    return candidates


def send_overdue_task_reminders(engine: Engine, data: pd.DataFrame, force: bool = False) -> tuple[int, list[str]]:
    """Send one reminder per overdue task/due-date.

    Streamlit Cloud is not a background worker. This check runs when the app is opened
    or when Administración explicitly forces it. The reminder table prevents duplicate
    sends for the same task and due date.
    """
    cfg = get_task_mail_config()
    if not cfg["enabled"]:
        return 0, ["Recordatorios de tareas desactivados por TASK_REMINDER_ENABLED."]
    if not smtp_is_configured():
        return 0, ["Correo no configurado: faltan secretos SMTP."]

    already_sent = get_sent_task_reminder_keys(engine)
    candidates = build_overdue_task_candidates(data, engine)
    to_send = [item for item in candidates if force or item["reminder_key"] not in already_sent]
    sent_count = 0
    errors: list[str] = []
    recipient = cfg["recipient"]
    now = datetime.now().isoformat(timespec="seconds")
    for task in to_send:
        subject = f"Tarea vencida: {task['accion_titulo']}"
        body = "\n".join([
            "Se ha detectado una tarea vencida en la herramienta Erandio Mugitzen ari da!.",
            "",
            f"Origen: {task['task_source']}",
            f"Actividad: {task['accion_titulo']}",
            f"Próximos pasos: {task['proximos_pasos'] or '-'}",
            f"Descripción: {task['descripcion'] or '-'}",
            f"Responsable de seguimiento: {task['responsable_seguimiento'] or '-'}",
            f"Fecha de vencimiento: {task['fecha_vencimiento']}",
            f"Estado: {task['estado']}",
            "",
            "Revisa la pestaña Tareas para actualizar el estado o la fecha de vencimiento.",
        ])
        record_payload = {
            "reminder_key": task["reminder_key"],
            "task_source": task["task_source"],
            "task_identifier": safe_text(task["task_identifier"]),
            "id_accion": task.get("id_accion"),
            "accion_titulo": task["accion_titulo"],
            "proximos_pasos": task["proximos_pasos"],
            "responsable_seguimiento": task["responsable_seguimiento"],
            "fecha_vencimiento": task["fecha_vencimiento"],
            "recipient": recipient,
            "sent_at": now,
            "status": "sent",
            "error_message": "",
        }
        try:
            send_task_email(subject, body, recipient)
            record_task_reminder(engine, record_payload)
            sent_count += 1
        except Exception as exc:
            record_payload["status"] = "error"
            record_payload["error_message"] = safe_text(exc)[:500]
            try:
                record_task_reminder(engine, record_payload)
            except Exception:
                pass
            errors.append(f"{task['accion_titulo']}: {exc}")
    clear_data_caches()
    return sent_count, errors


def maybe_send_overdue_task_reminders(engine: Engine, data: pd.DataFrame) -> None:
    """Run a light automatic reminder check at most once per browser session/day."""
    session_key = f"task_reminder_check_{date.today().isoformat()}"
    if st.session_state.get(session_key):
        return
    st.session_state[session_key] = True
    sent, errors = send_overdue_task_reminders(engine, data, force=False)
    st.session_state["task_reminder_last_result"] = {"sent": sent, "errors": errors, "checked_at": datetime.now().isoformat(timespec="seconds")}


def combine_task_rows_for_calendar(data: pd.DataFrame, engine: Engine) -> pd.DataFrame:
    """Construye una tabla de tareas con fecha para mostrar en Calendario."""
    rows: list[dict[str, Any]] = []
    try:
        overrides = get_task_overrides(engine)
        action_tasks = build_action_task_rows(data, overrides)
    except Exception:
        action_tasks = pd.DataFrame()
    if not action_tasks.empty:
        for _, row in action_tasks.iterrows():
            due_date = safe_text(row.get("Fecha de vencimiento"))
            request_date = safe_text(row.get("Fecha de petición"))
            date_value = due_date or request_date
            if not date_value:
                continue
            rows.append({
                "fecha": date_value,
                "hora": "",
                "tipo": "Tarea",
                "origen": "Ficha de evaluación",
                "titulo": safe_text(row.get("Id_acción + Título")),
                "descripcion": safe_text(row.get("Próximos pasos")),
                "responsable": safe_text(row.get("Responsable de seguimiento")),
                "estado": safe_text(row.get("Estado")) or "No iniciado",
                "ubicacion": "",
                "fecha_vencimiento": due_date,
            })
    try:
        manual_tasks = get_manual_tasks(engine)
    except Exception:
        manual_tasks = pd.DataFrame()
    if not manual_tasks.empty:
        for _, row in manual_tasks.iterrows():
            due_date = safe_text(row.get("fecha_vencimiento"))
            request_date = safe_text(row.get("fecha_peticion"))
            date_value = due_date or request_date
            if not date_value:
                continue
            id_accion = safe_text(row.get("id_accion"))
            accion = safe_text(row.get("accion_titulo"))
            if id_accion and id_accion not in {"0", "0.0"}:
                try:
                    title = f"{int(float(id_accion)):03d} - {accion}"
                except Exception:
                    title = accion or "Tarea manual"
            else:
                title = "Tarea manual"
            task_text = safe_text(row.get("proximos_pasos")) or safe_text(row.get("descripcion")) or "Tarea manual"
            rows.append({
                "fecha": date_value,
                "hora": "",
                "tipo": "Tarea",
                "origen": "Manual",
                "titulo": title,
                "descripcion": task_text,
                "responsable": safe_text(row.get("responsable_seguimiento")),
                "estado": safe_text(row.get("estado")) or "No iniciado",
                "ubicacion": "",
                "fecha_vencimiento": due_date,
            })
    df = pd.DataFrame(rows, columns=["fecha", "hora", "tipo", "origen", "titulo", "descripcion", "responsable", "estado", "ubicacion", "fecha_vencimiento"])
    if not df.empty:
        df["fecha_dt"] = pd.to_datetime(df["fecha"], errors="coerce")
        df = df.dropna(subset=["fecha_dt"]).sort_values(["fecha_dt", "hora", "titulo"])
    return clean_interface_dataframe(df)


@st.cache_data(ttl=300, show_spinner=False)
def get_calendar_events(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    with engine.begin() as conn:
        df = pd.read_sql_query(text("SELECT * FROM calendar_events ORDER BY start_date ASC, start_time ASC, event_id ASC"), conn)
    if df.empty:
        return pd.DataFrame(columns=[
            "event_id", "title", "event_type", "description", "location", "start_date", "start_time",
            "end_date", "end_time", "id_accion", "accion_titulo", "responsable", "created_by", "created_at", "updated_by", "updated_at",
        ])
    df["event_id"] = df["event_id"].astype(int)
    return clean_interface_dataframe(df)


def add_calendar_event(engine: Engine, payload: dict[str, Any], user_name: str) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    payload = {**payload, "created_by": user_name, "created_at": now, "updated_by": user_name, "updated_at": now}
    fields = list(payload.keys())
    with engine.begin() as conn:
        conn.execute(
            text(f"""
                INSERT INTO calendar_events ({", ".join(fields)})
                VALUES ({", ".join([":" + f for f in fields])})
            """),
            payload,
        )
    clear_data_caches()


def update_calendar_events(engine: Engine, edited: pd.DataFrame, user_name: str) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as conn:
        for _, row in edited.iterrows():
            event_id = int(row["event_id"])
            conn.execute(
                text("""
                    UPDATE calendar_events SET
                        title = :title,
                        event_type = :event_type,
                        description = :description,
                        location = :location,
                        start_date = :start_date,
                        start_time = :start_time,
                        end_date = :end_date,
                        end_time = :end_time,
                        responsable = :responsable,
                        updated_by = :updated_by,
                        updated_at = :updated_at
                    WHERE event_id = :event_id
                """),
                {
                    "event_id": event_id,
                    "title": safe_text(row.get("title")) or "Sin título",
                    "event_type": safe_text(row.get("event_type")) or "Evento",
                    "description": safe_text(row.get("description")),
                    "location": safe_text(row.get("location")),
                    "start_date": safe_text(row.get("start_date")),
                    "start_time": safe_text(row.get("start_time")),
                    "end_date": safe_text(row.get("end_date")),
                    "end_time": safe_text(row.get("end_time")),
                    "responsable": safe_text(row.get("responsable")),
                    "updated_by": user_name,
                    "updated_at": now,
                },
            )
    clear_data_caches()


def delete_calendar_event(engine: Engine, event_id: int) -> None:
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM calendar_events WHERE event_id = :event_id"), {"event_id": int(event_id)})
    clear_data_caches()


def render_calendar(data: pd.DataFrame, user_name: str, engine: Engine) -> None:
    st.subheader("Calendario")
    st.caption("Vista de tres meses de eventos, reuniones, hitos y tareas programadas.")

    events = get_calendar_events(engine)
    task_calendar = combine_task_rows_for_calendar(data, engine)

    with st.expander("Añadir evento o reunión", expanded=False):
        action_options = ["Sin acción asociada"] + data.apply(lambda row: f"{int(row['id_accion']):03d} - {row['Título']}", axis=1).tolist()
        with st.form("calendar_event_form", clear_on_submit=True):
            title = st.text_input("Título del evento o reunión")
            event_type = st.selectbox("Tipo", EVENT_TYPE_OPTIONS)
            selected_action = st.selectbox("Acción asociada", action_options)
            description = st.text_area("Descripción")
            location = st.text_input("Lugar / enlace de reunión")
            c1, c2, c3, c4 = st.columns(4)
            start_date = c1.date_input("Fecha de inicio", value=date.today())
            start_time = c2.text_input("Hora de inicio", placeholder="09:30")
            end_date = c3.date_input("Fecha de fin", value=date.today())
            end_time = c4.text_input("Hora de fin", placeholder="10:30")
            responsable = st.text_input("Responsable")
            submitted = st.form_submit_button("Añadir al calendario")
        if submitted:
            if not user_name:
                st.error("Antes de añadir un evento, entra con un perfil identificado.")
            elif not safe_text(title):
                st.error("El título es obligatorio.")
            else:
                id_accion = None
                accion_titulo = ""
                if selected_action != "Sin acción asociada":
                    id_accion = int(selected_action.split(" - ")[0])
                    accion_titulo = selected_action.split(" - ", 1)[1]
                add_calendar_event(engine, {
                    "title": title,
                    "event_type": event_type,
                    "description": description,
                    "location": location,
                    "start_date": start_date.isoformat(),
                    "start_time": start_time,
                    "end_date": end_date.isoformat(),
                    "end_time": end_time,
                    "id_accion": id_accion,
                    "accion_titulo": accion_titulo,
                    "responsable": responsable,
                }, user_name)
                st.success("Evento añadido al calendario.")
                clear_data_caches()
                st.rerun()

    today = date.today()
    month_names = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    weekday_names = ["L", "M", "X", "J", "V", "S", "D"]

    def add_months(year: int, month: int, offset: int) -> tuple[int, int]:
        absolute = year * 12 + (month - 1) + offset
        return absolute // 12, absolute % 12 + 1

    def month_label(value: str) -> str:
        y, m = value.split("-")
        return f"{month_names[int(m) - 1]} {y}"

    month_values = [f"{y}-{m:02d}" for y in range(today.year - 1, today.year + 4) for m in range(1, 13)]
    default_month = f"{today.year}-{today.month:02d}"
    if "calendar_start_month" not in st.session_state:
        st.session_state["calendar_start_month"] = default_month
    if st.session_state["calendar_start_month"] not in month_values:
        st.session_state["calendar_start_month"] = default_month

    st.markdown("### Vista de calendario")
    bar_left, bar_center, bar_right = st.columns([1, 5, 1])
    with bar_left:
        if st.button("◀ Anteriores", use_container_width=True):
            current_index = month_values.index(st.session_state["calendar_start_month"])
            st.session_state["calendar_start_month"] = month_values[max(0, current_index - 3)]
            st.rerun()
    with bar_center:
        selected_month = st.select_slider(
            "Mes inicial visible",
            options=month_values,
            value=st.session_state["calendar_start_month"],
            format_func=month_label,
            key="calendar_start_month_slider",
        )
        st.session_state["calendar_start_month"] = selected_month
    with bar_right:
        if st.button("Posteriores ▶", use_container_width=True):
            current_index = month_values.index(st.session_state["calendar_start_month"])
            st.session_state["calendar_start_month"] = month_values[min(len(month_values) - 1, current_index + 3)]
            st.rerun()

    c1, c2 = st.columns([1, 2])
    item_filter = c1.selectbox("Mostrar", CALENDAR_ITEM_TYPES, key="calendar_type_filter")
    search = c2.text_input("Buscar", placeholder="Título, descripción, responsable, lugar...")

    start_year, start_month = [int(x) for x in st.session_state["calendar_start_month"].split("-")]
    visible_months = [add_months(start_year, start_month, offset) for offset in range(3)]
    first_visible_day = date(visible_months[0][0], visible_months[0][1], 1)
    end_y, end_m = add_months(visible_months[-1][0], visible_months[-1][1], 1)
    last_visible_day = date(end_y, end_m, 1) - pd.Timedelta(days=1)
    last_visible_day = last_visible_day.date() if hasattr(last_visible_day, "date") else last_visible_day

    agenda_rows: list[dict[str, Any]] = []
    if item_filter in {"Todos", "Eventos/Reuniones"} and not events.empty:
        for _, row in events.iterrows():
            agenda_rows.append({
                "fecha": safe_text(row.get("start_date")),
                "hora": safe_text(row.get("start_time")),
                "tipo": safe_text(row.get("event_type")) or "Evento",
                "origen": "Calendario",
                "titulo": safe_text(row.get("title")),
                "descripcion": safe_text(row.get("description")),
                "responsable": safe_text(row.get("responsable")),
                "estado": "",
                "ubicacion": safe_text(row.get("location")),
                "fecha_vencimiento": "",
            })
    if item_filter in {"Todos", "Tareas"} and not task_calendar.empty:
        agenda_rows.extend(task_calendar.drop(columns=[c for c in ["fecha_dt"] if c in task_calendar.columns]).to_dict("records"))

    agenda = pd.DataFrame(agenda_rows, columns=["fecha", "hora", "tipo", "origen", "titulo", "descripcion", "responsable", "estado", "ubicacion", "fecha_vencimiento"])
    if not agenda.empty:
        agenda["fecha_dt"] = pd.to_datetime(agenda["fecha"], errors="coerce")
        agenda = agenda.dropna(subset=["fecha_dt"])
        agenda = agenda[(agenda["fecha_dt"].dt.date >= first_visible_day) & (agenda["fecha_dt"].dt.date <= last_visible_day)]
        if search:
            haystack = agenda[["titulo", "descripcion", "responsable", "ubicacion", "estado", "tipo", "origen"]].fillna("").agg(" ".join, axis=1).str.lower()
            agenda = agenda[haystack.str.contains(search.lower(), regex=False)]
        agenda = agenda.sort_values(["fecha_dt", "hora", "tipo", "titulo"])

    def event_color(tipo: str, origen: str, estado: str) -> str:
        tipo_l = safe_text(tipo).lower()
        origen_l = safe_text(origen).lower()
        estado_l = safe_text(estado).lower()
        if "tarea" in origen_l or "tarea" in tipo_l:
            if estado_l == "completado":
                return COLOR_SKY
            if estado_l == "descartado":
                return COLOR_MAUVE
            return COLOR_GOLD
        if "reun" in tipo_l:
            return COLOR_CORAL
        if "actividad" in tipo_l:
            return COLOR_SKY
        if "hito" in tipo_l:
            return COLOR_MAUVE
        return COLOR_NAVY

    def build_events_by_day(df: pd.DataFrame) -> dict[date, list[dict[str, str]]]:
        grouped: dict[date, list[dict[str, str]]] = {}
        if df.empty:
            return grouped
        for _, row in df.iterrows():
            day_value = row.get("fecha_dt")
            if pd.isna(day_value):
                continue
            day = day_value.date()
            grouped.setdefault(day, []).append({
                "hora": safe_text(row.get("hora")),
                "tipo": safe_text(row.get("tipo")),
                "origen": safe_text(row.get("origen")),
                "titulo": safe_text(row.get("titulo")),
                "descripcion": safe_text(row.get("descripcion")),
                "responsable": safe_text(row.get("responsable")),
                "estado": safe_text(row.get("estado")),
                "ubicacion": safe_text(row.get("ubicacion")),
            })
        return grouped

    events_by_day = build_events_by_day(agenda)
    if agenda.empty:
        st.info("No hay eventos ni tareas para los meses y filtros seleccionados.")
    else:
        total_items = len(agenda)
        total_days = agenda["fecha_dt"].dt.date.nunique()
        st.caption(f"{total_items} entradas distribuidas en {total_days} días entre {month_label(st.session_state['calendar_start_month'])} y {month_names[visible_months[-1][1] - 1]} {visible_months[-1][0]}.")

    st.markdown(
        f"""
        <style>
        .three-month-calendar-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 18px;
            margin-top: 12px;
        }}
        .month-card {{
            background: #FFFFFF;
            border: 1px solid rgba(28,48,84,0.18);
            border-radius: 18px;
            padding: 14px;
            box-shadow: 0 4px 14px rgba(28,48,84,0.06);
        }}
        .month-title {{
            font-weight: 800;
            color: {COLOR_NAVY};
            font-size: 1.05rem;
            margin-bottom: 10px;
            letter-spacing: .2px;
        }}
        .month-weekdays, .month-days {{
            display: grid;
            grid-template-columns: repeat(7, minmax(0, 1fr));
            gap: 4px;
        }}
        .weekday-cell {{
            text-align: center;
            color: {COLOR_NAVY};
            font-size: .70rem;
            font-weight: 700;
            opacity: .72;
            padding-bottom: 4px;
        }}
        .day-cell {{
            min-height: 72px;
            border-radius: 10px;
            background: #FFFFFF;
            border: 1px solid rgba(28,48,84,0.10);
            padding: 5px;
            overflow: hidden;
        }}
        .day-cell.muted {{
            background: rgba(28,48,84,0.025);
            opacity: .38;
        }}
        .day-cell.today {{
            border: 2px solid {COLOR_CORAL};
        }}
        .day-number {{
            color: {COLOR_NAVY};
            font-weight: 800;
            font-size: .75rem;
            line-height: 1;
            margin-bottom: 4px;
        }}
        .day-items {{
            display: flex;
            flex-direction: column;
            gap: 2px;
        }}
        .day-item {{
            color: #FFFFFF;
            font-size: .60rem;
            line-height: 1.15;
            border-radius: 5px;
            padding: 2px 3px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        .more-items {{
            color: {COLOR_NAVY};
            font-size: .58rem;
            font-weight: 700;
            padding-left: 2px;
        }}
        @media (max-width: 1100px) {{
            .three-month-calendar-grid {{ grid-template-columns: 1fr; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    cal = calendar.Calendar(firstweekday=0)
    html_parts = ['<div class="three-month-calendar-grid">']
    for year, month in visible_months:
        html_parts.append('<div class="month-card">')
        html_parts.append(f'<div class="month-title">{month_names[month - 1]} {year}</div>')
        html_parts.append('<div class="month-weekdays">')
        for wd in weekday_names:
            html_parts.append(f'<div class="weekday-cell">{wd}</div>')
        html_parts.append('</div><div class="month-days">')
        for week in cal.monthdatescalendar(int(year), month):
            for day in week:
                classes = ["day-cell"]
                if day.month != month:
                    classes.append("muted")
                if day == today:
                    classes.append("today")
                day_items = events_by_day.get(day, []) if day.month == month else []
                html_parts.append(f'<div class="{" ".join(classes)}">')
                html_parts.append(f'<div class="day-number">{day.day}</div>')
                if day_items:
                    html_parts.append('<div class="day-items">')
                    for item in day_items[:3]:
                        label = safe_text(item.get("titulo")) or safe_text(item.get("descripcion")) or safe_text(item.get("tipo"))
                        if safe_text(item.get("hora")):
                            label = f"{safe_text(item.get('hora'))} · {label}"
                        color = event_color(item.get("tipo", ""), item.get("origen", ""), item.get("estado", ""))
                        html_parts.append(f'<div class="day-item" style="background:{color};" title="{html.escape(label)}">{html.escape(label)}</div>')
                    if len(day_items) > 3:
                        html_parts.append(f'<div class="more-items">+{len(day_items) - 3} más</div>')
                    html_parts.append('</div>')
                html_parts.append('</div>')
        html_parts.append('</div></div>')
    html_parts.append('</div>')
    st.markdown("".join(html_parts), unsafe_allow_html=True)

    st.markdown("### Detalle")
    if agenda.empty:
        st.info("No hay detalle para mostrar.")
    else:
        agenda_display = agenda.copy()
        agenda_display["Fecha"] = agenda_display["fecha_dt"].dt.strftime("%Y-%m-%d")
        agenda_display = agenda_display.rename(columns={
            "hora": "Hora",
            "tipo": "Tipo",
            "origen": "Origen",
            "titulo": "Título / actividad",
            "descripcion": "Descripción / próximos pasos",
            "responsable": "Responsable",
            "estado": "Estado",
            "ubicacion": "Lugar / enlace",
            "fecha_vencimiento": "Fecha de vencimiento",
        })
        with st.expander("Ver agenda en tabla", expanded=False):
            st.dataframe(
                clean_interface_dataframe(agenda_display[["Fecha", "Hora", "Tipo", "Origen", "Título / actividad", "Descripción / próximos pasos", "Responsable", "Estado", "Lugar / enlace", "Fecha de vencimiento"]]),
                use_container_width=True,
                hide_index=True,
            )
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                clean_interface_dataframe(agenda_display).drop(columns=["fecha", "fecha_dt"], errors="ignore").to_excel(writer, index=False, sheet_name="Calendario")
            st.download_button(
                "Descargar calendario en Excel",
                buffer.getvalue(),
                "calendario.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.markdown("### Editar eventos y reuniones")
    if events.empty:
        st.info("Todavía no hay eventos o reuniones añadidos manualmente.")
    else:
        edit_cols = ["event_id", "title", "event_type", "description", "location", "start_date", "start_time", "end_date", "end_time", "responsable"]
        edited_events = st.data_editor(
            events[edit_cols],
            use_container_width=True,
            hide_index=True,
            disabled=["event_id"],
            column_config={
                "event_id": None,
                "title": st.column_config.TextColumn("Título", width="large"),
                "event_type": st.column_config.SelectboxColumn("Tipo", options=EVENT_TYPE_OPTIONS, required=True),
                "description": st.column_config.TextColumn("Descripción", width="large"),
                "location": st.column_config.TextColumn("Lugar / enlace", width="medium"),
                "start_date": st.column_config.TextColumn("Fecha inicio", help="Formato AAAA-MM-DD"),
                "start_time": st.column_config.TextColumn("Hora inicio"),
                "end_date": st.column_config.TextColumn("Fecha fin", help="Formato AAAA-MM-DD"),
                "end_time": st.column_config.TextColumn("Hora fin"),
                "responsable": st.column_config.TextColumn("Responsable", width="medium"),
            },
            key="calendar_events_editor",
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Guardar cambios en calendario"):
                if not user_name:
                    st.error("Antes de guardar, entra con un perfil identificado.")
                else:
                    update_calendar_events(engine, edited_events, user_name)
                    clear_data_caches()
                    st.success("Calendario actualizado.")
                    st.rerun()
        with c2:
            event_ids = events["event_id"].astype(int).tolist()
            delete_event_id = st.selectbox("Eliminar evento", event_ids, format_func=lambda x: f"Evento {x}", key="delete_calendar_event_select")
            if st.button("Eliminar evento seleccionado"):
                delete_calendar_event(engine, int(delete_event_id))
                clear_data_caches()
                st.success("Evento eliminado.")
                st.rerun()

def render_tasks(data: pd.DataFrame, user_name: str, engine: Engine) -> None:
    st.subheader("Tareas")
    st.caption("Seguimiento operativo de próximos pasos derivados de las fichas de evaluación y tareas añadidas manualmente.")

    overrides = get_task_overrides(engine)
    action_tasks = build_action_task_rows(data, overrides)

    st.markdown("### Tareas generadas desde Próximos pasos")
    st.caption("Se nutren automáticamente del campo Próximos pasos de cada Ficha de evaluación. Puedes completar descripción, vencimiento y estado.")
    if action_tasks.empty:
        st.info("Todavía no hay próximos pasos registrados en las fichas de evaluación.")
    else:
        edited_action_tasks = st.data_editor(
            action_tasks,
            use_container_width=True,
            hide_index=True,
            disabled=["task_key", "Id_acción + Título", "Próximos pasos", "Responsable de seguimiento", "Fecha de petición"],
            column_config={
                "task_key": None,
                "Id_acción + Título": st.column_config.TextColumn("Id_acción + Título", width="large"),
                "Próximos pasos": st.column_config.TextColumn("Próximos pasos", width="large"),
                "Descripción": st.column_config.TextColumn("Descripción", width="large"),
                "Responsable de seguimiento": st.column_config.TextColumn("Responsable de seguimiento", width="medium"),
                "Fecha de petición": st.column_config.TextColumn("Fecha de petición", width="small"),
                "Fecha de vencimiento": st.column_config.TextColumn("Fecha de vencimiento", help="Formato recomendado: AAAA-MM-DD", width="small"),
                "Estado": st.column_config.SelectboxColumn("Estado", options=TASK_STATUS_OPTIONS, required=True, width="small"),
            },
            key="action_tasks_editor",
        )
        if st.button("Guardar cambios en tareas de fichas", key="save_action_tasks"):
            if not user_name:
                st.error("Antes de guardar, escribe tu nombre o entra con un perfil identificado.")
            else:
                save_task_overrides(engine, edited_action_tasks, user_name)
                st.success("Tareas actualizadas.")
                st.rerun()

    st.markdown("### Tareas manuales")
    with st.expander("Añadir tarea manual", expanded=False):
        action_options = ["Sin acción asociada"] + data.apply(lambda row: f"{int(row['id_accion']):03d} - {row['Título']}", axis=1).tolist()
        with st.form("manual_task_form", clear_on_submit=True):
            selected_action = st.selectbox("Id_acción + Título", action_options)
            manual_next_steps = st.text_area("Próximos pasos")
            manual_description = st.text_area("Descripción")
            c1, c2 = st.columns(2)
            manual_responsible = c1.text_input("Responsable de seguimiento")
            manual_request_date = c2.date_input("Fecha de petición", value=date.today())
            c3, c4 = st.columns(2)
            manual_due_date = c3.date_input("Fecha de vencimiento", value=date.today())
            manual_status = c4.selectbox("Estado", TASK_STATUS_OPTIONS)
            submitted = st.form_submit_button("Añadir tarea")
        if submitted:
            if not user_name:
                st.error("Antes de añadir una tarea, escribe tu nombre o entra con un perfil identificado.")
            elif not safe_text(manual_next_steps) and not safe_text(manual_description):
                st.error("Introduce al menos Próximos pasos o Descripción.")
            else:
                id_accion = None
                action_title = ""
                if selected_action != "Sin acción asociada":
                    id_accion = int(selected_action.split(" - ")[0])
                    action_title = selected_action.split(" - ", 1)[1]
                add_manual_task(engine, {
                    "id_accion": id_accion,
                    "accion_titulo": action_title,
                    "proximos_pasos": manual_next_steps,
                    "descripcion": manual_description,
                    "responsable_seguimiento": manual_responsible,
                    "fecha_peticion": manual_request_date.isoformat(),
                    "fecha_vencimiento": manual_due_date.isoformat(),
                    "estado": manual_status,
                }, user_name)
                st.success("Tarea manual añadida.")
                st.rerun()

    manual_tasks = get_manual_tasks(engine)
    if manual_tasks.empty:
        st.info("Todavía no hay tareas manuales.")
    else:
        visible_manual = manual_tasks.copy()
        visible_manual["Id_acción + Título"] = visible_manual.apply(
            lambda row: f"{int(row['id_accion']):03d} - {safe_text(row.get('accion_titulo'))}" if pd.notna(row.get("id_accion")) and safe_text(row.get("id_accion")) not in ["", "0"] else "Sin acción asociada",
            axis=1,
        )
        editor_cols = [
            "task_id", "id_accion", "accion_titulo", "Id_acción + Título", "proximos_pasos", "descripcion",
            "responsable_seguimiento", "fecha_peticion", "fecha_vencimiento", "estado",
        ]
        edited_manual = st.data_editor(
            visible_manual[editor_cols],
            use_container_width=True,
            hide_index=True,
            disabled=["task_id", "id_accion", "accion_titulo", "Id_acción + Título"],
            column_config={
                "task_id": None,
                "id_accion": None,
                "accion_titulo": None,
                "Id_acción + Título": st.column_config.TextColumn("Id_acción + Título", width="large"),
                "proximos_pasos": st.column_config.TextColumn("Próximos pasos", width="large"),
                "descripcion": st.column_config.TextColumn("Descripción", width="large"),
                "responsable_seguimiento": st.column_config.TextColumn("Responsable de seguimiento", width="medium"),
                "fecha_peticion": st.column_config.TextColumn("Fecha de petición", width="small"),
                "fecha_vencimiento": st.column_config.TextColumn("Fecha de vencimiento", help="Formato recomendado: AAAA-MM-DD", width="small"),
                "estado": st.column_config.SelectboxColumn("Estado", options=TASK_STATUS_OPTIONS, required=True, width="small"),
            },
            key="manual_tasks_editor",
        )
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("Guardar cambios en tareas manuales", key="save_manual_tasks"):
                if not user_name:
                    st.error("Antes de guardar, escribe tu nombre o entra con un perfil identificado.")
                else:
                    update_manual_tasks(engine, edited_manual.rename(columns={
                        "proximos_pasos": "proximos_pasos",
                        "descripcion": "descripcion",
                        "responsable_seguimiento": "responsable_seguimiento",
                        "fecha_peticion": "fecha_peticion",
                        "fecha_vencimiento": "fecha_vencimiento",
                        "estado": "estado",
                    }), user_name)
                    st.success("Tareas manuales actualizadas.")
                    st.rerun()
        with c2:
            task_ids = visible_manual["task_id"].astype(int).tolist()
            delete_id = st.selectbox("Eliminar tarea manual", task_ids, format_func=lambda x: f"Tarea {x}", key="delete_manual_task_select")
            if st.button("Eliminar tarea seleccionada", key="delete_manual_task"):
                delete_manual_task(engine, int(delete_id))
                st.success("Tarea eliminada.")
                st.rerun()

    combined = []
    if not action_tasks.empty:
        derived = action_tasks.copy()
        derived["Origen"] = "Ficha de evaluación"
        combined.append(derived[["Origen", "Id_acción + Título", "Próximos pasos", "Descripción", "Responsable de seguimiento", "Fecha de petición", "Fecha de vencimiento", "Estado"]])
    if not manual_tasks.empty:
        manual_export = visible_manual.rename(columns={
            "proximos_pasos": "Próximos pasos",
            "descripcion": "Descripción",
            "responsable_seguimiento": "Responsable de seguimiento",
            "fecha_peticion": "Fecha de petición",
            "fecha_vencimiento": "Fecha de vencimiento",
            "estado": "Estado",
        }).copy()
        manual_export["Origen"] = "Manual"
        combined.append(manual_export[["Origen", "Id_acción + Título", "Próximos pasos", "Descripción", "Responsable de seguimiento", "Fecha de petición", "Fecha de vencimiento", "Estado"]])
    if combined:
        export_df = pd.concat(combined, ignore_index=True)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="Tareas")
        st.download_button(
            "Descargar tareas en Excel",
            buffer.getvalue(),
            "tareas.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def render_dashboard(df: pd.DataFrame) -> None:
    chart_df = df.copy()

    # Limpieza específica para gráficos: evita categorías técnicas como
    # undefined/null/nan y no las convierte en leyenda visible.
    chart_df["Ámbito"] = chart_df["Ámbito"].apply(safe_text)
    chart_df["estado_evaluacion"] = chart_df["estado_evaluacion"].apply(safe_text)
    chart_df["estado_evaluacion"] = chart_df["estado_evaluacion"].where(
        chart_df["estado_evaluacion"].isin(STATUS_OPTIONS),
        "Sin evaluar",
    )

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Acciones filtradas", len(chart_df))
    col2.metric("Iniciadas", int((chart_df["estado_evaluacion"] == "Iniciado").sum()))
    col3.metric("Completadas", int((chart_df["estado_evaluacion"] == "Completado").sum()))
    col4.metric("Sin evaluar", int((chart_df["estado_evaluacion"] == "Sin evaluar").sum()))

    left, right = st.columns(2)
    with left:
        st.subheader("Acciones por ámbito y estado")
        chart_data = (
            chart_df[chart_df["Ámbito"].astype(str).str.strip() != ""]
            .groupby(["Ámbito", "estado_evaluacion"], dropna=False)
            .size()
            .reset_index(name="acciones")
        )
        if not chart_data.empty:
            fig = px.bar(
                chart_data,
                x="Ámbito",
                y="acciones",
                color="estado_evaluacion",
                text="acciones",
                barmode="stack",
                category_orders={"estado_evaluacion": STATUS_OPTIONS},
                color_discrete_sequence=BRAND_COLORS,
                labels={
                    "Ámbito": "Ámbito",
                    "acciones": "Acciones",
                    "estado_evaluacion": "Estado de evaluación",
                },
            )
            fig.update_traces(textposition="inside")
            st.plotly_chart(apply_plotly_brand_layout(fig), use_container_width=True)
        else:
            st.info("No hay datos suficientes para mostrar este gráfico.")
    with right:
        st.subheader("Estado de evaluación")
        chart_data = (
            chart_df[chart_df["estado_evaluacion"].astype(str).str.strip() != ""]
            .groupby("estado_evaluacion", dropna=False)
            .size()
            .reset_index(name="acciones")
        )
        chart_data = chart_data[chart_data["estado_evaluacion"].isin(STATUS_OPTIONS)]
        if not chart_data.empty:
            fig = px.pie(
                chart_data,
                names="estado_evaluacion",
                values="acciones",
                category_orders={"estado_evaluacion": STATUS_OPTIONS},
                color_discrete_sequence=BRAND_COLORS,
                labels={"estado_evaluacion": "Estado de evaluación", "acciones": "Acciones"},
            )
            st.plotly_chart(apply_plotly_brand_layout(fig), use_container_width=True)
        else:
            st.info("No hay datos suficientes para mostrar este gráfico.")

    st.subheader("Acciones pendientes de completar")
    columns = ["id_accion", "Ámbito", "Título", "Estado", "estado_evaluacion", "cumplimiento_indicadores", "responsable_seguimiento", "updated_by"]
    pending = chart_df[chart_df["estado_evaluacion"].isin(["Sin evaluar", "Iniciado"])]
    st.dataframe(clean_interface_dataframe(pending.sort_values(["estado_evaluacion", "id_accion"])[columns].head(20)), use_container_width=True, hide_index=True)

def render_matrix(df: pd.DataFrame, all_data: pd.DataFrame, contacts: pd.DataFrame, assignments: pd.DataFrame, user_name: str, engine: Engine) -> None:
    visible_columns = [
        "id_accion", "Ámbito", "Título", "Tipo", "Estado", "Agente promotor", "Temporalidad",
        "personas_red_asignadas", "num_personas_red",
        "estado_evaluacion", "cumplimiento_indicadores", "responsable_seguimiento", "fecha_actualizacion", "updated_by", "updated_at",
    ]
    st.dataframe(df[visible_columns], use_container_width=True, hide_index=True)

    st.markdown("### Asignar personas de la Red Local de Salud")
    if contacts.empty:
        st.info("No se ha encontrado el archivo data/Contactos.xlsx o no contiene contactos.")
    elif df.empty:
        st.info("No hay acciones para los filtros seleccionados.")
    else:
        options = df.apply(lambda row: f"{int(row['id_accion']):03d} - {row['Título']}", axis=1).tolist()
        selected = st.selectbox("Selecciona una acción para asignar personas", options, key="matrix_assignment_action")
        selected_id = int(selected.split(" - ")[0])
        current_ids = assignments.loc[assignments["id_accion"] == selected_id, "contacto_id"].astype(int).tolist() if not assignments.empty else []
        label_by_id = dict(zip(contacts["contacto_id"], contacts["contacto_label"]))
        contact_options = contacts["contacto_id"].astype(int).tolist()
        selected_contacts = st.multiselect(
            "Personas asignadas",
            contact_options,
            default=[cid for cid in current_ids if cid in set(contact_options)],
            format_func=lambda cid: label_by_id.get(cid, str(cid)),
            key="matrix_assignment_contacts",
        )
        if st.button("Guardar asignación de personas", key="save_matrix_assignment"):
            if not user_name:
                st.error("Antes de guardar, escribe tu nombre en la barra lateral.")
            else:
                save_action_contacts(engine, selected_id, selected_contacts, user_name)
                st.success("Asignación guardada.")
                st.rerun()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Evaluacion")
    st.download_button(
        "Descargar evaluación filtrada en Excel", buffer.getvalue(), "evaluacion_filtrada.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_contacts(contacts: pd.DataFrame, assignments: pd.DataFrame, data: pd.DataFrame, user_name: str, engine: Engine) -> None:
    st.subheader("Red Local de Salud")
    st.caption("Listado compartido de personas y entidades de la Red Local de Salud. Puedes importar el Excel inicial, añadir personas y editar los campos directamente desde esta pestaña.")

    with st.expander("Cargar o actualizar listado de contactos", expanded=contacts.empty):
        st.write("Usa este bloque solo si el listado está vacío o necesitas incorporar contactos desde un Excel.")
        c_load1, c_load2 = st.columns([1.2, 1])
        with c_load1:
            uploaded_contacts = st.file_uploader("Subir Excel de contactos", type=["xlsx"], key="contacts_excel_uploader")
        with c_load2:
            replace_contacts = st.checkbox("Sustituir listado completo", value=False, help="Si se activa, borra el listado actual y sus asignaciones antes de importar.")
        if uploaded_contacts is not None:
            try:
                imported_df = read_contacts_bytes(uploaded_contacts.getvalue())
                st.write(f"Contactos detectados en el Excel: {len(imported_df)}")
                if st.button("Importar contactos desde Excel", key="import_uploaded_contacts"):
                    if not user_name:
                        st.error("Antes de importar, escribe tu nombre en la barra lateral.")
                    else:
                        inserted = import_contacts_dataframe(engine, imported_df, user_name, replace_contacts)
                        st.success(f"Contactos importados: {inserted}")
                        st.rerun()
            except Exception as exc:
                st.error(f"No se pudo leer el Excel de contactos: {exc}")
        elif contacts.empty and CONTACTS_EXCEL.exists():
            if st.button("Cargar Contactos.xlsx incluido en la app", key="import_packaged_contacts"):
                inserted = import_contacts_dataframe(engine, load_default_contacts(), user_name or "carga inicial desde panel", replace=False)
                st.success(f"Contactos cargados: {inserted}")
                st.rerun()

    if contacts.empty:
        st.info("Todavía no hay contactos cargados. Puedes importarlos desde el Excel o crear el primer registro desde el formulario inferior.")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Contactos", len(contacts))
    col2.metric("Entidades", contacts["Entidad"].replace("", pd.NA).dropna().nunique() if not contacts.empty else 0)
    col3.metric("Comisiones", contacts["Comisión"].replace("", pd.NA).dropna().nunique() if not contacts.empty else 0)
    col4.metric("Contactos asignados", assignments["contacto_id"].nunique() if not assignments.empty else 0)

    with st.expander("Añadir nueva persona", expanded=contacts.empty):
        with st.form("add_contact_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            persona = c1.text_input("Persona")
            entidad = c2.text_input("Entidad")
            c3, c4 = st.columns(2)
            categoria = c3.text_input("Categoría")
            subcategoria = c4.text_input("Subcategoría")
            c5, c6 = st.columns(2)
            comision = c5.text_input("Comisión")
            perfil = c6.text_input("Perfil")
            c7, c8 = st.columns(2)
            telefono = c7.text_input("Teléfono")
            mail = c8.text_input("Mail")
            submitted = st.form_submit_button("Añadir persona")
        if submitted:
            if not user_name:
                st.error("Antes de añadir una persona, escribe tu nombre en la barra lateral.")
            elif not safe_text(persona) and not safe_text(entidad):
                st.error("Introduce al menos la persona o la entidad.")
            else:
                add_contact(engine, {
                    "Categoria": categoria,
                    "Subcategoria": subcategoria,
                    "Comisión": comision,
                    "Perfil": perfil,
                    "Entidad": entidad,
                    "Teléfono": telefono,
                    "Persona": persona,
                    "Mail": mail,
                }, user_name)
                st.success("Persona añadida a la Red Local de Salud.")
                st.rerun()

    if contacts.empty:
        return

    with st.expander("Filtros", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        categoria = c1.multiselect("Categoría", sorted([v for v in contacts["Categoria"].unique() if safe_text(v)]))
        comision = c2.multiselect("Comisión", sorted([v for v in contacts["Comisión"].unique() if safe_text(v)]))
        perfil = c3.multiselect("Perfil", sorted([v for v in contacts["Perfil"].unique() if safe_text(v)]))
        texto = c4.text_input("Buscar", placeholder="Persona, entidad, email...")

    filtered = contacts.copy()
    if categoria:
        filtered = filtered[filtered["Categoria"].isin(categoria)]
    if comision:
        filtered = filtered[filtered["Comisión"].isin(comision)]
    if perfil:
        filtered = filtered[filtered["Perfil"].isin(perfil)]
    if texto:
        haystack = filtered[["Persona", "Entidad", "Mail", "Teléfono", "Comisión", "Categoria", "Subcategoria"]].fillna("").agg(" ".join, axis=1).str.lower()
        filtered = filtered[haystack.str.contains(texto.lower(), regex=False)]

    if not assignments.empty:
        assigned_meta = assignments.merge(data[["id_accion", "Título", "Ámbito"]], on="id_accion", how="left")
        assigned_text = assigned_meta.groupby("contacto_id").agg(
            acciones_asignadas=("Título", lambda values: "; ".join([safe_text(v) for v in values if safe_text(v)])),
            num_acciones=("id_accion", "nunique"),
        ).reset_index()
        filtered = filtered.merge(assigned_text, on="contacto_id", how="left")
    else:
        filtered["acciones_asignadas"] = ""
        filtered["num_acciones"] = 0
    filtered["acciones_asignadas"] = filtered["acciones_asignadas"].fillna("")
    filtered["num_acciones"] = filtered["num_acciones"].fillna(0).astype(int)

    st.markdown("### Listado editable")
    st.caption("Edita los campos directamente en la tabla y pulsa Guardar cambios. El ID, las acciones asignadas y el número de acciones son campos de consulta.")
    editable_cols = ["contacto_id", *CONTACT_COLUMNS, "num_acciones", "acciones_asignadas"]
    edited = st.data_editor(
        filtered[editable_cols],
        use_container_width=True,
        hide_index=True,
        disabled=["contacto_id", "num_acciones", "acciones_asignadas"],
        key="contacts_editor",
        column_config={
            "contacto_id": st.column_config.NumberColumn("ID", width="small"),
            "acciones_asignadas": st.column_config.TextColumn("Acciones asignadas", width="large"),
            "num_acciones": st.column_config.NumberColumn("Nº acciones", width="small"),
        },
    )

    b1, b2, b3 = st.columns([1.15, 1.15, 4])
    with b1:
        if st.button("Guardar cambios", key="save_contacts_changes"):
            if not user_name:
                st.error("Antes de guardar cambios, escribe tu nombre en la barra lateral.")
            else:
                update_contacts(engine, edited[["contacto_id", *CONTACT_COLUMNS]], user_name)
                st.success("Cambios guardados en la Red Local de Salud.")
                st.rerun()
    with b2:
        with st.popover("Eliminar persona"):
            delete_options = filtered["contacto_id"].astype(int).tolist()
            label_by_id = dict(zip(filtered["contacto_id"], filtered["contacto_label"]))
            selected_delete = st.selectbox(
                "Selecciona la persona a eliminar",
                delete_options,
                format_func=lambda cid: label_by_id.get(cid, str(cid)),
                key="delete_contact_select",
            )
            st.warning("Al eliminarla también se borrarán sus asignaciones a acciones.")
            if st.button("Confirmar eliminación", key="confirm_delete_contact"):
                delete_contact(engine, int(selected_delete))
                st.success("Persona eliminada.")
                st.rerun()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        filtered[editable_cols].to_excel(writer, index=False, sheet_name="Red Local Salud")
    st.download_button(
        "Descargar contactos filtrados en Excel",
        buffer.getvalue(),
        "red_local_salud_contactos.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def render_evolution(df: pd.DataFrame, all_data: pd.DataFrame, history: pd.DataFrame) -> None:
    st.subheader("Registro de modificaciones")
    st.caption("Tabla simplificada con quién modificó cada acción, cuándo lo hizo y qué campos cambió.")

    if history.empty:
        st.info("Todavía no hay modificaciones registradas. Cada vez que guardes una ficha de evaluación se añadirá una línea aquí.")
        return

    meta_columns = ["id_accion", "Ámbito", "Título", "Tipo", "Estado"]
    history_meta = history.merge(all_data[meta_columns], on="id_accion", how="left")
    history_meta = history_meta[history_meta["id_accion"].isin(set(df["id_accion"].astype(int)))].copy()
    if history_meta.empty:
        st.info("No hay modificaciones para los filtros seleccionados.")
        return

    def parse_bool(value: Any) -> bool:
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in {"true", "1", "sí", "si", "yes", "y"}

    def display_value(field: str, value: Any) -> str:
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        if field in ["link_promotora_enviado", "link_participante_enviado"]:
            return "Sí" if parse_bool(value) else "No"
        text_value = safe_text(value)
        if len(text_value) > 160:
            return text_value[:157] + "..."
        return text_value

    tracked_fields = [
        ("estado_evaluacion", "Estado de evaluación"),
        ("cumplimiento_indicadores", "Cumplimiento de indicadores"),
        ("valoracion_tecnica", "Valoración técnica"),
        ("observaciones", "Observaciones"),
        ("responsable_seguimiento", "Responsable seguimiento"),
        ("fecha_actualizacion", "Fecha de evaluación"),
        ("evidencias", "Evidencias / enlaces"),
        ("riesgos", "Riesgos / bloqueos"),
        ("proximos_pasos", "Próximos pasos"),
        ("prioridad", "Prioridad"),
        ("link_promotora", "Link Persona Promotora"),
        ("link_promotora_enviado", "Check envío Persona Promotora"),
        ("link_promotora_fecha", "Fecha envío Persona Promotora"),
        ("link_participante", "Link Persona participante"),
        ("link_participante_enviado", "Check envío Persona participante"),
        ("link_participante_fecha", "Fecha envío Persona participante"),
    ]
    tracked_fields = [(field, label) for field, label in tracked_fields if field in history_meta.columns]

    sort_cols = [c for c in ["id_accion", "updated_at_dt", "updated_at", "id"] if c in history_meta.columns]
    history_meta = history_meta.sort_values(sort_cols).copy()

    rows = []
    for _, group in history_meta.groupby("id_accion", sort=False):
        previous = None
        for _, item in group.iterrows():
            changed_labels = []
            changed_details = []
            if previous is None:
                for field, label in tracked_fields:
                    value = display_value(field, item.get(field))
                    if value:
                        changed_labels.append(label)
                        changed_details.append(f"{label}: {value}")
                if not changed_labels:
                    changed_labels = ["Registro inicial"]
                    changed_details = ["Registro inicial sin campos informados."]
            else:
                for field, label in tracked_fields:
                    old_value = display_value(field, previous.get(field))
                    new_value = display_value(field, item.get(field))
                    if old_value != new_value:
                        changed_labels.append(label)
                        changed_details.append(f"{label}: {old_value or 'vacío'} → {new_value or 'vacío'}")
                if not changed_labels:
                    changed_labels = ["Guardado sin cambios visibles"]
                    changed_details = ["Se guardó la ficha, pero no se detectaron cambios en los campos principales."]

            updated_when = item.get("updated_at_dt")
            if pd.isna(updated_when) if updated_when is not None else True:
                updated_when = pd.to_datetime(item.get("updated_at"), errors="coerce")

            rows.append({
                "Cuándo": updated_when,
                "Quién": safe_text(item.get("updated_by")) or "Sin identificar",
                "Actividad": safe_text(item.get("Título")) or f"Acción {int(item.get('id_accion'))}",
                "Ámbito": safe_text(item.get("Ámbito")),
                "Qué ha modificado": "; ".join(changed_labels),
                "Detalle del cambio": "\n".join(changed_details),
            })
            previous = item

    changes = pd.DataFrame(rows)
    if changes.empty:
        st.info("No hay modificaciones para mostrar.")
        return

    changes = changes.sort_values("Cuándo", ascending=False).copy()
    changes["Fecha"] = changes["Cuándo"].dt.strftime("%d/%m/%Y %H:%M").fillna("")

    with st.expander("Filtros", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            users = sorted([u for u in changes["Quién"].dropna().unique().tolist() if safe_text(u)])
            selected_users = st.multiselect("Quién", users, default=users)
        with col2:
            scopes = sorted([a for a in changes["Ámbito"].dropna().unique().tolist() if safe_text(a)])
            selected_scopes = st.multiselect("Ámbito", scopes, default=scopes)
        with col3:
            text_filter = st.text_input("Buscar", placeholder="Actividad, campo o detalle...")

    filtered_changes = changes.copy()
    if selected_users:
        filtered_changes = filtered_changes[filtered_changes["Quién"].isin(selected_users)]
    if selected_scopes:
        filtered_changes = filtered_changes[filtered_changes["Ámbito"].isin(selected_scopes)]
    if text_filter:
        haystack = filtered_changes[["Actividad", "Qué ha modificado", "Detalle del cambio", "Quién", "Ámbito"]].fillna("").agg(" ".join, axis=1).str.lower()
        filtered_changes = filtered_changes[haystack.str.contains(text_filter.lower(), regex=False)]

    st.metric("Modificaciones registradas", len(filtered_changes))
    st.dataframe(
        filtered_changes[["Quién", "Actividad", "Fecha", "Qué ha modificado", "Detalle del cambio"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Actividad": st.column_config.TextColumn("Qué actividad", width="large"),
            "Detalle del cambio": st.column_config.TextColumn("Qué ha modificado", width="large"),
            "Fecha": st.column_config.TextColumn("Cuándo", width="medium"),
        },
    )

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        filtered_changes[["Quién", "Actividad", "Ámbito", "Fecha", "Qué ha modificado", "Detalle del cambio"]].to_excel(writer, index=False, sheet_name="Registro cambios")
    st.download_button(
        "Descargar registro de modificaciones",
        buffer.getvalue(),
        "registro_modificaciones.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_questionnaire_link_card(title: str, url: str, description: str = "") -> None:
    """Renderiza una tarjeta limpia para abrir y copiar enlaces de cuestionario."""
    clean_title = safe_text(title)
    clean_url = safe_text(url)
    clean_description = safe_text(description)
    if not clean_url:
        st.info("Todavía no hay enlace disponible.")
        return

    st.markdown(
        f"""
        <div class="questionnaire-card">
            <div class="questionnaire-card-title">{clean_title}</div>
            <div class="questionnaire-card-description">{clean_description}</div>
            <a class="questionnaire-card-button" href="{clean_url}" target="_blank" rel="noopener noreferrer">Abrir cuestionario</a>
        </div>
        """,
        unsafe_allow_html=True,
    )
    safe_key = re.sub(r"[^a-zA-Z0-9_]+", "_", clean_title.lower())
    with st.expander(f"Copiar enlace - {clean_title}", expanded=False):
        st.text_input(
            "Enlace",
            value=clean_url,
            key=f"copy_link_{safe_key}_{abs(hash(clean_url))}",
            label_visibility="collapsed",
        )

def render_action_detail(df: pd.DataFrame, all_data: pd.DataFrame, user_name: str, engine: Engine) -> None:
    if df.empty:
        st.info("No hay acciones para los filtros seleccionados.")
        return
    options = df.apply(lambda row: f"{int(row['id_accion']):03d} - {row['Título']}", axis=1).tolist()
    selected = st.selectbox("Selecciona una acción", options)
    selected_id = int(selected.split(" - ")[0])
    row = all_data[all_data["id_accion"] == selected_id].iloc[0]

    st.subheader(row["Título"])
    meta1, meta2, meta3 = st.columns(3)
    meta1.write(f"**Ámbito:** {row['Ámbito']}")
    meta2.write(f"**Tipo:** {row['Tipo']}")
    meta3.write(f"**Estado original:** {row['Estado']}")
    personas_asignadas = safe_text(row.get("personas_red_asignadas"))
    if personas_asignadas:
        st.info(f"Personas de la Red Local asignadas: {personas_asignadas}")

    with st.expander("Información de la acción", expanded=True):
        st.write(f"**Medida:** {row.get('Medida', '')}")
        st.write(f"**Agente promotor:** {row.get('Agente promotor', '')}")
        st.write(f"**Personas destinatarias:** {row.get('Personas destinatarias', '')}")
        st.write(f"**Temporalidad:** {row.get('Temporalidad', '')}")
        st.write("**Descripción:**")
        st.write(row.get("Descripción", ""))
        st.write("**Indicadores previstos en la matriz:**")
        indicadores_matriz = safe_text(row.get("Indicadores", ""))
        st.info(indicadores_matriz or "Esta acción no tiene indicadores definidos en la matriz.")
        st.write("**Recursos:**")
        st.write(row.get("Recursos", ""))

    current_estado = safe_text(row.get("estado_evaluacion")) or "Sin evaluar"
    current_cumplimiento = safe_text(row.get("cumplimiento_indicadores")) or "Sin evaluar"
    current_priority = safe_text(row.get("prioridad")) or "Media"

    with st.form("evaluation_form"):
        st.markdown("### Evaluación")
        estado_evaluacion = st.selectbox("Estado de evaluación", STATUS_OPTIONS, index=STATUS_OPTIONS.index(current_estado) if current_estado in STATUS_OPTIONS else 0)
        st.markdown("#### Cumplimiento de indicadores")
        st.caption("Indicadores definidos en la matriz")
        st.info(indicadores_matriz or "Sin indicadores definidos en la matriz.")
        cumplimiento = st.selectbox("Valoración del cumplimiento", INDICATOR_OPTIONS, index=INDICATOR_OPTIONS.index(current_cumplimiento) if current_cumplimiento in INDICATOR_OPTIONS else 0)

        prioridad = st.selectbox("Prioridad", PRIORITY_OPTIONS, index=PRIORITY_OPTIONS.index(current_priority) if current_priority in PRIORITY_OPTIONS else 1)
        responsable = st.text_input("Responsable de seguimiento", value=safe_text(row.get("responsable_seguimiento")))
        fecha_value = safe_text(row.get("fecha_actualizacion"))
        try:
            parsed_date = datetime.fromisoformat(fecha_value).date() if fecha_value else date.today()
        except ValueError:
            parsed_date = date.today()
        fecha_actualizacion = st.date_input("Fecha de actualización", value=parsed_date)
        valoracion = st.text_area("Valoración técnica", value=safe_text(row.get("valoracion_tecnica")), height=100)
        observaciones = st.text_area("Observaciones", value=safe_text(row.get("observaciones")), height=100)
        riesgos = st.text_area("Riesgos / bloqueos", value=safe_text(row.get("riesgos")), height=80)
        proximos_pasos = st.text_area("Próximos pasos", value=safe_text(row.get("proximos_pasos")), height=80)

        st.markdown("### Links de evaluación")
        st.caption("La aplicación genera enlaces propios de cuestionario. Usa los botones para abrirlos y el desplegable para copiar el enlace cuando tengas que enviarlo.")
        generated_promotora = build_questionnaire_link(selected_id, "promotora")
        generated_participante = build_questionnaire_link(selected_id, "participante")
        current_link_promotora = safe_text(row.get("link_promotora")) or generated_promotora
        current_link_participante = safe_text(row.get("link_participante")) or generated_participante

        l1, l2 = st.columns(2)
        with l1:
            render_questionnaire_link_card(
                "Persona Promotora",
                current_link_promotora,
                "Formulario para la entidad o persona que promueve o coordina la actividad.",
            )
            link_promotora = st.text_input(
                "Enlace guardado - Persona Promotora",
                value=current_link_promotora,
                help="Puedes sustituirlo por un enlace externo si alguna acción usa otro formulario.",
            )
            link_promotora_enviado = st.checkbox("Enviado a Persona Promotora", value=bool(row.get("link_promotora_enviado") or False))
            promotora_fecha_value = safe_text(row.get("link_promotora_fecha"))
            try:
                promotora_date = datetime.fromisoformat(promotora_fecha_value).date() if promotora_fecha_value else date.today()
            except ValueError:
                promotora_date = date.today()
            link_promotora_fecha = st.date_input("Fecha de envío a Persona Promotora", value=promotora_date)
        with l2:
            render_questionnaire_link_card(
                "Persona participante",
                current_link_participante,
                "Formulario para las personas que han participado o se han apuntado a la actividad.",
            )
            link_participante = st.text_input(
                "Enlace guardado - Persona participante",
                value=current_link_participante,
                help="Puedes sustituirlo por un enlace externo si alguna acción usa otro formulario.",
            )
            link_participante_enviado = st.checkbox("Enviado a Persona participante", value=bool(row.get("link_participante_enviado") or False))
            participante_fecha_value = safe_text(row.get("link_participante_fecha"))
            try:
                participante_date = datetime.fromisoformat(participante_fecha_value).date() if participante_fecha_value else date.today()
            except ValueError:
                participante_date = date.today()
            link_participante_fecha = st.date_input("Fecha de envío a Persona participante", value=participante_date)

        submitted = st.form_submit_button("Guardar evaluación")

    if submitted:
        if not user_name:
            st.error("Antes de guardar, escribe tu nombre en la barra lateral.")
            return
        save_evaluation(engine, {
            "id_accion": selected_id,
            "avance": state_to_legacy_progress(estado_evaluacion),
            "estado_evaluacion": estado_evaluacion,
            "cumplimiento_indicadores": cumplimiento,
            "valoracion_tecnica": valoracion,
            "observaciones": observaciones,
            "responsable_seguimiento": responsable,
            "fecha_actualizacion": fecha_actualizacion.isoformat(),
            "evidencias": "",
            "riesgos": riesgos,
            "proximos_pasos": proximos_pasos,
            "prioridad": prioridad,
            "link_promotora": link_promotora,
            "link_promotora_enviado": int(link_promotora_enviado),
            "link_promotora_fecha": link_promotora_fecha.isoformat() if link_promotora_enviado else "",
            "link_participante": link_participante,
            "link_participante_enviado": int(link_participante_enviado),
            "link_participante_fecha": link_participante_fecha.isoformat() if link_participante_enviado else "",
            "updated_by": user_name,
        })
        st.success("Evaluación guardada y añadida al histórico.")
        st.rerun()

    st.markdown("### Repositorio documental de la actividad")
    st.caption("Sube imágenes, actas, documentos o evidencias vinculadas a esta acción. Los archivos se guardan en la base de datos de la aplicación.")
    uploaded_docs = st.file_uploader(
        "Añadir imágenes o documentos",
        type=["pdf", "doc", "docx", "xls", "xlsx", "png", "jpg", "jpeg", "webp", "txt"],
        accept_multiple_files=True,
        key=f"activity_docs_{selected_id}",
    )
    doc_notes = st.text_input("Nota para los archivos", key=f"activity_docs_notes_{selected_id}")
    if st.button("Guardar documentos", key=f"save_activity_docs_{selected_id}"):
        if not user_name:
            st.error("Antes de subir documentos, escribe tu nombre en la barra lateral.")
        elif not uploaded_docs:
            st.error("Selecciona al menos un archivo.")
        else:
            inserted = save_activity_documents(engine, selected_id, uploaded_docs, doc_notes, user_name)
            st.success(f"Documentos guardados: {inserted}")
            st.rerun()

    docs = get_activity_documents(engine, selected_id)
    if docs.empty:
        st.info("Todavía no hay documentos o imágenes vinculados a esta acción.")
    else:
        st.write(f"Documentos guardados: {len(docs)}")
        for _, doc in docs.iterrows():
            with st.container(border=True):
                c1, c2, c3 = st.columns([4, 2, 1])
                c1.write(f"**{safe_text(doc.get('filename'))}**")
                c1.caption(f"Subido por {safe_text(doc.get('uploaded_by')) or 'sin identificar'} · {safe_text(doc.get('uploaded_at'))}")
                if safe_text(doc.get("notes")):
                    c1.write(safe_text(doc.get("notes")))
                c2.write(f"{int(doc.get('size_bytes') or 0) / 1024:.1f} KB")
                raw = base64.b64decode(safe_text(doc.get("content_base64")) or "")
                c2.download_button(
                    "Descargar",
                    data=raw,
                    file_name=safe_text(doc.get("filename")) or "documento",
                    mime=safe_text(doc.get("mime_type")) or "application/octet-stream",
                    key=f"download_doc_{int(doc['document_id'])}",
                )
                with c3:
                    if st.button("Eliminar", key=f"delete_doc_{int(doc['document_id'])}"):
                        delete_activity_document(engine, int(doc["document_id"]))
                        st.success("Documento eliminado.")
                        st.rerun()


    if st.session_state.get("role") == "admin":
        render_questionnaire_results(engine, selected_id)


@st.cache_data(ttl=120, show_spinner=False)
def get_task_email_reminders(_engine: Engine) -> pd.DataFrame:
    engine = _engine
    try:
        with engine.begin() as conn:
            df = pd.read_sql_query(text("SELECT * FROM task_email_reminders ORDER BY sent_at DESC"), conn)
    except Exception:
        return pd.DataFrame(columns=[
            "reminder_key", "task_source", "task_identifier", "id_accion", "accion_titulo", "proximos_pasos",
            "responsable_seguimiento", "fecha_vencimiento", "recipient", "sent_at", "status", "error_message",
        ])
    return clean_interface_dataframe(df)


def render_task_reminder_admin(data: pd.DataFrame, engine: Engine) -> None:
    st.markdown("### Recordatorios de tareas vencidas")
    cfg = get_task_mail_config()
    if not cfg["enabled"]:
        st.warning("Los recordatorios están desactivados por TASK_REMINDER_ENABLED.")
    elif smtp_is_configured():
        st.success(f"Correo configurado. Destinatario de avisos: {cfg['recipient']}")
    else:
        st.warning("Correo no configurado. Añade SMTP_HOST, SMTP_USER, SMTP_PASSWORD, MAIL_FROM y TASK_REMINDER_TO en Secrets.")

    overdue = build_overdue_task_candidates(data, engine)
    sent_keys = get_sent_task_reminder_keys(engine)
    pending = [item for item in overdue if item["reminder_key"] not in sent_keys]
    c1, c2, c3 = st.columns(3)
    c1.metric("Tareas vencidas activas", len(overdue))
    c2.metric("Avisos pendientes", len(pending))
    c3.metric("Avisos ya enviados", len(sent_keys))

    if pending:
        st.dataframe(pd.DataFrame(pending)[["accion_titulo", "proximos_pasos", "responsable_seguimiento", "fecha_vencimiento", "estado"]], use_container_width=True, hide_index=True)
    if st.button("Comprobar y enviar avisos pendientes", key="send_overdue_task_reminders_admin"):
        sent, errors = send_overdue_task_reminders(engine, data, force=False)
        if sent:
            st.success(f"Avisos enviados: {sent}")
        if errors:
            st.error("No se han podido enviar algunos avisos:\n" + "\n".join(errors[:5]))
        if not sent and not errors:
            st.info("No había avisos pendientes de envío.")
        st.rerun()

    reminders = get_task_email_reminders(engine)
    if not reminders.empty:
        st.markdown("#### Últimos avisos registrados")
        st.dataframe(reminders.head(30), use_container_width=True, hide_index=True)

def render_admin(data: pd.DataFrame, history: pd.DataFrame, engine: Engine) -> None:
    st.subheader("Administración")
    st.write("Estado de conexión:", "PostgreSQL/Supabase" if is_postgres(engine) else "SQLite local")
    st.write("Acciones cargadas:", len(data))
    st.write("Mediciones históricas:", len(history))
    st.write("Perfiles activos:", ", ".join(PROFILE_CONFIG.keys()))
    access_log = get_access_log(engine)
    st.markdown("### Últimos accesos")
    if access_log.empty:
        st.info("Todavía no hay accesos registrados.")
    else:
        st.dataframe(clean_interface_dataframe(access_log[["perfil", "role", "scope", "logged_at"]].head(50)), use_container_width=True, hide_index=True)
    render_task_reminder_admin(data, engine)
    render_questionnaire_results(engine)
    st.info("Para uso compartido real, configura DATABASE_URL con Supabase/PostgreSQL en los Secrets de Streamlit Cloud. Las contraseñas de perfiles pueden cambiarse desde Secrets.")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Estado actual")
        history.to_excel(writer, index=False, sheet_name="Historico")
        get_access_log(engine).to_excel(writer, index=False, sheet_name="Accesos")
        get_questionnaire_responses(engine).to_excel(writer, index=False, sheet_name="Cuestionarios")
        get_task_email_reminders(engine).to_excel(writer, index=False, sheet_name="Avisos tareas")
    st.download_button("Descargar copia completa", buffer.getvalue(), "copia_evaluacion_completa.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")



def render_navigation(available_pages: list[str]) -> str:
    """Navegación rápida.

    Se usa radio horizontal en lugar de st.tabs porque st.tabs ejecuta el contenido
    de todas las pestañas en cada interacción. Esta navegación solo calcula la
    página activa, por lo que reduce consultas a Supabase y acelera la app.
    """
    st.markdown("""
    <div class="nav-helper">Selecciona el apartado de trabajo</div>
    """, unsafe_allow_html=True)
    return st.radio(
        "Apartado",
        available_pages,
        horizontal=True,
        label_visibility="collapsed",
        key="main_navigation",
    )


@st.cache_data(ttl=300, show_spinner=False)
def load_contacts_and_assignments(_engine: Engine) -> tuple[pd.DataFrame, pd.DataFrame]:
    engine = _engine
    seed_contacts_if_empty(engine)
    contacts = clean_interface_dataframe(get_contacts(engine))
    assignments = clean_interface_dataframe(get_action_contacts(engine))
    return contacts, assignments


def main() -> None:
    page_icon = str(LOGO_PATH) if LOGO_PATH.exists() else None
    st.set_page_config(page_title="Evaluación y seguimiento", page_icon=page_icon, layout="wide")
    apply_brand_theme()
    st.markdown("""
    <style>
    .nav-helper {
        color: #1C3054;
        font-size: 0.92rem;
        font-weight: 700;
        margin-top: 0.25rem;
        margin-bottom: 0.4rem;
    }
    div[role="radiogroup"] {
        gap: 0.65rem !important;
        flex-wrap: wrap !important;
        margin-bottom: 1.1rem !important;
    }
    div[role="radiogroup"] label {
        background: #FFFFFF !important;
        border: 1px solid rgba(28, 48, 84, 0.20) !important;
        border-radius: 999px !important;
        padding: 0.52rem 0.86rem !important;
        margin-right: 0.2rem !important;
        box-shadow: 0 1px 5px rgba(28, 48, 84, 0.06) !important;
    }
    div[role="radiogroup"] label:hover {
        border-color: #32A4CF !important;
        background: rgba(50, 164, 207, 0.08) !important;
    }
    div[role="radiogroup"] label span {
        color: #1C3054 !important;
        font-weight: 650 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    engine = get_engine()
    if render_public_questionnaire(engine):
        return
    if not authenticate(engine):
        return
    init_db(engine)

    render_sidebar_logo()
    user_name = sidebar_user()
    render_internal_header()

    matrix = load_default_matrix()
    evaluations = clean_interface_dataframe(get_evaluations(engine))
    data = merge_matrix_evaluations(matrix, evaluations)
    data = clean_interface_dataframe(data)
    data["avance"] = data.get("avance", pd.Series([0] * len(data))).fillna(0).astype(int)
    data["estado_evaluacion"] = data["estado_evaluacion"].fillna("Sin evaluar")
    data["cumplimiento_indicadores"] = data["cumplimiento_indicadores"].fillna("Sin evaluar")
    data["prioridad"] = data["prioridad"].fillna("Media")
    for optional_col, default_value in {
        "link_promotora": "", "link_promotora_enviado": 0, "link_promotora_fecha": "",
        "link_participante": "", "link_participante_enviado": 0, "link_participante_fecha": "",
    }.items():
        if optional_col not in data.columns:
            data[optional_col] = default_value
        data[optional_col] = data[optional_col].fillna(default_value)

    data = apply_profile_scope(data)
    maybe_send_overdue_task_reminders(engine, data)
    filtered = filter_dataframe(data)

    if st.session_state.get("role") == "admin":
        available_pages = ["Panel general", "Matriz", "Ficha de evaluación", "Tareas", "Calendario", "Red Local de Salud", "Evolución", "Administración"]
    else:
        available_pages = ["Panel general", "Matriz", "Ficha de evaluación", "Tareas", "Calendario", "Red Local de Salud"]

    selected_page = render_navigation(available_pages)

    if selected_page == "Panel general":
        render_dashboard(filtered)

    elif selected_page == "Matriz":
        with st.spinner("Cargando asignaciones y contactos..."):
            contacts, assignments = load_contacts_and_assignments(engine)
            data_with_assignments = merge_assignments(data, assignments, contacts)
            data_with_assignments = clean_interface_dataframe(data_with_assignments)
            filtered_with_assignments = data_with_assignments[data_with_assignments["id_accion"].isin(filtered["id_accion"])].copy()
        render_matrix(filtered_with_assignments, data_with_assignments, contacts, assignments, user_name, engine)

    elif selected_page == "Ficha de evaluación":
        render_action_detail(filtered, data, user_name, engine)

    elif selected_page == "Tareas":
        render_tasks(data, user_name, engine)

    elif selected_page == "Calendario":
        render_calendar(data, user_name, engine)

    elif selected_page == "Red Local de Salud":
        with st.spinner("Cargando Red Local de Salud..."):
            contacts, assignments = load_contacts_and_assignments(engine)
            data_with_assignments = merge_assignments(data, assignments, contacts)
            data_with_assignments = clean_interface_dataframe(data_with_assignments)
        render_contacts(contacts, assignments, data_with_assignments, user_name, engine)

    elif selected_page == "Evolución" and st.session_state.get("role") == "admin":
        with st.spinner("Cargando registro de evolución..."):
            history = clean_interface_dataframe(get_history(engine))
        render_evolution(filtered, data, history)

    elif selected_page == "Administración" and st.session_state.get("role") == "admin":
        with st.spinner("Cargando administración..."):
            history = clean_interface_dataframe(get_history(engine))
        render_admin(data, history, engine)


if __name__ == "__main__":
    main()
