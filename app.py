from __future__ import annotations

import io
import os
import re
from datetime import date, datetime
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

STATUS_OPTIONS = [
    "Sin evaluar", "No iniciado", "Avance inicial", "Avance medio", "Avance alto", "Completado", "No procede",
]
INDICATOR_OPTIONS = ["Sin evaluar", "No iniciado", "Parcial", "Cumplido", "Sobrecumplido", "No procede"]
PRIORITY_OPTIONS = ["Baja", "Media", "Alta", "Crítica"]

# Paleta corporativa Erandio Mugitzen ari da!
COLOR_NAVY = "#1C3054"
COLOR_SKY = "#32A4CF"
COLOR_CORAL = "#E95C47"
COLOR_MAUVE = "#AE7CAA"
COLOR_GOLD = "#F9C14F"
BRAND_COLORS = [COLOR_NAVY, COLOR_SKY, COLOR_CORAL, COLOR_MAUVE, COLOR_GOLD]


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


def safe_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalize_column_name(value: Any) -> str:
    return str(value).replace("\n", " ").strip()


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


def init_db(engine: Engine) -> None:
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
    for table_name in ["evaluations", "evaluation_history"]:
        if not column_exists(engine, table_name, "updated_by"):
            with engine.begin() as conn:
                conn.execute(text(f"ALTER TABLE {table_name} ADD COLUMN updated_by TEXT"))


def get_evaluations(engine: Engine) -> pd.DataFrame:
    with engine.begin() as conn:
        df = pd.read_sql_query(text("SELECT * FROM evaluations"), conn)
    if df.empty:
        return pd.DataFrame(columns=[
            "id_accion", "avance", "estado_evaluacion", "cumplimiento_indicadores", "valoracion_tecnica",
            "observaciones", "responsable_seguimiento", "fecha_actualizacion", "evidencias", "riesgos",
            "proximos_pasos", "prioridad", "updated_by", "updated_at",
        ])
    df["id_accion"] = df["id_accion"].astype(int)
    return df


def get_history(engine: Engine) -> pd.DataFrame:
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


def get_action_contacts(engine: Engine) -> pd.DataFrame:
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



def seed_contacts_if_empty(engine: Engine) -> None:
    """Carga Contactos.xlsx en la base de datos una sola vez si la tabla contacts está vacía."""
    with engine.begin() as conn:
        count = conn.execute(text("SELECT COUNT(*) FROM contacts")).scalar() or 0
    if count > 0 or not CONTACTS_EXCEL.exists():
        return
    seed = load_default_contacts()
    if seed.empty:
        return
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as conn:
        for _, row in seed.iterrows():
            payload = {col: safe_text(row.get(col)) for col in CONTACT_COLUMNS}
            payload.update({"updated_by": "carga inicial", "updated_at": now})
            fields = list(payload.keys())
            conn.execute(
                text(f"""
                    INSERT INTO contacts ({", ".join(fields)})
                    VALUES ({", ".join([":" + f for f in fields])})
                """),
                payload,
            )


def get_contacts(engine: Engine) -> pd.DataFrame:
    with engine.begin() as conn:
        df = pd.read_sql_query(text("SELECT * FROM contacts ORDER BY contacto_id ASC"), conn)
    if df.empty:
        return pd.DataFrame(columns=["contacto_id", *CONTACT_COLUMNS, "updated_by", "updated_at", "contacto_label"])
    df["contacto_id"] = df["contacto_id"].astype(int)
    for col in CONTACT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].apply(safe_text)
    df["contacto_label"] = df.apply(make_contact_label, axis=1)
    return df[["contacto_id", *CONTACT_COLUMNS, "updated_by", "updated_at", "contacto_label"]]


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


def delete_contact(engine: Engine, contacto_id: int) -> None:
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM action_contacts WHERE contacto_id = :contacto_id"), {"contacto_id": contacto_id})
        conn.execute(text("DELETE FROM contacts WHERE contacto_id = :contacto_id"), {"contacto_id": contacto_id})


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


def merge_matrix_evaluations(matrix: pd.DataFrame, evaluations: pd.DataFrame) -> pd.DataFrame:
    if evaluations.empty:
        merged = matrix.copy()
        defaults = {
            "avance": 0, "estado_evaluacion": "Sin evaluar", "cumplimiento_indicadores": "Sin evaluar",
            "valoracion_tecnica": "", "observaciones": "", "responsable_seguimiento": "",
            "fecha_actualizacion": "", "evidencias": "", "riesgos": "", "proximos_pasos": "",
            "prioridad": "Media", "updated_by": "", "updated_at": "",
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


def authenticate() -> bool:
    app_password = get_secret("APP_PASSWORD", "").strip()
    if not app_password:
        st.warning("No hay clave de acceso configurada. Para publicar la app, define APP_PASSWORD en Secrets.")
        return True
    if st.session_state.get("authenticated"):
        return True
    render_login_title()
    password = st.text_input("Clave de acceso", type="password")
    if st.button("Entrar"):
        if password == app_password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Clave incorrecta.")
    return False


def sidebar_user() -> str:
    with st.sidebar:
        st.header("Sesión")
        user_name = st.text_input("Nombre de quien actualiza", value=st.session_state.get("user_name", ""))
        st.session_state["user_name"] = user_name
        if st.button("Cerrar sesión"):
            st.session_state.clear()
            st.rerun()
    return user_name.strip()


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
        avance_min, avance_max = st.slider("Avance evaluado (%)", 0, 100, (0, 100))
        texto = st.text_input("Buscar texto", placeholder="Título, descripción, indicadores...")

    filtered = df[
        df["Ámbito"].isin(ambitos)
        & df["Estado"].isin(estados)
        & df["Tipo"].isin(tipos)
        & df["Agente promotor"].isin(promotores)
        & df["avance"].fillna(0).between(avance_min, avance_max)
    ].copy()
    if texto:
        haystack = filtered[["Título", "Descripción", "Indicadores", "Medida"]].fillna("").agg(" ".join, axis=1).str.lower()
        filtered = filtered[haystack.str.contains(texto.lower(), regex=False)]
    return filtered


def render_dashboard(df: pd.DataFrame) -> None:
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Acciones filtradas", len(df))
    col2.metric("Avance medio", f"{df['avance'].fillna(0).mean():.1f}%" if len(df) else "0%")
    col3.metric("Completadas", int((df["estado_evaluacion"] == "Completado").sum()))
    col4.metric("Sin evaluar", int((df["estado_evaluacion"].fillna("Sin evaluar") == "Sin evaluar").sum()))

    left, right = st.columns(2)
    with left:
        st.subheader("Acciones por ámbito")
        chart_data = df.groupby("Ámbito", dropna=False).size().reset_index(name="acciones")
        if not chart_data.empty:
            fig = px.bar(chart_data, x="Ámbito", y="acciones", text="acciones", color_discrete_sequence=BRAND_COLORS)
            st.plotly_chart(apply_plotly_brand_layout(fig), use_container_width=True)
    with right:
        st.subheader("Estado de evaluación")
        chart_data = df.groupby("estado_evaluacion", dropna=False).size().reset_index(name="acciones")
        if not chart_data.empty:
            fig = px.pie(chart_data, names="estado_evaluacion", values="acciones", color_discrete_sequence=BRAND_COLORS)
            st.plotly_chart(apply_plotly_brand_layout(fig), use_container_width=True)

    st.subheader("Acciones con menor avance")
    columns = ["id_accion", "Ámbito", "Título", "Estado", "avance", "estado_evaluacion", "responsable_seguimiento", "updated_by"]
    st.dataframe(df.sort_values(["avance", "id_accion"])[columns].head(20), use_container_width=True, hide_index=True)


def render_matrix(df: pd.DataFrame, all_data: pd.DataFrame, contacts: pd.DataFrame, assignments: pd.DataFrame, user_name: str, engine: Engine) -> None:
    visible_columns = [
        "id_accion", "Ámbito", "Título", "Tipo", "Estado", "Agente promotor", "Temporalidad",
        "personas_red_asignadas", "num_personas_red", "avance",
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
    if contacts.empty:
        st.info("Todavía no hay contactos. Puedes crear el primer registro desde el formulario inferior.")

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
    st.subheader("Evolución temporal")
    if history.empty:
        st.info("Todavía no hay histórico. Cada vez que guardes una evaluación se añadirá una nueva medición con fecha.")
        return
    history_meta = history.merge(all_data[["id_accion", "Ámbito", "Título", "Tipo", "Estado"]], on="id_accion", how="left")
    history_meta = history_meta[history_meta["id_accion"].isin(set(df["id_accion"].astype(int)))].copy()
    if history_meta.empty:
        st.info("No hay datos históricos para los filtros seleccionados.")
        return
    history_meta["fecha_grafico"] = history_meta["fecha_actualizacion_dt"].fillna(history_meta["updated_at_dt"])
    history_meta = history_meta.dropna(subset=["fecha_grafico"])

    col1, col2, col3 = st.columns(3)
    col1.metric("Mediciones históricas", len(history_meta))
    col2.metric("Acciones con histórico", history_meta["id_accion"].nunique())
    col3.metric("Último avance medio", f"{df['avance'].fillna(0).mean():.1f}%" if len(df) else "0%")

    st.markdown("### Evolución global del avance medio")
    global_series = history_meta.sort_values(["fecha_grafico", "updated_at_dt", "id"]).groupby("fecha_grafico", as_index=False)["avance"].mean()
    fig = px.line(global_series, x="fecha_grafico", y="avance", markers=True, labels={"avance": "Avance medio (%)", "fecha_grafico": "Fecha"}, color_discrete_sequence=BRAND_COLORS)
    st.plotly_chart(apply_plotly_brand_layout(fig), use_container_width=True)

    st.markdown("### Evolución media por ámbito")
    scope_series = history_meta.sort_values(["fecha_grafico", "updated_at_dt", "id"]).groupby(["fecha_grafico", "Ámbito"], as_index=False)["avance"].mean()
    fig = px.line(scope_series, x="fecha_grafico", y="avance", color="Ámbito", markers=True, labels={"avance": "Avance medio (%)", "fecha_grafico": "Fecha"}, color_discrete_sequence=BRAND_COLORS)
    st.plotly_chart(apply_plotly_brand_layout(fig), use_container_width=True)

    st.markdown("### Evolución de una acción concreta")
    options = df.apply(lambda row: f"{int(row['id_accion']):03d} - {row['Título']}", axis=1).tolist()
    selected = st.selectbox("Selecciona una acción", options, key="history_action")
    selected_id = int(selected.split(" - ")[0])
    action_history = history_meta[history_meta["id_accion"] == selected_id].sort_values(["fecha_grafico", "updated_at_dt", "id"])
    if action_history.empty:
        st.info("Esta acción todavía no tiene mediciones históricas.")
    else:
        fig = px.line(action_history, x="fecha_grafico", y="avance", markers=True, labels={"avance": "Avance (%)", "fecha_grafico": "Fecha"}, color_discrete_sequence=[COLOR_CORAL])
        st.plotly_chart(apply_plotly_brand_layout(fig), use_container_width=True)
        st.dataframe(action_history[["fecha_actualizacion", "avance", "estado_evaluacion", "cumplimiento_indicadores", "responsable_seguimiento", "updated_by", "valoracion_tecnica", "observaciones", "riesgos", "proximos_pasos", "updated_at"]], use_container_width=True, hide_index=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        history_meta.to_excel(writer, index=False, sheet_name="Historico")
    st.download_button("Descargar histórico en Excel", buffer.getvalue(), "historico_evaluacion.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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
        st.write("**Indicadores previstos:**")
        st.write(row.get("Indicadores", ""))
        st.write("**Recursos:**")
        st.write(row.get("Recursos", ""))

    current_avance = int(row["avance"]) if pd.notna(row.get("avance")) else 0
    current_estado = safe_text(row.get("estado_evaluacion")) or "Sin evaluar"
    current_cumplimiento = safe_text(row.get("cumplimiento_indicadores")) or "Sin evaluar"
    current_priority = safe_text(row.get("prioridad")) or "Media"

    with st.form("evaluation_form"):
        st.markdown("### Evaluación")
        avance = st.slider("Avance (%)", 0, 100, current_avance)
        estado_evaluacion = st.selectbox("Estado de evaluación", STATUS_OPTIONS, index=STATUS_OPTIONS.index(current_estado) if current_estado in STATUS_OPTIONS else 0)
        cumplimiento = st.selectbox("Cumplimiento de indicadores", INDICATOR_OPTIONS, index=INDICATOR_OPTIONS.index(current_cumplimiento) if current_cumplimiento in INDICATOR_OPTIONS else 0)
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
        evidencias = st.text_area("Evidencias / enlaces a documentos", value=safe_text(row.get("evidencias")), height=80, placeholder="Pega enlaces a Drive, actas, fotos, informes o evidencias relevantes")
        submitted = st.form_submit_button("Guardar evaluación")

    if submitted:
        if not user_name:
            st.error("Antes de guardar, escribe tu nombre en la barra lateral.")
            return
        save_evaluation(engine, {
            "id_accion": selected_id,
            "avance": avance,
            "estado_evaluacion": estado_evaluacion,
            "cumplimiento_indicadores": cumplimiento,
            "valoracion_tecnica": valoracion,
            "observaciones": observaciones,
            "responsable_seguimiento": responsable,
            "fecha_actualizacion": fecha_actualizacion.isoformat(),
            "evidencias": evidencias,
            "riesgos": riesgos,
            "proximos_pasos": proximos_pasos,
            "prioridad": prioridad,
            "updated_by": user_name,
        })
        st.success("Evaluación guardada y añadida al histórico.")
        st.rerun()


def render_admin(data: pd.DataFrame, history: pd.DataFrame, engine: Engine) -> None:
    st.subheader("Administración")
    st.write("Estado de conexión:", "PostgreSQL/Supabase" if is_postgres(engine) else "SQLite local")
    st.write("Acciones cargadas:", len(data))
    st.write("Mediciones históricas:", len(history))
    st.info("Para uso compartido real, configura DATABASE_URL con Supabase/PostgreSQL en los Secrets de Streamlit Cloud.")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Estado actual")
        history.to_excel(writer, index=False, sheet_name="Historico")
    st.download_button("Descargar copia completa", buffer.getvalue(), "copia_evaluacion_completa.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main() -> None:
    page_icon = str(LOGO_PATH) if LOGO_PATH.exists() else None
    st.set_page_config(page_title="Evaluación y seguimiento", page_icon=page_icon, layout="wide")
    apply_brand_theme()
    if not authenticate():
        return

    engine = get_engine()
    init_db(engine)
    render_sidebar_logo()
    user_name = sidebar_user()

    render_internal_header()

    with st.sidebar:
        st.header("Datos")
        uploaded = st.file_uploader("Sustituir matriz Excel en esta sesión", type=["xlsx"])
        st.caption("En la versión web compartida se recomienda mantener una matriz base estable en la carpeta data/.")

    matrix = read_excel_bytes(uploaded.getvalue()) if uploaded is not None else load_default_matrix()
    evaluations = get_evaluations(engine)
    history = get_history(engine)
    seed_contacts_if_empty(engine)
    contacts = get_contacts(engine)
    assignments = get_action_contacts(engine)
    data = merge_matrix_evaluations(matrix, evaluations)
    data = merge_assignments(data, assignments, contacts)
    data["avance"] = data["avance"].fillna(0).astype(int)
    data["estado_evaluacion"] = data["estado_evaluacion"].fillna("Sin evaluar")
    data["cumplimiento_indicadores"] = data["cumplimiento_indicadores"].fillna("Sin evaluar")
    data["prioridad"] = data["prioridad"].fillna("Media")

    filtered = filter_dataframe(data)
    tabs = st.tabs(["Panel general", "Matriz", "Ficha de evaluación", "Red Local de Salud", "Evolución", "Administración"])
    with tabs[0]:
        render_dashboard(filtered)
    with tabs[1]:
        render_matrix(filtered, data, contacts, assignments, user_name, engine)
    with tabs[2]:
        render_action_detail(filtered, data, user_name, engine)
    with tabs[3]:
        render_contacts(contacts, assignments, data, user_name, engine)
    with tabs[4]:
        render_evolution(filtered, data, history)
    with tabs[5]:
        render_admin(data, history, engine)


if __name__ == "__main__":
    main()
