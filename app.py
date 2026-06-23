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
DEFAULT_EXCEL = DATA_DIR / "Matriz Acciones Fase II.xlsx"
DEFAULT_SHEET = "MATRIZ Evaluación"
LOCAL_SQLITE = f"sqlite:///{APP_DIR / 'evaluacion.db'}"

BASE_COLUMNS = [
    "Ámbito", "Título", "Tipo", "Estado", "Medida", "Agente promotor",
    "Descripción", "Personas destinatarias", "Temporalidad", "Indicadores", "Recursos",
]

STATUS_OPTIONS = [
    "Sin evaluar", "No iniciado", "Avance inicial", "Avance medio", "Avance alto", "Completado", "No procede",
]
INDICATOR_OPTIONS = ["Sin evaluar", "No iniciado", "Parcial", "Cumplido", "Sobrecumplido", "No procede"]
PRIORITY_OPTIONS = ["Baja", "Media", "Alta", "Crítica"]


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


def authenticate() -> bool:
    app_password = get_secret("APP_PASSWORD", "").strip()
    if not app_password:
        st.warning("No hay clave de acceso configurada. Para publicar la app, define APP_PASSWORD en Secrets.")
        return True
    if st.session_state.get("authenticated"):
        return True
    st.title(get_secret("APP_TITLE", "Evaluacion y seguimiento Emad!"))
    st.caption("Acceso restringido")
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
            st.plotly_chart(px.bar(chart_data, x="Ámbito", y="acciones", text="acciones"), use_container_width=True)
    with right:
        st.subheader("Estado de evaluación")
        chart_data = df.groupby("estado_evaluacion", dropna=False).size().reset_index(name="acciones")
        if not chart_data.empty:
            st.plotly_chart(px.pie(chart_data, names="estado_evaluacion", values="acciones"), use_container_width=True)

    st.subheader("Acciones con menor avance")
    columns = ["id_accion", "Ámbito", "Título", "Estado", "avance", "estado_evaluacion", "responsable_seguimiento", "updated_by"]
    st.dataframe(df.sort_values(["avance", "id_accion"])[columns].head(20), use_container_width=True, hide_index=True)


def render_matrix(df: pd.DataFrame) -> None:
    visible_columns = [
        "id_accion", "Ámbito", "Título", "Tipo", "Estado", "Agente promotor", "Temporalidad", "avance",
        "estado_evaluacion", "cumplimiento_indicadores", "responsable_seguimiento", "fecha_actualizacion", "updated_by", "updated_at",
    ]
    st.dataframe(df[visible_columns], use_container_width=True, hide_index=True)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Evaluacion")
    st.download_button(
        "Descargar evaluación filtrada en Excel", buffer.getvalue(), "evaluacion_filtrada.xlsx",
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
    st.plotly_chart(px.line(global_series, x="fecha_grafico", y="avance", markers=True, labels={"avance": "Avance medio (%)", "fecha_grafico": "Fecha"}), use_container_width=True)

    st.markdown("### Evolución media por ámbito")
    scope_series = history_meta.sort_values(["fecha_grafico", "updated_at_dt", "id"]).groupby(["fecha_grafico", "Ámbito"], as_index=False)["avance"].mean()
    st.plotly_chart(px.line(scope_series, x="fecha_grafico", y="avance", color="Ámbito", markers=True, labels={"avance": "Avance medio (%)", "fecha_grafico": "Fecha"}), use_container_width=True)

    st.markdown("### Evolución de una acción concreta")
    options = df.apply(lambda row: f"{int(row['id_accion']):03d} - {row['Título']}", axis=1).tolist()
    selected = st.selectbox("Selecciona una acción", options, key="history_action")
    selected_id = int(selected.split(" - ")[0])
    action_history = history_meta[history_meta["id_accion"] == selected_id].sort_values(["fecha_grafico", "updated_at_dt", "id"])
    if action_history.empty:
        st.info("Esta acción todavía no tiene mediciones históricas.")
    else:
        st.plotly_chart(px.line(action_history, x="fecha_grafico", y="avance", markers=True, labels={"avance": "Avance (%)", "fecha_grafico": "Fecha"}), use_container_width=True)
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
    st.set_page_config(page_title="Evaluacion y seguimiento", layout="wide")
    if not authenticate():
        return

    engine = get_engine()
    init_db(engine)
    user_name = sidebar_user()

    st.title(get_secret("APP_TITLE", "Evaluacion y seguimiento Emad!"))
    st.caption("Herramienta compartida de evaluación, seguimiento histórico y actualización de acciones")

    with st.sidebar:
        st.header("Datos")
        uploaded = st.file_uploader("Sustituir matriz Excel en esta sesión", type=["xlsx"])
        st.caption("En la versión web compartida se recomienda mantener una matriz base estable en la carpeta data/.")

    matrix = read_excel_bytes(uploaded.getvalue()) if uploaded is not None else load_default_matrix()
    evaluations = get_evaluations(engine)
    history = get_history(engine)
    data = merge_matrix_evaluations(matrix, evaluations)
    data["avance"] = data["avance"].fillna(0).astype(int)
    data["estado_evaluacion"] = data["estado_evaluacion"].fillna("Sin evaluar")
    data["cumplimiento_indicadores"] = data["cumplimiento_indicadores"].fillna("Sin evaluar")
    data["prioridad"] = data["prioridad"].fillna("Media")

    filtered = filter_dataframe(data)
    tabs = st.tabs(["Panel general", "Matriz", "Ficha de evaluación", "Evolución", "Administración"])
    with tabs[0]:
        render_dashboard(filtered)
    with tabs[1]:
        render_matrix(filtered)
    with tabs[2]:
        render_action_detail(filtered, data, user_name, engine)
    with tabs[3]:
        render_evolution(filtered, data, history)
    with tabs[4]:
        render_admin(data, history, engine)


if __name__ == "__main__":
    main()
