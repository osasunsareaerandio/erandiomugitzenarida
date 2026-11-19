from __future__ import annotations

import calendar
import hashlib
import hmac
import os
import secrets
import smtplib
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from html import escape
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text

ROOT = Path(__file__).resolve().parent
LOGO = ROOT / 'logo_gazteria.png'
COLORS = {
    'Azul noche': '#15263D', 'Rosa gazteria': '#EA1450', 'Turquesa': '#31C9CE',
    'Amarillo': '#F6B729', 'Verde petróleo': '#125A5D', 'Blanco': '#FFFFFF',
    'Fondo suave': '#F7FAFB'
}
NAV = ['Inicio', 'Calendario', 'Tareas', 'Contactos', 'Programas', 'Contraseñas', 'Diseño']
PROGRAMAS = [
    ('Plan Local de Salud 2026-2028', 'Salud', 'Red Local de Salud, mesas y seguimiento de actuaciones'),
    ('Plan de Juventud 2027-2031', 'Juventud', 'Diagnóstico, participación y planificación'),
    ('Gazte Kontseilua', 'Juventud', 'Consejo de juventud, convocatorias y compromisos'),
    ('Umeen Kontseilua', 'Infancia', 'Participación infantil y propuestas municipales'),
    ('Lonjagune', 'Juventud', 'Lonjas juveniles, coordinación y seguimiento'),
]
ESTADOS = ['Pendiente', 'En curso', 'Esperando respuesta', 'Bloqueada', 'Completada']
TIPOS = ['Reunión', 'Actividad', 'Hito', 'Entrega', 'Otro']

st.set_page_config(page_title='Erandio · Mi intranet', page_icon='🗓️', layout='wide', initial_sidebar_state='collapsed')
st.markdown('''<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;500;600;700;800&display=swap');
html, body, .stApp, input, textarea, button, [data-testid="stMarkdownContainer"] {font-family:Nunito,Arial,sans-serif!important}
.stApp {background:#F7FAFB;color:#15263D}
.main .block-container {max-width:1360px;padding-top:1rem}
h1,h2,h3 {color:#15263D!important;letter-spacing:-.02em}
h1 {font-size:1.9rem!important} h2 {font-size:1.35rem!important}
[data-testid="stHeader"]{background:rgba(247,250,251,.97)}
[data-testid="stTabs"] [data-baseweb="tab-list"] {gap:.3rem;flex-wrap:wrap;background:#fff;padding:.45rem;border-radius:17px;border:1px solid #DFE8EC}
[data-testid="stTabs"] button[data-baseweb="tab"] {border-radius:11px!important;padding:.65rem .85rem!important;font-weight:800!important;color:#15263D!important}
[data-testid="stTabs"] button[aria-selected="true"]{background:#EAF9FA!important;color:#125A5D!important;border-bottom:3px solid #31C9CE!important}
[data-testid="stTabs"] button p {font-size:.94rem!important}
[data-testid="stMetric"]{background:white;padding:1rem;border-radius:14px;border:1px solid #DDE8EB;border-left:5px solid #31C9CE;box-shadow:0 3px 14px #15263D08}
[data-testid="stForm"], [data-testid="stExpander"] {background:#fff!important;border-radius:15px!important;border:1px solid #DFE8EC!important}
.stButton button[kind="primary"], [data-testid="stFormSubmitButton"] button[kind="primary"] {background:#125A5D!important;border-color:#125A5D!important;color:#fff!important;border-radius:10px!important}
.stButton button {border-radius:10px!important}
.gaz-banner{background:#fff;border:1px solid #DDE8EB;border-radius:20px;padding:14px 20px;margin-bottom:14px;border-bottom:5px solid #F6B729}
.gaz-overline{font-size:.8rem;color:#125A5D;font-weight:800;letter-spacing:.08em;text-transform:uppercase}
.gaz-note{color:#516276;font-size:.9rem}
.gaz-chip{display:inline-block;padding:5px 10px;border-radius:99px;background:#E7F8F9;color:#125A5D;font-weight:800;font-size:.85rem}
.gaz-swatch{height:48px;border-radius:10px;border:1px solid #dde5ea}
@media (max-width:760px){.main .block-container{padding-left:.7rem;padding-right:.7rem}}
</style>''', unsafe_allow_html=True)


def secret(name: str, default: str = '') -> str:
    try:
        value = st.secrets.get(name, os.environ.get(name, default))
        return str(value) if value is not None else default
    except Exception:
        return os.environ.get(name, default)


def authenticate() -> None:
    expected = secret('APP_PASSWORD')
    if not expected:
        st.error('Falta APP_PASSWORD en Streamlit Secrets. No se permite acceder sin contraseña.')
        st.stop()
    if st.session_state.get('er_authenticated'):
        return
    if LOGO.exists():
        st.image(str(LOGO), width=380)
    st.title('Erandio · Mi intranet')
    st.caption('Acceso a gestión de proyectos, tareas, reuniones y contactos.')
    with st.form('login'):
        password = st.text_input('Contraseña de acceso', type='password')
        submit = st.form_submit_button('Entrar', type='primary')
    if submit:
        if hmac.compare_digest(password.encode(), expected.encode()):
            st.session_state.er_authenticated = True
            st.rerun()
        else:
            st.error('Contraseña incorrecta.')
    st.stop()


authenticate()
url = secret('DATABASE_URL').strip()
if not url:
    st.error('Falta DATABASE_URL en Secrets. No se utilizará almacenamiento temporal para datos personales.')
    st.stop()
if url.startswith('postgres://'):
    url = 'postgresql+psycopg2://' + url[len('postgres://'):]
elif url.startswith('postgresql://'):
    url = 'postgresql+psycopg2://' + url[len('postgresql://'):]

@st.cache_resource
def db_engine():
    return create_engine(url, pool_pre_ping=True, pool_recycle=240, connect_args={'connect_timeout': 10})


def read(sql: str, params: dict | None = None) -> pd.DataFrame:
    with db_engine().connect() as connection:
        return pd.read_sql_query(text(sql), connection, params=params or {})


def write(sql: str, params: dict | None = None) -> None:
    with db_engine().begin() as connection:
        connection.execute(text(sql), params or {})


def init_db() -> None:
    # Migración aditiva: nunca se borran las tablas ni los registros de la versión 1.
    statements = [
        '''CREATE TABLE IF NOT EXISTS er_programas (
            id SERIAL PRIMARY KEY, nombre TEXT NOT NULL UNIQUE, area TEXT, descripcion TEXT,
            estado TEXT NOT NULL DEFAULT 'Activo', notas TEXT NOT NULL DEFAULT '')''',
        '''CREATE TABLE IF NOT EXISTS er_contactos (
            id SERIAL PRIMARY KEY, nombre TEXT NOT NULL, entidad TEXT NOT NULL DEFAULT '',
            cargo TEXT NOT NULL DEFAULT '', email TEXT NOT NULL DEFAULT '', telefono TEXT NOT NULL DEFAULT '',
            programa_id INTEGER REFERENCES er_programas(id) ON DELETE SET NULL,
            notas TEXT NOT NULL DEFAULT '', actualizado TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP)''',
        '''CREATE TABLE IF NOT EXISTS er_tareas (
            id SERIAL PRIMARY KEY, titulo TEXT NOT NULL, descripcion TEXT NOT NULL DEFAULT '',
            programa_id INTEGER REFERENCES er_programas(id) ON DELETE SET NULL,
            contacto_id INTEGER REFERENCES er_contactos(id) ON DELETE SET NULL,
            estado TEXT NOT NULL DEFAULT 'Pendiente', prioridad TEXT NOT NULL DEFAULT 'Media',
            fecha_limite DATE, fecha_aviso DATE, proximo_paso TEXT NOT NULL DEFAULT '',
            actualizado TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP)''',
        '''CREATE TABLE IF NOT EXISTS er_eventos (
            id SERIAL PRIMARY KEY, titulo TEXT NOT NULL, tipo TEXT NOT NULL DEFAULT 'Reunión',
            fecha DATE NOT NULL, hora TIME, programa_id INTEGER REFERENCES er_programas(id) ON DELETE SET NULL,
            contacto_id INTEGER REFERENCES er_contactos(id) ON DELETE SET NULL,
            lugar TEXT NOT NULL DEFAULT '', notas TEXT NOT NULL DEFAULT '')''',
        'ALTER TABLE er_tareas ADD COLUMN IF NOT EXISTS fecha_peticion DATE',
        'ALTER TABLE er_tareas ADD COLUMN IF NOT EXISTS fecha_ejecucion DATE',
        # Mantener sincronizada la fecha antigua con la nueva para preservar plazos existentes.
        'UPDATE er_tareas SET fecha_ejecucion = fecha_limite WHERE fecha_ejecucion IS NULL AND fecha_limite IS NOT NULL',
    ]
    with db_engine().begin() as connection:
        for statement in statements:
            connection.execute(text(statement))
        for name, area, description in PROGRAMAS:
            connection.execute(text('''INSERT INTO er_programas(nombre,area,descripcion) VALUES(:n,:a,:d)
                ON CONFLICT(nombre) DO NOTHING'''), {'n': name, 'a': area, 'd': description})
        # Las tablas pertenecen a la conexión PostgreSQL de la app; nunca a la API pública.
        for table in ('er_programas', 'er_contactos', 'er_tareas', 'er_eventos'):
            connection.execute(text(f'REVOKE ALL PRIVILEGES ON TABLE public.{table} FROM anon, authenticated'))
            connection.execute(text(f'ALTER TABLE public.{table} ENABLE ROW LEVEL SECURITY'))
        for seq in ('er_programas_id_seq', 'er_contactos_id_seq', 'er_tareas_id_seq', 'er_eventos_id_seq'):
            connection.execute(text(f'REVOKE ALL PRIVILEGES ON SEQUENCE public.{seq} FROM anon, authenticated'))

try:
    init_db()
except Exception:
    st.error('No se puede conectar a Supabase o migrar las tablas. Revisa DATABASE_URL, permisos y logs de Streamlit; no se ha borrado ningún registro.')
    st.stop()


def programs() -> pd.DataFrame:
    return read('SELECT * FROM er_programas ORDER BY nombre')


def contacts() -> pd.DataFrame:
    return read('''SELECT c.*, p.nombre AS programa FROM er_contactos c
        LEFT JOIN er_programas p ON p.id=c.programa_id ORDER BY c.nombre''')


def tasks() -> pd.DataFrame:
    return read('''SELECT t.*, p.nombre AS programa FROM er_tareas t
        LEFT JOIN er_programas p ON p.id=t.programa_id
        ORDER BY t.fecha_ejecucion NULLS LAST, t.id DESC''')


def events() -> pd.DataFrame:
    return read('''SELECT e.*, p.nombre AS programa FROM er_eventos e
        LEFT JOIN er_programas p ON p.id=e.programa_id ORDER BY e.fecha,e.hora NULLS LAST''')


def opt_map(df: pd.DataFrame) -> dict[str, int | None]:
    return {'— Sin asignar —': None, **{f"{int(r['id'])} · {r['nombre']}": int(r['id']) for _, r in df.iterrows()}}


def selection(label: str, mapping: dict[str, int | None], *, key: str, current: int | None = None) -> int | None:
    labels = list(mapping)
    index = next((i for i, label_ in enumerate(labels) if mapping[label_] == current), 0)
    return mapping[st.selectbox(label, labels, index=index, key=key)]


def parse_date(value) -> date | None:
    if value is None or pd.isna(value):
        return None
    return pd.to_datetime(value).date()


def table(df: pd.DataFrame, fields: list[str], titles: list[str]) -> None:
    if df.empty:
        st.info('Todavía no hay registros.')
        return
    shown = df[fields].copy()
    shown.columns = titles
    st.dataframe(shown, use_container_width=True, hide_index=True)


pr = programs()
co = contacts()
ta = tasks()
ev = events()
pr_opts = opt_map(pr)
today = date.today()
active = ta[ta['estado'] != 'Completada'].copy() if len(ta) else ta.copy()

with st.container():
    if LOGO.exists():
        c1, c2 = st.columns([2.6, 1], vertical_alignment='center')
        with c1:
            st.image(str(LOGO), width=345)
        with c2:
            st.markdown(f'<div style="text-align:right" class="gaz-note">MI INTRANET · {today:%d/%m/%Y}</div>', unsafe_allow_html=True)
    else:
        st.title('GAZTERIA · ERANDIO')
    st.caption('Tu espacio para organizar programas, tareas, reuniones y contactos.')

tabs = st.tabs(NAV)

with tabs[0]:
    st.title('Inicio')
    st.caption('Todas las tareas por realizar, ordenadas por fecha de ejecución. Las completadas no se muestran.')
    due = active[active.fecha_ejecucion.notna() & (pd.to_datetime(active.fecha_ejecucion).dt.date < today)] if len(active) else active
    soon = active[active.fecha_ejecucion.notna() & (pd.to_datetime(active.fecha_ejecucion).dt.date <= today + timedelta(days=7))] if len(active) else active
    m1, m2, m3 = st.columns(3)
    m1.metric('Tareas pendientes', len(active))
    m2.metric('Fuera de plazo', len(due))
    m3.metric('Próximos 7 días', len(soon))
    st.subheader('Mis tareas')
    if len(active):
        display = active.assign(fecha=active['fecha_ejecucion'].apply(lambda x: parse_date(x).strftime('%d/%m/%Y') if parse_date(x) else 'Sin fecha'))
        table(display, ['titulo', 'programa', 'fecha', 'descripcion'], ['Tarea', 'Programa', 'Fecha', 'Descripción'])
    else:
        st.success('No tienes tareas pendientes. Puedes incorporarlas en la pestaña Tareas.')

with tabs[1]:
    st.title('Calendario')
    st.caption('Las fechas de ejecución de Tareas se incorporan automáticamente al calendario.')
    mcol, dcol = st.columns([1, 1])
    with mcol:
        month = st.date_input('Mes a visualizar', value=today.replace(day=1), key='month_view')
    with dcol:
        selected_day = st.date_input('Agenda del día', value=today, key='selected_day')
    calmap: dict[date, list[tuple[str, str]]] = {}
    for _, row in ev.iterrows():
        d = parse_date(row['fecha'])
        if d:
            calmap.setdefault(d, []).append(('reunion', str(row['titulo'])))
    for _, row in ta.iterrows():
        d = parse_date(row['fecha_ejecucion'])
        if d and row['estado'] != 'Completada':
            calmap.setdefault(d, []).append(('tarea', str(row['titulo'])))
    out = '<div style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:5px">'
    for day_name in ['L', 'M', 'X', 'J', 'V', 'S', 'D']:
        out += f'<div style="text-align:center;font-weight:800;padding:8px">{day_name}</div>'
    for week in calendar.monthcalendar(month.year, month.month):
        for number in week:
            if not number:
                out += '<div style="min-height:100px;background:#F0F4F5;border-radius:10px"></div>'
                continue
            day = date(month.year, month.month, number)
            items = calmap.get(day, [])
            background = '#E4F8F8' if day == today else 'white'
            lines = [f'<div style="font-weight:800;color:#15263D">{number}</div>']
            for typ, label in items[:3]:
                dot = '#EA1450' if typ == 'tarea' else '#125A5D'
                lines.append(f'<div style="font-size:11px;overflow:hidden;text-overflow:ellipsis;color:{dot}">● {escape(label[:28])}</div>')
            if len(items) > 3:
                lines.append(f'<small>+{len(items)-3} más</small>')
            out += f'<div style="min-height:100px;padding:8px;background:{background};border:1px solid #DFE8EC;border-radius:10px;overflow:hidden">{"".join(lines)}</div>'
    st.markdown(out + '</div>', unsafe_allow_html=True)
    st.caption('Rosa: tareas pendientes · Verde petróleo: reuniones y eventos')
    st.subheader(f'Agenda · {selected_day:%d/%m/%Y}')
    day_tasks = ta[ta.fecha_ejecucion.apply(parse_date) == selected_day] if len(ta) else ta
    day_events = ev[ev.fecha.apply(parse_date) == selected_day] if len(ev) else ev
    st.markdown('**Tareas**')
    table(day_tasks, ['titulo', 'programa', 'estado'], ['Tarea', 'Programa', 'Estado'])
    st.markdown('**Reuniones y eventos**')
    table(day_events, ['titulo', 'hora', 'programa', 'lugar'], ['Reunión / evento', 'Hora', 'Programa', 'Lugar'])
    with st.expander('＋ Añadir reunión al calendario', expanded=False):
        with st.form('add_event', clear_on_submit=True):
            title = st.text_input('Título *', key='evt_title')
            c1, c2 = st.columns(2)
            with c1:
                when = st.date_input('Fecha de reunión', value=today, key='evt_date')
                hour = st.time_input('Hora', value=datetime.now().time().replace(second=0, microsecond=0), key='evt_time')
            with c2:
                pid = selection('Programa', pr_opts, key='evt_prog')
                place = st.text_input('Lugar / enlace')
            notes = st.text_area('Descripción / acuerdos')
            if st.form_submit_button('Guardar reunión', type='primary'):
                if not title.strip():
                    st.error('El título es obligatorio.')
                else:
                    write('''INSERT INTO er_eventos (titulo,tipo,fecha,hora,programa_id,lugar,notas)
                        VALUES(:t,'Reunión',:f,:h,:p,:l,:n)''', {'t': title.strip(), 'f': when, 'h': hour, 'p': pid, 'l': place, 'n': notes})
                    st.rerun()
    with st.expander('Eliminar reunión', expanded=False):
        if len(ev):
            event_ids = {f"{int(r['id'])} · {r['titulo']} · {r['fecha']}": int(r['id']) for _, r in ev.iterrows()}
            chosen = st.selectbox('Reunión', list(event_ids), key='delete_evt')
            if st.button('Eliminar reunión seleccionada', key='delete_evt_btn'):
                write('DELETE FROM er_eventos WHERE id=:id', {'id': event_ids[chosen]})
                st.rerun()
        else:
            st.caption('No hay reuniones guardadas.')

with tabs[2]:
    st.title('Tareas')
    st.caption('Registra aquí el trabajo. Todas las tareas pendientes aparecerán en Inicio y en su programa; su fecha de ejecución aparecerá en Calendario.')
    with st.container(border=True):
        st.subheader('Nueva tarea')
        with st.form('new_task', clear_on_submit=True):
            title = st.text_input('Título *', key='task_title')
            pid = selection('Programa', pr_opts, key='task_program')
            desc = st.text_area('Descripción', key='task_desc')
            left, right = st.columns(2)
            with left:
                request_date = st.date_input('Fecha petición', value=today, key='task_request')
            with right:
                execution_date = st.date_input('Fecha ejecución', value=today + timedelta(days=7), key='task_execution')
            if st.form_submit_button('Guardar tarea', type='primary'):
                if not title.strip():
                    st.error('Escribe el título de la tarea.')
                elif pid is None:
                    st.error('Selecciona un programa.')
                else:
                    write('''INSERT INTO er_tareas(titulo,descripcion,programa_id,fecha_peticion,fecha_ejecucion,fecha_limite,estado)
                        VALUES(:t,:d,:p,:fp,:fe,:fe,'Pendiente')''',
                        {'t': title.strip(), 'd': desc, 'p': pid, 'fp': request_date, 'fe': execution_date})
                    st.rerun()
    st.subheader('Listado y seguimiento')
    left, right = st.columns(2)
    with left:
        prog_filter = st.selectbox('Filtrar programa', ['Todos'] + pr.nombre.tolist(), key='tasks_filter_program')
    with right:
        state_filter = st.selectbox('Filtrar estado', ['Pendientes (todos)', 'Todos'] + ESTADOS, key='tasks_filter_state')
    filtered = ta.copy()
    if prog_filter != 'Todos':
        filtered = filtered[filtered['programa'] == prog_filter]
    if state_filter == 'Pendientes (todos)':
        filtered = filtered[filtered['estado'] != 'Completada']
    elif state_filter != 'Todos':
        filtered = filtered[filtered['estado'] == state_filter]
    table(filtered, ['titulo', 'programa', 'descripcion', 'fecha_peticion', 'fecha_ejecucion', 'estado'],
          ['Título', 'Programa', 'Descripción', 'Fecha petición', 'Fecha ejecución', 'Estado'])
    with st.expander('Editar tarea o marcar completada', expanded=False):
        if len(ta):
            task_ids = {f"{int(r['id'])} · {r['titulo']}": int(r['id']) for _, r in ta.iterrows()}
            picked = st.selectbox('Selecciona una tarea', list(task_ids), key='edit_task_select')
            record = ta[ta.id == task_ids[picked]].iloc[0]
            with st.form('edit_task_form'):
                title2 = st.text_input('Título', value=str(record['titulo']))
                desc2 = st.text_area('Descripción', value=str(record['descripcion'] or ''))
                current_program = int(record['programa_id']) if pd.notna(record['programa_id']) else None
                pid2 = selection('Programa', pr_opts, key=f'edit_task_prog_{int(record["id"])}', current=current_program)
                a, b = st.columns(2)
                with a:
                    request2 = st.date_input('Fecha petición', value=parse_date(record['fecha_peticion']) or today)
                with b:
                    execution2 = st.date_input('Fecha ejecución', value=parse_date(record['fecha_ejecucion']) or today)
                current_state = str(record['estado'])
                state2 = st.selectbox('Estado', ESTADOS, index=ESTADOS.index(current_state) if current_state in ESTADOS else 0)
                if st.form_submit_button('Guardar cambios', type='primary'):
                    if title2.strip():
                        write('''UPDATE er_tareas SET titulo=:t,descripcion=:d,programa_id=:p,fecha_peticion=:fp,
                            fecha_ejecucion=:fe,fecha_limite=:fe,estado=:s,actualizado=CURRENT_TIMESTAMP WHERE id=:id''',
                            {'t': title2.strip(), 'd': desc2, 'p': pid2, 'fp': request2, 'fe': execution2, 's': state2, 'id': int(record['id'])})
                        st.rerun()
                    else:
                        st.error('El título no puede estar vacío.')
        else:
            st.caption('Primero añade una tarea.')

with tabs[3]:
    st.title('Contactos')
    st.caption('El listado se muestra primero. Utiliza el buscador y abre el formulario solo cuando necesites añadir una persona.')
    search = st.text_input('Buscar por persona, entidad, cargo o correo', key='search_contacts')
    filtered = co.copy()
    if search.strip() and len(filtered):
        columns = ['nombre', 'entidad', 'cargo', 'email']
        mask = filtered[columns].fillna('').astype(str).apply(lambda col: col.str.contains(search.strip(), case=False, regex=False)).any(axis=1)
        filtered = filtered[mask]
    table(filtered, ['nombre', 'entidad', 'cargo', 'telefono', 'email', 'programa'],
          ['Persona', 'Entidad', 'Cargo', 'Teléfono', 'Correo', 'Programa'])
    with st.expander('＋ Añadir contacto', expanded=False):
        with st.form('new_contact', clear_on_submit=True):
            name = st.text_input('Persona *')
            a, b = st.columns(2)
            with a:
                entity = st.text_input('Entidad')
                position = st.text_input('Cargo / función')
                email = st.text_input('Correo electrónico')
            with b:
                phone = st.text_input('Teléfono')
                pid = selection('Programa principal', pr_opts, key='new_contact_program')
            notes = st.text_area('Notas profesionales (sin datos sensibles innecesarios)')
            if st.form_submit_button('Guardar contacto', type='primary'):
                if not name.strip():
                    st.error('El nombre es obligatorio.')
                else:
                    write('''INSERT INTO er_contactos(nombre,entidad,cargo,email,telefono,programa_id,notas)
                        VALUES(:n,:e,:c,:m,:t,:p,:x)''',
                        {'n': name.strip(), 'e': entity, 'c': position, 'm': email, 't': phone, 'p': pid, 'x': notes})
                    st.rerun()
    with st.expander('Editar contacto', expanded=False):
        if len(co):
            choices = {f"{int(r['id'])} · {r['nombre']}": int(r['id']) for _, r in co.iterrows()}
            chosen = st.selectbox('Contacto', list(choices), key='edit_contact_pick')
            record = co[co.id == choices[chosen]].iloc[0]
            with st.form('edit_contact'):
                name = st.text_input('Persona', value=str(record['nombre'] or ''))
                entity = st.text_input('Entidad', value=str(record['entidad'] or ''))
                position = st.text_input('Cargo', value=str(record['cargo'] or ''))
                email = st.text_input('Correo', value=str(record['email'] or ''))
                phone = st.text_input('Teléfono', value=str(record['telefono'] or ''))
                notes = st.text_area('Notas', value=str(record['notas'] or ''))
                if st.form_submit_button('Guardar cambios', type='primary'):
                    if name.strip():
                        write('''UPDATE er_contactos SET nombre=:n,entidad=:e,cargo=:c,email=:m,telefono=:t,notas=:x,
                            actualizado=CURRENT_TIMESTAMP WHERE id=:id''',
                            {'n': name.strip(), 'e': entity, 'c': position, 'm': email, 't': phone, 'x': notes, 'id': int(record['id'])})
                        st.rerun()
                    else:
                        st.error('El nombre no puede estar vacío.')
        else:
            st.caption('Aún no hay contactos.')

with tabs[4]:
    st.title('Programas')
    st.caption('Selecciona un programa para consultar todas sus tareas pendientes.')
    program_names = pr['nombre'].tolist()
    if program_names:
        selected = st.selectbox('Selecciona un programa', program_names, key='program_panel_selector')
        record = pr[pr['nombre'] == selected].iloc[0]
        st.markdown(f'<span class="gaz-chip">{escape(str(record["area"] or "Proyecto"))}</span>', unsafe_allow_html=True)
        st.write(record['descripcion'])
        program_tasks = ta[(ta['programa_id'] == record['id']) & (ta['estado'] != 'Completada')]
        st.subheader(f'Tareas pendientes · {len(program_tasks)}')
        table(program_tasks, ['titulo', 'descripcion', 'fecha_peticion', 'fecha_ejecucion', 'estado'],
              ['Tarea', 'Descripción', 'Fecha petición', 'Fecha ejecución', 'Estado'])
        program_events = ev[ev['programa_id'] == record['id']]
        with st.expander('Reuniones de este programa', expanded=False):
            table(program_events, ['fecha', 'hora', 'titulo', 'lugar'], ['Fecha', 'Hora', 'Reunión', 'Lugar'])
        with st.expander('Seguimiento del programa', expanded=False):
            with st.form('edit_program'):
                status_options = ['Activo', 'En preparación', 'En pausa', 'Finalizado']
                current = str(record['estado'])
                status = st.selectbox('Situación', status_options, index=status_options.index(current) if current in status_options else 0)
                notes = st.text_area('Notas / decisiones', value=str(record['notas'] or ''))
                if st.form_submit_button('Guardar seguimiento', type='primary'):
                    write('UPDATE er_programas SET estado=:s,notas=:n WHERE id=:id',
                          {'s': status, 'n': notes, 'id': int(record['id'])})
                    st.rerun()
    with st.expander('＋ Añadir otro programa', expanded=False):
        with st.form('new_program', clear_on_submit=True):
            name = st.text_input('Nombre *')
            area = st.text_input('Área')
            desc = st.text_area('Descripción')
            if st.form_submit_button('Crear programa', type='primary'):
                if name.strip():
                    write('''INSERT INTO er_programas(nombre,area,descripcion) VALUES(:n,:a,:d)
                        ON CONFLICT(nombre) DO NOTHING''', {'n': name.strip(), 'a': area, 'd': desc})
                    st.rerun()
                else:
                    st.error('Introduce un nombre.')

with tabs[5]:
    st.title('Contraseñas')
    st.caption('Gestión de accesos de esta intranet y de los servicios asociados.')
    st.info('Por seguridad, las contraseñas reales no se muestran ni se guardan en esta pantalla.')
    st.markdown('**Acceso a la intranet**')
    st.write('Para cambiar la contraseña de entrada, abre Streamlit Community Cloud → tu app → Settings → Secrets y modifica **APP_PASSWORD**. Guarda los cambios y vuelve a iniciar sesión.')
    st.markdown('**Base de datos Supabase**')
    st.write('Para cambiar la contraseña de PostgreSQL, hazlo en Supabase → Intranet Erandio → Connect → Reset database password. Después actualiza **DATABASE_URL** en Streamlit Secrets.')
    st.markdown('**Servicios vinculados**')
    st.markdown('- [GitHub](https://github.com) · código de la intranet\n- [Streamlit](https://share.streamlit.io) · alojamiento y acceso\n- [Supabase](https://supabase.com/dashboard) · base de datos')
    st.warning('No reutilices la misma contraseña en distintos servicios ni publiques archivos de Secrets en GitHub.')
    if st.button('Cerrar mi sesión', key='logout'):
        st.session_state.er_authenticated = False
        st.rerun()

with tabs[6]:
    st.title('Diseño')
    st.caption('Identidad visual de GAZTERIA Erandio, aplicada a toda la intranet.')
    if LOGO.exists():
        st.image(str(LOGO), width=480)
    st.subheader('Tipografía')
    st.markdown('<div style="font-family:Nunito,Arial,sans-serif;font-size:27px;font-weight:800;color:#15263D">Nunito · Titulares y navegación</div>', unsafe_allow_html=True)
    st.write('Nunito se utiliza en títulos, menús, formularios y tablas, con Arial como alternativa.')
    st.subheader('Paleta de colores')
    swatches = [(name, value) for name, value in COLORS.items() if name not in {'Blanco', 'Fondo suave'}]
    for row in (swatches[:3], swatches[3:]):
        cols = st.columns(len(row))
        for col, (name, value) in zip(cols, row):
            with col:
                st.markdown(f'<div class="gaz-swatch" style="background:{value}"></div>', unsafe_allow_html=True)
                st.markdown(f'**{name}**')
                st.code(value, language=None)
    st.caption('Los colores se han tomado como referencia visual del logo suministrado. Puedes ajustar estos códigos y la tipografía en el archivo app.py.')

st.divider()
st.caption('Erandio · Gestión interna de proyectos · La publicación debe ser privada. Antes de incorporar datos personales, verifica el acceso y las políticas de Supabase.')
