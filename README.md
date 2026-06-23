# Herramienta compartida de evaluacion y seguimiento

Aplicacion web desarrollada en Streamlit para evaluar y seguir la evolucion de la Matriz de Acciones Fase II.

La version actual funciona de dos maneras:

1. **Modo local**: usa SQLite en el propio ordenador.
2. **Modo compartido/web**: usa una base de datos PostgreSQL/Supabase configurada mediante `DATABASE_URL`.

## Funcionalidades incluidas

- Acceso mediante clave compartida.
- Identificacion de la persona que actualiza cada registro.
- Filtros por ambito, estado, tipo, agente promotor, avance y texto.
- Panel general con indicadores y graficos.
- Matriz consultable y exportable.
- Ficha de evaluacion por accion.
- Historico de mediciones por fecha.
- Graficos de evolucion global, por ambito y por accion.
- Exportacion completa de datos e historico.
- Preparada para publicarse en Streamlit Cloud y conectarse a Supabase/PostgreSQL.

## Ejecucion local en Mac

Desde la carpeta de la aplicacion:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

La aplicacion se abrira normalmente en:

```text
http://localhost:8501
```

## Configuracion local con clave

Copia el archivo de ejemplo:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Edita `.streamlit/secrets.toml` y cambia:

```toml
APP_PASSWORD = "cambiar-esta-clave"
APP_TITLE = "Evaluacion y seguimiento Emad!"
DATABASE_URL = ""
```

Si `DATABASE_URL` esta vacio, la aplicacion usa SQLite local.

## Publicacion compartida recomendada

Arquitectura recomendada:

- Streamlit Community Cloud para publicar la aplicacion.
- GitHub para alojar el codigo.
- Supabase para base de datos compartida PostgreSQL.

### Paso 1. Crear proyecto en Supabase

1. Entrar en Supabase.
2. Crear un nuevo proyecto.
3. Guardar la clave de base de datos.
4. Copiar la cadena de conexion PostgreSQL compatible con SQLAlchemy.

El formato habitual es similar a este:

```text
postgresql+psycopg2://postgres.PROJECT_REF:PASSWORD@aws-0-eu-west-1.pooler.supabase.com:6543/postgres?sslmode=require
```

Sustituir `PROJECT_REF` y `PASSWORD` por los datos reales del proyecto.

### Paso 2. Subir la aplicacion a GitHub

Subir al repositorio estos archivos y carpetas:

```text
app.py
requirements.txt
README.md
data/Matriz Acciones Fase II.xlsx
.streamlit/secrets.toml.example
```

No subir nunca `.streamlit/secrets.toml` si contiene claves reales.

### Paso 3. Desplegar en Streamlit Cloud

1. Entrar en Streamlit Community Cloud.
2. Crear nueva app.
3. Conectar el repositorio de GitHub.
4. Seleccionar `app.py` como archivo principal.
5. En `Settings > Secrets`, pegar algo como:

```toml
APP_PASSWORD = "clave-compartida-segura"
APP_TITLE = "Evaluacion y seguimiento Emad!"
DATABASE_URL = "postgresql+psycopg2://postgres.PROJECT_REF:PASSWORD@aws-0-eu-west-1.pooler.supabase.com:6543/postgres?sslmode=require"
```

6. Guardar y desplegar.

La aplicacion creara automaticamente las tablas necesarias la primera vez que se ejecute.

## Consideraciones importantes

### Base de datos

Para uso compartido real, no usar SQLite. SQLite esta pensado para uso local. En web debe usarse Supabase/PostgreSQL.

### Evidencias documentales

Esta version guarda evidencias como texto o enlaces, por ejemplo enlaces a Google Drive, actas, fotos o informes. No se recomienda guardar ficheros subidos directamente en Streamlit Cloud porque el almacenamiento local puede no ser persistente.

### Seguridad

La version actual usa una clave compartida. Es suficiente para un piloto controlado, pero no sustituye a un sistema completo de usuarios y permisos.

Para una version institucional seria, se recomienda incorporar:

- usuarios individuales;
- roles de lectura, edicion y administracion;
- auditoria de cambios;
- copias de seguridad revisadas;
- politica de evidencias y datos personales.

## Coste tecnico orientativo

- GitHub: 0 euros para empezar.
- Streamlit Community Cloud: 0 euros para empezar.
- Supabase Free: 0 euros para piloto.
- Supabase Pro: aproximadamente 25 USD/mes para uso real.
- Dominio propio opcional: aproximadamente 10-25 euros/anio.

## Estructura del proyecto

```text
evaluacion_app/
├── app.py
├── requirements.txt
├── README.md
├── data/
│   └── Matriz Acciones Fase II.xlsx
└── .streamlit/
    └── secrets.toml.example
```
