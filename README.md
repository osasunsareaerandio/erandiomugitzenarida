# Erandio · Mi intranet — versión 2

Actualización de la versión anterior. Mantiene las cuatro tablas `er_programas`, `er_contactos`, `er_tareas` y `er_eventos` y **no borra datos**. Añade `fecha_peticion` y `fecha_ejecucion` a las tareas; migra las fechas límite anteriores a `fecha_ejecucion` cuando esta es nula.

## Actualizar el repositorio existente

En `osasunsareaerandio/intraneterandio` (privado), sustituye `app.py` por el de esta versión y sube también `logo_gazteria.png` **en la raíz del repositorio**, junto a `app.py`. No subas un ZIP ni ningún secreto. El archivo `requirements.txt` se conserva sin cambios; puedes sustituirlo por el incluido si lo prefieres.

La nueva interfaz incluye pestañas arriba: Inicio, Calendario, Tareas, Contactos, Programas, Contraseñas y Diseño.

## Configuración

Conserva en Streamlit Secrets las variables existentes `APP_PASSWORD` y `DATABASE_URL`. No son necesarios cambios en su valor por actualizar el diseño. La conexión debe usar el proyecto Supabase independiente de la intranet. Utiliza contraseñas distintas y rota cualquier credencial compartida anteriormente.

## Seguridad

- Las tablas tienen RLS habilitado y se revocan los permisos a `anon` y `authenticated` en el arranque. Una conexión PostgreSQL propietaria/privilegiada puede seguir escribiendo. Esta medida no sustituye una revisión de permisos SQL y del hosting.
- Se necesita una conexión PostgreSQL con privilegios suficientes para crear/migrar tablas, habilitar RLS y revocar permisos. Si la base es gestionada por un usuario sin esos privilegios, solicita una migración administrada.
- Mantén Streamlit configurado como **privado**. El inicio de sesión con `APP_PASSWORD` por sí solo no sustituye el control del alojamiento.
- El apartado Contraseñas explica cómo actualizar el acceso; nunca almacena contraseñas en registros de la base de datos.
- Los avisos automáticos de correo y la sincronización con Google Calendar no están implementados en esta versión. Los vencimientos sí aparecen en Inicio y Calendario.

## Comprobación posterior al despliegue

1. Confirma que aparecen las siete pestañas y el logo.
2. Comprueba que los contactos/tareas anteriores siguen visibles.
3. Crea una tarea ficticia, comprueba que aparece en Inicio y en Programas, y que su fecha figura en Calendario.
4. Crea una reunión ficticia en Calendario.
5. Marca la tarea como completada: dejará de aparecer en Inicio y en las tareas pendientes del programa.
6. Revisa en Supabase que `er_contactos`, `er_tareas`, `er_eventos` y `er_programas` tengan RLS activado y permisos de roles públicos revocados.

Si aparece un error al arrancar, consulta los Logs de Streamlit ocultando credenciales antes de compartirlos.
