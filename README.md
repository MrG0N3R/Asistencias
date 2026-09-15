# Dashboard Checadores

Panel Streamlit de Recursos Humanos con login para sus tres vistas: asistencia de
checador, incidencias verificadas RH y consolidado mensual.

## Conexión central

La base central es **UserActivity**. Las cuentas propias de esta aplicación se
guardan en `public.checadores_users`. El historial compartido usa
`public.app_user_activity` con `app_name = "Dashboard Checadores"`.

Copia `.streamlit/activity_db.example.toml` como `.streamlit/activity_db.toml`
y configura el servidor y la contraseña PostgreSQL. La contraseña de conexión
pertenece al usuario de PostgreSQL, no a una cuenta de la aplicación. Ese archivo
queda excluido de Git. En producción, `127.0.0.1` apunta al equipo o contenedor
que ejecuta Streamlit; utiliza el host real de la base central cuando sea remoto.

## Instalación y usuarios

Desde la carpeta del proyecto, con Python 3.11 o posterior:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe manage_users.py init-db
.\.venv\Scripts\python.exe manage_users.py create-user operador --name "Operador de RH"
.\.venv\Scripts\python.exe -m streamlit run app.py
```

En Linux usa `python3 -m venv .venv` y `.venv/bin/python` para los comandos
siguientes. El cliente PostgreSQL se instala mediante `psycopg2-binary` en ambos
sistemas. Este proyecto procesa archivos Excel y no utiliza Firebird ni su DLL.

`init-db` crea las tablas si no existen, conservando el historial existente.
No hay cuentas o contraseñas predeterminadas. `create-user` pide la contraseña
con entrada oculta y requiere al menos 12 caracteres.

Administración de cuentas:

```powershell
.\.venv\Scripts\python.exe manage_users.py reset-password operador
.\.venv\Scripts\python.exe manage_users.py disable-user operador
.\.venv\Scripts\python.exe manage_users.py enable-user operador
```

Las contraseñas se guardan con PBKDF2-HMAC-SHA256 y sal aleatoria. Cinco intentos
fallidos bloquean la cuenta durante cinco minutos. La sesión vence tras 30 minutos
sin interacción; la caducidad y las desactivaciones se verifican en la siguiente
interacción. Restablecer una contraseña no cierra por sí solo una sesión abierta.

## Registro de accesos

El primer login correcto inserta aplicación, usuario, identificador de sesión,
primer acceso y última actividad. Las interacciones siguientes actualizan esa
misma fila. Cerrar sesión borra el estado del usuario; otro login genera un nuevo
identificador. Las visitas anónimas no se registran como accesos correctos.

El programa aparecerá en el historial al ocurrir el primer acceso real. Si falla
la primera inserción, se reintenta durante la siguiente interacción.

## Pruebas

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Para incluir las pruebas de PostgreSQL con la conexión configurada:

```powershell
$env:RUN_DB_TESTS = "1"
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
Remove-Item Env:RUN_DB_TESTS
```

En Linux: `RUN_DB_TESTS=1 .venv/bin/python -m unittest discover -s tests -v`.
Los usuarios y accesos de prueba se crean dentro de transacciones que se revierten.
Se conservan las pruebas de importación y cálculo de los reportes de Recursos Humanos.

La configuración acepta UTF-8 con o sin BOM. Un rechazo de credenciales por
PostgreSQL se comunica de forma controlada incluso si el cliente Windows recibe
un mensaje de error localizado con otra codificación.
