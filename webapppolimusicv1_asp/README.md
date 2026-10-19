# PoliMusic v1 - Aplicacion Web ASP Clasico (VBScript)

## Estructura de archivos

```
webapppolimusicv1_asp/
├── conn.asp       - Conexion a SQL Server (funcion GetConnection)
├── default.asp    - Pagina de inicio (redirige segun sesion)
├── login.asp      - Inicio de sesion con ADODB.Command parametrizado
├── songs.asp      - Listado de canciones con reproductor de audio
├── logout.asp     - Cierre de sesion (Session.Abandon)
└── style.css      - Hoja de estilos basica
```

## Requisitos

- IIS habilitado en Windows
- SQL Server local con base de datos `BDD_PoliMusic`
- Usuario SQL: `usr_polimusic` / contraseña: `Politecnica1`
- Proveedor MSOLEDBSQL instalado

## Como desplegar en IIS

1. Copiar la carpeta `webapppolimusicv1_asp` al directorio web de IIS
   (por ejemplo `C:\inetpub\wwwroot\webapppolimusicv1_asp`).

2. Abrir **Administrador de IIS** (IIS Manager).

3. Si se desea usar un puerto especifico (ej. 8080):
   - Click derecho en "Sitios" > "Agregar sitio web"
   - Nombre: PoliMusic
   - Ruta fisica: `C:\inetpub\wwwroot\webapppolimusicv1_asp`
   - Puerto: 8080

4. Asegurarse de que el grupo de aplicaciones este en modo
   "Sin codigo administrado" (Classic) o "Integrado".

5. Habilitar ASP clasico en IIS:
   - Panel de control > Programas > Activar o desactivar caracteristicas
   - Marcar "ASP" bajo Servicios de World Wide Web.

## Como probar

Abrir el navegador y navegar a:

```
http://localhost:8080/login.asp
```

(o `http://localhost:8080/default.asp` que redirige automaticamente)

## Flujo de la aplicacion

1. `default.asp` detecta si hay sesion y redirige a `songs.asp` o `login.asp`.
2. `login.asp` valida contra `TBL_USER` con parametros (anti-inyeccion SQL).
3. `songs.asp` lista las canciones de `TBL_SONG` con reproductores `<audio>`.
4. `logout.asp` cierra la sesion y redirige a `login.asp`.

## Seguridad

- Todas las consultas usan `ADODB.Command` con parametros.
- Todo lo mostrado en pantalla pasa por `Server.HTMLEncode`.
- Los errores de conexion muestran mensajes amigables, no tecnicos.
