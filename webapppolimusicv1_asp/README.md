# PoliMusic · ASP Clásico + SQL Server

Aplicación web de gestión de biblioteca musical desarrollada en **ASP Clásico (VBScript)**
con **Microsoft SQL Server**. Incluye autenticación por sesión, control de acceso por
roles (Administrador / Usuario Estándar), CRUD de canciones y un reproductor de audio
global con contador de reproducciones.

---

## 1. Requisitos

| Componente | Detalle |
|---|---|
| Servidor web | IIS 7+ con el rol **ASP** habilitado (ASP Clásico, no ASP.NET) |
| Base de datos | SQL Server 2012 o superior, base `BDD_PoliMusic` creada |
| Navegador | Chrome / Edge / Firefox actualizado (usa `fetch` y `<audio>`) |
| Cuenta SQL | `sa` / `1234` (autenticación SQL Server) |

> El script `BDD_PoliMusic (1).sql` usa `DROP DATABASE IF EXISTS` y `OFFSET/FETCH`,
> por lo que requiere **SQL Server 2016+** para ejecutarse sin ajustes.

---

## 2. Estructura de archivos

```
webapppolimusicv1_asp/
├── conexion.asp           Núcleo: configuración, ADODB, utilidades y resolución de rutas
├── seguridad.asp          Guards de sesión y rol (RequiereSesion / RequiereAdmin)
├── header.asp             Cabecera común + navbar (usuario, rol, logout)
├── footer.asp             Cierre de layout, scripts y cierre de conexión
├── login.asp              Formulario de acceso (validación parametrizada)
├── logout.asp             Destrucción de sesión y retorno al login
├── canciones.asp          Listado + búsqueda + paginación + modal CRUD + reproductor
├── cancion_accion.asp     Endpoint POST: insertar / actualizar / eliminar (transaccional)
├── api_play.asp           Endpoint AJAX: incrementa PLAYS y devuelve JSON
├── css/
│   └── estilos.css        Diseño completo (tema oscuro violeta/cian, responsive)
├── js/
│   └── app.js             Reproductor, modal, confirmaciones, alertas
└── SongFilesAndImages_resources/
    ├── SongFiles/         Archivos .mp3
    ├── UserImages/        Fotos de usuario
    └── PlaylistImages/    Portadas de canciones / álbumes
```

---

## 3. Instalación en IIS

1. **Habilitar ASP Clásico**
   *Panel de control → Programas → Activar o desactivar características de Windows →
   Internet Information Services → Servicios World Wide Web → Características de
   desarrollo de aplicaciones → **ASP***.

2. **Crear el sitio o la aplicación virtual**
   *Administrador de IIS → clic derecho en **Sitios** → Agregar sitio web* (o
   *Agregar aplicación* dentro de un sitio existente) apuntando a la carpeta
   `webapppolimusicv1_asp`.

3. **Documento predeterminado**
   En *IIS → la aplicación → Documento predeterminado*, agrega `login.asp` como primera
   entrada (o navega directamente a `http://localhost/webapppolimusicv1_asp/login.asp`).

4. **Permisos de carpeta**
   El usuario del grupo `IIS_IUSRS` necesita **lectura** sobre toda la carpeta. Si vas a
   usar `ListarAudio()` (sugerencia de rutas en el formulario), también necesita lectura
   sobre `SongFilesAndImages_resources/SongFiles`.

5. **Base de datos**
   Ejecuta el script `BDD_PoliMusic (1).sql` en SQL Server Management Studio. Verifica que
   el servicio **SQL Server Browser** esté activo y que las conexiones TCP/IP estén
   habilitadas en el puerto `1433`.

6. **Ajustar la conexión** (si tu entorno difiere)
   ```vbscript
   ' conexion.asp
   Const DB_SERVIDOR = "localhost"
   Const DB_PUERTO   = "1433"
   Const DB_NOMBRE   = "BDD_PoliMusic"
   Const DB_USUARIO  = "sa"
   Const DB_CLAVE    = "1234"
   ```

> **No es necesario** habilitar *Rutas primarias* (Parent Paths): todos los `include` son
> relativos y la raíz virtual se calcula en tiempo de ejecución con `BaseUrl()`.

### Nota sobre los archivos de audio

La base de datos de ejemplo referencia cinco pistas
(`bensound-adventure.mp3`, `-dreams.mp3`, `-energy.mp3`, `-scifi.mp3`, `-tomorrow.mp3`),
pero en `SongFilesAndImages_resources/SongFiles/` sólo viene **`bensound-adventure.mp3`**.
Las demás filas se listan correctamente y el reproductor muestra el aviso
«No se pudo cargar el audio»: basta con copiar los `.mp3` restantes a esa carpeta
(o editar la ruta desde el formulario de administración) para que suenen.

Lo mismo ocurre con las portadas: la base apunta a `~/albumImages/albumN.png` y a
`~/ArtistImages/artistN.jpg`, carpetas que no existen en el proyecto. `UrlImagen()`
detecta el archivo faltante y devuelve `PlaylistImages/default.jpg`, así que nunca se
ven imágenes rotas.

---

## 4. Credenciales de prueba

| Usuario | Contraseña | Rol | Acceso |
|---|---|---|---|
| `userAdmin` | `12345678` | Administrador (`USER_TYPE = 1`) | CRUD completo |
| `userNormal` | `12345678` | Usuario Estándar (`USER_TYPE = 0`) | Sólo consulta y reproducción |
| `username1` | *(hash MD5)* | Usuario Estándar | Demuestra el soporte de MD5 |

El login acepta **contraseñas en texto plano y en MD5**. La comparación se hace en SQL:

```sql
WHERE USERNAME = ?
  AND ( PASSWORD = ?
        OR UPPER(PASSWORD) = UPPER(CONVERT(varchar(32), HASHBYTES('MD5', ?), 2)) )
```

---

## 5. Decisiones técnicas relevantes

### 5.1 Resolución de rutas `~/`
La base de datos guarda rutas tipo `~/songFiles/bensound-adventure.mp3`. Como el navegador
no entiende `~/`, `conexion.asp` incluye dos funciones:

* `RutaRecurso(ruta)` → traduce la carpeta lógica a la carpeta física real
  (`songFiles` → `SongFilesAndImages_resources/SongFiles`), sin distinguir mayúsculas.
* `UrlRecurso(ruta)` → antepone la raíz virtual de la aplicación (`BaseUrl()`), de modo
  que funciona igual en `http://localhost/` que en `http://localhost/webapppolimusicv1_asp/`.

`UrlImagen(ruta, porDefecto)` comprueba con `FileSystemObject` que el archivo exista y, si
no, devuelve `PlaylistImages/default.jpg`. Esto evita íconos rotos porque la base de datos
referencia carpetas (`~/albumImages/`, `~/ArtistImages/`) que no están en el proyecto.

### 5.2 Consultas parametrizadas
Todas las consultas usan `ADODB.Command` con parámetros (`?`). Nunca se concatena entrada
del usuario en el SQL, lo que neutraliza la inyección SQL.

### 5.3 Borrado seguro
`TBL_PLAYER`, `TBL_TRACKLIST` y `TBL_PLAYLIST_SONG` tienen claves foráneas hacia
`TBL_SONG`. El borrado se ejecuta dentro de una **transacción** y limpia primero las tres
tablas relacionadas; si algo falla, se hace `RollbackTrans`.

### 5.4 Relaciones N:M sin duplicar filas
Un artista puede tener varias canciones y una canción varios artistas. El listado concatena
los nombres con `STUFF(... FOR XML PATH(''))` para mostrar **una fila por canción**.

### 5.5 Contador de reproducciones
Al iniciar la reproducción, `js/app.js` llama a `api_play.asp`, que ejecuta un
`UPDATE ... OUTPUT INSERTED.PLAYS` (incremento atómico) y devuelve el nuevo valor para
actualizar la celda sin recargar la página.

### 5.6 Patrón POST-Redirect-GET
`cancion_accion.asp` siempre termina en una redirección con un mensaje *flash* guardado en
sesión, evitando que al refrescar se reenvíe el formulario.

---

## 6. Roles y permisos

| Elemento | Administrador | Usuario Estándar |
|---|---|---|
| Ver listado, buscar, paginar | Sí | Sí |
| Reproducir audio | Sí | Sí |
| Botón «+ Nueva Canción» | Sí | Oculto |
| Columna «Acciones» (editar / eliminar) | Sí | Oculto |
| Acceso directo a `cancion_accion.asp` | Permitido | **Bloqueado en servidor** |

Ocultar los botones no es suficiente: `RequiereAdmin` se ejecuta al inicio de
`cancion_accion.asp`, por lo que un usuario normal no puede forzar la acción escribiendo la
URL a mano.

---

## 7. Solución de problemas

| Error | Causa probable | Solución |
|---|---|---|
| `Error 500.19` / ASP no reconocido | Rol ASP no instalado | Habilitar *ASP* en Características de Windows |
| `ADODB.Connection error '80004005'` | SQL Server apagado, TCP deshabilitado o credenciales erróneas | Iniciar el servicio, habilitar TCP/IP en 1433, revisar `conexion.asp` |
| `No fue posible conectar…` (pantalla propia) | La app ya captura el error | Verifica la base `BDD_PoliMusic` y el usuario `sa` |
| `Syntax error` al abrir una página | Archivo guardado con BOM o codificación distinta de UTF-8 sin BOM | Guardar como **UTF-8 sin BOM** |
| El audio no suena (error 404 en consola) | El `.mp3` no existe en `SongFiles/` o el nombre no coincide | Verifica el nombre exacto y la columna `SONG_PATH` |
| Las tildes se ven como `Ã¡` | Codificación del archivo | Guardar en UTF-8 sin BOM; las páginas ya declaran `CodePage=65001` |
| `El proveedor 'SQLOLEDB' no está registrado` | Windows moderno sin el proveedor antiguo | La app reintenta con `Driver={SQL Server}` y `MSOLEDBSQL` automáticamente |

---

## 8. Codificación

Todos los archivos `.asp`, `.css` y `.js` están en **UTF-8 sin BOM**. Cada página declara:

```asp
<%@ Language="VBScript" CodePage="65001" %>
```

y `conexion.asp` establece `Response.CodePage`, `Response.CharSet` y `Session.CodePage`
en `65001`. Si editas los archivos con Visual Studio Code, usa la opción
*Guardar con codificación → UTF-8* (sin BOM).
