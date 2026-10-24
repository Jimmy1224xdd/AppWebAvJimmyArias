# PoliMusic — notas permanentes del proyecto

## Stack y entorno
- **ASP Clásico (VBScript) + SQL Server**. Nada de ASP.NET, PHP ni Node en el servidor.
- Base: `BDD_PoliMusic` en `localhost,1433`, autenticación SQL (`sa` / `1234`), configurada
  en las constantes `DB_*` de `conexion.asp`.
- Raíz de la app: `webapppolimusicv1_asp/` (puede colgar de un sitio o de una carpeta virtual).

## Convenciones obligatorias
1. **Codificación**: todos los `.asp` en **UTF-8 sin BOM**; cada página declara
   `<%@ Language="VBScript" CodePage="65001" %>` y `conexion.asp` fija
   `Response.CodePage/CharSet` y `Session.CodePage = 65001`. Con BOM aparecen caracteres
   basura al inicio de la respuesta.
2. **Rutas de recursos**: en BD se guardan como `~/songFiles/x.mp3`. En HTML **nunca** se
   imprimen directo: usar `UrlRecurso()` (audio) o `UrlImagen()` (imágenes, con fallback).
3. **SQL**: siempre parametrizado con `NuevoComando` + `AgregarParam` (nunca concatenar
   entrada del usuario).
4. **Roles**: `USER_TYPE = 1` administrador, `0` usuario estándar. Ocultar botones no basta:
   `cancion_accion.asp` llama a `RequiereAdmin` al inicio.
5. **Borrado de canciones**: limpiar `TBL_PLAYLIST_SONG`, `TBL_TRACKLIST` y `TBL_PLAYER`
   antes de `TBL_SONG`, dentro de una transacción (las tres tienen FK).
6. **Escrituras**: patrón POST-Redirect-GET con mensajes flash en sesión.

## Estructura de includes
`conexion.asp` → `seguridad.asp` → (contenido) → `header.asp` / `footer.asp`.
`footer.asp` cierra la conexión al final de cada página.

## Datos de prueba
`userAdmin` / `12345678` (admin) · `userNormal` / `12345678` (estándar).
