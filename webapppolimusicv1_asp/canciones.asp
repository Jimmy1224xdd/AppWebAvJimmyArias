<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp"-->
<!--#include file="seguridad.asp"-->
<%
'==============================================================================
'  PoliMusic - canciones.asp
'  Página principal: listado de canciones con búsqueda, paginación, reproductor
'  global y operaciones CRUD restringidas al rol Administrador (USER_TYPE = 1).
'==============================================================================

'--- 1) Protección de sesión (antes de cualquier salida HTML) ----------------
RequiereSesion

Dim esUsuarioAdmin, usuarioActual, tituloPagina
Dim sFlashTipo, sFlashTexto
Dim q, pagina, porPagina, totalRegs, totalPaginas, offsetSql, totalPlays, queryFiltro
Dim modoForm, editId, editNombre, editRuta, editArtista, editAlbum
Dim cmd, rs, sql, whereSql
Dim idFila, nombreFila, rutaFila, playsFila, artistasFila, albumesFila, imagenFila
Dim idxFila, desdeReg, hastaReg, i, enlace
Dim rsArt, rsAlb, esEdicion, tituloModal, accionForm
Dim listaAudios, arrAudios, k, nomArt

esUsuarioAdmin = EsAdmin()
usuarioActual  = NombreUsuario()
tituloPagina   = "Canciones"

'--- 2) Mensajes flash (resultado de una acción previa) ----------------------
sFlashTipo  = FlashTipo()
sFlashTexto = FlashTexto()
LimpiarFlash

'--- 3) Filtros y paginación -------------------------------------------------
q = Trim(CStr(Request.QueryString("q") & ""))
If Len(q) > 60 Then q = Left(q, 60)

pagina    = AEntero(Request.QueryString("p"), 1)
If pagina < 1 Then pagina = 1
porPagina = 8

queryFiltro = ""
If Len(q) > 0 Then queryFiltro = "q=" & Server.URLEncode(q) & "&"

AbrirConexion

'--- 4) Carga del registro a editar (sólo administradores) -------------------
modoForm    = ""
editId      = 0
editNombre  = ""
editRuta    = "~/songFiles/"
editArtista = 0
editAlbum   = 0

If esUsuarioAdmin Then
    If CStr(Request.QueryString("nueva") & "") = "1" Then modoForm = "nueva"

    If IsNumeric(Request.QueryString("editar")) Then
        editId = AEntero(Request.QueryString("editar"), 0)
        If editId > 0 Then
            sql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, " & _
                  " (SELECT TOP 1 p.ID_ARTIST FROM TBL_PLAYER p " & _
                  "   WHERE p.ID_SONG = s.ID_SONG ORDER BY p.ID_ARTIST) AS ID_ARTIST, " & _
                  " (SELECT TOP 1 t.ID_ALBUM FROM TBL_TRACKLIST t " & _
                  "   WHERE t.ID_SONG = s.ID_SONG ORDER BY t.ID_ALBUM) AS ID_ALBUM " & _
                  "FROM TBL_SONG s WHERE s.ID_SONG = ?"

            Set cmd = NuevoComando(sql)
            AgregarParam cmd, "@id", adInteger, 4, editId
            Set rs = cmd.Execute

            If Not rs.EOF Then
                modoForm   = "editar"
                editNombre = CStr(rs("SONG_NAME") & "")
                editRuta   = CStr(rs("SONG_PATH") & "")
                If Not IsNull(rs("ID_ARTIST")) Then editArtista = AEntero(rs("ID_ARTIST"), 0)
                If Not IsNull(rs("ID_ALBUM"))  Then editAlbum   = AEntero(rs("ID_ALBUM"), 0)
            End If

            rs.Close
            Set rs  = Nothing
            Set cmd = Nothing
        End If
    End If
End If

'--- 5) Totales (para estadísticas y paginación) ----------------------------
Set cmd = NuevoComando("SELECT COUNT(*) AS TOTAL, " & _
                       "ISNULL(SUM(ISNULL(PLAYS,0)),0) AS REPRODUCCIONES FROM TBL_SONG")
Set rs = cmd.Execute
totalRegs  = 0
totalPlays = 0
If Not rs.EOF Then
    totalRegs  = AEntero(rs("TOTAL"), 0)
    totalPlays = AEntero(rs("REPRODUCCIONES"), 0)
End If
rs.Close
Set rs  = Nothing
Set cmd = Nothing

totalPaginas = Int((totalRegs + porPagina - 1) / porPagina)
If totalPaginas < 1 Then totalPaginas = 1
If pagina > totalPaginas Then pagina = totalPaginas

offsetSql = (pagina - 1) * porPagina
desdeReg  = offsetSql + 1
hastaReg  = offsetSql + porPagina
If hastaReg > totalRegs Then hastaReg = totalRegs

'--- 6) Consulta del listado -------------------------------------------------
'  Se concatenan los artistas y los álbumes de cada canción (relaciones N:M)
'  con STUFF + FOR XML PATH para no duplicar filas.
whereSql = ""
If Len(q) > 0 Then whereSql = " WHERE s.SONG_NAME LIKE ? "

sql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " & _
      " ISNULL(STUFF((SELECT ', ' + a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,'') " & _
      "                FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
      "                WHERE p.ID_SONG = s.ID_SONG FOR XML PATH('')), 1, 2, ''), '') AS ARTISTAS, " & _
      " ISNULL(STUFF((SELECT ', ' + al.TITLE_ALBUM " & _
      "                FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
      "                WHERE t.ID_SONG = s.ID_SONG FOR XML PATH('')), 1, 2, ''), '') AS ALBUMES, " & _
      " (SELECT TOP 1 al2.ALBUM_IMAGE_PATH " & _
      "    FROM TBL_TRACKLIST t2 INNER JOIN TBL_ALBUM al2 ON al2.ID_ALBUM = t2.ID_ALBUM " & _
      "   WHERE t2.ID_SONG = s.ID_SONG ORDER BY al2.ID_ALBUM) AS IMAGEN_ALBUM " & _
      "FROM TBL_SONG s" & whereSql & _
      " ORDER BY s.ID_SONG ASC OFFSET " & offsetSql & " ROWS FETCH NEXT " & porPagina & " ROWS ONLY"

Set cmd = NuevoComando(sql)
If Len(q) > 0 Then AgregarParam cmd, "@q", adVarChar, 62, "%" & q & "%"
Set rs = cmd.Execute
%>
<!--#include file="header.asp"-->

    <!-- ===================== ENCABEZADO DE SECCIÓN ===================== -->
    <section class="hero">
        <div class="hero__content">
            <span class="hero__kicker">
                <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true">
                    <path d="M9 18V5l12-2v13"></path><circle cx="6" cy="18" r="3"></circle>
                </svg>
                Biblioteca musical
            </span>
            <h1 class="hero__title">
                Song List
                <span class="hero__title-alt">/ Listado de Canciones</span>
            </h1>
            <p class="hero__sub">
                Explora el catálogo, reproduce tus pistas favoritas y consulta los artistas
                y álbumes asociados a cada canción.
            </p>
        </div>

        <div class="hero__actions">
            <% If esUsuarioAdmin Then %>
            <a class="btn btn--primary" href="canciones.asp?<%= queryFiltro %>nueva=1#modal">
                <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true">
                    <path d="M12 5v14"></path><path d="M5 12h14"></path>
                </svg>
                Nueva Canción
            </a>
            <% End If %>
        </div>
    </section>

    <!-- ===================== ESTADÍSTICAS ===================== -->
    <section class="stats">
        <article class="stat">
            <span class="stat__ico stat__ico--violet" aria-hidden="true">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M9 18V5l12-2v13"></path><circle cx="6" cy="18" r="3"></circle><circle cx="18" cy="16" r="3"></circle>
                </svg>
            </span>
            <div>
                <span class="stat__num"><%= FormatoNumero(totalRegs) %></span>
                <span class="stat__lbl">Canciones en el catálogo</span>
            </div>
        </article>

        <article class="stat">
            <span class="stat__ico stat__ico--cyan" aria-hidden="true">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M4 14v-2a8 8 0 0 1 16 0v2"></path>
                    <rect x="2.5" y="13.5" width="4" height="7" rx="1.6"></rect>
                    <rect x="17.5" y="13.5" width="4" height="7" rx="1.6"></rect>
                </svg>
            </span>
            <div>
                <span class="stat__num"><%= FormatoNumero(totalPlays) %></span>
                <span class="stat__lbl">Reproducciones totales</span>
            </div>
        </article>

        <article class="stat">
            <span class="stat__ico stat__ico--amber" aria-hidden="true">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M12 3l2.6 5.3 5.9.9-4.3 4.1 1 5.8L12 16.4 6.8 19.1l1-5.8L3.5 9.2l5.9-.9z"></path>
                </svg>
            </span>
            <div>
                <span class="stat__num stat__num--sm"><%= EtiquetaRol() %></span>
                <span class="stat__lbl">Sesión de <%= Esc(usuarioActual) %></span>
            </div>
        </article>
    </section>

    <!-- ===================== ALERTAS ===================== -->
    <% If Len(sFlashTipo) > 0 And Len(sFlashTexto) > 0 Then %>
    <div class="alerta alerta--<%= Esc(sFlashTipo) %>" role="alert">
        <span class="alerta__ico" aria-hidden="true">
            <% If sFlashTipo = "exito" Then %>
            <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6L9 17l-5-5"></path>
            </svg>
            <% Else %>
            <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                <circle cx="12" cy="12" r="9"></circle><path d="M12 8v5"></path><path d="M12 16.5v.01"></path>
            </svg>
            <% End If %>
        </span>
        <span class="alerta__txt"><%= Esc(sFlashTexto) %></span>
        <button class="alerta__close" type="button" data-cerrar-alerta aria-label="Cerrar aviso">&times;</button>
    </div>
    <% End If %>

    <!-- ===================== BARRA DE HERRAMIENTAS ===================== -->
    <section class="panel">
        <div class="toolbar">
            <form class="buscador" method="get" action="canciones.asp" role="search">
                <span class="buscador__ico" aria-hidden="true">
                    <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                        <circle cx="11" cy="11" r="7"></circle><path d="M20 20l-3.5-3.5"></path>
                    </svg>
                </span>
                <input class="buscador__input" type="search" name="q" value="<%= Esc(q) %>"
                       placeholder="Buscar canción por título…" maxlength="60" aria-label="Buscar canción">
                <% If Len(q) > 0 Then %>
                <a class="buscador__clear" href="canciones.asp" title="Quitar filtro">&times;</a>
                <% End If %>
            </form>

            <div class="toolbar__meta">
                <% If totalRegs > 0 Then %>
                <span class="toolbar__count">
                    Mostrando <b><%= desdeReg %>–<%= hastaReg %></b> de <b><%= totalRegs %></b>
                </span>
                <% Else %>
                <span class="toolbar__count">Sin resultados</span>
                <% End If %>

                <% If Not esUsuarioAdmin Then %>
                <span class="chip chip--solo">
                    <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                        <path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7-10-7-10-7z"></path><circle cx="12" cy="12" r="3"></circle>
                    </svg>
                    Modo sólo lectura
                </span>
                <% End If %>
            </div>
        </div>

        <!-- ===================== TABLA DE CANCIONES ===================== -->
        <div class="tabla-wrap">
            <table class="tabla">
                <thead>
                    <tr>
                        <th class="th-id">ID</th>
                        <th class="th-cover">Portada</th>
                        <th>Título de la Canción</th>
                        <th>Artista</th>
                        <th class="th-album">Álbum</th>
                        <th class="th-plays">Plays</th>
                        <th class="th-play">Reproductor</th>
                        <% If esUsuarioAdmin Then %>
                        <th class="th-acciones">Acciones</th>
                        <% End If %>
                    </tr>
                </thead>
                <tbody>
                <%
                idxFila = 0
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="<% If esUsuarioAdmin Then Response.Write "8" Else Response.Write "7" %>">
                            <div class="vacio">
                                <span class="vacio__ico" aria-hidden="true">
                                    <svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
                                        <circle cx="11" cy="11" r="7"></circle><path d="M20 20l-3.5-3.5"></path>
                                    </svg>
                                </span>
                                <p class="vacio__titulo">No se encontraron canciones</p>
                                <p class="vacio__txt">
                                    <% If Len(q) > 0 Then %>
                                    No hay coincidencias para «<b><%= Esc(q) %></b>». Prueba con otro término.
                                    <% Else %>
                                    Todavía no hay canciones registradas en el catálogo.
                                    <% End If %>
                                </p>
                            </div>
                        </td>
                    </tr>
                <%
                Else
                    Do While Not rs.EOF
                        idxFila   = idxFila + 1
                        idFila    = AEntero(rs("ID_SONG"), 0)
                        nombreFila = CStr(rs("SONG_NAME") & "")
                        rutaFila  = CStr(rs("SONG_PATH") & "")
                        playsFila = AEntero(rs("PLAYS"), 0)
                        artistasFila = DecodificarXml(CStr(rs("ARTISTAS") & ""))
                        albumesFila  = DecodificarXml(CStr(rs("ALBUMES") & ""))
                        imagenFila   = CStr(rs("IMAGEN_ALBUM") & "")

                        If Len(artistasFila) = 0 Then artistasFila = "Artista desconocido"
                        If Len(albumesFila)  = 0 Then albumesFila  = "—"
                %>
                    <tr class="fila" data-song-row="<%= idFila %>">
                        <td class="td-id"><span class="num">#<%= idFila %></span></td>

                        <td>
                            <div class="cover">
                                <img src="<%= UrlImagen(imagenFila, IMG_CANCION_DEFECTO) %>"
                                     alt="Portada de <%= Esc(nombreFila) %>" loading="lazy">
                                <button class="cover__play js-play" type="button"
                                        data-src="<%= Esc(UrlRecurso(rutaFila)) %>"
                                        data-title="<%= Esc(nombreFila) %>"
                                        data-artist="<%= Esc(artistasFila) %>"
                                        data-cover="<%= UrlImagen(imagenFila, IMG_CANCION_DEFECTO) %>"
                                        data-id="<%= idFila %>"
                                        data-plays="<%= playsFila %>"
                                        aria-label="Reproducir <%= Esc(nombreFila) %>">
                                    <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor" aria-hidden="true">
                                        <path d="M8 5.5v13l11-6.5z"></path>
                                    </svg>
                                </button>
                            </div>
                        </td>

                        <td>
                            <div class="song">
                                <span class="song__name"><%= Esc(nombreFila) %></span>
                                <span class="song__path" title="<%= Esc(rutaFila) %>"><%= Esc(rutaFila) %></span>
                            </div>
                        </td>

                        <td>
                            <span class="pill pill--artist">
                                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                                    <circle cx="12" cy="8" r="3.4"></circle><path d="M5 20c0-3.5 3.1-5.2 7-5.2s7 1.7 7 5.2"></path>
                                </svg>
                                <%= Esc(artistasFila) %>
                            </span>
                        </td>

                        <td class="td-album"><span class="album"><%= Esc(albumesFila) %></span></td>

                        <td class="td-plays">
                            <span class="plays" data-plays-id="<%= idFila %>"><%= FormatoNumero(playsFila) %></span>
                        </td>

                        <td>
                            <button class="btn-play js-play" type="button"
                                    data-src="<%= Esc(UrlRecurso(rutaFila)) %>"
                                    data-title="<%= Esc(nombreFila) %>"
                                    data-artist="<%= Esc(artistasFila) %>"
                                    data-cover="<%= UrlImagen(imagenFila, IMG_CANCION_DEFECTO) %>"
                                    data-id="<%= idFila %>"
                                    data-plays="<%= playsFila %>">
                                <span class="btn-play__ico" aria-hidden="true">
                                    <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor">
                                        <path d="M8 5.5v13l11-6.5z"></path>
                                    </svg>
                                </span>
                                <span class="btn-play__txt">Reproducir</span>
                            </button>
                        </td>

                        <% If esUsuarioAdmin Then %>
                        <td>
                            <div class="acciones">
                                <a class="btn-icon btn-icon--edit"
                                   href="canciones.asp?<%= queryFiltro %>editar=<%= idFila %>#modal"
                                   title="Editar «<%= Esc(nombreFila) %>»">
                                    <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                        <path d="M4 20h4l10-10a2.8 2.8 0 0 0-4-4L4 16z"></path>
                                    </svg>
                                    <span class="sr">Editar</span>
                                </a>

                                <form class="inline-form" method="post" action="cancion_accion.asp"
                                      data-confirmar="¿Eliminar la canción «<%= Esc(nombreFila) %>»? Esta acción no se puede deshacer.">
                                    <input type="hidden" name="accion" value="eliminar">
                                    <input type="hidden" name="id" value="<%= idFila %>">
                                    <input type="hidden" name="q" value="<%= Esc(q) %>">
                                    <input type="hidden" name="p" value="<%= pagina %>">
                                    <button class="btn-icon btn-icon--del" type="submit"
                                            title="Eliminar «<%= Esc(nombreFila) %>»">
                                        <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                            <path d="M4 7h16"></path><path d="M10 11v6"></path><path d="M14 11v6"></path>
                                            <path d="M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12"></path>
                                            <path d="M9 7V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2"></path>
                                        </svg>
                                        <span class="sr">Eliminar</span>
                                    </button>
                                </form>
                            </div>
                        </td>
                        <% End If %>
                    </tr>
                <%
                        rs.MoveNext
                    Loop
                End If
                %>
                </tbody>
            </table>
        </div>

        <!-- ===================== PAGINACIÓN ===================== -->
        <% If totalPaginas > 1 Then %>
        <nav class="paginacion" aria-label="Paginación del listado">
            <%
            If pagina > 1 Then
            %>
            <a class="pag__btn" href="canciones.asp?<%= queryFiltro %>p=<%= pagina - 1 %>" rel="prev">
                <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 6l-6 6 6 6"></path></svg>
                Anterior
            </a>
            <% Else %>
            <span class="pag__btn is-disabled">
                <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 6l-6 6 6 6"></path></svg>
                Anterior
            </span>
            <% End If %>

            <div class="pag__nums">
                <%
                For i = 1 To totalPaginas
                    If i = pagina Then
                %>
                    <span class="pag__num is-active"><%= i %></span>
                <%
                    Else
                %>
                    <a class="pag__num" href="canciones.asp?<%= queryFiltro %>p=<%= i %>"><%= i %></a>
                <%
                    End If
                Next
                %>
            </div>

            <%
            If pagina < totalPaginas Then
            %>
            <a class="pag__btn" href="canciones.asp?<%= queryFiltro %>p=<%= pagina + 1 %>" rel="next">
                Siguiente
                <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6l6 6-6 6"></path></svg>
            </a>
            <% Else %>
            <span class="pag__btn is-disabled">
                Siguiente
                <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6l6 6-6 6"></path></svg>
            </span>
            <% End If %>
        </nav>
        <% End If %>

    </section>

<%
' Cerramos el recordset del listado antes de abrir otros cursores
rs.Close
Set rs  = Nothing
Set cmd = Nothing
%>

<%
'==============================================================================
'  MODAL DE ALTA / EDICIÓN  (sólo Administrador)
'==============================================================================
If esUsuarioAdmin And Len(modoForm) > 0 Then
    esEdicion   = (modoForm = "editar")
    accionForm  = "insertar"
    If esEdicion Then
        accionForm  = "actualizar"
        tituloModal = "Editar canción"
    Else
        tituloModal = "Nueva canción"
    End If
%>
    <div class="modal is-open" id="modalCancion" role="presentation">
        <div class="modal__backdrop" data-cerrar-modal></div>

        <div class="modal__dialog" role="dialog" aria-modal="true" aria-labelledby="modalTitulo">
            <header class="modal__head">
                <div>
                    <span class="modal__kicker"><% If esEdicion Then %>Actualizar registro<% Else %>Crear registro<% End If %></span>
                    <h2 class="modal__title" id="modalTitulo"><%= tituloModal %></h2>
                </div>
                <a class="modal__close" href="canciones.asp?<%= queryFiltro %>p=<%= pagina %>" aria-label="Cerrar">&times;</a>
            </header>

            <form class="modal__body" method="post" action="cancion_accion.asp" autocomplete="off">
                <input type="hidden" name="accion" value="<%= accionForm %>">
                <input type="hidden" name="id" value="<%= editId %>">
                <input type="hidden" name="q" value="<%= Esc(q) %>">
                <input type="hidden" name="p" value="<%= pagina %>">

                <div class="campo">
                    <label class="campo__label" for="songName">Título de la canción <span class="req">*</span></label>
                    <div class="campo__box">
                        <input class="campo__input" type="text" id="songName" name="song_name"
                               value="<%= Esc(editNombre) %>" maxlength="50" required
                               placeholder="Ej. Adventure" autofocus>
                    </div>
                    <span class="campo__hint">Máximo 50 caracteres.</span>
                </div>

                <div class="campo">
                    <label class="campo__label" for="songPath">Ruta del archivo de audio <span class="req">*</span></label>
                    <div class="campo__box">
                        <input class="campo__input" type="text" id="songPath" name="song_path"
                               value="<%= Esc(editRuta) %>" maxlength="255" required
                               list="rutasAudio" placeholder="~/songFiles/bensound-adventure.mp3">
                    </div>
                    <span class="campo__hint">
                        Usa el formato <code>~/songFiles/archivo.mp3</code>. Se convertirá
                        automáticamente a la carpeta real de recursos.
                    </span>
                    <datalist id="rutasAudio">
                        <%
                        listaAudios = ListarAudio()
                        If Len(listaAudios) > 0 Then
                            arrAudios = Split(listaAudios, "|")
                            For k = 0 To UBound(arrAudios)
                                If Len(arrAudios(k)) > 0 Then
                        %>
                        <option value="~/songFiles/<%= Esc(arrAudios(k)) %>"></option>
                        <%
                                End If
                            Next
                        End If
                        %>
                    </datalist>
                </div>

                <div class="grid-2">
                    <div class="campo">
                        <label class="campo__label" for="songArtist">Artista</label>
                        <div class="campo__box">
                            <select class="campo__input campo__select" id="songArtist" name="id_artist">
                                <option value="0">— Sin artista —</option>
                                <%
                                Set cmd = NuevoComando("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME " & _
                                                       "FROM TBL_ARTIST ORDER BY ARTIST_NAME")
                                Set rsArt = cmd.Execute
                                Do While Not rsArt.EOF
                                    nomArt = CStr(rsArt("ARTIST_NAME") & "")
                                    If Not IsNull(rsArt("ARTIST_LASTNAME")) Then
                                        If Len(CStr(rsArt("ARTIST_LASTNAME"))) > 0 Then
                                            nomArt = nomArt & " " & CStr(rsArt("ARTIST_LASTNAME"))
                                        End If
                                    End If
                                %>
                                <option value="<%= rsArt("ID_ARTIST") %>" <% If editArtista = AEntero(rsArt("ID_ARTIST"), -1) Then Response.Write "selected" %>><%= Esc(nomArt) %></option>
                                <%
                                    rsArt.MoveNext
                                Loop
                                rsArt.Close
                                Set rsArt = Nothing
                                Set cmd = Nothing
                                %>
                            </select>
                        </div>
                    </div>

                    <div class="campo">
                        <label class="campo__label" for="songAlbum">Álbum</label>
                        <div class="campo__box">
                            <select class="campo__input campo__select" id="songAlbum" name="id_album">
                                <option value="0">— Sin álbum —</option>
                                <%
                                Set cmd = NuevoComando("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR " & _
                                                       "FROM TBL_ALBUM ORDER BY TITLE_ALBUM")
                                Set rsAlb = cmd.Execute
                                Do While Not rsAlb.EOF
                                %>
                                <option value="<%= rsAlb("ID_ALBUM") %>" <% If editAlbum = AEntero(rsAlb("ID_ALBUM"), -1) Then Response.Write "selected" %>><%= Esc(rsAlb("TITLE_ALBUM")) %><% If Not IsNull(rsAlb("RELEASE_YEAR")) Then Response.Write " (" & rsAlb("RELEASE_YEAR") & ")" %></option>
                                <%
                                    rsAlb.MoveNext
                                Loop
                                rsAlb.Close
                                Set rsAlb = Nothing
                                Set cmd = Nothing
                                %>
                            </select>
                        </div>
                    </div>
                </div>

                <footer class="modal__foot">
                    <a class="btn btn--ghost" href="canciones.asp?<%= queryFiltro %>p=<%= pagina %>">Cancelar</a>
                    <button class="btn btn--primary" type="submit">
                        <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                            <path d="M20 6L9 17l-5-5"></path>
                        </svg>
                        <% If esEdicion Then %>Guardar cambios<% Else %>Crear canción<% End If %>
                    </button>
                </footer>
            </form>
        </div>
    </div>
<%
End If
%>

    <!-- ===================== REPRODUCTOR GLOBAL ===================== -->
    <div class="player" id="playerBar" aria-label="Reproductor de audio">
        <div class="player__track">
            <img class="player__cover" id="playerCover" src="<%= BaseUrl() & IMG_CANCION_DEFECTO %>" alt="">
            <div class="player__meta">
                <span class="player__title" id="playerTitle">Selecciona una canción</span>
                <span class="player__artist" id="playerArtist">PoliMusic</span>
            </div>
        </div>

        <div class="player__center">
            <div class="player__buttons">
                <button class="pbtn" id="btnPrev" type="button" aria-label="Canción anterior">
                    <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden="true">
                        <path d="M7 5h2v14H7zM19 5v14L9 12z"></path>
                    </svg>
                </button>

                <button class="pbtn pbtn--main" id="btnPlayPause" type="button" aria-label="Reproducir">
                    <svg class="ico-play" viewBox="0 0 24 24" width="20" height="20" fill="currentColor" aria-hidden="true">
                        <path d="M8 5.5v13l11-6.5z"></path>
                    </svg>
                    <svg class="ico-pause" viewBox="0 0 24 24" width="20" height="20" fill="currentColor" aria-hidden="true">
                        <path d="M7 5h3.2v14H7zM13.8 5H17v14h-3.2z"></path>
                    </svg>
                </button>

                <button class="pbtn" id="btnNext" type="button" aria-label="Canción siguiente">
                    <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden="true">
                        <path d="M15 5h2v14h-2zM5 5l10 7-10 7z"></path>
                    </svg>
                </button>
            </div>

            <div class="player__progress">
                <span class="player__time" id="timeActual">0:00</span>
                <input class="rango" type="range" id="seek" min="0" max="1000" value="0" step="1"
                       aria-label="Progreso de la canción">
                <span class="player__time" id="timeTotal">0:00</span>
            </div>
        </div>

        <div class="player__right">
            <span class="player__eq" id="playerEq" aria-hidden="true">
                <i></i><i></i><i></i><i></i>
            </span>
            <svg class="player__vol-ico" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M4 10v4h4l5 4V6l-5 4H4z"></path><path d="M17.5 9.5a4 4 0 0 1 0 5"></path>
            </svg>
            <input class="rango rango--vol" type="range" id="volumen" min="0" max="100" value="85"
                   aria-label="Volumen">
        </div>

        <audio id="audioGlobal" preload="metadata"></audio>
    </div>

    <script>
        /* Datos mínimos que el reproductor necesita conocer del servidor */
        window.POLIMUSIC = {
            esAdmin: <% If esUsuarioAdmin Then Response.Write "true" Else Response.Write "false" %>
        };
    </script>

<!--#include file="footer.asp"-->
