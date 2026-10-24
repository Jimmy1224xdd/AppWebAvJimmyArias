<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp"-->
<!--#include file="seguridad.asp"-->
<%
'==============================================================================
'  PoliMusic - cancion_accion.asp
'  Endpoint único para las operaciones de escritura sobre TBL_SONG:
'      accion = insertar    -> INSERT  + relaciones (TBL_PLAYER / TBL_TRACKLIST)
'      accion = actualizar  -> UPDATE  + resincronización de relaciones
'      accion = eliminar    -> DELETE  en cascada manual (respeta las FK)
'
'  Se ejecuta siempre por POST, dentro de una transacción, y termina con una
'  redirección (patrón POST-Redirect-GET) para evitar reenvíos duplicados.
'==============================================================================

'--- 1) Sólo administradores autenticados ------------------------------------
RequiereAdmin

'--- 2) Sólo por POST ---------------------------------------------------------
If UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) <> "POST" Then
    Response.Redirect "canciones.asp"
    Response.End
End If

Dim accion, idSong, nombre, ruta, idArtista, idAlbum
Dim q, pag, destino, sep, cmd, rs, sql, ok, nuevoId, afectados
Dim mensajeError, destinoForm

accion = LCase(Trim(CStr(Request.Form("accion") & "")))
idSong = AEntero(Request.Form("id"), 0)
q      = Trim(CStr(Request.Form("q") & ""))
pag    = AEntero(Request.Form("p"), 1)
If pag < 1 Then pag = 1

'--- 3) Destino de retorno (conserva el filtro y la página actual) -----------
destino = "canciones.asp"
sep = "?"
If Len(q) > 0 Then
    destino = destino & sep & "q=" & Server.URLEncode(q)
    sep = "&"
End If
If pag > 1 Then destino = destino & sep & "p=" & pag

If Len(accion) = 0 Then
    Flash "error", "No se especificó ninguna acción a realizar."
    Response.Redirect destino
    Response.End
End If

AbrirConexion

Select Case accion

'==============================================================================
'  ALTA
'==============================================================================
Case "insertar"

    nombre    = Trim(CStr(Request.Form("song_name") & ""))
    ruta      = Trim(CStr(Request.Form("song_path") & ""))
    idArtista = AEntero(Request.Form("id_artist"), 0)
    idAlbum   = AEntero(Request.Form("id_album"), 0)

    mensajeError = ""
    If Len(nombre) = 0 Then
        mensajeError = "El título de la canción es obligatorio."
    ElseIf Len(nombre) > 50 Then
        mensajeError = "El título no puede superar los 50 caracteres."
    ElseIf Len(ruta) = 0 Then
        mensajeError = "La ruta del archivo de audio es obligatoria."
    ElseIf Len(ruta) > 255 Then
        mensajeError = "La ruta no puede superar los 255 caracteres."
    ElseIf Not RutaConArchivo(ruta) Then
        mensajeError = "La ruta debe incluir el nombre del archivo (por ejemplo " & _
                       "~/songFiles/bensound-adventure.mp3)."
    End If

    If Len(mensajeError) > 0 Then
        Flash "error", mensajeError
        CerrarConexion
        Response.Redirect "canciones.asp?nueva=1"
        Response.End
    End If

    ok = True
    nuevoId = 0
    On Error Resume Next

    gConn.BeginTrans
    If Err.Number <> 0 Then ok = False

    If ok Then
        ' OUTPUT INSERTED nos devuelve el ID recién generado de forma fiable
        Set cmd = NuevoComando("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) " & _
                               "OUTPUT INSERTED.ID_SONG AS NUEVO_ID VALUES (?, ?, 0)")
        AgregarParam cmd, "@nombre", adVarChar, 50,  nombre
        AgregarParam cmd, "@ruta",   adVarChar, 255, ruta
        Set rs = cmd.Execute
        If Err.Number <> 0 Then
            ok = False
        Else
            If Not rs.EOF Then nuevoId = AEntero(rs("NUEVO_ID"), 0)
            rs.Close
            Set rs = Nothing
        End If
        Set cmd = Nothing
    End If

    '--- Relación con artista (TBL_PLAYER) -----------------------------------
    If ok And nuevoId > 0 And idArtista > 0 Then
        Set cmd = NuevoComando("INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)")
        AgregarParam cmd, "@song",   adInteger, 4, nuevoId
        AgregarParam cmd, "@artist", adInteger, 4, idArtista
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    '--- Relación con álbum (TBL_TRACKLIST) ----------------------------------
    If ok And nuevoId > 0 And idAlbum > 0 Then
        Set cmd = NuevoComando("INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)")
        AgregarParam cmd, "@album", adInteger, 4, idAlbum
        AgregarParam cmd, "@song",  adInteger, 4, nuevoId
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    If ok Then
        gConn.CommitTrans
        If Err.Number <> 0 Then ok = False
    Else
        gConn.RollbackTrans
    End If
    Err.Clear
    On Error GoTo 0

    CerrarConexion

    If ok Then
        Flash "exito", "La canción «" & nombre & "» se creó correctamente."
        Response.Redirect destino
    Else
        Flash "error", "No se pudo crear la canción. Verifica los datos e inténtalo de nuevo."
        Response.Redirect "canciones.asp?nueva=1"
    End If
    Response.End

'==============================================================================
'  MODIFICACIÓN
'==============================================================================
Case "actualizar"

    nombre    = Trim(CStr(Request.Form("song_name") & ""))
    ruta      = Trim(CStr(Request.Form("song_path") & ""))
    idArtista = AEntero(Request.Form("id_artist"), 0)
    idAlbum   = AEntero(Request.Form("id_album"), 0)

    destinoForm = "canciones.asp?editar=" & idSong

    mensajeError = ""
    If idSong <= 0 Then
        mensajeError = "No se identificó la canción a modificar."
    ElseIf Len(nombre) = 0 Then
        mensajeError = "El título de la canción es obligatorio."
    ElseIf Len(nombre) > 50 Then
        mensajeError = "El título no puede superar los 50 caracteres."
    ElseIf Len(ruta) = 0 Then
        mensajeError = "La ruta del archivo de audio es obligatoria."
    ElseIf Len(ruta) > 255 Then
        mensajeError = "La ruta no puede superar los 255 caracteres."
    ElseIf Not RutaConArchivo(ruta) Then
        mensajeError = "La ruta debe incluir el nombre del archivo (por ejemplo " & _
                       "~/songFiles/bensound-adventure.mp3)."
    End If

    If Len(mensajeError) > 0 Then
        Flash "error", mensajeError
        CerrarConexion
        If idSong > 0 Then
            Response.Redirect destinoForm
        Else
            Response.Redirect destino
        End If
        Response.End
    End If

    ok = True
    On Error Resume Next

    gConn.BeginTrans
    If Err.Number <> 0 Then ok = False

    '--- 1) Datos principales -----------------------------------------------
    If ok Then
        Set cmd = NuevoComando("UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ? WHERE ID_SONG = ?")
        AgregarParam cmd, "@nombre", adVarChar, 50,  nombre
        AgregarParam cmd, "@ruta",   adVarChar, 255, ruta
        AgregarParam cmd, "@id",     adInteger, 4,   idSong
        cmd.Execute afectados, , adExecuteNoRecords
        If Err.Number <> 0 Then
            ok = False
        ElseIf afectados = 0 Then
            ok = False
        End If
        Set cmd = Nothing
    End If

    '--- 2) Se eliminan las relaciones anteriores ---------------------------
    If ok Then
        Set cmd = NuevoComando("DELETE FROM TBL_PLAYER WHERE ID_SONG = ?")
        AgregarParam cmd, "@id", adInteger, 4, idSong
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    If ok Then
        Set cmd = NuevoComando("DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?")
        AgregarParam cmd, "@id", adInteger, 4, idSong
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    '--- 3) Se vuelven a crear según lo elegido en el formulario ------------
    If ok And idArtista > 0 Then
        Set cmd = NuevoComando("INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)")
        AgregarParam cmd, "@song",   adInteger, 4, idSong
        AgregarParam cmd, "@artist", adInteger, 4, idArtista
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    If ok And idAlbum > 0 Then
        Set cmd = NuevoComando("INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)")
        AgregarParam cmd, "@album", adInteger, 4, idAlbum
        AgregarParam cmd, "@song",  adInteger, 4, idSong
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    If ok Then
        gConn.CommitTrans
        If Err.Number <> 0 Then ok = False
    Else
        gConn.RollbackTrans
    End If
    Err.Clear
    On Error GoTo 0

    CerrarConexion

    If ok Then
        Flash "exito", "La canción «" & nombre & "» se actualizó correctamente."
        Response.Redirect destino
    Else
        Flash "error", "No se pudo actualizar la canción. Inténtalo nuevamente."
        Response.Redirect destinoForm
    End If
    Response.End

'==============================================================================
'  BAJA
'==============================================================================
Case "eliminar"

    If idSong <= 0 Then
        Flash "error", "No se identificó la canción a eliminar."
        CerrarConexion
        Response.Redirect destino
        Response.End
    End If

    ok = True
    afectados = 0
    On Error Resume Next

    gConn.BeginTrans
    If Err.Number <> 0 Then ok = False

    ' Las FK obligan a limpiar primero todas las tablas relacionadas
    If ok Then
        Set cmd = NuevoComando("DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?")
        AgregarParam cmd, "@id", adInteger, 4, idSong
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    If ok Then
        Set cmd = NuevoComando("DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?")
        AgregarParam cmd, "@id", adInteger, 4, idSong
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    If ok Then
        Set cmd = NuevoComando("DELETE FROM TBL_PLAYER WHERE ID_SONG = ?")
        AgregarParam cmd, "@id", adInteger, 4, idSong
        EjecutarSinRetorno cmd
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    If ok Then
        Set cmd = NuevoComando("DELETE FROM TBL_SONG WHERE ID_SONG = ?")
        AgregarParam cmd, "@id", adInteger, 4, idSong
        cmd.Execute afectados, , adExecuteNoRecords
        If Err.Number <> 0 Then ok = False
        Set cmd = Nothing
    End If

    If ok And afectados > 0 Then
        gConn.CommitTrans
        If Err.Number <> 0 Then ok = False
    Else
        ok = False
        gConn.RollbackTrans
    End If
    Err.Clear
    On Error GoTo 0

    CerrarConexion

    If ok Then
        Flash "exito", "La canción se eliminó correctamente."
    Else
        Flash "error", "No se pudo eliminar la canción. Puede que ya no exista."
    End If

    Response.Redirect destino
    Response.End

'==============================================================================
'  ACCIÓN DESCONOCIDA
'==============================================================================
Case Else
    CerrarConexion
    Flash "error", "Acción no reconocida."
    Response.Redirect destino
    Response.End

End Select
%>
