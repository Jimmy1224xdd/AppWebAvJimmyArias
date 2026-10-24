<%
'==============================================================================
'  PoliMusic - conexion.asp
'  Núcleo compartido de la aplicación:
'    1) Configuración de la base de datos y constantes ADO
'    2) Apertura / cierre controlado de la conexión
'    3) Utilidades: escape HTML, resolución de rutas "~/" y helpers varios
'
'  Este archivo NO genera salida visible: sólo define funciones y variables.
'  Debe incluirse SIEMPRE como primera línea de código de cada página:
'      <!--#include file="conexion.asp"-->
'==============================================================================

'--- Codificación: todo el sitio trabaja en UTF-8 -----------------------------
Response.Buffer   = True
Response.CodePage = 65001
Response.CharSet  = "utf-8"
Session.CodePage  = 65001
Session.Timeout   = 60

'==============================================================================
'  1) CONFIGURACIÓN  (ajustar aquí si cambia el entorno)
'==============================================================================
Const DB_SERVIDOR = "localhost"          ' Servidor / instancia de SQL Server
Const DB_PUERTO   = "1433"               ' Puerto TCP (por defecto 1433)
Const DB_NOMBRE   = "BDD_PoliMusic"      ' Nombre de la base de datos
Const DB_USUARIO  = "sa"                 ' Autenticación SQL Server
Const DB_CLAVE    = "1234"

' Carpeta de recursos (relativa a la raíz de la aplicación)
Const CARPETA_RECURSOS = "SongFilesAndImages_resources"

' Imágenes por defecto cuando el archivo referenciado no existe en disco
Const IMG_CANCION_DEFECTO = "SongFilesAndImages_resources/PlaylistImages/default.jpg"
Const IMG_USUARIO_DEFECTO = "SongFilesAndImages_resources/UserImages/default.jpg"

'--- Constantes ADO (VBScript no las trae predefinidas) ----------------------
Const adCmdText          = 1
Const adParamInput       = 1
Const adInteger          = 3
Const adVarChar          = 200
Const adExecuteNoRecords = 128
Const adStateOpen        = 1

'==============================================================================
'  2) CONEXIÓN
'==============================================================================
Dim gConn
Set gConn = Nothing

'------------------------------------------------------------------------------
' AbrirConexion: crea el objeto ADODB.Connection y lo abre probando, en orden,
' varias cadenas de conexión (SQLOLEDB -> ODBC Driver -> MSOLEDBSQL).
' Si ninguna funciona, muestra una página de error limpia y detiene la página.
'------------------------------------------------------------------------------
Sub AbrirConexion()
    Dim cadenas, i, ok, mensaje

    ' Si ya hay una conexión abierta, no hacemos nada
    If IsObject(gConn) Then
        If Not (gConn Is Nothing) Then
            If gConn.State = adStateOpen Then Exit Sub
        End If
    End If

    cadenas = Array( _
        "Provider=SQLOLEDB;Data Source=" & DB_SERVIDOR & "," & DB_PUERTO & _
            ";Initial Catalog=" & DB_NOMBRE & ";User ID=" & DB_USUARIO & _
            ";Password=" & DB_CLAVE & ";", _
        "Driver={SQL Server};Server=" & DB_SERVIDOR & "," & DB_PUERTO & _
            ";Database=" & DB_NOMBRE & ";Uid=" & DB_USUARIO & _
            ";Pwd=" & DB_CLAVE & ";", _
        "Provider=MSOLEDBSQL;Server=" & DB_SERVIDOR & "," & DB_PUERTO & _
            ";Database=" & DB_NOMBRE & ";UID=" & DB_USUARIO & _
            ";PWD=" & DB_CLAVE & ";TrustServerCertificate=True;" _
    )

    ok = False
    For i = 0 To UBound(cadenas)
        On Error Resume Next
        Set gConn = Server.CreateObject("ADODB.Connection")
        gConn.ConnectionTimeout = 10
        gConn.CommandTimeout    = 30
        gConn.Open cadenas(i)
        If Err.Number = 0 Then
            If gConn.State = adStateOpen Then ok = True
        End If
        Err.Clear
        On Error GoTo 0

        If ok Then Exit For
        Set gConn = Nothing
    Next

    If Not ok Then
        mensaje = "No fue posible establecer comunicación con el servidor SQL Server. " & _
                  "Verifica que el servicio esté iniciado, que la base de datos " & _
                  "<strong>" & DB_NOMBRE & "</strong> exista y que las credenciales " & _
                  "configuradas en <em>conexion.asp</em> sean correctas."
        MostrarErrorConexion mensaje
    End If
End Sub

'------------------------------------------------------------------------------
' CerrarConexion: cierra y libera el objeto de conexión de forma segura.
'------------------------------------------------------------------------------
Sub CerrarConexion()
    On Error Resume Next
    If IsObject(gConn) Then
        If Not (gConn Is Nothing) Then
            If gConn.State = adStateOpen Then gConn.Close
            Set gConn = Nothing
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' MostrarErrorConexion: página de error autocontenida (no depende del layout).
'------------------------------------------------------------------------------
Sub MostrarErrorConexion(mensaje)
    Response.Clear
    Response.Write "<!DOCTYPE html><html lang=""es""><head><meta charset=""utf-8"">" & _
        "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & _
        "<title>PoliMusic - Error de conexión</title><style>" & _
        "body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;" & _
        "background:#0b0b12;color:#eef0f7;font-family:'Segoe UI',system-ui,sans-serif;padding:24px}" & _
        ".box{max-width:560px;background:#161624;border:1px solid #2a2a3d;border-radius:18px;" & _
        "padding:36px;box-shadow:0 24px 60px rgba(0,0,0,.55)}" & _
        ".ico{width:54px;height:54px;border-radius:14px;display:flex;align-items:center;" & _
        "justify-content:center;background:rgba(244,63,94,.12);color:#f43f5e;font-size:26px;" & _
        "margin-bottom:18px}h1{margin:0 0 10px;font-size:20px}" & _
        "p{color:#9aa0b5;line-height:1.65;font-size:14px;margin:0 0 8px}" & _
        "code{background:#1d1d2e;padding:2px 6px;border-radius:6px;font-size:12.5px;color:#a78bfa}" & _
        "</style></head><body><div class=""box""><div class=""ico"">!</div>" & _
        "<h1>Sin conexión a la base de datos</h1><p>" & mensaje & "</p>" & _
        "<p style=""margin-top:16px"">Cadena usada: <code>Provider=SQLOLEDB;Data Source=" & _
        DB_SERVIDOR & ";Initial Catalog=" & DB_NOMBRE & "</code></p>" & _
        "</div></body></html>"
    CerrarConexion
    Response.End
End Sub

'==============================================================================
'  3) UTILIDADES GENERALES
'==============================================================================

'--- Esc: escapa texto para insertarlo con seguridad dentro del HTML ----------
Function Esc(valor)
    Dim s
    s = CStr(valor & "")
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    Esc = s
End Function

'--- BaseUrl: raíz virtual de la aplicación (funciona en sitio o en carpeta) --
'   Ej.: si la página es /webapppolimusicv1_asp/canciones.asp -> "/webapppolimusicv1_asp/"
Function BaseUrl()
    Dim ruta, pos
    ruta = Request.ServerVariables("SCRIPT_NAME")
    pos  = InStrRev(ruta, "/")
    If pos > 0 Then
        BaseUrl = Left(ruta, pos)
    Else
        BaseUrl = "/"
    End If
End Function

'--- UrlEncodeArchivo: codifica sólo el nombre del archivo (espacios, tildes) --
Function UrlEncodeArchivo(nombre)
    Dim s
    s = Server.URLEncode(nombre)
    s = Replace(s, "+", "%20")
    UrlEncodeArchivo = s
End Function

'------------------------------------------------------------------------------
' RutaRecurso: traduce las rutas guardadas en la BD (estilo "~/songFiles/x.mp3")
' a una ruta relativa real dentro de la carpeta de recursos.
'
'   '~/songFiles/bensound-adventure.mp3'
'       -> 'SongFilesAndImages_resources/SongFiles/bensound-adventure.mp3'
'   '~/userImages/default.jpg'
'       -> 'SongFilesAndImages_resources/UserImages/default.jpg'
'
' La comparación de carpetas es insensible a mayúsculas porque en la base de
' datos aparecen como 'songFiles', 'userImages', 'playlistImages', etc.
'------------------------------------------------------------------------------
Function RutaRecurso(ruta)
    Dim r, pos, carpeta, resto, archivo

    r = Trim(CStr(ruta & ""))
    If Len(r) = 0 Then
        RutaRecurso = ""
        Exit Function
    End If

    ' Normalizamos separadores y quitamos el prefijo ~/ o /
    r = Replace(r, "\", "/")
    If Left(r, 2) = "~/" Then r = Mid(r, 3)
    If Left(r, 1) = "/"  Then r = Mid(r, 2)

    pos = InStr(r, "/")
    If pos > 0 Then
        carpeta = LCase(Left(r, pos - 1))
        resto   = Mid(r, pos + 1)
    Else
        carpeta = LCase(r)
        resto   = ""
    End If

    ' Mapa carpeta-lógica -> carpeta-física
    Select Case carpeta
        Case "songfiles"      : carpeta = "SongFiles"
        Case "userimages"     : carpeta = "UserImages"
        Case "playlistimages" : carpeta = "PlaylistImages"
        Case "albumimages"    : carpeta = "PlaylistImages"   ' no existe: se usa Playlist
        Case "artistimages"   : carpeta = "UserImages"       ' no existe: se usa UserImages
        Case Else             : carpeta = "PlaylistImages"
    End Select

    If Len(resto) = 0 Then
        RutaRecurso = CARPETA_RECURSOS & "/" & carpeta
    Else
        ' Codificamos sólo el nombre del archivo (los separadores se conservan)
        pos = InStrRev(resto, "/")
        If pos > 0 Then
            archivo = UrlEncodeArchivo(Mid(resto, pos + 1))
            RutaRecurso = CARPETA_RECURSOS & "/" & carpeta & "/" & _
                          Left(resto, pos) & archivo
        Else
            RutaRecurso = CARPETA_RECURSOS & "/" & carpeta & "/" & _
                          UrlEncodeArchivo(resto)
        End If
    End If
End Function

'--- UrlRecurso: ruta absoluta (desde la raíz virtual) lista para el navegador -
Function UrlRecurso(ruta)
    Dim rel
    rel = RutaRecurso(ruta)
    If Len(rel) = 0 Then
        UrlRecurso = ""
    Else
        UrlRecurso = BaseUrl() & rel
    End If
End Function

'------------------------------------------------------------------------------
' UrlImagen: igual que UrlRecurso, pero comprueba que el archivo exista
' físicamente. Si no existe devuelve la imagen por defecto (evita íconos rotos).
'------------------------------------------------------------------------------
Function UrlImagen(ruta, porDefecto)
    Dim rel, fisica, fso, existe

    UrlImagen = BaseUrl() & porDefecto
    rel = RutaRecurso(ruta)
    If Len(rel) = 0 Then Exit Function

    existe = False
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    fisica = Server.MapPath(rel)
    If Err.Number = 0 Then
        If fso.FileExists(fisica) Then existe = True
    End If
    Set fso = Nothing
    Err.Clear
    On Error GoTo 0

    If existe Then UrlImagen = BaseUrl() & rel
End Function

'--- RutaConArchivo: valida que la ruta termine en un nombre de archivo -------
'   Evita guardar rutas que apuntan sólo a una carpeta ('~/songFiles/').
Function RutaConArchivo(ruta)
    Dim r, pos, archivo

    r = Trim(CStr(ruta & ""))
    If Len(r) = 0 Then
        RutaConArchivo = False
        Exit Function
    End If

    r = Replace(r, "\", "/")
    pos = InStrRev(r, "/")
    If pos > 0 Then
        archivo = Mid(r, pos + 1)
    Else
        archivo = r
    End If

    archivo = Trim(archivo)
    RutaConArchivo = (Len(archivo) > 0 And InStr(archivo, ".") > 1)
End Function

'--- DecodificarXml: revierte el escapado que aplica FOR XML PATH ------------
'   Las listas de artistas/álbumes se construyen en SQL con FOR XML PATH, que
'   convierte & en &amp;, < en &lt;, etc. Al mostrarlas con Esc() se volverían a
'   escapar (doble escapado). Esta función lo evita.
Function DecodificarXml(valor)
    Dim s
    s = CStr(valor & "")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&apos;", "'")
    s = Replace(s, "&amp;", "&")
    DecodificarXml = s
End Function

'--- ListarAudio: devuelve los .mp3 disponibles en la carpeta SongFiles -------
'   Se usa para sugerir rutas válidas en el formulario de alta/edición.
Function ListarAudio()
    Dim fso, carpeta, archivo, lista, fisica

    lista = ""
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    fisica = Server.MapPath(CARPETA_RECURSOS & "/SongFiles")
    If Err.Number = 0 Then
        If fso.FolderExists(fisica) Then
            Set carpeta = fso.GetFolder(fisica)
            For Each archivo In carpeta.Files
                If LCase(fso.GetExtensionName(archivo.Name)) = "mp3" Then
                    lista = lista & archivo.Name & "|"
                End If
            Next
        End If
    End If
    Set carpeta = Nothing
    Set fso = Nothing
    Err.Clear
    On Error GoTo 0

    ListarAudio = lista
End Function

'==============================================================================
'  4) HELPERS DE ACCESO A DATOS (consultas parametrizadas)
'==============================================================================

'--- NuevoComando: crea un ADODB.Command ligado a la conexión activa ----------
Function NuevoComando(sql)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = gConn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    Set NuevoComando = cmd
End Function

'--- AgregarParam: añade un parámetro (evita inyección SQL) -------------------
Sub AgregarParam(cmd, nombre, tipo, tam, valor)
    Dim p
    If tam = 0 Then tam = 1
    Set p = cmd.CreateParameter(nombre, tipo, adParamInput, tam, valor)
    cmd.Parameters.Append p
End Sub

'--- EjecutarSinRetorno: para INSERT/UPDATE/DELETE ----------------------------
Sub EjecutarSinRetorno(cmd)
    cmd.Execute , , adExecuteNoRecords
End Sub

'--- AEntero: convierte de forma segura un valor a Long ----------------------
Function AEntero(valor, porDefecto)
    If IsNumeric(valor) Then
        On Error Resume Next
        AEntero = CLng(valor)
        If Err.Number <> 0 Then
            AEntero = porDefecto
            Err.Clear
        End If
        On Error GoTo 0
    Else
        AEntero = porDefecto
    End If
End Function

'--- FormatoNumero: separador de miles con punto (1.234) ----------------------
Function FormatoNumero(n)
    Dim s, i, res, cuenta
    s = CStr(CLng(n))
    res = ""
    cuenta = 0
    For i = Len(s) To 1 Step -1
        res = Mid(s, i, 1) & res
        cuenta = cuenta + 1
        If cuenta Mod 3 = 0 And i > 1 Then res = "." & res
    Next
    FormatoNumero = res
End Function

'--- FechaCorta: dd/mm/aaaa ------------------------------------------------
Function FechaCorta(f)
    If IsNull(f) Or Len(f & "") = 0 Then
        FechaCorta = "-"
    Else
        FechaCorta = Right("0" & Day(f), 2) & "/" & Right("0" & Month(f), 2) & "/" & Year(f)
    End If
End Function

'==============================================================================
'  5) MENSAJES FLASH (feedback tras redirecciones - patrón PRG)
'==============================================================================
Sub Flash(tipo, texto)
    Session("flash_tipo")  = tipo
    Session("flash_texto") = texto
End Sub

Function FlashTipo()
    FlashTipo = CStr(Session("flash_tipo") & "")
End Function

Function FlashTexto()
    FlashTexto = CStr(Session("flash_texto") & "")
End Function

Sub LimpiarFlash()
    Session("flash_tipo")  = ""
    Session("flash_texto") = ""
End Sub
%>
