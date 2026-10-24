<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp"-->
<%
'==============================================================================
'  PoliMusic - api_play.asp
'  Endpoint interno (AJAX) que incrementa el contador de reproducciones de una
'  canción y devuelve el nuevo valor en formato JSON.
'
'  Petición : POST  id=<ID_SONG>
'  Respuesta: {"ok":true,"plays":12}
'==============================================================================

Response.ContentType = "application/json"
Response.CharSet     = "utf-8"
Response.Expires     = 0
Response.CacheControl = "no-cache"

Dim idSong, cmd, rs, plays

'--- 1) Debe existir sesión activa ------------------------------------------
If Len(Trim(CStr(Session("USERNAME") & ""))) = 0 Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""ok"":false,""error"":""Sesion no iniciada""}"
    Response.End
End If

'--- 2) Sólo se acepta POST -------------------------------------------------
If UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) <> "POST" Then
    Response.Status = "405 Method Not Allowed"
    Response.Write "{""ok"":false,""error"":""Metodo no permitido""}"
    Response.End
End If

'--- 3) Validación del identificador ----------------------------------------
idSong = AEntero(Request.Form("id"), 0)
If idSong <= 0 Then
    Response.Status = "400 Bad Request"
    Response.Write "{""ok"":false,""error"":""Identificador invalido""}"
    Response.End
End If

'--- 4) Incremento atómico y lectura del nuevo valor ------------------------
AbrirConexion

plays = -1
On Error Resume Next

Set cmd = NuevoComando("UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS,0) + 1 " & _
                       "OUTPUT INSERTED.PLAYS AS NUEVO_PLAYS " & _
                       "WHERE ID_SONG = ?")
AgregarParam cmd, "@id", adInteger, 4, idSong
Set rs = cmd.Execute

If Err.Number = 0 Then
    If Not rs.EOF Then plays = AEntero(rs("NUEVO_PLAYS"), 0)
    rs.Close
    Set rs = Nothing
End If
Err.Clear
On Error GoTo 0

Set cmd = Nothing
CerrarConexion

If plays >= 0 Then
    Response.Write "{""ok"":true,""plays"":" & plays & "}"
Else
    Response.Status = "404 Not Found"
    Response.Write "{""ok"":false,""error"":""Cancion no encontrada""}"
End If
Response.End
%>
