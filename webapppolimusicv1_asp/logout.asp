<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp"-->
<!--#include file="seguridad.asp"-->
<%
'==============================================================================
'  PoliMusic - logout.asp
'  Destruye por completo la sesión del usuario y vuelve al login.
'  No imprime nada antes de redireccionar (Response.Buffer ya está activo).
'==============================================================================

' 1) Borra todas las variables de sesión
Session.Contents.RemoveAll

' 2) Invalida la sesión en el servidor (nuevo identificador en la siguiente)
Session.Abandon

' 3) Evita que el navegador muestre páginas protegidas desde la caché
Response.Expires        = -1
Response.ExpiresAbsolute = DateAdd("d", -1, Date())
Response.CacheControl   = "no-cache"
Response.AddHeader "Pragma", "no-cache"

CerrarConexion

' 5) De vuelta al login
Response.Redirect "login.asp"
Response.End
%>
