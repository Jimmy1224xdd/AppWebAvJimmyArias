<%
'==============================================================================
'  PoliMusic - seguridad.asp
'  Control de acceso basado en la sesión del usuario.
'
'  Reglas del sistema:
'     USER_TYPE = 0  ->  Usuario estándar (sólo consulta y reproduce)
'     USER_TYPE = 1  ->  Administrador (CRUD completo de canciones)
'
'  Incluir después de conexion.asp:
'      <!--#include file="seguridad.asp"-->
'==============================================================================

'--- HaySesion: True si existe un usuario autenticado -------------------------
Function HaySesion()
    HaySesion = (Len(Trim(CStr(Session("USERNAME") & ""))) > 0)
End Function

'--- TipoUsuario: devuelve 0 o 1 sin errores de conversión --------------------
Function TipoUsuario()
    Dim t
    t = Session("USER_TYPE")
    If IsEmpty(t) Or IsNull(t) Then
        TipoUsuario = 0
        Exit Function
    End If
    If Trim(CStr(t)) = "" Then
        TipoUsuario = 0
    Else
        TipoUsuario = AEntero(t, 0)
    End If
End Function

'--- EsAdmin: True cuando el usuario logueado es administrador ---------------
Function EsAdmin()
    EsAdmin = (TipoUsuario() = 1)
End Function

'--- EtiquetaRol: texto visible del rol --------------------------------------
Function EtiquetaRol()
    If EsAdmin() Then
        EtiquetaRol = "Administrador"
    Else
        EtiquetaRol = "Usuario Estándar"
    End If
End Function

'--- ClaseRol: clase CSS del badge -------------------------------------------
Function ClaseRol()
    If EsAdmin() Then
        ClaseRol = "badge--admin"
    Else
        ClaseRol = "badge--user"
    End If
End Function

'------------------------------------------------------------------------------
' RequiereSesion: corta la ejecución y envía al login si no hay sesión.
' Debe llamarse ANTES de imprimir cualquier HTML.
'------------------------------------------------------------------------------
Sub RequiereSesion()
    If Not HaySesion() Then
        Response.Redirect "login.asp"
        Response.End
    End If
End Sub

'------------------------------------------------------------------------------
' RequiereAdmin: exige sesión + rol administrador.
' Se usa también en los endpoints de acción (no basta con ocultar botones).
'------------------------------------------------------------------------------
Sub RequiereAdmin()
    If Not HaySesion() Then
        Response.Redirect "login.asp"
        Response.End
    End If
    If Not EsAdmin() Then
        Flash "error", "No tienes permisos para realizar esta acción."
        Response.Redirect "canciones.asp"
        Response.End
    End If
End Sub

'--- NombreUsuario: nombre del usuario en sesión ----------------------------
Function NombreUsuario()
    NombreUsuario = CStr(Session("USERNAME") & "")
End Function

'--- IniciarSesion: guarda los datos del usuario tras un login correcto ------
Sub IniciarSesion(idUsuario, nombreUsuario, tipoUsuario, fotoUsuario)
    Session("ID_USER")    = idUsuario
    Session("USERNAME")   = nombreUsuario
    Session("USER_TYPE")  = tipoUsuario
    Session("USER_PHOTO") = fotoUsuario
End Sub

'--- DestruirSesion: limpia variables y abandona la sesión ------------------
Sub DestruirSesion()
    Session.Contents.RemoveAll
    Session.Abandon
End Sub
%>
