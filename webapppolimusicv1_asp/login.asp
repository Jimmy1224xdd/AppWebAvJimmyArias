<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp"-->
<!--#include file="seguridad.asp"-->
<%
'==============================================================================
'  PoliMusic - login.asp
'  Formulario de acceso. Valida credenciales contra TBL_USER mediante una
'  consulta parametrizada (sin concatenar datos del usuario).
'
'  Soporta los dos formatos de contraseña que existen en la base de datos:
'    - Texto plano  (ej. '12345678')
'    - MD5 en mayúsculas (ej. '7C222FB2927D828AF22F592134E8932480637C0D')
'==============================================================================

' Si ya hay sesión activa, no tiene sentido mostrar el formulario
If HaySesion() Then
    Response.Redirect "canciones.asp"
    Response.End
End If

Dim errorLogin, usuario, clave, cmd, rs, sql
Dim idUsuario, nombreBd, tipoBd, fotoBd, recordar, consultaOk

errorLogin = ""
usuario    = ""
recordar   = False
consultaOk = False

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    usuario = Trim(Request.Form("usuario") & "")
    clave   = CStr(Request.Form("clave") & "")
    If Request.Form("recordar") = "1" Then recordar = True

    '--- Validaciones básicas de entrada --------------------------------------
    If Len(usuario) = 0 Or Len(clave) = 0 Then
        errorLogin = "Debes ingresar tu usuario y tu contraseña."
    ElseIf Len(usuario) > 50 Then
        errorLogin = "El nombre de usuario no puede superar los 50 caracteres."
    Else

        AbrirConexion

        '--- Consulta parametrizada -------------------------------------------
        '  Se comparan las 3 posibilidades: texto plano y MD5 calculado por
        '  el propio motor, de modo que nunca se concatena la entrada.
        sql = "SELECT TOP 1 ID_USER, USERNAME, USER_TYPE, USER_PHOTO " & _
              "FROM TBL_USER " & _
              "WHERE USERNAME = ? " & _
              "  AND ( PASSWORD = ? " & _
              "        OR UPPER(PASSWORD) = UPPER(CONVERT(varchar(32), HASHBYTES('MD5', ?), 2)) )"

        Set cmd = NuevoComando(sql)
        AgregarParam cmd, "@usuario", adVarChar, 50,  usuario
        AgregarParam cmd, "@clave",   adVarChar, 50,  clave
        AgregarParam cmd, "@clave2",  adVarChar, 50,  clave

        On Error Resume Next
        Set rs = cmd.Execute
        If Err.Number <> 0 Then
            errorLogin = "Ocurrió un error al validar las credenciales. " & _
                         "Inténtalo nuevamente."
            Err.Clear
        Else
            consultaOk = True
        End If
        On Error GoTo 0

        If errorLogin = "" Then
            If Not consultaOk Then
                errorLogin = "No fue posible validar las credenciales."
            ElseIf rs.EOF Then
                errorLogin = "Usuario o contraseña incorrectos."
                rs.Close
                Set rs = Nothing
            Else
                idUsuario = rs("ID_USER")
                nombreBd  = CStr(rs("USERNAME") & "")
                tipoBd    = AEntero(rs("USER_TYPE"), 0)
                fotoBd    = CStr(rs("USER_PHOTO") & "")
                rs.Close
                Set rs = Nothing

                '--- Sesión iniciada ------------------------------------------
                IniciarSesion idUsuario, nombreBd, tipoBd, fotoBd
                Session("fecha_ingreso") = Now()

                ' "Mantener mi sesión iniciada" amplía el tiempo de vida a 8 h
                If recordar Then Session.Timeout = 480

                CerrarConexion
                Response.Redirect "canciones.asp"
                Response.End
            End If
        End If

        Set cmd = Nothing
        CerrarConexion
    End If
End If

'--- Foto del usuario (sólo si el POST falló y ya teníamos una) --------------
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>PoliMusic · Iniciar sesión</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="css/estilos.css">
</head>
<body class="auth">

    <div class="auth__bg" aria-hidden="true">
        <span class="orb orb--1"></span>
        <span class="orb orb--2"></span>
        <span class="orb orb--3"></span>
    </div>

    <main class="auth__wrap">

        <section class="auth__card">

            <header class="auth__head">
                <div class="brand brand--lg">
                    <span class="brand__mark" aria-hidden="true">
                        <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M9 18V5l12-2v13"></path>
                            <circle cx="6" cy="18" r="3"></circle>
                            <circle cx="18" cy="16" r="3"></circle>
                        </svg>
                    </span>
                    <span class="brand__text">Poli<b>Music</b></span>
                </div>
                <h1 class="auth__title">Bienvenido de vuelta</h1>
                <p class="auth__sub">Ingresa tus credenciales Jimmy Arias para acceder a tu biblioteca musical.</p>
            </header>

            <% If Len(errorLogin) > 0 Then %>
            <div class="alerta alerta--error" role="alert">
                <span class="alerta__ico" aria-hidden="true">
                    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                        <circle cx="12" cy="12" r="9"></circle><path d="M12 8v5"></path><path d="M12 16.5v.01"></path>
                    </svg>
                </span>
                <span class="alerta__txt"><%= Esc(errorLogin) %></span>
            </div>
            <% End If %>

            <form class="auth__form" method="post" action="login.asp" autocomplete="on" novalidate>

                <div class="campo">
                    <label class="campo__label" for="usuario">Usuario</label>
                    <div class="campo__box">
                        <span class="campo__ico" aria-hidden="true">
                            <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                                <circle cx="12" cy="8" r="4"></circle><path d="M4 21c0-4 3.6-6 8-6s8 2 8 6"></path>
                            </svg>
                        </span>
                        <input class="campo__input" type="text" id="usuario" name="usuario"
                               value="<%= Esc(usuario) %>" placeholder="Ej. userAdmin"
                               maxlength="50" required autofocus>
                    </div>
                </div>

                <div class="campo">
                    <label class="campo__label" for="clave">Contraseña</label>
                    <div class="campo__box">
                        <span class="campo__ico" aria-hidden="true">
                            <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                                <rect x="4" y="10" width="16" height="10" rx="2"></rect>
                                <path d="M8 10V7a4 4 0 0 1 8 0v3"></path>
                            </svg>
                        </span>
                        <input class="campo__input" type="password" id="clave" name="clave"
                               placeholder="••••••••" maxlength="50" required>
                        <button class="campo__ojo" type="button" data-toggle-password="clave"
                                aria-label="Mostrar u ocultar contraseña">
                            <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                                <path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7-10-7-10-7z"></path>
                                <circle cx="12" cy="12" r="3"></circle>
                            </svg>
                        </button>
                    </div>
                </div>

                <label class="check">
                    <input type="checkbox" name="recordar" value="1" <% If recordar Then Response.Write "checked" %>>
                    <span>Mantener mi sesión iniciada</span>
                </label>

                <button class="btn btn--primary btn--block" type="submit">
                    <span>Ingresar</span>
                    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                        <path d="M5 12h14"></path><path d="M13 6l6 6-6 6"></path>
                    </svg>
                </button>
            </form>

            <footer class="auth__foot">
                <div class="auth__demo">
                    <span class="auth__demo-title">Cuentas de prueba</span>
                    <span><b>userAdmin</b> / 12345678 &nbsp;·&nbsp; <b>userNormal</b> / 12345678</span>
                </div>
                <p class="auth__copy">PoliMusic &copy; <%= Year(Date()) %> · ASP Clásico + SQL Server</p>
            </footer>

        </section>
    </main>

    <script src="js/app.js"></script>
</body>
</html>
