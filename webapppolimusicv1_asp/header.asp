<%
'==============================================================================
'  PoliMusic - header.asp
'  Cabecera común: <head>, apertura del <body> y barra de navegación superior.
'
'  Variables que la página debe definir ANTES de incluir este archivo:
'      tituloPagina  (String)  -> título de la pestaña
'      paginaActiva  (String)  -> opcional, para resaltar el ítem de menú
'==============================================================================

Dim navUsuario, navFoto, navEsAdmin
navUsuario = NombreUsuario()
navFoto    = UrlImagen(Session("USER_PHOTO"), IMG_USUARIO_DEFECTO)
navEsAdmin = EsAdmin()

If Len(tituloPagina & "") = 0 Then tituloPagina = "PoliMusic"
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><%= Esc(tituloPagina) %> · PoliMusic</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="css/estilos.css">
</head>
<body class="app">

    <header class="navbar">
        <div class="navbar__inner">

            <a class="brand" href="canciones.asp" title="Ir al listado de canciones">
                <span class="brand__mark" aria-hidden="true">
                    <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M9 18V5l12-2v13"></path>
                        <circle cx="6" cy="18" r="3"></circle>
                        <circle cx="18" cy="16" r="3"></circle>
                    </svg>
                </span>
                <span class="brand__text">Poli<b>Music</b></span>
            </a>

            <nav class="navbar__nav" aria-label="Navegación principal">
                <a class="navbar__link is-active" href="canciones.asp">Canciones</a>
            </nav>

            <div class="navbar__user">
                <div class="navbar__identity">
                    <img class="avatar" src="<%= navFoto %>" alt="Foto de <%= Esc(navUsuario) %>">
                    <div class="navbar__data">
                        <span class="navbar__hello">Bienvenido/a, <strong><%= Esc(navUsuario) %></strong></span>
                        <span class="badge <%= ClaseRol() %>">
                            <% If navEsAdmin Then %>
                            <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                                <path d="M12 3l2.6 5.3 5.9.9-4.3 4.1 1 5.8L12 16.4 6.8 19.1l1-5.8L3.5 9.2l5.9-.9z"></path>
                            </svg>
                            <% Else %>
                            <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                                <circle cx="12" cy="8" r="3.6"></circle><path d="M4.5 20c0-3.6 3.3-5.4 7.5-5.4s7.5 1.8 7.5 5.4"></path>
                            </svg>
                            <% End If %>
                            <%= EtiquetaRol() %>
                        </span>
                    </div>
                </div>

                <a class="btn btn--ghost btn--sm" href="logout.asp" title="Cerrar la sesión actual">
                    <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                        <path d="M15 17l5-5-5-5"></path><path d="M20 12H9"></path>
                        <path d="M12 19H6a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2h6"></path>
                    </svg>
                    <span>Cerrar sesión</span>
                </a>
            </div>

        </div>
    </header>

    <main class="contenedor">
