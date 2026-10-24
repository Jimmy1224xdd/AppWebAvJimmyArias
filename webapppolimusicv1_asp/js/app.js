/* ==========================================================================
   PoliMusic · app.js
   JavaScript nativo (sin dependencias). Módulos:
     1. Utilidades
     2. Mostrar / ocultar contraseña
     3. Alertas descartables
     4. Confirmación antes de eliminar
     5. Modal de alta / edición
     6. Reproductor global + contador de reproducciones
   ========================================================================== */
(function () {
    'use strict';

    /* ----------------------------------------------------------------------
       1. UTILIDADES
       ---------------------------------------------------------------------- */
    var $  = function (sel, ctx) { return (ctx || document).querySelector(sel); };
    var $$ = function (sel, ctx) {
        return Array.prototype.slice.call((ctx || document).querySelectorAll(sel));
    };

    /** Convierte segundos a formato m:ss */
    function formatoTiempo(seg) {
        if (!isFinite(seg) || seg < 0) { return '0:00'; }
        var m = Math.floor(seg / 60);
        var s = Math.floor(seg % 60);
        return m + ':' + (s < 10 ? '0' : '') + s;
    }

    /** Pinta el relleno de un input range según su porcentaje */
    function pintarRango(el, pct, color) {
        if (!el) { return; }
        pct = Math.max(0, Math.min(100, pct));
        var c = color || 'var(--violet)';
        el.style.background =
            'linear-gradient(90deg,' + c + ' 0%,' + c + ' ' + pct + '%,' +
            '#2a2a40 ' + pct + '%,#2a2a40 100%)';
    }

    /* ----------------------------------------------------------------------
       2. MOSTRAR / OCULTAR CONTRASEÑA  (login)
       ---------------------------------------------------------------------- */
    $$('[data-toggle-password]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            var input = document.getElementById(btn.getAttribute('data-toggle-password'));
            if (!input) { return; }
            var esTexto = input.type === 'text';
            input.type = esTexto ? 'password' : 'text';
            btn.style.color = esTexto ? '' : 'var(--violet-lt)';
            input.focus();
        });
    });

    /* ----------------------------------------------------------------------
       3. ALERTAS DESCARTABLES
       ---------------------------------------------------------------------- */
    $$('[data-cerrar-alerta]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            var alerta = btn.closest('.alerta');
            if (!alerta) { return; }
            alerta.style.transition = 'opacity .25s ease, transform .25s ease';
            alerta.style.opacity = '0';
            alerta.style.transform = 'translateY(-6px)';
            setTimeout(function () { alerta.remove(); }, 260);
        });
    });

    /* ----------------------------------------------------------------------
       4. CONFIRMACIÓN ANTES DE ELIMINAR
       ---------------------------------------------------------------------- */
    $$('form[data-confirmar]').forEach(function (form) {
        form.addEventListener('submit', function (ev) {
            var mensaje = form.getAttribute('data-confirmar') || '¿Confirmas esta acción?';
            if (!window.confirm(mensaje)) {
                ev.preventDefault();
            }
        });
    });

    /* ----------------------------------------------------------------------
       5. MODAL DE ALTA / EDICIÓN
       ---------------------------------------------------------------------- */
    var modal = $('#modalCancion');

    function cerrarModal() {
        if (!modal) { return; }
        modal.classList.remove('is-open');
        // Limpia ?nueva=1 / ?editar=N de la barra de direcciones
        if (window.history && window.history.replaceState) {
            var url = window.location.pathname + window.location.search
                        .replace(/([?&])(nueva|editar)=[^&]*&?/g, '$1')
                        .replace(/[?&]$/, '')
                        .replace(/&&/g, '&');
            window.history.replaceState({}, '', url + window.location.hash);
        }
    }

    if (modal) {
        $$('[data-cerrar-modal]', modal).forEach(function (el) {
            el.addEventListener('click', function (ev) {
                ev.preventDefault();
                cerrarModal();
            });
        });

        document.addEventListener('keydown', function (ev) {
            if (ev.key === 'Escape' && modal.classList.contains('is-open')) {
                cerrarModal();
            }
        });

        // Enfoca el primer campo al abrir
        var primero = $('#songName', modal);
        if (primero) { setTimeout(function () { primero.focus(); }, 120); }
    }

    /* ----------------------------------------------------------------------
       6. REPRODUCTOR GLOBAL
       ---------------------------------------------------------------------- */
    var audio = $('#audioGlobal');
    var barra = $('#playerBar');

    if (!audio || !barra) { return; }

    var btnPlayPause = $('#btnPlayPause');
    var btnPrev      = $('#btnPrev');
    var btnNext      = $('#btnNext');
    var elCover      = $('#playerCover');
    var elTitle      = $('#playerTitle');
    var elArtist     = $('#playerArtist');
    var elSeek       = $('#seek');
    var elVol        = $('#volumen');
    var elActual     = $('#timeActual');
    var elTotal      = $('#timeTotal');

    /* Lista de pistas construida a partir de las filas de la tabla */
    var pistas = $$('.fila').map(function (fila) {
        var btn = $('.btn-play', fila);
        if (!btn) { return null; }
        return {
            fila:   fila,
            boton:  btn,
            src:    btn.getAttribute('data-src') || '',
            titulo: btn.getAttribute('data-title') || 'Sin título',
            artista: btn.getAttribute('data-artist') || 'Artista desconocido',
            cover:  btn.getAttribute('data-cover') || '',
            id:     btn.getAttribute('data-id') || '0'
        };
    }).filter(Boolean);

    var actual = -1;
    var contada = false;   // evita contar dos veces la misma carga

    /* Volumen inicial */
    audio.volume = 0.85;
    pintarRango(elVol, 85, 'var(--cyan)');

    /** Resalta la fila de la pista activa */
    function marcarFila(indice) {
        pistas.forEach(function (p, i) {
            p.fila.classList.toggle('is-playing', i === indice);
        });
    }

    /** Carga una pista en el reproductor */
    function cargar(indice, reproducir) {
        if (!pistas.length) { return; }
        if (indice < 0) { indice = pistas.length - 1; }
        if (indice >= pistas.length) { indice = 0; }

        var p = pistas[indice];
        actual  = indice;
        contada = false;

        audio.src = p.src;
        audio.load();

        elTitle.textContent  = p.titulo;
        elArtist.textContent = p.artista;
        if (p.cover) { elCover.src = p.cover; }
        document.title = p.titulo + ' · PoliMusic';

        marcarFila(indice);
        pintarRango(elSeek, 0);
        elActual.textContent = '0:00';
        elTotal.textContent  = '0:00';

        if (reproducir) {
            var promesa = audio.play();
            if (promesa && typeof promesa.catch === 'function') {
                promesa.catch(function () { /* autoplay bloqueado por el navegador */ });
            }
        }
    }

    /** Registra la reproducción en el servidor (api_play.asp) */
    function registrarReproduccion(id) {
        if (!id || id === '0') { return; }

        var datos = new URLSearchParams();
        datos.append('id', id);

        fetch('api_play.asp', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
            body: datos.toString(),
            credentials: 'same-origin'
        })
        .then(function (r) { return r.json(); })
        .then(function (res) {
            if (!res || !res.ok) { return; }
            var celda = document.querySelector('[data-plays-id="' + id + '"]');
            if (!celda) { return; }
            celda.textContent = Number(res.plays).toLocaleString('es-EC');
            celda.classList.add('is-bump');
            setTimeout(function () { celda.classList.remove('is-bump'); }, 420);
        })
        .catch(function () { /* silencioso: el contador no es crítico */ });
    }

    /* --- Botones de la tabla --- */
    pistas.forEach(function (p, i) {
        p.boton.addEventListener('click', function () {
            // Si es la misma pista: alterna play/pausa
            if (i === actual && !audio.paused) {
                audio.pause();
                return;
            }
            if (i === actual && audio.paused && audio.src) {
                audio.play();
                return;
            }
            cargar(i, true);
        });
    });

    /* --- Controles --- */
    if (btnPlayPause) {
        btnPlayPause.addEventListener('click', function () {
            if (actual === -1) {
                cargar(0, true);
                return;
            }
            if (audio.paused) {
                audio.play();
            } else {
                audio.pause();
            }
        });
    }

    if (btnPrev) {
        btnPrev.addEventListener('click', function () {
            if (audio.currentTime > 3) {
                audio.currentTime = 0;
                return;
            }
            cargar(actual - 1, true);
        });
    }

    if (btnNext) {
        btnNext.addEventListener('click', function () {
            cargar(actual + 1, true);
        });
    }

    /* --- Eventos del elemento audio --- */
    audio.addEventListener('play', function () {
        barra.classList.add('is-playing');
        if (btnPlayPause) { btnPlayPause.setAttribute('aria-label', 'Pausar'); }
        if (!contada && actual > -1) {
            contada = true;
            registrarReproduccion(pistas[actual].id);
        }
    });

    audio.addEventListener('pause', function () {
        barra.classList.remove('is-playing');
        if (btnPlayPause) { btnPlayPause.setAttribute('aria-label', 'Reproducir'); }
    });

    audio.addEventListener('loadedmetadata', function () {
        elTotal.textContent = formatoTiempo(audio.duration);
    });

    audio.addEventListener('timeupdate', function () {
        elActual.textContent = formatoTiempo(audio.currentTime);
        if (audio.duration && isFinite(audio.duration)) {
            var pct = (audio.currentTime / audio.duration) * 100;
            elSeek.value = Math.round(pct * 10);
            pintarRango(elSeek, pct);
        }
    });

    audio.addEventListener('ended', function () {
        contada = false;
        cargar(actual + 1, true);
    });

    audio.addEventListener('error', function () {
        elTitle.textContent  = 'No se pudo cargar el audio';
        elArtist.textContent = 'Verifica la ruta del archivo en la base de datos';
        barra.classList.remove('is-playing');
    });

    /* --- Barra de progreso --- */
    if (elSeek) {
        elSeek.addEventListener('input', function () {
            if (!audio.duration || !isFinite(audio.duration)) { return; }
            var pct = elSeek.value / 10;
            pintarRango(elSeek, pct);
            audio.currentTime = (pct / 100) * audio.duration;
        });
    }

    /* --- Volumen --- */
    if (elVol) {
        elVol.addEventListener('input', function () {
            var v = Number(elVol.value);
            audio.volume = v / 100;
            pintarRango(elVol, v, 'var(--cyan)');
        });
    }

    /* --- Atajos de teclado --- */
    document.addEventListener('keydown', function (ev) {
        var tag = (ev.target.tagName || '').toLowerCase();
        if (tag === 'input' || tag === 'select' || tag === 'textarea') { return; }
        if (modal && modal.classList.contains('is-open')) { return; }

        if (ev.code === 'Space') {
            ev.preventDefault();
            if (actual === -1) { cargar(0, true); return; }
            if (audio.paused) { audio.play(); } else { audio.pause(); }
        }
    });

})();
