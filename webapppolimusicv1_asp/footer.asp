    </main>

    <footer class="pie">
        <div class="pie__inner">
            <span><b>PoliMusic</b> · Gestión de biblioteca musical</span>
            <span class="pie__sep" aria-hidden="true">·</span>
            <span>ASP Clásico + SQL Server</span>
            <span class="pie__sep" aria-hidden="true">·</span>
            <span><%= Year(Date()) %></span>
        </div>
    </footer>

    <script src="js/app.js"></script>
</body>
</html>
<%
' Libera la conexión al finalizar la página (si seguía abierta)
CerrarConexion
%>
