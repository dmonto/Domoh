<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrGrid.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim Grid, sWhere
	
	if Request("tipo")<>"menu" then
		SetupGrid Grid

		sWhere= " WHERE id IS NOT NULL  "
		if Request("Activo")<>"" then sWhere= sWhere & " AND activo='" & Request("activo") & "'"
		if Request("email")<>"" then sWhere= sWhere & " AND usuario IN (SELECT usuario FROM usuarios WHERE email = '" & Request("email") & "') "
		if Request("id")<>"" then sWhere= sWhere & " AND id = " & Request("id")
	
		Grid.SQL = "SELECT * FROM Anuncios " & sWhere & " ORDER BY fechaultimamodificacion DESC"
		Grid.sTabla="Anuncios"

		Grid.ExtraFormItems = "<input type=hidden name=activo value=" & Request("activo") & ">" 
		Grid.ExtraFormItems = Grid.ExtraFormItems & "<input type=hidden name=email value=" & Request("email") & ">" 
		Grid.ExtraFormItems = Grid.ExtraFormItems & "<input type=hidden name=id value=" & Request("id") & ">" 
	end if
%>
<body>
<div class=container>
	<form action=TrAdminAnunciosGrid.asp>
    	<h1>Buscar piso num: <input title='Buscar un piso concreto' name=id /> e-mail: <input title='Buscar los pisos de un usuario' name=email /><input title=Dele type=submit value=Buscar class=btnLogin /></h1></form></div>
	<nav>
        Pisos <a title='Vuelca a Excel la tabla completa de Pisos' href='NuGrabaExcel.asp?tabla=Pisos'>Excel Pisos</a><a title='Ver solo los pisos activos' href='TrAdminAnunciosGrid.asp?Activo=Si'>Solo Activos</a>
		<a title='Ver todos los pisos' href=TrAdminAnunciosGrid.asp>Todos</a><a title='Volver a la pantalla de publicar' href='NuAdminPisos.asp?tipo=todos'>Tabla Simple</a></nav>
<%	if Request("tipo")<>"menu" then Grid.Display %>
<!-- #include file="IncPie.asp" -->																																								
</body>
