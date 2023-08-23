<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrGrid.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim Grid, sWhere
	
	if Request("tipo")<>"menu" then
		SetupGrid Grid

		sWhere= " WHERE u.id > 0 "
		if Request("tipousu")<>"" then sWhere= sWhere & " AND u.tipo='" & Request("tipousu") & "'"
		if Request("usuario")<>"" then sWhere= sWhere & " AND usuario='" & Request("usuario") & "'"
		if Request("email")<>"" then sWhere= sWhere & " AND email='" & Request("email") & "'"
		if Request("tipo")="nuevos" then sWhere= sWhere & " AND activo='Mail'"
	
		Grid.SQL = "SELECT * FROM Usuarios u " & sWhere & " ORDER BY fechaalta DESC"		

		Grid.sTabla="Usuarios"
		Grid.ExtraFormItems = "<input type=hidden name=tipousu value=" & Request("Tipousu") & ">" 
		Grid.ExtraFormItems = Grid.ExtraFormItems & "<input type=hidden name=usuario value=" & Request("usuario") & " autocompete=username >" 
		Grid.ExtraFormItems = Grid.ExtraFormItems & "<input type=hidden name=email value=" & Request("email") & ">" 		
	end if
%>
<div class=container>
	<div class=main>
		<form action=NuAdminUsuarios.asp>
			Buscar usuario: <input title='Pon el nombre de usuario que buscas' name=usuario /> e-mail: <input title='Pon el e-mail del usuario que buscas' name=email/><input type=submit value=Buscar /></form>
		<div>
			<div>Usuarios <a title='Pasa a Excel toda la tabla de Usuarios' href='NuGrabaExcel.asp?tabla=Usuarios'>Excel Usuarios</a></div>
			<div><a title='Ver sólo Caseros' href='NuAdminUsuarios.asp?Tipousu=Casero'>Solo Caseros</a></div>
			<div><a title='Ver sólo Inquilinos' href='NuAdminUsuarios.asp?Tipousu=Inquilino'>Solo Inquilinos</a></div>
			<div><a title='Ver todos los Usuarios' href=NuAdminUsuarios.asp>Todos</a></div></div></div>
	<%	if Request("Tipo")<>"menu" then Grid.Display %>
<!-- #include file="IncPie.asp" -->
