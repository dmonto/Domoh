<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrGrid.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim Grid, sWhere
	
	if Request("Tipo")<>"menu" then
		SetupGrid Grid	
		sWhere= " WHERE a.id IS NOT NULL  "
		if Request("Activo")<>"" then sWhere= sWhere & " AND a.activo='" & Request("activo") & "'"
		if Request("email")<>"" then sWhere= sWhere & " AND a.usuario IN (SELECT u.usuario FROM usuarios u WHERE u.email = '" & Request("email") & "') "
		if Request("id")<>"" then sWhere= sWhere & " AND a.id = " & Request("id")
	
		Grid.SQL = "SELECT a.id,usuario,activo,tipo,destacado,fechaalta,fechaultimamodificacion,fechaavisobaja,ciudad,provincia,"
		Grid.SQL = Grid.SQL & " zona,dir1,dir2,ciudadnombre,cp,descripcion,gente,idiomaes,idiomaen,fuma,mascota,rentaviv,rentavacas,precio,foto "
		Grid.SQL = Grid.SQL & " FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id " & sWhere & " ORDER BY fechaultimamodificacion DESC"
		Grid.sTabla="Pisos"
		
		Grid.ExtraFormItems = "<input type=hidden name=activo value=" & Request("activo") & "><input type=hidden name=email value=" & Request("email") & ">" 
		Grid.ExtraFormItems = Grid.ExtraFormItems & "<input type=hidden name=id value=" & Request("id") & ">" 
		Grid("foto").AttachExpression "<img src=mini[[foto]]>" 
		Grid("descripcion").AttachTextArea 30, 60 
		Grid("gente").AttachTextArea 30, 60 
	end if
%>
<form action=NuAdminPisosGrid.asp><h1>Buscar piso num: <input title='Buscar un piso concreto' name=id /> e-mail: <input title='Buscar los pisos de un usuario' name=email /><input type=submit value=Buscar /></h1></form>
<div class=container>
    Pisos
	<a title='Vuelca a Excel la tabla completa de Pisos' href='NuGrabaExcel.asp?tabla=Pisos'>	Excel Pisos</a>
	<a title='Ver solo los pisos activos' href='NuAdminPisosGrid.asp?Activo=Si'>Solo Activos</a>
	<a title='Ver todos los pisos' href=NuAdminPisosGrid.asp>Todos</a>
	<a title='Volver a la pantalla de publicar' href='NuAdminPisos.asp?tipo=todos'>Tabla Simple</a>
<%	if Request("Tipo")<>"menu" then Grid.Display %></div>
<!-- #include file="IncPie.asp" -->
