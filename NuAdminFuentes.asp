<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<head><title>Fuentes de Usuarios</title></head>
<%
	dim numFuentes, i
	
	sQuery= "SELECT COUNT(*) AS numFuentes FROM Fuentes"
	rst.Open sQuery, sProvider
	numFuentes=rst("numFuentes")
	rst.Close

	sQuery= "SELECT f.id, [desc], descen, COUNT(*) AS numUsuarios FROM Fuentes f LEFT JOIN Usuarios u ON f.id=u.fuente "
	sQuery= sQuery & "GROUP BY f.id, [desc], f.descen ORDER BY 4 DESC "
	rst.Open sQuery, sProvider
%>
<form method=post name=frm action=NuAdminFuentesGraba.asp>
<div class=container>
	<div class=logo>Fuentes</div>
    <div class=banner><a title='Exporta a Excel toda la tabla de Fuentes' href='NuGrabaExcel.asp?tabla=Fuentes'>Excel Fuentes</a><a title='Acceso a la tabla de Fuentes' href='NuAdminTabla.asp?tabla=Fuentes'>Tabla Fuentes</a></div>
        <% PagCabeza numFuentes, "Fuentes de Usuario" %></div>
<div title='Nombre de la fuente de usuarios'>Fuente</div><div title='Número de usuarios'>Usuarios</div><div title='Acciones Posibles'>Qué hago con ella</div>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
%>
<div>
    <div title='Nombre Español'><%=rst("desc")%></div><div title='Nombre Inglés'><%=rst("descen")%></div><div title='Número de Usuarios de esa Fuente'><%=rst("numUsuarios")%></div>
	<div title='Acciones Posibles'><select name=<%=rst("id")%>><option selected value=0>-Elige Una-</option><option value=Borrar>Borrar</option></select></div></div>
<%		
		end if
		rst.Movenext
	next
%>				
<div title='Nombre Español'><input name=desc /></div><div title='Nombre Inglés'><input name=descen /></div><div title='Número de Usuarios de esa Fuente'></div><div title='Borrar?'></div>
<div><% PagPie numFuentes, "NuAdminFuentes.asp?" %></div>
<div><input title=Grabar type=submit value='Insertar Referencia' class=btnForm /></div></form>
<!-- #include file="IncPie.asp" -->
