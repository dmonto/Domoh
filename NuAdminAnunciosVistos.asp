<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrGrid.asp" -->
<head>
    <!-- #include file="IncQuCabecera.asp" -->
    <title>Anuncios Vistos Por Dia</title>
</head>
<%
	dim Grid
		
	SetupGrid Grid
	Grid.SQL = "SELECT 0, fecha, SUM(visitas) AS anunciosVistos FROM AnunciosVistos GROUP BY fecha ORDER BY fecha DESC"		
	Grid.sTabla="AnunciosVistos"
	Grid.Cols(1).Hidden=True
%>
<div class=container>
	<div class=logo>Anuncios Vistos Por Dia</div>
	<a href=NuAdminAnunciosVistos.asp>NuAdminAnunciosVistos.asp</a> <a title='Generar Excel con los Todos Anuncios Vistos' href='NuGrabaExcel.asp?tabla=AnunciosVistos'>Excel Anuncios Vistos</a></div>
<%	Grid.Display %>
<!-- #include file="IncPie.asp" -->