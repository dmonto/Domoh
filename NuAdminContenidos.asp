<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrGrid.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim Grid
	
	SetupGrid Grid
	Grid.SQL = "SELECT * FROM Contenidos ORDER BY pagina, posicion"
	Grid.sTabla="Contenidos"
%>
<div class=container>
	<div class=banner>
		<h1>Contenidos</h1></div>
	<div class=main>
		<a title='Exportar a Excel toda la Tabla de Contenidos' href='NuGrabaExcel.asp?tabla=Contenidos'>Excel Contenidos</a>
<%	Grid.Display %>
<form method=post name=frm action=NuAdminContenidosGraba.asp enctype=multipart/form-data>
	Página: 
    <select title='Paginas disponibles para contenidos' name=pagina size=1> 
        <option value=''>--Selecciona una--</option>
<%		
	rst.Open "SELECT DISTINCT pagina FROM Contenidos ORDER BY 1", sProvider
	while not rst.Eof
%>
		<option value=<%=rst("pagina")%>><%=rst("pagina")%></option>
<%
		rst.Movenext
	wend
	rst.Close
%>				
    </select></div>
<div>Posición: <input title='0 para portada o número de artículo' name=posicion size=3 /></div>
<div>
	<div>
        Banners: 
        <input title='Elige ficheros para subir a la web' type=file name=banner1 />
        <input title='Elige ficheros para subir a la web' type=file name=banner2 />
		<input title='Elige ficheros para subir a la web' type=file name=banner3 />
        <input title='Elige ficheros para subir a la web' type=file name=banner4 />
		<input title='Elige ficheros para subir a la web' type=file name=banner5 /></div>
	<div>URL Destino/HTML: <input title='URL del banner' name=html size=50 /></div></div>
<div>Texto: <textarea title='Texto de la página' cols=70 rows=20 name=texto></textarea></div>
<div><input title=Grabalo type=submit value='Insertar Contenido' class=btnForm /></div></form></div></div>
<!-- #include file="IncPie.asp" -->