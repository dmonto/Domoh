<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim numInquilinos, i, sIdioma, sWhere

	Session("numResults")=1000
	sWhere=" WHERE a.activo='Si' AND u.activo='Si' AND (a.fechaultimamodificacion < DATEADD(DAY,-31,GETDATE()) "
	sWhere=sWhere & " AND (a.fechaavisobaja IS NULL OR a.fechaavisobaja < DATEADD(DAY,-2,GETDATE())))"
	
	sQuery= "SELECT COUNT(*) AS numInquilinos FROM (Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario "
	sQuery= sQuery & sWhere 
	rst.Open sQuery, sProvider
	numInquilinos=rst("numInquilinos")
	rst.Close

	sQuery= "SELECT a.activo AS aActivo, a.id AS aId, i.ciudad AS iCiudad, u.foto AS uFoto, * "
	sQuery= sQuery & " FROM (Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario ORDER BY a.fechaultimamodificacion"
	rst.Open sQuery, sProvider
%>
<form method=post name=frm action=NuAdminInquilinosViejosGraba.asp>
<div class=container>
	<div class=main>
		<h1>Inquilinos Viejos</h1>
    	<h2><a title='Acceso a la tabla completa de Inquilinos' href='NuAdminTabla.asp?tabla=inquilinos&order=fechaultimamodificacion'>Tabla Inquilinos</a></h2>
    	<a title='Generar Excel con toda la tabla de inquilinos' href='NuGrabaExcel.asp?tabla=Inquilinos'>Excel Inquilinos</a>
    <h3>
<%	if numInquilinos=0 then %>
        Nada viejo Carretier.
<%
		Response.End
	else
		PagCabeza numInquilinos, "Inquilino"
	end if 
%>
	    </h3>
    <h3><input title='Mandar mails, borrar...' type=submit value='Grabar Cambios' class=btnForm /><% PagPie numInquilinos, "NuAdminInquilinosViejos.asp?" %></h3>
    <div class=row>
    	<div class=middle title='Ciudad/Zona'>Dónde</div><div class=middle title='Título'>Anuncio</div><div class=middle title='¿El usuario ha confirmado?'>Activo</div><div class=middle title='Pulsa para enviar un e-mail'>Quién</div>
        <div title='Teléfono del que puso el anuncio'>Teléfonos</div><div title='Máximo alquiler'>Precio</div><div title='Día que se creó el anuncio'>Fecha Alta</div>
		<div title='Última Actualización'>Fecha Renovación</div><div title='Pulsa para ver la foto'>Foto?</div>
        <div title='Activa para enviar mail a este'>Mandar Mail Aviso</div><div title='Activa para Desactivar este'>Desactivar</div></div>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
			if UCase(rst("idiomaes"))="ON" then 
				if UCase(rst("idiomaen"))="ON" then sIdioma="b" else sIdioma="e" 
			else 
				sIdioma="i"
			end if
%>
    <div>
        <div title='Ciudad/Zona'><a href="javascript:detalle('QuAnuncioDetalle.asp?tabla=Inquilinos&id=<%=rst("aId")%>')"><%=rst("iCiudad")%></a></div>
    	<div title='Título'><%=rst("cabecera")%></div><div title='¿El usuario ha confirmado?'><%=rst("aActivo")%></div>
    	<div title='Pulsa para enviar un e-mail'><a href='mailto:<%=rst("email")%>'><%=rst("email")%></a></div>
    	<div title='Teléfono del que puso el anuncio'><%=rst("tel1")%></div><div title='Máximo alquiler'>€<%=rst("maximo")%></div>
	    <div title='Día que se creó el anuncio'><%=rst("fechaalta")%></div><div title='Última Actualización'><%=rst("fechaultimamodificacion")%></div>
	    <div title='Pulsa para ver la foto'>
<%			if rst("uFoto") <> "" then %>		
		    <a href="javascript:detalle('http://domoh.com/<%=rst("uFoto")%>')"><img src="http://domoh.com/mini<%=rst("uFoto")%>" alt='Pulsa para ver las fotos'/></a>
<%			end if %>    
		    </div>
<%			if IsNull(rst("fechaavisobaja")) then %>
	    <div title='Activa para enviar mail a este'><input type=checkbox name=m<%=sIdioma & rst("aId")%> /></div>
<%			else %>
	    <div title='Fecha del Aviso'><%=rst("fechaavisobaja")%></div><div title='Activa para Desactivar este'><input type=checkbox name=b<%=sIdioma & rst("aId")%>/></div>
<%			end if %>
	        </div>
<%
		end if
		rst.Movenext
	next
%>				
        <div><% PagPie numInquilinos, "NuAdminInquilinosViejos.asp?" %></div>
	    <div><input title=Grabalo type=submit value='Grabar Cambios' class=btnForm /></div></div></div>
</form>	
<!-- #include file="IncPie.asp" -->
