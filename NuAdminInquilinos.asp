<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim numInquilinos, i, sProvincias

	Session("Pantalla")=Request.Servervariables("SCRIPT_NAME") & "?tipo=" & Request("tipo")
	sQuery= "SELECT COUNT(*) AS numInquilinos FROM (Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario "
	if Request("tipo")="nuevos" then sQuery= sQuery & " WHERE a.activo='Si' AND a.provincia=0"
	rst.Open sQuery, sProvider
	numInquilinos=rst("numInquilinos")
	rst.Close

	rst.Open "SELECT * FROM Provincias ORDER BY nombre", sProvider
	while not rst.Eof
		sProvincias=sProvincias & "<option value=" & rst("id") &">Publicar en " & rst("nombre")
		rst.Movenext
	wend
	rst.Close

	sQuery= "SELECT a.id AS aId, u.activo AS uActivo, u.foto AS uFoto, i.tipoviv AS iTipoViv, i.ciudad AS iCiudad, i.maximo AS iMaximo, * "
	sQuery= sQuery & " FROM (Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario  "
	if Request("tipo")="nuevos" then sQuery= sQuery & " WHERE a.activo='Si' AND a.provincia=0 "
	rst.Open sQuery, sProvider
%>
<form method=post name=frm action=NuAdminAnunciosGraba.asp>
	<input type=hidden name=tabla value=Inquilinos /><input type=hidden name=campo value=provincia /><input type=hidden name=tipo value=<%=Request("tipo")%> />
<div class=container>
	<h1 class=banner>Nuevos Buscadores de Pisos</h1>
    <div class=row>
        <div class=col-4><a title='Ver toda la tabla de Buscadores de Pisos' href='NuAdminTabla.asp?tabla=Inquilinos&order=fechaultimamodificacion'>Tabla Busca Pisos</a></div>
	    <div class=col-8><a title='Generar Excel con todos los Inquilinos' href='NuGrabaExcel.asp?tabla=Inquilinos'>Excel Inquilinos</a></div></div>
	<div class=row>
        <h2>
<%	if numInquilinos=0 then %>
		Nada nuevo Carretier.
<%
		Response.End
	else
		PagCabeza numInquilinos, "Busca Piso"
	end if 
%>
	    </h2></div>
	<div class=col-12><input title='Publicar, Borrar...' type=submit value='Grabar Cambios' class=btnForm /><% PagPie numInquilinos, "NuAdminInquilinos.asp?" %></div>
	<div class=row>
		<div class=middle title='¿Que busca?'>Tipo</div><div class=middle title='Nombre del anunciante'>Nombre</div><div class=middle title='¿El usuario ha confirmado?'>Usuario Activo</div>
        <div class=middle title='Pulsa para mandarle un mail'>Usuario</div><div class=middle title='Teléfono del que ha puesto el anuncio'>Teléfonos</div><div class=middle title='Cabecera del Anuncio'>Titulo</div>
        <div class=middle title='Día que se creó el anuncio'>Fecha Alta </div><div class=middle title='Dónde busca'>Ciudad</div><div class=middle title='Pulsa para ver la foto'>Foto</div>
        <div class=middle title='Renta máxima que asumiría'>Máximo</div><div class=middle title='Publicar, Borrar...'>Qué hago con él</div></div>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
%>
	<div class=row>
		<div class=middle title='¿Qué busca?'><%=rst("iTipoViv")%>
<%			if UCase(rst("idiomaes"))="ON" then %>
			<img src=images/Espanol.gif alt='Anuncio en Español'/>
<%			
            end if 
            if UCase(rst("idiomaen"))="ON" then 
%>
			<img src=images/Ingles.gif alt='Anuncio en Inglés'/>
<%			end if %>
		</div>
		<div class=middle title='Nombre del anunciante'><a href="javascript:detalle('QuAnuncioDetalle.asp?tabla=Inquilinos&id=<%=rst("aId")%>')"><%=rst("nombre")%></a></div>
        <div title='¿El usuario ha confirmado?'><%=rst("uActivo")%></div><div class=middle title='Pulsa para mandarle un mail'><a href='mailto:<%=rst("email")%>'><%=rst("usuario")%></a></div>
        <div title='Teléfono del que ha puesto el anuncio'><%=rst("tel1")%></div><div class=middle title='Cabecera del Anuncio'><%=rst("cabecera")%></div><div title='Día que se creó el anuncio'><%=rst("fechaalta")%></div>
        <div title='Dónde busca'><%=rst("iCiudad")%></div>
		<div class=middle title='Pulsa para ver la foto'>
<%			if rst("uFoto")="Si" then %>
			<a href="javascript:detalle('http://domoh.com/<%=rst("uFoto")%>')"><img src=images/Camara.gif alt='Pulsa para ver las fotos'/></a>
<%			elseif rst("uFoto")<>"" then %>		
			<a href="javascript:detalle('http://domoh.com/<%=rst("uFoto")%>')"><img src='http://domoh.com/mini<%=rst("uFoto")%>' alt='Pulsa para ver las fotos'/>
			</a>
<%			end if %>    
    	</div>
		<div class=middle title='Renta máxima que asumiría'><%=rst("iMaximo")%></div>
		<div class=middle title='Publicar, Borrar...'>
			<select name=<%=rst("aId")%> 
				onchange="if(value=='Modificar') location='TrCasaRegDemandaFront.asp?op='+value+'&id=<%=rst("aId")%>';"
                >
				<option selected value=0>-Elige Una-</option>
<%          if rst("uActivo")="Si" then %>
				<option value=Borrar>Ocultar</option>
<%          else %>
				<option value=Reactivar>Reactivar</option>
<%          end if %>
				<option value=Modificar>Modificar</option>
				<option value=Eliminar>Eliminar Definitivamente</option>
				<%=sProvincias%>
			</select></div></div>
<%
		end if
		rst.Movenext
	next
%>				
	<div class=row><% PagPie numInquilinos, "NuAdminInquilinos.asp?" %></div>
	<div class=row><input title=Grabar type=submit value='Grabar Cambios' class=btnForm /></div>
</div></form>	
<!-- #include file="IncPie.asp" -->