<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim sWhere, numMisc, i, numCatMisc, sNombreCatMisc, sHora, sIdioma, numProvincia, sReq

	numCatMisc=0
	if Request("catmisc")<>"" then
		numCatMisc=CLng(Request("catmisc"))
		sReq="catmisc=" & Request("catmisc") & "&"
	end if

	numProvincia=0
	if Request.Form("Provincia")<>"" or Request.QueryString("Provincia")<>"" then
		numProvincia=CLng(Request("provincia"))
		Response.Cookies("ProvinciaViv")=numProvincia
		Response.Cookies("ProvinciaViv").Expires=Now+30
		sReq=sReq & "provincia=" & Request("provincia") & "&"
	end if

	sWhere= "WHERE a.activo='Si' AND u.activo='Si' "
	if numCatMisc<>0 then sWhere= sWhere & " AND m.catmisc = " & numCatMisc	else sWhere= sWhere & " AND m.catmisc <> 0 "
	if numProvincia<>0 then sWhere= sWhere & " AND a.provincia = " & numProvincia else sWhere= sWhere & " AND a.provincia <> 0 "
	
	sQuery="SELECT COUNT(*) AS numMisc FROM (Misc m INNER JOIN Anuncios a ON m.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en NuMiscBusco", sQuery & " - " & Err.Description
	numMisc=rst("numMisc")
	rst.Close

	if numCatMisc <> 0 then 
		rst.Open "SELECT nombre FROM CatMisc WHERE id=" & numCatMisc, sProvider
		sNombreCatMisc=rst("nombre")
		Session("CatMiscNombre")=sNombreCatMisc
		rst.Close
	end if
	
	if Session("Idioma")="En" then sIdioma="a.idiomaen" else sIdioma="a.IdiomaEs" 
	sQuery="SELECT a.id AS aId, cm.nombre as cmNombre, * FROM ((Misc m INNER JOIN Anuncios a ON m.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario) "
	sQuery=sQuery & " LEFT JOIN CatMisc cm ON m.CatMisc=cm.id " & sWhere & " ORDER BY UPPER(" & sIdioma & ") DESC, a.fechaultimamodificacion Desc"
	rst.Open sQuery, sProvider
%>
<body onload="window.parent.location.hash='top';"> 
<table>
	<tr>
        <td colspan=3 rowspan=2>
<%  if Request("msg")<>"" then %>
	<p><%=Request("msg")%></p>
<%  end if %>
</td><!--#include file="IncNuSubMenu.asp"--></tr>
	<tr>
		<td colspan=5>
<%	
	if numMisc=0 then
		if Request("tipobusqueda") = "nuevas" then
%>
		<b>Nada nuevo desde tu última visita. Puedes ver todos si pulsas <a title='Ver todos los anuncios disponibles' href='NuMiscBusco.asp?tipobusqueda=todas'>aquí</a>.</b>
<%		else %>
		<b>No hemos encontrado anuncios con esos criterios.</b>
<%		end if 
		Response.End
	end if

	PagCabeza numMisc, "Anuncio"
%>
	<br/></td></tr>
	<tr>	
		<td title='Última Modificación'><b>Publicado</b></td><td title='Idioma del Anuncio'><p><b>Idioma</b></p></td><td title='Los anuncios están clasificados'><b>Categoría</b></td>
		<td title='Pulsa para ver el anuncio'><b>Titulo</b></td><td title='Pulsa para ver la foto'><b>Foto</b></td><td title='Pulsa para ver el flyer'><b>Flyer</b></td></tr>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
%>
	<tr onclick="javascript:detalle('TrAnuncioDetalle.asp?tabla=Misc&id=<%=rst("aId")%>')">
		<td title='Última Modificación'>
<%
			sHora=Hour(rst("fechaultimamodificacion")) & ":" & Minute(rst("fechaultimamodificacion"))
			if sHora="0:0" then
				Response.Write FormatDateTime(rst("fechaultimamodificacion"),2)
			else
				Response.Write FormatDateTime(rst("fechaultimamodificacion"),2) & " " & FormatDateTime(rst("fechaultimamodificacion"),4)
			end if
%>
		</td>
   		<td title='Idioma del Anuncio'>
<%			if UCase(rst("idiomaes"))="ON" then %>
			<img src=images/Espanol.gif width=20 alt='Anuncio en Español'/>
<%			
            end if 
			if UCase(rst("idiomaen"))="ON" then 
%>
			<img src=images/Ingles.gif width=20 alt='Anuncio en Inglés'/>
<%			end if %>
		</td>
   		<td title='Los anuncios están clasificados'><%=rst("cmNombre")%></td>
    	<td title='Pulsa para ver el anuncio'><a href="javascript:detalle('TrAnuncioDetalle.asp?tabla=Misc&id=<%=rst("aId")%>')"><%=rst("cabecera")%></a></td>
		<td title='Pulsa para ver la foto'>
<%			if rst("foto") = "Si" then %>
			<a href="javascript:detalle('http://domoh.com/<%=rst("foto")%>')"><img src=images/Camara.gif height=50 alt='Pulsa para ver las fotos'/></a>
<%			elseif rst("foto") <> "" then %>		
		    <a href="javascript:detalle('http://domoh.com/<%=rst("foto")%>')"><img src='http://domoh.com/mini<%=rst("foto")%>' height=50 alt='Pulsa para ver las fotos'/></a>
<%			end if %>    
		</td>
		<td title='Pulsa para ver el flyer'>
<%			if rst("flyer") <> "" then %>		
			<a href="javascript:detalle('http://domoh.com/TrMiscFlyer.asp?flyer=<%=rst("flyer")%>')"><img src=http://domoh.com/<%=rst("flyer")%> height=50 alt='Pulsa para ver las fotos'/></a>
<%			end if %>    
			</td></tr>
<%
		end if
		rst.Movenext
	next

	PagPie numMisc, "NuMiscBusco.asp?" & sReq 
%>				
</tableIncPie.asp" -->
</body>
