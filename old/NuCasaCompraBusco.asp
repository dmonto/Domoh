<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim sWhere, sOrder, numPisos, i, sNombreProvincia, numProvincia, sHora, sIdioma, sReq

	numProvincia=0
	if Request.Form("provincia")<>"" or Request.QueryString("provincia")<>"" then
		numProvincia=CLng(Request("provincia"))
		Response.Cookies("ProvinciaViv")=numProvincia
		Response.Cookies("ProvinciaViv").Expires=Now+30
		sReq="provincia=" & Request("provincia") & "&"
	end if

	sWhere= "WHERE a.activo='Si' AND u.activo='Si' AND precio <> 0 " 
	if numProvincia<>0 then sWhere= sWhere & " AND a.provincia = " & numProvincia else sWhere= sWhere & " AND a.provincia <> 0 "
	
	if Request("precio") <> "" then 
		sWhere= sWhere & " AND precio <= " & Request("precio")
		sReq=sReq & "precio=" & Request("precio") & "&"
	end if
	
	sQuery="SELECT COUNT(*) AS numPisos FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON u.usuario=a.usuario  " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
	numPisos=rst("numPisos")
	rst.Close

	if Session("Idioma")="En" then sIdioma="idiomaen" else sIdioma="idiomaes"
	
	if Request("orden")="fecha" then
		sOrder= "destacado DESC, fechaultimamodificacion DESC, " & sIdioma & " DESC, precio"
	elseif Request("orden")="precio" then
		sOrder= "destacado DESC, precio, " & sIdioma & " DESC, fechaultimamodificacion DESC"
	else
		sOrder= "destacado DESC, " & sIdioma & " DESC, precio"
	end if
	
	sQuery="SELECT a.id AS aId, * FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON u.usuario=a.usuario  " & sWhere & " ORDER BY " & sOrder
	rst.Open sQuery, sProvider
%>
<body onload="window.parent.location.hash='top';"> 

<table>
	<tr><td colspan=3 rowspan=2 width=500>
		<% if Request("msg")<>"" then %><p style='font-size:16pt;color:red;border:1px solid red'><%=Request("msg")%></p>
<% end if %>
</td><!--#include file="IncNuSubMenu.asp"-->
	<tr>
		<td colspan=3 width=500>&nbsp;</td><td colspan=2><ul type=circle><li><a title='Poner un Anuncio' href=NuCasaCompraOfrezcoFront.asp><b>Publicar Piso o Casa</b></a></li></ul></td></tr>
	<tr>
		<td colspan=6>
<%	if numPisos=0 then %>
			<b>No hemos encontrado viviendas con esos criterios.</b>
<%		
		Response.End
	end if

	PagCabeza numPisos, "Vivienda"
%>
	</td></tr>
	<tr>	
		<td><a title='Ver las más recientes primero' href=NuCasaCompraBusco.asp?orden=fecha&<%=sReq%>>por Fecha</a></td>
    	<td><a title='Ver primero las que estén en español' href='NuCasaCompraBusco.asp?orden=idioma&<%=sReq%>'>por Idioma</a></td>
    	<td>&nbsp;</td><td><a title='Ver las más baratas primero' href='NuCasaCompraBusco.asp?orden=precio&<%=sReq%>'>por Precio</a></td><td>&nbsp;</td><td>&nbsp;</td></tr>
	<tr>	
		<td title='Última Modificación'><b>Publicado</b></td><td title='¿Está el anuncio en español?'><b>Idioma</b></td><td title='Ciudad/Zona'><b>Dónde</b></td>
    	<td title='Precio en miles de euros'><img alt='Precio' src=images/Eurito.gif width=25 height=25></td>
    	<td title='¿Qué es?' align=center><b>Tipo</b></td><td title='Haz click para ver la foto' align=center><b>Foto</b></td></tr>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
			if rst("Destacado") then
%>
	<tr valign=middle bgcolor=blue onmouseover="this.bgColor='cyan';" onmouseout="if (this.rowIndex % 2) {this.bgColor='blue'} else {this.bgColor='#FFFFFF'};">
<%			else %>
	<tr valign=middle <%if (i mod 2)=1 then Response.Write "bgcolor=#CC99FF"%> onmouseover="this.bgColor='cyan';" onmouseout="if (this.rowIndex % 2) {this.bgColor='#CC99FF'} else {this.bgColor='#FFFFFF'};">
<%			end if %>    
   		<td title='Última Modificación' width=150 align=left valign=middle >
<%
			sHora=Hour(rst("fechaultimamodificacion")) & ":" & minute(rst("FechaUltimaModificacion"))
			if sHora="0:0" then
				Response.Write formatdatetime(rst("FechaUltimaModificacion"),2)
			else
				Response.write formatdatetime(rst("FechaUltimaModificacion"),2) & " " & formatdatetime(rst("FechaUltimaModificacion"),4)
			end if
%>
    	</td>
   		<td title="¿Está el anuncio en español?" width=50 align=center valign="middle" >
<%			if UCase(rst("idiomaes"))="ON" then %>
			<img src=images/Espanol.gif width=20 alt="Anuncio en Español">
<%			end if %>
<%			if UCase(rst("idiomaen"))="ON" then %>
			<img src=images/Ingles.gif width=20 alt="Anuncio en Inglés">
<%			end if %>
		</td>
   		<td title="Ciudad/Zona" width=300 valign=middle>
    		<a href=javascript:detalle('TrAnuncioDetalle.asp?tabla=Pisos&ID=<%=rst("aId")%>')><%=rst("cabecera")%></a></td>
	   <td title="Precio en miles de euros" width=90 valign=middle align=center >
			<%=rst("Precio")%> mil <font face=Arial>€</font>
    </td>
    <td title="¿Qué es?" valign=middle align="center" ><%=rst("Tipo")%>
    </td>
    <td title="Haz click para ver la foto" valign=middle align="center" >
<%			if rst("Foto") = "Si" then %>
		<a href=javascript:foto('<%=rst("aId")%>')><img src="images/Camara.gif" height=50 alt="Pulsa para ver las fotos"></a>
<%			elseif rst("Foto") <> "" then %>		
		<a href=javascript:foto('<%=rst("aId")%>')><img src="http://domoh.com/mini<%=rst("Foto")%>" height=50 alt="Pulsa para ver las fotos"></a>
<%			End if %>    
    </td></tr>
<%
		end if
		rst.movenext
	next

	PagPie numPisos, "NuCasaCompraBusco.asp?" & sReq & "orden=" & request("orden") & "&"
%>
    </table>
<!-- #include file="IncPie.asp" -->