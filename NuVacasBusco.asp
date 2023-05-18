<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncNuMail.asp" -->
<head>
    <!-- #include file="IncTrCabecera.asp" -->
    <title>Vacaciones - Resultados</title>
</head>
<body onload="window.parent.location.hash='top';"> 
<%
	dim sWhere, sOrder, numPisos, i, sNombreProvincia, numProvincia, numPais, sIdioma, sReq
 	
	if Request.Form("provincia")<>"" or Request.QueryString("provincia")<>"" then
		numProvincia=CLng(Request("provincia"))
		sReq="provincia=" & Request("provincia") & "&"
	end if

	if Request("pais")<>"" then
		numPais=CLng(Request("pais"))
		sReq=sReq & "pais=" & Request("pais") & "&"
	end if

	sWhere= "WHERE a.activo='Si' AND rentavacas <> 0 AND u.activo='Si' "
	if numProvincia<>0 then 
		sWhere= sWhere & " AND a.provincia = " & numProvincia
	elseif numPais<>0 then 
		sWhere= sWhere & " AND a.provincia IN (SELECT id FROM Provincias WHERE pais= " & numPais & ")"
	else
		sWhere= sWhere & " AND a.provincia <> 0 "
	end if
		
	Err.Clear
	sQuery="SELECT COUNT(*) AS numPisos FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
	numPisos=rst("numPisos")
	rst.Close

	if Session("Idioma")="En" then sIdioma="a.idiomaen" else sIdioma="a.idiomaes" 
	if Request("orden")="fecha" then
		sOrder= sIdioma & " DESC, fechaultimamodificacion DESC, rentavacas"
	elseif Request("orden")="precio" then
		sOrder= sIdioma & " DESC, rentaviv, fechaultimamodificacion DESC"
	else
		sOrder= "a.destacado DESC, " & sIdioma & " DESC, rentavacas"
	end if
	
	sQuery="SELECT a.id AS aId, p.tipo AS pTipo, a.foto as aFoto, * FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere & " ORDER BY " & sOrder
	rst.Open sQuery, sProvider
%>
<div class=container>
	<div class=logo><ul><li><a title='Ir a Nuevo Anuncio' href=TrVacasRegOfrezcoFront.asp>Publicar Apartamento u Hotel</a></li></ul></div>
	<div class=topmenu><!--#include file="IncNuSubMenu.asp"--></div>
	<div class=main>
<%	
	if numPisos=0 then
		if Request("tipobusqueda") = "nuevas" then
%>
			Nada nuevo desde tu última visita. Puedes ver todas si pulsas <a title='Hacer otra búsqueda sin filtros' href='<%=Request.Servervariables("SCRIPT_NAME") %>?tipobusqueda=todas'>aquí</a>.
<%		else %>
			No hemos encontrado viviendas con esos criterios.
<%		
		end if
		Response.End
	end if

	PagCabeza numPisos, "Vivienda"
%>
    	</div>
	<div>	
  		<div title='Provincia'>Dónde</div><div title='¿Está en español?'>Idioma</div><div title='Precio por semana'><img src=images/Eurito.gif class="euro"/></div>
  		<div title='¿Qué es?'>Tipo</div><div title='¿Hay fotos?'>Foto</div><div title='¿Se permiten Fumadores?'>Fumadores</div><div title='¿Se permiten Mascotas?'>Mascotas</div></div>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
%>
	<div onmouseover="this.bgColor='cyan';" 
<%	        if rst("destacado") then %>
	    class="destacado" onmouseout="if (this.rowIndex % 2) {this.bgColor='blue'} else {this.bgColor='#FFFFFF'};">
<%			else %>
	    <% if (i mod 2)=1 then Response.Write "bgcolor=#CC99FF"%> onmouseout="if (this.rowIndex % 2) {this.bgColor='#CC99FF'} else {this.bgColor='#FFFFFF'};">
<%			end if %>    
   		<div title='Haz click para ver el anuncio' <% if rst("destacado") then Response.Write "class=destacado" %>>
    		<a href="javascript:detalle('TrAnuncioDetalle.asp?tabla=Pisos&id=<%=rst("aId")%>')"><%=rst("cabecera")%></a></div>
   		<div title='¿Está en español?' <% if rst("destacado") then Response.Write "class=destacado" %>>
<%			if UCase(rst("idiomaes"))="ON" then %>
			<img src=images/Espanol.gif alt='Anuncio en Español'/>
<%			
            end if 
            if UCase(rst("idiomaen"))="ON" then 
%>
			<img src=images/Ingles.gif alt='Anuncio en Inglés'/>
<%			end if %>
		</div>
        <div title='Precio por semana' <% if rst("destacado") then Response.Write "class=destacado" %>>	<%=rst("rentavacas")%> €/semana</div>
		<div title='¿Qué es?' <% if rst("destacado") then Response.Write "class=destacado" %>><%=rst("pTipo")%></div>
		<div title='Haz click para ver las fotos' <% if rst("destacado") then Response.Write "class=destacado" %>>
<%			if rst("aFoto") = "Si" then %>
			<a href="javascript:foto('<%=rst("aId")%>')"><img src=images/Camara.gif alt='Pulsa para ver las fotos'/></a>
<%			elseif rst("aFoto") <> "" then %>		
			<a href="javascript:foto('<%=rst("aId")%>')"><img src='http://domoh.com/mini<%=rst("aFoto")%>' alt='Pulsa para ver las fotos'/></a>
<%			end if %>    
			</div>
		<div title='¿Se permiten Fumadores?' <% if rst("destacado") then Response.Write "class=destacado" %>>
<%			if UCase(rst("Fuma")) = "ON" then %>
		    <img src=images/Cigarrito.gif alt='Admiten Fumadores'/>
<%			else %>
   			<img src=images/NoCigarrito.gif alt='No Admiten Fumadores'/>
<%			end if %>
			</div>
		<div title='¿Se permiten Mascotas?' <% if rst("destacado") then Response.Write "class=destacado" %>>
<%			if UCase(rst("Mascota")) = "ON" then %>
            <img src=images/Perrillos.gif alt='Admiten Mascotas'/>
<%			else %>
            <img src=images/NoPerrillos.gif alt='No Admiten Mascotas'/>
<%			end if %>
	        </div></div>
<%
		end if
		rst.Movenext
	next

	PagPie numPisos, Request.Servervariables("SCRIPT_NAME") & "?" & sReq 
%>
    </div>
<!-- #include file="IncPie.asp" -->
</body>