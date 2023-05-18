<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim sWhere, sOrder, numPisos, i, sIdioma, sReq

	sWhere= "WHERE a.activo='Si' AND u.activo='Si' AND rentavacas=0 AND rentaviv=0 AND precio=0" 
	sQuery="SELECT COUNT(*) AS numPisos FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON u.usuario=a.usuario " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en Busco", sQuery & " - " & Err.Description
	numPisos=rst("numPisos")
	rst.Close

	if Session("Idioma")="En" then sIdioma="idiomaen" else sIdioma="idiomaes" 
	sOrder= sIdioma &" DESC, fechaultimamodificacion DESC"
	sQuery="SELECT a.id as aId, p.Tipo as pTipo, * FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON u.usuario=a.usuario " & sWhere & " ORDER BY " & sOrder
	rst.Open sQuery, sProvider
%>
<head>
    <!-- #include file="IncTrCabecera.asp" -->
    <title>Vacaciones - Intercambio</title>
</head>
<body onload="window.parent.location.hash='top';"> 
<div class=container>
    <div>
<%  if Request("msg")<>"" then %>
		<p><%=Request("msg")%></p
<%  end if %>
        </div>
    <!--#include file="IncNuSubMenu.asp"--></div>
<div class="container"><ul><li><a title='Ir a Nuevo Anuncio' href=TrVacasSwapOfrezcoFront.asp>Publicar Apartamento para Intercambio</a></li></ul></div>
<div class="container">
	<p>
<%	
	if numPisos=0 then
		if Request("tipobusqueda") = "nuevas" then
%>
    	Nada nuevo desde tu última visita. Puedes ver todas si pulsas <a title='Ver todos los anuncios' href='<% Request.Servervariables("SCRIPT_NAME") %>?tipobusqueda=todas'>aquí</a>.
<%		else %>
		No hemos encontrado apartamentos para intercambio.
<%		
        end if 
		Response.End
	end if

	PagCabeza numPisos, "Vivienda"
%>
  	    </p></div>
    <div><a title='Ver primero los más recientes' href='?orden=fecha<%=sReq%>'>por Fecha</a></div>
  	<div>	
    	<div title='Última Modificación'>Publicado</div><div title='Idioma del texto'>Idioma</div><div title='Población/Zona'>Dónde</div><div title='¿Qué es?'>Tipo</div>
    	<div title='Pulsa para ver las Fotos'>Foto</div><div title='¿Se admiten fumadores?'>Fumadores</div><div title='¿Se admiten mascotas?'>Mascotas</div></div>

<%		
	for i=1 to RegHasta
		if i >= RegDesde then
%>
	<div onmouseover="this.bgColor='cyan';" onmouseout="if (this.rowIndex % 2) {this.bgColor='blue'} else {this.bgColor='#FFFFFF'};"
        style=
<%			if rst("destacado") then %>
		'font-weight:bold;font-size:small; color:white' 
<%			elseif (i mod 2)=1 then %>
		'background-color:#CC99FF' 
<%			end if %>    
   	>
   		<div title='Última Modificación' <% if rst("destacado") then Response.Write "" %>>
<%
			Fecha(rst("fechaultimamodificacion"))
%>
		    </div>
   		<div title='Idioma del texto' <% if rst("destacado") then Response.Write "" %>>
<%			if UCase(rst("idiomaes"))="ON" then %>
			<img src=images/Espanol.gif alt='Anuncio en Español'/>
<%			
            end if 
			if UCase(rst("idiomaen"))="ON" then 
%>
			<img src=images/Ingles.gif alt='Anuncio en Inglés'/>
<%			end if %>
		    </div>
   		<div title='Población/Zona' <% if rst("destacado") then Response.Write "" %>>
    		<a <% if rst("destacado") then Response.Write "" %> href="javascript:detalle('TrAnuncioDetalle.asp?tabla=Pisos&id=<%=rst("aId")%>')"><%=rst("cabecera")%></a></div>
        <div title='¿Qué es?' <% if rst("destacado") then Response.Write "" %>><%=rst("pTipo")%></div>
        <div title='Pulsa para ver las Fotos' <% if rst("destacado") then Response.Write "" %>>
<%			if rst("foto") = "Si" then %>
	    	<a href="javascript:foto('<%=rst("aId")%>')"><img src=images/Camara.gif alt='Pulsa para ver las fotos'/></a>
<%			elseif rst("foto") <> "" then %>		
		    <a href="javascript:foto('<%=rst("aId")%>')"><img src='http://domoh.com/mini<%=rst("foto")%>' alt='Pulsa para ver las fotos'/></a>
<%			end if %>    
            </div>
        <div <% if rst("destacado") then Response.Write "" %>>
<%			if UCase(rst("fuma")) = "ON" then %>
	        <img src=images/Cigarrito.gif alt='Admiten Fumadores'/>
<%			else %>
   		    <img src=images/NoCigarrito.gif alt='No Admiten Fumadores'/>
<%			end if %>
		    </div>
        <div <% if rst("destacado") then Response.Write "" %>>
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

	PagPie numPisos, Request.Servervariables("SCRIPT_NAME") & "?" & sReq & "&orden=" & Request("orden") & "&"
%>
<!-- #include file="IncPie.asp" -->
</body>
