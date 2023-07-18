<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';"> 
<%
	dim sWhere, sOrder, numPisos, i, j, sNombreProvincia, numProvincia, sHora, sReq, numBueno, numRegular, numMalo, numVotos
 	
	numProvincia=0
	if Request.Form("provincia")<>"" or Request.QueryString("provincia")<>"" then
		numProvincia=CLng(Request("provincia"))
		Response.Cookies("ProvinciaViv")=numProvincia
		Response.Cookies("ProvinciaViv").Expires=Now+30
		sReq="&provincia=" & Request("provincia")
	end if

	sWhere= "WHERE a.activo='Si' AND u.activo='Si' AND p.rentaviv <> 0 " 
	if numProvincia<>0 then sWhere= sWhere & " AND a.provincia = " & numProvincia else sWhere= sWhere & " AND a.provincia <> 0 "
	
	if Request("precio") <> "" then 
		sWhere= sWhere & " AND p.rentaviv <= " & Request("precio")
		sReq=sReq & "&precio=" & Request("precio")
	end if
	
	sQuery="SELECT COUNT(*) AS numPisos FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en TrCasaBusco", sQuery & " - " & Err.Description
	numPisos=rst("numPisos")
	rst.Close

	if Request("orden")="fecha" then
		sOrder= "a.destacado DESC, a.fechaultimamodificacion DESC, a.idiomaes DESC, p.rentaviv"
	elseif Request("orden")="precio" then
		sOrder= "a.destacado DESC, p.rentaviv, a.idiomaes DESC, a.fechaultimamodificacion DESC"
	elseif Request("orden")="idioma" then
		sOrder= "a.destacado DESC, a.idiomaes DESC, a.fechaultimamodificacion DESC, p.rentaviv"
	elseif Request("orden")="valor" then
		sOrder= "a.destacado DESC, 2 DESC, 1 DESC, numbueno DESC, numregular DESC, nummalo "
	else
		sOrder= "a.destacado DESC, a.idiomaes DESC, p.rentaviv"
	end if
	
	sQuery="SELECT FLOOR(SUM(av.visitas)/(DATEDIFF(DAY,MIN(av.fecha),GETDATE())+1)*7/50) AS vistoSemana, FLOOR(u.numanuncios/5) AS numAnuncios, a.id AS aId, p.mascota AS pMascota, p.fuma AS pFuma, "
    sQuery=sQuery & " a.fechaultimamodificacion AS aFechaUltimaModificacion, p.rentaviv AS pRentaViv, p.tipo AS pTipo, a.foto AS aFoto, a.destacado, "
	sQuery=sQuery & " a.idiomaes, a.idiomaen, esagencia, cabecera, zona, a.numbueno, a.nummalo, a.numregular "
	sQuery=sQuery & " FROM ((Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON p.usuario=u.usuario) LEFT JOIN AnunciosVistos av ON p.id=av.anuncio " 
	sQuery=sQuery & sWhere & " GROUP BY a.id, p.mascota, p.fuma, a.fechaultimamodificacion, p.rentaviv, p.tipo, a.foto, a.destacado, a.idiomaes, "
	sQuery=sQuery & " a.idiomaen, esagencia, cabecera, zona, a.numbueno, a.nummalo, a.numregular, numanuncios ORDER BY " & sOrder
	rst.Open sQuery, sProvider
%>
<div class=container>
    <div class=banner>
    <% if Request("msg")<>"" then %>
	    <p class=destacado><%=Request("msg")%></p>
    <% end if %>
        </div>
    <div class=topmenu><!--#include file="IncNuSubMenu.asp"--></div>
	<div class=main>
		<ul>
			<li><a title='Pon tu anuncio de b˙squeda de piso' href=TrCasaRegDemandaFront.asp>No Encuentro lo que Busco</a></li>
			<li><a title='Publica tu piso disponible' href=TrCasaRegOfrezcoFront.asp>Publicar Piso o Habitaciùn</a></li></ul>
<%	
	if numPisos=0 then
		if Request("tipobusqueda") = "nuevas" then
%>
        Nada nuevo desde tu ˙ltima visita. Puedes ver todas si pulsas <a title='Mostrar todas' href='TrCasaBuscoMapa.asp?tipobusqueda=todas'>aquùù</a>.
<%		else %>
		No hemos encontrado viviendas con esos criterios.
<%		
		end if
		Response.End
	end if

	PagCabeza numPisos, "Vivienda"
%>
		<a title='Ver m·s recientes primero' href='TrCasaBusco.asp?orden=fecha<%=sReq%>'>por Fecha</a><a title='Ver anuncios en EspaÒol primero' href='TrCasaBusco.asp?orden=idioma<%=sReq%>'>por Idioma</a>
    	<a title='Ver anuncios con DomohRank alto' href='TrCasaBusco.asp?orden=valor<%=sReq%>'>por ValoraciÛn</a></div><div><a title='Ver los m·s baratos primero' href='TrCasaBusco.asp?orden=precio<%=sReq%>'>por Precio</a>
		<div title='⁄ltima ActualizaciÛn'>Publicado</div><div title='øEst· en EspaÒol?'>Idioma</div><div title='Provincia/Zona'>DÛnde</div>
   		<div title='ValoraciÛn seg˙n domoh y resto de usuarios'>ValoraciÛn</div><div title='Precio en Euros'><img alt='Precio' src=images/Eurito.gif class=euro /></div>
   		<div title='øQuÈ es?'>Tipo</div><div>Foto</div><div>Mapa</div><div>Fumadores</div><div>Mascotas</div></div>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
%>
		<div onmouseout="if (this.rowIndex % 2) {this.bgColor='blue'} else {this.bgColor='#FFFFFF'};" onmouseover="this.bgColor='cyan';"
<%          if rst("destacado") then %>
    class="destacado"
<%			else %>
    <% if (i mod 2)=1 then Response.Write "bgcolor=#CC99FF"%> 
<%			end if %>    
	>
    	<div <% if rst("destacado") then Response.Write "class=destacado" %>><% Fecha(rst("aFechaUltimaModificacion")) %></div>
   		<div <% if rst("destacado") then Response.Write "class=destacado" %>>
<%			if UCase(rst("idiomaes"))="ON" then %>
			<img src=images/Espanol.gif alt='Anuncio en EspaÒol'/>
<%			
            end if 
            if UCase(rst("IdiomaEn"))="ON" then 
%>
			<img src=images/Ingles.gif alt='Anuncio en InglÈs'/>
<%			end if %>
		</div>
   		<div class=
<%	        if rst("destacado") then %>
			"destacado"
<%	        elseif rst("esagencia")="" then %>
			'font: bold;'	
<%	        end if %>
		    >
    		<a href="javascript:detalle('QuCasaDetalle.asp?id=<%=rst("aId")%>')"><%=rst("cabecera")%></a></div>
	    <div <% if rst("destacado") then Response.Write "class=destacado" %>>
<%
	        for j=1 to rst("numanuncios") 
		        Response.Write "<img title='Usuario VIP' src=images/VIP.gif />"
	        next
	        for j=1 to rst("vistoSemana") 
		        Response.Write "<img title='Anuncio muy Solicitado' src=images/BigEyes.gif />"
	        next
	        numBueno=rst("numbueno") 
	        numRegular=rst("numregular") 
	        numMalo=rst("nummalo") 
	        numVotos=numBueno+numRegular+numMalo
	        if numVotos>5 then
		        numBueno=int(numBueno/numVotos*5)
		        numRegular=int(numRegular/numVotos*5)
		        numMalo=int(numMalo/numVotos*5)
	        end if
	        for j=1 to numBueno
		        Response.Write "<img title='Anuncio votado como Mola!' src=images/Mola.gif />"
	        next
	        for j=1 to numRegular
		        Response.Write "<img title='Anuncio votado como OK' src=images/OK2.gif />"
	        next
	        for j=1 to numMalo
		        Response.Write "<img title='Anuncio votado como Malo' src=images/Malo.jpg />"
	        next
%>
            </div>
	    <div <% if rst("destacado") then Response.Write "class=destacado" %>><%=rst("pRentaViv")%> ù/mes</div>
        <div <% if rst("destacado") then Response.Write "class=destacado" %>><%=rst("pTipo")%></div>
        <div <% if rst("destacado") then Response.Write "class=destacado" %>>
<%			if rst("aFoto") = "Si" then %>
    		<a href="javascript:foto('<%=rst("aId")%>')"><img src=images/Camara.gif alt='Pulsa para ver las fotos'/></a>
<%			elseif rst("aFoto") <> "" then %>		
	    	<a href="javascript:foto('<%=rst("aId")%>')"><img src='http://domoh.com/mini<%=rst("aFoto")%>' alt='Pulsa para ver las fotos'/></a>
<%			end if %>    
            </div>
        <div <% if rst("destacado") then Response.Write "class=destacado" %>>
    		<a href="javascript:{sw1=window.open('TrMapa.asp?id=<%=rst("aId")%>','searchpage','toolbar=no,location=no,directories=no,status=no,scrollbars=yes,resizable=no,copyhistory=no,width=655,height=600');sw1.focus();}">
            <img src=images/Mapa.gif alt='Pulsa para ver el mapa'/></a></div>
        <div <% if rst("destacado") then Response.Write "class=destacado" %>>
<%			if UCase(rst("pFuma")) = "ON" then %>
	        <img src=images/Cigarrito.gif alt='Admiten Fumadores'/>
<%			else %>
   		    <img src=images/NoCigarrito.gif alt='No Admiten Fumadores'/>
<%			end if %>
		    </div>
        <div <% if rst("destacado") then Response.Write "class=destacado" %>>
<%			if UCase(rst("pMascota")) = "ON" then %>
            <img src=images/Perrillos.gif alt='Admiten Mascotas'/>
<%			else %>
            <img src=images/NoPerrillos.gif alt='No Admiten Mascotas'/>
<%			end if %>
	        </div>
<%
		end if
		rst.movenext
	next

	PagPie numPisos, "NuCasaBusco.asp?" & sReq & "&orden=" & request("orden")
%>
    </div></div>
<!-- #include file="IncPie.asp" -->
</body>