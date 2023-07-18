<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';"> 
<%
	dim sWhere, sOrder, i, j, sNombreProvincia, numProvincia, sHora, sReq, numPais, numBueno, numRegular, numMalo, numVotos, numPisos, sBack, sIdioma, sCampoPrecio, sTabla, sCampoTipo, sCampoFoto, sDesc, sClass, sInvDesc
				 	
	if Session("Donar")<>"Si" and Request("donar")<>"Si" and Request("op") <> "Test" then 
		Response.Redirect "TrDonar.asp?destino=" & Request.ServerVariables("SCRIPT_NAME") & "&" & Request.ServerVariables("QUERY_STRING")
	end if
	
	if Request("tipo")="inquilino" or Request("tipo")="comprador" then
		sTabla="Inquilinos"
		sCampoPrecio="maximo"
		sCampoTipo="tipoviv"
		sCampoFoto="u.foto"
	else
		sTabla="Pisos"
		sCampoTipo="tipo"
		sCampoFoto="a.foto"
		if Request("tipo")="alquiler" or Request("op")="Test" then
			sCampoPrecio="rentaviv"
		elseif Request("tipo")="venta" then
			sCampoPrecio="precio"
		elseif Request("tipo")="vacas" or Request("tipo")="vacasswap" then
			sCampoPrecio="rentavacas"
		else
			Response.Write "Por favor inténtelo de nuevo."
			Response.End
		end if
	end if
	sReq="tipo=" & Request("tipo") & "&"
			
	numProvincia=0
	if Request.Form("provincia")<>"" or Request.QueryString("provincia")<>"" then
		numProvincia=CLng(Request("provincia"))
		Response.Cookies("ProvinciaViv")=numProvincia
		Response.Cookies("ProvinciaViv").Expires=Now+30
		sReq=sReq & "provincia=" & Request("provincia") & "&"
	end if

	if Request("pais")<>"" then
		numPais=CLng(Request("pais"))
		sReq=sReq & "pais=" & Request("pais") & "&"
	end if

	sWhere= "WHERE a.activo='Si' AND u.activo='Si' "
	if Request("tipo")="vacasswap" then
		sWhere= sWhere & "AND rentavacas=0 AND rentaviv=0 AND precio=0 " 
	elseif Request("tipo")="inquilino" then
		sWhere= sWhere & "AND p.tipoviv <> 'Compra' " 
	elseif Request("tipo")="comprador" then		
		sWhere= sWhere & "AND p.tipoviv = 'Compra' " 
	else
		sWhere= sWhere & "AND p." & sCampoPrecio & " <> 0 " 
	end if
		
	if numProvincia<>0 then 
		sWhere= sWhere & " AND a.provincia = " & numProvincia
	elseif numPais<>0 then 
		sWhere= sWhere & " AND a.provincia IN (SELECT id FROM Provincias WHERE pais= " & numPais & ")"
	else
		sWhere= sWhere & " AND a.provincia <> 0 "
	end if
		
	if Request("precio") <> "" then 
		sWhere= sWhere & " AND p." & sCampoPrecio & " <= " & Request("precio")
		sReq=sReq & "&precio=" & Request("precio")
	end if
	
	sQuery="SELECT COUNT(*) AS numPisos FROM (" & sTabla & " p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.Usuario " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en " & Request.ServerVariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
	numPisos=rst("numPisos")
	rst.Close

	if Request("orden")="fecha" then
		if Request("desc")="" then sInvDesc="" else sInvDesc="DESC"
		sOrder= "a.fechaultimamodificacion " & sInvDesc & ", a.idiomaes DESC, p." & sCampoPrecio 
	elseif Request("orden")="tipo" then
		sOrder= "p.tipo " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="ciudad" then
		sOrder= "ciudadnombre " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="foto" then
		sOrder= sCampoFoto & " " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="fuma" then
		sOrder= "p.fuma " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="mascota" then
		sOrder= "p.mascota " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="precio" then
		sOrder= "p." & sCampoPrecio & " " & Request("desc") & ", a.idiomaes DESC, a.fechaultimamodificacion DESC"
	elseif Request("orden")="idioma" then
		sOrder= "a.idiomaes " & Request("desc") & ", a.fechaultimamodificacion DESC, p." & sCampoPrecio 
	elseif Request("orden")="valor" then
		sOrder= "2 " & Request("desc") & ", 1 " & Request("desc") & ", numbueno DESC, numregular DESC, NumMalo "
	else
		sOrder= "a.destacado DESC, a.idiomaes DESC, p." & sCampoPrecio 
	end if
	
	if Request("desc")="" then 
		sDesc="desc" 
	else 
		sDesc=""
	end if
	
	sQuery="SELECT FLOOR(SUM(av.visitas)/(DATEDIFF(DAY,min(av.fecha),GETDATE())+1)*7/50) AS numVistoSemana, FLOOR(u.numanuncios/5) as NumAnuncios, a.id AS aId, p.mascota AS pMascota, p.fuma AS pFuma, "
	sQuery=sQuery & " CONVERT(VARCHAR, a.fechaultimamodificacion, 3) AS aFechaUltimaModificacion, p." & sCampoPrecio & " AS pPrecio, p." & sCampoTipo & " AS pTipo, " & sCampoFoto & " AS aFoto, a.destacado, a.idiomaes, "
    sQuery=sQuery & " a.idiomaen, esagencia, cabecera, "
	if sTabla="Pisos" then sQuery=sQuery & " zona, ciudadnombre, "
	sQuery=sQuery & " a.numbueno, a.nummalo, a.numregular FROM ((" & sTabla & " p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.Usuario=u.Usuario) "
	sQuery=sQuery & " LEFT JOIN AnunciosVistos av ON a.id=av.anuncio " & sWhere & " GROUP BY a.id, p.mascota, p.fuma, a.fechaultimamodificacion, "
	sQuery=sQuery & "p." & sCampoPrecio & ", p." & sCampoTipo & ", " & sCampoFoto & ", a.destacado, a.idiomaes, a.idiomaen, esagencia, cabecera, "
	if sTabla="Pisos" then sQuery=sQuery & " zona, ciudadnombre, "
	sQuery=sQuery & " a.numbueno, a.nummalo, a.numregular, numanuncios ORDER BY " & sOrder
	rst.Open sQuery, sProvider
%>
<div class=container>
<%  if Request("msg")<>"" then %>
	<h2 class=banner><%=Request("msg")%></h2>
<% end if %>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class=topmenu>
<%  if Request("tipo")="alquiler" or Request("tipo")="venta" or Request("tipo")="inquilino" or Request("tipo")="comprador" then %>
			<a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a>
<%	else %>
			<a href=TrVacas.asp class=linkutils><%=MesgS("Anuncios Vacaciones","Holiday Adverts")%> &gt; </a>
<%	
    end if 
	if Request("tipo")="alquiler" then 
%>
			<a href=TrViviendaBuscoFront.asp class=linkutils><%=MesgS("Búsqueda","Search")%> &gt; </a>
<%	elseif Request("tipo")="vacas" then %>
			<a href=TrVacasBuscoFront.asp class=linkutils><%=MesgS("Búsqueda","Search")%> &gt; </a>
<%	
    end if  
	if Request("tipo")="alquiler" then 
%>
            <%=MesgS("Pisos y Habitaciones en ALQUILER","Flats and Rooms for RENT")%> 
<%	elseif Request("tipo")="venta" then %>
			<%=MesgS("Pisos y Casas en VENTA","Flats and Houses for SALE")%> 
<%	elseif Request("tipo")="vacas" then %>
			<%=MesgS("Apartamentos y Hoteles","Flats and Hotels for VACATIONS")%>
<%	elseif Request("tipo")="vacasswap" then %>
			<% Response.Write MesgS("Apartamentos y Hoteles para INTERCAMBIO", "Flats and Hotels for SWAPPING")%>
<%	elseif Request("tipo")="inquilino" then %>
			<%=MesgS("Ofertas de INQUILINOS","People looking for Flats and Houses for RENT")%>
<%	end if  %>
			</div></div>
<%	if Request("tipo")="comprador" then %>
    <div class=main>
        <p>
            <% Response.Write MesgS("Si deseas asesoramiento para la tramitación de la venta de tu piso envíanos un correo a", "If you need any assistance with the sale of your property send us an e-mail to")%> 
			<a href='mailto:atencioncliente@domoh.com'>atencioncliente@domoh.com</a></p>
<%		if numPisos=0 then %>
        <% Response.Write MesgS("Pulsa en 'Publicar Piso o Casa' para registrar tu anuncio", "Go to 'Post New House for Sale' for selling your property")%>.
<%		
			Response.End
		end if
	end if 
%>
        </div>
	<div>
	    <div><div><% PagCabeza numPisos, MesgS("Anuncio","Advert") %></div><div><% PagPie numPisos, "TrViviendaBusco.asp?" & sReq & "orden=" & request("orden") & "&" %></div></div>
        <div><% Response.Write MesgS("Para ver los detalles de un anuncio, haz clic en su icono <img src=images/TrDetalle.gif />", "Click <img src=images/TrDetalle.gif /> to see advert details")%></div></div>
	<div>
	    <div>
            <div title='<%=MesgS("Última Modificación","Last Update")%>' class=anuncios2_cab><a href='?orden=fecha&desc=<%=sDesc%>&<%=sReq%>&'><%=MesgS("Publicado","Posted")%></a></div>
			<div class=anuncios2_cab><a href='?orden=tipo&desc=<%=sDesc%>&<%=sReq%>&'><%=MesgS("Tipo","Type")%></a></div>
			<div class=anuncios2_cab><a href='?orden=ciudad&desc=<%=sDesc%>&<%=sReq%>&'><%=MesgS("Dónde","Where")%></a></div>
			<div class=anuncios2_cab><a href='?orden=foto&desc=<%=sDesc%>&<%=sReq%>&'><%=MesgS("Foto","Picture")%></a></div>
<%	if Request("tipo")<>"vacasswap" then %>
			<div class=anuncios2_cab><a href='?orden=precio&desc=<%=sDesc%>&<%=sReq%>&'><%=MesgS("Precio","Price")%></a></div>
<%	end if  %>
			<div class=anuncios2_cab><a href='?orden=valor&desc=<%=sDesc%>&<%=sReq%>&'><%=MesgS("Valoración","Ranking")%></a></div>
<%	if sTabla="Pisos" then %>
			<div class=anuncios2_cab><%=MesgS("MAPA","MAP")%></div>
<%	
    end if 
	if Request("tipo")<>"venta" then 
%>
			<div class=anuncios2_cab><a href='?orden=fuma&desc=<%=sDesc%>&<%=sReq%>&'><%=MesgS("Fumadores","Smokers")%></a></div>
			<div class=anuncios2_cab><a href='?orden=mascota&desc=<%=sDesc%>&<%=sReq%>&'><%=MesgS("Mascotas","Pets")%></a></div>
<%	end if  %>
			</div>
<%		

	for i=1 to RegHasta
		if i >= RegDesde then
			if rst("destacado") then sClass="anunciosDest" else sClass="anuncios2"
%>
        <div>
            <div class=<%=sClass%>>
                <a href=# onclick="window.open('QuAnuncioDetalle.asp?tabla=<%=sTabla%>&id=<%=rst("aId")%>','anunciospop','scrollbars=yes,width=700,height=550')"> <img alt='Detalle' src=images/TrDetalle.gif /></a></div>
			<div class=<%=sClass%>>
                <%=rst("aFechaUltimaModificacion") %>
<%	        if UCase(rst("IdiomaEs"))="ON" then %>
			    <img src=images/Espanol.gif alt='<%=MesgS("Anuncio en Español","Advert in Spanish")%>'/>
<%	        
            end if 
            if UCase(rst("IdiomaEn"))="ON" then 
%>
				<img src=images/Ingles.gif alt='<%=MesgS("Anuncio en Inglés","Advert in English")%>'/>
<%	        end if %>
				</div>
			<div class='<%=sClass%>'>
<%
	        if InStr(rst("pTipo"),"Habi") then
		        Response.Write MesgS("Habitación","Room")
	        elseif rst("pTipo") ="Piso" then
		        Response.Write MesgS("Piso","Flat")
	        else
		        Response.Write rst("pTipo") & " "
	        end if
%>
			    </div>
		    <div class='<%=sClass%>'>
    		    <a href='QuAnuncioDetalle.asp?tabla=<%=sTabla%>&id=<%=rst("aId")%>'>
<%	        if sTabla="Pisos" then %>
	    		    <%=rst("ciudadnombre")%> <%=rst("zona")%>
<%	        else %>
		    		<%=rst("cabecera")%>
<%	        end if %>
                    </a></div>
			<div class='<%=sClass%>'>
<%	        if rst("aFoto") <> "" then %>
                <a
<%		        if sCampoFoto="u.foto" then %>
				    href="javascript:detalle('http://domoh.com/<%=rst("aFoto")%>')">
<%		        else %>    
					href="javascript:foto('<%=rst("aId")%>')">
<%	            end if %>    
					<img src='http://domoh.com/mini<%=rst("aFoto")%>' alt='Pulsa para ver las fotos'/></a>
<%	        else %>    
				&nbsp;
<%	        end if %>    
				</div>
<%	        if Request("tipo")<>"vacasswap" then %>
			<div class=<%=sClass%>>
                <a href='QuAnuncioDetalle.asp?tabla=<%=sTabla%>&id=<%=rst("aId")%>'><%=rst("pPrecio")%>
<%		        if Request("tipo")="alquiler" or Request("tipo")="inquilino" then %>
                    &euro;/<%=MesgS("mes","month")%>
<%		        elseif Request("tipo")="vacas" then %>
                    &euro;/<%=MesgS("semana","week")%>
<%		        elseif request("tipo")="venta" then %>
                    <%=MesgS("mil ","thousand ")%>&euro;
<%		        end if %>
                    </a></div>
<%	        end if %>
            <div class=<%=sClass%>>
<%
	        for j=1 to rst("numanuncios") 
		        Response.Write "<img title='Usuario VIP' src=images/VIP.gif />"
		        if j=3 then exit for
	        next
	        if sTabla="Pisos" and not(IsNull(rst("numVistoSemana"))) then
	            for j=1 to rst("numVistoSemana") 
		            Response.Write "<img title='Anuncio muy Solicitado' src='images/BigEyes.gif'/>"
	            next
	        end if
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
		        Response.Write "<img title='Anuncio votado como Mola!' src='images/Mola.gif'/>"
	        next
	        for j=1 to numRegular
		        Response.Write "<img title='Anuncio votado como OK' src='images/OK2.gif'/>"
	        next
	        for j=1 to numMalo
		        Response.Write "<img title='Anuncio votado como Malo' src='images/Malo.jpg'/>"
	        next
%>
                &nbsp;</div>
<%	        if Request("tipo")<>"inquilino" then %>
            <div class=<%=sClass%>>
                <a href="javascript:{sw1=window.open('TrMapa.asp?id=<%=rst("aId")%>','searchpage','toolbar=no,location=no,directories=no,status=no,scrollbars=yes,resizable=no,copyhistory=no,width=650,height=450'); sw1.focus();}">
                    <img src=images/Mapa.gif alt='Pulsa para ver el mapa'/></a></div>
<%	        
            end if  
	        if Request("tipo")<>"venta" then 
%>
            <div class=<%=sClass%>>
<%              if UCase(rst("pFuma")) = "ON" then %>
                <img src=images/Cigarrito.gif alt='Admiten Fumadores'/>
<%		        else %>
   		        <img src=images/NoCigarrito.gif alt='No Admiten Fumadores'/>
<%		        end if %>
                </div>
            <div class=<%=sClass%>>
<%	Mascota rst("pMascota") %>
                </div>
<%	        end if %>
            </div>
<%
            if sBack="" then sBack="bgcolor=#ededed" else sBack=""
	    end if
		rst.Movenext
	next
%>                            
        </div></div>
<div><!-- #include file="IncPie.asp" --></div></body>
