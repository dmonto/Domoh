<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<%
	dim sWhere, sOrder, i, j, sNombreProvincia, numProvincia, sHora, sReq, numPais, numBueno, numRegular, numMalo, numVotos, numPisos, sBack, sIdioma, sCampoPrecio, sTabla, sCampoTipo, sCampoFoto, sDesc, sClass, sInvDesc
				 	
	if Request("tipo")="inquilino" or Request("tipo")="comprador" then
		sTabla="Inquilinos"
		sCampoPrecio="maximo"
		sCampoTipo="tipoviv"
		sCampoFoto="u.foto"
	else
		sTabla="Pisos"
		sCampoTipo="tipo"
		sCampoFoto="a.foto"
		if Request("tipo")="alquiler" or Request("tipo")="" then
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
		sOrder= "p.tipo " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, nummalo "
	elseif Request("orden")="ciudad" then
		sOrder= "ciudadnombre " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, nummalo "
	elseif Request("orden")="foto" then
		sOrder= sCampoFoto & " " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, nummalo "
	elseif Request("orden")="fuma" then
		sOrder= "p.fuma " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, nummalo "
	elseif Request("orden")="mascota" then
		sOrder= "p.mascota " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, nummalo "
	elseif Request("orden")="precio" then
		sOrder= "p." & sCampoPrecio & " " & Request("desc") & ", a.idiomaes DESC, a.fechaultimamodificacion DESC"
	elseif Request("orden")="idioma" then
		sOrder= "a.idiomaes " & Request("desc") & ", a.fechaultimamodificacion DESC, p." & sCampoPrecio 
	elseif Request("orden")="valor" then
		sOrder= "2 " & Request("desc") & ", 1 " & Request("desc") & ", numbueno DESC, numregular DESC, nummalo "
	else
		sOrder= "a.destacado DESC, a.idiomaes DESC, p." & sCampoPrecio 
	end if
	
	if Request("desc")="" then 
		sDesc="desc" 
	else 
		sDesc=""
	end if
	
	sQuery="SELECT FLOOR(SUM(av.visitas)/(DATEDIFF(DAY,min(av.fecha),GETDATE())+1)*7/50) AS numVistoSemana, FLOOR(u.numanuncios/5) as numAnuncios, a.id AS aId, p.mascota AS pMascota, p.fuma AS pFuma, "
	sQuery=sQuery & " CONVERT(VARCHAR, a.fechaultimamodificacion, 3) AS aFechaUltimaModificacion, p." & sCampoPrecio & " AS pPrecio, p." & sCampoTipo & " AS pTipo, " & sCampoFoto & " AS aFoto, a.destacado, a.idiomaes, "
    sQuery=sQuery & " a.idiomaen, esagencia, cabecera, "
	if sTabla="Pisos" then sQuery=sQuery & " zona, ciudadnombre, "
	sQuery=sQuery & " a.numbueno, a.nummalo, a.numregular "
	sQuery=sQuery & " FROM ((" & sTabla & " p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.Usuario=u.Usuario) "
	sQuery=sQuery & " LEFT JOIN AnunciosVistos av ON a.id=av.anuncio " & sWhere & " GROUP BY a.id, p.mascota, p.fuma, a.fechaultimamodificacion, "
	sQuery=sQuery & "p." & sCampoPrecio & ", p." & sCampoTipo & ", " & sCampoFoto & ", a.destacado, a.idiomaes, a.idiomaen, esagencia, cabecera, "
	if sTabla="Pisos" then sQuery=sQuery & " zona, ciudadnombre, "
	sQuery=sQuery & " a.numbueno, a.nummalo, a.numregular, numanuncios ORDER BY " & sOrder
	rst.Open sQuery, sProvider
%>
	    <h3><% PagCabeza numPisos, MesgS("Anuncio","Advert") %></h3>
        <div id=accordion>
<% 	for i=1 to numPisos %>
            <h3><%=rst("cabecera")%></h3>
            <div>
                <div class=left>
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
					<div class=left>
		    	    <a href=# onclick="window.open('QuAnuncioDetalle.asp?tabla=<%=sTabla%>&id=<%=rst("aId")%>','anunciospop','scrollbars=yes,width=700,height=550')"><img alt='Detalle' src=images/TrDetalle.gif /></a></div>
				<div class=middle>
<%	    
        Response.Write rst("aFechaUltimaModificacion") 
	    if UCase(rst("IdiomaEs"))="ON" then 
%>
					<img src=images/Espanol.gif alt='<%=MesgS("Anuncio en Español","Advert in Spanish")%>'/>
<%	        
            end if 
	        if UCase(rst("IdiomaEn"))="ON" then 
%>
					<img src=images/Ingles.gif alt='<%=MesgS("Anuncio en Inglés","Advert in English")%>'/>
<%	        end if %>
					</div>
			    <div class=middle>
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
				<div class=middle>
					<a href='QuAnuncioDetalle.asp?tabla=<%=sTabla%>&id=<%=rst("aId")%>'>
<%	        if sTabla="Pisos" then %>
						<%=rst("ciudadnombre")%> (<%=rst("zona")%>)
<%	        else %>
						<%=rst("cabecera")%>
<%	        end if %>
                        </a></div>
<%	        if Request("tipo")<>"vacasswap" then %>
				<div class=middle>
                    <a href="QuAnuncioDetalle.asp?tabla=<%=sTabla%>&id=<%=rst("aId")%>"><%=rst("pPrecio")%>
<%		        if Request("tipo")="alquiler" or Request("tipo")="inquilino" then %>
			    		&euro;/<%=MesgS("mes","month")%>
<%		        elseif Request("tipo")="vacas" then %>
                        &euro;/<%=MesgS("semana","week")%>
<%		        elseif request("tipo")="venta" then %>
                        <%=MesgS("mil ","thousand ")%>&euro;
<%		        end if %>
                        </a></div>
<%	        end if %>
                <div class=middle>
<%
	        for j=1 to rst("numAnuncios") 
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
                <div class=middle>
                    <a href="javascript:{sw1=window.open('TrMapa.asp?id=<%=rst("aId")%>','searchpage','toolbar=no,location=no,directories=no,status=no,scrollbars=yes,resizable=no,copyhistory=no,width=650,height=450'); 
                        sw1.focus();}">
                    <img src=images/Mapa.gif alt='Pulsa para ver el mapa'/></a></div>
<%	        
            end if  
	        if Request("tipo")<>"venta" then 
%>
                <div class='middle'>
<%              if UCase(rst("pFuma")) = "ON" then %>
                    <img src=images/Cigarrito.gif alt='Admiten Fumadores'/>
<%		        else %>
   		            <img src=images/NoCigarrito.gif alt='No Admiten Fumadores'/>
<%		        end if %>
                    </div>
                <div class='middle'>
<%		        if UCase(rst("pMascota")) = "ON" then %>
                    <img src=images/Perrillos.gif alt='Admiten Mascotas'/>
<%		        else %>
                    <img src=images/NoPerrillos.gif alt='No Admiten Mascotas'/>
<%		        end if %>
                    </div>
<%	        end if %>
                </div>
<% 	
    	rst.Movenext
    next 
%>
			</div><!-- accordion-->
		