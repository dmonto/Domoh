<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';"> 
<%
	dim sWhere, sOrder, numInquilinos, i, sHora, sIdioma
 	
	sWhere= "WHERE a.activo='Si' AND i.tipoviv='Compra' AND a.provincia <> 0 AND u.activo='Si' "
	sQuery="SELECT COUNT(*) AS numInquilinos FROM (Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) INNER JOIN Usuarios u ON u.usuario=a.usuario " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
	numInquilinos=rst("numInquilinos")
	rst.Close

	if Session("Idioma")="En" then sIdioma="a.idiomaen" else sIdioma="a.idiomaes" 
	
	if Request("orden")="fecha" then
		sOrder= "a.fechaultimamodificacion DESC, " & sIdioma & " DESC "
	elseif Request("orden")="precio" then
		sOrder= "i.maximo DESC, a.idiomaes DESC"
	elseif Request("orden")="idioma" then
		sOrder= "a.idiomaes DESC, a.fechaultimamodificacion DESC"
	else
		sOrder= "p.nombre DESC, a.idiomaes DESC"
	end if
	
	sQuery="SELECT p.nombre AS pNombre, a.id AS aId, * FROM "
	sQuery=sQuery & " ((Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) INNER JOIN Usuarios u ON u.usuario=a.usuario) INNER JOIN Provincias p ON a.provincia=p.id " & sWhere & " ORDER BY " & sOrder
	rst.Open sQuery, sProvider
%>
<div class=container>
	<div class=row><!--#include file="IncNuSubMenu.asp"--></div>
	<div><ul><li><a title='<%=MesgS("Poner un anuncio","Post a New Ad for selling your property")%>' href=NuCasaCompraOfrezcoFront.asp>Publicar Piso o Casa</a></li></ul></div>
	<div>
		Si deseas asesoramiento para la tramitación de la venta de tu piso envíanos un correo a <a title='Enviar mail a Domoh' href=mailto:atencioncliente@domoh.com> atencioncliente@domoh.com</a>
<%	if numInquilinos=0 then %>
		Pulsa en "Publicar Piso o Casa" para registrar tu anuncio.
        <!-- #include file="IncPie.asp" -->
<%		
		Response.End
	end if

	PagCabeza numInquilinos, "Personas que buscan Piso o Casa"
%>
	</div>
	<div>	
		<div><a title='Ver los mas recientes primero' href='<%=Request.Servervariables("SCRIPT_NAME")%>?orden=fecha'>por Fecha</a></div>
    	<div><a title='Ver los que estén en español primero' href='<%=Request.Servervariables("SCRIPT_NAME")%>?orden=idioma'>por Idioma</a></div>
    	<div><a title='Ver los que ofrecen más dinero primero' href='<%=Request.Servervariables("SCRIPT_NAME")%>?orden=precio'>por Precio</a></div></div>
	<div>	
		<div title='Última Actualización'>Publicado</div><div title='¿Está en Español?'>Idioma</div><div title='Ciudad/Zona'>Dónde</div><div title='Pulsa para ver el anuncio'>Anuncio</div>
    	<div title='Precio en miles de Euros'><img alt='Precio' src=images/Eurito.gif class="euro" /></div><div title='Casa/Piso/Otros'>Tipo</div></div>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
%>
	<div <%if (i mod 2)=1 then Response.write "bgcolor=#CC99FF"%> onmouseover="this.bgColor='cyan';" 
  		onmouseout="if (this.rowIndex % 2) {this.bgColor='#CC99FF'} else {this.bgColor='#FFFFFF'};" onclick="javascript:detalle('TrAnuncioDetalle.asp?tabla=Inquilinos&id=<%=rst("aId")%>')">
   		<div title='Última Actualización'>
<%	
			sHora=Hour(rst("fechaultimamodificacion")) & ":" & Minute(rst("fechaultimamodificacion"))
			if sHora="0:0" then
				Response.Write FormatDateTime(rst("fechaultimamodificacion"),2)
			elseif rst("fechaultimamodificacion")<>"" then
				Response.Write FormatDateTime(rst("fechaultimamodificacion"),2) & " " & FormatDateTime(rst("fechaultimamodificacion"),4)
			end if
%>
		</div>
   		<div title='¿Está en Español?'>
<%			if UCase(rst("idiomaes"))="ON" then %>
			<img src=images/Espanol.gif alt='Anuncio en Español'/>
<%			
            end if 
			if UCase(rst("idiomaen"))="ON" then 
%>
			<img src=images/Ingles.gif alt='Anuncio en Inglés'/>
<%			end if %>
		</div>
   		<div title='Ciudad/Zona'><%=rst("pNombre")%> (<%=rst("ciudad")%>)</div>
   		<div title='Pulsa para ver el anuncio'><a href="javascript:detalle('TrAnuncioDetalle.asp?tabla=Inquilinos&id=<%=rst("aId")%>')"><%=rst("cabecera")%></a></div>
		<div title='Precio en miles de Euros'><%=rst("maximo")%> mil €</div><div title='Casa/Piso/Otros'><%=rst("tipoviv")%></div></div>
<%
		end if
		rst.Movenext
	next

	PagPie numInquilinos, Request.Servervariables("SCRIPT_NAME") & "?orden=" & Request("orden") & "&"
%>
</div>
<!-- #include file="IncPie.asp" -->
</body>
