<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<%
	dim numPisos, i, sPantalla, numPiso, vfoto

	Session("Id")=""
	if Session("Usuario")="hector" then 
		if Request("msg") like "Piso*" then Response.Redirect Session("Pantalla")
	elseif Session("Usuario")="" or Session("Activo")<>"Si" then 
		Response.Redirect "NuDomoh.asp?msg=" & MesgS("Sesión+finalizada","Session+Ended")
	end if
	
	sQuery= "SELECT foto FROM Usuarios WHERE usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
	vfoto=rst("foto")
	rst.Close

	sQuery= "SELECT COUNT(*) AS numPisos FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id "
	sQuery= sQuery & " WHERE a.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
	numPisos=rst("numPisos")
	rst.Close

	if numPisos then
		sQuery= "SELECT p.dir1 AS pDir1, a.foto AS aFoto, a.activo AS aActivo, "
		sQuery= sQuery & " a.provincia AS aProvincia, a.id AS aId, p.tipo AS pTipo, * "
		sQuery= sQuery & " FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) "
		sQuery= sQuery & " INNER JOIN Usuarios u ON a.usuario=u.usuario "
		sQuery= sQuery & " WHERE a.usuario='"& Session("Usuario") & "' "
		rst.Open sQuery, sProvider

		numPiso=rst("aId")
	end if
%>
<!--#include file="IncTrCabecera.asp"-->
<title>Domoh - Zona de Usuario</title>
<body onLoad="window.parent.location.hash='top';"> 
<table width=100%>
	<tr>
		<td width=500 colspan=3><% if Request("msg")<>"" then %>
			<p><%=Request("msg")%></p>
<% end if %>
</td><!--#include file="IncNuSubMenu.asp"--></tr>
	<tr>
		<td width=400 colspan=5>
			<p align=left><%=Session("Nombre")%>, <%=MesgS("modifica","update")%> 
				<a title="Modificar Ficha de Usuario" href=TrUsuario.asp?origen=QuHomeUsuario.asp>
				aquí</a> tus datos y preferencias.</tr>
<% if vfoto<>"" and not(IsNull(vfoto)) then %>
		<td rowspan=3>
			<img title="Para cambiar esta foto ve a Modificar Datos" width=100 src="http://domoh.com/<%=vfoto%>"></td>
<% end if %>
	<tr>
		<td width=400 colspan=5>			
			<p align=left><b><font color=#FFFFFF size=4>Pisos y Habitaciones</font></b>
	<tr>
		<td>
			<br><br>
<%	if numPisos>0 then %>
	<tr>
		<td title="Indica si tu anuncio es visible por usuarios o no" width=50><b>Estado</b></td>
		<td title="¿Qué es?" width=50 ><b>Tipo</b></td>
		<td title="Ciudad/Zona" width=200 ><b>Dónde</b></td>
		<td title="Mensual/Semanal/Venta" width=50 >
			<b>Precio</b></td>
		<td title="Pulsa para ver las fotos" width=80>
			<b>¿Foto?</b></td>
		<td title="Modificar, Ocultar..." width=150>
			<b>Qué hago con ella</b></td></tr>
<%		
	end if

	for i=1 to numPisos
%>
	<tr>
		<td title="Indica si tu anuncio es visible por usuarios o no" >
<%	if rst("aActivo")="No" then %>
			Oculto
<%	elseif rst("aProvincia")=0 then %>
			En proceso de publicación
<%	else %>
			Publicado
<%	end if %>									
		</td>
		<td title="¿Qué es?" height=50><%=rst("pTipo")%></td>
		<td title="Ciudad/Zona" height=50>
		<a href="javascript:detalle('TrAnuncioDetalle.asp?tabla=Pisos&id=<%=rst("aId")%>&admin=si')">
			<%=rst("cabecera")%></a></td>
		<td title="Mensual/Semanal/Venta" height=50>
<%	if rst("rentaviv")>0 then %>
			<%=rst("rentaviv")%> <font face=arial>€</font>/mes
<%			sPantalla="TrCasaRegOfrezcoFront.asp"
	elseif rst("rentavacas")>0 then %>
			<%=rst("rentavacas")%> <font face=arial>€</font>/semana
<%			sPantalla="TrVacasRegOfrezcoFront.asp"
	else %>
<%		if rst("precio")>0 then %>
			<%=rst("precio")%> mil <font face=arial>€</font>
<%			sPantalla="TrCasaCompraOfrezcoFront.asp"
		else %>
			Intercambio
<%			sPantalla="TrVacasSwapOfrezcoFront.asp"
		end if %>
<%	end if %>
		</td>
		<td title="Pulsa para ver las fotos" height=50>
<%			if rst("aFoto") = "Si" then %>
			<a href=javascript:foto('<%=rst("aId")%>')>
			<img src="../images/Camara.gif" height=50 alt="Pulsa para ver las fotos"></a>
<%			elseif rst("aFoto") <> "" then %>		
			<a href=javascript:foto('<%=rst("aId")%>')>
			<img src="http://domoh.com/mini<%=rst("aFoto")%>" height=50 
				alt="Pulsa para ver las fotos"></a>
<%			end if %>    
		</td>
		<td title="Modificar, Ocultar..." height=50>
			<select onchange="if(value) location='<%=sPantalla%>?op='+value+'&id=<%=rst("aId")%>';" 
					name=<%=rst("aId")%>>
				<option selected value=0>-Elige Una-
				<option value=Modificar>Modificar</option>
<%	if rst("aActivo")="No" then %>
				<option value=Activar>Reactivar</option>
<%	elseif rst("aProvincia")<>0 then %>
				<option value=Borrar>Ocultar</option>
<%	end if %>									
				</select></td></tr>
<%
		rst.Movenext
	next
	if numPisos then rst.Close
%>				
	<tr>
		<td colspan=2>
			<form method=get name=frm action=../TrCasaRegOfrezcoFront.asp?id=nuevo>
			<input title="Nuevo Anuncio" type=submit value="Añadir Piso o Habitación para Vivir" >
			</form></td>
		<td>
			<form method=get name=frm action=../TrVacasRegOfrezcoFront.asp?id=nuevo>
			<input type=submit value="Para Vacaciones">
			</form></td>
		<td>
			<form method=get name=frm action=../TrCasaCompraOfrezcoFront.asp?id=nuevo>
			<input type=submit value="Para Vender">
			</form></td>
		<td>
			<FORM method=get name=frm action=../TrVacasSwapOfrezcoFront.asp?id=nuevo>
			<input type="submit" value="Para Intercambio" 
			></form></td></tr>
<%	if numPiso then %>				
    <tr>
		<td align="center" colspan=3>
           ¿Quieres que te lo <b>traduzcamos al inglés</b>?<br>
           Envía <font size="4">DOMOH ING<%=numPiso%> </font>al <b>25522 </b>
           (Coste 1,20<font face=Arial>€</font> + IVA)<br><br> 
           Y nos encargaremos de todo</TD>
		<td align="center" colspan=3>
           Nueva opción de <b>ANUNCIO DESTACADO</b>
           Envía <font size="4">DOMOH DES<%=numPiso%> </font>al <b>25522 </b>
           (Coste 1,20<font face=Arial>€</font> + IVA)<br><br> 
           Y aparecerá durante <b>DIEZ DÍAS </b>de los primeros y en un formato 
           <a href=../NuDestacado.asp title="Ver Ejemplo" target=destacado>destacado</a>.&nbsp; 
	       </td></tr>
<%	end if %>				
</table>
<%
	dim numInquilinos

	sQuery= "SELECT CONT(*) as numInquilinos FROM Inquilinos i "
	sQuery= sQuery= " INNER JOIN Anuncios a ON i.id=a.id WHERE a.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
	numInquilinos=rst("numInquilinos")
	rst.Close

	sQuery= "SELECT a.Activo as aActivo, a.Provincia as aProvincia, a.Id as aId, p.Nombre as pNombre, i.TipoViv as iTipoViv, * "
	sQuery= sQuery & " FROM ((Inquilinos i INNER JOIN Anuncios a on i.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario) "
	sQuery= sQuery & " LEFT JOIN Provincias p ON a.provincia=p.id "
	sQuery= sQuery & " WHERE a.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
%>
<table width="100%">
	<tr>
		<td width="438" colspan=4 >			
			<p align="left"><b><font color="#FFFFFF" size="4">Anuncios de Búsqueda de Piso</font></b></tr>
	<tr>
		<td>
			<br>
<%	if numInquilinos>0 then %>
	<tr>
		<td title="¿Es visible tu anuncio?" width="100"><b>Estado</b></td>
		<td width="100" ><b>Tipo</b></td>
		<td width="200"><b>Título</b></td>
		<td width="100" ><b>Dónde</b></td>
		<td width="170"><b>Qué hago con él</b></td></tr>
<%		
	end if

	for i=1 to numInquilinos
%>
	<tr>
		<td width="100">
<%	if rst("aActivo")="No" then %>
			Oculto
<%	elseif rst("aProvincia")=0 then %>
			En proceso de publicación
<%	else %>
			Publicado
<%	end if %>									
		</td>
		<td width="100" height="50"><%=rst("iTipoViv")%></td>
		<td width="200" height="50"><a href=javascript:detalle('TrAnuncioDetalle.asp?tabla=Inquilinos&ID=<%=rst("aId")%>&admin=si')><%=rst("Cabecera")%></a></td>
		<td width="100" height="50"><%=rst("pNombre")%></td>
		<td width="143" height="50">
			<select onchange="if(value) 
<%	if rst("iTipoViv")="Compra" then %>
			location='TrCasaCompraDemandaFront.asp?op='+value+'&id=<%=rst("aId")%>';" 
<%	else %>
			location='TrCasaRegDemandaFront.asp?op=' 
<%	end if %>									
				name=<%=rst("aId")%> +value+'&id="<%=rst("aId")%>';&quot;">
				<option selected value=0>-Elige Una-
				<option value=Modificar>Modificar</option>
<%	if rst("aActivo")="No" then %>
				<option value=Activar>Reactivar</option>
<%	elseif rst("aProvincia")<>0 then %>
				<option value=Borrar>Ocultar</option>
<%	end if %>									
			</select></td></tr>
<%
		rst.movenext
	next
	rst.Close
%>				
	<tr>
		<td colspan="2" width="170">
<FORM method=get name=frm action=../TrCasaRegDemandaFront.asp?id=nuevo>
			<input type="submit" value="Añadir Anuncio de Búsqueda de Piso para Alquilar">
</form></td>
		<td>
<FORM method=get name=frm action=../TrCasaCompraDemandaFront.asp?id=nuevo>
			<input type="submit" value="Para Comprar" >
</form></td>
	</tr>
</table></form>
<%
	dim numCurris

	Session("Id")=""
	if Session("Usuario")="" or Session("Activo")<>"Si" then Response.Redirect "NuDomoh.asp?msg=Sesión+finalizada"
	
	sQuery= "SELECT count(*) as numCurris FROM Curriculums c INNER JOIN Anuncios a ON c.id=a.id WHERE a.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
	numCurris=rst("numCurris")
	rst.Close

	sQuery= "SELECT a.Activo as aActivo, a.Provincia as aProvincia, a.Id as aId, p.Nombre as pNombre, u.Foto as uFoto, * "
	sQuery= sQuery & " FROM ((Curriculums c INNER JOIN Anuncios a ON c.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario) "
	sQuery= sQuery & " LEFT JOIN Provincias p ON a.provincia=p.id "
	sQuery= sQuery & " WHERE a.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
%>
<FORM method=post name=frm action="../NuCurroRegDemandaFront.asp">
	<input type=hidden name=id value="nuevo">
<table width="100%">
	<tr>
		<td width="438" colspan=4 >			
			<p align="left"><b><font color="#FFFFFF" size="4">Curriculum</font></b></tr>
	<tr>
		<td>
			<br><br>
<%	if numCurris>0 then %>
	<tr>
		<td title="¿Visible para Usuarios?" width="50" ><b>Estado</b></td>
		<td width="100" ><b>Título</b></td>
		<td width="200" ><b>Dónde</b></td>
		<td width="200" ><b>Foto</b></td>
		<td width="170" ><b>Qué hago con él</b></td></tr>
<%		
	end if

	for i=1 to numCurris
%>
	<tr>
		<td width="50">
<%	if rst("aActivo")="No" then %>
			Oculto
<%	elseif rst("CatCurro")=0 then %>
			En proceso de publicación
<%	else %>
			Publicado
<%	end if %>									
		</td>
		<td width="100" height="50"><a href=javascript:detalle('TrAnuncioDetalle.asp?tabla=Curriculums&ID=<%=rst("aId")%>&admin=si')><%=rst("Cabecera")%></a></td>
		<td width="300" height="50"><%=rst("pNombre")%></td>
		<td valign="middle" align="center">
<% 			if rst("uFoto") <>"" then %>
				<a href=javascript:foto('http://domoh.com/<%=rst("uFoto")%>')><img src="http://domoh.com/mini<%=rst("uFoto")%>" height=50 alt="Pulsa para ver la foto"></a>
<%			end if %>
		</td>
		<td width="143" height="50">
			<select onchange="if(value) location='NuCurroRegDemandaFront.asp?op='+value+'&id=<%=rst("aId")%>';" 
				name=<%=rst("aId")%>>
				<option selected value=0>-Elige Una-
				<option value=Modificar>Modificar</option>
<%	if rst("aActivo")="No" then %>
				<option value=Activar>Reactivar</option>
<%	elseif rst("CatCurro")<>0 then %>
				<option value=Borrar>Ocultar</option>
<%	end if %>									
			</select></td></tr>
<%
		rst.movenext
	next
	rst.Close
%>				
	<tr>
		<td colspan="2" width="170">
			<input type="submit" value="Añadir Curriculum">
	</td></tr>
</table></form>
<%
	dim numTrabajos

	sQuery= "SELECT count(*) as numTrabajos FROM Trabajos t INNER JOIN Anuncios a ON t.id=a.id WHERE a.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
	numTrabajos=rst("numTrabajos")
	rst.Close

	sQuery= "SELECT a.Activo as aActivo, p.Nombre as pNombre, a.Id as aId, u.Foto as uFoto, * "
	sQuery= sQuery & " FROM ((Trabajos t INNER JOIN Anuncios a ON t.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario) "
	sQuery= sQuery & " LEFT JOIN Provincias p ON a.provincia=p.id "
	sQuery= sQuery & " WHERE a.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
%>
<FORM method=post name=frm action="../NuCurroRegOfrezcoFront.asp">
	<input type=hidden name=id value="nuevo">
<table width="100%">
	<tr>
		<td width="438" colspan=4 >			
			<p align="left"><b><font color="#FFFFFF" size="4">Ofertas de Trabajo</font></b></tr>
	<tr>
		<td>
			<br><br>
<%	if numTrabajos>0 then %>
	<tr>
		<td title="¿Los usuarios pueden ver la oferta?" width="50" ><b>Estado</b></td>
		<td width="100" ><b>Título</b></td>
		<td width="200"><b>Dónde</b></td>
		<td width="200" ><b>Foto</b></td>
		<td width="170" ><b>Qué hago con ella</b></td></tr>
<%		
	end if

	for i=1 to numTrabajos
%>
	<tr>
		<td width="50">
<%	if rst("aActivo")="No" then %>
			Oculto
<%	elseif rst("CatCurro")=0 then %>
			En proceso de publicación
<%	else %>
			Publicado
<%	end if %>									
		</td>
		<td width="100" height="50"><a href=javascript:detalle('TrAnuncioDetalle.asp?tabla=Trabajos&ID=<%=rst("aId")%>&admin=si')><%=rst("cabecera")%></a></td>
		<td width="300" height="50"><%=rst("pNombre")%></td>
		<td valign="middle" align="center">
<% 			if rst("uFoto") <>"" then %>
				<a href=javascript:foto('http://domoh.com/<%=rst("uFoto")%>')><img src="http://domoh.com/mini<%=rst("uFoto")%>" height=50 alt="Pulsa para ver la foto"></a>
<%			end if %>
		</td>
		<td width="143" height="50">
			<select onchange="if(value) location='NuCurroRegOfrezcoFront.asp?op='+value+'&id=<%=rst("aId")%>';" 
				name=<%=rst("aId")%>>
				<option selected value=0>-Elige Una-
				<option value=Modificar>Modificar</option>
<%	if rst("aActivo")="No" then %>
				<option value=Activar>Reactivar</option>
<%	elseif rst("CatCurro")<>0 then %>
				<option value=Borrar>Ocultar</option>
<%	end if %>									
			</select></td></tr>
<%
		rst.movenext
	next
	rst.Close
%>				
	<tr>
		<td colspan="2" width="170">
			<input type="submit" value="Añadir Oferta de Trabajo" ></td></tr>
</table></form>
<!-- #include file="IncPie.asp" -->





































































































