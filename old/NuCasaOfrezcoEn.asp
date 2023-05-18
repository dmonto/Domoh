<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim sWhere, sOrder, numInquilinos, i, sHora

	sWhere= "WHERE a.activo='Si' AND i.tipoviv<>'Compra' "
	sWhere= sWhere & " AND a.usuario IN (SELECT usuario FROM Usuarios u WHERE u.activo='Si') "
	sWhere= sWhere & " AND a.provincia <> 0 "

	sQuery="SELECT COUNT(*) AS numInquilinos FROM Inquilinos i "
	sQuery=sQuery & " INNER JOIN Anuncios a ON i.id=a.id " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), _
		sQuery & " - " & Err.Description
	numInquilinos=rst("numInquilinos")
	rst.Close

	if Request("orden")="fecha" then
		sOrder= "a.fechaalta DESC"
	elseif Request("orden")="precio" then
		sOrder= "i.maximo"
	elseif Request("orden")="idioma" then
		sOrder= "a.idiomaen DESC, a.fechaultimamodificacion DESC"
	else
		sOrder= "a.fechaalta DESC"
	end if
	
	sQuery="SELECT p.nombre AS pNombre, a.id AS aId, u.foto AS uFoto, i.ciudad AS iCiudad, * "
	sQuery=sQuery & " FROM ((Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) "
	sQuery=sQuery & " INNER JOIN Provincias p ON a.provincia=p.id) "
	sQuery=sQuery & " INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere & " ORDER BY " & sOrder
	rst.Open sQuery, sProvider
%>
<body onLoad="window.parent.location.hash='top';
"> 

<table>
	<tr>
		<td colspan=4 rowspan=2 width=800><% if Request("msg")<>"" then %>
<p ><%=Request("msg")%></p>
<% end if %>
</td>
<!--#include file="IncNuSubMenu.asp"-->
	<tr>
		<td colspan=4>
			<ul type=circle>
				<li><a title="Publish your Flat/House for Rent" href=../NuCasaRegOfrezcoFrontEn.asp>
					<b>Post Flat or House</b></a></li></ul>
	</td></tr>
	<tr>
		<td colspan=6 width=500>
<%	
	if numInquilinos=0 then
%>
		<b>Click on "Post Flat or House" and post your advert.</b>
<%		
		Response.End
	end if

	PagCabezaEn numInquilinos, "People seeking House or Flat"
%>
	</td></tr>
	<tr>	
		<td><a title="See most recent first" href=NuCasaOfrezcoEn.asp?orden=fecha>by Date</a></td>
    	<td><a title="See adverts in English first" href=NuCasaOfrezcoEn.asp?orden=idioma>by Lang</a></td>
    	<td></td>
    	<td></td>
    	<td>
    		<a title="See people willing to pay more first" href="NuCasaOfrezcoEn.asp?orden=precio">by Rent</a></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td></tr>
	<tr>	
		<td title="Last Update"><b>Posted</b></td>
		<td title="Is it in English?"><b>Lang</b></td>
    	<td title="City/Area"><b>Where</b></td>
    	<td title="Click to see the advert"><b>Ad</b></td>
    	<td title="Monthly Rent">
    		<img src="../images/Eurito.gif" width="25" height="25" ></td>
    	<td title="What is it?" align=center><b>Type</b></td>
    	<td title="Click to see picture"><b>Picture</b></td>
    	<td><b>Smokers?</b></td>
    	<td><b>Pets?</b></td></tr>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
%>
	<tr valign=middle <%if (i mod 2)=1 then response.write "bgcolor=#CC99FF"%>
  		onmouseover="this.bgColor='cyan';" 
  		onmouseout="if (this.rowIndex % 2) {this.bgColor='#CC99FF'} else {this.bgColor='#FFFFFF'};"
  		onclick="javascript:detalle('TrAnuncioDetalle.asp?tabla=Inquilinos&ID=<%=rst("aId")%>')">
   		<td align=center valign="middle">
<%
			sHora=Hour(rst("FechaUltimaModificacion")) & ":" & minute(rst("FechaUltimaModificacion"))
			if sHora="0:0" then
				Response.write formatdatetime(rst("FechaUltimaModificacion"),2)
			else
				Response.write formatdatetime(rst("FechaUltimaModificacion"),2) & " " & formatdatetime(rst("FechaUltimaModificacion"),4)
			end if
%> 		
    	</td>
   		<td width=50 align="center" valign="middle">
<%			if UCase(rst("IdiomaEs"))="ON" then %>
			<img src="../images/Espanol.gif" width=20 alt="Advert in Spanish">
<%			end if %>
<%			if UCase(rst("IdiomaEn"))="ON" then %>
			<img src="../images/Ingles.gif" width=20 alt="Advert in English">
<%			end if %>
		</td>
   		<td valign="middle">
    		<%=rst("pNombre")%> (<%=rst("iCiudad")%>)</td>
	<td valign="middle">
		<a href=javascript:detalle('TrAnuncioDetalles.asp?Tabla=Inquilinos&ID=<%=rst("aId")%>')>
		<%=rst("Cabecera")%></a></td>
		<td valign="middle"><%=rst("Maximo")%> <font face=Arial>€</font>/month</td>
    </td>
    <td valign="middle" align="center">
	<% if rst("TipoViv")="Alquiler Piso" then Response.Write "Flat" else Response.Write "Room" %>
    </td>
    <td valign="middle" align="center">
<%			If rst("uFoto") = "Si" then %>
		<a href=javascript:foto('http://domoh.com/<%=rst("uFoto")%>')><img src="../images/Camara.gif" height=50 alt="Click here to watch the pictures"></a>
<%			elseif rst("uFoto") <> "" then %>		
		<a href=javascript:foto('http://domoh.com/<%=rst("uFoto")%>')><img src="http://domoh.com/mini<%=rst("uFoto")%>" height=50 alt="Click here to watch the pictures"></a>
<%			End if %>    
    </td>
    <td valign="middle" align="center">
<%			If UCase(rst("Fuma")) = "ON" then %>
	    <img src="../images/Cigarrito.gif" alt="Smoker">
<%			else %>
   		<img src="../images/NoCigarrito.gif" alt="Non Smoker" >
<%			end if %>
		</td>
    <td valign="middle" align="center">
<%			If UCase(rst("Mascota")) = "ON" then %>
    <img src="../images/Perrillos.gif" alt="Pet">
<%			else %>
    <img src="../images/NoPerrillos.gif" alt="No Pets">
<%			end if %>
	</td></tr>
<%
		end if
		rst.movenext
	next

	PagPie numInquilinos, "NuCasaOfrezcoEn.asp?orden=" & request("orden") & "&"
%>
    </table>
<!-- #include file="IncPie.asp" -->



























































