<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<title>Owner Profile</title>
<%
	dim numPisos, i, sPantalla

	Session("Id")=""
	if Session("Usuario")="" or Session("Activo")<>"Si" then 
		Response.Redirect "NuDomoh.asp?msg=Session+Finished"
	end if
	if Session("Usuario")="hector" then Response.Redirect "NuAdminPisos.asp?tipo=nuevos"
	
	sQuery= "SELECT COUNT(*) AS numPisos FROM Pisos p WHERE p.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
	numPisos=rst("numPisos")
	rst.Close

	if numPisos then
		sQuery= "SELECT p.dir1 AS pDir1, p.activo AS pActivo, p.provincia AS pProvincia, p.id AS pId, "
		sQuery= sQuery & " p.tipo AS pTipo, * "
		sQuery= sQuery & " FROM Pisos p INNER JOIN Usuarios u ON p.usuario=u.usuario "
		sQuery= sQuery & " WHERE p.usuario='"& Session("Usuario") & "' "
		rst.Open sQuery, sProvider

		if rst("rentaviv")<>0 then
			sPantalla="TrCasaRegOfrezcoFrontEn.asp"
		elseif rst("rentavacas")<>0 then
			sPantalla="TrVacasRegOfrezcoFront.asp"
		else
			if rst("precio")<>0 then
				sPantalla="TrCasaCompraOfrezcoFront.asp"
			else
				sPantalla="TrVacasSwapOfrezcoFront.asp"
			end if
		end if
	end if
%>
<!--#include file="IncTrCabecera.asp"-->
<script language=JavaScript>
function foto(id){
	url="TrFotos.asp?id="+id;
	sw1=window.open(url,'searchpage',
		'toolbar=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,' +
		'copyhistory=no,width=700,height=600');
	sw1.focus();
}
</script>
<body onLoad="window.parent.location.hash='top';"> 
<form method=post name=frm action=<%=sPantalla%>>
	<input type=hidden name=id value=nuevo>
<table width=100% >
	<tr><td colspan=3 width=500><%	if Request("msg")<>"" then %>
	<tr><td colspan=4><font size=4><b><%=Request("msg")%></font></b></td></tr>
<%	end if %>
</td><!--#include file="IncNuSubMenu.asp"--></tr>
	<tr>
		<td width=438 colspan=4>
			<p align=left>
				Update <a title="Update Profile" href=../TrUsuario.asp?origen=QuHomeUsuario.asp>here</a> 
				your personal data.
			<p align=left><b>Your Flats and Houses, <%=Session("Nombre")%></b>
			<br><br>
<%	if numPisos=0 then %>
			You haven't got any house/flat yet.
<%	
	end if

	if numPisos>0 then 
%>
	<tr>
		<td title="Can users access your ad?" width=50 >
			<b>Status</b></td>
		<td title="What is it?" width=100 ><b>Type</b></td>
		<td title="City/Town" width=200 ><b>Location</b></td>
		<td title="Weekly/Monthly/All-in" width=93><b>Price</b></td>
		<td title="Click to see pictures" width=80><b>Picture?</b>
			</td>
		<td title="Modify or Remove" width=170 ><b>What shall I do?
			</b></td></tr>
<%		
	end if

	for i=1 to numPisos
%>
	<tr>
		<td title="Can users access your ad?" width=50>
<%	if rst("pActivo")="No" then %>
			Hidden
<%	elseif rst("pProvincia")=0 then %>
			Awaiting publication
<%	else %>
			Published
<%	end if %>									
		</td>
		<td title="What is it?" width=100 height=50>
<% if rst("pTipo")="Casa" then Response.Write "House" else Response.Write "Flat" %>
		</td>
		<td title="City/Town" width=300 height=50>
			<a href=javascript:detalle('TrAnuncioDetalle.asp?tabla=Pisos&admin=si&id=<%=rst("pId")%>')>
			<%=rst("zona")%> - <%=rst("ciudadnombre")%></a></td>
		<td title="Weekly/Monthly/All-in" width=150 height=50>
<%	if rst("rentaviv")>0 then %>
			<%=rst("rentaviv")%> <font face=arial>€</font>/month
<%	elseif rst("rentavacas")>0 then %>
			<%=rst("rentavacas")%> <font face=arial>€</font>/week
<%	else %>
<%		if rst("Precio")>0 then %>
			<%=rst("Precio")%> thousand <font face=arial>€</font>
<%		else %>
			For Swapping
<%		end if %>
<%	end if %>
		</td>
		<td title="Click to see pictures" width=80 height=50>
<%	if rst("Foto") = "Si" then %>
		<a href=javascript:foto('<%=rst("pId")%>')>
		<img src="../images/Camara.gif" height=50 alt="Click here to watch pictures"></a>
<%	elseif rst("foto") <> "" then %>		
		<a href=javascript:foto('<%=rst("pId")%>')>
		<img src="http://domoh.com/mini<%=rst("Foto")%>" height=50 alt="Click here to watch pictures">
		</a>
<%	end if %>    
		</td>
		<td title="Modify or Remove" width=143 height=50>
			<select onchange="if(value) location='<%=sPantalla%>?op='+value+'&id=<%=rst("pId")%>';" 
					name=<%=rst("pId")%>>
				<option selected value=0>-Choose One-
				<option value=Modificar>Update</option>
<%	if rst("pActivo")="No" then %>
				<option value=Activar>Reactivate</option>
<%	elseif rst("pProvincia")<>0 then %>
				<option value=Borrar>Hide</option>
<%	end if %>									
				</select></td></tr>
<%
		rst.Movenext
	next
%>				
	<tr>
		<td colspan=3>
			<input title=Grabar type=submit value="Add New" >
	</td></tr>
</table></form>
<!-- #include file="IncPie.asp" -->

