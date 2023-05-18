<!-- #include file="IncNuBD.asp" -->
<%
	dim numCat, sCatName, sFoto, sDesc

	numCat=Request("catcode")
	rst.Open "SELECT * FROM CatShop WHERE id=" & numCat, sProvider
	sCatName=rst("nombre")
	sFoto=rst("foto")
	sDesc=rst("descripcion")
	if rst.Eof then Response.Write "Unable to locate category!"
	rst.Close
%>
<html>
<head><title><%=sNombreTienda%></title></head>
<body onload="window.parent.location.hash='top';" bgcolor=<%=sColorBG%> topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 text=<%=sColorText%> alink=<%=sColorLight%> link=<%=sColorLight%> vlink=<%=sColorLink%>>
	<font face=Arial>
	<% categorymenu %></font>
<table>
	<tr>
		<td valign=top align=left>
			<font face=helvetica size=3 color=<%=sColorDark%>><b><%=sCatName%></b></font><br>
			<font size=1 face=helvetica color=<%=sColorLight%>>
<%
	'---- Display list of other products in category
	'---- Get names and codes of all products in that category
	sQuery="SELECT nombre, id FROM ShopProducts WHERE catshop=" & numCat & " ORDER BY nombre"
	rst.Open sQuery, sProvider
	if not rst.Eof then
		rst.Movefirst
		while not rst.Eof
%>
			<a title='Ver el Producto' href=TrShopProducto.asp?id=<%=rst("id")%>><%=rst("nombre")%></a><br><br>
<%
			rst.Movenext
		wend
	end if
%>
			</font></td>
<% if sFoto<>"" then %>
		<td title='Foto de la Categoria' width=180 align=left><%=sFoto%>
<% else %>
		<td width=300 align=center height=250 valign=center>
				<font face='Verdana,Arial' size=9><b><%=LCase(sCatName) %></b>@<%=lcase(sNombreTienda) %></font><br><br>
				<font face='Verdana,Arial' size=3><b><%=LCase(sDesc) %></b></font>
<% end if	%>
				</td></tr></table>
<table>
	<tr>
	<tr><td height=30>&nbsp;</td></tr>
	<tr><td align=center valign=top><!-- #include file="IncPie.asp" --></td></tr>
	<tr><td height=20 align=center>&nbsp;</td></tr>
	<tr><td>&nbsp;</td><td height=5 align=center valign=top background=images/TrSombraAbajo.gif class=filaestrecha>&nbsp;</td><td>&nbsp;</td></tr></table>
<hr color=#CCCCCC size=1 noshade>
</body>
