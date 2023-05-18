<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim sPedido, sRespuesta, i, numPrecio, esVacio, sBody

	sPedido=Request("ds_order")
	sRespuesta=Request("ds_response")

	'---- Check if anything in shopping cart
	if IsArray(Session("cart"))=false or sPedido="" then
		dim tbCesta(19,1)
		Session("cart") = tbCesta
		Response.Write MesgS("Tu cesta esta vacía.","Your shopping cart is empty.")
	end if

	tbCesta=Session("cart")
	esVacio=True

	for i=LBound(tbCesta) to UBound(tbCesta)
		if tbCesta(i,0)<>"" and tbCesta(i,1)<>"" then esVacio=false
	next

	if esVacio then Response.Redirect "TrShopRevisa.asp"

	sQuery="INSERT INTO ShopOrders (id, usuario, nombre, apellidos, dir1, dir2, ciudad, provincia, cp, pais, fecha) "
	sQuery=sQuery & " VALUES (" & sPedido & ",'"& Session("Usuario") & "', ' " & Session("fname") & "','" & Session("lname") & "','" & Session("address1") & "','" & Session("address2") & "','" & Session("city") & "',"
	sQuery=sQuery & "'" & Session("state") & "','" & Session("zip") & "','" & Session("country") & "','" & Date() & " " & Time() & "')"
 	 	
	rst.Open sQuery, sProvider
	sBody="Pedido completado.<br>"
    sBody=sBody & "Número:<br>" 
    sBody=sBody & sPedido & "<br>"
    sBody=sBody & "Usuario:" & Session("Usuario")
	sBody=sBody & "<br>Enviar a:<br>" & Session("fname") & " " & Session("lname")
	sBody=sBody & "<br>" & Session("address1") & "<br>" & Session("address2")& "<br>" & Session("city") & " - " & Session("zip") & " " & session("state")
	sBody=sBody & "<br>" & Session("country") & "<br><br>Detalle:<br>" 
		
	'--- Next we need to store each of the items
	for i=LBound(tbCesta) to UBound(tbCesta)	
		if tbCesta(i,0)<>"" and tbCesta(i,1)<>"" then
			'--- Look up price per unit
			rst.Open "SELECT precio FROM ShopProducts WHERE id=" & tbCesta(i,0), sProvider
			if rst.Eof then	Response.Redirect "error.asp?msg=" & Server.URLEncode("We are unable to process your requst at present.")

			numPrecio=Replace(rst("precio"),",",".")
			rst.Close
			sQuery= "INSERT INTO ShopOrderDetalle ([order],producto,cantidad,precio) VALUES (" & sPedido & "," & tbCesta(i,0) & "," & tbCesta(i,1) & "," & numPrecio & ")"
			rst.Open sQuery, sProvider
			sBody=sBody & "Producto: " & tbCesta(i,0) & " - " & tbCesta(i,1) & " unidades - " & numPrecio & "€/unidad<br>"
		end if
	next
	Mail "hector@domoh.com", "Pedido Pagado", sBody

	'--- Finally, empty the cart
	Session("cart")=null
%>
<html>
<head><title><%= sNombreTienda%></title></head>
<body onload="window.parent.location.hash='top';" bgcolor=<%=sColorBG%> topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 text=<%=sColorText%> alink=<%=sColorLight%> link=<%=sColorLight%> vlink=<%=sColorLink%>>
<font face=Arial>
<% 
	header 
	categorymenu
%>

<center>
<table border=0>
	<tr>
		<td valign=top>
			<font face=helvetica size=6 color='<%= sColorDark %>'>

			<p align=left>
				<%=MesgS("Tu orden ha sido procesada","Your order has been processed")%><br>
				<font face='helvetica, arial' size=2 color=<%=sColorText%>>
				&#183; <%=MesgS("Gracias por comprar con ","Thanks for buying at ")%><%= sNombreTienda %>.<br><br></font></font>
	            <font face='helvetica' color=<%= sColorDark %>>
                <%=MesgS("Tu número de pedido es","Your order number is")%><font face=arial><font face=helvetica size=6 color=<%= sColorDark %>><font face='helvetica, arial' size="3"  color="<%= sColorText %>">
	<font face="helvetica, arial" size="3"  color="<%= COLerror %>">
	<b><%= sPedido %></b><font face="helvetica, arial" size="3"  color="<%= sColorText %>">.<br>
	<%=MesgS("Necesitarás esta información si tienes que contactar con nosotr@s","We'll ask for that information when you call us")%>.
	</td>

</table>
<br><br>

</font>

</body>
</html>


































