<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javascript>
	function ClickPais() {document.location="QuViviendaBuscoFront.asp?Pais="+frm.pais.value; return true;}
	function checkForm() {if(!checkNumber(document.frm.precio, "Precio Máximo")) return false; return true;}
</script>
<% 
	if Request("pais")<>"" then
		Session("Pais")=CLng(Request("pais"))
	elseif Session("Pais")=0 then
		Session("Pais")=34
	end if
%>
<body onload="window.parent.location.hash='top';"> 
<div class=container>
<%  if Request("msg")<>"" then %>
	<h1><%=Request("msg")%></h1>
<%  end if %>
	<div class=logo><a href=PeMenu.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class=banner><a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a><%=MesgS("Buscar Vivienda en Alquiler","Look for Renting")%></div>
	<div class=main>
	    <form name=frm onsubmit='return checkForm();' action=QuViviendaBusco.asp method=get>
   		    <input type=hidden name=tipo value=alquiler />
		<h1><%=MesgS("Mira de forma gratuita ","Look for free ")%> <% Response.Write MesgS("las casas, pisos o habitaciones en casa compartidas que están disponibles", "houses, flats and rooms currently available")%>.</h1>
		<h2><%=MesgS("País","Country")%>:</h2>
		<h2><% SelectPais Session("Idioma"), "Vivir", Session("Pais") %></h2>
		<h2><%=MesgS("Provincia","Region")%>:</h2>
		<h2>
<% 
	if CLng(Request.Cookies("ProvinciaViv"))=0 then 
        SelectProvincia "Vivir", 0 
    else
        SelectProvincia "Vivir", CLng(Request.Cookies("ProvinciaViv"))
    end if
%>
		</h2>
		<h1>
		    <%=MesgS("Máximo alquiler que estás dispuesto a pagar","Maximum rent you are willing to pay")%> 
            <input title='<% Response.Write MesgS("Pon el precio en Euros","Rent in Euros")%>' size=5 name=precio /> &euro;/<%=MesgS("Mes","Month")%></h1>
		<h1><input name=Submit type=submit class=btnLogin value='<%=MesgS("Buscar","Search")%>'/></h1></form></div>
    <div class=foot><!-- #include file="IncPie.asp" --></div></div>
</body>
