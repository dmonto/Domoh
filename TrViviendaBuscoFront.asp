<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javascript>
	function ClickPais() {document.location="TrViviendaBuscoFront.asp?Pais="+frm.pais.value; return true;}
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
<%  if Request("head")="si" then %>
<!--#include file="IncFrHead.asp"-->
<%  end if %>
<%  if Request("msg")<>"" then %>
	<h2><%=Request("msg")%></h2>
<%  end if %>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class=topmenu><a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a><%=MesgS("Buscar Vivienda en Alquiler","Look for Renting")%></div>
    <div class=formentrada>
        <div class=row>
    	    <form name=frm onsubmit='return checkForm();' action=TrViviendaBusco.asp method=get>
			    <input type=hidden name=tipo value=alquiler />
			<h1><%=MesgS("Mira de forma gratuita ","Look for free ")%> <%=MesgS("las casas, pisos o habitaciones en casa compartidas que están disponibles", "houses, flats and rooms currently available")%>.</h1>
			<h2><%=MesgS("País","Country")%>:<% SelectPais Session("Idioma"), "Vivir", Session("Pais") %></h2>
            <div>
                <div><p><%=MesgS("Provincia","Region")%>:</p></div>
                <div>
		            <p>
<% 
    if CLng(Request.Cookies("ProvinciaViv"))=0 then 
        SelectProvincia "Vivir", 0 
	else
	    SelectProvincia "Vivir", CLng(Request.Cookies("ProvinciaViv"))
	end if
%>
	                        </p></div></div>
			    <div>
			        <p>
                        <%=MesgS("Máximo alquiler","Maximum rent")%> <% Response.Write MesgS("que pagarías", "you are willing to pay")%>
                        <input title='<% Response.Write MesgS("Pon el precio en Euros","Rent in Euros")%>' size=5 name=precio /> €/<%=MesgS("Mes","Month")%></p></div>
			    <div><input name=Submit type=submit class=btnLogin value='<%=MesgS("Buscar","Search")%>'/></div></form></div>
            <div id=banners_lat>
                <div><a href=http://www.communitas.es target=_blank><img src=conts/Communitas.gif alt='Con la boca abierta' /></a><a href=# target=_blank></a></div>
                <!-- #include file="IncTrGoogleAn.asp" -->
                </div></div>
        <div><!-- #include file="IncPie.asp" --></div></div>
</body>
