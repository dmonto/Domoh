<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<% 
	dim sWhere
	
	if Request("pais")<>"" then
		Session("Pais")=CLng(Request("pais"))
	elseif Session("Pais")=0 then
		Session("Pais")=34
	end if
%>
<script type=text/javascript>function ClickPais() {document.location="TrVacasBuscoFront<%=Session("Idioma")%>.asp?pais="+frm.pais.value; return true;}</script>
<body onload="window.parent.location.hash='top';"> 
<span id=bolatexto class=bolatexto></span>
<div class=container>
<% if Request("msg")<>"" then %>
	<h1 class=banner><%=Request("msg")%></h1>
<% end if %>
	<nav>
		<div class=col-3><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
		<div class=col-9><a href=TrVacas.asp class=linkutils><%=MesgS("Anuncios Vacaciones","Holiday Adverts")%> &gt;</a><%=MesgS("BUSCAR Alojamiento de Vacaciones","LOOK FOR Holiday Lettings")%></div></nav>
	<section class=col-1>
		<div><a href=http://www.aulacreactiva.com/ target=_blank><img src=conts/Creactiva.gif alt='Con la boca abierta' /></a><a href=# target=_blank></a></div>
		<div><a href=http://www.materiagris.es target=_blank><img alt='Materia Gris' src=conts/MatGris.gif id=banner_4 /></a></div></section>
    <section class=col-11>
	    <form name=frm onsubmit='return checkForm();' action=TrViviendaBusco.asp method=get>
   		    <input type=hidden name=tipo value=vacas />
        <h1>
            <% Response.Write MesgS("Mira de forma gratuita ", "Search now for free ")%> 
            <% Response.Write MesgS("los hoteles, apartamentos o habitaciones en casa compartidas que están disponibles", "available hotels, flats and rooms in shared houses")%>.</h1>
        <div>
<%		
	if Session("Pais")=34 then
		sWhere = "WHERE pais=" & Session("Pais") & " AND id IN (SELECT provincia FROM Anuncios WHERE activo='Si' AND tabla='Pisos') "
		rst.Open "SELECT * FROM Provincias " & sWhere & " ORDER BY nombre", sProvider
%>
	        </div>
        <div>
            <img alt='Mapa' title='<% Response.Write MesgS("En esta zona aún no tenemos alojamientos!", "No available lodging there at this time")%>' src=/images/MapaCiudades.gif usemap=#mapaEspana />
            <map name=mapaEspana>
<%		while not rst.Eof %>
		        <area onmouseout="document.getElementById('bolatexto').innerText='';" onmouseover="document.getElementById('bolatexto').innerText='<%=rst("nombre")%>';" 
    			    shape=poly coords='<%=rst("Poligono")%>' href='TrViviendaBusco.asp?tipo=vacas&provincia=<%=rst("Id")%>' alt="<%=rst("Nombre")%>">
<%
			rst.Movenext
		wend
		rst.Close
%>	
			    </map></div>
<%  end if %>
        <h2><%=MesgS("País","Country")%>:<% SelectPais "", "Vacas", Session("Pais") %></h2>
        <div><div><p><%=MesgS("Provincia","Region")%>:</p></div><div><p><% SelectProvincia "Vacas", 0 %></p></div></div>
        <div>
            <%=MesgS("Máximo precio","Maximum price")%> <%=MesgS("que prefieres","you are willing to pay")%>
            <input title='<%=MesgS("Pon el precio en Euros","Weekly Price in Euros")%>' size=5 name=precio /> €/<%=MesgS("Semana","Week")%></div>
        <div> <input name=Submit type=submit class=btnLogin value='<%=MesgS("Buscar","Search")%>'/></div></form></section>
    <div><!-- #include file="IncPie.asp" --></div></div>
</body>