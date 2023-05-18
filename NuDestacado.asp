<!-- #include file="IncNuBD.asp" -->
<html>
<head>
    <title>Formato Destacado</title>
    <!-- #include file="IncTrCabecera.asp" -->
</head>
<body onload='window.resizeTo(800,300);'>
<div class=contanuncio>
	<div class=col-1><%=MesgS("Publicado","Posted")%></div><div><%=MesgS("Idioma","Lang")%></div><div><%=MesgS("Dónde","Where")%></div>
		<div><img alt='Precio' src=images/Eurito.gif class=euro /></div><div><%=MesgS("Tipo","Type")%></div>
		<div><%=MesgS("Foto","Picture")%></div><div><%=MesgS("Fumadores","Smokers?")%></div><div><%=MesgS("Mascotas","Pets?")%></div></div>
	<div class=anunciosDest onmouseover="this.bgColor='cyan';" onmouseout="this.bgColor='#CC99FF';">
		<div>30/11/2005 20:30</div><div><img src=images/Espanol.gif alt='Anuncio en Español'/></div><div><%=MesgS("Ibiza (Centro)","Ibiza (Downtown)")%></div>
		<div>220 €/<%=MesgS("mes","month")%></div><div><%=MesgS("Habitación","Room")%></div><div><img src=images/Cigarrito.gif alt='<%=MesgS("Admiten Fumadores","Smokers Allowed")%>'/></div>
		<div><img src=images/Perrillos.gif alt='<%=MesgS("Admiten Mascotas","Pets Allowed")%>'/></div></div></div></div>
</body></html>
