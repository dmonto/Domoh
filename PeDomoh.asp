<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';">
<div class=row>
	<a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href='QuVivienda.asp?idioma=<%=MesgS("Es","En")%>'>
	<img class=ban_portada alt='Vivienda' src=images/TrCabViviendaT<%=Session("Idioma")%>.gif /></a>
	<a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href="TrVacas.asp?idioma=<%=MesgS("Es","En")%>">
	<img class=ban_portada alt='Vacaciones' src=images/TrCabVacacionesT<%=Session("Idioma")%>.gif /></a></div>
<div class=row>
	<a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href='TrViviendaBuscoFront.asp?idioma=<%=MesgS("Es","En")%>'>
	<img class=ban_portada alt='Buscar Vivienda' src=images/QuPortada<%=Session("Idioma")%>.jpg /></a></div>
<!-- #include file="IncPie.asp" -->
</body>