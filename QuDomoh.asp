<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<% 
    dim sTarget
    sTarget="_self"
    if Request("head")<>"no" then 
        sTarget="principal"
%>
<style>.domoh {padding: 0; width: 100%; border: 0; background-image: url('images/TrFondo.gif'); background-attachment: fixed}</style>
<body onload="window.parent.location.hash='top';" class=domoh>
<% end if %>
<div>
	<a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=<%=sTarget%> href='TrVivienda.asp?idioma=<%=MesgS("Es","En")%>&head=si'>
	<img alt='Vivienda' src=images/TrCabViviendaT<%=Session("Idioma")%>.gif /></a></div>
<div><img alt='Blanco' src=images/TrEspacio.gif /></div>
<div>
    <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=<%=sTarget%> href="TrVacas.asp?idioma=<%=MesgS("Es","En")%>&head=si">
	<img alt='Vacaciones' src=images/TrCabVacacionesT<%=Session("Idioma")%>.gif /></a></div>
<div>
	<a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=<%=sTarget%> href="TrViviendaBuscoFront.asp?idioma=<%=MesgS("Es","En")%>&head=si">
	<img alt='Busca Casa' src=images/QuPortada<%=Session("Idioma")%>.jpg /></a></div>
<% if Request("head")<>"no" then %>
    <div><!-- #include file="IncPie.asp" --></div>
<% end if %>
</body>
