<!-- #include file="IncNuBD.asp" -->
<head><link href=TrDomoh.css rel=stylesheet type=text/css><title><%=MesgS("Vacaciones","Holidays")%></title></head>
<body onload="window.parent.location.hash='top';">
<%  if Request("head")="si" then %>
    <!--#include file="IncFrHead.asp"-->
<%  end if %>
<div class=container>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
    <div class=main><img alt='Vacaciones' src=images/TrCabVacaciones<%=Session("Idioma")%>.gif /></div>
        <h1 class=tituVerAnuncios><% Response.Write MesgS("<span class=tituVerAnunciosBold>Ver</span> anuncios de", "<span class=tituVerAnunciosBold>Browse</span> adverts")%>:</h1>
        <div class=col-1><img alt='Punto' src=images/TrBulletAnuncios.gif /></div>
        <div><a title='<%=MesgS("Ver anuncios de vacaciones","Browse holiday adverts")%>' href=TrVacasBuscoFront.asp><%=MesgS("Apartamentos y Hoteles","Flats and Hotels")%></a></div
        <div><div><img alt='Bullet' src=images/TrBulletAnuncios.gif /></div><div><a href='TrViviendaBusco.asp?tipo=vacasswap'><%=MesgS("Apartamentos en Intercambio","Flat Swapping")%></a></div></div>
        <div class=tituVerAnuncios><% Response.Write MesgS("<span class=tituVerAnunciosBold>Publicar</span> anuncios de","<span class=tituVerAnunciosBold>Post</span> adverts")%>:</div>
        <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href=TrVacasRegOfrezcoFront.asp><%=MesgS("Apartamentos y Hoteles","Flats and Hotels")%></a></div></div>
		</div></div>
    <div class=row id=banners_lat>
        <!-- #include file="IncTrGoogle.asp" -->
        <div><a href=http://www.materiagris.es target=_blank><img src=conts/MatGris.gif id=banner_4 /></a></div>
		<div><a title='Aula Creactiva' href=http://www.aulacreactiva.com/ target=_blank><img src=conts/Creactiva.gif /></a></div></div>
    <div><!-- #include file="IncPie.asp" --></div>
</body>
