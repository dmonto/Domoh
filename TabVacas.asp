<!-- #include file="IncNuBD.asp" -->
<div class=container>
	<nav><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></nav>
    <div class=row><img alt='Vacaciones' src=images/TrCabVacaciones<%=Session("Idioma")%>.gif /></div>
        <h1 class=tituVerAnuncios><% Response.Write MesgS("<span class=tituVerAnunciosBold>Ver</span> anuncios de", "<span class=tituVerAnunciosBold>Browse</span> adverts")%>:</h1>
        <div class=row>
    	    <div class=col-1><img alt='Punto' src=images/TrBulletAnuncios.gif /></div>
			<div><a title='<%=MesgS("Ver anuncios de vacaciones","Browse holiday adverts")%>' href=TrVacasBuscoFront.asp><%=MesgS("Apartamentos y Hoteles","Flats and Hotels")%></a></div></div>
        <div><div><img alt='Bullet' src=images/TrBulletAnuncios.gif /></div><div><a href='TrViviendaBusco.asp?tipo=vacasswap'><%=MesgS("Apartamentos en Intercambio","Flat Swapping")%></a></div></div>
        <div class=tituVerAnuncios><% Response.Write MesgS("<span class=tituVerAnunciosBold>Publicar</span> anuncios de","<span class=tituVerAnunciosBold>Post</span> adverts")%>:</div>
        <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href=TrVacasRegOfrezcoFront.asp><%=MesgS("Apartamentos y Hoteles","Flats and Hotels")%></a></div></div>
		</div></div>
</div>
