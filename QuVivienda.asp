<!-- #include file="IncNuBD.asp" -->
<head><link href=TrDomoh.css rel=stylesheet type=text/css><title>Domoh - <%=MesgS("Anuncios de Vivienda","Lodging Adverts")%></title></head>
<body onload="window.parent.location.hash='top';">
<div class=container>
	<div class=logo><a title='<%=MesgS("Página Principal", "Main Page")%>' href=PeMenu.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
    <div class=main>
        <h2><img alt='Vivienda' src=images/TrCabVivienda<%=Session("Idioma")%>.gif /></h2>
        <h2><% Response.Write MesgS("<span class=tituVerAnunciosBold>Ver</span> anuncios de","<span class=tituVerAnunciosBold>Browse</span> adverts")%>:</h2>
        <h3>
            <img src=images/TrBulletAnuncios.gif />
            <a title='<% Response.Write MesgS("Ver anuncios de vivienda", "Browse housing adverts")%>' href='QuViviendaBuscoFront.asp?tipo=alquiler'>
                    <% Response.Write MesgS("Pisos y Habitaciones en Alquiler", "Flats and Rooms for Rent")%></a></h3>
        <h3><img src=images/TrBulletAnuncios.gif /><a href='QuViviendaBusco.asp?tipo=inquilino'><%=MesgS("Inquilinos","Tenants")%></a></h3>
        <h3><% Response.Write MesgS("<span class=tituVerAnunciosBold>Publicar</span> anuncios de","<span class=tituVerAnunciosBold>Post</span> adverts")%>:</h3>
        <h3><img src=images/TrBulletAnuncios.gif /><a href=TrCasaRegOfrezcoFront.asp><%=MesgS("Pisos y Habitaciones en Alquiler","Flats and Rooms on Rent")%></a></h3>
        <h3><img src=images/TrBulletAnuncios.gif /><a href=TrCasaRegDemandaFront.asp><%=MesgS("Inquilinos","Tenants")%></a></h3>
    <div class=foot><!-- #include file="IncPie.asp" --></div></div>
</body>