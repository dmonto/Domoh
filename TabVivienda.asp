<!-- #include file="IncNuBD.asp" -->
<div class=container>
	<div class=logo><a title='<%=MesgS("Página Principal", "Main Page")%>' href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
    <div class=row>
        <div class=row><img alt='Vivienda' src=images/TrCabVivienda<%=Session("Idioma")%>.gif /></div>
        <div class='row tituVerAnuncios'><% Response.Write MesgS("<span class=tituVerAnunciosBold>Ver</span> anuncios de","<span class=tituVerAnunciosBold>Browse</span> adverts")%>:</div>
        <div class=row>
    	    <div class="left"><img src=images/TrBulletAnuncios.gif /></div>
            <div>
    		    <a title='<% Response.Write MesgS("Ver anuncios de vivienda", "Browse housing adverts")%>' href='QuViviendaBuscoFront.asp?tipo=alquiler'>
                    <% Response.Write MesgS("Pisos y Habitaciones en Alquiler", "Flats and Rooms for Rent")%></a></div></div>
        <div><img src=images/TrBulletAnuncios.gif /></div>
        <div><a href='QuViviendaBusco.asp?tipo=venta'><%=MesgS("Pisos o casas en Venta","Flats & Houses for Sale")%></a></div>
        <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href='QuViviendaBusco.asp?tipo=inquilino'><%=MesgS("Inquilinos","Tenants")%></a></div></div></div></div>
<div class=publiAnuncios1>
    <div class=tituVerAnuncios><% Response.Write MesgS("<span class=tituVerAnunciosBold>Publicar</span> anuncios de","<span class=tituVerAnunciosBold>Post</span> adverts")%>:</div>
    <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href=JQViviendaBusco.asp><%=MesgS("Pisos y Habitaciones en Alquiler","Flats and Rooms on Rent")%></a></div></div>
    <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href=TrCasaRegDemandaFront.asp><%=MesgS("Inquilinos","Tenants")%></a></div></div></div></div>