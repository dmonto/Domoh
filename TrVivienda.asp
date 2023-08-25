<!-- #include file="IncNuBD.asp" -->
<head><link href=TrDomoh.css rel=stylesheet type=text/css><title>Domoh - <%=MesgS("Anuncios de Vivienda","Lodging Adverts")%></title></head>
<body onload="window.parent.location.hash='top';">
<%  if Request("head")="si" then %>
<!--#include file="IncFrHead.asp"-->
<%  end if %>
<div class=container>
	<div class=row><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
    <div class=col-2><a href=http://www.materiagris.es target=_blank><img alt='Materia Gris' src=conts/MatGris.gif id=banner_4 /></a></div>
    <div class=col-8>
        <div id=verAnuncios1>
            <div><img src=images/TrCabVivienda<%=Session("Idioma")%>.gif /></div>
            <div class=tituVerAnuncios><% Response.Write MesgS("<span class=tituVerAnunciosBold>Ver</span> anuncios de", "<span class=tituVerAnunciosBold>Browse</span> adverts")%>:</div>
            <div>
  			    <div><img alt='Bullet' src=images/TrBulletAnuncios.gif /></div>
                <div>
				    <a title='<% Response.Write MesgS("Ver anuncios de vivienda", "Browse housing adverts")%>' href='TrViviendaBuscoFront.asp?tipo=alquiler'>
					<% Response.Write MesgS("Pisos y Habitaciones en Alquiler", "Flats and Rooms for Rent")%></a></div></div>
            <div><div><img alt='Bullet' src=images/TrBulletAnuncios.gif /></div><div><a href='TrViviendaBusco.asp?tipo=venta'><%=MesgS("Pisos o casas en Venta","Flats & Houses for Sale")%></a></div></div>
            <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href='TrViviendaBusco.asp?tipo=inquilino'><%=MesgS("Inquilinos","Tenants")%></a></div></div></div>
		<div class=publiAnuncios1>
		    <div class=tituVerAnuncios><% Response.Write MesgS("<span class=tituVerAnunciosBold>Publicar</span> anuncios de","<span class=tituVerAnunciosBold>Post</span> adverts")%>:</div>
            <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href=TrCasaRegOfrezcoFront.asp><%=MesgS("Pisos y Habitaciones en Alquiler","Flats and Rooms on Rent")%></a></div></div>
            <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href=TrCasaCompraOfrezcoFront.asp><%=MesgS("Pisos o casas en Venta","Flats & Houses on Sale")%></a></div></div>
            <div><div><img src=images/TrBulletAnuncios.gif /></div><div><a href=TrCasaRegDemandaFront.asp><%=MesgS("Inquilinos","Tenants")%></a></div></div></div></div>
    <div id=banners_lat>
	    <div>
            <script type=text/javascript>google_ad_client = "pub-0985599315730502"; /* Ppal */ google_ad_slot = "4588720766"; google_ad_width = 120; google_ad_height = 240;</script>
            <script type=text/javascript src='http://pagead2.googlesyndication.com/pagead/show_ads.js'></script></div>
        <div><a href=http://www.communitas.es target=_blank><img src=conts/Communitas.gif /></a></div><div><a href=http://www.aulacreactiva.com/ target=_blank><img src=conts/Creactiva.gif /></a></div></div></div>
<div><!-- #include file="IncPie.asp" --></div>
</body>
