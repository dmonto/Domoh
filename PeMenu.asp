<!-- #include file="IncNuBD.asp" -->
<html>
<head>
<% 
	if Request("idioma")="Es" then 
		Session("Idioma")=""
	elseif Request("idioma")="En" then 
		Session("Idioma")="En"
	end if

	if Request("nuevoidioma")="Es" then 
		Session("Idioma")=""
		Response.Cookies("Idioma")="Es"
		Response.Cookies("Idioma").Expires=Now+365
	elseif Request("nuevoidioma")="En" then 
		Session("Idioma")="En"
		Response.Cookies("Idioma")="En"
		Response.Cookies("Idioma").Expires=Now+365
	end if
%>
    <title>
<%  if Session("Idioma")="" then %>
        Domoh - Pisos y Habitaciones Gay-Friendly
<%  else %>		
	    Domoh - Your Gay-Friendly Flats and Rooms 
<%  end if %>		
    </title>
    <meta content='Vacaciones, Alquiler, Alquileres para gay, gays, lesbiana, lesbianas de pisos y habitaciones en casas compartidas' name=description>
    <meta content='Vacaciones, Alquiler, Alquileres, gay, gays, lesbiana, lesbianas, pisos, habitaciones, casas compartidas, madrid, barcelona' name=keywords>
    <link href=TrDomoh.css rel=stylesheet type=text/css>
    <link rel='shortcut icon' href=favicon.ico type=image/icon><link rel=icon href=favicon.ico type=image/icon>
    <script async type=text/javascript src=forms<%=Session("Idioma")%>.js></script>
    <style>.full {width:100%}</style>
</head>
<body>
<div class=container>
    <div class=logo><img alt='Domoh' src=images/TrLogo.gif class=imagen /></div>
    <div class=banner><img src="images/TrBannerEnswf.jpg" class=full></div>
    <div class=topmenu>
        <a title='<%=MesgS("Conócenos","About us")%>' href='TrAbout.asp?idioma=<%=MesgS("Es","En")%>' target=menu class=linkutils>
            <img alt='<%=MesgS("Conócenos","About us")%>' src=images/Info.png class=icono /></a>
        <a title='<%=MesgS("English version","En Español")%>' href='PeMenu.asp?nuevoidioma=<%=MesgS("En","Es")%>'><img src='images/<%=MesgS("Eng","Esp")%>.svg' class=icono /></a></div>
    <div class=main> 
<%  if Session("MenuAction")<>"" then %>
        <div><%=Session("MenuAction")%></div>
<%  elseif Request("id")<>"" then %>
        <div>'QuAnuncioDetalle.asp?idioma=<%=MesgS("Es","En")%>&tabla=Pisos&id=<%=Request("id")%>'</div>
<%  else %>
        <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=menu href='QuVivienda.asp?idioma=<%=MesgS("Es","En")%>'>
	    <img class=ban_portada alt='Vivienda' src=images/TrCabViviendaT<%=Session("Idioma")%>.gif /></a>
	    <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href="TrVacas.asp?idioma=<%=MesgS("Es","En")%>">
	    <img class=ban_portada alt='Vacaciones' src=images/TrCabVacacionesT<%=Session("Idioma")%>.gif /></a>
        <a title='<%=MesgS("Entrar en tu zona privada","Enter your private area")%>' target=principal href='http://www.domoh.com/PeDomohLogOn.asp?idioma=<%=MesgS("Es","En")%>'>
            <img class=pestana src=images/TrHomeUsuarios<%=Session("Idioma")%>.gif id=Image1 onmouseover="ImgSwap('Image1','images/TrHomeUsuarios<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a>
        <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href='TrViviendaBuscoFront.asp?idioma=<%=MesgS("Es","En")%>'>
    	<img class=ban_portada alt='Buscar Vivienda' src=images/QuPortada<%=Session("Idioma")%>.jpg /></a></div>
    <!-- #include file="IncPie.asp" -->
    </div>
<%  end if %>
</body></html>
