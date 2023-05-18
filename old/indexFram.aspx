<!-- #include file="IncNuBD.aspx" -->
<html>
<head data-hayidioma='@Session("hayidioma")' data-idioma='@Session("Idioma")' data-idiomaesen='@Session("IdiomaEsEn")' data-req='@Request("idioma")' data-nuevo='@Request("nuevoidioma")'>
	<title>Domoh - Alojamiento Gay, Alquiler Gay, Pisos Gay, Habitación Gay - Hébergement Gay</title>
	<meta name=google-site-verification content=TmUeHVRRwXj313YY12Q9Zep315nSOIh5a-wbWheC1VM><meta name=msvalidate.01 content='E42AAFC279BD732454AEBD988B18EDEC'>
    <meta name=viewport content='width=device-width, initial-scale=1'>
    <meta content='text/html;charset=iso-8859-1' http-equiv=Content-Type>
	<meta content='Vacaciones, Alquiler, Alquileres para gay, gays, lesbiana, lesbianas de pisos y habitaciones en casas compartidas' name=description>
	<meta content='Vacaciones, Alquiler, Alquileres, gay, gays, lesbiana, lesbianas, pisos, habitaciones, casas compartidas, madrid, barcelona' name=keywords>
    <meta http-equiv=content-language content='es-es'>
    <link rel='shortcut icon' href=favicon.ico type=image/icon>
    <link rel=stylesheet type=text/css href=TrDomoh.css>
    <!-- #include file="IncTrGoogleAn.asp" -->
    <script type=text/javascript src=efectos.js></script><script type=text/javascript src='forms@Session("Idioma").js'></script>
</head>
<body style = "background-image: url('images/TrFondo.gif'); background-attachment: fixed">
    if Request("nuevoidioma")="Es" then
        Session("Idioma")=""
        Session("IdiomaEsEn")="Es"
        Response.Cookies("Idioma")="Es"
        Response.Cookies("Idioma").Expires=Now+365
    elseif Request("nuevoidioma")="En" then
        Session("Idioma")="En"
        Session("IdiomaEsEn")="En"
        Response.Cookies("Idioma")="En"
        Response.Cookies("Idioma").Expires=Now+365
    elseif Request("idioma")="Es" then
        Session("Idioma")=""
    elseif Request("idioma")="En" then
        Session("Idioma")="En"
    end if
<a name=top />
<header class=container>
    <div><a title='Página Principal' href='QuDomoh.asp?idioma=@MesgS("Es","En")' target=principal><img alt='Domoh' src=images/TrLogo.gif /></a></div>
    <div>
        <script type=text/javascript>
            banner1 = new Banner('banner1');
            sQuery="SELECT * FROM Contenidos WHERE pagina='Home" & Session("Idioma") & "' AND posicion>=10"
            rst.Open sQuery, sProvider
            while not rst.Eof
                if Right(rst("banner"),3) = "swf" then
                   banner1.add("IMAGE", "@Replace(rst("banner"),".","")&".jpg"", 15, 60, 468,"@rst("html")");
                elseif rst("banner")<>"" then 
                   banner1.add("IMAGE", "@rst("banner")", 15, 60, 468,"@rst("html")");
                end if
                rst.Movenext
            wend
            rst.Close
            document.write(banner1); banner1.start();
            </script></div>
    <div>
        <div><img alt='Contacto' src=images/TrContact.gif /></div>
        <div><a title='@MesgS("Contactanos","Contact us")' href='TrContactanos.asp?idioma=@MesgS("Es","En")' target=principal class=linkutils>@MesgS("Contacto","Contact us")</a></div></div>
        <div>
            <div><img alt='Conócenos' src=images/TrConocenos.gif /></div>
            <div><a title='@MesgS("Conócenos","About us")' href='TrAbout.asp?idioma=@MesgS("Es","En")' target=principal class=linkutils>@MesgS("Conócenos","About us")</a></div></div>
        <div>
            <div><a title='@MesgS("English version","En Español")' href='QuMenu.asp?nuevoidioma=@MesgS("En","Es")'><img src='images/Tr@MesgS("Eng","Esp").gif' /></a></div>
            <div><a title='@MesgS("English version","En Español")' href='QuMenu.asp?nuevoidioma=@MesgS("En","Es")' class=linkutils>@MesgS("English","Español")</a></div></div>
        <div>
            <div><img src=images/TrDonar.gif /></div>
            <div><a title='@MesgS("Donar","Donate")' href='TrDonar.asp?idioma=@MesgS("Es","En")&destino=QuDomoh.asp' target=principal class=linkutils>@MesgS("Donaciones","Donations")</a></div></div></header>
 <div>
    <div>
        <a title='@MesgS("Entrar en tu zona privada","Enter your private area")' target=principal href='http://www.domoh.com/TrDomohLogOn.asp?idioma=@MesgS("Es","En")'>
            <img src='images/TrHomeUsuarios@Session("Idioma").gif' id=Image1 onmouseover="ImgSwap('Image1','images/TrHomeUsuarios@Session("Idioma")_O.gif')" onmouseout=ImgRestore() /></a></div>
    <div>@MesgS("Mis anuncios","My Ads") &gt; </div>
    <div>
        <a title='@MesgS("Ver y Publicar Anuncios","Watch and Post Ads")' target=principal href='TrVivienda.asp?idioma=@MesgS("Es","En")'>
            <img src='images/TrAdsVivienda@Session("Idioma").gif' id=boton_vivienda onmouseover="ImgSwap('boton_vivienda','images/TrAdsVivienda@Session("Idioma")_O.gif')" onmouseout=ImgRestore() /></a></div>
    <div>
        <a title='@MesgS("Ver y Publicar Anuncios","Watch and Post Ads")' target=principal href='TrVacas.asp?idioma=@MesgS("Es","En")'>
            <img src='images/TrAdsVacaciones@Session("Idioma").gif' id=boton_vacaciones onmouseover="ImgSwap('boton_vacaciones','images/TrAdsVacaciones@Session("Idioma")_O.gif')" onmouseout=ImgRestore() /></a></div>
    <!-- AddThis Button BEGIN -->
    <div class='addthis_toolbox addthis_default_style'>
        <a href='http://addthis.com/bookmark.php?v=250&amp;username=dmonto' class=addthis_button_compact>Share</a><span class=addthis_separator>|</span><a class=addthis_button_facebook></a>
        <a class=addthis_button_myspace></a><a class=addthis_button_google></a><a class=addthis_button_twitter></a></div><script type=text/javascript src='http://s7.addthis.com/js/250/addthis_widget.js#username=dmonto'></script>
        <!-- AddThis Button END -->
        </div>
<div>
    <iframe name=principal id=principal 
        if Request("id")<>"" then
            src='QuAnuncioDetalle.asp?idioma=@MesgS("Es","En")&tabla=Pisos&id=@Request("id")'
        else 
            src='QuDomoh.asp?idioma=@MesgS("Es","En")'
        end if 
        ></iframe></div>
</body></html>
