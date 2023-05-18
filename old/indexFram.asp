<!-- #include file="IncNuBD.asp" -->
<%
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
%>
<html>
<head>
	<title>Domoh - Alojamiento Gay, Alquiler Gay, Pisos Gay, Habitación Gay - Hébergement Gay</title>
	<meta name=google-site-verification content=TmUeHVRRwXj313YY12Q9Zep315nSOIh5a-wbWheC1VM /><meta name=msvalidate.01 content='E42AAFC279BD732454AEBD988B18EDEC'/>
    <meta name=viewport content='width=device-width, initial-scale=1'/>
	<meta content='text/html;charset=iso-8859-1' http-equiv=Content-Type />
	<meta content='Vacaciones, Alquiler, Alquileres para gay, gays, lesbiana, lesbianas de pisos y habitaciones en casas compartidas' name=description />
	<meta content='Vacaciones, Alquiler, Alquileres, gay, gays, lesbiana, lesbianas, pisos, habitaciones, casas compartidas, madrid, barcelona' name=keywords />
    <meta http-equiv='content-language' content='es-es'/>
    <link rel='shortcut icon' href=favicon.ico type=image/icon />
	<script type=text/javascript src=efectos.js></script>
    <script src='http://code.jquery.com/jquery-1.7.1.min.js'></script><script src='http://code.jquery.com/mobile/1.1.0/jquery.mobile-1.1.0.min.js'></script>
    <script>$.get("QuDomoh.asp?head=no&idioma=<%=MesgS("Es","En")%>", function(data) {$( "#principal2" ).html(data);});</script>
    <!-- #include file="IncTrGoogleAn.asp" -->
</head>
<body>
<a name=top />
<header class=container>
    <div class=row>
        <div class=col-2><a title='Página Principal' href='indexFram.asp?idioma=<%=MesgS("Es","En")%>' target=_self><img alt='Domoh' src=images/TrLogo.gif /></a></div>
        <div class=col-5>
            <script type=text/javascript>
                banner1 = new Banner('banner1');
<%
    sQuery="SELECT * FROM Contenidos WHERE pagina='Home" & Session("Idioma") & "' AND posicion>=10"
    rst.Open sQuery, sProvider
                                         
    while not rst.Eof
        if Right(rst("banner"),3) = "swf" then
%>
                banner1.add("IMAGE", "<%=Replace(rst("banner"),".","")&".jpg"%>", 15, 60, 468,"<%=rst("html")%>");
<%		elseif rst("banner")<>"" then %>
                banner1.add("IMAGE", "<%=rst("Banner")%>", 15, 60, 468,"<%=rst("html")%>");
<%
        end if
        rst.Movenext
    wend
    rst.Close                                            
%>
                // document.write(banner1); banner1.start();
                </script></div>
        <div class=col-4>
            <div class=row>
                <img alt='Contacto' src=images/TrContact.gif class=icono />
                <a title='<%=MesgS("Contactanos","Contact us")%>' href='TrContactanos.asp?idioma=<%=MesgS("Es","En")%>' target=_self class=linkutils><%=MesgS("Contacto","Contact us")%></a></div>
            <div class=row>
                <img alt='Conócenos' src=images/TrConocenos.gif class=icono />
                <a title='<%=MesgS("Conócenos","About us")%>' href='TrAbout.asp?idioma=<%=MesgS("Es","En")%>&head=si' target=_self class=linkutils><%=MesgS("Conócenos","About us")%></a></div>
            <div class=row>
                <a title='<%=MesgS("English version","En Español")%>' href='indexFram.asp?nuevoidioma=<%=MesgS("En","Es")%>' target=_self><img src='images/Tr<%=MesgS("Eng","Esp")%>.gif' class=icono /></a>
                <a title='<%=MesgS("English version","En Español")%>' href='indexFram.asp?nuevoidioma=<%=MesgS("En","Es")%>' target=_self class=linkutils><%=MesgS("English","Español")%></a></div>
            <div class=row>
                <div><img src=images/TrDonar.gif class=icono /></div>
                <div><a title='<%=MesgS("Donar","Donate")%>' href='TrDonar.asp?idioma=<%=MesgS("Es","En")%>&head=si&destino=QuDomoh.asp' target=_self class=linkutils><%=MesgS("Donaciones","Donations")%></a></div></div></div></div>
    <div class=row>
        <a title='<%=MesgS("Entrar en tu zona privada","Enter your private area")%>' target=_self href='http://www.domoh.com/TrDomohLogOn.asp?idioma=<%=MesgS("Es","En")%>&head=si'>
            <img src='images/TrHomeUsuarios<%=Session("Idioma")%>.gif' id=Image1 onmouseover="ImgSwap('Image1','images/TrHomeUsuarios<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a>
        <%=MesgS("Mis anuncios","My Ads")%> &gt; </div>
    <div class=row>
        <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=_self href='TrVivienda.asp?idioma=<%=MesgS("Es","En")%>&head=si'>
        <img src='images/TrAdsVivienda<%=Session("Idioma")%>.gif' id=boton_vivienda onmouseover="ImgSwap('boton_vivienda','images/TrAdsVivienda<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a></div>
    <div>
        <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=_self href='TrVacas.asp?idioma=<%=MesgS("Es","En")%>&head=si'>
        <img src='images/TrAdsVacaciones<%=Session("Idioma")%>.gif' id=boton_vacaciones onmouseover="ImgSwap('boton_vacaciones','images/TrAdsVacaciones<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a></div>
    <div>
        <!-- AddThis Button BEGIN -->
        <div class='addthis_toolbox addthis_default_style'>
            <a href='http://addthis.com/bookmark.php?v=250&amp;username=dmonto' class=addthis_button_compact>Share</a><span class=addthis_separator>|</span>
            <a class=addthis_button_facebook></a><a class=addthis_button_myspace></a><a class=addthis_button_google></a><a class=addthis_button_twitter></a></div>
        <script type=text/javascript src='http://s7.addthis.com/js/250/addthis_widget.js#username=dmonto'></script>
        <!-- AddThis Button END -->
        </div></header>
    <div id=principal2>&nbsp;</div>
</body></html>