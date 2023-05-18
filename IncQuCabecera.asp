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
	<title>Domoh - Alojamiento Gay, Alquiler Gay, Pisos Gay, Habitacion Gay - Hebergement Gay</title>
	<meta name=google-site-verification content=TmUeHVRRwXj313YY12Q9Zep315nSOIh5a-wbWheC1VM><meta name=msvalidate.01 content='E42AAFC279BD732454AEBD988B18EDEC'>
    <meta name=viewport content='width=device-width, initial-scale=1'>
	<meta content='text/html;charset=iso-8859-1' http-equiv=Content-Type>
	<meta content='Vacaciones, Alquiler, Alquileres para gay, gays, lesbiana, lesbianas de pisos y habitaciones en casas compartidas' name=description>
	<meta content='Vacaciones, Alquiler, Alquileres, gay, gays, lesbiana, lesbianas, pisos, habitaciones, casas compartidas, madrid, barcelona' name=keywords>
    <meta http-equiv='content-language' content='es-es'>
    <link rel='shortcut icon' href=favicon.ico type=image/icon>
    <script src='http://code.jquery.com/jquery-1.7.1.min.js'></script><script src='http://code.jquery.com/mobile/1.1.0/jquery.mobile-1.1.0.min.js'></script>
    <script>$.get("QuDomoh.asp?head=no&idioma=<%=MesgS("Es","En")%>", function(data) {$( "#principal2" ).html(data);});</script>
    <!-- #include file="IncTrGoogleAn.asp" -->
</head>
<body>
<div class=container>
    <div class=logo>
        <a title='Página Principal' href=indexFram.asp?idioma =<%=MesgS("Es","En")%> target=_self><img alt='Domoh' src=images/TrLogo.gif /></a></div>
    <div class=banner>
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
            banner1.add("IMAGE", "<%=rst("banner")%>", 15, 60, 468,"<%=rst("html")%>");
<%
        end if
        rst.Movenext
    wend
    rst.Close                                            
%>
            // document.write(banner1); banner1.start();
            </script></div>
    <div class=menu>
        <p><img alt='Conócenos' src=images/TrConocenos.gif /><a title='<%=MesgS("Conócenos","About us")%>' href='TrAbout.asp?idioma=<%=MesgS("Es","En")%>&head=si' target=_self class=linkutils><%=MesgS("Conócenos","About us")%></a></p>
        <p><a title='<%=MesgS("English version","En Español")%>' href='indexFram.asp?nuevoidioma=<%=MesgS("En","Es")%>' target=_self><img src='images/Tr<%=MesgS("Eng","Esp")%>.gif'/></a>
        <a title='<%=MesgS("English version","En Español")%>' href='indexFram.asp?nuevoidioma=<%=MesgS("En","Es")%>' target=_self class=linkutils><%=MesgS("English","Español")%></a></p>
        </div>
    <div class=main>
        <a title='<%=MesgS("Entrar en tu zona privada","Enter your private area")%>' target=_self href='http://www.domoh.com/TrDomohLogOn.asp?idioma=<%=MesgS("Es","En")%>&head=si'>
        <img src='images/TrHomeUsuarios<%=Session("Idioma")%>.gif' id=Image1 onmouseover="ImgSwap('Image1','images/TrHomeUsuarios<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a>
        <%=MesgS("Mis anuncios","My Ads")%> &gt; 
        <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=_self href='TrVivienda.asp?idioma=<%=MesgS("Es","En")%>&head=si'>
        <img src='images/TrAdsVivienda<%=Session("Idioma")%>.gif' id=boton_vivienda onmouseover="ImgSwap('boton_vivienda','images/TrAdsVivienda<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a>
        <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=_self href='TrVacas.asp?idioma=<%=MesgS("Es","En")%>&head=si'>
        <img src='images/TrAdsVacaciones<%=Session("Idioma")%>.gif' id=boton_vacaciones onmouseover="ImgSwap('boton_vacaciones','images/TrAdsVacaciones<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a>
    