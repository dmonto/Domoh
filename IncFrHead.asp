<a name=top />
<div class=container>
    <div class=row>
        <a title='Página Principal' href='indexFram.asp?idioma=<%=MesgS("Es","En")%>' target=_self><img alt='Domoh' src=images/TrLogo.gif /></a>
        <div>
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
                </script>
    <nav>
    <div>
        <div><img alt='Conócenos' src=images/TrConocenos.gif /></div>
        <div><a title='<%=MesgS("Conócenos","About us")%>' href='TrAbout.asp?idioma=<%=MesgS("Es","En")%>&head=si' target=_self class=linkutils><%=MesgS("Conócenos","About us")%></a></div></div>
    <div>
        <div><a title='<%=MesgS("English version","En Español")%>' href='indexFram.asp?nuevoidioma=<%=MesgS("En","Es")%>' target=_self><img src='images/Tr<%=MesgS("Eng","Esp")%>.gif' /></a></div>
        <div><a title='<%=MesgS("English version","En Español")%>' href='indexFram.asp?nuevoidioma=<%=MesgS("En","Es")%>' target=_self class=linkutils><%=MesgS("English","Español")%></a></div></div>
    <div>
        <div>
            <a title='<%=MesgS("Entrar en tu zona privada","Enter your private area")%>' target=_self href='http://www.domoh.com/TrDomohLogOn.asp?idioma=<%=MesgS("Es","En")%>&head=si'>
            <img src='images/TrHomeUsuarios<%=Session("Idioma")%>.gif' id=Image1 onmouseover="ImgSwap('Image1','images/TrHomeUsuarios<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a></div>
        <div><%=MesgS("Mis anuncios","My Ads")%> &gt; </div>
        <div>
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
            </div>
    </div></div>
    