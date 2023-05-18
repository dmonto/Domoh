<!-- #include file="IncNuBD.asp" -->
<html>
<head data-hayidioma='<%=Session("HayIdioma")%>' data-idioma='<%=Session("Idioma")%>' data-idiomaesen='<%=Session("IdiomaEsEn")%>' data-req='<%=Request("idioma")%>' data-nuevo='<%=Request("nuevoidioma")%>'>
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
    <title>Domoh - Pisos y Habitaciones Gay-Friendly</title>
    <link rel='shortcut icon' href=favicon.ico type=image/icon><link rel=icon href=favicon.ico type=image/icon>
<% if Request("nuevoidioma")="En" then  %>
    <link rel='alternate' href='http://domoh.com' hreflang='es'/>
<% end if %>
    <script type=text/javascript src=efectos.js></script><script type=text/javascript src=forms<%=Session("Idioma")%>.js></script></script>
    <base target=principal />
    <meta name=viewport content='width=device-width'>
</head>
<body>
<a name=top></a>
<div class=container>
    <div class=logo><a title='Página Principal' href='QuDomoh.asp?idioma=<%=MesgS("Es","En")%>' target=principal><img alt='Domoh' src=images/TrLogo.gif /></a></div>
    <div class=banner>
            <script type=text/javascript>
                banner1 = new Banner('banner1');
<% 
	sQuery="SELECT * FROM Contenidos WHERE pagina='Home" & Session("Idioma") & "' AND posicion>=10"
	rst.Open sQuery, sProvider  
	while not rst.Eof
		if Right(rst("banner"),3) = "swf" then 
%>
                // banner1.add("FLASH", "<%=rst("banner")%>", 15, 60, 468,"");
                banner1.add("IMAGE", "<%=Replace(rst("banner"),".","")&".jpg"%>", 15, 70, 600,"<%=rst("html")%>");
<%		elseif rst("banner")<>"" then %>
	            banner1.add("IMAGE", "<%=rst("banner")%>", 15, 70, 600,"<%=rst("html")%>");
<%		
	    end if 
	    rst.Movenext
	wend
	rst.Close 
%>
	            document.write(banner1); banner1.start();
                </script></div>
    <div class=topmenu>
            <div>
    	        <div class=icono><img class="icono" alt='Contacto' src=images/TrContact.gif /></div>
                <div><a title='<%=MesgS("Contactanos","Contact us")%>' href='TrContactanos.asp?idioma=<%=MesgS("Es","En")%>' target=principal class=linkutils><%=MesgS("Contacto","Contact us")%></a></div></div>
            <div>
    		    <div class=icono><img alt='Conócenos' src=images/TrConocenos.gif /></div>
                <div><a title='<%=MesgS("Conócenos","About us")%>' href='TrAbout.asp?idioma=<%=MesgS("Es","En")%>' target=principal class=linkutils><%=MesgS("Conócenos","About us")%></a></div></div>
        <div>
            <div class=icono><a title='<%=MesgS("English version","En Español")%>' href='QuMenu.asp?nuevoidioma=<%=MesgS("En","Es")%>' target=menu><img src='images/Tr<%=MesgS("Eng","Esp")%>.gif' /></a></div>
            <div><a title='<%=MesgS("English version","En Espa?ol")%>' href='QuMenu.asp?nuevoidioma=<%=MesgS("En","Es")%>' target=menu class=linkutils><%=MesgS("English","Español")%></a></div></div></div></div>
    <div>
	    <div>
		    <a title='<%=MesgS("Entrar en tu zona privada","Enter your private area")%>' target=principal href='http://www.domoh.com/PeDomohLogOn.asp?idioma=<%=MesgS("Es","En")%>'>
			    <img src='images/TrHomeUsuarios<%=Session("Idioma")%>.gif' id=Image1 onmouseover="ImgSwap('Image1','images/TrHomeUsuarios<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a>
            </div>
		<div class="menu"><%=MesgS("Mis anuncios","My Ads")%> &gt; </div>
		<div>
            <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href='TrVivienda.asp?idioma=<%=MesgS("Es","En")%>'>
                <img src='images/TrAdsVivienda<%=Session("Idioma")%>.gif' id=boton_vivienda 
                    onmouseover="ImgSwap('boton_vivienda','images/TrAdsVivienda<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a></div>
   		<div><a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href='TrVacas.asp?idioma=<%=MesgS("Es","En")%>'>
            <img src='images/TrAdsVacaciones<%=Session("Idioma")%>.gif' id=boton_vacaciones 
                onmouseover="ImgSwap('boton_vacaciones','images/TrAdsVacaciones<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore() /></a></div>
        <!-- AddThis Button BEGIN -->
        <div class='addthis_toolbox addthis_default_style'>
            <a href='http://addthis.com/bookmark.php?v=250&amp;username=dmonto' class=addthis_button_compact>Share</a>
            <span class=addthis_separator>|</span><a class=addthis_button_facebook></a><a class=addthis_button_myspace></a><a class=addthis_button_google></a><a class=addthis_button_twitter></a></div>
            <script type=text/javascript src='http://s7.addthis.com/js/250/addthis_widget.js#username=dmonto'></script>
        <!-- AddThis Button END -->
        </div>
    <div>
        <iframe name=principal id=principal src=
<%  if Request("id")<>"" then %>
            'QuAnuncioDetalle.asp?idioma=<%=MesgS("Es","En")%>&tabla=Pisos&id=<%=Request("id")%>'
<%  else %>
            'QuDomoh.asp?idioma=<%=MesgS("Es","En")%>' 
<%  end if %>
            /></div></div>
</body></html>
