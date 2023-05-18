<!-- #include file="IncNuBD.asp" -->
<html>
<head data-hayidioma="<%=Session("HayIdioma")%>" idioma="<%=Session("Idioma")%>" idiomaesen="<%=Session("IdiomaEsEn")%>" req="<%=Request("idioma")%>" nuevo="<%=Request("nuevoidioma")%>">
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
<title>
<%	if Session("Idioma")="" then %>
	Domoh - Alojamiento Gay, Alquiler Gay, Pisos Gay, Habitación Gay
<%  else %>		
	domoh - Vacations, Lodging, Rent, Apartments for gay, gays, lesbian, lesbians flats and rooms to share
<%  end if %>		
</title>
<meta name=description content='Vacaciones, Alquiler, Alquileres para gay, gays, lesbiana, lesbianas de pisos y habitaciones en casas compartidas'>
<meta content="Vacaciones, Alquiler, Alquileres, gay, gays, lesbiana, lesbianas, pisos, habitaciones, casas compartidas, madrid, barcelona" name=keywords>		
<link href=TrDomoh.css rel=stylesheet type=text/css>
<link rel='shortcut icon' href=favicon.ico type=image/icon><link rel=icon href=favicon.ico type=image/icon>
<script type=text/javascript src=efectos.js></script><script type=text/javascript src=forms<%=Session("Idioma")%>.js></script>
<base target=principal/>
</head>
<body>
<a name=top></a>
<div>
	<table border=0 id=container>
		<tr>
			<td></td>
			<td>
				<table border=0 id=padding>
					<tr>
						<td>
							<table border=0>
								<tr>
									<td><a title='Página Principal' href='QuDomoh.asp?idioma=<%=MesgS("Es","En")%>' target=principal><img alt='Domoh' src=images/TrLogo.gif width=141 height=90 /></a></td>
									<td>
                                        <script type=text/javascript>
	                                        banner1 = new Banner('banner1');
<% 
	sQuery="SELECT * FROM Contenidos WHERE pagina='Home" & Session("Idioma") & "' AND posicion>=10"
	rst.Open sQuery, sProvider  
	while not rst.Eof
		if Right(rst("banner"),3) = "oldswf" then 
%>
	                                        banner1.add("FLASH", "<%=rst("banner")%>", 15, 60, 468,"");
<%		elseif Right(rst("banner"),3) = "swf" then %>
                                            banner1.add("IMAGE", "<%=Replace(rst("banner"),".","")&".jpg"%>", 15, 60, 468,"<%=rst("HTML")%>");
<%		elseif rst("banner")<>"" then %>
	                                        banner1.add("IMAGE", "<%=rst("banner")%>", 15, 60, 468,"<%=rst("HTML")%>");
<%		
		end if 
		rst.Movenext
	wend
	rst.Close 
%>
	                                        document.write(banner1);
	                                        banner1.start();
                                        </script></td>
									<td>
										<table border=0>
											<tr><td colspan=2>&nbsp;</td><td></td><td></td></tr>
											<tr>
												<td>&nbsp;</td><td><img alt='Contacto' src=images/TrContact.gif width=18 height=14 /></td><td>&nbsp;</td>
												<td>
													<a title='<%=MesgS("Contactanos","Contact us")%>' href='TrContactanos.asp?idioma=<%=MesgS("Es","En")%>' target=principal class=linkutils>
                                                        <%=MesgS("Contacto","Contact us")%></a></td></tr>
											<tr>
												<td>&nbsp;</td><td><img alt='Conócenos' src=images/TrConocenos.gif width=18 height=14 /></td><td>&nbsp;</td>
												<td>
													<a title='<%=MesgS("Conócenos","About us")%>' href="TrAbout.asp?idioma=<%=MesgS("Es","En")%>" target=principal class=linkutils><%=MesgS("Conócenos","About us")%></a></td>
                                                    </tr>
											<tr>
												<td>&nbsp;</td>
												<td>
													<a title='<%=MesgS("English version","En Español")%>' href="QuMenu.asp?nuevoidioma=<%=MesgS("En","Es")%>">
													<img src='images/Tr<%=MesgS("Eng","Esp")%>.gif' width=18 height=16 border=0 /></a></td>
												<td>&nbsp;</td>
												<td><a title='<%=MesgS("English version","En Español")%>' href="QuMenu.asp?nuevoidioma=<%=MesgS("En","Es")%>" class=linkutils><%=MesgS("English","Español")%></a></td></tr>
											<tr>
												<td>&nbsp;</td><td><img src=images/TrDonar.gif width=18 height=14 /></td><td>&nbsp;</td>
												<td>
													<a title='<%=MesgS("Donar","Donate")%>' href="TrDonar.asp?idioma=<%=MesgS("Es","En")%>&destino=QuDomoh.asp" target=principal class=linkutils>
													<%=MesgS("Donaciones","Donations")%></a></td></tr></table></td></tr></table></td>
						            <td rowspan=4></td></tr>
					            <tr>
						            <td>
							            <table border=0>
								            <tr>
						                        <td>
							                        <a title='<%=MesgS("Entrar en tu zona privada","Enter your private area")%>' target=principal href='http://www.domoh.com/TrDomohLogOn.asp?idioma=<%=MesgS("Es","En")%>'>
							                        <img src='images/TrHomeUsuarios<%=Session("Idioma")%>.gif' width=82 height=34 border=0 id=Image1 
						                        <td><%=MesgS("Mis anuncios","My Ads")%> &gt; </td>
						                        <td>
                                                    <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href='TrVivienda.asp?idioma=<%=MesgS("Es","En")%>'>
                                                    <img src=images/TrAdsVivienda<%=Session("Idioma")%>.gif width=104 height=47 border=0 id=boton_vivienda
                                                        onmouseover="ImgSwap('boton_vivienda','images/TrAdsVivienda<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore()/></a></td>
						                        <td>
                                                    <a title='<%=MesgS("Ver y Publicar Anuncios","Watch and Post Ads")%>' target=principal href="TrVacas.asp?idioma=<%=MesgS("Es","En")%>">
                                                    <img src=images/TrAdsVacaciones<%=Session("Idioma")%>.gif width=104 height=47 border=0 id=boton_vacaciones
                                                        onmouseover="ImgSwap('boton_vacaciones','images/TrAdsVacaciones<%=Session("Idioma")%>_O.gif')" onmouseout=ImgRestore()/></a></td              <td>
                                                    <!-- AddThis Button BEGIN -->
                                                    <div class='addthis_toolbox addthis_default_style'>
                                                    <a href='http://addthis.com/bookmark.php?v=250&amp;username=dmonto' class=addthis_button_compact>Share</a>
                                                    <span class=addthis_separator>|</span><a class=addthis_button_facebook></a><a class=addthis_button_myspace></a><a class=addthis_button_google></a>
                                                    <a class=addthis_button_twitter></a></div>
                                                    <script type=text/javascript src='http://s7.addthis.com/js/250/addthis_widget.js#username=dmonto'></script>
                                                    <!-- AddThis Button END -->
													</td></tr                                         <tr>
                                                <td>
                                                    <iframe width=770 height=4000 name=principal id=principal src=
<%  if Request("id")<>"" then %>
                                                        'QuAnuncioDetalle.asp?idioma=<%=MesgS("Es","En")%>&tabla=Pisos&id=<%=Request("id")%>'
<%  else %>
                                                        'QuDomoh.asp?idioma=<%=MesgS("Es","En")%>' 
<%  end if %>
                                                        />
                                                        </td>
</body></html>
