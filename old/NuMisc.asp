<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<title>Anuncios Varios</title>
<% 
	if Request("pais")<>"" then
		Session("Pais")=CLng(Request("pais"))
	elseif Session("Pais")=0 then
		Session("Pais")=34
	end if
%>
<body onload="window.parent.location.hash='top';">
<form name=frm action=NuMiscBusco<%=Session("Idioma")%>.asp method=post>
<table>
	<tr>
		<td rowspan=6 width=30>&nbsp;</td><td colspan=3 width=550>&nbsp;</td>
		<!--#include file="IncNuSubMenu.asp"-->
	<tr>
		<td colspan=3><img alt='Varios' src=images/MiscBanda<%=Session("Idioma")%>.gif width=545></td>
		<td>
			<ul type=circle>
				<li><a title='<%=MesgS("Ir a la pantalla de nuevo anuncio","Go to the New Ad form")%>' href=../NuMiscRegOfrezcoFront.asp>
				<b><u><%=MesgS("Publicar Anuncio","Post Ad")%></u></b></a></li></ul></td></tr>
	<tr>
		<td colspan=3 height=35>
			<p align=center>
				<% Response.Write MesgS("Mira y publica de forma gratuita anuncios de cualquier cosa", "Check out and post ads of anything you want")%>.
				</p></td></tr>
	<tr>
		<td rowspan=3><!-- #include file="IncTrGoogle.asp" --></td>
	<tr>
		<td valign=top height=30><p align=right><b><%=MesgS("Categoría","Category")%>:</b></p></td>
		<td valign=top><% SelectTabla "CatMisc", "Misc", 0, "Nombre" & Session("Idioma") %></td>
		<td height=27 rowspan=2 align=left width=250>
			<button title='<%=MesgS("Ver Anuncios","See the adverts")%>' type=submit onmouseout=ImgRestore() 
				onmouseover="ImgSwap('Buscar','images/BotonBuscar<%=Session("Idioma")%>_O.gif');">
       		<img alt=Buscar name=Buscar src=images/BotonBuscar<%=Session("Idioma")%>.gif width=55></button></td></tr>
	<tr valign=top>
		<td height=30><p align=right><b><%=MesgS("Provincia","Region")%>:</b></td>
		<td>
<% 
	if CLng(Request.Cookies("ProvinciaViv"))=0 then 
		SelectProvincia "Anuncios", 0 
	else
		SelectProvincia "Anuncios", CLng(Request.Cookies("ProvinciaViv"))
	end if
%>
		</td></tr></table></form>
<!-- #include file="IncPie.asp" -->