<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<title>Property for Sale</title>
<script type=text/javascript>function ClickPais() {document.location="NuCasaCompraEn.asp?Pais="+frm.pais.value; return true;}</script>
<% 
	if Request("pais")<>"" then
		Session("Pais")=CLng(Request("pais"))
	elseif Session("Pais")=0 then
		Session("Pais")=34
	end if
%>
<body onload="window.parent.location.hash='top';"> 
<form name=frm action=old/NuCasaCompraBuscoEn.asp method=post>
<table width=100%>
	<tr>
		<td rowspan=6 width=30></td>
		<td colspan=3>
			<a title='Inicio de Mi Casa' href=old/NuCasaEn.asp><img src=images/CasaBandaEn.gif width=500></a></td>
<!--#include file="IncNuSubMenu.asp"-->
			<ul type=circle>
				<li><b><a title='New Ad for Selling your property' href=TrCasaCompraOfrezcoFront.asp>Post House or Flat</a></b></li>
				<li><b><a title='New Ad for requesting property' href=TrCasaCompraDemandaFront.asp>Cannot find what you are looking for</a></b></li>
				<li><b><a title='Go to the Renting Area' href=old/NuCasaBuscoFrontEn.asp>Flats and Rooms to Rent</a></b></li></ul></tr>
	<tr><td colspan=3 height=40><p align=center>Look at our selection of flats and rooms on sale.</p></td></tr>
	<tr>
		<td valign=top align=right height=28><b>Country:</b><% SelectPais "", "Compra", Session("Pais") %></td>
		<td rowspan=3 align=left height=60>
			<button title='Watch Results' type=submit onmouseout=ImgRestore() onmouseover=ImgSwap('Buscar','images/BuscarEn_O.gif') >
				<img alt='Buscar' name=Buscar src=images/BuscarEn.gif width=75 height=77></button></td></tr>
	<tr><td valign=top align=right height=28><b>Region:</b><% SelectProvincia "Compra", 0 %></td></tr>
	<tr>
		<td bordercolor=#800080 align=right height=40>    
			<b>Maximum price </b> you are willing to pay <input title='Price in thousand Euros' size=5 name=precio><font face=arial> thousand €</font></td></tr>
</table></form>
<!-- #include file="IncPie.asp" -->








