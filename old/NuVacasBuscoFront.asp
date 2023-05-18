<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<script language=JavaScript>
function ClickPais() {
	document.location="NuVacasBuscoFront<%=Session("Idioma")%>.asp?Pais="+frm.pais.value;
   	return true;
}
</script>
<% 
	dim sWhere
	
	if Request("pais")<>"" then
		Session("Pais")=CLng(Request("pais"))
	elseif Session("Pais")=0 then
		Session("Pais")=34
	end if
%>
<body onload="window.parent.location.hash='top';"> 
<span id=bolatexto></span>
<table width=100%>
	<tr><td colspan=2 width=500 align=right></td><!--#include file="IncNuSubMenu.asp"--></tr>
	<tr>
		<td colspan=2 align=center>
			Mira de forma gratuita los hoteles y apartamentos para
			vacaciones que están disponibles.</td>
		<td rowspan=2>
			<ul type=circle>
				<li><a title="Ir a Nuevo Anuncio para Vacaciones" href=../TrVacasRegOfrezcoFront.asp>
					<b>Publicar Hotel o Apartamento</b></a></li>
				<li><a title="Busca Apartamentos disponibles para Intercambio" href=../NuVacasSwap.asp>
					<b>Intercambio de Apartamentos</b></a></li></ul></td></tr>
	<tr><td colspan=2 align=center><b>País:</b><% SelectPais "", "Vacas", Session("Pais")%></td></tr>
	<tr>
<%		
	if Session("Pais")<>0 then
		sWhere = "WHERE pais=" & Session("Pais") 
		sWhere = sWhere & " AND id IN "
		sWhere = sWhere & " (SELECT provincia from Anuncios WHERE activo='Si' AND tabla='Pisos') "
		rst.Open "SELECT * FROM Provincias " & sWhere & " ORDER BY nombre", sProvider
	end if

	if Session("Pais")=34 then
%>
	<tr>
   		<td bordercolor=#800080 colspan=2>
			<img title="En esta zona aún no tenemos alojamientos!" src=/images/MapaCiudades.gif 
				width=500 align=middle usemap="#mapaEspana">
			<map name=mapaEspana>
<%
		while not rst.Eof
%>
				<area onmouseout="document.getElementById('bolatexto').innerText='';" 
					onmouseover="document.getElementById('bolatexto').innerText='<%=rst("Nombre")%>';" 
					shape=poly coords="<%=rst("Poligono")%>" 
					href="TrVacasBusco.asp?provincia=<%=rst("Id")%>" ALT="<%=rst("Nombre")%>">
<%
			rst.Movenext
		wend
		rst.Close
%>	
			</map></td></tr>
	<tr>
		<td align=left valign=top>
<%
	else
		rst.Close
		end if
%>
<form name=frm action=NuVacasBusco<%=Session("Idioma")%>.asp method=post>
  		<tr>
			<td align=center>
              <p valign=top nobr><b>Provincia:&nbsp;</b>
				<% SelectProvincia "Vacas", 0 %></p>
        		<button title="Ver alojamientos" type=submit onmouseout="ImgRestore()" 
        			onmouseover="ImgSwap('Buscar','images/BotonBuscar_O.gif')" >
        		<img name=Buscar src=../images/BotonBuscar.gif width=70></button></td>
        </tr>
</form>
</table>
<!-- #include file="IncPie.asp" -->

































