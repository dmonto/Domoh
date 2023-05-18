<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<%
	dim sBody
	on error goto 0
%>
<!-- #include file="IncTrCabecera.asp" -->
<title>Test Upload</title>
<body onload="window.parent.location.hash='top';
<%	if instr(request("email"),"hotmail")>0 then %>
	detalles('images/AntiSpam.gif');
<%	end if %>
"> 
<form name=frm action=TestUpload.asp method=post enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Session("Id")%>>
<div align=center>
<table width=770 height=100% border=0 cellpadding=0 cellspacing=0 bgcolor=#FFFFFF id=container>
	<tr>
		<td align=center valign=top>
			<table width=728 border=0 cellspacing=0 cellpadding=0>
				<tr>
					<td width=567 valign=top>
						<table width='100%' border=0 cellspacing=0 cellpadding=0>
				            <tr>
                                <td width=285 align=left valign=top><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></td>
					            <td align=left><a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios de Vivienda","Property Adverts")%> &gt;</a><b>&gt; <%=MesgS("Subir Fotos","Upload Pictures")%></b></td></tr>
				            <tr><td colspan=2 align=left class=tituSec>&nbsp;</td></tr>
				<tr><td colspan=2 align=left class=tituSec><%=MesgS("Publicación de Piso o Habitación para Vivir","Post New Advert for Rent")%></td></tr>
                <!-- #include file="IncTrUFotos.asp" -->
</table></form>
<!-- #include file="IncPie.asp" -->















