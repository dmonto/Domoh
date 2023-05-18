<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrJpeg.asp" -->
<!-- #include file="freeASPUpload.asp" -->
<%
	dim numBytes, Upload, sNombreFich, i, Fich, sNombrePreview, Jpeg
	on error goto 0		
%>
<script type=text/javascript>function loca() { alert(window.parent.location); }</script>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';
<%	if instr(request("email"),"hotmail")>0 then %>
	detalles('images/AntiSpam.gif');
<%	end if %>
"> 
<a name=top />
<div class=container>
	<div class=logo>
		<div class=left><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
		<nav>
			<a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios de Vivienda","Property Adverts")%> &gt;</a>
			<a href=TrCasaRegOfrezcoFront.asp class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a>  &gt; 
			<a href=old/TestUploadFront.asp class=linkutils><%=MesgS("Subir Fotos","Upload Pictures")%></a> &gt; <%=MesgS("Anuncio Grabado","Advert Saved")%></nav></div>
	<div class=tituSec>&nbsp;</div>
	<div class=tituSec><%=MesgS("Publicación de Piso o Habitación para Vivir","Advert Publication for Rent")%></div>
    <div>
        <p>o bien pulsa <a title='Volver a la pantalla anterior' href="javascript:document.getElementById('principal').contentWindow.location.replace(url);}">aquí</a> para utilizar otro nombre de usuario.</p>
        <a href="javascript:alert(document.getElementById('menu').contentWindow.location);">tst</a><a href='javascript:loca();'>tst2</a></div>
    <div>
<%
	set Upload = new FreeASPUpload
	'Upload.maxFileSize = 150000
	'Upload.fileExt = "jpg,gif,bmp,png,jpeg"
	Upload.Upload()

	if Err then
		Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Upload.ErrorDesc
		Response.Write MesgS("Hay un problema con las fotos: ", "Problem uploading pictures: ") & Upload.ErrorDesc
		Response.End
	end if

	rst.Open "SELECT max(id)+1 AS nextid FROM Fotos", sProvider
	i=rst("nextid")
    rst.Close

    dim sNombreFinal, j
    j=0
	for each Fich in Upload.Files()
		sNombreFich="fotos/" & i & Fich.Ext
		i=i+1
		
		Upload.SaveOne Server.MapPath("/"), j, sNombreFich, sNombreFinal
	'	fich.SavetoDisk Server.MapPath("fotos/")
		set Jpeg = new DigJpeg
		Jpeg.Open sNombreFich
		Jpeg.Height = 100
		Jpeg.Save "mini" & sNombreFich
		j=j+1
	next
%>
</div></div>
<!-- #include file="IncPie.asp" -->
</body>
