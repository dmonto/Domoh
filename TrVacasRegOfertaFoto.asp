<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrJpeg.asp" -->
<!-- #include file="IncTrUpload.asp" -->
<%
	dim Upload, sNombreFich, i, Fich, sNombrePreview, Jpeg
	on error goto 0
		
	if Request("borrar")<>"" then
		rst.Open "DELETE FROM Fotos WHERE id=" & Request("borrar"), sProvider
		if Err then Mail "diego@domoh.com", "Error en TrVacasRegOfertaFoto", Err.Description

		if Session("Id")<>"" then
			rst.Open "SELECT COUNT(*) AS numFotos FROM Fotos WHERE piso=" & Session("Id"), sProvider
			if Err then Mail "diego@domoh.com", "Error en TrVacasRegOfertaFoto", Err.Description
			if rst("numFotos")=0 then
				rst.Close
				sQuery="UPDATE Anuncios SET foto='', fechaultimamodificacion=GETDATE() WHERE id=" & Session("Id")
				rst.Open sQuery, sProvider
				if Err then Mail "diego@domoh.com", "Error en TrVacasRegOfertaFoto", Err.Description
			end if
		end if
		
		Response.Redirect "TrVacasRegOferta.asp?op=FotoBorrada"
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<header class=container>
	<div class=logo><a href=QuDomoh.asp title='Página Principal' class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class=topmenu>
		<a href=TrVacas.asp title='Ver Anuncios' class=linkutils><%=MesgS("Anuncios de Vacaciones","Holiday Adverts")%> &gt;</a>
    	<a href=TrVacasRegOfrezcoFront.asp class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a>  
        &gt; <a href=TrVacasRegOferta.asp class=linkutils><%=MesgS("Subir Fotos","Upload Pictures")%></a>&gt; <%=MesgS("Anuncio Grabado","Advert Saved")%></div></header>
<div class=tituSec><% Response.Write MesgS("Publicación de Alojamiento de Vacaciones", "Holiday Lodging Publication")%></div>
<div>
    <p>
<%
	set Upload = new FileUploader
	Upload.maxFileSize = 200000
	Upload.fileExt = "jpg,gif,bmp,png,jpeg"
	Upload.Upload()

	if Upload.Error then
		Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Upload.ErrorDesc
		Response.Write UMesgS(Upload.Form("idioma"), "Hay un problema con las fotos: ", "Problem uploading pictures: ") & Upload.ErrorDesc
		Response.End
	end if

	dim sNombreFinal, j
	j=0
	
	for each Fich in Upload.Files()
		rst.Open "SELECT MAX(id)+1 AS nextid FROM Fotos", sProvider
		sNombreFich="fotos/" & rst("nextid") & Fich.Ext
		rst.Close

		if Session("Id")="" then Session("Id")=Upload.Form("id")
		if Session("Id")="" then 
			Response.Write "Sobrepasaste el Tiempo Máximo. Vuelve a intentarlo por favor."
			Response.End
		end if

		if sNombrePreview="" then 
			sNombrePreview=sNombreFich
			sQuery="UPDATE Anuncios SET foto='" & sNombrePreview & "', fechaultimamodificacion=GETDATE() WHERE id=" & Session("Id")
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrVacasRegOfertaFoto", Err.Description
		end if	
		
		Upload.SaveOne Server.MapPath("/"), j, sNombreFich, sNombreFinal
		set Jpeg = New DigJpeg
		Jpeg.Open sNombreFich
		Jpeg.Height = 100
		Jpeg.Save "mini" & sNombreFich

		sQuery= "INSERT INTO Fotos (piso, foto) VALUES(" & Session("Id") & ", '" & sNombreFich & "')"

		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en TrVacasRegOfertaFoto", Err.Description
		j=j+1
	next
	
	if Session("Usuario")="hector" then 
		Response.Redirect "NuAdminPisos.asp"
	elseif Session("Activo")="Si" then 
		Response.Redirect "QuHomeUsuario.asp?msg=" & Server.URLEncode(UMesgS(Upload.Form("idioma"), "Vivienda modificada correctamente.","Advert Correctly Updated."))
 	end if
%>
        </p></div>
<!-- #include file="IncTrUEnvioMail.asp" -->
<!-- #include file="IncUPie.asp" -->
</body>