<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrJpeg.asp" -->
<!-- #include file="IncTrUpload.asp" -->
<%
	dim Upload, sNombreFich, i, Fich, sNombrePreview, Jpeg
	on error goto 0
	
	if Request("borrar")<>"" then
		rst.Open "DELETE FROM Fotos WHERE id=" & Request("borrar"), sProvider
		if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOfertaFoto", Err.Description

		if Session("Id")<>"" then
			rst.Open "SELECT COUNT(*) AS numfotos FROM Fotos WHERE piso=" & Session("Id"), sProvider
			if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOfertaFoto", Err.Description
			if rst("numFotos")=0 then
				rst.Close
				sQuery= "UPDATE Anuncios SET foto='', fechaultimamodificacion=GETDATE() WHERE id=" & Session("Id")
				rst.Open sQuery, sProvider
				if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOfertaFoto", Err.Description
			end if
		end if
		
		Response.Redirect "TrVacasSwapOferta.asp"
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<div class=container>
	<div class=logo>
		<a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a>
		<a href=TrVacas.asp class=linkutils><%=MesgS("Anuncios Vacaciones","Holiday Adverts")%> </a> &gt; <a href=TrVacasSwapOferta.asp class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a>
			&gt; <%=MesgS("Subir Fotos","Upload Pictures")%> &gt; <%=MesgS("Anuncio Grabado","Advert Saved")%></div>
    <h1 class=tituSec><% Response.Write MesgS("Publicación de Apartamento para Intercambio de Vacaciones","Flat Swap Advert Publication")%></h1>
    <div class=row>
<%
	set Upload = New FileUploader
	Upload.maxFileSize = 300000
	Upload.fileExt = "jpg,gif,bmp,png,jpeg"
	Upload.Upload()

	if Upload.Error then
		Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Upload.ErrorDesc
		Response.Write UMesgS(Upload.Form("idioma"), "Hay un problema con las fotos: ", "Problem uploading pictures: ") & Upload.ErrorDesc
		Response.End
	end if

	if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOfertaFoto", Err.Description
	dim sNombreFinal, j
	j=0
	
	for each Fich in Upload.Files()
		rst.Open "SELECT MAX(id)+1 AS nextid FROM Fotos", sProvider
		sNombreFich="fotos/" & rst("nextid") & Fich.Ext
		rst.Close

		if Session("Id")="" then Session("Id")=Upload.Form("id")
		if Session("Id")="" then 
			sQuery="SELECT MAX(id) AS maxid FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOfertaFoto", Err.Description
			Session("Id")=rst("maxid")
			rst.Close
		end if	
	
		if sNombrePreview="" then 
			sNombrePreview=sNombreFich
			sQuery="UPDATE Anuncios SET foto='" & sNombrePreview & "', fechaultimamodificacion=GETDATE() WHERE id=" & Session("Id")
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOfertaFoto", Err.Description
		end if	

		Upload.SaveOne Server.MapPath("/"), j, sNombreFich, sNombreFinal
		set Jpeg = new DigJpeg
		Jpeg.Open sNombreFich
		Jpeg.Height = 100
		Jpeg.Save "mini" & sNombreFich

		sQuery= "INSERT INTO Fotos (piso, foto) VALUES(" & Session("Id") & ", '" & sNombreFich & "')"

		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOfertaFoto", Err.Description
		j=j+1
	next
	
	if Session("Usuario")="hector" then 
		Response.Redirect "NuAdminPisos.asp"
	elseif Session("Activo")="Si" then 
		Response.Redirect "QuHomeUsuario.asp?msg=" & Server.URLEncode(UMesgS(Upload.Form("idioma"), "Vivienda modificada correctamente.", "Advert updated correctly"))
 	end if
%>
        </div>
	<div><!-- #include file="IncTrUEnvioMail.asp" --></div>
	</div>
<!-- #include file="IncUPie.asp" -->
</body>