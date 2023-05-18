<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrUpload.asp" -->
<%
	dim Upload, Fich, sNombreFich

	set Upload = New FileUploader
	Upload.maxFileSize = 150000
	Upload.fileExt = "jpg,gif,bmp,png,jpeg"
	Upload.Upload()

	if Upload.Error then
		Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Upload.ErrorDesc
		Response.Write UMesgS(Upload.Form("idioma"), "Hay un problema con las fotos: ", "Problem uploading pictures: ") & Upload.ErrorDesc
		Response.End
	end if

	if Upload.Files.Count then
		set Fich=Upload.Files(0)
		sNombreFich="'conts/" & Fich.FileName & "'"
		Upload.Save "conts/"
	else
		sNombreFich="null"
	end if

	sQuery="DELETE FROM Contenidos WHERE pagina='" & Upload.Form("pagina") & "' AND posicion=" & Upload.Form("posicion")
	rst.Open sQuery, sProvider

	sQuery= "INSERT INTO Contenidos (pagina, posicion, banner, html, texto) "
	sQuery= sQuery & "VALUES ('" & Upload.Form("pagina") & "', " & Upload.Form("posicion") & ", " & sNombreFich & ", '" & Upload.Form("html") & "', '" & Replace(Upload.Form("texto"),"'","''") & "')"

	rst.Open sQuery, sProvider
	Response.Redirect "NuAdminContenidos.asp"
%>