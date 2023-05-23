<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrJpeg.asp" -->
<!-- #include file="IncTrUpload.asp" -->
<%
	dim Upload, sNombreFich, sBody
	if Session("Usuario") ="" then Response.Redirect "TrLogOn.asp"
		
	set Upload = New FileUploader
	Upload.maxFileSize = 150000
	Upload.fileExt = "jpg,gif,bmp,png,jpeg"
	Upload.Upload()

	if Upload.Error then
		Mail "diego@domoh.com", "Error en " & Request.ServerVariables("SCRIPT_NAME"), Upload.ErrorDesc
		Response.Write UMesgS(Upload.Form("idioma"), "Hay un problema con las fotos: ", "Problem uploading pictures: ") & Upload.ErrorDesc
		Response.End
	end if
			
	sNombreFich=GrabaFoto(Upload, "usuario")

	if Upload.Form("password")="" then
		Response.Redirect Upload.Form("origen") & "?msg=" & UMesgS(Upload.Form("idioma"), "Problemas+Actualizando+Perfil","Problems+Updating+Profile")
	end if
	
 	sQuery = "UPDATE Usuarios SET "	
	sQuery = sQuery & " nombre='" & Replace(Upload.Form("nombre"),"'","''") & "', esagencia='" & Upload.Form("esagencia") & "',"
	sQuery = sQuery & " perfil='" & Upload.Form("perfil") & "', mostrarperfil='" & Upload.Form("mostrarperfil") & "',"
	sQuery = sQuery & " dir1='" & Replace(Upload.Form("dir1"),"'","''") & "', ciudaddir='" & Replace(Upload.Form("ciudaddir"),"'","''") & "',"
	sQuery = sQuery & " dir2='" & Replace(Upload.Form("dir2"),"'","''") & "', cp='" & Upload.Form("cp") & "',"
	sQuery = sQuery & " tel1='" & LimpiaNum(Upload.Form("tel1")) & "', tel2='" & LimpiaNum(Upload.Form("tel2")) & "',"
	sQuery = sQuery & " tel3='" & LimpiaNum(Upload.Form("tel3")) & "', tel4='" & LimpiaNum(Upload.Form("tel4")) & "', password='" & Upload.Form("password") & "',"
	if sNombreFich<>"null" then sQuery = sQuery & " foto=" & sNombreFich & ","
	sQuery = sQuery & " email='" & Upload.Form("email") & "', instrucciones='" & Replace(Upload.Form("instrucciones"),"'","''") & "',"
	sQuery = sQuery & " mostrartel4='" & Upload.Form("mostrartel4") & "', mostraremail='" & Upload.Form("mostraremail") 	& "',"
	sQuery = sQuery & " mostrartel1='" & Upload.Form("mostrartel1") & "', mostrartel2='" & Upload.Form("mostrartel2") & "',"
	sQuery = sQuery & " mostrartel3='" & Upload.Form("mostrartel3") & "' WHERE usuario='" & Session("Usuario") & "'"
	rst.Open sQuery, sProvider

	Response.Redirect Upload.Form("origen") & "?msg=" & UMesgS(Upload.Form("idioma"), "Perfil+Modificado+Correctamente","User+Profile+Updated")
%>