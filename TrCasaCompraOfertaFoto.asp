<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrJpeg.asp" -->
<!-- #include file="IncTrUpload.asp" -->
<%
	dim Upload, sNombreFich, i, Fich, sNombrePreview, Jpeg
	
	if Request("borrar")<>"" then
		sQuery="DELETE FROM Fotos WHERE id=" & Request("borrar")
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

		if Session("Id")<>"" then
			sQuery="SELECT COUNT(*) AS NumFotos FROM Fotos WHERE piso=" & Session("Id")
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
			if rst("numFotos")=0 then
				rst.Close
				sQuery="UPDATE Anuncios SET foto='', fechaultimamodificacion=GETDATE() WHERE id=" & Session("Id")
				rst.Open sQuery, sProvider
				if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
			end if
		end if
		
		Response.Redirect "TrCasaCompraOferta.asp?op=FotoBorrada"
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<div class=container>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class=topmenu>
		<a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios de Vivienda","Property Adverts")%> &gt;</a><a href=TrCasaCompraOfrezcoFront.asp class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a>  
		&gt; <%=MesgS("Subir Fotos","Upload Pictures")%> &gt; <%=MesgS("Anuncio Grabado","Advert Saved")%></div>
    <div class=banner><% Response.Write MesgS("Publicación de Piso o Casa para Vender", "Advert Publication - For Sale")%></div>
    <div>
	    <p>
<%
	set Upload = New FileUploader
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
	
		if sNombrePreview="" then 
			sNombrePreview=sNombreFich
			sQuery="UPDATE Anuncios SET foto='" & sNombrePreview & "', fechaultimamodificacion=GETDATE() WHERE id=" & Session("Id")
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
		end if	
		
		Upload.SaveOne Server.MapPath("/"), j, sNombreFich, sNombreFinal
		set Jpeg = new DigJpeg
		Jpeg.Open sNombreFich
		Jpeg.Height = 100
		Jpeg.Save "mini" & sNombreFich

		sQuery= "INSERT INTO Fotos (piso, foto) VALUES (" & Session("Id") & ", '" & sNombreFich & "')"

		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
		j=j+1
	next
	
	if Session("Activo")="Si" then 
		Response.Redirect "QuHomeUsuario.asp?msg=" & Server.URLEncode(UMesgS(Upload.Form("idioma"), "Vivienda modificada correctamente.","Advert correctly updated."))
	end if
%>
<!-- #include file="IncTrUEnvioMail.asp" -->
</p></div></div>
<!-- #include file="IncUPie.asp" -->
</body>
