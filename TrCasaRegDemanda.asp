<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<!-- #include file="IncTrJpeg.asp" -->
<!-- #include file="IncTrUpload.asp" -->
<%
	dim Upload, sNombreFich, sBody, numMaximo, sIdi
				
	if Request("op")="NuevoUsuario" then
		sQuery = "SELECT * FROM Usuarios WHERE usuario='" & Request.Form("usuario") & "'"
		rst.Open sQuery, sProvider
		if Not(rst.Eof) then 
			Response.Redirect "TrUsuarioYaExiste.asp?origen=TrCasaRegDemanda.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto") & "&idiomaen=" & Request.Form("idiomaen")
		end if
		rst.Close
		sQuery = "UPDATE Usuarios SET usuario='" & Request.Form("usuario") & "' WHERE usuario='" & Session("usuario") & "'"
		rst.Open sQuery, sProvider
		sQuery = "UPDATE Anuncios SET usuario='" & Request.Form("usuario") & "' WHERE usuario='" & Session("usuario") & "'"
		rst.Open sQuery, sProvider
		Session("Usuario")=Request.Form("usuario")
		PreBienvenida "TrCasaRegDemanda.asp"
	elseif Request("op")<>"Test" then
		set Upload = New FileUploader
		Upload.maxFileSize = 300000
		Upload.fileExt = "jpg,gif,bmp,png,jpeg"
		Upload.Upload()

		if Upload.Error then
			Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Upload.ErrorDesc
			Response.Write UMesgS(Upload.Form("idioma"), "Hay un problema con las fotos: ", "Problem uploading pictures: ") & Upload.ErrorDesc
			Response.End
		end if
			
		numMaximo=0
		numMaximo=CLng(Upload.Form("maximo"))
		Err.Clear
	
		if Session("Usuario")="" then
            '--- Usuario Nuevo 
			if Upload.Form("password")="" then 
				Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), Request.Form("usuario")
                Response.Write "Algo fue mal"
				Response.End
			end if

			Session("TipoUsuario")="Inquilino"
			InsertUsuario Upload
			if Session("Usuario")=Upload.Form("usuario") then PreBienvenida "TrCasaRegDemanda.asp"
		end if	

		sNombreFich=GrabaFoto(Upload, "inquilino")
	    if sNombreFich<>"null" then
	    	sQuery="UPDATE Usuarios SET foto=" & sNombreFich & " WHERE usuario='" & Session("Usuario") & "'"
		    Err.Clear
		    rst.Open sQuery, sProvider
		    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
	    end if

	    if Upload.Form("id")="" or Upload.Form("id")="nuevo" then
		    '--- Inquilino Nuevo 
		    if Upload.Form("cabecera")="" or Session("usuario")="" then Response.Redirect "TrCasaRegDemandaFront.asp"
		
	 	    sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
		    sQuery = sQuery & "VALUES ('Inquilinos','Si','" & Session("Usuario") & "','" & Replace(Upload.Form("cabecera"),"'","''") & "',"
		    sQuery = sQuery & "'" & Replace(Upload.Form("anuncio"),"'","''") & "',0,'" & Upload.Form("idiomaes") & "', '" & Upload.Form("idiomaen") & "', GETDATE(), GETDATE())"

		    Err.Clear
		    rst.Open sQuery, sProvider
		    if Err then 
    			Mail "diego@domoh.com", "Error en TrCasaRegDemanda", sQuery & " - " & Err.Description
	    		Response.Redirect "TrCasaRegDemandaFront.asp"
		    end if
				
    		sQuery="SELECT MAX(id) AS maxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
	    	rst.Open sQuery, sProvider
		    if Err then Mail "diego@domoh.com", "Error en TrCasaRegDemanda", sQuery & " - " & Err.Description
    		Session("Id")=rst("maxId")
	    	rst.Close

	 	    sQuery = "INSERT INTO Inquilinos (id, tipoviv, ciudad, maximo, fuma, mascota) "
		    sQuery = sQuery & "VALUES (" & Session("Id") & ",'"& Upload.Form("tipoviv") & "','" & Upload.Form("ciudad") & "'," & numMaximo & ",'" & Upload.Form("fuma") & "','"& Upload.Form("mascota") & "')"

		    Err.Clear
		    rst.Open sQuery, sProvider
		    if Err then 
    			Mail "diego@domoh.com", "Error en TrCasaRegDemanda", sQuery & " - " & Err.Description
			    Response.Redirect "TrCasaRegDemandaFront.asp"
		    end if

		    sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Anuncio de búsqueda de piso para alquilar introducido por " & Session("Usuario") & InquilinoDetalles("", Session("Id"))
		    Mail "hector@domoh.com", "Nuevo Inquilino", sBody
		    NuevoAnuncio
		    if Upload.Form("id")="nuevo" then 
    			Response.Redirect "QuHomeUsuario.asp?msg=" & UMesgS(Upload.Form("idioma"), "Piso+Añadido.","Advert+Added.")
	    	elseif Session("Usuario")<>Upload.Form("usuario") then 
		    	Response.Redirect "TrUsuarioYaExiste.asp?origen=TrCasaRegDemanda.asp&usuario=" & Upload.Form("usuario") & "&foto=" & Upload.Form("foto")& "&idiomaen=" & Upload.Form("idiomaen")
		    end if
	    else
		    '---- Modificar Inquilino
	 	    sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Upload.Form("cabecera"),"'","''") & "', descripcion='" & Replace(Upload.Form("anuncio"),"'","''") & "',"
		    sQuery = sQuery & " idiomaes='" & Upload.Form("idiomaes") & "', idiomaen='" & Upload.Form("idiomaen") & "', fechaultimamodificacion=GETDATE() WHERE id=" & Upload.Form("id") 
    
		    rst.Close
		    Err.Clear
		    rst.Open sQuery, sProvider
		    if Err then Mail "diego@domoh.com", "Error en TrCasaRegDemanda", sQuery & " - " & Err.Description

    	 	sQuery = "UPDATE Inquilinos SET tipoviv='" & Upload.Form("tipoviv") & "', ciudad='" & Replace(Upload.Form("ciudad"),"'","''") & "',"
	    	sQuery = sQuery & " maximo=" & numMaximo & ", fuma='" & Upload.Form("fuma") & "', mascota='" & Upload.Form("mascota") & "' WHERE id=" & Upload.Form("id") 

		    if rst.State then rst.Close
		    Err.Clear
		    rst.Open sQuery, sProvider
		    if Err then 
			    Mail "diego@domoh.com", "Error en TrCasaRegDemanda", sQuery & " - " & Err.Description
			    Response.Redirect "TrCasaRegDemandaFront.asp"
		    end if
		    sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Anuncio de inquilino modificado por " & Session("Usuario") & InquilinoDetalles("", Session("id"))
		    Mail "hector@domoh.com", "Anuncio de Inquilino Modificado", sBody
		    Response.Redirect "QuHomeUsuario.asp?msg=" & Server.URLEncode(UMesgS(Upload.Form("idioma"), "Anuncio Modificado","Advert Updated"))
	    end if
	end if
    sIdi=Upload.Form("idioma")
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';"> 
<div class=container>
    <div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <% Response.Write UMesgS(sIdi, "Volver a Home","Back to Home Page")%> </a></div>
		<div class=right>
		    <a href=TrVivienda.asp class=linkutils><% Response.Write UMesgS(sIdi, "Anuncios de Vivienda","Property Adverts")%> &gt;</a>
			<a href=TrCasaRegDemandaFront.asp class=linkutils><%=UMesgS(sIdi, "Nuevo Anuncio","New Advert")%></a>
            &gt; <%=UMesgS(sIdi,"Subir Fotos","Upload Pictures")%></div></div>
    <div class=tituSec><%=UMesgS(sIdi, "Publicación de Anuncio de Búsqueda para Alquilar","Flat or Room Request for Rent Advert Publication")%></div>
    <!-- #include file="IncTrEnvioMail.asp" -->
    </div>
<!-- #include file="IncUPie.asp" -->
</body>