<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<!-- #include file="IncTrJpeg.asp" -->
<!-- #include file="IncTrUpload.asp" -->
<%
	dim Upload, Fich, sNombreFich, sNombreFlyer, sBody, Jpeg, JpLogo, i, j, sNombreFinal
			
	if Request("op")<>"Test" then
        set Upload = new FileUploader
	    Upload.maxFileSize = 300000
	    Upload.fileExt = "jpg,gif,bmp,png,jpeg"
	    Upload.Upload()

	    if Upload.Error then
		    Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), "Upload: " & Upload.ErrorDesc & " - " & Err.Description
		    Response.Write UMesgS(Upload.Form("idioma"), "Hay un problema con las fotos: ", "Problem uploading pictures: ") & Upload.ErrorDesc
		    Response.End
	    end if
			
	    i=0
	    j=-1
	    for each Fich in Upload.Files()
		    if Upload.Files.Key(j)="foto" then j=i
		    i=i+1
	    next

	    if j=-1 then
		    sNombreFlyer= "NULL"
	    else
		    set Fich=Upload.Files("flyer")
		    sNombreFich="flyer" & Month(Now) & Day(Now) & Hour(Now) & Minute(now) & fich.Ext
		    Upload.SaveOne Server.MapPath("/"), j, sNombreFlyer, sNombreFinal
		    sNombreFlyer= "'conts/" & sNombreFinal & "'"
	    end if

	    sNombreFich=GrabaFoto(Upload, "anuncio" & Month(Now) & Day(Now) & Hour(Now) & Minute(now))

	    if Session("Usuario")="" then
		    '--- Usuario Nuevo 
		    if Request("password")="" then 
			    Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), Request.Form("usuario")
                Response.Write "Algo fue mal"
			    Response.End
		    end if

		    Session("TipoUsuario")="Anunciante"
		    Session("Activo")="Mail"
		    InsertUsuario Upload
		    PreBienvenida "TrMiscRegOferta" & Session("Idioma") & ".asp"
	    end if	

	    if Upload.Form("id")="" or Upload.Form("id")="nuevo" then
		    '---- Anuncio Nuevo 
	 	    sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, foto, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
		    sQuery = sQuery & "VALUES ('Misc','Si','" & Session("Usuario") & "','" & Replace(Upload.Form("descripcion"),"'","''") & "','" & Replace(Upload.Form("anuncio"),"'","''") & "'," 
            sQuery = sQuery & Upload.Form("provincia") & "," & sNombreFich & ",'" & Upload.Form("idiomaes") & "', '" & Upload.Form("idiomaen") & "', GETDATE(),GETDATE())"

		    Err.Clear
		    rst.Open sQuery, sProvider
		    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
		
		    sQuery="SELECT MAX(id) AS maxid FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
		    rst.Open sQuery, sProvider
		    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Err.Description
		    Session("Id")=rst("maxId")
		    rst.Close

	 	    sQuery = "INSERT INTO Misc (id, flyer) VALUES (" & Session("Id") & ", " & sNombreFlyer & ")"

		    Err.Clear
		    if rst.State then rst.Close
		    rst.Open sQuery, sProvider
		    if Err then 
			    Mail "diego@domoh.com", "Error en NuMiscRegOferta", sQuery & " - " & Err.Description
		    end if

		    sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Anuncio introducido por " & Session("Usuario") & MiscDetalles("", Session("Id"))
		    Mail "hector@domoh.com", "Nuevo Anuncio", sBody
		    NuevoAnuncio
		    if Upload.Form("id")="nuevo" then Response.Redirect "QuHomeUsuario.asp?msg=" & UMesgS(Upload.Form("idioma"), "Anuncio+Añadido.","Advert+Added.")
	    else
		    '---- Modificar Anuncio
	 	    sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Upload.Form("descripcion"),"'","''") & "', descripcion='" & Replace(Upload.Form("anuncio"),"'","''") & "', provincia=" & Upload.Form("provincia") & ","
		    if sNombreFich<>"null" then sQuery = sQuery & "foto=" & sNombreFich & ","
		    sQuery = sQuery & " idiomaes='" & Upload.Form("idiomaes") & "', idiomaen='" & Upload.Form("idiomaen") & "', fechaultimamodificacion=GETDATE() WHERE id=" & Upload.Form("id") 

		    Err.Clear
		    rst.Open sQuery, sProvider
		    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

		    if sNombreFlyer<>"null" then 
		 	    sQuery = "UPDATE Misc SET flyer=" & sNombreFlyer & " WHERE id=" & Upload.Form("id") 

			    Err.Clear
			    rst.Open sQuery, sProvider
			    if Err then 
				    Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
			    end if
		    end if
		    sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Anuncio modificado por " & Session("Usuario") & MiscDetalles("", Session("Id"))
		    Mail "hector@domoh.com", "Anuncio Modificado", sBody
		    Response.Redirect "QuHomeUsuario.asp?msg=" & UMesgS(Upload.Form("idioma"), "Anuncio+Modificado.","Advert+Updated.")
	    end if
    end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<div id=container>
	<div><a href=QuDomoh.asp class=linkutils>&lt;&lt; <% Response.Write UMesgS(Upload.Form("idioma"), "Volver a Home ","Back to Home Page ")%></a></div><div class=tituSec>&nbsp;</div>
	<div class=tituSec><% Response.Write UMesgS(Upload.Form("idioma"), "Publicación de Anuncio","Advert Publication")%></div>
    <!-- #include file="IncTrEnvioMail.asp" -->
</div>
<!-- #include file="IncUPie.asp" -->
</body>
