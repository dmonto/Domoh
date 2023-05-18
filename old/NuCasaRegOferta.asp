<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncQuDetalles.asp" -->
<%
	dim sBody, sSubject
		
	if Request("op")<>"FotoBorrada" and Request("op")<>"test" then
		if Session("Usuario")="" then
			'--- Usuario Nuevo 
			if Request("nombre")="" then Response.Redirect "NuCasaRegOfrezcoFront" & Session("Idioma") & ".asp"
    	    
    	    if Request("password")="" then 
	    	    Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), " Usuario " & Session("Usuario")
                Response.Write "Algo fue mal"
	    	    Response.End
    	    end if

			Session("TipoUsuario")="Casero"
			Session("Activo")="Mail"
			InsertUsuario Request
			PreBienvenida Request.Servervariables("SCRIPT_NAME")
		end if	

		if Request("id")="" or Request("id")="nuevo" then
			'--- Piso Nuevo 		
			if Request("tipo")="" or Request("ciudadnombre")="" or Session("Usuario")="" or Session("EMail")="" then 
				Response.Redirect "TrCasaRegOfrezcoFront.asp"
			end if
	
		 	sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
			sQuery = sQuery & "VALUES ('Pisos','Si','" & Session("Usuario") & "','" & Replace(Request("ciudadnombre"),"'","''") 
			sQuery = sQuery & " (" & Replace(Request("zona"),"'","''") & ")','" & Replace(Request("descripcion"),"'","''") & "',0,'" & Request("idiomaes") & "', '" & Request("idiomaen") & "', GETDATE(), GETDATE())"

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then 
				Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
				Response.Redirect "NuCasaRegOfrezcoFront.asp"
			end if
				
			sQuery="SELECT MAX(id) AS maxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
			Session("Id")=rst("maxId")
			rst.Close

		 	sQuery = "INSERT INTO Pisos (id, tipo, dir1, ciudadnombre, dir2, cp, zona, rentaviv, gente, fuma, mascota) "
			sQuery = sQuery & "VALUES (" & Session("Id") & ",'" & Request("tipo") & "','" & Replace(Request("dir1"),"'","''") & "','" & Replace(Request("ciudadnombre"),"'","''") & "',"
            sQuery = sQuery & "'" & Replace(Request("dir2"),"'","''") & "','" & Request("cp") & "','" & Replace(Request("zona"),"'","''") & "'," & LimpiaNum(Request("rentaviv")) & ","
            sQuery = sQuery & "'" & Replace(Request("gente"),"'","''") & "', '" & Request("fuma") & "','" & Request("mascota") & "')"
	
			Err.Clear	
			rst.Open sQuery, sProvider
			if Err then 
				Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
				Response.Redirect "NuCasaRegOfrezcoFront.asp"
			end if

			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>"
			sBody=sBody & Request("tipo") & " introducid@ por " & Session("Usuario") & " en " & Request("ciudadnombre") & CasaDetalles("", Session("Id"), "Admin", "ON")
			sSubject="Nuev@ " & Request("tipo") & " para Vivir en " & Request("ciudadnombre")
			Mail "hector@domoh.com", sSubject, sBody

			NuevoAnuncio
			if Request("id")="nuevo" and Request("Foto")<>"on" then Response.Redirect "QuHomeUsuario.asp?msg=Piso+Añadido."
		else
			'---- Modificar Piso
			if Request("tipo")="" or Request("ciudadnombre")="" or Session("usuario")="" or Session("email")="" then 
				Response.Redirect "TrDomoh.asp?Sesión+Finalizada"
			end if

		 	sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Request("ciudadnombre"),"'","''") & " (" & Replace(Request("zona"),"'","''") & ")', descripcion='" & Replace(Request("descripcion"),"'","''") & "', "
            sQuery = sQuery & " idiomaes='" & Request("idiomaes") & "', idiomaen='" & Request("idiomaen") & "', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id") 

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

		 	sQuery = "UPDATE Pisos SET tipo='" & Request("tipo") & "', dir1='" & Replace(Request("dir1"),"'","''") & "', ciudadnombre='" & Request("ciudadnombre") & "', dir2='" & Replace(Request("dir2"),"'","''") & "',"
			sQuery = sQuery & " cp='" & Request("cp") & "', zona='" & Replace(Request("zona"),"'","''") & "', rentaviv=" & LimpiaNum(Request("rentaviv")) & ", gente='" & Replace(Request("gente"), "'", "''")  & "',"
		 	sQuery = sQuery & " fuma='" & Request("fuma")  & "', mascota='" & Request("mascota")  & "' WHERE id=" & Request("id") 
			rst.Open sQuery, sProvider
			if Err then 
				Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
				Response.Redirect "NuCasaRegOfrezcoFront.asp"
			end if
			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>" & Request("tipo") & " modificad@ por " & Session("Usuario") & CasaDetalles("", Session("id"), "Admin", "ON")
			Mail "hector@domoh.com", "Anuncio de " & Request("tipo") & " para Vivir Modificado", sBody
			if Request("Foto")<>"on" then Response.Redirect "QuHomeUsuario.asp?msg=Piso+Modificado."
			Session("Id")=Request("id")
		end if
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<form name=frm action=TrCasaRegOfertaFoto.asp method=post enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Session("Id")%> />
<div class=container>
    <!-- #include file="IncTrDestacado.asp" -->
	<div class=main>
<%	if Request("foto")="on" or Request("op")="FotoBorrada" then %>
        <a title=
<%		if Request("op")<>"FotoBorrada" and (Request("id")="" or Request("id")="nuevo") then %>
    		'Página Inicial' href='QuDomoh.asp?msg=<%=Server.URLEncode("Te hemos enviado un e-mail para completar tu registro, por favor pulsa el enlace que verás en él.")%>'>
<% 		else %>
			'Volver a Tu Perfil' href=QuHomeUsuario.asp?msg=Piso+Modificado.>
<% 		end if %>
			No quiero subir fotos</a></div>
	<div>
		<p>La primera foto se utilizará como vista previa</p>
<% 
	    if Session("Id")<>"" then
		    rst.Open "SELECT * FROM Fotos WHERE piso=" & Session("id"), sProvider
		    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Err.Description
		    while not rst.Eof 
%>
		<img alt='Foto' title='Para borrar, pulsa en el link de la derecha' src=<%=rst("foto")%> /><a title='Borrar Foto y Volver' href='TrCasaRegOfertaFoto.asp?borrar=<%=rst("id")%>'>Borrar</a></div></div>
<% 
		        rst.Movenext
		    wend 
		    rst.Close
	    end if 
%>
<!-- #include file="IncTrUFotos.asp" -->
<%	else %>
<!-- #include file="IncTrUEnvioMail.asp" -->
<%	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>