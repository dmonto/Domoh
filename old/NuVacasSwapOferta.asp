<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim sBody
		
	if Session("Usuario")="" then
		'---- Usuario Nuevo 
	    if Request("password")="" then 
		    Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), " Usuario " & Request.Form("usuario")
            Response.Write "Algo fue mal"
	    	Response.End
    	end if

		Session("TipoUsuario")="Casero"
		Session("Activo")="Mail"
		InsertUsuario Request
		PreBienvenida "NuVacasSwapOferta" & Session("Idioma") & ".asp"
	end if	

	if Request("id")="" or Request("id")="nuevo" then
		'---- Piso Nuevo 
	 	sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
		sQuery = sQuery & "VALUES ('Pisos','Si','" & Session("Usuario") & "','" & Replace(Request("ciudadnombre"),"'","''") 
		sQuery = sQuery & " (" & Replace(Request("zona"),"'","''") & ")','" & Replace(Request("descripcion"),"'","''") & "',0,'" & Request("idiomaes") & "', '" & Request("idiomaen") & "', GETDATE(), GETDATE())"

		Err.Clear
		rst.Open sQuery, sProvider
		if Err then 
			Mail "diego@domoh.com", "Error en NuVacasSwapOferta", sQuery & " - " & Err.Description
			Response.Redirect "TrVacasSwapOfrezcoFront.asp"
		end if
				
		sQuery="SELECT MAX(id) AS maxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuVacasSwapOferta", sQuery & " - " & Err.Description
		Session("Id")=rst("maxId")
		rst.Close

	 	sQuery = "INSERT INTO Pisos (id, tipo, dir1, ciudadnombre, dir2, cp, zona, rentaviv, gente, fuma, mascota) "
		sQuery = sQuery & "VALUES (" & Session("Id") & ",'Piso','" & Replace(Request("dir1"),"'","''") & "','" & Replace(Request("ciudadnombre"),"'","''") & "','" & Replace(Request("dir2"),"'","''") & "',"
        sQuery = sQuery & "'" & Request("cp") & "','" & Replace(Request("zona"),"'","''") & "'," & LimpiaNum(Request("rentaviv")) 	& ","
		sQuery = sQuery & "'" & Replace(Request("gente"),"'","''") & "', '" & Request("fuma") & "','" & Request("mascota") & "') "
	
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuVacasSwapOferta", sQuery & " - " & Err.Description

		Session("Ciudad")=Request("ciudad")
		sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>"
		sBody=sBody & "Piso para Intercambio introducido por " & Session("Usuario") & " en " & Request("ciudadnombre") & CasaDetalles("", Session("Id"), "Admin", "ON")
		Mail "hector@domoh.com", "Nuevo Piso para Intercambio en " & Request("ciudadnombre"), sBody
		NuevoAnuncio
		if Request("id")="nuevo" and Request("Foto")<>"on" then Response.Redirect "QuHomeUsuario.asp?msg=Piso+Añadido."
		end if
	else
		'---- Modificar Piso
	 	sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Request("ciudadnombre"),"'","''") & " (" & Replace(Request("zona"),"'","''") & ")', descripcion='" & Replace(Request("descripcion"),"'","''") & "',"
        sQuery = sQuery & " idiomaes='" & Request("idiomaes") & "', idiomaen='" & Request("idiomaen") & "', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id") 

		Err.Clear
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuVacasSwapOferta", sQuery & " - " & Err.Description

	 	sQuery = "UPDATE Pisos SET "
		sQuery = sQuery & " dir1='" & Replace(Request("dir1"),"'","''") & "', ciudadnombre='" & Replace(Request("ciudadnombre"),"'","''") & "',"
		sQuery = sQuery & " dir2='" & Replace(Request("dir2"),"'","''") & "', cp='" & Request("cp") & "',"
		sQuery = sQuery & " zona='" & Replace(Request("zona"),"'","''") & "', fuma='" & Request("fuma")  & "',"
		sQuery = sQuery & " mascota='" & Request("mascota")  & "' WHERE id=" & Request("id") 
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
		sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>"
		sBody=sBody & "Piso para intecambio modificado por " & Session("Usuario") & CasaDetalles("", Session("Id"), "Admin", "ON")
		Mail "hector@domoh.com", "Piso para Intercambio Modificado", sBody
		Session("Id")=Request("id")
	end if
%>
<title>Vacaciones - Subir Fotos Intercambio</title>
<body onload="window.parent.location.hash='top';
<%	if instr(Session("email"),"hotmail")>0 then %>
	detalles('images/AntiSpam.gif');
<%	end if %>
"> 
<form name=frm action=TrVacasSwapOfertaFoto.asp method=post enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Request("id")%>>
<table width=100%>
	<tr>
<!--#include file="IncNuSubMenu.asp"-->
<%	if Request("foto")="on" or Request("op")="FotoBorrada" then %>
	</tr>
	<tr>
		<td align=center>
			<p>La primera foto se utilizará como vista previa</p>
<%
 		if Session("Id")<>"" then 
			if Request("id")="" or Request("id")="nuevo" then
%>
				<a title='Ir a Tus Pisos' href=QuHomeUsuario.asp?msg=Piso+Añadido.>No quiero subir más fotos</a><br>
<% 			else %>
				<a title='Ir a Tus Pisos' href=QuHomeUsuario.asp?msg=Piso+Modificado.>No quiero subir más fotos</a><br>
<%
			end if
			
			rst.Open "SELECT * FROM Fotos WHERE piso=" & Session("id"), sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Err.Description
			while not rst.Eof 
%>
			<img alt='Foto' title='Para quitarla, pulsa en el enlace de la derecha' src=http://domoh.com/<%=rst("foto")%>>
			<a title='Borrar esta Foto' href='TrVacasRegOfertaFoto.asp?borrar=<%=rst("id")%>'>Borrar</a>
<% 
				rst.Movenext
			wend 
			rst.Close
		else
%>
			<a title='Ir a la página principal' 
				href=NuDomoh.asp?msg=<% Response.Write Server.URLEncode("En breve recibirás un correo con un enlace para confirmar la publicación. " & _
				"Por favor púlsalo y podremos publicar el anuncio.")%>>
			No quiero subir fotos</a><br>
<% 		end if %>
			<p>Ficheros con Fotos Digitales (máximo 300 KB cada una)</p>
			<p><input title='Elegir el fichero de la primera foto' type=file name=foto1></p>
			<p><input title='Elegir el fichero de la segunda foto' type=file name=foto2></p>
			<p><input title='Elegir el fichero de la tercera foto' type=file name=foto3></p>
			<p><input title="Elegir el fichero de la cuarta foto" type=file name=foto4></p>
			<p><input title="Elegir el fichero de la quinta foto" type=file name=foto5></p>
			<p><input title="Elegir el fichero de la sexta foto" type=file name=foto6></p></td>
	<tr>
		<td align=center>
			<input title=Grabar type=submit value="Subir Fotos" 
				></td></tr>
<%	
	elseif Request("id")<>"" then 
		Response.Redirect "QuHomeUsuario.asp?msg=Piso+Modificado."		
	end if
%>
</form>
<!-- #include file="IncTrEnvioMail.asp" -->
</table>
<!-- #include file="IncPie.asp" -->