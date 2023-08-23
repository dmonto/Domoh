<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<%
	dim sBody
		
	if Request("op")="NuevoUsuario" then
		sQuery = "SELECT * FROM Usuarios where usuario='" & Request.Form("usuario") & "'"
		rst.Open sQuery, sProvider
		if Not(rst.Eof) then 
			Response.Redirect "TrUsuarioYaExiste.asp?origen=TrCasaRegOferta.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto") & "&idiomaen=" & Request.Form("idiomaen")
		end if
		rst.Close
		sQuery = "UPDATE Usuarios SET usuario='" & Request.Form("usuario") & "'" & " WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		sQuery = "UPDATE Anuncios SET usuario='" & Request.Form("usuario") & "' where usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		Session("Usuario")=Request.Form("usuario")
		PreBienvenida "NuCasaRegOferta.asp"
	elseif Request("op")<>"FotoBorrada" and Request("op")<>"Test" then
		if Session("Usuario")="" then
			'--- Usuario Nuevo 
    	    if Request("password")="" or Request("nombre")="" or Request("usuario")="" then 
	    	    Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), " Usuario " & Session("Usuario")
                Response.Write "Algo fue mal"
	    	    Response.End
    	    end if

			Session("TipoUsuario")="Casero"
			Session("Activo")="Mail"
			InsertUsuario Request
			if Session("Usuario")=Request.Form("usuario") then PreBienvenida Request.Servervariables("SCRIPT_NAME")
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
				Mail "diego@domoh.com", "Error en TrCasaRegOferta", sQuery & " - " & Err.Description
				Response.Redirect "TrCasaRegOfrezcoFront.asp"
			end if
				
			sQuery="SELECT MAX(id) AS numMaxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrCasaRegOferta", sQuery & " - " & Err.Description
			Session("Id")=rst("numMaxId")
			rst.Close

		 	sQuery = "INSERT INTO Pisos (id, tipo, dir1, ciudadnombre, dir2, cp, zona, rentaviv, gente, fuma, mascota) "
			sQuery = sQuery & "VALUES (" & Session("Id") & ",'" & Request("tipo") & "','" & Replace(Request("dir1"),"'","''") & "','" & Replace(Request("ciudadnombre"),"'","''") & "',"
            sQuery = sQuery & "'" & Replace(Request("dir2"),"'","''") & "','" & Request("cp") & "','" & Replace(Request("zona"),"'","''") & "'," & LimpiaNum(Request("rentaviv")) & ",'" 
            sQuery = sQuery & Replace(Request("gente"),"'","''") & "', " & "'" & Request("fuma") & "','" & Request("mascota") & "')"
	
			Err.Clear	
			rst.Open sQuery, sProvider
			if Err then 
				Mail "diego@domoh.com", "Error en TrCasaRegOferta", sQuery & " - " & Err.Description
				Response.Redirect "TrCasaRegOfrezcoFront.asp"
			end if

			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>"
			sBody=sBody & Request("tipo") & " introducid@ por " & Session("Usuario") & " en " & Request("ciudadnombre") & CasaDetalles("", Session("Id"), "Admin", "ON")
			Mail "hector@domoh.com", "Nuev@ " & Request("tipo") & " para Vivir en " & Request("ciudadnombre"), sBody

			NuevoAnuncio
			if Request("id")="nuevo" then 
				if Request("Foto")<>"on" then Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Piso+Añadido.","Advert+Added.")
			elseif Session("Usuario")<>Request.Form("usuario") then 
				Response.Redirect "TrUsuarioYaExiste.asp?origen=TrCasaRegOferta.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto")& "&idiomaen=" & Request.Form("idiomaen")
			end if
		else
			'---- Modificar Piso
			if Request("tipo")="" or Request("ciudadnombre")="" or Session("usuario")="" or Session("email")="" then Response.Redirect "QuDomoh.asp?msg=" & MesgS("Sesión+Finalizada.","Session+Ended.")

		 	sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Request("ciudadnombre"),"'","''") & " (" & Replace(Request("zona"),"'","''") & ")', descripcion='" & Replace(Request("descripcion"),"'","''") & "',"
			sQuery = sQuery & " idiomaes='" & Request("idiomaes") & "', idiomaen='" & Request("idiomaen") & "', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id") 

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrCasaRegOferta", sQuery & " - " & Err.Description

		 	sQuery = "UPDATE Pisos SET tipo='" & Request("tipo") & "', dir1='" & Replace(Request("dir1"),"'","''") & "', ciudadnombre='" & Request("ciudadnombre") & "', dir2='" & Replace(Request("dir2"),"'","''") & "',"
			sQuery = sQuery & " cp='" & Request("cp") & "', zona='" & Replace(Request("zona"),"'","''") & "', rentaviv=" & LimpiaNum(Request("rentaviv")) & ", gente='" & Replace(Request("gente"), "'", "''")  & "', "
            sQuery = sQuery & " fuma='" & Request("fuma")  & "', mascota='" & Request("mascota")  & "' WHERE id=" & Request("id") 
			rst.Open sQuery, sProvider
			if Err then 
				Mail "diego@domoh.com", "Error en TrCasaRegOferta", sQuery & " - " & Err.Description
				Response.Redirect "TrCasaRegOfrezcoFront.asp"
			end if
			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>" & Request("tipo") & " modificad@ por " & Session("Usuario") & CasaDetalles("", Session("Id"), "Admin", "ON")
			Mail "hector@domoh.com", "Anuncio de " & Request("tipo") & " para Vivir Modificado", sBody
			if Request("foto")<>"ON" then 
				Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Piso+Modificado.","Advert+Updated.")
			end if
			Session("Id")=Request("id")
		end if
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<form name=frm action=TrCasaRegOfertaFoto.asp method=post enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Session("Id")%> />
<div class=container>
	<div class=logo>
        <div class=left><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
	    <div class=right>
    		<a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios de Vivienda","Property Adverts")%> &gt;</a>
		    <a href=TrCasaRegOfrezcoFront.asp class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a> &gt; <%=MesgS("Subir Fotos","Upload Pictures")%></div></div></div>
<h1 class=tituSec><% Response.Write MesgS("Publicación de Piso o Habitación para Vivir", "Post New Advert for Rent")%></h1>
    <div>
<%	if Request("Foto")="on" or Request("op")="FotoBorrada" then %>
<%		if Request("op")<>"FotoBorrada" and (Request("id")="" or Request("id")="nuevo") then %>
	    <a title='Página Inicial' href='QuDomoh.asp?msg=
            <% Response.Write Server.URLEncode(MesgS("Te hemos enviado un e-mail para completar tu registro, por favor pulsa el enlace que verás en él.", _
            "We have sent you an e-mail for completing your registration, you will find a link on it, please click on it."))%>'>
<% 		else %>
            </a>
	    <a title='Volver a Tu Perfil' href='QuHomeUsuario.asp?msg=<%=MesgS("Piso+Modificado.","Advert+Updated.")%>'>
<% 		end if %>
    	    <%=MesgS("No quiero subir fotos","I won't upload pictures at this moment")%></a></div>
	<div>
        <p><%=MesgS("La primera foto se utilizará como vista previa","The first picture will be used as preview")%></p>
<% 
	    if Session("id")<>"" then
		    rst.Open "SELECT * FROM Fotos WHERE Piso=" & Session("id"), sProvider
		    if Err then Mail "diego@domoh.com", "Error en TrCasaRegOferta", Err.Description
            while not rst.Eof 
%>
        <img title='<%=MesgS("Para borrar, pulsa en el link de la derecha","To delete click the link on the right")%>' src=<%=rst("foto")%> />
        <a href='TrCasaRegOfertaFoto.asp?borrar=<%=rst("id")%>'><%=MesgS("Borrar","Delete")%></a>
<% 
		        rst.Movenext
            wend 
            rst.Close
        end if 
%>
        <p><%=MesgS("Ficheros con Fotos Digitales (máximo 300 KB cada una)","Files with digital pictures (max 300KB each)")%></p></div>
	<!-- #include file="IncTrUFotos.asp" -->
<%	else %>
	<!-- #include file="IncTrEnvioMail.asp" -->
<%  end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>