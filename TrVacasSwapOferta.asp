<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<%
	dim sBody
	on error goto 0
		
	if Request("op")="NuevoUsuario" then
		sQuery = "SELECT * FROM Usuarios WHERE usuario='" & Request.Form("usuario") & "'"
		rst.Open sQuery, sProvider
		if Not(rst.Eof) then Response.Redirect "TrUsuarioYaExiste.asp?origen=TrVacasSwapOferta.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto") & "&idiomaen=" & Request.Form("idiomaen")
		rst.Close
		sQuery = "UPDATE Usuarios SET usuario='" & Request.Form("usuario") & "' " & " WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		sQuery = "UPDATE Anuncios SET usuario='" & Request.Form("usuario") & "' " & " WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		Session("Usuario")=Request.Form("usuario")
		PreBienvenida "TrVacasSwapOferta.asp"
	elseif Request("op")<>"FotoBorrada" and Request("op")<>"Test" then
		if Session("Usuario")="" then
			'--- Usuario Nuevo 
			Session("TipoUsuario")="Casero"
			Session("Activo")="Mail"

			if Request("password")="" then 
				Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), Request.Form("usuario")
                Response.Write "Algo fue mal"
				Response.End
			end if

			InsertUsuario Request
			if Session("Usuario")=Request.Form("usuario") then PreBienvenida "TrVacasSwapOferta" & Session("Idioma") & ".asp"
		end if	

		if Request("id")="" or Request("id")="nuevo" then
			'--- Piso Nuevo 
		 	sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
			sQuery = sQuery & "VALUES ('Pisos','Si','" & Session("Usuario") & "','" & Replace(Request("ciudadnombre"),"'","''") & " (" & Replace(Request("zona"),"'","''") & ")',"
            sQuery = sQuery & "'" & Replace(Request("descripcion"),"'","''") & "',0,'" & Request("idiomaes") & "', '" & Request("idiomaen") & "', GETDATE(), GETDATE())"

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then 
				Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
				Response.Redirect "TrVacasSwapOfrezcoFront.asp"
			end if
				
			sQuery="SELECT MAX(id) AS maxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOferta", sQuery & " - " & Err.Description
			Session("Id")=rst("maxId")
			rst.Close

		 	sQuery = "INSERT INTO Pisos (id, tipo, dir1, ciudadnombre, dir2, cp, zona, rentaviv, gente, fuma, mascota) "
			sQuery = sQuery & "VALUES (" & Session("Id") & ",'Piso','" & Replace(Request("dir1"),"'","''") & "',"
			sQuery = sQuery & "'" & Replace(Request("ciudadnombre"),"'","''") & "','" & Replace(Request("dir2"),"'","''") & "','" & Request("cp") & "','" & Replace(Request("zona"),"'","''") & "',"
			sQuery = sQuery & LimpiaNum(Request("rentaviv")) & ",'" & Replace(Request("gente"),"'","''") & "', '" & Request("fuma") & "','" & Request("mascota") & "') "
	
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOferta", sQuery & " - " & Err.Description

			Session("Ciudad")=Request("ciudad")
			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>"
			sBody=sBody & "Piso para Intercambio introducido por " & Session("Usuario") & " en " & Request("ciudadnombre") & CasaDetalles("", Session("Id"), "Admin", "ON")
			Mail "hector@domoh.com", "Nuevo Piso para Intercambio en " & Request("ciudadnombre"), sBody
			NuevoAnuncio
			if Request("id")="nuevo" then 
				if Request("foto")<>"on" then Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Piso+Aùadido.","Advert+Added.")
			elseif Session("Usuario")<>Request.Form("usuario") then 
				Response.Redirect "TrUsuarioYaExiste.asp?origen=TrVacasSwapOferta.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto") & "&idiomaen=" & Request.Form("idiomaen")
			end if
		else
			'---- Modificar Piso
		 	sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Request("ciudadnombre"),"'","''") & " (" & Replace(Request("zona"),"'","''") & ")'," & " descripcion='" & Replace(Request("descripcion"),"'","''") & "',"
			sQuery = sQuery & " idiomaes='" & Request("idiomaes") & "', idiomaen='" & Request("idiomaen") & "', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id") 

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrVacasSwapOferta", sQuery & " - " & Err.Description

		 	sQuery = "UPDATE Pisos SET dir1='" & Replace(Request("dir1"),"'","''") & "', ciudadnombre='" & Replace(Request("ciudadnombre"),"'","''") & "',"
			sQuery = sQuery & " dir2='" & Replace(Request("dir2"),"'","''") & "', cp='" & Request("cp") & "', zona='" & Replace(Request("zona"),"'","''") & "', fuma='" & Request("fuma")  & "',"
			sQuery = sQuery & " mascota='" & Request("mascota")  & "' WHERE id=" & Request("id") 
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
			Session("Id")=Request("id")
			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Piso para intercambio modificado por " & Session("Usuario") & CasaDetalles("", Session("Id"), "Admin", "ON")
			Mail "hector@domoh.com", "Piso para Intercambio Modificado", sBody

			if Request("foto")="" then	
				Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Anuncio+Modificado.","Advert+Updated")
			end if
		end if
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<form name=frm action=TrVacasSwapOfertaFoto.asp method=post enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
	<div class=topmenu>
			<a href=TrVacas.asp class=linkutils><%=MesgS("Anuncios Vacaciones","Holiday Adverts")%> </a> &gt; 
			<a href=TrVacasSwapOfrezcoFront.asp class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a>  &gt; <%=MesgS("Subir Fotos","Upload Pictures")%></div></div>
    <div class=banner><% Response.Write MesgS("PublicaciÛn de Apartamento para Intercambio de Vacaciones","Flat Swap Advert Publication")%></div>
    <div>
<%	if Request("foto")="on" or Request("op")="FotoBorrada" then %>
        <a title=
<%		if Request("op")<>"FotoBorrada" and (Request("id")="" or Request("id")="nuevo") then %>
		    '<%=MesgS("P·gina Inicial","Home Page")%>' 
            href="QuDomoh.asp?msg=<% Response.Write _
            Server.URLEncode(MesgS("Te hemos enviado un e-mail para completar tu registro, por favor pulsa el enlace que ver·s en Èl.", "We have sent you an e-mail, click on the link within")) %>"
<% 		else %>
            '<%=MesgS("Volver a tu Perfil","Back to Your Home")%>' href='QuHomeUsuario.asp?msg=<% Response.Write MesgS("Piso+Modificado", "Advert+Updated")%>.'>
<% 		end if %>
			<% Response.Write MesgS("No quiero subir fotos", "I will not upload pictures right now")%></a></div></div>
<div>
    <p><% Response.Write MesgS("La primera foto se utilizar· como vista previa","The first picture will be used as preview")%></p>
<% 
	    if Session("Id")<>"" then
		    rst.Open "SELECT * FROM Fotos WHERE piso=" & Session("Id"), sProvider
		    if Err then 
			    Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Err.Description
		    end if
		    while not rst.Eof 
%>
	<img alt='Foto' title='<% Response.Write MesgS("Para borrar la foto pulsa a la derecha","Press the button on the right to delete")%>' src=<%=rst("foto")%> />
	<a title='<%=MesgS("Borrar esta Foto","Delete this picture")%>' href='TrVacasRegOfertaFoto.asp?borrar=<%=rst("id")%>'><%=MesgS("Borrar","Delete")%></a>
<% 
		        rst.Movenext
		    wend 
		    rst.Close
	    end if 
%>
    </div>
<!-- #include file="IncTrUFotos.asp" -->
<%	else %>
<!-- #include file="IncTrEnvioMail.asp" -->
<%	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>
