<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<%
	dim sBody, sSubject
		
	if Request("op")="NuevoUsuario" then
		sQuery = "SELECT * FROM Usuarios WHERE usuario='" & Request.Form("usuario") & "'"
		rst.Open sQuery, sProvider
		if Not(rst.Eof) then 
			Response.Redirect "TrUsuarioYaExiste.asp?origen=TrCasaCompraOferta.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto") & "&idiomaen=" & Request.Form("idiomaen")
		end if
		rst.Close
		sQuery = "UPDATE Usuarios SET usuario='" & Request.Form("usuario") & "' WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		sQuery = "UPDATE Anuncios SET usuario='" & Request.Form("usuario") & "' WHERE usuario='" & Session("usuario") & "'"
		rst.Open sQuery, sProvider
	elseif Request("op")<>"FotoBorrada" and Request("op")<>"Test" then
		if Session("Usuario")="" then
			'--- Usuario Nuevo 
			if Request("password")="" then
                Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), Request.Form("usuario")
                Response.Write "Algo fue mal"
                Response.End        
            end if

            Session("TipoUsuario")="VendeCasa"
			InsertUsuario Request
			if Session("Usuario")=Request.Form("usuario") then PreBienvenida "TrCasaCompraOferta.asp"
		else
			sQuery = "UPDATE Usuarios SET numanuncios=numanuncios+1 WHERE usuario='" & Session("Usuario") & "'"
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
		end if	

		if Request("id")="" or Request("id")="nuevo" then
			'--- Piso Nuevo 
			if Request("tipo")="" then Response.Redirect "TrCasaCompraOfrezcoFront.asp"

	 		sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
			sQuery = sQuery & "VALUES ('Pisos','Si','" & Session("Usuario") & "','" & Replace(Request("ciudadnombre"),"'","''")
			sQuery = sQuery & " (" & Replace(Request("zona"),"'","''") & ")','" & Replace(Request("descripcion"),"'","''") & "',0,'" & Request("idiomaes") & "', '" & Request("idiomaen") & "', GETDATE(), GETDATE())"

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then 
				Mail "diego@domoh.com", "Error en TrCasaCompraOferta", sQuery & " - " & Err.Description
				Response.Redirect "TrCasaCompraOfrezcoFront.asp"
			end if
				
			sQuery="SELECT MAX(id) AS maxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrCasaCompraOferta", sQuery & " - " & Err.Description
			Session("Id")=rst("maxId")
			rst.Close

		 	sQuery = "INSERT INTO Pisos (id, tipo, dir1, ciudadnombre, dir2, cp, zona, precio) "
			sQuery = sQuery & "VALUES (" & Session("Id") & ", '" & Request("tipo") 	& "','" & Replace(Request("dir1"),"'","''") & "','" & Replace(Request("ciudadnombre"),"'","''") & "',"
			sQuery = sQuery & "'" & Replace(Request("dir2"),"'","''") & "','" & Request("cp") & "','" & Replace(Request("zona"),"'","''") & "'," & LimpiaNum(Request("precio")) & ")"
	
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrCasaCompraOferta", sQuery & " - " & Err.Description

			Session("Ciudad")=Request("ciudad")
			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>"
			sBody=sBody & Request("tipo") & " introducid@ por " & Session("Usuario") & " en " & Request("ciudadnombre") & CasaDetalles("", Session("Id"), "Admin", "ON")
			sSubject="Nuev@ " & Request("tipo") & " para Vender en " & Request("ciudadnombre")
			Mail "hector@domoh.com", sSubject, sBody
			if Request("id")="nuevo" then 
				if Request("foto")<>"on" then Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Piso+Añadido.","Advert+Added.")
			elseif Session("Usuario")<>Request.Form("usuario") then 
				Response.Redirect "TrUsuarioYaExiste.asp?origen=TrCasaCompraOferta.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto")& "&idiomaen=" & Request.Form("idiomaen")
			end if
		else
			'---- Modificar Piso
		 	sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Request("ciudadnombre"),"'","''") & " (" & Replace(Request("zona"),"'","''") & ")',"
			sQuery = sQuery & " descripcion='" & Replace(Request("descripcion"),"'","''") & "', idiomaes='" & Request("idiomaes") & "', idiomaen='" & Request("idiomaen") & "',"
		 	sQuery = sQuery & " fechaultimamodificacion=GETDATE() WHERE id=" & Request("id") 

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrCasaRegOferta", sQuery & " - " & Err.Description

		 	sQuery = "UPDATE Pisos SET tipo='" & Request("tipo") & "', dir1='" & Replace(Request("dir1"),"'","''") & "', ciudadnombre='" & Replace(Request("ciudadnombre"),"'","''") & "',"
			sQuery = sQuery & " dir2='" & Replace(Request("dir2"),"'","''") & "', cp='" & Request("cp") & "',"
			sQuery = sQuery & " zona='" & Replace(Request("zona"),"'","''") & "', precio=" & LimpiaNum(Request("precio")) & " WHERE id=" & Session("Id") 
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>" & Request("tipo") & " modificad@ por " & Session("Usuario") & CasaDetalles("", Session("Id"), "Admin", "ON")
			Mail "hector@domoh.com", Request("tipo") & " Modificad@", sBody
			if Request("foto")<>"on" then Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Piso+Modificado.","Advert+Updated.")
		end if
	end if
%>
<!--#include file="IncTrCabecera.asp"-->
<body onload="window.parent.location.hash='top'"> 
<form name=frm action=TrCasaCompraOfertaFoto.asp method=post enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Session("Id")%> />
<div class=container>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
    <div>
	    <a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios de Vivienda","Property Adverts")%> &gt;</a>
		<a href=TrCasaCompraOfrezcoFront.asp class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a>  &gt; <%=MesgS("Subir Fotos","Upload Pictures")%></div></div>
<div class=tituSec>&nbsp;</div>
<div class=tituSec><% Response.Write MesgS("Publicación de Piso o Casa para Venta", "House Sale Advert Publication")%></div>
<div>
<%	if Request("foto")="on" or Request("op")="FotoBorrada" then %>
    <div>
<%		if Request("op")<>"FotoBorrada" and (Request("id")="" or Request("id")="nuevo") then %>
	    <a title='<%=MesgS("Página Inicial","Home Page")%>' 
		    href='QuDomoh.asp?msg=
            <% Response.Write Server.URLEncode(MesgS("Te hemos enviado un e-mail para completar tu registro, por favor pulsa el enlace que verás en él.", _
            "We have sent you an e-mail to tha address you just gave us. Please click on the link so we can finish your registration."))%>'>
<% 		else %>
            </a>
		<a title='<%=MesgS("Volver a Tu Perfil","Back to Personal Page")%>' href='QuHomeUsuario.asp?msg=<% Response.Write MesgS("Piso+Modificado","Advert+Updated")%>.'>
<% 		end if %>
		    <% Response.Write MesgS("No quiero subir fotos", "I won't upload pictures at this moment")%></a></div></div>
<% 
	    if Session("Id")<>"" then
		    rst.Open "SELECT * FROM Fotos WHERE Piso=" & Session("Id"), sProvider
		    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Err.Description
		    while not rst.Eof 
%>
<div>
    <img title='<% Response.Write MesgS("Para borrarla pulsa en el enlace", "To remove, click the link on the right")%>' src=<%=rst("foto")%> />
    <a href='TrCasaCompraOfertaFoto.asp?borrar=<%=rst("id")%>'><%=MesgS("Borrar","Remove")%></a></div>
<% 
		        rst.Movenext
		    wend 
		    rst.Close
	    end if 
%>
<!-- #include file="IncTrUFotos.asp" -->
<%	else %>
<!-- #include file="IncTrEnvioMail.asp" -->
<%	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>