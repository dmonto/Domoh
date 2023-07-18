<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<%
	dim numMaximo, sBody
	
	numMaximo=0
	numMaximo=CLng(Request("maximo"))
	
	if Request("op")="NuevoUsuario" then
		sQuery = "SELECT * FROM Usuarios where usuario='" & Request.Form("usuario") & "'"
		rst.Open sQuery, sProvider
		if Not(rst.Eof) then 
			Response.Redirect "TrUsuarioYaExiste.asp?origen=TrCasaCompraDemanda.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto") & "&idiomaen=" & Request.Form("idiomaen")
		end if
		rst.Close
		sQuery = "UPDATE Usuarios SET usuario='" & Request.Form("usuario") & "'" & " where usuario='" & Session("usuario") & "'"
		rst.Open sQuery, sProvider
		sQuery = "UPDATE Anuncios SET usuario='" & Request.Form("usuario") & "' WHERE usuario='" & Session("usuario") & "'"
		rst.Open sQuery, sProvider
		Session("Usuario")=Request.Form("usuario")
		PreBienvenida "TrCasaCompraDemanda.asp"
	elseif Request("op")<>"Test" then
		if Session("Usuario")="" then
			'--- Usuario Nuevo 
			if Request("password")="" then
                Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), Request.Form("usuario")
                Response.Write "Algo fue mal"
                Response.End        
            end if

			Session("TipoUsuario")="Inquilino"
			Session("Activo")="Mail"
	
			InsertUsuario Request
			if Session("Usuario")=Request.Form("usuario") then PreBienvenida "TrCasaCompraDemanda.asp"
		else
			sQuery="UPDATE Usuarios SET numanuncios=numanuncios+1 WHERE usuario='" & Session("Usuario") & "'"
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
		end if	

		if Request("id")="" or Request("id")="nuevo" then
			'--- Inquilino Nuevo 
			if Request("cabecera")="" then Response.Redirect "TrCasaCompraDemandaFront.asp"
		
		 	sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
			sQuery = sQuery & "VALUES ('Inquilinos','Si','" & Session("Usuario") & "','" & Replace(Request("cabecera"),"'","''") & "',"
			sQuery = sQuery & "'" & Replace(Request("anuncio"),"'","''") & "',0,'" & Request("idiomaes") & "', '" & Request("idiomaen") & "', GETDATE(), GETDATE())"

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrCasaCompraDemanda", sQuery & " - " & Err.Description
		
			sQuery="SELECT MAX(id) AS maxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en TrCasaCompraDemanda", sQuery & " - " & Err.Description
			Session("Id")=rst("maxid")
			rst.Close

	 		sQuery = "INSERT INTO Inquilinos (id, tipoviv, ciudad, maximo) VALUES (" & Session("Id") & ",'Compra'," & "'" & Request("ciudad") & "'," & numMaximo & ")"

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then 
				if Err then Mail "diego@domoh.com", "Error en TrCasaCompraDemanda", sQuery & " - " & Err.Description
				Response.Redirect "TrCasaCompraDemandaFront.asp"
			end if

			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Anuncio de compra de piso introducido por " & Session("Usuario") & InquilinoDetalles("", Session("Id"))
			Mail "hector@domoh.com", "Nuevo Comprador", sBody
			if Request("id")="nuevo" then Response.Redirect "QuHomeUsuario.asp?msg=" & Server.URLEncode(MesgS("Anuncio Añadido.","Advert Added"))
			if Session("Usuario")<>Request.Form("usuario") then 
				Response.Redirect "TrUsuarioYaExiste.asp?origen=TrCasaCompraDemanda.asp&usuario=" & Request.Form("usuario") & "&foto=" & Request.Form("foto") & "&idiomaen=" & Request.Form("idiomaen")
			end if
		else
			'---- Modificar Inquilino
		 	sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Request("cabecera"),"'","''") & "', descripcion='" & Replace(Request("anuncio"),"'","''") & "',"
			sQuery = sQuery & " idiomaes='" & Request("idiomaes") & "', idiomaen='" & Request("idiomaen") & "', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id") 

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
	 		sQuery = "UPDATE Inquilinos SET ciudad='" & Replace(Request("ciudad"),"'","''") & "', maximo=" & numMaximo & " WHERE id=" & Request("id") 

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
			sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Anuncio de compra de piso modificado por " & Session("Usuario") & InquilinoDetalles("", Session("Id"))
			Mail "hector@domoh.com", "Anuncio de compra de piso Modificado", sBody
			Response.Redirect "QuHomeUsuario.asp?msg=" & Server.URLEncode(MesgS("Anuncio Modificado.","Advert Updated"))
		end if
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<div class=container>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
		<div class=right>
			<a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios de Vivienda","Property Adverts")%> &gt;</a>
			<a href=TrCasaCompraDemandaFront.asp class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a>  &gt; <%=MesgS("Subir Fotos","Upload Pictures")%></div></div>
    <h1 class=tituSec><% Response.Write MesgS("Publicación de Anuncio de Búsqueda para Comprar", "Flat Request for Buying Advert Publication")%></h1>
    <!-- #include file="IncTrEnvioMail.asp" -->
    </div>
<!-- #include file="IncPie.asp" -->
</body>