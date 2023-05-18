<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<title>Advert Confirmation</title>
<%
	dim numMaximo, sBody
	
	if Request("maximo")="" then 
        numMaximo=0 
    else 
        numMaximo=CLng(Request("maximo"))
    end if
	
	if Session("Usuario")="" then
		'--- Usuario Nuevo 
		Session("TipoUsuario")="Inquilino"
		Session("Activo")="Mail"

	    if Request("password")="" then 
		    Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), " Usuario " & Request.Form("usuario")
            Response.Write "Algo fue mal"
	    	Response.End
    	end if

		InsertUsuario Request
		PreBienvenida "NuCasaCompraDemandaEn.asp"
	else
		sQuery="UPDATE Usuarios SET numanuncios=numanuncios+1 WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
	end if	

	if Request("id")="" or Request("id")="nuevo" then
		'--- Inquilino Nuevo 
	 	sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
		sQuery = sQuery & "VALUES ('Inquilinos','Si','" & Session("Usuario") & "','" & Replace(Request("cabecera"),"'","''") & "',"
		sQuery = sQuery & "'" & Replace(Request("anuncio"),"'","''") & "',0,'" & Request("idiomaes") & "', '" & Request("idiomaen") & "', #" & Format(Now) & "#, #" & Format(Now) & "#)"

		Err.Clear
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuCasaCompraDemandaEn", sQuery & " - " & Err.Description
		
		sQuery="SELECT MAX(id) AS maxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuCasaCompraDemandaEn", sQuery & " - " & Err.Description
		Session("Id")=rst("maxId")
		rst.Close

	 	sQuery = "INSERT INTO Inquilinos (id, tipoviv, ciudad, maximo) VALUES (" & Session("Id") & ",'Compra','" & Request("ciudad") & "'," & numMaximo & ")"

		Err.Clear
		rst.Open sQuery, sProvider
		if Err then 
			Mail "diego@domoh.com", "Error en NuCasaCompraDemandaEn", sQuery & " - " & Err.Description
			Response.Redirect Request.Servervariables("SCRIPT_NAME")
		end if

		sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Anuncio de compra de piso introducido por " & Session("Usuario") & InquilinoDetalles("En", Session("Id"))
		Mail "hector@domoh.com", "Nuevo Inquilino", sBody
		if Request("id")="nuevo" then Response.Redirect "QuHomeUsuario.asp?msg=Anuncio+Añadido."
	else
		'Modificar Inquilino
	 	sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Upload.Form("cabecera"),"'","''") & "',descripcion='" & Replace(Upload.Form("anuncio"),"'","''") & "',"
		sQuery = sQuery & "idiomaes='" & Upload.Form("idiomaes") & "', idiomaen='" & Upload.Form("idiomaen") & "',fechaultimamodificacion=#" & Format(Now) & "# WHERE id=" & Upload.Form("id") 

		Err.Clear
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuCasaCompraDemandaEn", sQuery & " - " & Err.Description

	 	sQuery = "UPDATE Inquilinos SET ciudad='" & Replace(Request("ciudad"),"'","''") & "',maximo=" & numMaximo & " WHERE id=" & Request("id") 

		Err.Clear
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuCasaCompraDemandaEn", sQuery & " - " & Err.Description
		sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>Anuncio de Comprador modificado por " & Session("Usuario") & InquilinoDetalles("En", Request("id"))
		Mail "hector@domoh.com", "Anuncio de Comprador Modificado", sBody
		Response.Redirect "QuHomeUsuario.asp?msg=Anuncio+Modificado."
	end if
%>
<!--#include file="IncTrCabecera.asp"-->
<body onLoad="window.parent.location.hash='top';
<%	if instr(Session("email"),"hotmail")>0 then %>
	detalles('images/AntiSpam.gif');
<%	end if %>
"> 
<table width=100%>
	<tr><!--#include file="IncNuSubMenu.asp"--></tr>
	<tr>
   		<td>
			<p><p align=center>
			<font color=#FF0000>WARNING</font>
			<font color=red>: PLEASE CHECK YOUR EMAIL ADDRESS. WE WILL SEND YOU AN EMAIL WITH INSTRUCTIONS TO ACTIVATE YOUR ADVERT. YOU WILL HAVE TO ANSWER THIS EMAIL.&nbsp; 
            <u><b>IF YOU DON'T FOLLOW THE INSTRUCTIONS WE WILL NOT POST YOUR ADVERT. <br></b></u></font>
            <u><b><font size=5 color=#FF00FF>HOTMAIL</font><font size=5 color=#008000> USERS AND OTHERS: PLEASE CHECK OUT YOUR “Junk Mail” or “Spam”</font></b></u></p><p>&nbsp;</p></td></tr></table>
<!-- #include file="IncPie.asp" -->