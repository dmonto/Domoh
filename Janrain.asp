<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	'--- NO BORRAR Janrain.asp acceso desde Janrain
	dim sMens, numId, sBody, sURL, objXML, strApiKey, strToken, sResponse, tbOrder, i, nLoc, nLoc2, sEMail
	on error goto 0

	Function TraeToken(sString, sToken)
		dim nLoc, nLoc2, nLoc3, sSubStr
	
		nLoc=0
		nLoc=InStr(sString, """" & sToken & """")

		if nLoc>0 then
			nLoc2=nLoc+Len(sToken)+4
			sSubStr=Right(sString,Len(sString)-nLoc2+1)
			nLoc3=InStr(sSubStr, """")-1
			TraeToken=Left(sSubStr, nLoc3)
		else
			TraeToken=""
		end if		
	End Function
	
	strApiKey = "bfa89ce089b06cc2754641716a42c03ef25f5831"
	strToken = Request("token")
	sURL="https://rpxnow.com/api/v2/auth_info"
	set objXML = Server.CreateObject("Microsoft.XMLHTTP")
	objXML.Open "POST", sURL, False
	objXML.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	objXML.Send "apiKey=" & strApiKey & "&token=" & strToken
	sResponse=objXML.ResponseText
	sEMail=TraeToken(sResponse, "email")

	if sEMail<>"" then
		sQuery= "SELECT * FROM Usuarios WHERE email='" & sEMail & "'"
		rst.Open sQuery, sProvider
		if Session("MenuAction")="" then Session("MenuAction")="QuHomeUsuario.asp"
		Session("Usuario")=Right(TraeToken(sResponse, "identifier"),50)
		Session("EMail")=sEMail
		Session("Password")=TraeToken(sResponse, "providerName")
		Session("Nombre")=Right(TraeToken(sResponse, "displayName"),50)
        Session("Id")=""
        Session("Activo")="Si"

        if rst.Eof then
            Response.Write "Dar de alta el e-mail de:" & sEMail 
            rst.Close
			sQuery = "INSERT INTO Usuarios (tipo, esagencia, activo, numanuncios, nombre, dir1, ciudaddir, dir2, cp, fuente, tel1, tel2, tel3, tel4, email, usuario, password, mostrartel1, mostrartel2, "
			sQuery = sQuery & " mostrartel3, instrucciones, mostrartel4, mostraremail, fechaalta, ultimavisita) "
			sQuery = sQuery & "VALUES ('Janrain','','Si',0,'" & Right(Replace(TraeToken(sResponse, "displayName"),"'","''"),50) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" 
            sQuery = sQuery & TraeToken(sResponse, "email") & "',"
		    sQuery = sQuery & "'" & Right(TraeToken(sResponse, "identifier"),50) & "','" & TraeToken(sResponse, "providerName") & "',NULL,NULL,NULL,NULL,NULL,'ON',GETDATE(),GETDATE())"

			Err.Clear
			rst.Open sQuery, sProvider
			if Err then
				Mail "diego@domoh.com", "Error en Janrain", sQuery & " - " & Err.Description
				Response.End
			end if
			Response.Write "Redirect TrUsuario"
			Response.Redirect "TrUsuario.asp?origen=PeMenu.asp"
		else
			Response.Redirect "PeMenu.asp"
		end if
	end if
	set objXML=Nothing
%>