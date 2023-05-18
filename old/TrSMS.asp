<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	'--- NO BORRAR TrSMS.asp acceso desde SEPOMO
	dim sMens, numId, sBody
	
	sMens=Right(Request("msg"), Len(Request("msg"))-6)

	if sMens="" then
		Mail "", "Donación Recibida", "El número " & Request("movil")	& " ha enviado una donación"
        Response.Write "OK Gracias por tu donación"
		Response.End
	end if
	
	sBody = SMens & " recibido del movil " & Request("movil")	& "<br>"
	
	numId=CLng(Right(sMens,Len(sMens)-3))

	if UCase(Left(sMens,3))="INF" then
		if numId then
			sQuery=" SELECT * FROM Sorteos WHERE id="& numId
			rst.Open sQuery, sProvider
		end if
	
		sBody = sBody & "<br/>" 
        sBody = sBody & rst("email") & " confirma Sorteo InfinitamenteGay<br/>" 
	
		rst.Close
		sQuery=" UPDATE Sorteos SET sms='" & Format(Date) & "' WHERE id=" & numId
		rst.Open sQuery, sProvider
        Response.Write "OK Has entrado en el sorteo"
    end if

	if UCase(Left(sMens,3))="ING" then
		if numId then
			sQuery=" SELECT * FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.id="& numId
			rst.Open sQuery, sProvider
		end if
	
		sBody = sBody & "<br/>" 
        sBody = sBody & rst("usuario") & " quiere que traduzcamos el piso (" & rst("cabecera") & ") de " & rst("ciudadnombre") & "<br/>" 
        Response.Write "OK tu piso será traducido"
    end if

	if UCase(Left(sMens,3))="DES" then
		if numId then
			sQuery=" UPDATE Anuncios SET destacado=destacado+1, fechafindestacado= DATEADD(DAY,10,GETDATE()) WHERE id="& numId
			rst.Open sQuery, sProvider
			sQuery=" SELECT * FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.id="& numId
			rst.Open sQuery, sProvider
		end if
	
		sBody = sBody & "<br/>" 
        sBody = sBody & rst("usuario") & " ha destacado el piso (" & rst("cabecera") & ") de " & rst("ciudadnombre") & "<br/>" 
        Response.Write "OK tu piso será destacado"
	end if

	Mail "", "SMS Recibido", sBody
%>
