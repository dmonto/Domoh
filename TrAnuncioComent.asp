<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
 	dim numVal, sUsuario, sBody
 	
	if Request("comentario")="" then Response.Redirect "QuAnuncioDetalle.asp?tabla=Pisos&id=" & Request("id") & "&coment=si"

 	numVal=3
 	numVal=CLng(Request("valoracion"))

	sUsuario=Request("usuario")
	if sUsuario="" then sUsuario=Session("Usuario")
	if sUsuario="" then sUsuario="Anónimo"
	
 	sQuery = "INSERT INTO AnuncioComent (anuncio,usuario,fechaalta,texto,valoracion) VALUES (" & Request("id") & ", '" & sUsuario & "', GETDATE(), '" & Replace(Request("comentario"),"'","''") & "'," & numVal & ")"
	rst.Open sQuery, sProvider
	Response.Cookies("Valoracion"&Request("id"))=Request("valoracion")

 	sQuery = "SELECT * FROM Usuarios u INNER JOIN Anuncios a ON u.usuario=a.usuario WHERE a.id=" & Request("id")
	rst.Open sQuery, sProvider
	sBody= sUsuario & " dice: " & numVal & " - " & Request("comentario") & " sobre el piso (" & rst("cabecera") & ")"
	Mail "hector@domoh.com", "Nuevo Comentario a Piso", sBody

	sBody= "Estimado " & rst("nombre") & ", alguien ha dejado un nuevo comentario sobre tu anuncio (" & rst("cabecera") & ")."
	sBody= sBody & "Para verlo entra ahora en <a title='Ir a Domoh' href='http://www.domoh.com/TrDomohLogOn.asp'>Domoh</a> con tu usuario y password."
	Mail rst("email"), "Nuevo Comentario a tu Anuncio", sBody
	rst.Close
	Response.Redirect "QuAnuncioDetalle.asp?tabla=Pisos&id=" & Request("id") & "&coment=si"
%>
