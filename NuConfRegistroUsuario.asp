<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	'--- NO BORRAR Link desde IncNuMail
	if (Request("usuario")="" or Request("id")="") and Request("op")<>"Test" then Response.Redirect "index.htm"

 	sQuery = "SELECT * FROM Usuarios WHERE usuario='" & Request("usuario") & "'"
	rst.Open sQuery, sProvider

	Session("Idioma")=Request("idioma")		
	if rst("id")<>CLng(Request("id")) then Response.Redirect "index.htm"
	Session("EMail")=rst("email")
	Session("Usuario")=rst("usuario")
	Session("Nombre")=rst("nombre")
	Session("Apellidos")=rst("apellidos")		
	Session("Password")=rst("password")		
	rst.Close
	if Session("Tipo")="Inquilino" then	Bienvenida "RegistroDemanda.asp" else Bienvenida "RegistroOferta.asp"
			 	
 	sQuery = "UPDATE Usuarios SET activo='Si' WHERE usuario='" & Session("Usuario") & "'"
	rst.Open sQuery, sProvider
%>
<!--#include file="IncTrCabecera.asp"-->
<div class=container>
	<div><a title='<%=MesgS("Home de Domoh","Domoh's Home Page")%>' href=index.htm target=_top><img alt='Domoh' src=images/NuLogo.gif /></a></div>
	<div><!--#include file="IncTrSubMenu.asp"--></div>
	<div>
		<%=MesgS("¡Enhorabuena! ya estás registrado en ", "Congratulations! You are now registered in ")%>
		<%=MesgS("domoh.com. En breve recibirás un correo con la confirmación del registro, ", "Domoh. You will shortly receive an e-mail with your registration, ")%>
		<%=MesgS("así como tus datos","as well as your login information")%>.</div>
    <div><%=MesgS("Pulsa","Click")%><a title='<%=MesgS("Home de Domoh","Home Page")%>' href=index.htm><%=MesgS("aquí","here")%></a> para comenzar.</div></div>
<!-- #include file="IncPie.asp" -->
