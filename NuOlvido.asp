<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim sBody
	
	if Request("idioma")="En" then Session("Idioma")="En"
	sQuery="SELECT nombre, password, usuario FROM Usuarios WHERE email = '" & Request("email") &"'"
	rst.Open sQuery, sProvider
	if rst.Eof then Response.Redirect "NuOlvidoFront.asp?msg=" & MesgS("e-mail+incorrecto","incorrect+address")
	
	if Session("Idioma")="" then
        sBody="<p>Tu usuario es: " & rst("usuario") & " y tu clave es: " & rst("password") & ". <a title='Home de Domoh' href='http://domoh.com'>Entrar a domoh.com</a></p>"
	else
        sBody="Your user name is: " & rst("usuario") & " and your password is: " & rst("password") & ". <a title='Domoh''s Home' href='http://domoh.com'>Enter domoh.com</a>"
	end if
	Mail Request("email"), MesgS("Acceso a domoh.com", "Access to domoh.com"), sBody
	if Err then Response.Redirect "NuOlvidoFront.asp?msg=" & MesgS("e-mail+incorrecto","incorrect+address")
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';"> 
<div class=container>
	<div class=topmenu><!--#include file="IncNuSubMenu.asp"--></div>
	<div>
<%	if Err then %>
		<%=MesgS("Hubo un problema enviándote la clave:","Delivery was not completed:")%> <%=Err.Description%>
<%	else %> 
		<%=MesgS("Te hemos enviado tu clave a:","We have sent your password to:")%> <%=Request("email")%>
<%	end if %>
	</div></div>
<!-- #include file="IncPie.asp" -->
</body>