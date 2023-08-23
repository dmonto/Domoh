<!--#include file="IncNuBD.asp"-->
<!--#include file="IncNuMail.asp" -->
<!--#include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top'"> 
<div class=container>
	<div class=logo><a title='Salir' href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class=col-6><a href="<%=Request("origen")%>" class=linkutils><%=MesgS("Nuevo Anuncio","New Advert")%></a></div>
    <div class=tituSec><%=MesgS("Usuario Duplicado","User Already Created")%></div>
    <h1>
<% 
	dim sUsuario, sBody
	sUsuario=Request("usuario")
	if sUsuario="" then sUsuario=Session("Usuario")
	sBody="Pidiendo otro usuario para sustituir a _Nuevo" & sUsuario & ". Desde " & Request("origen")
	Mail "","Usuario Duplicado",sBody
	if Session("Idioma")="" then 
%>
        El usuario <%=sUsuario%> ya está registrado en domoh.
<% else %>
        Username <%=sUsuario%> is already registered.
<% end if %>
    	</h1>
    <form name=frm action=<%=Request("origen")%> method=post>
	    <input type=hidden name=op value=NuevoUsuario /><input type=hidden name=foto value=<%=Request("foto")%> /><input type=hidden name=idiomaen value=<%=Request("idiomaen")%> />
	<h2>* <%=MesgS("Nuevo Nombre de Usuario","New User Name")%>:  <input maxlength=60 name=usuario size=20 /> <input name=submit type=submit class=btnLogin id=submit value='<%=MesgS("Continuar","Next")%>'/></h2></form>
    </div>
<!-- #include file="IncPie.asp" -->
</body>
