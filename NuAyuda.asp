<!--#include file="IncNuBD.asp"-->
<head>
    <title>Domoh - Pisos y Habitaciones Gay-Friendly - Ayuda</title>
    <!--#include file="IncTrCabecera.asp"-->
</head>
<body onload="window.parent.location.hash='top'">
<div class=container>
	<div class=menu><!--#include file="IncTrSubMenu.asp"--></nav>
	<h1><%=MesgS("AYUDA","HELP")%></h1>
    <h2>
		<% Response.Write MesgS("Si tienes algún problema y necesitas ayuda ", "Having problems using this website? ")%>
		<% Response.Write MesgS("para utilizar nuestra web escribe un correo a ", "Please, write us an email to")%>
		<a title="<%=MesgS("Enviarnos un e-mail","Send us an e-mail!")%>" href=mailto:atencioncliente@domoh.com>atencioncliente@domoh.com</a> 
		<% Response.Write MesgS("con tus preguntas y nos pondremos en contacto contigo para ayudarte.", "and we will help you immediately!") %></p></div>
<!-- #include file="IncPie.asp" -->
</body>