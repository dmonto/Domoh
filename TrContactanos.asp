<!--#include file="IncNuBD.asp"-->
<!--#include file="IncTrCabecera.asp"-->
<body onload="window.parent.location.hash='top';">
<div class=container>
    <div class=logo><a href=PeMenu.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<h1 class=banner><%=MesgS("Cont�ctanos","Contact Us")%></h1>
    <div class=main>
		<h2><% Response.Write MesgS("Si tienes cualquier duda comentario o sugerencia, no dudes en enviar un correo a la siguiente direcci�n:", "Please address any comment or suggestion to:")%>
    	<a title='<% Response.Write MesgS("Mandar un mail de sugerencias", "Send an e-mail with comments")%>' href=mailto:sugerencias@domoh.com><%=MesgS("sugerencias@domoh.com","comments@domoh.com")%></a></h2>
		<h2><% Response.Write MesgS("�Problemas utilizando la p�gina? Te ayudamos", "Any problem using the website? We'll be glad to help you")%>
		<a title="<% Response.Write MesgS("Mandar un mail con dudas o problemas", "Send an e-mail with questions or problems")%>" href=mailto:webmaster@domoh.com><%=MesgS("aqu�","here")%></a></h2>
		<h2><% Response.Write MesgS("Si est�s interesado en anunciarte en domoh.com escr�benos a", "Interested in advertising on domoh? Contact us on")%> 
		<a title="<% Response.Write MesgS("Dinos c�mo te gustar�a anunciarte", "Tell us how you want to advertise in domoh")%>" href="<% Reponse.Write MesgS("mailto:publicidad@domoh.com", "mailto:advertising@domoh.com")%>">
        <%=MesgS("publicidad@domoh.com","advertising@domoh.com")%></a></h2></div></div>
<!-- #include file="IncPie.asp" -->
</body>
