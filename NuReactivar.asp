<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';">
<div class=container>
	<div class=logo><!--#include file="IncNuSubMenu.asp"--></div>
	<div class=main><h1><% Response.Write MesgS("En breve nos pondremos en contacto contigo.", "We will get in touch with you shortly.")%></h1></div></div>
<% if Request("usuario")<>"" then Mail "hector@domoh.com", "Reactivación", Request("usuario") & " quiere reactivar su usuario." %>
<!-- #include file="IncPie.asp" -->
</body>