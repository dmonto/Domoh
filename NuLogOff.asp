<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%	Session.Abandon %>
<body onload="window.parent.location.hash='top';">
<div class=container>
	<div class=main>
	<h1><%=MesgS("Gracias por visitar Domoh.com.","Thanks for visiting domoh.com!!")%></h1>
	<p><% Response.Write MesgS("Permítenos sugerirte unas recomendaciones para seguir navegando...", "Some interesting recommendations to keep surfing...")%></p>
</div>
</div>
<!-- #include file="IncPie.asp" -->
</body>
