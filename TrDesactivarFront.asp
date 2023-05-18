<!--#include file="IncNuBD.asp"-->
<!--#include file="IncTrCabecera.asp"-->
<body onload="window.parent.location.hash='top';">
<form action=NuDesactivar.asp method=post>
<div class=container>
	<div class=main>
		<h1><% Response.Write MesgS("Al confirmar, tus datos y los de tus pisos serán eliminados de nuestra base de datos.", "Upon pressing OK, all your information and posts will be erased from our database.")%></h1>
		<h2><input title='<%=MesgS("Borrar tus datos","Erase your info")%>' type=submit value=<%=MesgS("Desactivar","Go")%> class=btnForm /></h2></div></div></form>
<!-- #include file="IncPie.asp" -->
</body>
