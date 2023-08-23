<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';"> 
<form action=NuOlvido.asp method=post>
<div class=container>
<% if Request("msg")<>"" then %>
	<p class=banner><%=Request("msg")%></p>
<% end if %>
	<h1><%=MesgS("¿No recuerdas tu clave?","Can't remember your password?")%></h1>
    <h2><% Response.Write MesgS("Introduce tu e-mail y te la enviaremos:", "Input here your e-mail address and we'll send it to you:")%> 
	<input title='<% Response.Write MesgS("Dirección de correo electrónico que nos diste al registrarte", "Enter the address you used when you registered")%>' name=email size=20 autocomplete=email />
	<input title="<%=MesgS("Te enviamos tu clave","We will send you the password")%>" type=submit value=<%=MesgS("Enviar","Send")%> class=boton /> </h2></div></form>
<!-- #include file="IncPie.asp" -->
</body>