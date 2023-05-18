<% 
    dim sIdi 
    sIdi="Es"
%>
<div class=aviso>
    <h2>
	<% Response.Write UMesgS(sIdi, "Antes de empezar a utilizar domoh.com debes pulsar el enlace que hay ","Before starting to use domoh.com we need you to click on the link ")%>
	<% Response.Write UMesgS(sIdi, "en el correo 'Confirmación Registro en domoh.com' que te hemos enviado ", "that you will find in the e-mail we just sent you ")%>
	<% Response.Write UMesgS(sIdi, "a la dirección con la que te registraste. Así completaremos tu registro. ", "to the address you just gave us. That way your registration will be complete. ")%>
	<% Response.Write UMesgS(sIdi, "Muchas gracias","Thank you very much")%>.</h2>
    <h2>
    <% Response.Write UMesgS(sIdi, "SI NO HAS RECIBIDO NUESTRO CORREO, POR FAVOR COMPRUEBA TU CORREO 'SPAM' ", "IF YOU HAVEN'T RECEIVED OUR E-MAIL, PLEASE CHECK YOUR 'SPAM' ")%>
    <% Response.Write UMesgS(sIdi, "O 'CORREO NO DESEADO' ", "OR 'UNWANTED' FOLDERS")%></h2></div>
    