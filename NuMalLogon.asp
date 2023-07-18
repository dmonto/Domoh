<!--#include file="IncNuBD.asp"-->				
<!--#include file="IncTrCabecera.asp"-->				
<body onload="window.parent.location.hash='top';">
<div class=container>
	<div class=logo><!--#include file="IncNuSubMenu.asp"--></div>
	<div class=main>
<% 	
    on error goto 0
    if Session("Activo")="No" then %>
<% 		if Session("Idioma")="" then %>
		Tu usuario está desactivado. Pulsa <a title='Enviar solicitud de reactivación' href='NuReactivar.asp?Usuario=<%=Server.URLEncode(Session("Usuario"))%>'>aquí</a> para activarlo de nuevo.
<% 		else %>
		Your account is inactive. Click <a title='Sent reactivation request' href='NuReactivar.asp?Usuario=<%=Server.URLEncode(Session("Usuario"))%>'>here</a> in order to reactivate it.
<% 		
        end if 
 	elseif Session("Activo")="Mail" then 
%>
        <!--#include file="IncTrEnvioMail.asp"-->				
<% 	else %>
		<%=MesgS("Usuario o Clave Incorrectos.","Incorrect User name/password.")%> <a title='Volver a introducir claves' href=TrDomohLogOn.asp>Volver a intentarlo</a>
<% 
	end if 
	Session.Abandon
%>
    </div></div>
<!-- #include file="IncPie.asp" -->
</body>
