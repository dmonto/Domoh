<!--#include file="IncNuBD.asp"-->
<html>
<!--#include file="IncTrCabecera.asp"-->
<body onload="window.parent.location.hash='top';"> 
<%
	dim sAmigo
	
	if Session("Usuario")="" then 
		Response.Write MesgS("Sesión Finalizada","Session Ended")
		Response.End
	end if
	
	sQuery="SELECT * FROM Amigos WHERE usuario='" & Session("Usuario") & "' AND amigo='" & Request("Usuario") & "'"
	rst.Open sQuery, sProvider
	if rst.Eof then sAmigo="No" else sAmigo=rst("activo")
	rst.Close
	rst.Open "SELECT * FROM Usuarios WHERE usuario='" & Request("usuario") & "'", sProvider
%>
<div>
	<div title='<%=MesgS("Nombre Completo","Complete Name")%>'><%=rst("nombre") & " " & rst("apellidos")%></div>
	<div title='<%=MesgS("Nombre de Usuario","Username")%>'>Nick: <%=rst("usuario")%></div>
	<div>
<%  if rst("mostrarperfil")<>"" or sAmigo="Si" then %>
		<%=rst("perfil")%>
<% 		if rst("foto")<>"" then %>
		<img alt='Foto' title='<%=MesgS("Foto de tu amigo","Picture of your Friend")%>' width=100 src='http://domoh.com/<%=rst("foto")%>'/>
<% 	
		end if 
	elseif sAmigo="Mail" then 
%>
		<%=rst("nombre")%> <% Response.Write MesgS("aún no te ha confirmado como amigo", "has not yet confirmed you as a friend")%>. 
		Pulsa 
		<a title='<% Response.Write MesgS("Reenviar el e-mail para que te confirme como amigo", "re-send invitation to your friend")%>' href='TrAmigoNuevo.asp?usuario=<%=Server.URLEncode(rst("Usuario"))%>'>
            <%=MesgS("aquí­","here")%></a> <%=MesgS("para recordárselo","to send a reminder")%>.
<% 	else %>
		<% Response.Write MesgS("Para ver el perfil y blog de este usuario debe aceptarte como amigo", "In order to see this user's profile and blog, he/she has to confirm you as a friend")%>. 
		<%=MesgS("Pulsa","Click")%> 
		<a title='<% Response.Write MesgS("Enviarle un e-mail para que te confirme como amigo", "Send invitation to your friend")%>' href='TrAmigoNuevo.asp&usuario=<%=Server.URLEncode(rst("Usuario"))%>'>
            <%=MesgS("aquí­","here")%></a>.
<% 
    end if 
    if rst("mostraremail")<>"" then 
%>
	</div></div>
	<div> e-Mail: <a title='<%=MesgS("Pulsa para enviar un mail","Click to send an e-mail")%>' href="mailto:<%=rst("email")%>"><%=rst("email")%></a>
<% end if %>
		</div>
</body>
</html>
