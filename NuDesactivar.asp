<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	if Request("op")<>"Test" then 
        if Session("Usuario")="" then Response.Redirect "TrLogOn.asp"
	
	    sQuery ="DELETE FROM Fotos WHERE piso IN (SELECT id FROM Anuncios WHERE usuario='" & Session("Usuario") & "')"
	    rst.Open sQuery, sProvider
	    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

	    sQuery ="DELETE FROM Pisos WHERE id IN (SELECT id FROM Anuncios WHERE usuario='" & Session("Usuario") & "')"
	    rst.Open sQuery, sProvider
	    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

	    sQuery ="DELETE FROM Curriculums WHERE id IN (SELECT id FROM Anuncios WHERE usuario='" & Session("Usuario") & "')"
	    rst.Open sQuery, sProvider
	    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

	    sQuery ="DELETE FROM Inquilinos WHERE id IN (SELECT id FROM Anuncios WHERE usuario='" & Session("Usuario") & "')"
	    rst.Open sQuery, sProvider
	    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

	    sQuery ="DELETE FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
	    rst.Open sQuery, sProvider
	    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

	    sQuery ="DELETE FROM Misc WHERE id IN (SELECT id FROM Anuncios WHERE usuario='" & Session("Usuario") & "')"
	    rst.Open sQuery, sProvider
	    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

	    sQuery ="DELETE FROM Trabajos WHERE id IN (SELECT id FROM Anuncios WHERE usuario='" & Session("Usuario") & "')"	
	    rst.Open sQuery, sProvider
	    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

	    sQuery ="DELETE FROM Usuarios WHERE usuario='" & Session("Usuario") & "'"
	    rst.Open sQuery, sProvider
	    if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

	    Mail "hector@domoh.com", "Usuario Eliminado", Session("Usuario") & " se ha autodestruido."
	    Session.Abandon
    end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';">
<div class=container>
	<div class=topmenu><!--#include file="IncNuSubMenu.asp"--></div>
	<div><h1><% Response.Write MesgS("Toda tu información ha sido eliminada de nuestros registros.", "All of your details have been deleted.") %></h1></div>
</div>
<!-- #include file="IncPie.asp" -->
</body>