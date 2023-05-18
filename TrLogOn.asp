<!-- #include file="IncNuBD.asp" -->
<%
	dim rstUpdate
	on error goto 0

	if Request("id")<>"" then
		sQuery= "SELECT * FROM Usuarios WHERE usuario='" & Request("id") &"'"
		rst.Open sQuery, sProvider

		if rst.Eof then	
			Response.Redirect "NuMalLogon.asp" 
		elseif rst("Password") <> Request("password") then 
			Response.Redirect "NuMalLogon.asp"
		end if
		
		Session("EMail")=rst("email")
		Session("Usuario")=rst("usuario")
		Session("Nombre")=rst("nombre")
		Session("Apellidos")=rst("apellidos")		
		Session("Password")=rst("password")		
		Session("Activo")=rst("activo")		

		if rst("activo")<>"Si" then Response.Redirect "NuMalLogon.asp"

		Session("Provincia")=rst("provincia")
		Session("UltimaVisita")=rst("ultimavisita")
		Session("TipoUsuario")=rst("tipo")
		Session("ContrataFin")=rst("contratafin")
		Session("TipoResiVaca")=rst("tiporesivaca")
		Session.LCID=1036
		Session.Timeout=60
		rst.Close

		if Session("Provincia") > 0 then 
			sQuery= "SELECT * FROM Provincias WHERE id = " & Session("Provincia")
			rst.Open sQuery, sProvider
			Session("NombreProvincia")=rst("nombre")
		end if
		
		Set rstUpdate = Server.CreateObject("ADODB.recordset")
		sQuery="UPDATE Usuarios SET ultimavisita=GETDATE() WHERE usuario='" & Session("Usuario") & "'"
		rstUpdate.Open sQuery, sProvider
	elseif Session("Usuario")="" then
		Response.Redirect "QuDomoh.asp"
	end if

	if Session("TipoUsuario") = "Admin" then Response.Redirect "NuAdmin.asp" else Response.Redirect "QuHomeUsuario.asp"
%>
