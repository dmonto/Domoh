<!-- #include file="IncNuBD.asp" -->
<%
	'---- Back de NuAdminFuentes --------------------------------------------
	dim sFuente
	
	for each sFuente in Request.Form
		if Request.Form.Item(sFuente)="Borrar" then	rst.Open "DELETE FROM Fuentes WHERE id=" & sFuente, sProvider
	next

	if Request("desc")<>"" or Request("descen")<>"" then
		sQuery="INSERT INTO Fuentes (""desc"", descen) VALUES ('" & Request("desc") & "','" & Request("descen") & "')"
		rst.Open sQuery, sProvider
	end if 

	Response.Redirect "NuAdminFuentes.asp"
	