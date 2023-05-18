<!-- #include file="IncNuBD.asp" -->
<% 
	Response.Buffer=false 
	Server.ScriptTimeout=1000
%>
<!-- #include file="IncTrGrid.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim Grid, sWhere, bEjec, bExiste, rstExec, dFecha, numViejo
	on error goto 0
	
	set rstExec = Server.CreateObject("ADODB.recordset")
	set fso = CreateObject("Scripting.FileSystemObject")
	if Request("op")="Ejecutar" then bEjec=3
%>
<h1><a title='Hacer cambios en BD' href='?op=Ejecutar'>Ejecutar</a></h1>