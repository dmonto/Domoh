<% 
	option explicit
	dim rst, sProvider, sProvider2, sQuery

    sProvider="PROVIDER=SQLOLEDB;DATA SOURCE=94.130.53.116;UID=domoh;PWD=lx_8L3p2;DATABASE=domoh"
    set rst = Server.CreateObject("ADODB.recordset")
	Server.ScriptTimeout=1000
%>
<html><head><title>Test</title></head>
<body>
<%
	on error goto 0
	
	Response.Write "Test<br/>"
	sQuery="SELECT * FROM Provincias"
	rst.Open sQuery, sProvider
	if not rst.Eof then	Response.Write "Test OK<br/>"
	rst.Close
%>
</body>
</html>
