<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<% 
    sProvider="PROVIDER=SQLOLEDB;DATA SOURCE=64.71.180.27;UID=domoh;PWD=hecaran@123;DATABASE=domoh"
    set rst = Server.CreateObject("ADODB.recordset")
	Server.ScriptTimeout=1000
%>
<html><head><title>Test</title></head>
<body>
    <!-- #include file="IncPie.asp" -->
<%
	on error goto 0
%>	
</body>
</html>
