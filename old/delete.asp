<%@ Language=VBScript %>
<!-- #include file="connection.asp" -->
<%

sqlStm = "Delete from product_master where n_prod_id = "

for each thing in Request.Form

	if i = 0 then
		sqlStmEnd = Request.Form(thing)
	end if

i = i + 1	
next

sqlStm = sqlStm&""&sqlStmEnd

con.Execute sqlStm
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY onLoad="document.tempFrm.submit();">

<form name="tempFrm" method="post" action="demo.asp">
	<input type="hidden" name="mess" value="Record Deleted Successfully!">
</form>
</BODY>
</HTML>