<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncGrabaExcel.asp" -->
<div>
<%	
	Response.ContentType = "application/vnd.ms-excel"
	Response.Write Graba(Request("tabla"), "")
%>
</div>
