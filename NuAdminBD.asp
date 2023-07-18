<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrGrid.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<form name=frm action=NuAdminBD.asp method=post>
<div class=container>
	<div class=main>
<%
	dim Grid
	on error goto 0
	
	if Left(Request("sent"), 6) = "select" then
		SetupGrid Grid
		Grid.SQL = Request("sent")
        Grid.Display 
%>
	<div class=footer>
<%
	elseif Request("sent")<>"" then
		sQuery = Request("sent")
		Response.Write "<h2>" & sQuery & "</h2>"
		rst.Open sQuery, sProvider
	end if
%>
    Instrucción para BD: <textarea title='Mucho cuidaditorrrr!!!' cols=50 rows=7 name=sent></textarea><input title='Segurorhr??' class=boton type=submit value=Ejecutar /></div></div>
</form>
<!-- #include file="IncPie.asp" -->