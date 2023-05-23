<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrGrid.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim Grid, i
	
	SetupGrid Grid
	Grid.SQL = "SELECT * FROM " & Request("tabla") 
	Grid.sTabla=Request("tabla") 
	if Request("id")<>"" then Grid.SQL = Grid.SQL & " WHERE id=" & Request("id")
	Grid.SQL = Grid.SQL & " ORDER BY 2"
	if Request("auto")="no" then Grid.bAuto=False
	Grid.ExtraFormItems = "<input type=hidden name=tabla value=" & Request("tabla") & "><input type=hidden name=id value=" & Request("id") & ">"
	Grid.ExtraFormItems = Grid.ExtraFormItems & "<input type=hidden name=nombretabla value=" & Request("nombretabla") & ">" 
%>
<h2><%=Request("nombretabla")%> <a title='Vuelca a Excel la tabla de <%=Request("nombretabla")%>' href='NuGrabaExcel.asp?tabla=<%=Request("tabla")%>'>Excel <%=Request("nombretabla")%></a></h2>
<%	Grid.Display %>
<!-- #include file="IncPie.asp" -->