<!-- #include file="IncNuBD.asp" -->
<div>
<%	
	dim numFot, i
	
	rst.Open "SELECT COUNT(*) AS numFot FROM Fotos WHERE piso=" & Request("id"), sProvider 
	numFot=CLng(rst("numFot"))
	rst.Close	
    if numFot=0 then
        Response.Write "<h1>No hay fotos</h1>"
        Response.End
    end if
	rst.Open "SELECT * FROM Fotos WHERE piso=" & Request("id"), sProvider 
	i=1
	do while not rst.Eof
%>
	<img id=foto<%=i%> class=fotogrande title='Foto de la Vivienda' src="http://domoh.com/<%=rst("foto")%>" onmouseover="pos=<%=i%>;foco();"/>	
<%		
		if (i mod 3)=0 then Response.Write "</div><div>"
		i=i+1
		rst.Movenext
	loop
%>
</div>