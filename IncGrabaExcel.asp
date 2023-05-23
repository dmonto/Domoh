<%
'---- Exporta a Excel una Tabla -------------------------------
Function Graba(vTabla, vTitulo)
	dim sFld, s

	on error goto 0
	rst.Open "SELECT * FROM " & vTabla, sProvider
				
	if vTitulo <> "" then s="<header>" & vTitulo & "</header>"
    s=s & "<table>"

	while Not rst.Eof 
		s=s & "<tr>"
 		for each sFld in rst.Fields 
 			s=s & "<td>" & sFld & "</td>"
 		next 
 		s=s & "</tr>"
 		rst.Movenext
	wend
    s=s & "</table>"
	rst.Close
	Response.Write vTabla 
    Graba=s
End Function 

'---- Exporta a Excel una Tabla con WHERE 
Function GrabaParte(vTabla, vTitulo, vCond)
	dim sFld, s

	on error goto 0
	rst.Open "SELECT * FROM " & vTabla & " WHERE " & vCond, sProvider
				
	if vTitulo <> "" then s="<header>" & vTitulo & "</header>"
    s=s & "<table>"

    while Not rst.Eof 
		s=s & "<tr>"
 		for each sFld in rst.Fields 
 			s=s & "<td>" & sFld & "</td>"
 		next 
 		s=s & "</tr>"
 		rst.Movenext
	wend
	rst.Close
    s=s & "</table>"
    Response.Write vTabla & " " & vCond
	GrabaParte=s
End Function 

'---- Exporta a Excel una Tabla con objeto Office
Sub GrabaOWC(vTabla, vTitulo, vSheet)
	dim numRow, numCol, sFld
	numRow=1
	numCol=1
	
	vSheet.Name = vTitulo
	rst.Open "SELECT * FROM " & vTabla, sProvider
	
	with vSheet
		while Not rst.Eof 
			for each sFld in rst.Fields             
				.Cells(numRow, numCol).Value = sFld
				numCol=numCol+1
			Next
			numRow=numRow+1
 		rst.Movenext
      wend
   end with
End Sub
%>
