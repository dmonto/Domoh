<%
'--- SetupGrid Inicializa el Objeto [vGrid]
Sub SetupGrid(vGrid)
	set vGrid = new DigGrid
	vGrid.Connect sProvider, "", ""
	vGrid.ImagePath = "images/grid/"
	vGrid.MaxRows = 50
	vGrid.bAuto=true
End Sub

'--- DigGrid implementa un Grid para edición
Class DigGrid
	Public Prov, ImagePath, MaxRows, SQL, ExtraFormItems, sTabla, tblBorder, editTbl, newTbl, rowNo, curPage, updateImg, deleteImg, bAuto
		
    '--- Connect inicializa el Conector SQL
	Public Sub Connect(vProv, v2, v3)
		Prov=vProv
	End Sub
	
	'--- Display genera el HTML del Grid
    Public default Sub Display()
		dim tblWidth, title, rowLower, rowUpper, r, colLower, colUpper, c, idVal, thing, rstId
		
		editTbl = CInt(Request("editTbl"))
		newTbl = CInt(Request("newTbl"))
		rowNo = CInt(Request("rowNo"))
		curpage = CInt(Request("curpage"))
		if curpage = 0 then curpage = 1

		updateImg = true
		deleteImg = false
		tblWidth = 100
		tblBorder = 0
%>
<script>
    function cancel() {document.tmpFrm.rowNo.value = 0; document.tmpFrm.editTbl.value = 0; document.tmpFrm.newTbl.value = 0; document.tmpFrm.submit();}
    function check(i) {document.tmpFrm.rowNo.value = i; document.tmpFrm.newTbl.value = 0; document.tmpFrm.submit();}
    function checkNew(i) {document.tmpFrm.rowNo.value = 0; document.tmpFrm.editTbl.value = 0; document.tmpFrm.submit();}
    function newRecord() {document.dgridFrm.op.value = "insert"; document.dgridFrm.submit();}
    		
    function deleteYesNo(i)	{
	    var ok;
	
	    ok = confirm("Are you sure you want to delete this record?");
	    if(ok) {document.dgridFrm.op.value = "delete"+i; document.dgridFrm.submit();} 
	} //end of deleteYesNo()
</script>
<form name=tmpFrm method=post><input type=hidden name=editTbl value=1 /><input type=hidden name=newTbl value=1 /><input type=hidden name=rowNo /></form>
<form name=dgridFrm method=post>
	<input type=hidden name=op value=update />
<% 		if Trim(Request.Form("mess")) <> "" then %>
	<h1><%=UCase(Trim(Request.Form("mess")))%></h1>
<%	
		end if

		if Trim(Request.Form("op")) = "update" then
            sQuery = "UPDATE " & sTabla & " SET"
            dim i, sWhere
			i=1
			
			for each thing in Request.Form
				if UCase(thing)="ID" then
					sWhere= " WHERE " & thing & " = "& Request.Form(thing) &""
				elseif thing<>"op" and thing<>"del" then
					sQuery = sQuery & " " & thing & " = '"& Request.Form(thing) &"',"
				end if
				i = i + 1	
			next

			sQuery= left(sQuery,Len(sQuery)-1)&sWhere
			rst.Open sQuery, sProvider
		elseif Trim(Request.Form("op")) = "insert" then
			dim sCampos, numCampos
			sQuery = "INSERT INTO " & sTabla & " ("	
	
			for each thing in Request.Form
				if i > 0 and thing<>"op" and thing<>"del" then
					sQuery = sQuery & thing & ","
					sCampos=sCampos & "'" & Request.Form(thing) & "',"
				end if
				i = i + 1	
			next

			sQuery = Left(sQuery,Len(sQuery)-1) &") VALUES (" & Left(sCampos,Len(sCampos)-1) & ")"
			rst.Open sQuery, sProvider
		elseif left(Request.Form("op"),6) = "delete" then
			sQuery = "DELETE FROM " & sTabla & " WHERE id = " & Right(Request.Form("op"),Len(Request.Form("op"))-6)
			rst.Open sQuery, sProvider
		end if

		rst.CursorType=1
		rst.Open SQL, Prov
		rowLower = CLng(Request("regdesde"))
		rowUpper = rst.RecordCount
			
		if rowUpper<0 then rowUpper=10
			if rowUpper>99 then rowUpper=rowLower+99
				Session("numResults")=100
				colLower = 0
				colUpper = rst.Fields.Count + 2
%>
<div class=container>
	<div class=main>
<%
				PagPie rst.RecordCount,"?"
				Session("numResults")=0
				TableStart tblWidth
				for r = 1 to rowLower
					rst.Movenext
				next
                for r = rowLower to rowUpper
		            RowStart(r)
		            for c = colLower to colUpper
			            if r = rowLower then
				            mainCellStart(0)
					        if c = colLower then
						        Response.Write "<a href='javascript:cancel();'><img src=grid/recycle.gif border=0 alt='Cancel'/></a>"
					        else
						        if c > rst.Fields.Count then 
                                    Response.Write "&nbsp;" 
                                else 
                                    Response.Write rst.Fields(c-1).Name '--- "HEADING"
                                end if
					        end if
				            mainCellEnd()
			            else
				            if c = colLower then
					            mainCellStart(0)
					            Response.Write "&nbsp;"
					            mainCellEnd()
				            else
					            CellStart()
						        if rowNo = r then
							        if c > rst.Fields.Count then
								        if updateImg then
									        Response.Write "<a href='javascript:document.dgridFrm.submit();'><img src=grid/save.gif border=0 alt='Save'/></a>"
									        updateImg = false
								        else
									        Response.Write "<a href='javascript:deleteYesNo("& rst.Fields(0).Value &");'><img src=grid/delete.gif border=0 alt='Delete'/></a>"
									        updateImg = true
								        end if
							        else
								        if c = 1 then '--- readonly field
									        Response.Write "<input type=hidden name='"& rst.Fields(c-1).Name & "' value='"& rst.Fields(c-1).Value & "'/>" & rst.Fields(c-1).Value
								        else	
									        Response.Write "<input type=text name='"& rst.Fields(c-1).Name & "' taborder=0 value='" & rst.Fields(c-1).Value & "'/>"
								        end if
							        end if
						        else
							        if c > rst.Fields.Count then
							            'dgrid.CellStart(0)
								        if updateImg then
									        Response.Write "<a href='javascript:check(" & r & ");'><img src=grid/edit.gif border=0 alt='Edit'/></a>"
									        updateImg = false
								        else
									        Response.Write "<a href='javascript:deleteYesNo(" & rst.Fields(0).Value & ");'><img src=grid/delete.gif border=0 alt='Delete'/></a>"
									        updateImg = true
								        end if
							        else
							            'dgrid.CellStart(40)
								        Response.Write Trim(rst.Fields(c-1).Value)'"test"
							        end if
						        end if
					            CellEnd()
				            end if
			            end if
		            next
	                RowEnd()
	                if r > rowLower then
		                on error goto 0
		                rst.MoveNext
	                end if
	            next
	            numCampos=rst.Fields.Count 
	            '--- For New Record 

                if newTbl = 0 then
	                RowStart(idVal)
		            mainCellStart(0)
			        '--- Response.Write "&nbsp;"
			        Response.Write "<a href='javascript:checkNew("&r&");'><img src=grid/add.gif alt='New Record'/></a>"
		            mainCellEnd()
		            for c = colLower to rst.Fields.Count + 1
			            CellStart()
				        if c = colLower then
					        '--- Response.Write "<a href='JavaScript:checkNew("&r&");'><img src='grid/add.gif' border='0' alt='New Record'></a>"
					        Response.Write "&nbsp;"
				        else
					        '--- if c = rs.Fields.Count + 1 then
					        '---	Response.Write "<a href='JavaScript:checkNew("&r&");'><img src='grid/recycle.gif' border='0' alt='Cancel'></a>"
					        '--- else
						    Response.Write "&nbsp;"
					        '--- end if
				        end if
			            CellEnd()
		            next
	                RowEnd()
                end if
                '--- For New Record 
                '--- Once you clickon New Record 
                if newTbl = 1 then
	                RowStart(idVal)
		            mainCellStart(0)
			        Response.Write "&nbsp;"
		            mainCellEnd()
		            for c = colLower to numCampos + 1
			            CellStart()
				        if c = colLower then
				            if bAuto then
					            Response.Write "(siguiente)"
				            else
					            set rstId = Server.CreateObject("ADODB.recordset")	
					            rstId.Open "SELECT MAX(id)+1 AS nextId FROM " & sTabla, sProvider
					            Response.Write rstId("nextId") & "<input type=hidden name=Id value=" & rstId("nextId") & ">"
					            rstId.Close
				            end if
				        else
					        if c <= numCampos-1 then
						        Response.Write "<input type=text name='" & rst.Fields(c).Name & "'>"
					        else
						        if c = numCampos then
							        Response.Write "<a href='javascript:newRecord();'><img src=grid/save.gif border=0 alt='Save'/></a>"
						        else
							        Response.Write "&nbsp;"
						        end if
					        end if
				        end if
			            CellEnd()
		            next
	                RowEnd()
                end if
                '--- Once you clickon New Record 
	            rst.Close	
                TableEnd()
	        End Sub

	        '--- TableStart genera la cabecera
            Sub TableStart(tblWidth)
		        Response.Write "<div id=grdTbl class='gridtbl container'>" & Chr(13)
	        End Sub 

	        '--- TableStart genera el final
            Sub TableEnd()
		        Response.Write "</div></form>" & Chr(13)
	        End Sub  
 
	        '--- TableStart genera el comienzo de fila
            Sub RowStart(idVal)
		        Response.Write "<div class=gridrow id='" & idVal & "'>" & Chr(13)
	        End Sub

	        '--- TableStart genera el cierre de fila
            Sub RowEnd()
		        Response.Write "</div>" & Chr(13)
	        End Sub	 

	        '--- mainCellStart marca una celda indice
            Sub mainCellStart(cellWidth)
		        Response.Write "<div class='tdOutset middle' width='" & cellWidth & "%'>" & Chr(13)
	        End Sub
    
	        '--- mainCellEnd cierra una celda indice
            Sub mainCellEnd()
		        Response.Write "</div>" & Chr(13)
	        End Sub

	        '--- CellStart abre una celda normal
            Sub CellStart()
		        Response.Write "<div class='tdplain middle'>" & Chr(13)
	        End Sub

	        '--- CellEnd cierra una celda normal
            Sub CellEnd()
		        Response.Write "</div>" & Chr(13)
	        End Sub
        End Class
%>
</div></div></form>