<%	
	dim RegDesde, RegHasta, sTipo
	
'---- PagCabeza() muestra la cabecera de [numItems] de algo llamado [nomItem]
Sub PagCabeza(vItems, nomItem) 
	dim numResults, sS
		
	if Session("numResults")=0 then Session("numResults")=30
	numResults=Session("numResults")
	if vItems>1 then sS="s" else sS=""

	if vItems=0 then 
%>
	<%=MesgS("Nada nuevo","Nothing new")%>.
<%
		Response.End
	elseif vItems <= numResults then
		RegDesde=1
		RegHasta=vItems
%>
	<%=MesgS("Hay ","There are ")%><%=vItems & " " & nomItem & sS & MesgS(" que te pueden interesar"," which seem interesting for you")%>.
<%
	elseif Request("regdesde") = "" then
		RegDesde=1
		RegHasta=numResults
%>
	<%=MesgS("Hay ","There are ")%><%=vItems & " " & nomItem & sS%>. <%=MesgS("Te ense�o del 1 al ","Showing from 1 to ")%><%=numResults%>.
<%
	else
		RegDesde=CLng(Request("regdesde"))
		if RegDesde<1 then RegDesde=1
		if vItems - RegDesde > numResults then RegHasta=RegDesde+numResults-1 else RegHasta=vItems
%>
	<%=MesgS("Hay ","There are ")%><%=vItems & " " & nomItem & sS%>. <%=MesgS("Te ense�o del ","Showing from ")%><%=RegDesde%><%=MesgS(" al "," to ")%><%=RegHasta%>.
<%	end if 
End Sub
	
'---- PagCabezaEn() muestra la cabecera en ingles de [numItems] de algo llamado [nomItem]
Sub PagCabezaEn(numItems, nomItem) 
		dim numResults, sS
		
		if Session("numResults")=0 then Session("numResults")=20
		numResults=Session("numResults")

		if numItems=0 then 
%>
	Nothing new.
<%
			Response.End
		elseif numItems <= numResults then
			RegDesde=1
			RegHasta=numItems
%>
	We have <%=numItems & " " & nomItem & sS%> available.
<%
		elseif Request("regdesde") = "" then
			RegDesde=1
			RegHasta=numResults
%>
	We have <%=numItems & " " & nomItem & sS%>. Showing from 1 to <%=numResults%>.
<%
		else
			RegDesde=CLng(Request("regdesde"))
			if RegDesde<1 then RegDesde=1
			if numItems - RegDesde > numResults then
				RegHasta=RegDesde+numResults-1
			else
				RegHasta=numItems
			end if	
%>
	We have <%=numItems & " " & nomItem & sS%>. Showing from <%=RegDesde%> to <%=RegHasta%>.
<%		
	end if 
End Sub

'---- PagPie() muestra el pie y flechitas para navegar por [numItems] elementos de [sURL]
Sub PagPie(numItems, sURL) 
	dim i
	
	if Session("numResults")=0 then Session("numResults")=20
	if numItems<=Session("numResults") then Exit Sub
	if InStr(sURL,"Nu") then
%>
	<nav>
<%		
		end if
		if RegDesde=1 then %>
				<img alt='Atr�s' title='Est�s en la primera p�gina' src=images/FlechaPrevioGris.gif />
<%		else %>
				<a title='P�gina Anterior' href='<%=sURL%>RegDesde=<%=RegDesde-Session("numResults")%>'><img alt='P�gina Anterior' src=images/FlechaPrevio.gif /></a>
<%	
		end if
	
		for i=0 to CInt(numItems/Session("numResults"))
			if i*Session("numResults") > numitems then Exit For
			if CInt(RegDesde/Session("numResults")) = i then
				Response.Write i+1
			else
%>
				<a title='Ir a una p�gina' href='<%=sURL%>RegDesde=<%=i*Session("numResults")+1%>'><%=i+1%></a>
<%	
			end if
		next
	
		if RegHasta>=NumItems then
%>
				<img alt='No hay siguiente' title='Est�s en la �ltima p�gina' src=images/FlechaProxGris.gif />
<%		else %>
				<a title='P�gina Siguiente' href='<%=sURL%>RegDesde=<%=RegHasta+1%>'><img alt='P�gina Siguiente' src=images/FlechaProx.gif /></a>
<%		end if %>				
		</nav>
<%	End Sub %>