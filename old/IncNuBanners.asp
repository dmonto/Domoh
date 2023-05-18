<%
	rst.Open "SELECT * FROM Contenidos WHERE pagina='Aconsejamos' ORDER BY posicion", sProvider
	while not rst.Eof
%>
<div>
    <div>
        <%=rst("texto")%>
<%		if right(rst("banner"),3) = "oldswf" then %>
		    <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0' id=banner>
			    <param name=movie value='<%=rst("banner")%>'/>
			    <param name=quality value=high />
			    <param name=bgcolor value=#FFFFFF />
			    <embed src='<%=rst("banner")%>' type=application/x-shockwave-flash pluginspace=http://www.macromedia.com/go/getflashplayer />
		        </object>
<%		elseif right(rst("banner"),3) = "swf" then %>
            <img alt='Banner' src="<%=Replace(rst("banner"),".","")&".jpg"%>"/>
<%		elseif rst("banner")<>"" then %>
		    <img alt='Banner' src="<%=rst("banner")%>"/>
<%		end if %>
		</div></div>
<%
		rst.Movenext
	wend
	rst.Close
%>
