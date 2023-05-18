<!-- #include file="IncNuBD.asp" -->
<%
	dim numMes, numAno, sMonth, rstComent, sTitulo, numBlogs, vpage, numPage, Block, Blocks, i, totale, maxpages, PagStart, PagStop, RemainingRecords

	numMes=Request("month")
	numAno=Request("year")
	if numMes="" then numMes=Month(Now)
	sMonth=NameFromMonth(numMes)
	if numAno="" then numAno=Year(Now)

	block=Request("Block")
	numPage=Request("page")
	if Block="" then Block=1 
	if numPage="" then numPage=1
%>
<html>
    <head><title>Blogs de Domoh</title><link href=include/styles.css rel=stylesheet></head>
<body>
<table border=0>
	<tr><td colspan=2><div class=hrgreen><img alt='Blanco' src=images/spacer.gif/></div></td></tr>
	<tr><td colspan=2><u>Editar/Borrar Blog</u> &nbsp;</td></tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr><td colspan=2>&nbsp;</td></tr></table>
<table border=0>
	<tr> 
		<td rowspan=2><img alt=blank src=images/spacer.gif/></td>
		<td>
			<%= Request("msg") %><br/><br/>
<%
	set rstComent = Server.CreateObject("ADODB.Recordset")
	sQuery= "SELECT COUNT(*) AS numBlogs FROM BlogDetalle"
	rst.Open sQuery, sProvider
	numBlogs=rst("numBlogs")
	rst.Close
	sQuery= "SELECT * FROM BlogDetalle ORDER BY blog_id DESC;"
	rst.Open sQuery, sProvider
	rst.PageSize = 20

	if rst.EOF then
		Response.Write "<div align='center'><br><br><br>No hay ningún Blog</div>"
		Response.End
	else	
	    totale = numBlogs
%>
			<div>BLOGS EN LA BASE DE DATOS :  <b><% = totale %></b></div><br/>
	        <table border=0 class=tablemenu>
                <tr> 
                    <td title='Título del Blog'><b>Título</b></td><td title='Autor del Blog'><b>Autor</b></td>
                    <td title='Número de textos'><b>Comentarios</b></td><td title='Borra, Modifica...'><b>Qué hago - Blog</b></td>
                    <td title='Borra, Modifica...'><b>Qué hago - Comentarios</b></td></tr>
<%
	    for i=1 to 20
		    if rst.Eof then exit for
		    sQuery = "SELECT Count(blog_id) AS numerocommenti FROM BlogComent WHERE blog_id = " & CLng(rst("blog_id")) & ";"
            rstComent.Open sQuery, sProvider
%>
                <tr> 
                    <td>
<%
            sTitulo = rst("blog_titolo")
            Response.Write anteprima(sTitulo, 3)
%>
                    </td>
                    <td><% if rst("blog_email") <> "" then Response.Write("<a href='mailto:" & rst("blog_email") & "' >" & rst("blog_autore") & "</a>") else Response.Write(rst("blog_autore")) %></td>
                    <td>
<%          
            if CBool(rst("commenti"))=true then 
                if not rstComent.Eof then 
                    Response.Write rstComent("numerocommenti")
                else
                    Response.Write  "0"
                end if
            else
                Response.Write "No Permitidos"
            end if
%>
                    </td>
                    <td>
<%          if rstComent("numerocommenti")<>"0" then %>
                        <a href='TrAdminBlogComents.asp?blog_id=<%=rst("blog_id")%>'><img src=images/BlogEditComent.gif alt="editar los comentarios del blog '<%=rst("blog_titolo")%>'" border=0/></a>
<%          else %>
					    &nbsp;
<%          end if %>
                        </td>
                    <td>
                        <a href='TrBlogEdit.asp?blog_id=<%=rst("blog_id")%>&page=<%=numPage %>&Block=<%=Block%>'><img src=images/BlogEdit.gif alt='editar blog' border=0/></a>   
                        <a href='TrBlogDelete.asp?blog_id=<%=rst("blog_id")%>&page=<%=numPage%>&Block=<%=Block%>' onclick="return confirm('¿Estás seguro de borrar el blog y todos sus comentarios?');">
                        <img src=images/BlogDelete.gif alt='borrar blog' border=0/></a></td></tr>
<%
		rstComent.Close
		rst.MoveNext
	next
end if
%>
                        </table></td></tr></table>
<br/><br/>
<table border=0><tr><td>
<%
    '--- paginazione
    maxpages = int(totale / 20)
    if (totale mod 20) <> 0 then maxpages = maxpages + 1 
    Blocks=0
    Blocks = int(maxpages / 10)
    if (maxpages mod 10) <> 0 then Blocks = Blocks + 1 
    Response.Write "página " &  numPage &  " de " & maxpages & "</td>"
    Response.Write "<td>"
    PagStop=Block*10
    PagStart=(PagStop-10)+1

    i=0
    if maxpages>1 then 
	    for numPag=PagStart to PagStop
            i=i+1 
            if Block=1 then i=0    
        
            if i=1 and Block>1 then 
                Response.Write "<a href='" & Request.Servervariables("SCRIPT_NAME") & "?Block=" & (Block-1) & "&page=" & numPag-1 & "' title='Páginas Anteriores'>" & _
                    "<img src=immagini/frecciasx.gif border=0><img src=immagini/frecciasx.gif border=0></a>  "
            end if
            
            RemainingRecords = totale-(numPag*intRecordsPerPage)  
            if RemainingRecords > 0 then 
    		    if numPag=CInt(numPage) then
                    Response.Write numPag & " "
		        else
                    Response.Write "<a href='" & Request.Servervariables("SCRIPT_NAME") & "?Block=" & Block & "&page=" & numPag & "'>" & numPag & "</a> "
		        end if
            else 
    		    if numPag=CInt(numPage) then
                    Response.Write numPag & " "
		        else
                    Response.Write "<a href='" & Request.Servervariables("SCRIPT_NAME") & "?Block=" & Block & "&page=" & numPag & "'>" & numPag & "</a> "
		        end if
                exit for
            end if
            
            if numPag=PagStop and Blocks>1 and int(Block-1)<int(Blocks) then 
                Response.Write "  <a href='" & Request.Servervariables("SCRIPT_NAME") & "?Block=" & (Block+1) & "&page=" & numPag+1 &  "' title='Próximas Páginas'>" & _ 
		            "<img src=immagini/freccia.gif border=0><img src=immagini/freccia.gif border=0></a>"
            end if
	    next
    end if
%>
</td></tr></table><br/><br/>	
<% rst.Close %>
<table><tr><td>&nbsp;</td><td>&nbsp;</td></tr>
</table></body></html>