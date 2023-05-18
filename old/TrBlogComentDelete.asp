<!-- #include file="IncNuBD.asp" -->
<%
	dim numBlog, numComment, Block, msg

	Response.Buffer = True
	Block=Request("Block")
	numComment=Request("commento_id")
	numBlog=Request("blog_id")

	sQuery = "DELETE FROM BlogComent WHERE commento_id =" & numComment & ""
	rst.Open sQuery, sProvider

	msg = "Comentario borrado"
	Response.Redirect "TrAdminBlogComents.asp?page=" &  Request("page") & "&block=" & Block & "&msg="& msg &"&blog_id=" & numBlog & ""
%>