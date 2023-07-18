<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	if Request("id")="" and Request("op")="" then 
        Response.Redirect "index.htm"
    elseif Request("id")<>"" then
	    if Request("idioma")<>"" then Session("Idioma")=Request("idioma")
 	    sQuery = "UPDATE Anuncios SET fechaultimamodificacion=GETDATE(), fechaavisobaja=NULL WHERE id=" & Request("id")
	    rst.Open sQuery, sProvider
 	    sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
	    rst.Open sQuery, sProvider
	    Mail "hector@domoh.com", "Anuncio Renovado", " Se ha renovado el anuncio (" & rst("cabecera") & ")"
	    rst.Close
    end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<div class=container>
	<div class=logo><!--#include file="IncNuSubMenu.asp"--></div>
    <div><h1><%=MesgS("Tu anuncio ha sido renovado.","Your advert has been renewed.")%></h1></div>
	<div><h2><a title='<%=MesgS("Página Inicial de Domoh","Domoh''s home page")%>' href=index.htm><%=MesgS("Pulsa aquí para entrar en domoh","Click here to enter domoh")%>.</a></h2></div></div>
<!-- #include file="IncPie.asp" -->
