<!-- #include file="IncNuMail.asp" -->
<% 
	'------ NO BORRAR Ref.asp - Para links desde otras webs
	Mail "diego@domoh.com", "Referencia de Web", "Acceso a Domoh desde " & Request("from")
	Response.Redirect "index.htm"
%>