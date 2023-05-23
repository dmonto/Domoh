<!-- #include file="IncNuBD.asp" -->
<%
 	sQuery = "UPDATE Anuncios SET " & Request("voto") & "=" & Request("voto") & "+1 WHERE id=" & Request("id")
	rst.Open sQuery, sProvider
	Response.Cookies("Voto" & Request("id"))=Request("voto")
	Response.Cookies("Voto" & Request("id")).Expires=Now+30
	Response.Redirect "QuAnuncioDetalle.asp?id=" & Request("id") & "&tabla=" & Request("tabla") & "&voto=si"
%>
