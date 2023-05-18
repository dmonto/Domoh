<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<html>
<head>
	<title><%=MesgS("Detalles de Anuncio","Advert Detail")%></title>
    <link rel=stylesheet type=text/css href=TrDomoh.css>
	<script type=text/javascript src=forms<%=Session("Idioma")%>.js></script></head>
<body>
	<div class=container>
		<div class=main>
			<form name=frm action=TrAnuncioComent.asp method=post class='contanuncio'>
<%
    if Request("tabla")="Pisos" then
		Response.Write CasaDetalles(Session("Idioma"), Request("id"), "Inquilino", "ON")
	elseif Request("tabla")="Inquilinos" then
		Response.Write InquilinoDetalles(Session("Idioma"), Request("id"))
	else
		Response.Write "Tipo de anuncio obsoleto"
	end if

    if Request("voto")="si" then 
%>
	¡Gracias por tu opinión!
<%      elseif Request.Cookies("Voto" & Request("id"))<>"" then %>
	Este anuncio lo has votado como 
<%		
		dim sVoto
		sVoto=Request.Cookies("Voto" & Request("id"))
		if sVoto="NumBueno" then
			Response.Write """Mola"""
		elseif sVoto="NumRegular" then
			Response.Write """OK"""
		else
			Response.Write """Confuso/Erróneo"""
		end if
	elseif Request("admin")<>"si" then
%>
	<div class=anvoto>
		<p><%=MesgS("¿Que te parece el anuncio?","What do you think of the advert?")%>
		<input title='Pulsa para dar tu opinión' onclick="detalle('TrCasaVoto.asp?voto=NumBueno&id=<%=Request("id")%>&tabla=<%=Request("tabla")%>')" type=radio value=Bueno name=voto /> ¡Mola!
		<input title='Pulsa para dar tu opinión' onclick="detalle('TrCasaVoto.asp?voto=NumRegular&id=<%=Request("id")%>&tabla=<%=Request("tabla")%>')" type=radio value=Regular name=voto /> OK
		<input title='Pulsa para dar tu opinión' onclick="detalle('TrCasaVoto.asp?voto=NumMalo&id=<%=Request("id")%>&tabla=<%=Request("tabla")%>')" type=radio value=Malo name=voto />Confuso/Erróneo</p>
<% 
		end if 
	
		sQuery="SELECT * FROM AnuncioComent WHERE anuncio=" & Request("id")
		rst.Open sQuery, sProvider
		if not rst.Eof then
%>
	    Comentarios de gente que ha estado allí:
<%		    do while not rst.Eof %>
	    <%=rst("valoracion")%> - <%=rst("texto")%>. Por <%=rst("usuario")%> el <%=rst("fechaalta")%>
<%
				rst.Movenext
			loop
		end if	

		if Request("coment")<>"si" then
%>

	        <input type=hidden value=<%=Request("id")%> name=id />
        <p>Añade tus comentarios: <textarea title='Tus comentarios servirán a los demás usuarios como guía' rows=3 cols=60 name=comentario></textarea></p>
        <%=MesgS("Valoración","Rating")%><input title='Pulsa para dar tu opinión' value=1 name=valoracion type=radio />1<input title='Pulsa para dar tu opinión' value=2 name=valoracion type=radio />2
        <input title='Pulsa para dar tu opinión' checked=checked value=3 name=valoracion type=radio />3<input title='Pulsa para dar tu opinión' value=4 name=valoracion type=radio />4
        <input title='Pulsa para dar tu opinión' value=5 name=valoracion type=radio />5<input title='Pulsa para grabarlo' type=submit value='Añadir Comentario' />
	    
<% 
		end if 
		rst.Close

		if Request("admin")="si" then
			sQuery="SELECT SUM(visitas) AS numVisitas FROM AnunciosVistos WHERE anuncio=" & Request("Id")
			rst.Open sQuery, sProvider
%>
		</div>
    <div class=anentradas>
         <%=rst("numVisitas")%> <%=MesgS("personas han entrado a ver tu anuncio.","people have checked your advert.")%>
         <% Response.Write MesgS("Si quieres poner un enlace a este anuncio desde otra web, copia y pega el siguiente texto:", "If you want to link your advert to any other web, copy and paste the following text:")%>
         http://domoh.com/Piso.asp?id=<%=Request("id")%>
<% 
		else
			sQuery="SELECT * FROM AnunciosVistos WHERE anuncio=" & Request("id") & " AND fecha='" & Format(Date) & "'"
			rst.Open sQuery, sProvider
			
			if rst.Eof then
				rst.Close
				sQuery = "INSERT INTO AnunciosVistos (anuncio, fecha, visitas) VALUES (" & Request("id") & ", '" & Format(Date) & "', 1)"
			else
				rst.Close
				sQuery = "UPDATE AnunciosVistos SET visitas=visitas+1 WHERE anuncio=" & Request("id") & " AND fecha='" & Format(Date) & "'"
			end if 
			rst.Open sQuery, sProvider	
		end if 
%>
	</form></div></div>
<script type=text/javascript>
    var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
    document.write(unescape("%3Cscript async src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type=text/javascript>try {var pageTracker = _gat._getTracker("UA-3713908-1"); pageTracker._trackPageview();} catch (err) { }</script></div></body>
</html>
