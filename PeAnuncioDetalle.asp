<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncQuDetalles.asp" -->
<html>
<head>
	<title><%=MesgS("Detalles de Anuncio","Advert Detail")%></title>
    <meta http-equiv=Content-Type content="text/html;charset=iso-8859-1">
	<meta content='Vacaciones, Alquiler, Alquileres para gay, gays, lesbiana, lesbianas de pisos y habitaciones en casas compartidas' name=description>
	<meta content='Vacaciones, Alquiler, Alquileres, gay, gays, lesbiana, lesbianas, pisos, habitaciones, casas compartidas, madrid, barcelona' name=keywords>		
	<link href=TrDomoh.css rel=stylesheet type=text/css>
	<style type=text/css>body {margin-left:10px;margin-top:15px;margin-right:10px;margin-bottom:5px;}</style>
	<script type=text/javascript src=forms<%=Session("Idioma")%>.js></script>
</head>
<body>
	<div class=contmanuncio>
<%
	dim numVisitas, sBody, rId
	on error goto 0
	
	if Request("op")="Test" then
		rId="12669"
	else
		rId=Request("id")
	end if

	if Request("tab")="Fotos" then
		rst.Open "SELECT * FROM Fotos where Piso=" & Request("id"), sProvider
%>
		<div class=manfoto>
<%		while not rst.Eof	 %>
			<img alt='Foto' src='http://domoh.com/mini<%=rst("foto")%>'/>
<%
			rst.MoveNext
		wend
		rst.Close
%>
		</div>
<%
		Response.End
	end if
	
	if Request("tabla")="Pisos" or Request("tabla")="Vacas" or Request("op")="Test" then
		Response.Write CasaDetalles(Session("Idioma"), Request("id"), "Inquilino", "", 0)
	elseif Request("tabla")="Inquilinos" then
		Response.Write InquilinoDetalles(Session("Idioma"), Request("id"))
	end if

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
%>
	<div class=mancomment>
<%
	sQuery="SELECT * FROM AnuncioComent WHERE anuncio=" & Request("id")
	rst.Open sQuery, sProvider
	if not rst.Eof then
%>
    <%=MesgS("Comentarios de gente","What other people think")%>:
<%		do while not rst.Eof %>
    <%=rst("valoracion")%> - <%=rst("texto")%>. <%=MesgS("Por","By")%> <%=rst("usuario")%> <%=MesgS("el ","on ")%><%=rst("fechaalta")%>
<%
			rst.Movenext
		loop
	end if	
	rst.Close
	sQuery="SELECT SUM(visitas) AS numVisitas FROM AnunciosVistos WHERE anuncio=" & Request("id")
	rst.Open sQuery, sProvider
	numVisitas=rst("numVisitas")
	rst.Close

	if numVisitas=10 then
		sQuery="SELECT u.email AS uEmail, u.usuario AS uUsuario, a.foto AS aFoto, * FROM usuarios u INNER JOIN Anuncios a ON u.usuario=a.usuario WHERE a.id="& Request("id")
		rst.Open sQuery, sProvider
		if rst("aFoto")="" then
			if Session("Idioma")="" then
				sBody="Estimad@ " & rst("nombre") & ", Tu anuncio (" & rst("cabecera") & ") está publicado en domoh.com y ya ha sido visitado por " & numVisitas & " personas. "
				sBody=sBody & "Muchos de nuestros usuarios no contactan si no ven fotos. Si quieres que te publiquemos fotos, contesta a este e-mail y adjuntalas."
                sBody=sBody & "O también puedes publicarlas tú mismo pulsando "
				sBody=sBody & "<a title='Home de Domoh' href=http://domoh.com>aquí</a>, y entrando en tu perfil con tu usuario " & rst("uUsuario") & ". "
				sBody=sBody & "Si no recuerdas tu clave pulsa <a title='Enviar clave por e-mail' href='http://domoh.com/NuOlvido.asp?email=" & rst("uEmail") & "'>aqui</a>."
				Mail rst("uEmail"), "Fotos para tu Anuncio", sBody
			else
				sBody="Dear " & rst("nombre") & ", Your advert (" & rst("cabecera") & ") is posted on domoh.com."
				sBody=sBody & "Your advert has been already visited by " & numVisitas & " people. Many users don't contact if they can't see pictures. "
				sBody=sBody & "If you want us to publish some pictures, please reply to this e-mail attaching the files. You can also upload them yourself by clicking "
				sBody=sBody & "<a title='Domoh Home' href=http://domoh.com>here</a>, and logging in with your username " & rst("uUsuario") & ". "
				sBody=sBody & "If you cannot remember your password just click <a href='http://domoh.com/NuOlvido.asp?idioma=En&email=" & rst("uEmail") & "'>here</a>."
				Mail rst("uEmail"), "Ad Pictures", sBody
			end if
		end if
		rst.Close
	end if
	
	if Request("admin")="si" then
%>
        <%=numVisitas%> <%=MesgS("personas han entrado a ver tu anuncio.","people have checked your advert.")%>
<% 
		Response.Write MesgS("Si quieres poner un enlace a este anuncio desde otra web, copia y pega el siguiente texto:", "If you want to link your advert to any other web, copy and paste the following text:")%>
		http://domoh.com/Piso.asp?id=<%=Request("id")%>
<% 
    else
        if Request("coment")<>"si" then
%>
        <form name=frm action=TrAnuncioComent.asp method=post>
		    <input type=hidden value=<%=Request("id")%> name=id />
        <% Response.Write MesgS("¿Has llamado o has quedado ya? Añade tus comentarios", "Already called? Please give us feedback")%>:
        <textarea title='<% Response.Write MesgS("Tus comentarios servirán a los demás usuarios como guía", "Your comments will be very useful to others")%>' rows=3 cols=60 name=comentario></textarea>
        <%=MesgS("Valoración","Rating")%> <input title='Pulsa para dar tu opinión' value=1 name=valoracion type=radio />1<input title='Pulsa para dar tu opinión' value=2 name=valoracion type=radio />2
        <input title='Pulsa para dar tu opinión' checked=checked value=3 name=valoracion type=radio />3<input title='Pulsa para dar tu opinión' value=4 name=valoracion type=radio />4
        <input title='Pulsa para dar tu opinión' value=5 name=valoracion type=radio />5
        <input title='<%=MesgS("Pulsa para grabarlo","Press to save")%>' type=submit value='<%=MesgS("Añadir Comentario","Add Comment")%>' class=btnForm /></form>
<% 
        end if
    	if Request("voto")="si" then 
%>
	<%=MesgS("¡Gracias por tu opinión!","Thanks for your opinion!")%>
<%      elseif Request.Cookies("Voto" & Request("id"))<>"" then %>
	<%=MesgS("Este anuncio lo has votado como","You voted this advert as")%> 
<%		
		    dim sVoto
            sVoto=Request.Cookies("Voto"&request("id"))
		    if sVoto="NumBueno" then
			    Response.Write MesgS("""Mola""","""Cool""")
		    elseif sVoto="NumRegular" then
			    Response.Write """OK"""
		    else
			    Response.Write MesgS("""Confuso/Erróneo""","""Misleading""")
		    end if
	    else 
%>
        <%=MesgS("¿Que te parece el anuncio?","What do you think of the ad?")%>
		<input title='<%=MesgS("Pulsa para dar tu opinión","Click to vote")%>' onclick="detalle('TrCasaVoto.asp?voto=NumBueno&id=<%=Request("id")%>')" value=Bueno name=voto type=radio /> <%=MesgS("¡Mola!","Cool!")%>
        <input title='<%=MesgS("Pulsa para dar tu opinión","Click to vote")%>' onclick="detalle('TrCasaVoto.asp?voto=NumRegular&id=<%=Request("id")%>')" value=Regular name=voto type=radio /> OK
		<input title='<%=MesgS("Pulsa para dar tu opinión","Click to vote")%>' onclick="detalle('TrCasaVoto.asp?voto=NumMalo&id=<%=Request("id")%>')" value=Malo name=voto type=radio /> 
        <%=MesgS("Confuso/Erróneo","Misleading")%>
<%	
        end if 		
    end if 
%>
</div></body>
</html>
