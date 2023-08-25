<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncQuDetalles.asp" -->
<html>
<head>
	<title>Domoh - <%=MesgS("Detalles de Anuncio","Advert Detail")%></title>
    <meta http-equiv=Content-Type content='text/html;charset=iso-8859-1'>
	<meta content='Vacaciones, Alquiler, Alquileres para gay, gays, lesbiana, lesbianas de pisos y habitaciones en casas compartidas' name=description>
	<meta content='Vacaciones, Alquiler, Alquileres, gay, gays, lesbiana, lesbianas, pisos, habitaciones, casas compartidas, madrid, barcelona' name=keywords>		
	<link href=TrDomoh.css rel=stylesheet type=text/css>
	<style type=text/css>body {margin-left:10px;margin-top:15px;margin-right:10px;margin-bottom:5px;}</style>
	<script type=text/javascript src=forms<%=Session("Idioma")%>.js></script>
	<base target=principal />
</head>
<body>
	<div class="contanuncio">
<%
	dim numVisitas, sBody, idAnuncio
	on error goto 0
	
	idAnuncio = Request("id")
	if idAnuncio = "" then idAnuncio = "12559"

	rst.Open "SELECT * FROM Anuncios WHERE id=" & idAnuncio, sProvider
	if rst.Eof then 
		Mail "diego@domoh.com", "Anuncio Erróneo", "Alguien intentó acceder al Anuncio " & idAnuncio
		Response.Write "El Anuncio ya no está disponible"
		Response.End
	end if
	rst.Close
	
	if Request("tabla")="Pisos" then
		Response.Write CasaDetalles(Session("Idioma"), idAnuncio, "Inquilino", "")
	elseif Request("tabla")="Inquilinos" then
		Response.Write InquilinoDetalles(Session("Idioma"), idAnuncio)
	end if

	sQuery="SELECT * FROM AnunciosVistos WHERE anuncio=" & idAnuncio & " AND fecha='" & Format(Date) & "'"
	rst.Open sQuery, sProvider
	if rst.EOF then
		rst.Close
		sQuery = "INSERT INTO AnunciosVistos (anuncio, fecha, visitas) VALUES (" & idAnuncio & ", '" & Format(Date) & "', 1)"
	else
		rst.Close
		sQuery = "UPDATE AnunciosVistos SET visitas=visitas+1 WHERE anuncio=" & idAnuncio & " AND fecha='" & Format(Date) & "'"
	end if 
	rst.Open sQuery, sProvider

	sQuery="SELECT * FROM AnuncioComent WHERE anuncio=" & idAnuncio
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

	sQuery="SELECT SUM(visitas) AS numVisitas FROM AnunciosVistos WHERE anuncio=" & idAnuncio
	rst.Open sQuery, sProvider
	numVisitas=rst("numVisitas")
	rst.Close

	if numVisitas=10 then
		sQuery="SELECT u.email AS uEmail, u.usuario AS uUsuario, a.foto AS aFoto, * FROM usuarios u INNER JOIN Anuncios a ON u.usuario=a.usuario WHERE a.id="& idAnuncio
		rst.Open sQuery, sProvider
		if rst("aFoto")="" then
			if Session("Idioma")="" then
				sBody="Estimad@ " & rst("nombre") & ","
				sBody=sBody & " Tu anuncio (" & rst("cabecera") & ") está publicado en domoh.com y ya ha sido visitado por " & numVisitas & " personas. Muchos de nuestros usuarios no contactan si no ven fotos. "
				sBody=sBody & "Si quieres que te publiquemos fotos, contesta a este e-mail y adjuntalas. O también puedes publicarlas tú mismo pulsando <a title='Home de Domoh' href=http://domoh.com>aquí</a>, "
				sBody=sBody & "y entrando en tu perfil con tu usuario " & rst("uUsuario") & ". Si no recuerdas tu clave pulsa "
				sBody=sBody & "<a title='Solicitar Clave por e-mail' href='http://domoh.com/NuOlvido.asp?email=" & rst("uEmail") & "'>aqui</a>."
				Mail rst("uEmail"), "Fotos para tu Anuncio", sBody
			else
				sBody="Dear " & rst("nombre") & ", Your advert (" & rst("cabecera") & ") is posted on domoh.com. Your advert has been already visited by " & numVisitas & " people. "
				sBody=sBody & "Many users don't contact if they can't see pictures. If you want us to publish some pictures, please reply to this e-mail attaching the files. You can also upload them yourself by clicking "
				sBody=sBody & "<a title='Home Page' href=http://domoh.com>here</a>, and logging in with your username " & rst("uUsuario") & ". "
				sBody=sBody & "If you cannot remember your password just click <a title='Send password by e-mail' href='http://domoh.com/NuOlvido.asp?idioma=En&email=" & rst("uEmail") & "'>here</a>."
				Mail rst("uEmail"), "Ad Pictures", sBody
			end if
		end if
		rst.Close
	end if
%>
	<div class=anvoto>
<%	
	if Request("admin")="si" then
%>
        <%=numVisitas%> <%=MesgS("personas han entrado a ver tu anuncio.","people have checked your advert.")%>
	    <%=MesgS("Si quieres poner un enlace a este anuncio desde otra web, copia y pega el siguiente texto:", "If you want to link your advert to any other web, copy and paste the following text:")%>
        http://domoh.com/Piso.asp?id=<%=idAnuncio%>	
<% 
	else
        if Request("coment")<>"si" then
%>
        <form name=frm action=TrAnuncioComent.asp method=post>
		<input type=hidden value=<%=Request("id")%> name=id />
		<% Response.Write MesgS("¿Has llamado o has quedado ya? Añade tus comentarios", "Already called? Please give us feedback")%>:
		<textarea title='<% Response.Write MesgS("Tus comentarios servirán a los demás usuarios como guía", "Your comments will be very useful to others")%>' rows=3 cols=60 name=comentario></textarea><br/>
        <%=MesgS("Valoración","Rating")%>
		<input title='Pulsa para dar tu opinión' value=1 name=valoracion type=radio />1<input title='Pulsa para dar tu opinión' value=2 name=valoracion type=radio />2
		<input title='Pulsa para dar tu opinión' checked=checked value=3 name=valoracion type=radio />3<input title='Pulsa para dar tu opinión' value=4 name=valoracion type=radio />4
		<input title='Pulsa para dar tu opinión' value=5 name=valoracion type=radio />5
		<input title='<%=MesgS("Pulsa para grabarlo","Press to save")%>' type=submit value='<%=MesgS("Añadir Comentario","Add Comment")%>' class=btnForm /></form>
<% 
		end if
	    if Request("voto")="si" then 
%>
        <%=MesgS("!Gracias por tu opinión!","Thanks for your opinion!")%>
<%      elseif Request.Cookies("Voto" & idAnuncio)<>"" then %>
        <%=MesgS("Este anuncio lo has votado como","You voted this advert as")%> 
<%		
            dim sVoto
			sVoto=Request.Cookies("Voto" & idAnuncio)
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
	<input title='<%=MesgS("Pulsa para dar tu opinión","Click to vote")%>' onclick="window.location='TrCasaVoto.asp?voto=NumBueno&tabla=<%=Request("tabla")%>&id=<%=idAnuncio%>'" value=Bueno name=voto type=radio /> 
    <%=MesgS("¡Mola!","Cool!")%>
    <input title='<%=MesgS("Pulsa para dar tu opinión","Click to vote")%>' onclick="window.location='TrCasaVoto.asp?voto=NumRegular&tabla=<%=Request("tabla")%>&id=<%=idAnuncio%>'" value=Regular name=voto type=radio /> OK
    <input title='<%=MesgS("Pulsa para dar tu opinión","Click to vote")%>' onclick="window.location='TrCasaVoto.asp?voto=NumMalo&tabla=<%=Request("tabla")%>&id=<%=idAnuncio%>'" value=Malo name=voto type=radio /> 
    <%=MesgS("Confuso/Erróneo","Misleading")%>
<%	
        end if 
    end if 
%>
	</div></div></body>
</html>