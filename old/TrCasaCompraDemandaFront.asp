<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javascript>
    function checkForm() {
	if(!check(document.frm.cabecera, "<%=MesgS("Cabecera","Heading")%>")) return false;
	if(!check(document.frm.anuncio, "<%=MesgS("Cómo eres y qué buscas","What are you looking for?")%>")) return false;
	if(!check(document.frm.ciudad, "<%=MesgS("Ciudad/Zona","City/Area")%>")) return false;
	if(!check(document.frm.maximo, "<%=MesgS("Máximo Precio","Top Price")%>")) return false;
<% if Request("id")="" then %>
	if(!check(document.frm.nombre, "<%=MesgS("Nombre","Name")%>")) return false;
	if(!check(document.frm.email, "E-mail")) return false;
	if(!check(document.frm.usuario, "<%=MesgS("Nombre de Usuario","User Name")%>")) return false;
 	if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	if(!checkSelect(document.frm.fuente, "<%=MesgS("Cómo nos Conociste","How did you get to Domoh")%>")) return false;
<% end if %>
   	return true;
}
</script>
<%
	dim vcabecera, vanuncio, vfuente, vprovincia, vciudad, vmaximo, vidiomaes, vidiomaen, sBody
	
	vprovincia=0
	vidiomaes=" checked"
	vidiomaen=""

	if Request("id")<>"" then
		if Request("op")="Borrar" then
		 	sQuery = "UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody= Session("Usuario") & " ha desactivado el anuncio (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Comprador Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Anuncio+Ocultado.","Advert+Hidden.")
		elseif Request("op")="Delete" then
		 	sQuery = "DELETE FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "DELETE FROM Inquilinos WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha borrado el anuncio (" & Request("id") &")"
			Mail "hector@domoh.com", "Comprador Borrado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Anuncio+Borrado.","Advert+Deleted.")
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el anuncio (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Comprador Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Anuncio+Activado.","Advert+Reposted.")
		elseif Request("op")="Modificar" then
			sQuery = "SELECT * FROM Inquilinos i INNER JOIN Anuncios a ON i.id=a.id WHERE a.id=" & Request("id")
			rst.Open sQuery, sProvider
			vcabecera=rst("cabecera")
			vanuncio=rst("descripcion")
			vprovincia=rst("provincia")
			vciudad=rst("ciudad")
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes=""
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen="" else vidiomaen=" checked"
			vmaximo=rst("maximo")
			rst.Close
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrCasaCompraDemanda.asp?op=FotoBorrada"
		end if
	end if
%>
<body onload="window.parent.location.hash='top';"> 
<form method=post name=frm onsubmit='return checkForm();' action=TrCasaCompraDemanda.asp>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<div class=topmenu>
		<a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a>
		<a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a><%=MesgS("Publicar Anuncio de Comprador","Post New Buyer")%></div>
	<h1 class=tituSec><% Response.Write MesgS("Publicación de anuncio de Búsqueda para Comprar","Buy Seeker Advert Publication")%></h1>
    <div>
	    <p>
            <% Response.Write MesgS("Al publicar este anuncio estás aceptando estas ", "By posting this advert you are agreeing to our ")%>
			<a title='<% Response.Write MesgS("Ver las condiciones de uso", "Read Terms & Conditions")%>' href=CondUso<%=Session("Idioma")%>.htm target=_blank>
			<%=MesgS("Condiciones Generales de Uso","Terms & Conditions")%></a></p></div>
<div class=seccion>*<%=MesgS("  son campos obligatorios","  are required fields")%></div>
<div class=seccion>
    * <%=MesgS("Cabecera del Anuncio","Advert Title")%>:
    <input title='<% Response.Write MesgS("Este es el titular que verán los compradores", "The first thing that users will see")%>' type=text name=cabecera size=50 value='<%=vcabecera%>'/></div>
<div><p><% Response.Write MesgS("¿Cómo es la vivienda que buscas?", "What are you looking for?")%></p></div>
<div>
    <div class=seccion>
	    * <% Response.Write MesgS("¿Cómo te gustaria que fuera?", "How would you like it to be?")%>: <textarea title='<%=MesgS("Texto del anuncio","Text of the Advert")%>' cols=35 rows=7 name=anuncio><%=vanuncio%></textarea>
        <input title='<% Response.Write MesgS("Pulsa si el anuncio está en Español","Check if the text is in Spanish")%>' type=checkbox name=idiomaes <%=vidiomaes%> value=ON/><%=MesgS("Texto en Español","Text in Spanish")%>
        <input title='<% Response.Write MesgS("Pulsa si el anuncio está en Inglés","Check if the text is in English")%>' type=checkbox name=idiomaen <%=vidiomaen%> value=ON/><%=MesgS("Texto en Inglés","Text in English")%></div>
    <div class="seccion">
	    * <% Response.Write MesgS("Ciudad/Zona que prefieres", "City/Area where you are looking")%>:
        <input title='<%=MesgS("Nombre de la ciudad","Name of the City")%>' type=text name=ciudad size=25 value='<%=vciudad%>' >
        * <%=MesgS("Máximo precio que pagarías","Highest price you would pay")%>
		<input title='<%=MesgS("Por el precio en Euros","Price in Euros")%>' size=5 name=maximo value='<%=vmaximo%>'/><%=MesgS("mil ","thousand ")%>€</div></div>
<% 
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
<div><p><input title='<%=MesgS("Grabar","Save")%>' type=submit value='<%=MesgS("Publicar Anuncio","Post Advert")%>' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
<div><p><input title='<%=MesgS("Grabar","Save")%>' type=submit value='<%=MesgS("Añadir Anuncio","Save Advert")%>' class=btnForm /></p></div>
<%	else %>
<div><p><input title='<%=MesgS("Grabar","Save")%>' type=submit value='<%=MesgS("Modificar Anuncio","Update Advert")%>' class=btnForm /></p></div>
<%	end if %>
    </div>
</form><!-- #include file="IncPie.asp" -->
</body>