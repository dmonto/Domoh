<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vcabecera, vanuncio, vfuente, vprovincia, vciudad, vmaximo, vfuma, vmascota, vtipoviv, vidiomaes, vidiomaen, sBody
	
	vfuma=""
	vmascota=""
	vidiomaes=" checked"
	vidiomaen=""

	if Request("id")<>"" then
		if Request("op")="Borrar" then
		 	sQuery = "UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha desactivado el anuncio (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Inquilino Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Anuncio+Ocultado.","Advert+Hidden")
		elseif Request("op")="Delete" then
		 	sQuery = "DELETE FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "DELETE FROM Inquilinos WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha borrado el anuncio (" & Request("id") &")"
			Mail "hector@domoh.com", "Inquilino Borrado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Anuncio+Borrado.","Advert+Deleted.")
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el anuncio (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Inquilino Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Anuncio+Activado.","Advert+Reposted")
		elseif Request("op")="Modificar" then
			sQuery="SELECT * FROM Inquilinos i INNER JOIN Anuncios a ON i.id=a.id WHERE a.id=" & Request("id")
			rst.Open sQuery, sProvider
			vcabecera=rst("cabecera")
			vanuncio=rst("descripcion")
			vprovincia=rst("provincia")
			vciudad=rst("ciudad")
			vmaximo=rst("maximo")
			vtipoviv=rst("tipoviv")
			if vtipoviv="Compra" then Response.Redirect "TrCasaCompraDemandaFront.asp?id=" & Request("id") & "&op=" & Request("op")
			if UCase(rst("fuma"))="ON" then vfuma=" checked"
			if UCase(rst("Mascota"))="ON" then vmascota=" checked"
			if UCase(rst("IdiomaEs"))<>"ON" then vidiomaes=""
			if UCase(rst("IdiomaEn"))<>"ON" then vidiomaen="" else vidiomaen=" checked"
			rst.Close
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrCasaRegDemanda.asp?op=FotoBorrada"
		end if
	end if
%>
<!--#include file="IncTrCabecera.asp"-->
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.cabecera, "<%=MesgS("Cabecera","Header")%>")) return false;
	    if(!check(document.frm.anuncio, "<%=MesgS("Cómo eres y qué buscas","How you are and what you look for")%>")) return false;
	    if(!check(document.frm.ciudad, "<%=MesgS("Ciudad/Zona","City/Area")%>")) return false;
	    if(!check(document.frm.maximo, "<%=MesgS("Máximo Alquiler","Top Rent")%>")) return false;
<%  if Request("id")="" then %>
	    if(!check(document.frm.nombre, "<%=MesgS("Nombre","Name")%>")) return false;
	    if(!check(document.frm.email, "E-mail")) return false;
	    if(!check(document.frm.usuario, "<%=MesgS("Nombre de Usuario","User Name")%>")) return false;
 	    if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if(!checkSelect(document.frm.fuente, "<%=MesgS("Cómo nos Conociste","How you got to Domoh")%>")) return false;
<% end if %>
   	return true;
}
</script>
<body onload="window.parent.location.hash='top';"> 
<form method=post name=frm onsubmit="return checkForm();" action=TrCasaRegDemanda.asp enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
<% if Request("msg")<>"" then %>
	<p class=banner><%=Request("msg")%></p>
<% end if %>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
	<div class=menu><a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a><%=MesgS("Publicar Anuncio de Inquilino","Post New Tenant")%></div></nav>
	<div class=main>
		<h1 class=tituSec><% Response.Write MesgS("Publicación de Anuncio de Búsqueda para Alquilar", "Flat or Room Request for Rent Advert Publication")%></h1>
	    <div class=seccion>
		    <h2>* <%=MesgS("Cabecera del Anuncio","Title of the Advert")%>:</h2>
            <input title='<% Response.Write MesgS("Título del anuncio que verán los usuarios","First thing your users will see")%>' type=text name=cabecera size=50 value='<%=vcabecera%>'/></div>
        <div class="seccion">
            <h2><p><% Response.Write MesgS("¿Cómo es la vivienda que buscas?", "How is the place you are looking for?")%></p></h2></div>
    <div><%=MesgS("¿Qué buscas?","What are looking for?")%></div>
    <div>
	    <input title='<% Response.Write MesgS("Pulsa aqui si no quieres compartir","Check if you do not want to share")%>' 
            type=radio value='Alquiler Piso' name=tipoviv <% if vtipoviv="Alquiler Piso" then Response.Write " checked"%> />  
	    <%=MesgS(" Un piso completo "," A whole apartment ")%>
		<input title='<% Response.Write MesgS("Pulsa aqui para compartir casa", "Check if you want to share")%>' type=radio 
            value='Alquiler Habitación' name=tipoviv <%if vtipoviv="Alquiler Habitación" then Response.Write " checked"%> />
		<% Response.Write MesgS(" Una habitación en casa compartida ", " A room in a shared apartment ")%>
		<input title="<% Response.Write MesgS("Pulsa aqui si te da igual", "Check if you don''t care")%>" type=radio value=Alquiler name=tipoviv <%if vtipoviv="Alquiler" then Response.Write " checked"%> checked />
        <%=MesgS(" Me da igual/Otros "," I don't care/Other ")%></div>
    <div>
	    <div class=seccion>
		    * <% Response.Write MesgS("¿Cómo te gustaría que fuera y cómo eres?", "How would you like it to be and how are you?")%>:
            <textarea title='<%=MesgS("Texto del anuncio","Text of the Advert")%>' cols=45 rows=7 name=anuncio><%=vanuncio%></textarea>
			<input title='<%=MesgS("Activa si has escrito el anuncio en español","Check if the advert is in Spanish")%>' type=checkbox name=idiomaes <%=vidiomaes%> value='ON'/>
            <%=MesgS("Texto en Español","Text in Spanish")%> <input title='<%=MesgS("Activa si has escrito el anuncio en inglés","Check if the advert is in English")%>' type=checkbox name=idiomaen <%=vidiomaen%> value='ON'/>
            <%=MesgS("Texto en Inglés","Text in English")%> <%=MesgS("Deja aquí tu foto ","Leave here your picture ")%>
            <input title='<% Response.Write MesgS("Pulsa en el botón de la derecha para buscar el fichero", "Press the button on the right to find the file with the picture")%>' type=file name=foto /></div>
        <div class=seccion>
            * <%=MesgS("Ciudad que prefieres","City you would like")%>: <input title='<%=MesgS("Ciudad/Pueblo/Zona","City/Town/Area")%>' type=text name=ciudad size=25 value='<%=vciudad%>'/>
			* <%=MesgS("Máximo alquiler que pagarías","Maximum rent you are willing to pay")%>
			<input title='<%=MesgS("Máximo precio mensual en euros que pagarías","Monthly rent in Euros")%>' size=5 name=maximo value='<%=vmaximo%>'/> €/<%=MesgS("Mes","Month")%>
            <input type=checkbox name=fuma <%=vfuma%> value='ON'/><%=MesgS("Fumas ","Smoker?")%>
			<input type=checkbox name=mascota <%=vmascota%> value='ON'/><%=MesgS("Tienes Mascota","Have pets?")%></div></div>
<% 
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
    <div><p><input type=submit value='<%=MesgS("Publicar Anuncio","Post Advert")%>' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
    <div><p><input type=submit value='<%=MesgS("Añadir Anuncio","New Advert")%>' class=btnForm /></p></div>
<%	else %>
    <div><p><input type=submit value='<%=MesgS("Modificar Anuncio","Update Advert")%>' class=btnForm /></p></div>
<%	end if %>
    </div></form></div><!-- #include file="IncPie.asp" -->
</body>
