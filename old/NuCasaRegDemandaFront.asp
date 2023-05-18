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
			Response.Redirect "NuCasaInquilino" & Session("Idioma") & ".asp?msg=" & MesgS("Anuncio+Ocultado.","Advert+Hidden")
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
			Response.Redirect "NuCasaInquilino.asp?msg=Anuncio+Activado."
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
			if UCase(rst("mascota"))="ON" then vmascota=" checked"
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes=""
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen="" else vidiomaen=" checked"
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
	    if(!check(document.frm.cabecera, "Cabecera")) return false; if(!check(document.frm.anuncio, "Cómo eres y qué buscas")) return false;
	    if(!check(document.frm.ciudad, "Ciudad/Zona")) return false; if(!check(document.frm.maximo, "Máximo Alquiler")) return false;
<%  if Request("id")="" then %>
	    if(!check(document.frm.nombre, "Nombre")) return false; if(!check(document.frm.email, "E-mail")) return false;
	    if(!check(document.frm.usuario, "Nombre de Usuario")) return false;	if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if(!checkSelect(document.frm.fuente, "Cómo nos Conociste")) return false;
<%  end if %>
   	    return true;
    }
</script>
<body onload="window.parent.location.hash='top';"> 
<form method=post name=frm onsubmit="return checkForm();" action=TrCasaRegDemanda.asp enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
<% if Request("msg")<>"" then %>
    <h1><%=Request("msg")%></h1>
<% end if %>
    <div><!--#include file="IncNuSubMenu.asp"--></div>
<%	if Request("op")="Destacar" then %>
    <!--#include file="IncTrDestacado.asp"-->
<%	
		Response.End
	end if 
%>
	<div><p>Publicación de Anuncio de Búsqueda para Alquilar</p></div>
	<div>
		<p>
            Al publicar este anuncio estás aceptando estas <a title='Ver condiciones' href=CondUso.htm target=_blank>Condiciones Generales de Uso</a> y nuestra 
		    <a title='Leer Política de Protección de Datos' href=ProtDatos.asp target=_blank>Política de Protección de Datos</a></p></div>
	<div class=seccion>*  son campos obligatorios</div>
	<div class=seccion>* Cabecera del Anuncio: <input title='Título del anuncio que verán los usuarios' type=text name=cabecera size=50 value='<%=vcabecera%>'/></div>
	<div><p>¿Cómo es la vivienda que buscas?</p></div>
	<div>¿Qué buscas?</div>
	<div>
		<input title='Pulsa aqui si no quieres compartir' type=radio value='Alquiler Piso' name=tipoviv <% if vtipoviv="Alquiler Piso" then Response.Write " checked"%>/>   Un piso completo 
		<input title='Pulsa aqui para compartir casa' type=radio value='Alquiler Habitación' name=tipoviv <% if vtipoviv="Alquiler Habitación" then Response.Write " checked"%>/>   Una habitación en casa compartida 
		<input title='Pulsa aqui si te da igual' type=radio value=Alquiler name=tipoviv <% if vtipoviv="Alquiler" then Response.Write " checked"%>/>   Me da igual/Otros</div>
	<div class=seccion>
		* ¿Cómo te gustaría que fuera y cómo eres?: <textarea title='Texto del anuncio' cols=35 rows=7 name=anuncio><%=vanuncio%></textarea>
		<input title='Activa si has escrito el anuncio en español' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Texto en Español
		<input title='Activa si has escrito el anuncio en inglés' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Texto en Inglés
		Deja aquí tu foto <input title='Pulsa en el botón de la derecha para buscar el fichero' type=file name=foto /></div>
	<div class=seccion>
		* Ciudad que prefieres: <input title='Ciudad/Pueblo/Zona' type=text name=ciudad size=25 value="<%=vciudad%>"/>* Máximo alquiler que pagarías
		<input title='Máximo precio mensual en euros que pagarías' size=5 name=maximo value="<%=vmaximo%>"/>  €/Mes
		<input title='Activa si Fumas' type=checkbox name=fuma <%=vfuma%> value=ON />Fumas 
		<input title='Activa si tienes perro, gato...' type=checkbox name=mascota <%=vmascota%> value=ON />Tienes Mascota</div>
<% 
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
	<div><p><input title=Grabar type=submit value='Publicar Anuncio' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
	<div><p><input title=Grabar type=submit value='Añadir Anuncio' class=btnForm /></p></div>
<%	else %>
	<div><p><input title=Grabar type=submit value='Modificar Anuncio' class=btnForm /></p></div>
<%	end if %>
</div></form>
<!-- #include file="IncPie.asp" -->
</body>
