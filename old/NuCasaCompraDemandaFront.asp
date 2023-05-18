<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.cabecera, "<%=MesgS("Cabecera","Heading")%>")) return false;
	    if(!check(document.frm.anuncio, "<%=MesgS("Cómo eres y qué buscas","What are you looking for?")%>")) return false;
	    if(!check(document.frm.ciudad, "Ciudad/Zona")) return false;
<%  if Request("id")="" then %>
	    if(!check(document.frm.nombre, "Nombre")) return false;	if(!check(document.frm.email, "E-mail")) return false;
	    if(!check(document.frm.usuario, "Nombre de Usuario")) return false;	if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if(!checkSelect(document.frm.fuente, "Cómo nos Conociste")) return false;
<%  end if %>
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
			sQuery=" SELECT * FROM Anuncios a WHERE a.id="& Request("id")
			rst.Open sQuery, sProvider
			sBody= Session("Usuario") & " ha desactivado el anuncio (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Comprador Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=Anuncio+Ocultado."
		elseif Request("op")="Delete" then
		 	sQuery = "DELETE FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "DELETE FROM Inquilinos WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha borrado el anuncio (" & Request("id") &")"
			Mail "hector@domoh.com", "Comprador Borrado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Anuncio+Borrado.","Advert+Deleted.")
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sQuery=" SELECT * FROM Anuncios a WHERE a.id="& Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el anuncio (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Comprador Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=Anuncio+Activado."
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
<header>
	<nav><!--#include file="IncNuSubMenu.asp"--></nav>
	<div class="row">
<%	if Request("op")="Destacar" then %>
        <!--#include file="IncTrDestacado.asp"-->
<%	
		Response.End
	end if 
%>
        Publicación de anuncio de Búsqueda para Comprar</div></header>
<div>
    <p>
        Al publicar este anuncio estás aceptando estas <a title='Ver las condiciones de uso' href=CondUso.htm target=_blank>Condiciones Generales de Uso</a> y nuestra 
		<a title='Ver la Politica de Protección de Datos' href=ProtDatos.asp target=_blank>Política de Protección de Datos</a></p></div>
<div class=seccion>* son campos obligatorios</div>
<div class=seccion>* Cabecera del Anuncio: <input title='Este es el titular que verán los compradores' type=text name=cabecera size=50 value='<%=vcabecera%>'/></div>
<div><p>¿Cómo es la vivienda que buscas?</p></div>
<div>
    <div class=seccion>
	    * ¿Cómo te gustaría que fuera y cómo eres?: <textarea title='Texto del anuncio' cols=35 rows=7 name=anuncio><%=vanuncio%></textarea>
		<input title='Pulsa si el anuncio está en Español' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Texto en Español
		<input title='Pulsa si el anuncio está en Inglés' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Texto en Inglés</div>
    <div class=seccion>
        * Ciudad/Zona que prefieres: <input title='Nombre de la ciudad' type=text name=ciudad size=25 value="<%=vciudad%>"/>
        Máximo precio que pagarías <input title='Pon el precio en Euros' size=5 name=maximo value='<%=vmaximo%>'/>  mil €</div></div>
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
</form>
<!-- #include file="IncPie.asp" -->
</body>