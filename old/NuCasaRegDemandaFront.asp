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
	    if(!check(document.frm.cabecera, "Cabecera")) return false; if(!check(document.frm.anuncio, "C�mo eres y qu� buscas")) return false;
	    if(!check(document.frm.ciudad, "Ciudad/Zona")) return false; if(!check(document.frm.maximo, "M�ximo Alquiler")) return false;
<%  if Request("id")="" then %>
	    if(!check(document.frm.nombre, "Nombre")) return false; if(!check(document.frm.email, "E-mail")) return false;
	    if(!check(document.frm.usuario, "Nombre de Usuario")) return false;	if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if(!checkSelect(document.frm.fuente, "C�mo nos Conociste")) return false;
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
	<div><p>Publicaci�n de Anuncio de B�squeda para Alquilar</p></div>
	<div>
		<p>
            Al publicar este anuncio est�s aceptando estas <a title='Ver condiciones' href=CondUso.htm target=_blank>Condiciones Generales de Uso</a> y nuestra 
		    <a title='Leer Pol�tica de Protecci�n de Datos' href=ProtDatos.asp target=_blank>Pol�tica de Protecci�n de Datos</a></p></div>
	<div class=seccion>*  son campos obligatorios</div>
	<div class=seccion>* Cabecera del Anuncio: <input title='T�tulo del anuncio que ver�n los usuarios' type=text name=cabecera size=50 value='<%=vcabecera%>'/></div>
	<div><p>�C�mo es la vivienda que buscas?</p></div>
	<div>�Qu� buscas?</div>
	<div>
		<input title='Pulsa aqui si no quieres compartir' type=radio value='Alquiler Piso' name=tipoviv <% if vtipoviv="Alquiler Piso" then Response.Write " checked"%>/>   Un piso completo 
		<input title='Pulsa aqui para compartir casa' type=radio value='Alquiler Habitaci�n' name=tipoviv <% if vtipoviv="Alquiler Habitaci�n" then Response.Write " checked"%>/>   Una habitaci�n en casa compartida 
		<input title='Pulsa aqui si te da igual' type=radio value=Alquiler name=tipoviv <% if vtipoviv="Alquiler" then Response.Write " checked"%>/>   Me da igual/Otros</div>
	<div class=seccion>
		* �C�mo te gustar�a que fuera y c�mo eres?: <textarea title='Texto del anuncio' cols=35 rows=7 name=anuncio><%=vanuncio%></textarea>
		<input title='Activa si has escrito el anuncio en espa�ol' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Texto en Espa�ol
		<input title='Activa si has escrito el anuncio en ingl�s' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Texto en Ingl�s
		Deja aqu� tu foto <input title='Pulsa en el bot�n de la derecha para buscar el fichero' type=file name=foto /></div>
	<div class=seccion>
		* Ciudad que prefieres: <input title='Ciudad/Pueblo/Zona' type=text name=ciudad size=25 value="<%=vciudad%>"/>* M�ximo alquiler que pagar�as
		<input title='M�ximo precio mensual en euros que pagar�as' size=5 name=maximo value="<%=vmaximo%>"/>  �/Mes
		<input title='Activa si Fumas' type=checkbox name=fuma <%=vfuma%> value=ON />Fumas 
		<input title='Activa si tienes perro, gato...' type=checkbox name=mascota <%=vmascota%> value=ON />Tienes Mascota</div>
<% 
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
	<div><p><input title=Grabar type=submit value='Publicar Anuncio' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
	<div><p><input title=Grabar type=submit value='A�adir Anuncio' class=btnForm /></p></div>
<%	else %>
	<div><p><input title=Grabar type=submit value='Modificar Anuncio' class=btnForm /></p></div>
<%	end if %>
</div></form>
<!-- #include file="IncPie.asp" -->
</body>
