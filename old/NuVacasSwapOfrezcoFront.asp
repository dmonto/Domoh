<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vtipo, vdir1, vciudadnombre, vdir2, vcp, vzona, vdescripcion, vmascota, vidiomaes, vidiomaen, vfoto, sBody

	vmascota=" checked"
	if Session("Idioma")="En" then vidiomaes="" else vidiomaes=" checked" 
	if Session("Idioma")="En" then vidiomaen=" checked" else vidiomaen="" 
	vfoto=" checked"

	if Request("id")<>"" then
		if Request("op")="Borrar" then
		 	sQuery = "UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Piso Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=Piso+Ocultado."
		elseif Request("op")="Delete" then
		 	sQuery = "DELETE FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "DELETE FROM Pisos WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha borrado el piso (" & Request("id") &")"
			Mail "hector@domoh.com", "Piso Borrado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Borrado.","House+Deleted.")
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el piso (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Piso Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=Piso+Activado."
		elseif Request("op")="Modificar" then
			sQuery="SELECT * FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.id=" & Request("id")
			rst.Open sQuery, sProvider
			vciudadnombre=rst("ciudadnombre")
			vzona=rst("zona")
			if rst("rentaviv") > 0 then vrentaviv=rst("rentaviv")
			vdescripcion=rst("descripcion")
			vdir1=rst("dir1")
			vdir2=rst("dir2")
			vcp=rst("cp")			
			if UCase(rst("mascota"))<>"ON" then vmascota=""
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes=""
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen="" else vidiomaen=" checked"
			if IsNull(rst("Foto")) or rst("Foto")="" then vfoto=""
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrVacasSwapOferta.asp?op=FotoBorrada"
		end if
	end if
%>
<head>
    <!-- #include file="IncTrCabecera.asp" -->
    <script type=text/javascript>
        function checkForm() {
	        if(!check(document.frm.ciudadnombre, "Ciudad")) return false; if(!check(document.frm.zona, "Zona")) return false;
	        if(!check(document.frm.descripcion, "Descripción")) return false;
<%	if Request("id")="" then %>
	        if(!check(document.frm.nombre, "Nombre")) return false;	if(!check(document.frm.email, "E-mail")) return false;
	        if(!check(document.frm.usuario, "Nombre de Usuario")) return false;
	        if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	        if(!checkSelect(document.frm.fuente, "Cómo nos Conociste")) return false;
	        if ((document.frm.email.value.length && document.frm.mostraremail.checked) || (document.frm.tel1.value.length && document.frm.mostrartel1.checked) ||
    		    (document.frm.tel2.value.length && document.frm.mostrartel2.checked) || (document.frm.tel3.value.length && document.frm.mostrartel3.checked) ||	(document.frm.tel4.value.length && document.frm.mostrartel4.checked))
	            return true;
	        alert("Debes activar alguna forma de contacto");
	        return false;
<%	else %>
	        return true;
<%	end if %>
        }
    </script>
    <title>Vacaciones - Nuevo Anuncio Intercambio</title>
</head>
<body onload="window.parent.location.hash='top';">
<form name=frm onsubmit="return checkForm();" action=TrVacasSwapOferta.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<div class=menu><!--#include file="IncNuSubMenu.asp"--></div>
<%	if Request("op")="Destacar" then %>
    <!--#include file="IncTrDestacado.asp"-->
<%	
		Response.End
	end if 
%>
    <h1>Publicación de Apartamento para Intercambio de Vacaciones</h1>
    <!-- #include file="IncNuCondUso.asp" -->
	<div class=seccion>* son campos obligatorios</div>
	<h2>¿Cómo es el apartamento que quieres intercambiar?</h2>
	<h3>¿Dónde está el apartamento?</h3>
	<div class=seccion>
		<div>
			* Ciudad: <input title='Población donde está la vivienda' maxlength=50 size=19 name=ciudadnombre value="<%=vciudadnombre%>"/>
			* Zona: <input title='Zona o Barrio de la vivienda' maxlength=45 size=30 name=zona value="<%=vzona%>"/></div>
		<div>
			Dirección Cuanto más exacta, mejor será el mapa <input title='Dirección de la vivienda' maxlength=100 name=dir1 value="<%=vdir1%>" size=25 />
            Dirección (sigue) <input title='Continúa la Dirección' maxlength=100 name=dir2 value="<%=vdir2%>" size=25 />
            Código Postal <input title='El código postal es importante para el Mapa' maxlength=20 size=10 name=cp value='<%=vcp%>'/></div></div>
	<div>
		* Descríbenos la vivienda:c<textarea title='Cómo es, accesibilidad...' cols=61 rows=7 name=descripcion><%=vdescripcion%></textarea>
		<input title='Activa si el texto está en español' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Texto en Español
		<input title='Activa si el texto está en inglés' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Texto en Inglés
		<input title='Si activas podrás subir ficheros con fotos' type=checkbox name=foto <%=vfoto%> />¿Quieres subir fotos de la casa? ¡No te llamarán si no pones fotos!</div>
	<div><input title='Activa si el que venga puede traer animales' type=checkbox name=mascota <%=vmascota%> value=ON />Admites Mascotas</div
<%	
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
	<div><p><input title=Graba type=submit value='Publicar Apartamento' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
	<div><p><input title=Graba type=submit value='Añadir Piso' class=btnForm /></p></div>
<%	else %>
	<div><p><input title=Graba type=submit value='Modificar Piso' class="btnForm"/></p></div>
<%	end if %>
    </div></form>
<!-- #include file="IncPie.asp" -->
</body>