<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vtipo, vdir1, vciudadnombre, vdir2, vcp, vzona, vrentavacas, vdescripcion, vgente, vfuma, vmascota, vidiomaes, vidiomaen, vfoto, sBody
	on error goto 0
	
	vfuma=" checked"
	vmascota=" checked"
	vtipo="Piso"
	vfoto=" checked"
	if Session("Idioma")="" then vidiomaes=" checked" else vidiomaes="" 
	if Session("Idioma")="En" then vidiomaen=" checked" else vidiomaen="" 

	if Request("id")<>"" then
		if Request("op")="Borrar" then
		 	sQuery = "UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") &")"
			rst.Close
			Mail "hector@domoh.com", "Piso Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=Piso+Ocultado."
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
			sBody=Session("Usuario") & " ha reactivado el piso (" & rst("cabecera") &")"
			rst.Close
			Mail "hector@domoh.com", "Piso Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=Piso+Activado."
		elseif Request("op")="Modificar" then
			sQuery="SELECT * FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.id=" & Request("id")
			rst.Open sQuery, sProvider
			vtipo=rst("tipo")
			vdir1=rst("dir1")
			vdir2=rst("dir2")
			vcp=rst("cp")
			vciudadnombre=rst("ciudadnombre")
			vzona=rst("zona")
			if rst("rentavacas") > 0 then vrentavacas=rst("rentavacas")
			vdescripcion=rst("descripcion")
			vgente=rst("gente")
			if UCase(rst("fuma"))<>"ON" then vfuma=""
			if UCase(rst("mascota"))<>"ON" then vmascota=""
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes=""
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen="" else vidiomaen=" checked"
			if IsNull(rst("foto")) or rst("foto")="" then vfoto=""
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrVacasRegOferta.asp?op=FotoBorrada"
		end if
	end if
%>
<head>
    <!-- #include file="IncTrCabecera.asp" -->
    <script type=text/javascript>
        function checkForm() {
	        if(!check(document.frm.ciudadnombre, "Ciudad")) return false; if(!check(document.frm.zona, "Zona")) return false;
	        if(!check(document.frm.rentavacas, "Renta Semanal")) return false; if(!check(document.frm.descripcion, "Descripción")) return false;
	        if(document.frm.tipo[1].checked) {if (!check(document.frm.gente, "Cómo son las personas")) return false};
<%	if Request("id")="" then %>
    	    if(!check(document.frm.nombre, "Nombre de Contacto")) return false;	if(!check(document.frm.email, "E-mail")) return false;
	        if(!check(document.frm.usuario, "Nombre de Usuario")) return false;
	        if(!checkSelect(document.frm.fuente, "Cómo nos Conociste")) return false;
	        if(!checkPassword(document.frm.password, document.frm.password2)) return false;
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
    <title>Vacaciones - Nuevo Anuncio</title>
</head>
<body onload="window.parent.location.hash='top';"> 
<form name=frm onsubmit='return checkForm();' action=TrVacasRegOferta.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<div class=menu><!--#include file="IncNuSubMenu.asp"--></div>
<%	if Request("op")="Destacar" then %>
	<!--#include file="IncTrDestacado.asp"-->
<%	
		Response.End
	end if 
%>
	<div class=banner><p>Publicación de Apartamento u Hotel para Vacaciones</p></div></div>
<!--#include file="IncNuCondUso.asp"-->
<div><p>¿Cómo es el alojamiento que quieres alquilar?</p></div>
<div>¿Qué es?</div>
<div>
    <input title='Pulsa si no compartirán con nadie' type=radio value=Piso name=tipo <% if vtipo="Piso" then Response.Write " checked"%> />  Es un apartamento completo 
	<input title='Pulsa si compartirán la casa' type=radio value=Habitación name=tipo <% if vtipo="Habitación" then Response.Write " checked"%> />Es una habitación en casa compartida
	<input title='Pulsa si es un hotel o albergue' type=radio value=Hotel name=tipo <% if vtipo="Hotel" then Response.Write " checked"%> /> Es un hotel</div>
<div>¿Dónde está el piso/habitación/hotel?</div>
<div>
    <div>
	    * Ciudad <input title='Nombre de la población' maxlength=50 size=19 name=ciudadnombre value="<%=vciudadnombre%>"/>
		* Zona <input title='Nombre del barrio o urbanización' maxlength=45 size=30 name=zona value="<%=vzona%>"/>
        *  Renta Semanal: <input title='Coste completo equivalente a una semana' maxlength=30 name=rentavacas value='<%=vrentavacas%>' size=8/></div>
	<div>
        Dirección Cuanto más exacta, mejor será el mapa <input title='Calle, número, piso...' maxlength=100 name=dir1 value="<%=vdir1%>" size=34 />
        Dirección (sigue) <input title='Urbanización...' maxlength=100 name=dir2 value="<%=vdir2%>" size=34 /> Codigo Postal <input maxlength=20 size=10 name=cp value='<%=vcp%>'/></div></div>
	<div>
		<div class=seccion>
			* Descríbenos el alojamiento: <textarea title='Tamaño, localización, comunicación...' cols=35 rows=7 name=descripcion><%=vdescripcion%></textarea>
			<input title='Activa si el texto esta en Español' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Texto en Español
			<input title='Activa si el texto esta en Inglés' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Texto en Inglés
			<input title='Si activas te pediremos los ficheros con las fotos' type=checkbox name=foto <%=vfoto%> />¿Quieres subir fotos del alojamiento? ¡No te llamarán si no pones fotos!</div>
		<div class=seccion>
			¿Cómo son las personas que viven ahí? * Obligatorio en Casas Compartidas <textarea title='Otros visitantes con los que convivirán' cols=36 rows=7 name=gente><%=vgente%></textarea> 
			<input title='Activa si permites fumadores' type=checkbox name=fuma <%=vfuma%> value=ON />Admitís Fumadores 
			<input title='Activa si permites animales' type=checkbox name=mascota <%=vmascota%> value=ON />Admitís Mascotas</div></div>
<% 
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
	<div><p><input title=Grabar type=submit value='Publicar Alojamiento' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
	<div><p><input title=Grabar type=submit value='Añadir Alojamiento' class=btnForm /></p></div>
<%	else %>
	<div><p><input title=Grabar type=submit value='Modificar Alojamiento' class=btnForm /></p></div>
<%	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>