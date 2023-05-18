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
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") &")"
			rst.Close
			Mail "hector@domoh.com", "Piso Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Piso+Ocultado.","Advert+Hidden.")
		elseif Request("op")="Delete" then
		 	sQuery = "DELETE FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "DELETE FROM Pisos WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha borrado el piso (" & Request("id") &")"
			Mail "hector@domoh.com", "Piso Borrado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Borrado.","Advert+Deleted.")
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
		 	rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el piso (" & rst("cabecera") &")"
			rst.Close
			Mail "hector@domoh.com", "Piso Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Piso+Activado.","Advert+Reposted.")
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
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javascript>
function checkForm() {
	if(!check(document.frm.ciudadnombre, "<%=MesgS("Ciudad","City")%>")) return false;
	if(!check(document.frm.zona, "<%=MesgS("Zona","Area")%>")) return false;
	if(!check(document.frm.descripcion, "<%=MesgS("Descripción","Description")%>")) return false;
<%	if Request("id")="" then %>
	if(!check(document.frm.nombre, "<%=MesgS("Nombre","Name")%>")) return false; if(!check(document.frm.email, "E-mail")) return false;
	if(!check(document.frm.usuario, "<%=MesgS("Nombre de Usuario","User Name")%>")) return false;
	if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	if(!checkSelect(document.frm.fuente, "<%=MesgS("Cómo nos Conociste","How did you get to Domoh")%>")) return false;
	if ((document.frm.email.value.length && document.frm.mostraremail.checked) || (document.frm.tel1.value.length && document.frm.mostrartel1.checked) ||
		(document.frm.tel2.value.length && document.frm.mostrartel2.checked) || (document.frm.tel3.value.length && document.frm.mostrartel3.checked) ||	(document.frm.tel4.value.length && document.frm.mostrartel4.checked))
	    return true;
	alert("<%=MesgS("Debes activar alguna forma de contacto","You must check one way of contact")%>");
	return false;
<%	else %>
	return true;
<%	end if %>
}
</script>
<body onload="window.parent.location.hash='top';">
<form name=frm onsubmit="return checkForm();" action=TrVacasSwapOferta.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<div class="logo"><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class="topmenu"><a href=TrVacas.asp class=linkutils><%=MesgS("Anuncios Vacaciones","Holiday Adverts")%> </a> &gt; <% Response.Write MesgS("Publicar Nuevo Lugar de Vacaciones", "Post New Holidays Advert")%></div>
	<h1 class=banner><% Response.Write MesgS("Publicación de Apartamento para Intercambio de Vacaciones","Flat Swap Advert Publication")%></h1>
    <div class=main>
		<!-- #include file="IncNuCondUso.asp" -->
		<p><% Response.Write MesgS("Cómo es el apartamento que quieres intercambiar?","How is the apartment you want to swap?")%></p>
    	<%=MesgS("¿Dónde está el apartamento?","Where is the apartment?")%>
    <div>
	    <div>
		    * <%=MesgS("Ciudad","City")%>: <input title='<% Response.Write MesgS("Población donde está la vivienda","City/Town where the lodging is")%>' maxlength=50 size=19 name=ciudadnombre value='<%=vciudadnombre%>' autocomplete="address-line3"
            * <%=MesgS("Zona","Area")%>: <input title='<% Response.Write MesgS("Zona o Barrio de la vivienda", "Name of the Neighbourhood")%>' maxlength=45 size=30 name=zona value='<%=vzona%>'/></div>
        <div>
		    <%=MesgS("Dirección ","Address ")%> <% Response.Write MesgS("Cuanto más exacta, mejor será el mapa", "Please be accurate so we can produce a better map")%>
            <input title='<% Response.Write MesgS("Calle, número, piso...", "Street, number, floor...")%>' maxlength=100 name=dir1 value='<%=vdir1%>' size=25 autocomplete="address-line1" />
            <%=MesgS("Dirección (sigue)","Address (cont.)")%> <input title='<%=MesgS("Urbanización...","Condo...")%>' maxlength=100 name=dir2 value='<%=vdir2%>' size=25 autocomplete="address-line2" /> <%=MesgS("Codigo Postal","Postal Code")%>&nbsp; 
			<input title='<% Response.Write MesgS("El código postal es importante para el Mapa","Postal Code is inportant for the map")%>' maxlength=20 size=10 name=cp value='<%=vcp%>' /></div></div>
	<div>
	    * <%=MesgS("Descríbenos la vivienda","Describe the property")%>:
        <textarea title='<% Response.Write MesgS("Cómo es, accesibilidad", "How it is, location")%>...' cols=61 rows=7 name=descripcion><%=vdescripcion%></textarea>
        <input title='<% Response.Write MesgS("Activa si el texto está en español","Check if the advert is in Spanish")%>' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />
    	<%=MesgS("Texto en Español","Text in Spanish")%> <input title='<% Response.Write MesgS("Activa si el texto está en inglés","Check if the advert is in English")%>' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />
        <%=MesgS("Texto en Inglés","Text in English")%> <input title='<%=MesgS("Si activas podrás subir ficheros con fotos","If you check we will ask for the files with the pictures")%>' type=checkbox name=foto <%=vfoto%> />
        <%=MesgS("¿Quieres subir fotos de la casa?","Do you want to upload pictures?")%> <%=MesgS("!No te llamarán si no pones fotos!","They won't call if they don't see pictures!")%></div>
    <div>
	    <input title='<%=MesgS("Activa si el que venga puede traer animales","Check if the guest can bring along pets")%>' type=checkbox name=mascota <%=vmascota%> value=ON /><%=MesgS("Admites Mascotas","Pets Allowed")%></div>
<%	
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
    <div><p><input type=submit value='<%=MesgS("Publicar Apartamento","Post Apartment")%>' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
    <div><p><input type=submit value='<%=MesgS("Añadir Piso","Add Apartment")%>' class=btnForm /></p></div>
<%	else %>
    <div><p><input type=submit value='<%=MesgS("Modificar Piso","Update Apartment")%>' class=btnForm /></p></div>
<%	end if %>
</div></div></form><!-- #include file="IncPie.asp" -->
</body>
