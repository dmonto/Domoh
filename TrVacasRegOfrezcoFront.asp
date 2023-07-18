<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vtipo, vdir1, vciudadnombre, vdir2, vcp, vzona, vrentavacas, vdescripcion, vgente, vfuma, vmascota, vidiomaes, vidiomaen, vfoto, sBody

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
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Piso Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Ocultado.","Advert+Hidden.")
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
			sBody=Session("Usuario") & " ha reactivado el piso (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Piso Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Activado.","Advert+Republished.")
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
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javaccript>
    function checkForm() {
	    if(!check(document.frm.ciudadnombre, "<%=MesgS("Ciudad","City")%>")) return false;
	    if(!check(document.frm.zona, "<%=MesgS("Zona","Area")%>")) return false;
	    if(!check(document.frm.rentavacas, "<%=MesgS("Renta Semanal","Weekly Rent")%>")) return false;
	    if(!check(document.frm.descripcion, "<%=MesgS("Descripción","Description")%>")) return false;
	    if(document.frm.tipo[1].checked) {if (!check(document.frm.gente, "<%=MesgS("Cómo son las personas","People in the house")%>")) return false};
<%	if Request("id")="" then %>
	    if(!check(document.frm.nombre, "<%=MesgS("Nombre de Contacto","Contact Name")%>")) return false;
	    if(!check(document.frm.email, "E-mail")) return false;
	    if(!check(document.frm.usuario, "<%=MesgS("Nombre de Usuario","User Name")%>")) return false;
	    if(!checkSelect(document.frm.fuente, "<%=MesgS("Cómo nos Conociste","How did you get to Domoh")%>")) return false;
	    if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if ((document.frm.email.value.length && document.frm.mostraremail.checked) || (document.frm.tel1.value.length && document.frm.mostrartel1.checked) ||
    		(document.frm.tel2.value.length && document.frm.mostrartel2.checked) || (document.frm.tel3.value.length && document.frm.mostrartel3.checked) || (document.frm.tel4.value.length && document.frm.mostrartel4.checked))
	        return true;
	    alert("<% Response.Write MesgS("Debes activar alguna forma de contacto", "Please check one of the ways of contact")%>");
	    return false;
<%	else %>
	    return true;
<%	end if %>
}
</script>
<body onload="window.parent.location.hash='top';"> 
<form name=frm onsubmit="return checkForm();" action=TrVacasRegOferta.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%>/>
<div class=container>
	<div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class=topmenu><a href=TrVacas.asp class=linkutils><%=MesgS("Anuncios Vacaciones","Holiday Adverts")%> &gt; </a><% Response.Write MesgS("Publicar Nuevo Lugar de Vacaciones", "Post New Holidays Advert")%></div>
<%	if Request("op")="Destacar" then %>
		<!-- #include file="IncTrDestacado.asp" -->
<%	
		Response.End
	end if 
%>
	<h1 class="banner"><% Response.Write MesgS("Publicación de Apartamento u Hotel para Vacaciones", "Vacation Advert Publication")%></h1>
	<div class=main>
		<!-- #include file="IncNuCondUso.asp" -->
		<h1><% Response.Write MesgS("¿Cómo es el alojamiento que quieres alquilar?", "How is the Lodging?")%></h1>
		<h2><%=MesgS("¿Qué es?","What is it?")%></h2>
        <h3><input title='<% Response.Write MesgS("Pulsa si no compartirán con nadie", "Check if the lodging will not be shared")%>' type=radio value=Piso name=tipo <%if vtipo="Piso" then Response.Write " checked"%> />  
        <%=MesgS("Es un apartamento completo","It's a complete apartment")%> </h3>
		<h3><input title='<% Response.Write MesgS("Pulsa si compartirán la casa", "Check if guests share the lodging")%>' type=radio value=Habitación name=tipo <%if vtipo="Habitación" then Response.Write " checked"%> />	
        <% Response.Write MesgS("Es una habitación en casa compartida", "Is a Room in a shared apartment")%></h3>
		<h3><input title='<% Response.Write MesgS("Pulsa si es un hotel o albergue", "Check if it is a hotel or B&B")%>' type=radio value=Hotel name=tipo <%if vtipo="Hotel" then Response.Write " checked"%> /> 
        <%=MesgS("Es un hotel","Is a Hotel")%></h3>
		<h2><% Response.Write MesgS("¿Dónde está el piso/habitación/hotel?", "Where is the apartment/room/hotel?")%></h2>
	    <h3>* <%=MesgS("Ciudad","City")%> <input title='<%=MesgS("Nombre de la población","Name of the City/Town")%>' maxlength=50 size=19 name=ciudadnombre value='<%=vciudadnombre%>' autocomplete="address-line3" /></h3>
        <h3>* <%=MesgS("Zona","Area")%> <input title='<% Response.Write MesgS("Nombre del barrio o urbanización", "Name of the neighbourhood or district")%>' maxlength=45 size=30 name=zona value='<%=vzona%>' /></h3>
        <h3>* <%=MesgS("Renta Semanal","Cost per Week")%>:
        <input title='<% Response.Write MesgS("Coste completo equivalente a una semana", "Complete cost of a week stay")%>' maxlength=30 name=rentavacas value='<%=vrentavacas%>' size=8 /></h3>
	    <h3><%=MesgS("Dirección ","Address ")%> <% Response.Write MesgS("Cuanto más exacta, mejor será el mapa", "Please be accurate for map correctness")%>
        <input title='<%=MesgS("Calle, número, piso...","Street, number, floor...")%>' maxlength=100 name=dir1 value='<%=vdir1%>' size=34 autocomplete="address-line1" /></h3>
        <h3><%=MesgS("Dirección (sigue)","Address (cont.)")%> <input title='<%=MesgS("Urbanización...","Name of the condo...")%>' maxlength=100 name=dir2 value='<%=vdir2%>' size=34 autocomplete="address-line2" />
        <%=MesgS("Código Postal","Postal Code")%> <input maxlength=20 size=10 name=cp value='<%=vcp%>'/></h3>
	    <h3>* <%=MesgS("Descríbenos el alojamiento","Describe the lodging")%>:
        <textarea title='<%=MesgS("Tamaño, localización, comunicación...","Size, location, distance to atractions")%>' cols=35 rows=7 name=descripcion><%=vdescripcion%></textarea>
        <input type=checkbox name=idiomaes <%=vidiomaes%> value=ON /><%=MesgS("Texto en Español","Text in Spanish")%>
        <input type=checkbox name=idiomaen <%=vidiomaen%> value=ON /><%=MesgS("Texto en Inglés","Text in English")%></h3>
        <h3><input type=checkbox name=foto <%=vfoto%> /><%=MesgS("¿Quieres subir fotos del alojamiento?","Do you want to upload pictures?")%>
        <%=MesgS("¿No te llamarán si no pones fotos!","They won't call you if you don't upload pictures!")%></h3>
	    <h2><%=MesgS("¿Cómo son las personas que viven ahí?","How are people that stay there?")%></h2>
        <h2>* <%=MesgS("Obligatorio en Casas Compartidas","Required for Shared Houses")%> <textarea cols=36 rows=7 name=gente><%=vgente%></textarea> </h2>
		<h3><input type=checkbox name=fuma <%=vfuma%> value=ON /><%=MesgS("Admitís Fumadores ","Smoking Allowed ")%></h3>
		<h3><input type=checkbox name=mascota <%=vmascota%> value=ON /><%=MesgS("Admitís Mascotas","Pets Allowed")%></h3>
<% 
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
		<h2><input type=submit value='<%=MesgS("Publicar Alojamiento","Post Lodging")%>' class=btnForm /></h2>
<%	elseif Request("id")="nuevo" then %>
		<h2><input type=submit value='<%=MesgS("Añadir Alojamiento","Add Lodging")%>' class=btnForm /></h2>
<%	else %>
		<h2><input type=submit value='<%=MesgS("Modificar Alojamiento","Update Lodging")%>' class=btnForm /></h2>
<%	end if %>
		</div>
	<!-- #include file="IncPie.asp" --></div></form></body>