<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vtipo, vdir1, vciudadnombre, vdir2, vcp, vzona, vrentaviv, vdescripcion, vgente, vfuma, vmascota, vfoto, vidiomaes, vidiomaen, sBody

	vfuma=" checked"
	vmascota=" checked"
	if Session("Idioma")="En" then 
		vidiomaes="" 
		vidiomaes=" checked" 
	else 
		vidiomaes=" checked"
		vidiomaen=""
	end if
	vtipo="Piso"
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
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Ocultado.","House+Hidden.")
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
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Activado.","House+Republished.")
		elseif Request("op")="Modificar" then
			sQuery="SELECT * FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.Id=" & Request("id")
			rst.Open sQuery, sProvider
			vtipo=rst("tipo")
			vdir1=rst("dir1")
			vdir2=rst("dir2")
			vcp=rst("cp")
			vciudadnombre=rst("ciudadnombre")
			vzona=rst("zona")
			if rst("rentaviv") > 0 then vrentaviv=rst("rentaviv")
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
			Response.Redirect "TrCasaRegOferta.asp?op=FotoBorrada"
		end if
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.ciudadnombre, "<%=MesgS("Ciudad","City")%>")) return false;
	    if(!check(document.frm.zona, "<%=MesgS("Zona","Area")%>")) return false;
	    if(!checkNumber(document.frm.rentaviv, "<%=MesgS("Renta Mensual","Monthly Rent")%>")) return false;
	    if(!check(document.frm.descripcion, "<%=MesgS("Descripción","Description")%>")) return false;
	    if(document.frm.tipo[1].checked) {if (!check(document.frm.gente, "<%=MesgS("Cómo son las personas","People living there")%>")) return false};
	    if(!check(document.frm.nombre, "<%=MesgS("Nombre","Name")%>")) return false;
<%	if Request("id")="" then %>
	    if(!check(document.frm.email, "E-mail")) return false; if(!check(document.frm.usuario, "<%=MesgS("Nombre de Usuario","Username")%>")) return false;
	    if(!checkSelect(document.frm.fuente, "<%=MesgS("Cómo nos Conociste","How did you get to us")%>")) return false;
	    if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if ((document.frm.email.value.length && document.frm.mostraremail.checked) ||
		    (document.frm.tel1.value.length && document.frm.mostrartel1.checked) || (document.frm.tel2.value.length && document.frm.mostrartel2.checked) ||
		    (document.frm.tel3.value.length && document.frm.mostrartel3.checked) ||	(document.frm.tel4.value.length && document.frm.mostrartel4.checked))
	        return true;
	    alert("<%=MesgS("Debes activar alguna forma de contacto","You must check one way of contact")%>");
	    return false;
<%	else %>
	    return true;
<%	end if %>
    }
</script>
<body onload="window.parent.location.hash='top';">
<form name=frm onsubmit="return checkForm();" action=TrCasaRegOferta.asp method=post>
	<input type=hidden name=id value='<%=Request("id")%>' />
<div class=container>
	<div class=topmenu>
        <a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a>
        <a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a><% Response.Write MesgS("Publicar Nuevo Lugar para Alquilar", "Post New Advert for Rent")%>
<%	if Request("op")="Destacar" then %>
    <!-- #include file="IncTrCabecera.asp" --> 
<%	
		Response.End
	end if 
%>
		</div>
	<div class=banner><h1><% Response.Write MesgS("Publicación de Piso o Habitación para Vivir", "Post New Advert for Rent")%></h1></div>
	<div class=main>
		<h2><% Response.Write MesgS("¿Cómo es la vivienda que quieres alquilar?","How is the property you want to rent?")%></h2>
		<input title='<% Response.Write MesgS("Pulsa aquí si alquilas todo el piso","Check if it is a complete apartment")%>' type=radio checked="<% if vtipo="Piso" then Response.Write " checked"%>" value='Piso' name=tipo />
		<%=MesgS("Es un piso completo ","Complete Apartment")%>
		<input title='<%=MesgS("Pulsa aquí si es para compartir","Check if it is for sharing")%>' type=radio value='Habitación' <% if vtipo="Habitación" then Response.Write " checked"%> name=tipo />
		<%=MesgS("Es una habitación","Room in a shared apartment")%>
		<input title='<%=MesgS("Pulsa aquí si no es una vivienda","Check if not for living")%>' type=radio value=Local <% if vtipo="Local" then Response.Write " checked"%> name=tipo />
		<%=MesgS("Es un local, garaje","Commercial prop., garage")%>...
		<%=MesgS("¿Dónde está el piso/habitación?","Where is the property?")%>
		* <%=MesgS("Ciudad","City")%> <input title='<%=MesgS("Nombre de la Población","Name of the City/Town")%>' maxlength=50 size=19 name=ciudadnombre value='<%=vciudadnombre%>'/>
		* <%=MesgS("Zona","Area")%> <input title='<%=MesgS("Barrio, Distrito...","Name of the Neighbourhood")%>' maxlength=45 size=30 name=zona value='<%=vzona%>'/>
		* <%=MesgS("Renta Mensual","Monthly Rent")%>: <input title='<%=MesgS("Renta Mensual en euros","Monthly Rent in Euros")%>' maxlength=30 size=8 name=rentaviv value='<%=vrentaviv%>'/>
		<%=MesgS("Dirección ","Address ")%>(<% Response.Write MesgS("Cuanto más preciso seas, mejor será el mapa. Ej.: Calle Preciados 5", "Please be accurate so we can produce a better map. For instance: Calle Preciados 5")%>)
		<input title='<%=MesgS("Calle, número, piso","Street, number, floor...")%>' maxlength=100 size=34 name=dir1 value='<%=vdir1%>'/>
		<%=MesgS("Dirección (sigue)","Address (cont.)")%> <input maxlength=100 size=34 name=dir2 value='<%=vdir2%>'/> <%=MesgS("Código Postal","Postal Code")%><input maxlength=20 size=10 name=cp value='<%=vcp%>'/>
		* <%=MesgS("Descríbenos la vivienda","Describe the Property")%>: <textarea name=descripcion rows=7 cols=30><%=vdescripcion%></textarea>
		<input type=checkbox <%=vidiomaes%> value=ON name=idiomaes /><%=MesgS("Texto en Español","Text in Spanish")%>
		<%=MesgS("¿Cómo son las personas que viven en la casa?","How are the people that are living there?")%> * <%=MesgS("Obligatorio en Casas Compartidas ","Required for Shared Houses ")%>
		<textarea name=gente rows=7 cols=34><%=vgente%></textarea> 
		<input type=checkbox <%=vfuma%> value=ON name=fuma /><%=MesgS("Admitís Fumadores ","Smokers Allowed ")%>
		<input type=checkbox <%=vmascota%> value=ON name=mascota /><%=MesgS("Admitís Mascotas","Pets Allowed ")%>
		<input type=checkbox <%=vidiomaen%> value=ON name=idiomaen /><%=MesgS("Texto en Inglés ","Text in English")%>
		<input type=checkbox <%=vfoto%> name=foto /><%=MesgS("¿Quieres subir fotos de la casa?","Do you want to upload pictures?")%>
		<%=MesgS("¡No te llamarán si no pones fotos!","People won't call if you don't upload pictures!")%>
	<%	
		if Request("id")="" then 
			if Session("Usuario")="" then FormDatosPersonales 
		end if
	%>
		<input value= 
<%	if Session("Usuario")="" then %>
			'<%=MesgS("Publicar anuncio","Post Advert")%>'
<%	elseif Request("id")="nuevo" then %>
			'Añadir piso'
<%	else %>
			'Modificar piso'
<%	end if %>
			name=submit type=submit class=btnForm id=submit /></div>
<!-- #include file="IncPie.asp" --></div></form>
</body>
