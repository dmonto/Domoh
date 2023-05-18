<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
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
			sQuery=" SELECT * FROM Anuncios a WHERE a.id="& Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Piso Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Ocultado.","House+Hidden.")
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sQuery=" SELECT * FROM Anuncios a WHERE a.id="& Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el piso (" & rst("cabecera") & ")"
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
			Response.Redirect "NuCasaRegOferta.asp?op=FotoBorrada"
		end if
	end if
%>
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.ciudadnombre, "Ciudad")) return false; if(!check(document.frm.zona, "Zona")) return false;
	    if(!checkNumber(document.frm.rentaviv, "Renta Mensual")) return false; if(!check(document.frm.descripcion, "Descripción")) return false;
	    if(document.frm.tipo[1].checked) {if (!check(document.frm.gente, "Cómo son las personas")) return false};
	    if(!check(document.frm.nombre, "Nombre")) return false;
<%	if Request("id")="" then %>
	    if(!check(document.frm.email, "E-mail")) return false; if(!check(document.frm.usuario, "Nombre de Usuario")) return false;
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
<body onload="window.parent.location.hash='top';">
<form name=frm onsubmit="return checkForm();" action=NuCasaRegOferta.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<div class="logo"><!--#include file="IncNuSubMenu.asp"--></div>
<%	if Request("op")="Destacar" then %>
    <!--#include file="IncTrDestacado.asp"-->
<%	
		Response.End
	end if 
%>
	<div class=banner><p>Publicación de Piso o Habitación para Vivir</p></div>
	<div class=main>
		<p>
            Al publicar este piso/habitación estás aceptando nuestras <a title='Leer Condiciones' href=CondUso.htm target=_blank>Condiciones Generales de Uso</a> 
		    y nuestra <a title='Leer Política de Protección de Datos' href=ProtDatos.asp target=_blank>Política de Protección de Datos</a></p>
		* son campos obligatorios</div>
<div><p>¿Cómo es la vivienda que quieres alquilar?</p></div>
<div> ¿Qué es?</div>
<div>
	<input title='Pulsa aquí si alquilas todo el piso' type=radio value=Piso name=tipo <% if vtipo="Piso" then Response.Write " checked"%>/>  Es un piso completo 
	<input title='Pulsa aquí si es para compartir' type=radio value=Habitación name=tipo <% if vtipo="Habitación" then Response.Write " checked"%>/>Es una habitación en casa compartida
	<input title='Pulsa aquí si no es una vivienda' type=radio value=Local name=tipo <% if vtipo="Local" then Response.Write " checked"%>/> Es un local, garaje...</div>
<div>¿Dónde está el piso/habitación?</div>
<div>
	<div>
		* Ciudad <input title='Nombre de la Población' maxlength=50 size=19 name=ciudadnombre value='<%=vciudadnombre%>'/>
		* Zona <input title='Barrio, Distrito...' maxlength=45 size=30 name=zona value="<%=vzona%>"/>
        * Renta Mensual: <input title='Renta Mensual en euros' maxlength=30 name=rentaviv value='<%=vrentaviv%>' size=8 /></div>
	<div>
		Dirección Cuanto más preciso seas, mejor será el mapa. Ej.: Calle Preciados 5 <input title='Calle, número, puerta' maxlength=100 name=dir1 value="<%=vdir1%>" size=34 />
        Dirección (sigue) <input title='Segunda línea de la dirección' maxlength=100 name=dir2 value="<%=vdir2%>" size=34/>
		Código Postal <input title='Usado para el mapa, etc' maxlength=20 size=10 name=cp value='<%=vcp%>'/></div></div>
<div>
	<div>
		* Descríbenos la vivienda: <textarea title='Número de habitaciones, metros, ubicación...' cols=36 rows=7 name=descripcion><%=vdescripcion%></textarea>
		<input title='Activa si el anuncio está en Español' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Texto en Español
		<input title='Activa si el anuncio está en Inglés' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Texto en Inglés
		<a title='Haz click y te explicamos como pedir la traducción' href=NuTraducir.asp target=traducir>¿Te lo traducimos?</a>
		<input type=checkbox name=foto <%=vfoto%> />¿Quieres subir fotos de la casa? !No te llamarán si no pones fotos!</div>
	<div>
		¿Cómo son las personas que viven en la casa? * Obligatorio en Casas Compartidas <textarea cols=34 rows=7 name=gente><%=vgente%></textarea> 
		<input type=checkbox name=fuma <%=vfuma%> value=ON />Admitís Fumadores <input type=checkbox name=mascota <%=vmascota%> value=ON />Admitís Mascotas</div></div>
<%	
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
<div><p><input type=submit value='Publicar Piso' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
<div><p><input type=submit value='Añadir Piso' class=btnForm /></p></div>
<%	else %>
<div><p><input type=submit value='Modificar Piso' class=btnForm /></p></div>
<%	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>