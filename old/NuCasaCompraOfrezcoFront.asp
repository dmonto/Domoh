<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vtipo, vdir1, vciudadnombre, vdir2, vcp, vzona, vprecio, vdescripcion, vfoto, vidiomaes, vidiomaen, sBody
	
	vtipo="Piso"
	vfoto=" checked"
	if Session("Idioma")="En" then
		vidiomaen=" checked"
		vidiomaes=""
	else
		vidiomaes=" checked"
		vidiomaen=""
	end if

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
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Ocultado.","House+Hidden")
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
			vtipo=rst("tipo")
			vdir1=rst("dir1")
			vdir2=rst("dir2")
			vcp=rst("cp")
			vciudadnombre=rst("ciudadnombre")
			vzona=rst("zona")
			if rst("precio") > 0 then vprecio=rst("Precio")
			vdescripcion=rst("descripcion")
			if IsNull(rst("Foto")) or rst("Foto")="" then vfoto=""
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes=""
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen="" else vidiomaen=" checked" 
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrCasaCompraOferta.asp?op=FotoBorrada"
		end if
	end if
%>
<!--#include file="IncTrCabecera.asp"-->
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.ciudadnombre, "Ciudad")) return false; if(!check(document.frm.zona, "Zona")) return false;
	    if(!check(document.frm.precio, "Precio")) return false; if(!check(document.frm.descripcion, "Descripción")) return false;
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
<form name=frm onsubmit='return checkForm();' action=TrCasaCompraOferta.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class="container">
	<nav><!--#include file="IncNuSubMenu.asp"--></nav>
<%	if Request("op")="Destacar" then %>
    <!--#include file="IncTrDestacado.asp"-->
<%	
		Response.End
	end if 
%>
	<h1>Publicación de Piso o Casa para Venta</h1>
	<div>
		<p>
            Al publicar este piso/casa estás aceptando nuestras <a title='Ver las Condiciones de Uso' href=CondUso.htm target=_blank>Condiciones Generales de Uso</a> y nuestra 
		    <a title='Ver la Política de Protección de Datos' href=ProtDatos.asp target=_blank>Política de Protección de Datos</a></p></div></div>
	<div>* son campos obligatorios</div>
	<div><p>¿Cómo es la vivienda que quieres vender?</p></div>
	<div> ¿Qué es?</div>
	<div>
		<input title='Pulsa si es un apartamento, piso, duplex' type=radio value=Piso name=tipo <% if vtipo="Piso" then Response.Write " checked"%> />Es un piso
		<input title='Pulsa si es una casa, chalet' type=radio value=Casa name=tipo <% if vtipo="Casa" then Response.Write " checked"%> /> Es una casa 
		<input title='Pulsa si no es ninguna de las anteriores' type=radio value=Local name=tipo <% if vtipo="Local" then Response.Write " checked"%> />Es un Local, Garaje... </div>
	<div>¿Dónde está el piso/casa?</div>
	<div>
		<div>
			* Ciudad <input title='Población donde se encuentra la vivienda' maxlength=50 size=19 name=ciudadnombre value='<%=vciudadnombre%>'/>
			* Zona <input title='Barrio/Distrito/Zona...' maxlength=45 size=30 name=zona value="<%=vzona%>"/>*  Precio (miles de €):
			<input title='Precio en miles de euros' maxlength=30 name=precio value='<%=vprecio%>' size=8 /></div>
		<div>
			Dirección (opcional) Cuanto más exacta, mejor será el mapa <input title='Primera línea de la dirección' maxlength=100 name=dir1 value="<%=vdir1%>" size=34 />
			Dirección (sigue) <input maxlength=100 name=dir2 value="<%=vdir2%>" size=34 /> Código Postal <input maxlength=20 size=10 name=cp value='<%=vcp%>'/></div></div>
	<div>
		* Descríbenos la vivienda: <textarea title='Metros, habitaciones, antigüedad...' cols=36 rows=7 name=descripcion><%=vdescripcion%></textarea>
		<input title='Activa si el texto está en Español' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Texto en Español
		<input title='Activa si el texto está en Inglés' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Texto en Inglés
		<a title='Activa si quieres que te añadamos el texto en inglés' href=NuTraducir.asp target=traducir>¿Te lo traducimos?</a>
		<input title='Activa si quieres que te pidamos las fotos' type=checkbox name=foto <%=vfoto%> />¿Quieres subir fotos de la casa?	¡No te llamarán si no pones fotos!</div>
<%	if Request("id")="" then %>
  	<div><p>Datos Personales</p></div>
	<div>*Nombre de Contacto <input title='Nombre de Contacto' maxlength=35 name=nombre size=20 /><input title='Activa si sois profesionales' type=checkbox name=esagencia />Somos una Agencia</div>
	<div>Forma de Contacto</div>
	<div>
		<div>Teléfono 1:  <input title='Todos los números' maxlength=30 name=tel1 size=20 /></div>
		<div><input title='Activa si quieres que lo mostremos' type=checkbox name=mostrartel1 checked />Visible para Compradores</div></div>
	<div>
		<div>Teléfono 2:  <input title='Todos los números' maxlength=30 name=tel2 size=20 /></div>
		<div><input title='Activa si quieres que lo mostremos' type=checkbox name=mostrartel2 checked />Visible para Compradores</div></div>
	<div>
		<div>Teléfono 3:  <input title='Todos los números' maxlength=30 name=tel3 size=20 /></div>
		<div><input title='Activa si quieres que lo mostremos' type=checkbox name=mostrartel3 checked />Visible para Compradores</div></div>
	<div>
		<div>Teléfono 4:  <input title='Todos los números' maxlength=30 name=tel4 size=20 /></div>
		<div><input title='Activa si quieres que lo mostremos' type=checkbox name=mostrartel4 checked />Visible para Compradores</div></div>
	<div>Instrucciones para llamarte: <input title='Horas...' name=instrucciones size=56 /></div>
	<div>
		<div>*  e-Mail:  <input title='Importante para contactarte' maxlength=60 name=email size=20 /></div>
		<div><input title='Activa si quieres que lo mostremos' type=checkbox name=mostraremail checked />Visible para Compradores</div></div>
	<div>
		<p>
            ADVERTENCIA: POR FAVOR RELLENA CORRECTAMENTE LA DIRECCIÓN DE CORREO. TE ENVIAREMOS UN MENSAJE CON INSTRUCCIONES PARA ACTIVAR TU ANUNCIO. 
			EN CASO DE QUE NO LO HAGAS NO PODREMOS PUBLICAR TU ANUNCIO. En caso de no recibir nuestro correo, por favor, comprueba tu correo “Spam”, “Junk Mail” o “Correo no deseado”</p></div>
	<div>Identificación en domoh</div>
	<div><div>* Nombre de Usuario:  </div><div><input title='Lo usarás para conectarte más tarde' maxlength=60 name=usuario size=20 /></div></div>
	<div><div>* Elige una Clave:       </div><div><input type=password maxlength=20 size=10 name=password /></div><div>*¿Cómo nos conociste? <% SelectFuente 0 %></div></div>
	<div><div>* Tecléala de Nuevo:&nbsp; </div><div><input type=password maxlength=20 size=10 name=password2 /></div></div>
	<div><p><input type=submit value='Publicar Piso' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
	<div><p><input type=submit value='Añadir Piso' class=btnForm /></p></div>
<%	else %>
	<div><p><input type=submit value='Modificar Piso' class=btnForm /></p></div>
<%	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>
