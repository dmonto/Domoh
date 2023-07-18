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
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") &")"
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
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Piso+Activado.","Advert+Republished")
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
	    if(!check(document.frm.ciudadnombre, "<%=MesgS("Ciudad","City")%>")) return false;
	    if(!check(document.frm.zona, "<%=MesgS("Zona","Area")%>")) return false;
	    if(!check(document.frm.precio, "<%=MesgS("Precio","Price")%>")) return false;
	    if(!check(document.frm.descripcion, "<%=MesgS("Descripción","Description")%>")) return false;
	    if(!check(document.frm.nombre, "<%=MesgS("Nombre","Name")%>")) return false;
<%	if Request("id")="" then %>
	    if(!check(document.frm.email, "E-mail")) return false;
	    if(!check(document.frm.usuario, "<%=MesgS("Nombre de Usuario","User name")%>")) return false;
	    if(!checkSelect(document.frm.fuente, "<%=MesgS("Cómo nos Conociste","How did you get to domoh")%>")) return false;
	    if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if ((document.frm.email.value.length && document.frm.mostraremail.checked) || (document.frm.tel1.value.length && document.frm.mostrartel1.checked) ||
		    (document.frm.tel2.value.length && document.frm.mostrartel2.checked) || (document.frm.tel3.value.length && document.frm.mostrartel3.checked) || (document.frm.tel4.value.length && document.frm.mostrartel4.checked))
	        return true;
	    alert("<%=MesgS("Debes activar alguna forma de contacto","Please check one way of contact")%>");
	    return false;
<%	else %>
	    return true;
<%	end if %>
    }
</script>
<body onload="window.parent.location.hash='top';">
<form name=frm onsubmit="return checkForm();" action=TrCasaCompraOferta.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
    <div class=logo>
        <a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a>
        <a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a><%=MesgS("Publicar Nuevo Lugar para Vender","Post New Advert for Sale")%></div>
	<div class=banner><% Response.Write MesgS("Publicación de Piso o Casa para Venta", "House Sale Advert Publication")%></div>
    <div class=main>
		<p><% Response.Write MesgS("¿Cómo es la vivienda que quieres vender?", "How is the property you want to sell?")%></p>
    	<%=MesgS("¿Qué es?","What is it?")%>
	    <input title='<% Response.Write MesgS("Pulsa si es un apartamento, piso, duplex","Check if it is an apartment, flat...")%>' 
		    type=radio value=Piso name=tipo <%if vtipo="Piso" then Response.Write " checked"%> /> <%=MesgS("Es un piso ","It's an apartment")%>
		<input title='<% Response.Write MesgS("Pulsa si es una casa, chalet", "Check if it is a complete house")%>' type=radio value=Casa name=tipo <%if vtipo="Casa" then Response.Write " checked"%> />
		<%=MesgS("Es una casa ","It's a complete house ")%>
		<input title='<% Response.Write MesgS("Pulsa si no es ninguna de las anteriores","Check if none of the above")%>' type=radio value=Local name=tipo <%if vtipo="Local" then Response.Write " checked"%> />
		<%=MesgS("Es un Local, Garaje... ","It's a business, parking... ")%>
		<%=MesgS("¿Dónde está el piso/casa?","Where is the property?")%>
		* <%=MesgS("Ciudad","City")%> <input title='<%=MesgS("Población donde se encuentra la vivienda","City/Town of the property")%>' maxlength=50 size=19 name=ciudadnombre value='<%=vciudadnombre%>'/>
		* <%=MesgS("Zona","Area")%> <input title='<%=MesgS("Barrio/Distrito/Zona...","Neighbourhood/District...")%>' maxlength=45 size=30 name=zona value='<%=vzona%>'/>
		*  <%=MesgS("Precio (miles de ","Price (thousand ")%> €): <input title='<%=MesgS("Precio en miles de euros","Price in thousand Euros")%>' maxlength=30 name=precio value='<%=vprecio%>' size=8 />
		<%=MesgS("Dirección (opcional) ","Address (optional) ")%><%=MesgS("Cuanto más exacta, mejor será el mapa","More accuracy will produce a better map")%>
		<input title='<%=MesgS("Primera línea de la dirección","Street, number, floor...")%>' maxlength=100 name=dir1 value='<%=vdir1%>' size=34 />
		<%=MesgS("Dirección (sigue)","Address (cont.)")%> <input maxlength=100 name=dir2 value='<%=vdir2%>' size=34 />
		<%=MesgS("Código Postal","Post Code")%> <input maxlength=20 size=10 name=cp value='<%=vcp%>'/>
        * <%=MesgS("Descríbenos la vivienda","Describe the property")%>: <textarea title='<%=MesgS("Metros, habitaciones, antigüedad...","Size, Age, Rooms...")%>' cols=36 rows=7 name=descripcion><%=vdescripcion%></textarea>
    	<input title='<%=MesgS("Activa si el texto está en Español","Check if the advert is in Spanish")%>' type=checkbox name=idiomaes <%=vidiomaes%> value=ON /><%=MesgS("Texto en Español","Text in Spanish")%>
        <input title='<%=MesgS("Activa si el texto está en Inglés","Check if the advert is in English")%>' type=checkbox name=idiomaen <%=vidiomaen%> value=ON /><%=MesgS("Texto en Inglés","Text in English")%>
        <input type=checkbox name=foto <%=vfoto%> /><%=MesgS("¿Quieres subir fotos de la casa?","Do you want to upload pictures?")%>
	    <%=MesgS("!No te llamarán si no pones fotos!","People won't call if they can't see pictures!")%>
<%	
    if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales 
    end if
%>
        <input name=submit type=submit class=btnForm id=submit value=
<%	if Request("id")="" then %>
            '<%=MesgS("Publicar anuncio","Post Advert")%>'/>
<%	elseif Request("id")="nuevo" then %>
            'Añadir piso'/>
<%	else %>
            'Modificar piso'/>
<%	end if %>
        </div></header></form>
<!-- #include file="IncPie.asp" --></body>