<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vdescripcion, vanuncio, vfoto, vflyer, vprovincia, vidiomaes, vidiomaen, sBody
	
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
			sBody=Session("Usuario") & " ha desactivado el anuncio (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Anuncio Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma")& ".asp?msg=" & MesgS("Anuncio+Ocultado.","Advert+Hidden")
		elseif Request("op")="Delete" then
		 	sQuery = "DELETE FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "DELETE FROM Misc WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha borrado el anuncio (" & Request("id") &")"
			Mail "hector@domoh.com", "Anuncio Borrado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Anuncio+Borrado.","Advert+Deleted.")
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el anuncio (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Anuncio Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=" & MesgS("Anuncio+Activado.","Advert+Posted")
		elseif Request("op")="Modificar" then
			sQuery= "SELECT * FROM Misc m INNER JOIN Anuncios a ON m.id=a.id WHERE a.id=" & Request("id")
			rst.Open sQuery, sProvider
			vdescripcion=rst("cabecera")
			vanuncio=rst("descripcion")
			vprovincia=rst("provincia")
			if rst("foto")<>"" then vfoto=rst("foto")
			if rst("flyer")<>"" then vflyer=rst("flyer")
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes=""
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen="" else vidiomaen=" checked"
			rst.close
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrMiscRegOferta.asp?op=FotoBorrada"
		end if
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.descripcion, '<%=MesgS("Cabecera","Title")%>')) return false;
	    if(!check(document.frm.anuncio, '<%=MesgS("Qué Ofreces","What are you offering")%>')) return false;
<%  if Request("id")="" then %>
	    if(!check(document.frm.nombre, '<%=MesgS("Nombre","Name")%>')) return false; if(!check(document.frm.email, "E-mail")) return false;
	    if(!check(document.frm.usuario, '<%=MesgS("Nombre de Usuario","Username")%>')) return false;
 	    if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if(!checkSelect(document.frm.fuente, '<%=MesgS("Cómo nos Conociste","How did you get to domoh")%>')) return false;
<%  end if %>
   	    return true;
    }
</script>
<body onload="window.parent.location.hash='top';"> 
<form method=post name=frm onsubmit='return checkForm();' action=TrMiscRegOferta.asp enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<div class=row>
        <div class=left><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	    <div class=right><a href=old/TrVarios.asp class=linkutils><%=MesgS("Anuncios Varios","Miscellaneous Adverts")%></a>&gt; <%=MesgS("PUBLICAR Anuncio","POST Advert")%></div></div>
<%	if Request("op")="Destacar" then %>
    <div class=seccion><!-- #include file="IncTrDestacado.asp" --></div>
<%	
		Response.End
	end if 
%>
    <h1 class=tituSec><%=MesgS("Publicación de Anuncio","Advert Publication")%></h1>
    <!-- #include file="IncNuCondUso.asp" -->
    <h2 class=seccion>
        * <%=MesgS("Cabecera del Anuncio","Advert Title")%>:
        <input title='<% Response.Write MesgS("Lo primero que verán de tu anuncio","This is what users will see first")%>' type=text name=descripcion size=50 value="<%=vdescripcion%>"/></h2>
<div class=seccion>
    * <%=MesgS("¿Qué Ofreces?","Your Ad")%>: <textarea title='<%=MesgS("Texto del Anuncio","Text of the Advert")%>' cols=35 rows=7 name=anuncio><%=vanuncio%></textarea>
    <input title='<% Response.Write MesgS("Activa si el texto está en Español","Check if the text is in Spanish")%>' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />
    <%=MesgS("Texto en Español","Advert in Spanish")%> <input title='<% Response.Write MesgS("Activa si el texto está en Inglés", "Check if the text is in English")%>' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />
    <%=MesgS("Texto en Inglés","Advert in English")%>
<%  if vfoto="" then %>
    <%=MesgS("Deja aquí tu foto ","Upload here your picture ")%><input title='<% Response.Write MesgS("Pulsa a la derecha para buscar el fichero","Press on the button to look for the file")%>' type=file name=foto />
<%  else %>
	<img title='<% Response.Write MesgS("Pulsa a la derecha para cambiar la foto","Press on the right to change the picture")%>' class=bordeazul src='http://domoh.com/<%=vfoto%>'/>
    <%=MesgS("Cambiar foto ","Change picture ")%> <input title='<% Response.Write MesgS("Pulsa a la derecha para cambiar la foto", "Press on the right to change the picture")%>' type=file name=foto />
<%  
    end if
    if vflyer="" then 
%>
        <%=MesgS("Si quieres, añade un flyer ","Upload here your flyer ")%><input title='<%=MesgS("Pulsa a la derecha para buscar el fichero","Press on the button to look for the file")%>' type=file name=flyer />
<%  else %>
		<img title='<%=MesgS("Pulsa a la derecha para cambiar el flyer","Press on the right to change the flyer")%>' class=bordeazul src='http://domoh.com/<%=vflyer%>'/>
        <%=MesgS("Cambiar flyer ","Change flyer ")%><input title='<%=MesgS("Pulsa a la derecha para cambiar el flyer","Press on the right to change the flyer")%>' type=file name=flyer />
<%  end if %>
        </div>
	<div class=seccion><%=MesgS("Provincia en la que estás","Your Region")%>: <% SelectProvincia "", vprovincia %></div></div>
<% 
	if Request("id")="" then 
		if Session("Usuario")="" then FormDatosPersonales
%>
<div><p><input title='<%=MesgS("Grabar","Save")%>' type=submit value='<%=MesgS("Publicar Anuncio","Post Ad")%>' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
<div><p><input type=submit value='<%=MesgS("Añadir Anuncio","Publish Ad")%>' class=btnForm /></p></div>
<%	else %>
<div><p><input type=submit value='<%=MesgS("Modificar Anuncio","Update Ad")%>' class=btnForm /></p></div>
<%	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>
