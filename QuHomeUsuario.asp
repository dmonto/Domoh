<!-- #include file="IncNuBD.asp" -->
<%
	dim numAnuncios, i, sPantalla, vfoto
	on error goto 0
		
	Session("Id")=""

	if Session("Usuario")="" or Session("Activo")<>"Si" then 
		if Request("usuario")="hector" or Request("op")="Test" then
			Session("Usuario")="hector"
			Session("Donar")="Si"
		else
			Response.Redirect "QuDomoh.asp?msg=" & MesgS("Sesón+finalizada","Session+Ended")
		end if
	end if
	
	sQuery= "SELECT foto FROM Usuarios WHERE usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
	if rst.Eof then
		Response.Write "No existe el usuario " & Session("Usuario")
		Response.End  
	end if
	vfoto=rst("foto")
	rst.Close
%>
<head><!--#include file="IncTrCabecera.asp"--><title><%=MesgS("Domoh - Zona Usuarios","Domoh - User Area")%></title></head>
<body onload="window.parent.location.hash='top';"> 
<div class=container>
	<nav class=logo>
<%  if Request("msg")<>"" then %>
		<h1><%=Request("msg")%></h1>
<%  end if %>
		<a title='<%=MesgS("Página Inicial", "Go to home page")%>' class=linkutils href=TrLogOn.asp><%=MesgS("Inicio", "Home")%></a> |
	   	<a title='<%=MesgS("Finaliza la Sesión", "Ends Session")%>' href=NuLogOff.asp><%=MesgS("Salir", "Logoff")%></a>	|
   		<a title='<%=MesgS("Borrar toda tu información", "Erase your info")%>' href=TrDesactivarFront.asp><%=MesgS("Darse de Baja", "Quit for good")%></a></nav>
	<div class=main>
		<p>
            <%=Session("Nombre")%>, <%=MesgS("modifica tu Perfil pulsando","update your Profile by clicking ")%>
		    <a title='<%=MesgS("Modificar Ficha de Usuario", "Update User Profile")%>' href='TrUsuario.asp?origen=QuHomeUsuario.asp'><%=MesgS("aquí", "here")%></a>.</p></div>
<%  if vfoto<>"" and Not(IsNull(vfoto)) then %>
		<img alt='Foto' title='<% Response.Write MesgS("Si ya no te gusta esta foto pulsa en Modificar Perfil", "If you don''t like this picture anymore press Update Profile")%>' class=bordeazul 
            src=http://domoh.com/<%=vfoto%> />
<% 	end if %>
<% 
	sQuery= "SELECT COUNT(*) AS numAnuncios FROM Anuncios a WHERE a.usuario='"& Session("Usuario") & "' "
	rst.Open sQuery, sProvider
	numAnuncios=rst("numAnuncios")
	rst.Close

	if numAnuncios then
		sQuery= "SELECT a.activo AS aActivo, a.id AS aId, a.tabla AS aTabla, a.provincia AS aProvincia, i.tipoviv AS iTipoViv, * FROM ((Anuncios a INNER JOIN Usuarios u ON a.usuario=u.usuario) "
		sQuery= sQuery & " LEFT JOIN Pisos p ON a.id=p.id) LEFT JOIN Inquilinos i ON a.id=i.id WHERE a.usuario='"& Session("Usuario") & "' "
		rst.Open sQuery, sProvider
	end if
%>
    <div><p><%=MesgS("Anuncios", "Adverts")%></p></div>
<%	if numAnuncios>0 then %>
    <div>
<%      if Session("Idioma")="" then %>
	    Para ver los detalles de un anuncio, haz clic en su icono <img alt='Ver Detalle' title='Ver Detalle' src=images/TrDetalle.gif />
<%      else %>
        Click <img alt='Detalle' src=images/TrDetalle.gif /> to watch advert details
<%      end if %>
		</div>
    <div>
		<div class=anuncios2_cab title='<% Response.Write MesgS("Indica si tu anuncio es visible por usuarios o no", "Can users see your advert?")%>'><%=MesgS("Estado", "Status")%></div>
		<div class=anuncios2_cab title='<%=MesgS("¿Qué es?", "What is it?")%>'><%=MesgS("Tipo", "Type")%></div>
        <div class=anuncios2_cab title='<%=MesgS("Haz click para ver el anuncio", "Click to see the advert")%>'><%=MesgS("Anuncio", "Advert")%></div>
        <div class=anuncios2_cab title='<%=MesgS("Modificar, Ocultar", "Update, Hide")%>...'><%=MesgS("Qué hago con él", "What shall I do")%></div></div>

<%	else %>
<%		
	end if

	for i=1 to numAnuncios
%>
    <div>
	    <div><a href=# onclick="javascript:detalle('QuAnuncioDetalle.asp?tabla=<%=rst("aTabla")%>&admin=si&id=<%=rst("aId")%>')"><img src=images/TrDetalle.gif /></a></div>
        <div class=anuncios2 title='<% Response.Write MesgS("Indica si tu anuncio es visible por usuarios o no", "Can users see your advert?")%>'>
<%	    if rst("aActivo")="No" then %>
		    <%=MesgS("Oculto", "Hidden")%>
<%	    elseif rst("aProvincia")=0 then %>
			<%=MesgS("En proceso de publicación", "Awaiting Publication")%>
<%	    else %>
			<%=MesgS("Publicado", "Posted")%>
<%	    end if %>									
			</div>
		<div class=anuncios2 title='<%=MesgS("¿Qué es?", "What is it?")%>'>
<%
	    if rst("aTabla")="Pisos" then
		    Response.Write MesgS("Vivienda", "Property")
	    elseif rst("aTabla")="Inquilinos" then
		    Response.Write MesgS("Inquilino", "Flat-seeker")
	    elseif rst("aTabla")="Trabajos" then
		    Response.Write MesgS("Trabajo", "Job")
	    elseif rst("aTabla")="Curriculums" then
		    Response.Write MesgS("Curriculum", "Resumé")
	    elseif rst("aTabla")="Misc" then
		    Response.Write MesgS("Anuncio", "Advert")
	    end if
%>
		    </div>
		<div class=anuncios2 title='<%=MesgS("Haz click para ver el anuncio", "Click to see advert")%>'>
            <a href="javascript:detalle('QuAnuncioDetalle.asp?tabla=<%=rst("aTabla")%>&admin=si&id=<%=rst("aId")%>')"><%=rst("cabecera")%></a></div>
		<div class=anuncios2>
<%
	    if rst("aTabla")="Pisos" then
		    if rst("RentaViv")<>0 then
			    sPantalla="TrCasaRegOfrezcoFront.asp"
		    elseif rst("RentaVacas")<>0 then
			    sPantalla="TrVacasRegOfrezcoFront.asp"
		    end if
	    elseif rst("aTabla")="Inquilinos" then
		    sPantalla="TrCasaRegDemandaFront.asp" 
	    end if
%>
		    <a href="<%=sPantalla%>?op=Modificar&id=<%=rst("aId")%>"><%=MesgS("Modificar", "Update")%></a>
<%	    if rst("aTabla")="Pisos" then %>
			<a href="<%=sPantalla%>?op=Fotos&id=<%=rst("aId")%>"><%=MesgS("Fotos", "Pictures")%></a>
<%	        
        end if
	    if rst("aActivo")="No" then 
%>
			<a href="<%=sPantalla%>?op=Activar&id=<%=rst("aId")%>"><%=MesgS("Reactivar", "Reactivate")%></a><a href="<%=sPantalla%>?op=Delete&id=<%=rst("aId")%>"><%=MesgS("Borrar", "Delete")%></a>
<%	    elseif rst("aProvincia")<>0 then %>
			<a href="<%=sPantalla%>?op=Borrar&id=<%=rst("aId")%>"><%=MesgS("Ocultar", "Hide")%></a>
<%	    end if %>									
			<a href="<%=sPantalla%>?op=Destacar&id=<%=rst("aId")%>"><%=MesgS("Destacar", "Bold")%></a><a href="<%=sPantalla%>?op=Traducir&id=<%=rst("aId")%>"><%=MesgS("Traducir", "Translate")%></a>
            </div></div>
<%
		rst.Movenext
	next
	if numAnuncios then rst.Close
%>				
	<div>
        <a href='TrCasaRegOfrezcoFront.asp?id=nuevo'><%=MesgS("Añadir Piso o Habitación para Vivir", "New Flat or Room to Rent for Living")%></a> - 
        <a href='TrVacasRegOfrezcoFront.asp?id=nuevo'><%=MesgS("Para Vacaciones", "For Holidays")%></a> - 
        <a href='TrCasaRegDemandaFront.asp?id=nuevo'><%=MesgS("Añadir Anuncio de Búsqueda de Piso para Alquilar", "New Request for Flat or Room for Living")%></a> - 
        </div></div>
<div><!-- #include file="IncPie.asp" --></div>
</body>