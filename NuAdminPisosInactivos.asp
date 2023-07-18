<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim numPisos, i, sIdioma, sWhere

	Session("numResults")=1000
	sWhere=" WHERE a.activo='No' AND p.rentaviv<>0 AND a.fechaultimamodificacion < DATEADD(DAY,-365,GETDATE()) AND (a.fechaavisobaja IS NULL OR a.fechaavisobaja < DATEADD(DAY,-7,GETDATE()))"
	
	sQuery= "SELECT COUNT(*) AS numPisos FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere
	rst.Open sQuery, sProvider
	numPisos=rst("numPisos")
	rst.Close

	sQuery= "SELECT a.activo AS aActivo, a.id AS aId, * FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere
	rst.Open sQuery, sProvider
%>
<form method=post name=frm action=NuAdminPisosInactivosGraba.asp>
<div class=container>
	<div class=banner>Pisos Inactivos</div>
	<nav>
    	<a title='Acceso a la tabla completa de Pisos' href=NuAdminPisosGrid.asp>Tabla Pisos</a>
		<div><a title='Generar Excel con la tabla completa de pisos' href='NuGrabaExcel.asp?tabla=Pisos'>Excel Pisos</a></div></nav>
<div>
<%	if numPisos=0 then %>
	Nada viejo Carretier.
<%
		Response.End
	else
		PagCabeza numPisos, "Piso"
	end if 
%>
	</div>
<div><input title='Mandar los Mails y Borrar' type=submit value='Grabar Cambios' class=btnForm /><% PagPie numPisos, "NuAdminPisosInactivos.asp?" %></div>
<div>
	<div title='Haz click para ver el piso'>Dónde</div><div title='¿Piso Activo?'>Activo</div><div title='Pulsa para mandar mail al pavo'>Quién</div><div title='Teléfono principal'>Teléfonos</div>
    <div title='Precio por mes/semana'>Precio</div><div title='Dia en el que se introdujo el piso'>Fecha Alta</div><div title='Dia de la última revisión'>Fecha Renovación</div>
    <div title='Haz click para ver fotos'>Foto? </div><div title='¿Dónde está?'>Ciudad</div><div title='Marca para incluir en envío'>Mandar Mail Aviso</div><div title='Avisar y Borrar'>Borrar</div></div>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
			if UCase(rst("idiomaes"))="ON" then 
				if UCase(rst("idiomaen"))="ON" then sIdioma="b" else sIdioma="e" 
			else 
				sIdioma="i"
			end if
%>
    <div>
	    <div title='Haz click para ver el piso'><a href='javascript:detalle("TrCasaDetalles.asp?id=<%=rst("aId")%>")'><%=rst("zona")%></a></div><div title='¿Piso Activo?'><%=rst("aActivo")%></div>
		<div title='Pulsa para mandar mail al pavo'><a href='mailto:<%=rst("email")%>'><%=rst("email")%></a></div><div title='Teléfono principal'><%=rst("tel1")%></div>
		<div title='Precio por mes/semana'>€ <%=rst("rentaviv")%></div><div title='Dia en el que se introdujo el piso'><%=rst("fechaalta")%></div>
		<div title='Dia de la última revisión'><%=rst("fechaultimamodificacion")%></div>
		<div title='Haz click para ver fotos'>
<%			if rst("foto") = "Si" then %>
		    <a href="javascript:foto('<%=rst("aId")%>')"><img src=images/Camara.gif alt='Pulsa para ver las fotos'/></a>
<%			elseif rst("foto") <> "" then %>		
			<a href="javascript:foto('<%=rst("aId")%>')"><img src='http://domoh.com/mini<%=rst("foto")%>' alt='Pulsa para ver las fotos'/></a>
<%			end if %>    
			</div>
		<div title='¿Dónde está?'><%=rst("cabecera")%></div>
<%			if IsNull(rst("fechaavisobaja")) then %>
		<div title='Marca para incluir en envío'><input type=checkbox name=m<%=sIdioma & rst("aId")%> /></div>
<%			else %>
		<div title='Marca para incluir en envío'><%=rst("fechaavisobaja")%></div><div title='Avisar y Borrar'><input type=checkbox name=b<%=sIdioma & rst("aId")%> /></div>
<%			end if %>
		</div>
<%
		end if
		rst.Movenext
	next
%>				
	<div><% PagPie numPisos, "NuAdminPisosInactivos.asp?" %></div>
    <div><input title=Grabar type=submit value='Grabar Cambios' class=btnForm /></div>
</div></form>	
<!-- #include file="IncPie.asp" -->