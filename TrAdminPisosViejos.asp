<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<%
	dim numPisos, i, sIdioma, sWhere
	on error goto 0
	
	Session("numResults")=1000
	sWhere=" WHERE a.activo='Si' AND u.activo='Si' AND (a.fechafindestacado< GETDATE() OR (p.rentaviv+p.precio<>0 AND a.fechaultimamodificacion < DATEADD(DAY,-31,GETDATE()) "
	sWhere=sWhere & " AND (a.fechaavisobaja IS NULL OR a.fechaavisobaja < DATEADD(DAY,-2,GETDATE())))"
	sWhere=sWhere & " OR (p.rentavacas<>0 AND a.fechaultimamodificacion < DATEADD(DAY,-365,GETDATE()) AND (a.fechaavisobaja IS NULL OR a.fechaavisobaja < DATEADD(DAY,-365,GETDATE()))))"
	
	sQuery= "SELECT COUNT(*) AS numPisos FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere
	rst.Open sQuery, sProvider
	numPisos=rst("numPisos")
	rst.Close

	sQuery= "SELECT a.activo AS aActivo, a.id AS aId, * FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario " & sWhere
	rst.Open sQuery, sProvider
%>
<head>
    <!-- #include file="IncTrCabecera.asp" -->
    <title>Limpiar Pisos Viejos</title>
</head>
<body>
<form method=post name=frm action=TrAdminPisosViejosGraba.asp>
<div class="container">
	<div class="logo"><h1>Pisos Viejos</h1></div>
	<div class="banner">
		<a title='Acceso a la tabla completa de Pisos' href=NuAdminPisosGrid.asp>Tabla Pisos</a>
		<a title='Generar Excel con toda la tabla de pisos' href='NuGrabaExcel.asp?tabla=Pisos'>Excel Pisos</a></div>
	<div class=main>
		<h2>
<%	if numPisos=0 then %>
    Nada viejo Carretier.
<%
		Response.End
	else
		PagCabeza numPisos, "Piso"
	end if 
%>
		</h2>
		<p><input type=submit value='Grabar Cambios' class=btnLogin /><% PagPie numPisos, "TrAdminPisosViejos.asp?" %></p>
<div class=row>
    <div class='middle anuncios2_cab' title='Qué es'>Tipo</div><div class=middle title='Población'>Dónde</div><div class=middle title='¿Está Publicado?'>Activo</div><div class=middle title='Pulsa para mandar mail'>Quién</div>
    <div class=middle title='Teléfonos del Usuario'>Teléfonos</div><div class=middle title='Precio por mes/semana/total'>Precio</div><div class=middle title='Día que se creó el anuncio'>Fecha Alta</div>
    <div class=middle title='Última Modificación'>Fecha Renovación</div><div class=middle title='Día que deja de estar destacado'>Destacado Hasta</div><div class=middle title='Haz click para ver las fotos'>Foto? </div>
    <div class=middle title='Ciudad y Zona'>Ciudad </div><div class=middle title='Solo mandaremos mails a los marcados'>Mandar Mail Aviso</div><div class=middle title='Solo desactivaremos los marcados'>Desactivar</div>
    <div class=middle title='Pasaremos los marcados a no destacados'>DesDestacar</div></div>
<%		
	for i=1 to RegHasta
		if i >= RegDesde then
			if UCase(rst("idiomaes"))="ON" then sIdioma="e" else sIdioma="i" 
%>
<div class=row>
	<div class=middle title='Qué es'> 
<%		    if rst("rentaviv")<>0 then %>
			Vivir
<% 		    elseif rst("precio")<>0 then %>
			Vender
<% 		    else %>
			Vacas
<% 		    end if %>
    		</div>
		<div class=middle title='Población'><a href="javascript:detalle('TrAnuncioDetalle.asp?tabla=Pisos&id=<%=rst("aId")%>')"><%=rst("zona")%></a></div><div class=middle title='¿Está Publicado?'><%=rst("aActivo")%></div>
        <div class=middle title='Pulsa para mandar mail'><a href='mailto:<%=rst("email")%>'><%=rst("email")%></a></div><div class=middle title='Teléfonos del Usuario'><%=rst("tel1")%></div>
		<div class=middle title='Precio por mes/semana/total'>
			€
<%		
		    if rst("rentaviv")<>0 then 
			    Response.Write rst("rentaviv")
 		    elseif rst("precio")<>0 then
    			Response.Write rst("precio")
 	    	else 
		    	Response.Write rst("rentavacas")
 		    end if 
%>
			</div>
		<div class=middle title='Día que se creó el anuncio'><%=rst("fechaalta")%></div><div class=middle title='Última Modificación'><%=rst("fechaultimamodificacion")%></div>
        <div class=middle title='Día que deja de estar destacado'><%=rst("fechafindestacado")%></div>
		<div class=middle title='Haz click para ver las fotos'>
<%		    if rst("foto")="Si" then %>
			<a href="javascript:foto('TrFotos.asp?id=<%=rst("id")%>')"><img src=images/Camara.gif alt='Pulsa para ver las fotos'/></a>
<%		    elseif rst("foto") <> "" then %>		
			<a href="javascript:foto('TrFotos.asp?id=<%=rst("id")%>')"><img src="http://domoh.com/mini<%=rst("foto")%>" alt='Pulsa para ver las fotos' /></a>
<%		    end if %>    
			</div>
		<div class=middle title='Ciudad y Zona'><%=rst("cabecera")%></div>
<%		    if not(IsNull(rst("fechaavisobaja"))) then %>
		<div class=middle title='Fecha en la que se avisó al usuario'><%=rst("fechaavisobaja")%></div><div title='Solo desactivaremos los marcados'><input type=checkbox name=b<%=sIdioma & rst("aId")%> /></div>
<%		    elseif rst("rentavacas")>0 then %>
		<div class=middle title='Solo desactivaremos los marcados'><input type=checkbox name=v<%=sIdioma & rst("aId")%> /></div>
<%		    else %>
		<div class=middle title='Solo desactivaremos los marcados'><input type=checkbox name=m<%=sIdioma & rst("aId")%> /></div>
<%			
            end if 
            if rst("destacado")>0 and rst("fechafindestacado")<Now then 
%>
		<div class=middle><input type=checkbox name=d<%=sIdioma & rst("aId")%>/></div>
<%		    end if %>
		</div>
<%
		end if
		rst.Movenext
	next
%>				
	<div class=row><% PagPie numPisos, "TrAdminPisosViejos.asp?" %></div>
	<div class=row><input title='Dele' type=submit value='Grabar Cambios' class=btnForm /></div>
</form>	</div>
<!-- #include file="IncPie.asp" -->
</body>
