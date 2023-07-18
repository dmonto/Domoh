<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim numPisos, i, sProvincias, sURLModif
	
	Session("Usuario")="hector"
	Session("Pantalla")="NuAdminPisos.asp?tipo=" & Request("tipo")
	sQuery= "SELECT COUNT(*) AS numPisos FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario "
	if Request("tipo")="nuevos" then sQuery= sQuery & " WHERE a.activo='Si' AND a.provincia=0 "
	rst.Open sQuery, sProvider
	numPisos=rst("numPisos")
	rst.Close

	rst.Open "SELECT * FROM Provincias ORDER BY nombre", sProvider
	while not rst.Eof
		sProvincias=sProvincias & "<option value=" & rst("id") &">Publicar en " & rst("nombre")
		rst.Movenext
	wend
	rst.Close

	sQuery= "SELECT a.activo AS aActivo, a.id AS aId, u.activo AS uActivo, * FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario "
	if Request("tipo")="nuevos" then sQuery= sQuery & " WHERE a.activo='Si' AND a.provincia=0 "
	rst.Open sQuery, sProvider
%>
<form method=post name=frm action=NuAdminAnunciosGraba.asp>
	<input type=hidden name=tipo value=<%=Request("tipo")%> /><input type=hidden name=tabla value=Pisos /><input type=hidden name=campo value=provincia />
<div class=container>
	<div class=banner><h1>Pisos Nuevos</h1></div>
	<div class=main>
		<h2><a title='Acceso a la tabla completa de Pisos' href='NuAdminPisosGrid.asp?tipo=menu'>Tabla Pisos</a></h2>
		<h2><a href=NuAdminPisos.asp>NuAdminPisos.asp</a><a title='Generar un Excel con toda la tabla de pisos' href='NuGrabaExcel.asp?tabla=Pisos'>Excel Pisos</a></h2>
		<h2>
		<%	if numPisos=0 then %>
			Nada nuevo Carretier.
		<%
				Response.End
			else
				PagCabeza numPisos, "Piso"
			end if 
		%>
		</h2>
		<h2><input title='Atizale para publicar, etc' type=submit value='Grabar Cambios' class=btnForm /><% PagPie numPisos, "NuAdminPisos.asp?" %></h2>
		<div class=gridtable>
			<div class=gridrow>
				<div class=left title='Para qué es el piso'>Tipo</div>
				<div class=left title='Ciudad/Zona'>Dónde</div>
				<div class=left title=¿El tipo lo ha ocultado?'>Piso Activo</div>
				<div class=left title='Haz click para mandarle un e-mail'>Quién</div>
				<div class=left title='Si el usuario ha confirmado o no'>Usuario Activo</div>
				<div class=left title='Cómo contactar con el tipo'>Teléfonos</div>
				<div class=left title='Precio por Mes/Semana/Total'>Precio</div>
				<div class=left title='Día que metieron el anuncio'>Fecha Alta</div>
				<div class=left title='Pulsa para ver las fotos'>Foto?</div>
				<div class=left title='Ciudad que ha introducido el usuario'>Ciudad</div>
				<div class=left title='Publicar, Borrar...'>Qué hago con él</div></div>
			<%		
				for i=1 to RegHasta
					if i >= RegDesde then
			%>
			<div class=gridrow>
				<div class=left title='Para qué es el piso'>
					<a href='javascript:detalle("QuAnuncioDetalle.asp?tabla=Pisos&id=<%=rst("aId")%>")'>
			<%
						if rst("rentaviv")=0 then
							if rst("rentavacas")>0 then 
								Response.Write "Vacaciones"
								sURLModif="TrVacasRegOfrezcoFront"
							end if
						else
							Response.Write "Vivir" 
							sURLModif="TrCasaRegOfrezcoFront"
						end if
			%>
					</a>
			<%	if UCase(rst("idiomaes"))="ON" then %>
					<img src=images/Espanol.gif alt='Anuncio en Español'/>
			<%	
				end if 
				if UCase(rst("idiomaen"))="ON" then 
			%>
					<img src=images/Ingles.gif alt='Anuncio en Inglés'/>
			<%	end if %>
					</div>
				<div class=left title='Ciudad/Zona'><a href='javascript:detalle("QuAnuncioDetalle.asp?tabla=Pisos&id=<%=rst("aId")%>")'><%=rst("cabecera")%></a></div>
				<div class=left title='¿El tipo lo ha ocultado?'><%=rst("aActivo")%></div><div title='Haz click para mandarle un e-mail'><a href='mailto:<%=rst("email")%>'><%=rst("usuario")%></a></div>
				<div class=left title='Si el usuario ha confirmado o no'><%=rst("uActivo")%></div><div title='Cómo contactar con el tipo'><%=rst("tel1")%></div>
				<div class=left title='Precio por Mes/Semana/Total'>¿<%=rst("rentaviv")+rst("rentavacas")+rst("precio")%></div><div title='Día que metieron el anuncio'><%=rst("fechaalta")%></div>
				<div class=left title='Pulsa para ver las fotos'>
			<%  if rst("foto") = "Si" then %>
					<a href="javascript:foto('<%=rst("id")%>')"><img src=images/Camara.gif alt='Pulsa para ver las fotos'/></a>
			<%			elseif rst("foto") <> "" then %>		
					<a href="javascript:foto('<%=rst("id")%>')"><img src="http://domoh.com/mini<%=rst("foto")%>" alt='Pulsa para ver las fotos'/></a>
			<%			end if %>    
					</div>
				<div class=left title='Ciudad que ha introducido el usuario'><%=rst("ciudadnombre")%></div>
				<div class=left title='Publicar, Borrar...'>
					<select name=<%=rst("aId")%> onchange="if(value=='Modificar') location='<%=sURLModif%>.asp?op='+value+'&id=<%=rst("aId")%>';">
						<option selected value=0>-Elige Una-</option>
			<% if rst("aActivo")="Si" then %>
						<option value=Borrar>Ocultar</option>
			<% else %>
						<option value=Reactivar>Reactivar</option>
			<%          end if %>
						<option value=Modificar>Modificar</option>
						<option value=Eliminar>Eliminar Definitivamente</option>
						<%=sProvincias%>
						</select></div></div>
			<%
					end if
					rst.Movenext
				next
			%>				
		<div class=gridrow><% PagPie numPisos, "NuAdminPisos.asp?" %></div>
		<div class=gridrow><input title=Graba type=submit value='Grabar Cambios' class=btnForm /></div>
</div></form>
<!-- #include file="IncPie.asp" -->
