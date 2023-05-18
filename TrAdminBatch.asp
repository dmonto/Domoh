<!-- #include file="IncNuBD.asp" -->
<% 
	Response.Buffer=false 
	Server.ScriptTimeout=1000
%>
<!-- #include file="IncTrGrid.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%
	dim Grid, sWhere, bEjec, bExiste, rstExec, dFecha, numViejo
	on error goto 0
	
	set rstExec = Server.CreateObject("ADODB.recordset")
	set fso = CreateObject("Scripting.FileSystemObject")
	if Request("op")="Ejecutar" then bEjec=3

	Response.Write "<h1>Usuarios Duplicados</h1>"
	sQuery="SELECT MIN(id) AS viejoId, usuario, COUNT(*) FROM Usuarios GROUP BY usuario HAVING COUNT(*)>1"
	rst.Open sQuery, sProvider
	if not rst.Eof then
		numViejo=rst("viejoId")	
		rst.Close
		sQuery="DELETE FROM Usuarios WHERE id=" & numViejo
		Response.Write "<h2><a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a>"
		if bEjec then
			rst.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
        Response.Write "</h2>"
	else
		rst.Close	
	end if
	
	Response.Write "<h1>Pisos Duplicados</h1>"
	sQuery="SELECT MIN(id) AS viejoId, COUNT(*) FROM Pisos GROUP BY id HAVING COUNT(*)>1"
	rst.Open sQuery, sProvider
	if not rst.Eof then
		numViejo=rst("viejoId")	
		rst.Close
		Response.Write "<h2>Tienes que borrar el piso " & numViejo & "</h2>"
	else
		rst.Close	
	end if

	Response.Write "<h1>Fotos sin Padre</h1>"
	sQuery="SELECT * FROM Fotos WHERE piso NOT IN (SELECT id FROM Pisos)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Fotos WHERE id=" & rst("id") 
		Response.Write "<a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Trabajos sin Categoria</h1>"
	sQuery="SELECT * FROM Trabajos WHERE catcurro NOT IN (SELECT id FROM CatCurro)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Trabajos WHERE id=" & rst("id") 
		Response.Write "<a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Usuarios Innecesarios</h1>"
	sQuery="SELECT * FROM Usuarios WHERE usuario IS NULL OR usuario IN ('a','s','d','f','g','h','j','k','l','ñ','ç') "
	sQuery=sQuery & " OR password IS NULL OR activo='Mail' AND fechaalta < DATEADD(DAY,-30,GETDATE()) OR numanuncios=0 AND ultimavisita < DATEADD(DAY,-365,GETDATE())"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Usuarios WHERE id=" & rst("id") 
		Response.Write "<h2><a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a></h2>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Usuarios Sin Anuncios</h1>"
	sQuery="SELECT * FROM Usuarios WHERE numanuncios<>0 AND usuario NOT IN (SELECT usuario FROM Anuncios)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="UPDATE Usuarios SET numanuncios=0 WHERE id=" & rst("id") 
		Response.Write "<a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & "></a>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Pisos Sin Padre</h1>"
	sQuery="SELECT * FROM Pisos WHERE id NOT IN (SELECT id FROM Anuncios)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Pisos WHERE id=" & rst("id") 
		Response.Write "<h2><a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a></h2>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Precios Absurdos</h1>"
	sQuery="SELECT * FROM Pisos WHERE precio>100000 OR rentavacas>10000"
	rst.Open sQuery, sProvider
	while not rst.Eof
		Response.Write "Precio de piso " & rst("id") & ": " & rst("precio") & "/" & rst("rentavacas") 
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Misc sin Padre</h1>"
	sQuery="SELECT * FROM Misc WHERE id NOT IN (SELECT id FROM Anuncios)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Misc WHERE id=" & rst("id") 
		Response.Write "<a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & "></a>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Inquilinos sin Padre</h1>"
	sQuery="SELECT * FROM Inquilinos WHERE id NOT IN (SELECT id FROM Anuncios)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Inquilinos WHERE id=" & rst("id") 
		Response.Write "<a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Curriculums sin Padre</h1>"
	sQuery="SELECT * FROM Curriculums WHERE id NOT IN (SELECT id FROM Anuncios)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Curriculums WHERE id=" & rst("id") 
		Response.Write "<a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Curriculums sin Categoria</h1>"
	sQuery="SELECT * FROM Curriculums WHERE catcurro NOT IN (SELECT id FROM CatCurro)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Curriculums WHERE id=" & rst("id") 
		Response.Write "<a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Comentarios Viejos</h1>"
	sQuery="SELECT * FROM AnuncioComent WHERE LTRIM(CAST(texto AS VARCHAR))='' OR anuncio NOT IN (SELECT id FROM anuncios) OR fechaalta < DATEADD(DAY,-365,GETDATE())"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM AnuncioComent WHERE id=" & rst("id") 
		Response.Write "<h2><a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a></h2>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close
	
	Response.Write "<h1>Anuncios sin Usuario</h1>"
	sQuery="SELECT * FROM Anuncios WHERE usuario NOT IN (SELECT usuario FROM usuarios)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Anuncios WHERE id=" & rst("id") 
		Response.Write "<h2><a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a></h2>"
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Fotos de Anuncios</h1>"
	sQuery="SELECT * FROM Anuncios WHERE foto IS NOT NULL AND foto<>''"
	rst.Open sQuery, sProvider
	on error resume next
	while not rst.Eof
		bExiste=False
		bExiste=fso.FileExists(Server.MapPath(rst("foto")))
		if not bExiste then 
			Response.Write "Anuncio " & rst("id") & ": FALTA Foto " & rst("foto") 
		end if 
		
		bExiste=False
		bExiste=fso.FileExists(Server.MapPath("mini" & rst("foto")))
		if not bExiste Then 
			Response.Write "Anuncio " & rst("id") & ": FALTA Foto mini" & rst("foto") 
		end if 

		rst.Movenext
	wend	
	rst.Close
	on error goto 0

	Response.Write "<h1>Anuncios sin Publicar</h1>"
	sQuery="SELECT * FROM Anuncios WHERE provincia NOT IN (SELECT id FROM Provincias)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Anuncios WHERE id=" & rst("id") 
		Response.Write "<h2><a target=bd href=NuAdminBD.asp?sent=" & Server.URLEncode(sQuery) & ">" & sQuery & "</a>: " & rst("fechaultimamodificacion")  & "</h2>"
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Fotos que faltan</h1>"
	sQuery="SELECT * FROM Fotos WHERE foto IS NOT NULL AND foto<>''"
	rst.Open sQuery, sProvider
	on error resume next
	while not rst.Eof
		bExiste=False
		bExiste=fso.FileExists(Server.MapPath(rst("foto")))
		if not bExiste then 
			Response.Write "<h2>Foto " & rst("id") & ": FALTA Foto " & rst("foto")  & "</h2>"
		end if 
		
		bExiste=False
		bExiste=fso.FileExists(Server.MapPath("mini" & rst("foto")))
		if not bExiste then 
			Response.Write "Foto " & rst("id") & ": FALTA Foto mini" & rst("foto") 
		end if 

		rst.Movenext
	wend	
	rst.Close
	on error goto 0

	Response.Write "<h1>Fotos de Usuarios</h1>"
	sQuery="SELECT * FROM Usuarios WHERE foto IS NOT NULL AND foto<>''"
	rst.Open sQuery, sProvider
	on error resume next
	while not rst.eof
		bExiste=False
		bExiste=fso.FileExists(Server.MapPath(rst("foto")))
		if not bExiste Then 
			Response.Write "Usuario " & rst("id") & ": FALTA Foto " & rst("foto") 
		end if 
		
		bExiste=False
		bExiste=fso.FileExists(Server.MapPath("mini" & rst("foto")))
		if not bExiste Then 
			Response.Write "Usuario " & rst("id") & ": FALTA Foto mini" & rst("foto") 
		end if 

		rst.Movenext
	wend	
	rst.Close
	on error goto 0

	Response.Write "<h1>Anuncios sin Piso</h1>"
	sQuery="SELECT * FROM Anuncios WHERE tabla='Pisos' AND id NOT IN (SELECT id FROM Pisos)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Anuncios WHERE id=" & rst("id") 
		Response.Write sQuery 
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Anuncios sin Inquilino</h1>"
	sQuery="SELECT * FROM Anuncios WHERE tabla='Inquilinos' AND id NOT IN (SELECT id FROM Inquilinos)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Anuncios WHERE id=" & rst("id") 
		Response.Write sQuery 
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Anuncios sin Curriculum</h1>"
	sQuery="SELECT * FROM Anuncios WHERE tabla='Curriculums' AND id NOT IN (SELECT id FROM Curriculums)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Anuncios WHERE id=" & rst("id") 
		Response.Write sQuery 
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Anuncios sin Misc</h1>"
	sQuery="SELECT * FROM Anuncios WHERE tabla='Misc' AND id NOT IN (SELECT id FROM Misc)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Anuncios WHERE id=" & rst("id") 
		Response.Write sQuery 
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Anuncios sin Trabajo</h1>"
	sQuery="SELECT * FROM Anuncios WHERE tabla='Trabajos' AND id NOT IN (SELECT id FROM Trabajos)"
	rst.Open sQuery, sProvider
	while not rst.Eof
		sQuery="DELETE FROM Anuncios WHERE id=" & rst("id") 
		Response.Write sQuery 
		if bEjec then
			rstExec.Open sQuery, sProvider
			Response.Write "->Hecho!!"
			bEjec=bEjec-1
		end if
		rst.Movenext
	wend	
	rst.Close

	Response.Write "<h1>Fotos sin Utilizar</h1>"
	dim fs,f, fi, i
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	set f=fs.GetFolder(Server.MapPath("/fotos"))
	for each fi in f.Files
		i=i+1
		if i=10 then
			Response.Write "."
			i=0
		end if
		sQuery="SELECT foto FROM Anuncios WHERE UPPER(foto)='FOTOS/" & UCase(fi.Name) & "'"
		rst.Open sQuery, sProvider		
		if rst.Eof then 
			rst.Close
			sQuery="SELECT foto FROM Fotos WHERE UPPER(foto)='FOTOS/" & UCase(fi.Name) & "'"
			rst.Open sQuery, sProvider		
			if rst.Eof then 
				rst.Close
				sQuery="SELECT foto FROM Usuarios WHERE UPPER(foto)='FOTOS/" & UCase(fi.Name) & "'"
				rst.Open sQuery, sProvider		
				if rst.Eof then 
                    Response.Write "<h2>Borrando " & f & "/" & fi.Name & "</h2>"
                    fs.DeleteFile f & "/" & fi.Name
                end if
				rst.Close
			end if
		end if 
		if rst.State then rst.Close
	next

	Response.Write "<h1>MiniFotos sin Utilizar</h1>"
	set f=fs.GetFolder(Server.MapPath("/minifotos"))
	for each fi in f.Files
		sQuery="SELECT foto FROM Anuncios WHERE UPPER(foto)='FOTOS/" & UCase(fi.Name) & "'"
		rst.Open sQuery, sProvider		
		if rst.Eof then 
			rst.Close
			sQuery="SELECT foto FROM Fotos WHERE UPPER(foto)='FOTOS/" & UCase(fi.Name) & "'"
			rst.Open sQuery, sProvider		
			if rst.Eof then 
				rst.Close
				sQuery="SELECT foto FROM Usuarios WHERE UPPER(foto)='FOTOS/" & UCase(fi.Name) & "'"
				rst.Open sQuery, sProvider		
				if rst.Eof then 
                    Response.Write "Borrando " & f & "/" & fi.Name 
                    fs.DeleteFile f & "/" & fi.Name
                end if
				rst.Close
			end if
		end if 
		if rst.State then rst.Close
	next

	sQuery="SELECT MIN(fecha) AS minFecha FROM AnunciosVistos WHERE anuncio<>0"
	rst.Open sQuery, sProvider
	dFecha=rst("minFecha")
	rst.Close
	sQuery="INSERT INTO AnunciosVistos (anuncio,fecha,visitas) SELECT 0, fecha, SUM(visitas) FROM AnunciosVistos WHERE fecha < '" & Format(dFecha+7) & "' AND anuncio <> 0 GROUP BY fecha "
	Response.Write sQuery 
	if bEjec then
		rst.Open sQuery, sProvider
		Response.Write "->Hecho!!"
		bEjec=bEjec-1
	end if
	sQuery="DELETE FROM AnunciosVistos WHERE fecha < DATEADD(DAY,7,'" & Format(dFecha) & "') AND anuncio <> 0"
	Response.Write sQuery 
	if bEjec then
		rst.Open sQuery, sProvider
		Response.Write "->Hecho!!"
		bEjec=bEjec-1
	end if
%>
<h1><a title='Hacer cambios en BD' href='?op=Ejecutar'>Ejecutar</a></h1>