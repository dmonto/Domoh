<%
dim rstDet

'---- Mesg devuelve [vEs] o [vEn] según el valor de [vIdioma]
Function Mesg(vIdioma, vEs, vEn)
	if vIdioma="" then Mesg=vEs else Mesg=vEn
End Function

'---- Banderas genera el html de las banderitas leyendo de [vRst] y con idioma [vIdioma]
Function Banderas(vRst, vIdioma)
	dim s
	
	if UCase(vRst("idiomaes"))="ON" then s=s & "<img src=http://domoh.com/images/Espanol.gif title='" & Mesg(vIdioma,"Anuncio en Español","Description in Spanish") & "'/>"
	if UCase(vRst("idiomaen"))="ON" then s=s & "<img src=http://domoh.com/images/Ingles.gif title='" & Mesg(vIdioma,"Anuncio en Inglés","Description in English") & "'/>"

	Banderas=s
End Function

'---- CasaDetalles devuelve el HTML de una casa [vId] en [vIdioma] con [vFotos]. [vModo] Obsoleto
Function CasaDetalles(vIdioma, vId, vModo, vFotos)
	dim s, sQuery
				
	set rstDet = rst
	sQuery="SELECT * FROM AnunciosVistos WHERE anuncio=" & vId & " AND fecha='" & Format(Date) & "'"
	rstDet.Open sQuery, sProvider
	if rstDet.Eof then
		rstDet.Close
		sQuery = "INSERT INTO AnunciosVistos (anuncio, fecha, visitas) VALUES (" & vId & ", '" & Format(Date) & "', 1)"
	else
		rstDet.Close
		sQuery = "UPDATE AnunciosVistos SET visitas=visitas+1 WHERE anuncio=" & vId & " AND fecha='" & Format(Date) & "'"
	end if 
	rstDet.Open sQuery, sProvider
	
	sQuery="SELECT p.tipo AS pTipo, p.fuma AS pFuma, p.mascota AS pMascota, p.dir1 AS pDir1, p.dir2 AS pDir2, p.cp AS pCP, a.id AS aId, p.ciudad AS pCiudad, a.descripcion AS aDescripcion, * "
	sQuery=sQuery & " FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario WHERE a.id=" & vId
	rstDet.Open sQuery, sProvider
	s="<div class=contanuncio>"
    s=s & "<div title='Ciudad del Piso' class=anciudad>" & rstDet("ciudadnombre") & Banderas(rstDet, vIdioma) & "</div><div class=anzona title=Zona>" & rstDet("zona") & "</div>"
	s=s & "<div class=antipo title=Tipo>"

	if rstDet("pTipo")="Piso" then
		s=s & Mesg(vIdioma,"Piso","Flat")
	elseif rstDet("pTipo")="Habitación" then
		s=s & Mesg(vIdioma,"Habitación","Room")
	elseif rstDet("pTipo")="Casa" then
		s=s & Mesg(vIdioma,"Casa","House")
	else
		s=s & rstDet("pTipo")
	end if

	if rstDet("pTipo")="Hotel" then 
        s=s & "<img title='Es un Hotel' src=http://domoh.com/images/Hotel.gif />"
    end if
	s=s & "</div>"
	s=s & "<div class=anprecio title='Precio'>"
	if rstDet("rentaviv")=0 then 
		if rstDet("rentavacas")=0 then
			if rstDet("precio")=0 then s=s & Mesg(vIdioma,"Para Intercambio","For Swapping") else s=s & rstDet("precio") & Mesg(vIdioma," mil Eur"," thousand Eur")
		else
			s=s & rstDet("rentavacas") & Mesg(vIdioma," Eur/semana"," Eur/Week")
		end if
	else 
		s=s & rstDet("rentaviv") & Mesg(vIdioma," Eur/mes"," Eur/Month")
		if rstDet("rentavacas")<>0 then s=s & rstDet("rentavacas") & Mesg(vIdioma," Eur/semana"," Eur/Week")
	end if 
	s=s & "</div>"
	s=s & "<div class=anfuma title='" & Mesg(vIdioma,"¿Fumadores?","Smokers?") & "'>"
	if UCase(rstDet("pFuma"))<>"ON" then s=s & Mesg(vIdioma,"No Aceptan Fumadores","Smokers Not Allowed") else s=s & Mesg(vIdioma,"Aceptan Fumadores","Smokers Allowed")
	s=s & "</div></div>"
	s=s & "<div class=anmascota title='" & Mesg(vIdioma,"¿Mascotas?","Pets?") & "'>"
	if UCase(rstDet("pMascota"))<>"ON" then s=s & Mesg(vIdioma,"No Aceptan Mascotas","Pets Not Allowed") else s=s & Mesg(vIdioma,"Aceptan Mascotas","Pets Allowed")
	s=s & "</div>"
	s=s & "<div class=andesc title='" & Mesg(vIdioma,"Ciudad y Zona","City/Area") & "'>" & rstDet("aDescripcion") & "</div>"
    s=s & "<div title='" & Mesg(vIdioma,"Quién vive en la casa","Who is already there") & "' colspan=7>" & rstDet("gente") & "</div>" & DetallesContacto(vIdioma)
	
	if UCase(vFotos)="ON" then
	    rstDet.Close
		rstDet.Open "SELECT * FROM Fotos WHERE piso=" & vId, sProvider
		s=s & "<div title='Fotos del Piso'>"
		while not rstDet.EOF		
		    s=s & "<img alt='Foto' src='http://domoh.com/" & rstDet("foto") & "'/>"
			rstDet.MoveNext
		wend
		s=s & "</div>"
	end if

	s=s & "</div></div>"
	rstDet.Close
   	CasaDetalles=s
End Function

'---- TrabajoDetalles devuelve el HTML de un trabajo [vId] en [vIdioma]
Function TrabajoDetalles(vIdioma, vId)
	dim s
	on error goto 0
					
	set rstDet = rst
	sQuery="SELECT p.nombre AS pNombre, ct.nombre" & vIdioma & " AS ctNombre, a.foto AS aFoto, *"
	sQuery=sQuery & " FROM (((Trabajos t INNER JOIN Anuncios a ON t.id=a.id) LEFT JOIN Usuarios u ON a.usuario=u.usuario) LEFT JOIN Provincias p ON a.provincia=p.id) LEFT JOIN CatCurro ct ON t.catcurro=ct.id "
	sQuery=sQuery & " WHERE a.id=" & vId
		
	rstDet.Open sQuery, sProvider
	
	s= "<div><div><div title='Título del Trabajo'>" & rstDet("cabecera") & Banderas(rstDet, vIdioma) & "</div>"
	s=s & "<div title='Categoría del Trabajo'>" & rstDet("ctNombre") & "</div><div>" & rstDet("pNombre") & "</div></div><div>" & rstDet("descripcion") & "</div>"
	s=s & DetallesContacto(vIdioma) & DetallesFoto() & "</div></div>"
	rstDet.Close
   
	TrabajoDetalles=s
End Function

'---- CurriDetalles devuelve el HTML de un CV [vId] en [vIdioma]
Function CurriDetalles(vIdioma, vId)
	dim s
	on error goto 0
					
	set rstDet = rst
	sQuery="SELECT p.nombre AS pNombre, ct.nombre" & vIdioma & " AS ctNombre, a.foto AS aFoto, * FROM "
	sQuery=sQuery & " (((Curriculums c INNER JOIN Anuncios a ON c.id=a.id) LEFT JOIN Usuarios u ON a.usuario=u.usuario) LEFT JOIN Provincias p ON a.provincia=p.id) LEFT JOIN CatCurro ct ON c.catcurro=ct.id "
	sQuery=sQuery & " WHERE a.id=" & vId
		
	rstDet.Open sQuery, sProvider
	s= "<div><div><div>" & rstDet("cabecera") & Banderas(rstDet, vIdioma) & "</div>"
	s=s & "<div>" & rstDet("ctNombre") & "<div>" & rstDet("pNombre") & "</div></div><div>" & rstDet("descripcion") & "</div>"
	s=s & "<div>" & rstDet("cv") & "</div>" & DetallesContacto(vIdioma) & DetallesFoto() & "</div></div>"
	rstDet.Close
	
	CurriDetalles=s
End Function

'---- MiscDetalles devuelve el HTML de un Anuncio diverso [vId] en [vIdioma]
Function MiscDetalles(vIdioma, vId)
	dim s
	on error goto 0
					
	set rstDet = rst
	sQuery= "SELECT p.nombre AS pNombre, cm.nombre" & vIdioma & " AS cmNombre, a.foto AS aFoto, * FROM "
	sQuery=sQuery & " (((Misc m INNER JOIN Anuncios a ON m.id=a.id) LEFT JOIN Usuarios u ON a.usuario=u.usuario) LEFT JOIN Provincias p ON a.provincia=p.id) LEFT JOIN CatMisc cm ON m.catmisc=cm.id WHERE a.id=" & vId
		
	rstDet.Open sQuery, sProvider
	s= "<div><div><div>" & rstDet("cabecera") & Banderas(rstDet, vIdioma) & "</div>"
	s=s & "<div>" & rstDet("cmNombre") & "</div><div>" & rstDet("pNombre") & "</div></div><div>" & rstDet("descripcion") & "</div>" & DetallesContacto(vIdioma) & DetallesFoto()
	if UCase(rstDet("flyer"))<>"" then s=s & "<div><img src='http://domoh.com/" & rstDet("flyer") & "'></div></div></div>"

	rstDet.Close
	MiscDetalles=s
End Function

'---- InquilinoDetalles devuelve el HTML de un Inquilino [vId] en [vIdioma]
Function InquilinoDetalles(vIdioma, vId)
	dim s
	on error goto 0
					
	set rstDet = rst
	sQuery= "SELECT p.nombre AS pNombre, i.ciudad AS iCiudad, i.maximo AS iMaximo, a.foto AS aFoto, * FROM "
	sQuery=sQuery & " ((Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) LEFT JOIN Usuarios u ON a.usuario=u.usuario) LEFT JOIN Provincias p ON a.provincia=p.id WHERE a.id=" & vId

	rstDet.Open sQuery, sProvider
	s= "<div><div><div>" & rstDet("cabecera") & Banderas(rstDet, vIdioma) & "</div>"
	s=s & "<div>" & rstDet("iCiudad") & "<div>" & rstDet("iMaximo") & "<div>" & rstDet("pNombre") & "</div></div><div>" & rstDet("descripcion") & "</div>"
	s=s & DetallesContacto(vIdioma) & DetallesFoto() & "</div></div>"
   
	rstDet.Close
	InquilinoDetalles=s
End Function

'---- DetallesContacto devuelve el HTML del contacto de rstDet en [vIdioma]
Function DetallesContacto(vIdioma)
	dim s
	
	s="<div><div>Identificador: " & rstDet("id") & "</div></div>"
	s=s & "<div><div>"
	if rstDet("esagencia")<>"" then s=s & Mesg(vIdioma,"Agencia","Agency") & ":" else s=s & Mesg(vIdioma,"Contacto","Contact") & ":" 
	s=s & " " &	rstDet("nombre") & "</div></div>"
	s=s & "<div><div>" & rstDet("dir1") & " " & rstDet("dir2") 
	if rstDet("dir1")<>"" or rstDet("dir2")<>"" then s=s & " - " 
	s=s & rstDet("cp") & " " & rstDet("ciudaddir") & "</div></div>"
	s=s & "<div><div>"
	if (UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel1")<>0) or (UCase(rstDet("mostrartel2")) = "ON" and rstDet("tel2")<>0) or (UCase(rstDet("mostrartel3")) = "ON" and rstDet("tel3")<>0) _
        or (UCase(rstDet("mostrartel4")) = "ON" and rstDet("tel4")<>0) then
		s=s & "<img src='http://domoh.com/images/Telefonillo.gif' class=telefono"
	end if
	if UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel1")<>0 then s=s & rstDet("tel1")
	s=s & "<div>"
	if UCase(rstDet("mostrartel2")) = "ON" and rstDet("tel2")<>0 then s=s & rstDet("tel2") & " </div>"
	s=s & "<div>"
	if UCase(rstDet("mostrartel3")) = "ON" and rstDet("tel3")<>0 then s=s & rstDet("tel3") & " </div>"
	s=s & "<div>"
	if UCase(rstDet("mostrartel4")) = "ON" and rstDet("tel4")<>0 then s=s & rstDet("tel4") & " </div></div>"
	s=s & "<div>" & rstDet("instrucciones") & "</div>"
	s=s & "<div><div>"
	if UCase(rstDet("mostraremail")) = "ON" then s=s & "e-Mail: <a href=mailto:" & rstDet("email") & ">" & rstDet("email") & "</a></div>"
	
	DetallesContacto=s
End Function

'---- DetallesFoto devuelve el HTML de la foto de rstDet
Function DetallesFoto()
	dim s, sFoto
	
	sFoto=rstDet("aFoto")
	if sFoto="" then sFoto=rstDet("foto")
	if sFoto<>"" then s=s & "<div><img class=bordeazul src='http://domoh.com/" & sFoto & "'/></div>"
	DetallesFoto=s
End Function
%>