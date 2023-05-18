<%
	dim rstDet

'---- Mesg devuelve [vEs] o [vEn] según el valor de [vIdioma]
Function Mesg(vIdioma, vEs, vEn)
	if vIdioma="" then Mesg=vEs else Mesg=vEn
End Function

'---- Banderas genera el html de las banderitas leyendo de [vRst] y con idioma [vIdioma]
Function Banderas(vRst, vIdioma)
	dim s
	
	if UCase(vRst("idiomaes"))="ON" then s=s & "<img src=http://domoh.com/images/Espanol.gif width=20 title='" & Mesg(vIdioma,"Anuncio en Español","Description in Spanish") & "'>"
	if UCase(vRst("idiomaen"))="ON" then s=s & "<img src=http://domoh.com/images/Ingles.gif width=20 title='" & Mesg(vIdioma,"Anuncio en Inglés","Description in English") & "'>"

	Banderas=s
End Function

'---- CasaDetalles devuelve el HTML de una casa [vId] en [vIdioma] con [vFotos]. [vModo] Obsoleto
Function CasaDetalles(vIdioma, vId, vModo, vFotos)
	dim s, sQuery
				
	set rstDet = Server.CreateObject("ADODB.recordset")
	sQuery="SELECT p.tipo AS pTipo, p.fuma AS pFuma, p.mascota AS pMascota, p.dir1 AS pDir1, p.dir2 AS pDir2, p.cp AS pCp, a.id AS aId, p.ciudad AS pCiudad, a.descripcion AS aDescripcion, a.foto AS aFoto, * "
	sQuery=sQuery & " FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario WHERE a.id=" & vId
	rstDet.Open sQuery, sProvider

	if rstDet.Eof then 
		CasaDetalles="El Anuncio ya no está disponible"
		Exit Function
	end if

	s= "<div><table border=0 class=bordeBurdeos><tr><td><table border=0>"
	s=s & "<tr><td title='Ciudad' class=anuncios3_cab><b>" & rstDet("ciudadnombre") & "</b></td><td title='Zona' class=anuncios3_cab>" & rstDet("zona") & "</td><td title='Qué es' class=anuncios3_cab>"

	if rstDet("pTipo")="Piso" then
		s=s & MesgS("Piso","Flat")
	elseif rstDet("pTipo")="Habitación" then
		s=s & MesgS("Habitación","Room")
	elseif rstDet("pTipo")="Casa" then
		s=s & MesgS("Casa","House")
	else
		s=s & rstDet("pTipo")
	end if
	
	if rstDet("pTipo")="Hotel" then 
        s=s & "<br/>"
        s=s & "<img title='Es un Hotel' src=http://domoh.com/images/Hotel.gif />"
    end if
	s=s & "</td>"

	s=s & "<td title=Precio class=anuncios3_cab><b>"
	if rstDet("rentaviv")=0 then 
		if rstDet("rentavacas")=0 then
			if rstDet("precio")=0 then s=s & MesgS("Para Intercambio","For Swapping") & "</b>" else s=s & rstDet("precio") & "</b>" & MesgS(" mil Eur"," thousand Eur")
		else
			s=s & rstDet("rentavacas") & "</b>" & MesgS(" Eur/semana"," Eur/Week")
		end if
	else 
		s=s & rstDet("rentaviv") & "</b>" & MesgS(" Eur/mes"," Eur/Month")
		if rstDet("rentavacas")<>0 then s=s & "<br>"& rstDet("rentavacas") & "</b>" & MesgS(" Eur/semana"," Eur/Week")
    	s=s & "</td>"
    end if
	if rstDet("precio")=0 then
		s=s & "<td title='Mascotas'>"
		if UCase(rstDet("pMascota"))<>"ON" then s=s & "<img alt='No Perros' src=http://domoh.com/images/NoPerrillos.gif>" else s=s & "<img src=http://domoh.com/images/Perrillos.gif width=32 height=22>"
		s=s & "</td>"
		s=s & "<td title='Fumadores'>"
		if UCase(rstDet("pFuma"))<>"ON" then 
            s=s & "<img alt='No Fumadores' src=http://domoh.com/images/NoCigarrito.gif width=32 height=22>"
		else
			s=s & "<img src=http://domoh.com/images/Cigarrito.gif width=32 height=22>"
		end if
		s=s & "</td>"
	end if
	s=s & "</tr></table></td></tr>"
	s=s & "<tr><td><table border=0><tr><td>" & rstDet("aDescripcion") & " </td><td rowspan=3>" & DetallesFoto() & " </td></tr>"
	s=s & "<tr><td>&nbsp;</td></tr>"
	s=s & "<tr><td>" & rstDet("gente") & " </td></tr>"
	s=s & "<tr><td><table border=0><tr>" & DetallesContacto(vIdioma) & " </tr>"
	
	if UCase(vFotos)="ON" then
		rstDet.Close
		rstDet.Open "SELECT * FROM Fotos WHERE piso=" & vId, sProvider
		s=s & "<tr><td colspan=6>"
		while not rstDet.EOF		
			s=s & "<img width=100 src='http://domoh.com/" & rstDet("foto") & "'>"
			rstDet.MoveNext
		wend
		s=s & "</td></tr>"
	end if

	s=s & "</tr></table></tr></table></div>"
	rstDet.Close
   	CasaDetalles=s
End Function

'---- TrabajoDetalles devuelve el HTML de un trabajo [vId] en [vIdioma]
Function TrabajoDetalles(vIdioma, vId)
	dim s, sQuery
				
	set rstDet = rst
	sQuery="SELECT p.nombre AS pNombre, a.descripcion AS aDescripcion, u.foto AS aFoto, ct.nombre" & vIdioma & " AS ctNombre, * "
	sQuery=sQuery & " FROM (((Trabajos t INNER JOIN Anuncios a ON t.id=a.id) LEFT JOIN Usuarios u ON a.usuario=u.usuario) "
	sQuery=sQuery & " LEFT JOIN Provincias p ON a.provincia=p.id) LEFT JOIN CatCurro ct ON t.catcurro=ct.id WHERE a.id=" & vId
		
	rstDet.Open sQuery, sProvider
	s= "<div><table border=0 class=bordeBurdeos><tr><td><table border=0>"
	s=s & "<tr><td title=Título class=anuncios3_cab><b>" & rstDet("cabecera") & "</b></td><td title=Categoría class=anuncios3_cab>" & rstDet("ctNombre") & "</td>"
	s=s & "<td title=Provincia class=anuncios3_cab>" & rstDet("pNombre") & "</td></tr>"
	s=s & "<tr><td colspan=3><table border=0><tr><td>" & rstDet("aDescripcion") & " </td>" & DetallesFoto() & " </tr>"
	s=s & "<tr><td>&nbsp;</td></tr>"
	s=s & "<tr><td colspan=2><table border=0><tr>" & DetallesContacto(vIdioma) & "</tr></table></div>"
   
	rstDet.Close   
	TrabajoDetalles=s
End Function

'---- CurriDetalles devuelve el HTML de un curriculum [vId] en [vIdioma]
Function CurriDetalles(vIdioma, vId)
	dim s, sQuery
	on error goto 0
				
	set rstDet = rst
	sQuery="SELECT u.foto AS aFoto, p.nombre AS pNombre, a.descripcion AS aDescripcion, ct.Nombre" & vIdioma & " AS ctNombre, * FROM "
	sQuery=sQuery & " (((Curriculums c INNER JOIN Anuncios a ON c.id=a.id) LEFT JOIN Usuarios u ON a.usuario=u.usuario) LEFT JOIN Provincias p ON a.provincia=p.id) "
	sQuery=sQuery & " LEFT JOIN CatCurro ct ON c.catcurro=ct.id WHERE a.id=" & vId
		
	rstDet.Open sQuery, sProvider
	s= "<div><table border=0 class=bordeBurdeos><tr><td><table border=0>"
	s=s & "<tr><td title='Título' class=anuncios3_cab><b>" & rstDet("cabecera") & "</b></td><td title='Categoría' class=anuncios3_cab>" & rstDet("ctNombre") & "</td>"
	s=s & "<td title='Provincia' class=anuncios3_cab>" & rstDet("pNombre") & "</td></tr>"
	s=s & "<tr><td colspan=3><table border=0><tr><td>" & rstDet("aDescripcion") & "</td>" & DetallesFoto() & "</tr>"
	s=s & "<tr><td>&nbsp;</td></tr>"
	s=s & "<tr><td>" & rstDet("cv") & "</td></tr>"
	s=s & "<tr><td colspan=2><table border=0><tr>" & DetallesContacto(vIdioma) & "             </tr></table>"
	rstDet.Close
   
	CurriDetalles=s
End Function

'---- MiscDetalles devuelve el HTML de un anuncio [vId] en [vIdioma]
Function MiscDetalles(vIdioma, vId)
	dim s, sQuery
				
	set rstDet = rst
	sQuery="SELECT p.nombre AS pNombre, a.descripcion as aDescripcion, a.foto AS aFoto, cm.nombre" & vIdioma & " AS cmNombre, a.foto, * FROM "
	sQuery=sQuery & " (((Misc m INNER JOIN Anuncios a ON m.id=a.id) LEFT JOIN Usuarios u ON a.usuario=u.usuario) LEFT JOIN Provincias p ON a.provincia=p.id) LEFT JOIN CatMisc cm ON m.catmisc=cm.id WHERE a.id=" & vId
		
	rstDet.Open sQuery, sProvider
	
	s= "<div><table border=0 class=bordeBurdeos><tr><td><table border=0><tr><td title='Cabecera' class=anuncios3_cab><b>" & rstDet("cabecera") &"</b></td>"
	s=s & "<td title='Categoría' class=anuncios3_cab>" & rstDet("cmNombre") &"</td><td title='Dónde' class=anuncios3_cab>" & rstDet("pNombre")&"</td></tr>"
	s=s & "<tr><td colspan=3><table border=0><tr><td>" & rstDet("aDescripcion") & " </td>" & DetallesFoto() & " </tr>"
	if UCase(rstDet("flyer"))<>"" then s=s & "<tr><td colspan=6><img width=100 style = ""border: 1 solid cyan""" src='http://domoh.com/" & rstDet("flyer") & "'></td></tr>"
	s=s & "<tr><td>&nbsp;</td></tr>"
	s=s & "<tr><td title='" & Mesg(vIdioma,"Ciudad y Zona","City/Area") & "' colspan=7 style = ""text-align: justify; color: #3333CC""">" & rstDet("aDescripcion") & "</td></tr>"
	s=s & DetallesContacto(vIdioma) & "             </tr></table>"
   
	rstDet.Close
	MiscDetalles=s
End Function

'---- InquilinoDetalles devuelve el HTML de un inquilino [vId] en [vIdioma]
Function InquilinoDetalles(vIdioma, vId)
	dim s, sQuery
	on error goto 0
				
	set rstDet = rst
	sQuery="SELECT a.descripcion AS aDescripcion, u.foto AS aFoto, p.nombre AS pNombre, i.ciudad AS iCiudad, i.maximo AS iMaximo, * FROM "
	sQuery=sQuery & " ((Inquilinos i INNER JOIN Anuncios a ON i.id=a.id) LEFT JOIN Usuarios u ON a.usuario=u.usuario) LEFT JOIN Provincias p ON a.provincia=p.id WHERE a.id=" & vId
	rstDet.Open sQuery, sProvider
		
	s="<div><table border=0 class=bordeBurdeos><tr><td><table border=0><tr><td title='Ciudad' class=anuncios3_cab><b>" & rstDet("cabecera") &"</b></td>"
	s=s & "<td title='Zona' class=anuncios3_cab>" & rstDet("iCiudad") &"</td><td title='Qué es' class=anuncios3_cab>" & rstDet("iMaximo") & "€</td><td class=anuncios3_cab><b>" & rstDet("pNombre") & "</td>"
	s=s & "<td>"
	if UCase(rstDet("Mascota"))<>"ON" then s=s & "<img src='images/NoPerrillos.gif' width=32 height=22 />" else s=s & "<img src='images/Perrillos.gif' width=32 height=22 />"
	s=s & "</td>"
	s=s & "<td>"
	if UCase(rstDet("Fuma"))<>"ON" then s=s & "<img src='images/NoCigarrito.gif' width=32 height=22 />" else s=s & "<img src='images/Cigarrito.gif' width=32 height=22 />"
	s=s & "</td></tr></table></td></tr>"
	s=s & "<tr><td><table border=0><tr><td>" & rstDet("aDescripcion") & "</td><td rowspan=3>" & DetallesFoto() &" </td></tr>"
	s=s & "<tr><td>&nbsp;</td></tr>"
	s=s & "<tr><td><table border=0><tr>" & DetallesContacto(vIdioma) & "             </tr></table>"
   
	rstDet.Close
	InquilinoDetalles=s
End Function

'---- DetallesContacto devuelve el HTML del contacto en [vIdioma]
Function DetallesContacto(vIdioma)
	dim s
	
	s=s & "<table><tr><td>"
	if rstDet("esagencia")<>"" then s=s & "<b>" & Mesg(vIdioma,"Agencia","Agency") & ";</b>" else s=s & "<b>" & Mesg(vIdioma,"Contacto","Contact") & ";</b>" 
    s=s & "<td> " & rstDet("nombre") & "</td></tr>"
	s=s & "<tr><td>&nbsp;</td></tr>"
	s=s & "<tr><td colspan=7>" & rstDet("dir1") & " " & rstDet("dir2") 
	if rstDet("dir1")<>"" or rstDet("dir2")<>"" then s=s & " - " 
	s=s & rstDet("cp") & " " & rstDet("ciudaddir") & "</td></tr>"
	s=s & "<tr>"
	if (UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel1")<>0) or (UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel2")<>0) or (UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel3")<>0) _
        or (UCase(rstDet("mostrartel4")) = "ON" and rstDet("tel4")<>0) then 
		s=s & "<td><b>" & Mesg(vIdioma,"Telef","Phone") & ";</b><td><img src=http://domoh.com/images/telefonillo.gif width=11 height=16 />"
	end if 
	if UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel1")<>0 then s=s & rstDet("tel1") & " "
	if UCase(rstDet("mostrartel2")) = "ON" and rstDet("tel2")<>0 then s=s & rstDet("tel2") & " "
	if UCase(rstDet("mostrartel3")) = "ON" and rstDet("tel3")<>0 then s=s & rstDet("tel3") & " "
	if UCase(rstDet("mostrartel4")) = "ON" and rstDet("tel4")<>0 then s=s & rstDet("tel4") & " "
	if rstDet("instrucciones")<>"" then s=s & "(" & rstDet("instrucciones") & ")"
	s=s & "</td></tr>"
	s=s & "<tr>"
	if UCase(rstDet("mostraremail")) = "ON" then s=s & "<td><b>eMail;</b></td><td><a href=mailto:" & rstDet("email") & ">" & rstDet("email") & "</a></td>"
	s=s & "</tr></table>"
	
	DetallesContacto=s
End Function

'---- DetallesFoto devuelve el HTML de la foto en rstDet
Function DetallesFoto()
	dim s
	if UCase(rstDet("aFoto"))<>"" then s=s & "<td rowspan=3><img src='http://domoh.com/" & rstDet("aFoto") &"' width=180></td>"
	DetallesFoto=s
end function
%>