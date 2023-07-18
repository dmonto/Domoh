<%
	dim rstDet

'---- Mesg devuelve [vEs] o [vEn] según el valor de [vIdioma]
Function Mesg(vIdioma, vEs, vEn)
	if vIdioma="" then Mesg=vEs else Mesg=vEn
End Function

'---- Banderas genera el html de las banderitas leyendo de [vRst] y con idioma [vIdioma]
Function Banderas(vRst, vIdioma)
	dim s
	
	if UCase(vRst("idiomaes"))="ON" then s=s & "<img src=http://domoh.com/images/Espanol.gif title='" & Mesg(vIdioma,"Anuncio en Español","Description in Spanish") & "'>"
	if UCase(vRst("idiomaen"))="ON" then s=s & "<img src=http://domoh.com/images/Ingles.gif title='" & Mesg(vIdioma,"Anuncio en Inglés","Description in English") & "'>"

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
		CasaDetalles="<h1>El Anuncio ya no está disponible</h1>"
		Exit Function
	end if

	s= "<div class='anposted'>Publicado</div><div title='Ciudad' class='anciudad'>" & rstDet("ciudadnombre") & "</div><div title='Zona' class='anzona'>" & rstDet("zona") &"</div>"
    s= s & "<div title='Qué es' class='antipo'>"

	if rstDet("pTipo")="Piso" then
		s=s & MesgS("Piso","Flat")
	elseif rstDet("pTipo")="Habitación" then
		s=s & MesgS("Habitación","Room")
	elseif rstDet("pTipo")="Casa" then
		s=s & MesgS("Casa","House")
	else
		s=s & rstDet("pTipo")
	end if
	
	if rstDet("pTipo")="Hotel" then s=s & "<img title='Es un Hotel' src=http://domoh.com/images/Hotel.gif />"	
	s=s & "</div>"
	s=s & "<div title=Precio class='anprecio'>"
	if rstDet("rentaviv")=0 then 
		if rstDet("rentavacas")=0 then
			if rstDet("precio")=0 then s=s & MesgS("Para Intercambio","For Swapping") else s=s & rstDet("precio") & MesgS(" mil Eur"," thousand Eur")
		else
			s=s & rstDet("rentavacas") & MesgS(" Eur/semana"," Eur/Week")
		end if
	else 
		s=s & rstDet("rentaviv") & MesgS(" Eur/mes"," Eur/Month")
		if rstDet("rentavacas")<>0 then 
            s=s & rstDet("rentavacas") & MesgS(" Eur/semana"," Eur/Week")
        end if
	end if 
	s=s & "</div>"

	if rstDet("precio")=0 then
		s=s & "<div title='Mascotas' class='anmascota'>"
		if UCase(rstDet("pMascota"))<>"ON" then 
			s=s & "<img alt='No Perros' src=http://domoh.com/images/NoPerrillos.gif />"
		else
			s=s & "<img alt='Perros Admitidos' src=http://domoh.com/images/Perrillos.gif />"
		end if
		s=s & "</div>"
		s=s & "<div title='Fumadores' class='anfuma'>"
		if UCase(rstDet("pFuma"))<>"ON" then 
			s=s & "<img alt='No Fumadores' src=http://domoh.com/images/NoCigarrito.gif />"
		else
			s=s & "<img alt='Fumadores' src=http://domoh.com/images/Cigarrito.gif />"
		end if
		s=s & "</div>"
	end if
	
	s=s & "<div class='andesc'>" & rstDet("aDescripcion") & "</div><div class='anfoto'>" & DetallesFoto() & "</div><div class='angente'>" & rstDet("gente") & "</div><div class='ancontacto'>" & DetallesContacto(vIdioma) & "</div>"
	
	if UCase(vFotos)="ON" then
		rstDet.Close
		rstDet.Open "SELECT * FROM Fotos WHERE Piso=" & vId, sProvider
		s=s & "<div class='anfotoxtra'>"
		while not rstDet.EOF		
			s=s & "<img class=bordeazul src='http://domoh.com/"& rstDet("foto") & "'/>"
			rstDet.MoveNext
		wend
		s=s & "</div>"
	end if

	rstDet.Close
   	CasaDetalles=s
End Function


'---- DetallesContacto devuelve el HTML común de contacto de rstDet en [vIdioma]
Function DetallesContacto(vIdioma)
	dim s
	
	s = ""
	if rstDet("esagencia")<>"" then s=s & Mesg(vIdioma,"Agencia","Agency") & ": " else s=s & Mesg(vIdioma,"Contacto","Contact") & ": " 
    s=s & rstDet("nombre") & " - " & rstDet("dir1") & " - " & rstDet("dir2") & " - " 
	if rstDet("dir1")<>"" or rstDet("dir2")<>"" then s=s & " - " 
	s=s & rstDet("cp") & " " & rstDet("ciudaddir") & " - " 
	if (UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel1")<>0) or (UCase(rstDet("mostrartel2")) = "ON" and rstDet("tel2")<>0) or (UCase(rstDet("mostrartel3")) = "ON" and rstDet("tel3")<>0) _
        or (UCase(rstDet("mostrartel4")) = "ON" and rstDet("tel4")<>0) then 
		s=s & Mesg(vIdioma,"Telef","Phone") & "<img src=http://domoh.com/images/telefonillo.gif />" & " "
	end if 
	if UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel1")<>0 then s=s & rstDet("tel1") & " "
	if UCase(rstDet("mostrartel2")) = "ON" and rstDet("tel2")<>0 then s=s & rstDet("tel2") & " "
	if UCase(rstDet("mostrartel3")) = "ON" and rstDet("tel3")<>0 then s=s & rstDet("tel3") & " "
	if UCase(rstDet("mostrartel4")) = "ON" and rstDet("tel4")<>0 then s=s & rstDet("tel4") & " "
	if rstDet("instrucciones")<>"" then s=s & "(" & rstDet("instrucciones") & ") "
	if UCase(rstDet("mostraremail")) = "ON" then s=s & "eMail: <a href=mailto:" & rstDet("email") & ">" & rstDet("email") & "</a>"
	
	DetallesContacto=s
End Function

'---- DetallesFoto devuelve el HTML de la foto de rstDet
Function DetallesFoto()
	dim s
	
	if UCase(rstDet("aFoto"))<>"" then s=s & " <img src='http://domoh.com/" & rstDet("aFoto") &"' />"
	DetallesFoto=s
End Function
%>
