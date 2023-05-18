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

'---- CasaDetalles devuelve el HTML de una casa [vId] en [vIdioma] con [vFotos]. [vModo] Obsoleto. [vSeq] es el número de id del DOM
Function CasaDetalles(vIdioma, vId, vModo, vFotos, vSeq)
	dim s, sQuery, sQueryDet
				
	set rstDet = Server.CreateObject("ADODB.recordset")
	sQuery="SELECT p.tipo AS pTipo, p.fuma AS pFuma, p.mascota AS pMascota, p.dir1 AS pDir1, p.dir2 AS pDir2, p.cp AS pCp, a.id AS aId, p.ciudad AS pCiudad, a.descripcion AS aDescripcion, a.foto AS aFoto, * "
	sQuery=sQuery & " FROM (Pisos p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.usuario WHERE a.id=" & vId
	rstDet.Open sQuery, sProvider
	
	s=s & "<div class=manfoto>" & DetallesFoto() & "</div>"
	s=s & "<div class=mantitulo id='cabtr" & vSeq & "'>" & rstDet("ciudadnombre") & rstDet("zona") & "</div>"
	s=s & "<div class=mantipo title='Qué es'>"

	if rstDet("pTipo")="Piso" then
		s=s & MesgS("Piso","Flat")
	elseif rstDet("pTipo")="Habitación" then
		s=s & MesgS("Habitación","Room")
	elseif rst("pTipo")="Casa" then
		s=s & MesgS("Casa","House")
	else
		s=s & rst("pTipo")
	end if
	
	if rstDet("pTipo")="Hotel" then s=s & "<img alt='Es un Hotel' src='http://domoh.com/images/Hotel.gif'/>"
	s=s & "</div>"
	s=s & "<div class=manprecio title='Precio'>"

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
	s=s & "<div class=manicon><a title='Mostrar Entero' href=# id='bloquea" & vSeq & "'><img src=http://domoh.com/images/EyeAd.jpg /></a><img src=http://domoh.com/images/spacer.gif />"
	s=s & "<a title='No Me Interesa' href=# id='quita" & vSeq & "'><img src=http://domoh.com/images/DeleteAd.jpg /></a></div>"
	s=s & "<div class=mancontacto id='tpg" & vSeq & ".4'>" & DetallesContacto(vIdioma) & "</div>"
	if rstDet("precio")=0 then
		s=s & "<div class=manperro title='Mascotas'>"
		if UCase(rstDet("pMascota"))<>"ON" then 
            s=s & "<img alt='No Perros' src=http://domoh.com/images/NoPerrillos.gif />"
        else
			s=s & "<img src=http://domoh.com/images/Perrillos.gif />"
		end if
		s=s & "</div>"
		s=s & "<div class=mancigarro title='Fumadores'>"
		if UCase(rstDet("pFuma"))<>"ON" then s=s & "<img src=http://domoh.com/images/NoCigarrito.gif />" else s=s & "<img src=http://domoh.com/images/Cigarrito.gif />"
		s=s & "</div>"
		s=s & "<div class=mandetalle>" & rstDet("aDescripcion") & "</div>"
  
        s=s & "<div class=manadicional>" & rstDet("gente") 
	    sQueryDet="SELECT * FROM AnuncioComent WHERE anuncio=" & rstDet("aId")
        rstDet.Close  
	    rstDet.Open sQueryDet, sProvider
	    if not rstDet.Eof then
            s=s & MesgS("Comentarios de gente","What other people think") 
		    do while not rstDet.Eof 
                s=s & rstDet("valoracion") & " - " & rstDet("texto") & ". " & MesgS("Por","By") & rstDet("usuario") & "<p>" & MesgS("el ","on ") & rstDet("fechaalta") 
			    rstDet.Movenext
		    loop
	    end if	
	    rstDet.Close
		if Request("coment")<>"si" then
            s=s & "<form name=frm action=TrAnuncioComent.asp method=post><input type=hidden value=" & Request("id") & " name=id>Already called? Please give us feedback: <textarea rows=3 cols=60 name=comentario></textarea></form>"
	    end if		
        s=s & "</div>"
  
   	    CasaDetalles=s
    end if
End Function


'---- DetallesContacto devuelve el HTML del contacto de rstDet en [vIdioma]
Function DetallesContacto(vIdioma)
	dim s
	
	if rstDet("esagencia")<>"" then s=s & Mesg(vIdioma,"Agencia","Agency") & ";" else s=s & Mesg(vIdioma,"Contacto","Contact") & ";" 
    s=s & rstDet("nombre") & rstDet("dir1") & " " & rstDet("dir2") 
	if rstDet("dir1")<>"" or rstDet("dir2")<>"" then s=s & " - " 
	s=s & rstDet("cp") & " " & rstDet("ciudaddir") 
	if (UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel1")<>0) or (UCase(rstDet("mostrartel2")) = "ON" and rstDet("tel2")<>0) or (UCase(rstDet("mostrartel3")) = "ON" and rstDet("tel3")<>0) _
        or (UCase(rstDet("mostrartel4")) = "ON" and rstDet("tel4")<>0) then 
		s=s & Mesg(vIdioma,"Telef","Phone") & ";<img src=http://domoh.com/images/telefonillo.gif />"
	end if 
	if UCase(rstDet("mostrartel1")) = "ON" and rstDet("tel1")<>0 then s=s & rstDet("tel1") & " "
	if UCase(rstDet("mostrartel2")) = "ON" and rstDet("tel2")<>0 then s=s & rstDet("tel2") & " "
	if UCase(rstDet("mostrartel3")) = "ON" and rstDet("tel3")<>0 then s=s & rstDet("tel3") & " "
	if UCase(rstDet("mostrartel4")) = "ON" and rstDet("tel4")<>0 then s=s & rstDet("tel4") & " "
	if rstDet("instrucciones") <>"" then s=s & "(" & rstDet("instrucciones") & ")"
	if UCase(rstDet("mostraremail")) = "ON" then s=s & " eMail;<a href=mailto:" & rstDet("email") & ">" & rstDet("email") & "</a>"
	
	DetallesContacto=s
End Function

'---- DetallesFoto devuelve el HTML de la foto en rstDet
Function DetallesFoto()
	dim s
	
	if UCase(rstDet("aFoto"))<>"" then s=s & "<img src='http://domoh.com/mini" & rstDet("aFoto") & "' />"
	DetallesFoto=s
End Function
%>
