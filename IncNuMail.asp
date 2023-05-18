<%
	dim esPrueba
	esPrueba=False

'---- Mail() Manda un Mail con to=[sAquien], Asunto=[sSubject] y Texto=[sBody]
Sub Mail(sAquien, sSubject, sBody)
	dim Mail, Conf
	on error goto 0
		
	'Session.CodePage = 65001
	set Mail=CreateObject("CDO.Message")
	set Conf=CreateObject("CDO.Configuration")
	Conf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com" 
	Conf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 
	Conf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 
	Conf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True 
	Conf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "domoh@domoh.com" 
	Conf.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "hecaran" 
	Conf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	Conf.Fields.Update
	set Mail.Configuration=Conf
	Mail.From = "domoh@domoh.com"
	'Mail.CharSet = "UTF-8"
	'Mail.ContentTransferEncoding = "Quoted-Printable"
	if not esPrueba then Mail.To=sAquien
	if not esPrueba then Mail.BCC="hector@domoh.com"
	Mail.BCC= Mail.BCC & ",diego@domoh.com"

	if esPrueba then 
		Mail.Subject = "Prueba: " & sSubject 
	else 
		Mail.Subject = sSubject
	end if

	if InStr(sSubject,"Anunciate")=0 then 
		Mail.HTMLBody = "<body>" 
        Mail.HTMLBody = Mail.HTMLBody & "<p>" & sBody & "</p>"
		Mail.HTMLBody = Mail.HTMLBody & "<style>.mail {font-size:12pt;font-family:Arial}</style>>"
		Mail.HTMLBody = Mail.HTMLBody & "<span class=mail>"
		if Session("Idioma")="En" then
			Mail.HTMLBody = Mail.HTMLBody & "<p>Best Regards,</p>"
            Mail.HTMLBody = Mail.HTMLBody & "<p>Domoh.com team</p>"
			Mail.HTMLBody = Mail.HTMLBody & "<p><a title='Go to Domoh' href='http://www.domoh.com/'>www.domoh.com</a></p>"
			Mail.HTMLBody = Mail.HTMLBody & "<p>Your Gay Friendly Classifieds website</p></span>"
		else
			Mail.HTMLBody = Mail.HTMLBody & "<p>Un saludo,</p>"
            Mail.HTMLBody = Mail.HTMLBody & "<p>El equipo de domoh.com</p>"
			Mail.HTMLBody = Mail.HTMLBody & "<p><a title='Ir a Domoh' href=http://www.domoh.com>www.domoh.com</a></p>"
			Mail.HTMLBody = Mail.HTMLBody & "<p>La web de anuncios clasificados Gay Friendly</p></span></body>"
		end if
	else
		Mail.HTMLBody = "<body>" & sBody & "</body>"
	end if	
	
	if esPrueba then 
        Response.Write "A: " & sAquien 
		Response.Write "BCC: hector@domoh.com" & Mail.BCC 
		Response.Write Mail.HTMLBody
	end if
	Err.Clear
	Mail.Send
End Sub

'---- Bienvenida() Manda un Mail de bienvenida y vuelve a la URL [sURL]
Sub Bienvenida(sURL)
	dim em, us, nom, pass, sBody
		
 	if Session("Bienvenida")="Si" then Exit Sub

	em=request("EMail")
	if em="" then em=Session("Email")

	us=request("Usuario")
	if us="" then us=Session("Usuario")

	nom=request("Nombre")
	if nom="" then nom=Session("Nombre")

	pass=request("Password")
	if pass="" then pass=Session("Password")

	if Session("Idioma")="" then
		sBody="Bienvenid@ a <a title='Ir a domoh' href=http://domoh.com>domoh.com</a>, " & nom & ". "
		sBody= sBody & "Lo primero de todo agradecerte el que te hayas registrado y que anuncies tu piso/habitación en nuestra página. Puedes acceder pulsando <a title='Ir a domoh' href=http://domoh.com>aquí</a> "
		sBody= sBody & "y rellenando: usuario: " & us & "clave: " & pass & "Tu anuncio será publicado en breve una vez hayamos comprobado que los datos que nos has facilitado son los correctos."
		sBody= sBody & "Aprovechamos esta ocasión para recordarte que puedes modificar y/o dar de baja tus datos y los del anuncio en cualquier momento. "
		sBody= sBody & "Para ello sólo tienes que entrar en <a href=http://domoh.com>domoh.com</a> e introducir tu nombre de usuario y tu clave. "
		sBody= sBody & "Asimismo te agradeceremos que una vez el anuncio no sea válido lo des de baja u ocultes (que seguirá en nuestros registros para que puedas reactivarlo si te es necesario). Muchas gracias"
	else
		sBody="Welcome to <a title='Enter Domoh' href=http://domoh.com>domoh.com</a>, " & nom & ". First of all, thank you for registering and posting your ad."
		sBody= sBody & "You can access domoh.com clicking <a title='Go to domoh' href=http://domoh.com>here</a> and filling in: user name: " & us & "password: " & pass 
		sBody= sBody & "Your advert will be published shortly once we validate the information."
		sBody= sBody & "You can update your personal data any time. Just access <a title='Go to domoh' href=http://domoh.com>domoh.com</a> "
		sBody= sBody & "and fill in your user name and password. Also, we would be thankful if you could deactivate the ad once it is no longer valid. We will keep it in our records, in case you need it again."
		sBody= sBody & "Thank you very much."
	end if
	
	if Session("Idioma")="" then
		Mail em, "Bienvenid@ a domoh.com", sBody
	 	if Err then Response.Redirect sURL & "?msg=" & Server.URLEncode("Mail Incorrecto.Inténtalo de nuevo.")
	else
		Mail em, "Welcome to domoh.com", sBody
	 	if Err then Response.redirect sURL & "?msg=Incorrect+e-mail.Please+try+again."
	end if
 	Session("Bienvenida")="Si"
End Sub

'---- PreBienvenida() Manda un Mail de Confirmación y vuelve a la URL [sURL]
Sub PreBienvenida(sURL)
	dim sBody, sWeb, sSubject
	on error goto 0
				
	if esPrueba then 
		sWeb="abyss"
	else
		sWeb="domoh.com"
	end if

	rst.Open "SELECT Id FROM Usuarios WHERE usuario='" & Session("Usuario") & "'", sProvider
	if Session("Idioma")="" then
		sBody="Por favor pulsa <a title='Confirmar Registro' href=http://" & sWeb & "/NuConfRegistroUsuario.asp?usuario=" 
		sBody=sBody & Server.URLEncode(Session("Usuario")) & "&id=" & rst("id") & ">aquí</a> para completar tu registro en domoh.com, " & Session("Nombre") & "."
	else
		sBody="Please click <a title='Confirm Registration' href=http://" & sWeb & "/NuConfRegistroUsuario.asp?idioma=En&usuario=" 
		sBody=sBody & Server.URLEncode(Session("Usuario")) & "&id=" & rst("id") & ">here</a> in order to complete your registration in domoh.com, " & Session("Nombre") & ".<br/>"
	end if
	rst.Close

	if sURL="TrCasaRegDemanda.asp" then
		sSubject="Confirmación de Registro para Buscar Alojamiento en Alquiler en domoh.com"
		Mail Session("Email"), sSubject, sBody
	elseif sURL="TrCasaRegDemandaEn.asp" then
		Mail Session("Email"), "Ad Confirmation for Looking for House in domoh.com", sBody
	elseif sURL="TrCasaRegOferta.asp" then
		sSubject="Confirmación de Registro para Ofrecer Alojamiento en Alquiler en domoh.com"
		Mail Session("Email"), sSubject, sBody
	elseif sURL="TrCasaRegOfertaEn.asp" then
		Mail Session("Email"), "Ad Confirmation for Posting House for Rent in domoh.com", sBody
	elseif sURL="TrVacasRegOferta.asp" then
		sSubject="Confirmación de Registro para Ofrecer Alojamiento en Vacaciones en domoh.com"
		Mail Session("Email"), sSubject, sBody
	elseif sURL="TrVacasRegOfertaEn.asp" then
		Mail Session("Email"), "Ad Confirmation for Posting Lodging for Holidays in domoh.com", sBody
	elseif sURL="NuCurroRegOferta.asp" then
		Mail Session("Email"), "Confirmación de Registro para Publicar Oferta de Trabajo en domoh.com", sBody
	elseif sURL="NuCurroRegOfertaEn.asp" then
		Mail Session("Email"), "Ad Confirmation for Posting Job Offer in domoh.com", sBody
	elseif sURL="NuMiscRegOferta.asp" then
		Mail Session("Email"), "Confirmación de Registro para Publicar Anuncio en domoh.com", sBody
	elseif sURL="NuMiscRegOfertaEn.asp" then
		Mail Session("Email"), "Confimation for Posting Advert in domoh.com", sBody
	else
		Mail Session("Email"), "Confirmación de Registro para Ofrecer Alojamiento en domoh.com", sBody
	end if
	
 	if Err then Response.Redirect sURL & "?msg=Mail+Incorrecto.Inténtalo+de+nuevo."
End Sub
%>
