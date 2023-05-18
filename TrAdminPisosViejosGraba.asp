<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim sPiso, numId, numVisitas, minFecha, sBody
	
	for each sPiso in Request.Form
		if UCase(Request.Form.Item(spiso))="ON" then
			numId=Mid(sPiso, 3, Len(sPiso))
			sQuery="SELECT * FROM Usuarios u INNER JOIN Anuncios a ON u.usuario=a.usuario WHERE a.id="& numId 
			rst.Open sQuery, sProvider

			if Left(sPiso,1)="m" then
				if Left(sPiso,2)="me" then
					sBody="Estimad@ " & rst("nombre") & ", Tu anuncio del piso (" & rst("cabecera") & ") está a punto de expirar."
					sBody=sBody & "Si deseas renovar el anuncio pulsa <a title='Renovar ahora' href=http://domoh.com/TrRenovarAnuncio.asp?id=" & numId & ">aquí</a>"
					sBody=sBody & ", en caso contrario lo desactivaremos el día " & FormatDateTime(Now+3,2) & "."
					sBody=sBody & "Siempre podrás reactivarlo entrando en <a title='Home de Domoh' href=http://domoh.com>domoh.com</a>."
					Mail rst("email"), "Renovación de Anuncio", sBody
				else
					sBody="Dear " & rst("nombre") & ", Your advert is about to expire (" & rst("cabecera") & "). "
					sBody=sBody & "If you want to renew it, please click <a title='Renew Advert' href=http://domoh.com/TrRenovarAnuncio.asp?idioma=En&id=" & numId & ">here</a>"
					sBody=sBody & ", otherwise we will hide it next " & FormatDateTime(Now+3,2) & ". You can always reactivate your advert entering <a title='Home Page' href=http://domoh.com>domoh.com</a>."
					Mail rst("email"), "Ad Renewal", sBody
				end if
				rst.Close
				sQuery="UPDATE Anuncios SET fechaavisobaja=GETDATE() WHERE id="& numId
				rst.Open sQuery, sProvider
			elseif left(sPiso,1)="b" then
				if left(sPiso,2)="be" then
					sBody="Estimad@ " & rst("nombre") & ", Tu anuncio del piso (" & rst("cabecera") & ") ha expirado."
					sBody=sBody & "Hemos procedido a desactivarlo. Puedes reactivarlo entrando en <a title='Home de Domoh' href=http://domoh.com>domoh.com</a>."
					Mail rst("email"), "Desactivación de Anuncio", sBody
				else
					sBody="Dear " & rst("nombre") & ", Your advert has expired. (" & rst("cabecera") & "). We have hidden it. You can reactivate your advert entering "
					sBody=sBody & "<a title='Domoh Home Page' href=http://domoh.com>domoh.com</a>."
					Mail rst("email"), "Advert Hidden", sBody
				end if
				rst.Close
				sQuery="UPDATE Anuncios SET activo='No', destacado=0, fechaultimamodificacion=GETDATE(), fechaavisobaja=NULL WHERE id="& numId
				rst.Open sQuery, sProvider
			elseif Left(sPiso,1)="d" then
				if Left(sPiso,2)="de" then
					sBody="Estimad@ " & rst("nombre") & ", Tu anuncio del piso (" & rst("cabecera") & ") ha dejado de estar destacado. Puedes renovarlo enviando DOMOH DES" & numId & " al 27722."
					Mail rst("email"), "Fin de Anuncio Destacado", sBody
				else
					sBody="Dear " & rst("nombre") & ", Your advert (" & rst("cabecera") & ") has finished to be in bold. You can renew it by sending DOMOH DES" & numId & " to 27722."
					Mail rst("email"), "End of Advert in Bold", sBody
				end if
				rst.Close
				sQuery="UPDATE Anuncios SET destacado=0, fechaultimamodificacion=GETDATE(), fechafindestacado=NULL WHERE id=" & numId
				rst.Open sQuery, sProvider
			elseif Left(sPiso,1)="v" then
				rst.Close
				sQuery="SELECT SUM(visitas) AS numVisitas, MIN(fecha) AS minFecha FROM AnunciosVistos WHERE anuncio="& numId
				rst.Open sQuery, sProvider
				if rst("numVisitas")=0 then numVisitas=2 else numVisitas=rst("numVisitas")
				if rst("minFecha")=0 then minFecha=now-31 else minFecha=rst("minFecha")
				rst.Close
				sQuery="SELECT u.email AS uEmail, u.usuario AS uUsuario, * FROM Usuarios u INNER JOIN Anuncios a ON u.usuario=a.usuario WHERE a.id="& numId
				rst.Open sQuery, sProvider

				if left(sPiso,2)="ve" then
					sBody="Estimad@ " & rst("nombre") & ", Te recordamos que tu anuncio de alojamiento para vacaciones (" & rst("cabecera") & ") sigue publicado en domoh.com."
					sBody=sBody & "Tu anuncio ha sido visitado por " & numVisitas & " personas desde " & minFecha & ". "
					sBody=sBody & "Si deseas añadir fotos, actualizar la información o desactivar el anuncio pulsa <a title='Home de Domoh' href=http://domoh.com>aquí</a>, "
					sBody=sBody & "y entra en tu perfil con tu usuario " & rst("uUsuario") & ". "
					sBody=sBody & "Si no recuerdas tu clave pulsa <a title='Página de Recordatorio' href='http://domoh.com/NuOlvido.asp?email=" & rst("uEmail") & "'>aquí</a>."
					Mail rst("uEmail"), "Recordatorio de Anuncio", sBody
				else
					sBody="Dear " & rst("nombre") & ", We just want to remind you that your advert for holidays (" & rst("cabecera") & ") is still posted on domoh.com."
					sBody=sBody & "Your advert has been visited " & numVisitas & " times since " & minFecha & ". "
					sBody=sBody & "If you want to upload pictures, update the information or hide the advert click <a title='Domoh Home Page' href=http://domoh.com>here</a>, "
					sBody=sBody & "and log in with your username " & rst("uUsuario") & ". "
					sBody=sBody & "If you cannot remember your password just click <a href='http://domoh.com/NuOlvido.asp?idioma=En&email=" & rst("uemail") & "'>here</a>."
					Mail rst("uEmail"), "Ad Reminder", sBody
				end if
				rst.Close
				rst.Open "UPDATE Anuncios SET fechaavisobaja=GETDATE() WHERE id="& numId, sProvider
			end if
		end if 
	next

	Response.Redirect "TrAdminPisosViejos.asp"
%>