<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim sPiso, numId, sUsuario, sTitle, sBody

	for each sPiso in Request.Form
		if UCase(Request.Form.Item(sPiso))="ON" then
			numId=Mid(sPiso, 3, Len(sPiso))
			sQuery="SELECT * FROM Usuarios u INNER JOIN Anuncios a ON u.usuario=a.usuario WHERE a.id="& numId 
			rst.Open sQuery, sProvider
			sUsuario=rst("usuario")
			sTitle=""
			sBody=""			

			if left(sPiso,1)="m" then
				if left(sPiso,2)="me" or left(sPiso,2)="mb" then
					sTitle="Renovación de Anuncio"
					sBody="Estimad@ " & rst("nombre") & ", Tu anuncio del piso (" & rst("cabecera") & ") lleva un año inactivo en nuestra Base de Datos."
					sBody=sBody & " Si deseas mantener la información para utilizarla en el futuro pulsa <a title='Renovar Anuncio' href=http://domoh.com/TrRenovarAnuncio.asp?id=" & numId & ">aquí</a>, "
					sBody=sBody & " en caso contrario lo borraremos el día " & FormatDateTime(Now+7,2) & "."
					sBody=sBody & " Siempre podrás introducirlo de nuevo entrando en <a title='Ir a Domoh' href=http://domoh.com>domoh.com</a>."
				end if
				if left(sPiso,2)="mb" then sTitle=sTitle & " / "
				if left(sPiso,2)="mi" or left(sPiso,2)="mb" then
					sTitle=sTitle & "Ad Renewal"
					sBody=sBody & "Dear " & rst("nombre") & ", Your advert (" & rst("cabecera") & ") has been inactive for one year."
					sBody=sBody & "If your want to keep its contents in our records for future use, please click "
					sBody=sBody & "<a title='Renew Advert' href=http://domoh.com/TrRenovarAnuncio.asp?idioma=En&id=" & numId & ">here</a>"
					sBody=sBody & ", otherwise we will delete it next " & FormatDateTime(Now+7,2) & ". You can advert your lodging again entering <a title='Go to Domoh' href='http://domoh.com'>domoh.com</a>."
				end if
				Mail rst("email"), sTitle, sBody
				rst.Close
				sQuery="UPDATE Anuncios SET fechaavisobaja=GETDATE() WHERE id="& numId
				rst.Open sQuery, sProvider
			elseif left(sPiso,1)="b" then
				if left(sPiso,2)="be" or left(sPiso,2)="bb" then
					sTitle="Borrado de Anuncio"
					sBody="Estimad@ " & rst("nombre") & ", Tu anuncio del piso (" & rst("cabecera") & ") ha expirado. Hemos procedido a borrarlo."
				    sBody=sBody & "Puedes volver a introducirlo entrando en <a title='Ir a Domoh' href='http://domoh.com'>domoh.com</a>."
				end if
				if left(sPiso,2)="bb" then sTitle=sTitle & " / "
				if left(sPiso,2)="bi" or left(sPiso,2)="bb" then
					sTitle=sTitle & "Ad Deletion"
					sBody=sBody & "Dear " & rst("nombre") & ", Your advert (" & rst("cabecera") & ") has expired. We have deleted it from our records."
					sBody=sBody & "You can re-post it entering <a title='Go to Domoh' href='http://domoh.com'>domoh.com</a>."
				end if
				Mail rst("email"), sTitle, sBody
				rst.Close
				rst.Open "DELETE FROM Fotos WHERE piso="& numId, sProvider
				rst.Open "DELETE FROM Pisos WHERE id="& numId, sProvider
				rst.Open "DELETE FROM Anuncios WHERE id="& numId, sProvider
			end if
		end if 
	next

	Response.Redirect "NuAdminPisosInactivos.asp"
%>