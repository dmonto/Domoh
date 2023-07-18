<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim sInqui, numId, sTitle, sBody

	for each sInqui in Request.Form
		if UCase(Request.Form.Item(sInqui))="ON" then
			numId=mid(sInqui, 3, len(sInqui))
			sQuery="SELECT * FROM Usuarios u INNER JOIN Anuncios a ON u.usuario=a.usuario WHERE a.id="& numId 
			rst.Open sQuery, sProvider

			if left(sInqui,1)="m" then
				if left(sInqui,2)="me" or left(sInqui,2)="mb" then
					sTitle="Renovación de Anuncio"
					sBody="Estimad@ " & rst("nombre") & ", Tu anuncio de búsqueda de piso  (" & rst("cabecera") & ") está a punto de expirar. Si deseas renovar el anuncio pulsa "
					sBody=sBody & " <a title='Renovar Anuncio' href=http://domoh.com/TrRenovarAnuncio.asp?id=" & numId & ">aquí</a>, en caso contrario lo desactivaremos el día " & FormatDateTime(Now+3,2) & "."
					sBody=sBody & " Siempre podrás reactivarlo entrando en <a title='Ir a Domoh' href=http://domoh.com>domoh.com</a>."
				end if
				if left(sInqui,2)="mb" then sTitle=sTitle & " / "
				if left(sInqui,2)="mi" or left(sInqui,2)="mb" then				
					sTitle=sTitle & "Ad Renewal"
                    sBody=sBody & "Dear " & rst("nombre") & ", Your advert is about to expire "
					sBody=sBody & "(" & rst("cabecera") & "). If you want to renew it, please click <a title='Renew Advert' href=http://domoh.com/TrRenovarAnuncio.asp?idioma=En&id=" & numId & ">here</a>, "
					sBody=sBody & "otherwise we will hide it next " & FormatDateTime(Now+3,2) & " . You can always reactivate your advert entering <a title='Go to Domoh' href=http://domoh.com>domoh.com</a>."
				end if
				Mail rst("email"), sTitle, sBody
				rst.Close
				sQuery="UPDATE Anuncios SET fechaavisobaja=GETDATE() WHERE id="& numId
				rst.Open sQuery, sProvider
			elseif left(sInqui,1)="b" then
				if left(sInqui,2)="be" or left(sInqui,2)="bb" then
					sTitle="Desactivación de Anuncio"
					sBody="Estimad@ " & rst("nombre") & ", Tu anuncio de búsqueda de piso (" & rst("cabecera") & ") ha expirado. Hemos procedido a desactivarlo."
					sBody=sBody & "Puedes reactivarlo entrando en " & "<a title='Ir a Domoh' href=http://domoh.com>domoh.com</a>."
				end if
				if left(sInqui,2)="bb" then sTitle=sTitle & " / "
				if left(sInqui,2)="bi" or left(sInqui,2)="bb" then				
					sTitle=sTitle & "Advert Hidden"
                    sBody=sBody & "Dear " & rst("nombre") & ", Your advert has expired. (" & rst("cabecera") & "). We have hidden it."
					sBody=sBody & "You can reactivate your advert entering <a title='Go to Domoh' href=http://domoh.com>domoh.com</a>."
				end if
				Mail rst("email"), sTitle, sBody
				rst.Close
				sQuery="UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE(), fechaavisobaja=NULL WHERE id="& numId
				rst.Open sQuery, sProvider
			end if
		end if 
	next

	Response.Redirect "NuAdminInquilinosViejos.asp"
%>
