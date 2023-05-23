<!-- #include file="IncNuBD.asp" -->
<%
	dim sAnuncio, sTablaCampo
	on error goto 0
	
	sTablaCampo=Request("tablacampo")
	if sTablaCampo="" then sTablaCampo="Anuncios"
	
	for each sAnuncio in Request.Form
		if Request.Form.Item(sAnuncio)="Borrar" then
			sQuery="UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE() WHERE id=" & sAnuncio
			rst.Open sQuery, sProvider
		elseif Request.Form.Item(sAnuncio)="Reactivar" then
			sQuery="UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & sAnuncio
			rst.Open sQuery, sProvider
		elseif Request.Form.Item(sAnuncio)="Eliminar" then
			rst.Open "DELETE FROM " & Request("tabla") & " WHERE id=" & sAnuncio, sProvider
			rst.Open "DELETE FROM Fotos WHERE piso=" & sAnuncio, sProvider
			rst.Open "DELETE FROM Anuncios WHERE id=" & sAnuncio, sProvider
		elseif Request.Form.Item(sAnuncio)<>"0" and sAnuncio<>"tipo" and sAnuncio<>"tabla" and sAnuncio<>"campo" and sAnuncio<>"tablacampo" then
			sQuery="UPDATE " & sTablaCampo & " SET " & Request("campo") & "=" & Request.Form.Item(sAnuncio) & " WHERE id=" & sAnuncio
			rst.Open sQuery, sProvider
			sQuery="UPDATE Anuncios SET fechaultimamodificacion=GETDATE() WHERE id=" & sAnuncio
			rst.Open sQuery, sProvider
		end if 
	next

	Response.Redirect "NuAdmin" & Request("tabla") & ".asp?tipo=" & Request("tipo")
%>
