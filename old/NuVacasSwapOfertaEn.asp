<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrDetalles.asp" -->
<%
	dim sBody
		
	if Request("tipo")="" then Response.Redirect "NuVacasSwapOfrezcoFront.asp"

	if Session("Usuario")="" then
		'--- Usuario Nuevo 
    	if Request("password")="" then
            Mail "diego@domoh.com", "Password vacia en " & Request.Servervariables("SCRIPT_NAME"), Request.Form("usuario")
            Response.Write "Algo fue mal"
            Response.End        
        end if

		Session("TipoUsuario")="Casero"
		Session("Activo")="Mail"
		InsertUsuario Request
		PreBienvenida Request.Servervariables("SCRIPT_NAME")
	else
		sQuery="UPDATE Usuarios SET numanuncios=numanuncios+1 WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
	end if	

	if Request("id")="" or Request("id")="nuevo" then
		'--- Piso Nuevo 
	 	sQuery = "INSERT INTO Anuncios (tabla, activo, usuario, cabecera, descripcion, provincia, idiomaes, idiomaen, fechaalta, fechaultimamodificacion) "
		sQuery = sQuery & "VALUES ('Pisos','Si','" & Session("Usuario") & "','" & Replace(Request("ciudadnombre"),"'","''") & " (" & Replace(Request("zona"),"'","''") & ")',"
		sQuery = sQuery & "'" & Replace(Request("descripcion"),"'","''") & "',0,'" & Request("idiomaes") & "', '" & Request("idiomaen") & "', GETDATE(), GETDATE())"

		Err.Clear
		rst.Open sQuery, sProvider
		if Err then 
			Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
			Response.Redirect "TrVacasRegOfrezcoFront.asp"
		end if
				
		sQuery="SELECT MAX(id) AS maxId FROM Anuncios WHERE usuario='" & Session("Usuario") & "'"
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuVacasSwapOferta", sQuery & " - " & Err.Description
		Session("Id")=rst("maxid")
		rst.Close

	 	sQuery = "INSERT INTO Pisos (id, activo, tipo, dir1, ciudadnombre, dir2, cp, zona, rentaviv, usuario, descripcion, gente, fuma, mascota, idiomaes, idiomaen, fechaalta, destacado, fechaultimamodificacion) "
		sQuery = sQuery & "VALUES (" & Session("Id") & ",'Si','Piso','" & Replace(Request("dir1"),"'","''") & "','" & Replace(Request("ciudadnombre"),"'","''") & "','" & Replace(Request("dir2"),"'","''") & "',"
		sQuery = sQuery & "'" & Request("cp") & "','" & Replace(Request("zona"),"'","''") & "'," & LimpiaNum(Request("rentaviv")) & ",'" & Session("Usuario") & "','" & Replace(Request("descripcion"),"'","''") & "', "
		sQuery = sQuery & "'" & Replace(Request("gente"),"'","''") & "', '" & Request("fuma") & "','" & Request("mascota") & "', '" & Request("IdiomaEs") & "','" & Request("IdiomaEn") & "', GETDATE(), 0, GETDATE())"
	
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description

		Session("Ciudad")=Request("ciudad")
		sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>"
		sBody=sBody & "Piso para Intercambio introducido por " & Session("Usuario") & " en " & Request("ciudadnombre") & CasaDetalles("", Session("Id"), "Admin", "ON")
		Mail "hector@domoh.com", "Nuevo Piso para Intercambio en " & Request("ciudadnombre"), sBody
		if Request("id")="nuevo" and Request("Foto")<>"on" then Response.Redirect "QuHomeUsuario.asp?msg=Apartment+Added."
	else
		'--- Modificar Piso
	 	sQuery = "UPDATE Anuncios SET cabecera='" & Replace(Request("ciudadnombre"),"'","''") & " (" & Replace(Request("zona"),"'","''") & ")',descripcion='" & Replace(Request("descripcion"),"'","''") & "',"
		sQuery = sQuery & "idiomaes='" & Request("idiomaes") & "', idiomaen='" & Request("idiomaen") & "', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id") 

		Err.Clear
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en NuVacasSwapOferta", sQuery & " - " & Err.Description

	 	sQuery = "UPDATE Pisos SET dir1='" & Replace(Request("dir1"),"'","''") & "', ciudadnombre='" & Replace(Request("ciudadnombre"),"'","''") & "',"
		sQuery = sQuery & " dir2='" & Replace(Request("dir2"),"'","''") & "', cp='" & Request("cp") & "', zona='" & Replace(Request("zona"),"'","''") & "', fuma='" & Request("fuma")  & "',"
		sQuery = sQuery & " mascota='" & Request("mascota")  & "' WHERE id=" & Request("id") 
		rst.Open sQuery, sProvider
		if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), sQuery & " - " & Err.Description
		sBody="<head><link href=http://domoh.com/TrDomoh.css rel=stylesheet type=text/css></head>" & Request("tipo") & " modificado por " & Session("Usuario") & CasaDetalles("", Session("Id"), "Admin", "ON")
		Mail "hector@domoh.com", Request("tipo") & " para Intercambio Modificado", sBody
		Session("Id")=Request("id")
	end if
%>
<head>
    <!-- #include file="IncTrCabecera.asp" -->
    <title>Holidays - Upload Pictures for Swapping</title>
</head>
<body onload="window.parent.location.hash='top';
<%	if instr(Session("EMail"),"hotmail")>0 then %>
	detalles('images/AntiSpam.gif');
<%	end if %>
"> 
<form name=frm action=NuVacasSwapOfertaFotoEn.asp method=post enctype=multipart/form-data>
	<input type=hidden name=id value=<%=Session("Id")%> />
<div class=container>
	<div>
        <!--#include file="IncNuSubMenu.asp"-->
<%	if Request("foto")="on" or Request("op")="FotoBorrada" or Request("op")="Test" then %>
    	</div>
	<div>
		<p>First image will be used as a preview</p>
<%
 		if Session("Id")<>"" then 
			if Request("foto")="on" then
%>
		<a title='Go to your profile' href='QuHomeUsuario.asp?msg=House+Added.'>I will not upload pictures at this moment.</a>
<% 			else %>
		<a title='Go to your profile' href='QuHomeUsuario.asp?msg=House+Updated.'>I will not upload pictures at this moment.</a>
<%
			end if
			
			rst.Open "SELECT * FROM Fotos WHERE piso=" & Session("Id"), sProvider
			if Err then Mail "diego@domoh.com", "Error en " & Request.Servervariables("SCRIPT_NAME"), Err.Description
			while not rst.eof 
%>
			<img alt='Foto' title='To remove this picture click on the right' src='http://domoh.com/<%=rst("foto")%>'/><a title='Delete this Picture' href='TrVacasRegOfertaFoto.asp?borrar=<%=rst("id")%>'>Remove</a>
<% 
				rst.Movenext
			wend 
			rst.Close
		else
%>
			<a title='Return to Home Page' href='old/NuDomohEn.asp?msg=We+have+sent+you+an+e-mail+containing+a+link+which+you+have+to+click+in+order+to+confirm+the+publication.'>I will not upload pictures at this moment.</a>
<% 		end if %>
			<p>Digital images files (maximum 300Kb each)</p>
			<p><input title='Press on the right to choose the first picture' type=file name=foto /></p>
			<p><input title='Press on the right to choose the second picture' type=file name=foto /></p>
			<p><input title='Press on the right to choose the third picture' type=file name=foto /></p>
			<p><input title='Press on the right to choose the fourth picture' type=file name=foto /></p>
			<p><input title='Press on the right to choose the fifth picture' type=file name=foto /></p>
			<p><input title="Press on the right to choose the sixth picture" type=file name=foto /></p></div></div>
	<div><input title=Save type=submit value='Upload Pictures' class=btnForm /></div>
<%	
	elseif Request("id")<>"" then 
		Response.Redirect "QuHomeUsuario.asp?msg=House+Updated."		
	else 
%>
    <!-- #include file="IncTrEnvioMail.asp" -->
<% 	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>