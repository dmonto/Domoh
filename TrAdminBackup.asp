<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncGrabaExcel.asp" -->
<% 
	dim s
	on error goto 0
	
	if Request("fase")="" then
%>
<html>
    <!-- #include file="IncTrCabecera.asp" -->
    <head>
        <style>div, h1 {display: table-row;} </style>
        <title>Domoh - Backup</title>
    </head>
<body>
    <div class=container>
		<div class=main>
			<h1><a href='TrAdminBackup.asp?fase=1'>Stage 1</a></h1>
			<h1><a href='TrAdminBackup.asp?fase=2'>Stage 2</a></h1>
			<h1><a href='TrAdminBackup.asp?fase=3'>Stage 3</a></h1>
			<h1><a href='TrAdminBackup.asp?fase=4'>Stage 4</a></h1>
<% 	elseif Request("fase")="1" then 
    	Response.Buffer=False
		s=Graba("Anuncios", "Anuncios") 
        Mail "", "Backup I", s
		Response.Write "<h1>Enviado I</h1>"
    elseif Request("fase")="2" then
		Response.Buffer=False
        s= Graba("Inquilinos", "Inquilinos") & Graba("Provincias", "Provincias") & Graba("Pisos", "Pisos") & Graba("Paises", "Paises") & Graba("Contenidos", "Contenidos") & Graba("Fuentes", "Fuentes") & Graba("AnuncioComent", "AnuncioComent") & Graba("Fotos", "Fotos") 
        Mail "", "Backup II", s
		Response.Write "Enviado II"
	elseif Request("fase")="3" then
		s= Graba("Usuarios", "Usuarios") 
		Mail "", "Backup III", s
		Response.Write "Enviado III"
	elseif Request("fase")="4" then
		s="<div>" & Graba("AnunciosVistos", "AnunciosVistos") & "</div>"
		Mail "", "Backup IV", s
		Response.Write "<h1>Enviado IV</h1>"
	end if
%>
		</div></div></body></html>