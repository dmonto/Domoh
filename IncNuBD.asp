<%
	option explicit
	dim rst, sProvider, sProvider2, sProviderOld, sQuery, sOpcion, fso
	on error resume next

	sProvider="PROVIDER=SQLOLEDB;DATA SOURCE=94.130.53.116;UID=domoh;PWD=lx_8L3p2;DATABASE=domoh"
	sProvider2=sProvider
    set rst = Server.CreateObject("ADODB.recordset")

'---- MesgS devuelve [vEs] si la sesión está en español y [vEn] si está en inglés
Function MesgS(vEs, vEn)
	MesgS=UMesgS(Request("idioma"), vEs, vEn)
End Function

'---- UMesgS devuelve [vEs] si vIdi es "Es" y [vEn] si vIdi es "En"
Function UMesgS(vIdi, vEs, vEn)
	if Session("HayIdioma")="" then
		if vIdi="Es" then 
			Session("Idioma")=""
			Session("HayIdioma")="Si"
			Response.Cookies("Idioma")="Es"
			Response.Cookies("Idioma").Expires=Now+365
		elseif vIdi="En" then 
			Session("Idioma")="En"
			Session("HayIdioma")="Si"
			Response.Cookies("Idioma")="En"
			Response.Cookies("Idioma").Expires=Now+365
		end if
	end if
	
	if Session("Idioma")="" then 
		UMesgS=vEs
	else
		UMesgS=vEn
	end if
End Function

'---- LimpiaNum devuelve [vTel] sin espacios ni simbolos que no sean numeros
Function LimpiaNum(vTel) 
	dim i, c, ret
	on error resume next
		
	for i=1 to Len(vTel)
		c=Mid(vTel,i,1)
		if c>="0" and c<="9" then ret=ret & c
	next

	if IsEmpty(ret) then ret="0"
	LimpiaNum=ret
End Function

'---- Format devuelve una [Fecha] en formato SQL
Function Format(vFecha) 
	on error resume next

	Format=Year(vFecha) & "/" & Month(vFecha) & "/" & Day(vFecha) & " " & Hour(vFecha) & ":" & Minute(vFecha)
End Function
	
'---- SelectProvincia muestra un combo con la provincias de [vTabla] y selecciona [vProvincia]
Sub SelectProvincia(vTabla, vProvincia)
%>
	<select title='<%=MesgS("Provincias disponibles","Available Regions")%>' name=provincia> 
		<option value=0 <% if vprovincia=0 then Response.Write "selected"%>><%=MesgS("-Cualquiera-","-Anyone-")%></option>
<%	
	sQuery = "SELECT * FROM Provincias"
	if vTabla<>"" then
		if vTabla="Vacas" then 
			sQuery = sQuery & " WHERE id IN (SELECT a.provincia FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.activo='Si' AND rentavacas<>0) "
		elseif vTabla="Vivir" then 
			sQuery = sQuery & " WHERE id IN (SELECT a.provincia FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.activo='Si' AND rentaviv<>0) "
		elseif vTabla="Compra" then 
			sQuery = sQuery & " WHERE id IN (SELECT a.provincia FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.activo='Si' AND precio<>0) "
		else
			sQuery = sQuery & " WHERE id IN (SELECT provincia FROM " & vTabla & " WHERE activo='Si') "
		end if
	end if

	if Session("Pais")<>0 then
		if vTabla<>"" then
			sQuery = sQuery & " AND pais=" & Session("Pais")
		else
			sQuery = sQuery & " WHERE pais=" & Session("Pais")
		end if
	end if

	sQuery = sQuery & " ORDER BY nombre"
	rst.Open sQuery, sProvider
	while not rst.Eof
%>
		<option value=<%=rst("id")%> <% if vprovincia=rst("id") then Response.Write " selected"%>><%=rst("nombre")%>
<%
		rst.Movenext
	wend
	rst.Close
%>									
		</option></select>
<%
End Sub

'---- SelectPais muestra un combo con la provincias de [vTabla] y selecciona [vPais]. En [vIdioma]
Sub SelectPais(vIdioma, vTabla, vPais)
	on error goto 0
	dim sFiltro
%>
	<select title='<%=MesgS("Países disponibles","Available Countries")%>' name=pais id=pais onchange='ClickPais();'> 
		<option value=0 <% if vPais=0 then Response.Write "selected"%>>
		<%=MesgS("-Cualquiera-","-Anyone-")%></option>
<%		
	sQuery = "SELECT * FROM Paises "
	if vTabla<>"" then
		if vTabla="Vacas" then 
			sFiltro=" (SELECT a.provincia FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.activo='Si' AND rentavacas<>0) "
		elseif vTabla="Vivir" then 
			sFiltro=" (SELECT a.provincia from Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.activo='Si' AND rentaviv<>0) "
		elseif vTabla="Compra" then 
			sFiltro=" (SELECT a.provincia from Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.activo='Si' AND precio<>0) "
		else
			sFiltro=" (SELECT provincia FROM " & vTabla & " WHERE activo='Si') "
		end if
		sQuery = sQuery & " WHERE id IN (SELECT pais FROM Provincias WHERE id IN " & sFiltro & " ) "
	end if
	sQuery = sQuery & " ORDER BY nombre"

	rst.Open sQuery, sProvider
	while not rst.Eof
%>
		<option value=<%=rst("id")%> <% if vpais=rst("id") then Response.Write " selected"%>><%=rst("Nombre" & vIdioma)%>
<%
		rst.Movenext
	wend
	rst.Close
%>									
		</option></select>
<%
End Sub

'---- SelectTabla muestra un combo con la provincias de [vTabla] y selecciona [vValor]. [vPadre] es la tabla en la que se buscan los valores posibles, y se muestra [vCampo]
Sub SelectTabla(vTabla, vPadre, vValor, vCampo)
%>
	<select title='<%=vTabla%> <%=MesgS("disponibles","available")%>' name=<%=vTabla%>> 
		<option value=0 <% if vValor=0 then Response.Write "selected"%>>
		<%=MesgS("-Cualquiera-","-Anyone-")%></option>
<%		
	if vPadre<>"" then
		sQuery="SELECT * FROM " & vTabla & " WHERE id IN (SELECT " & vTabla & " FROM " & vPadre & ") ORDER BY " & vCampo
		rst.Open sQuery, sProvider
	else
		rst.Open "SELECT * FROM " & vTabla & " ORDER BY " & vCampo, sProvider
	end if

	while not rst.Eof
%>
		<option value=<%=rst("id")%> 
			<% if vValor=rst("id") then Response.Write " selected"%>><%=rst(vCampo)%>
<%
		rst.Movenext
	wend
	rst.Close
%>									
		</option></select>
<% 
End Sub 

'---- SelectFuente muestra un combo con la provincias de Fuentes y selecciona [vValor]. 
Sub SelectFuente(vValor)
%>
	<select title='<%=MesgS("Elige una referencia","Choose a Reference")%>' name=fuente> 
		<option value=0 <% if vValor=0 then Response.Write "selected"%>/><%=MesgS("-Elige una-","-Select One-")%>
<%		
	rst.Open "SELECT * FROM Fuentes ORDER BY ""desc" & Session("Idioma") & """", sProvider

	while not rst.Eof
		if rst("desc" & Session("Idioma")) <> "" then
%>
		<option value=<%=rst("id")%> <% if vValor=rst("id") then Response.Write " selected"%>><%=rst("desc" & Session("Idioma"))%>
<%
		end if
		rst.Movenext
	wend
	rst.Close
%>									
		</option></select>
<% 
End Sub 

'---- InsertUsuario crea el usuario del formulario Req. Si ya existe le añade sufijo _Nuevo
Sub InsertUsuario(Req)
	sQuery = "SELECT * FROM Usuarios WHERE usuario='" & Req.Form("usuario") & "'"
	rst.Open sQuery, sProvider
	if not(rst.Eof) then 
		Session("Usuario")="_Nuevo" & Req.Form("usuario")
	else
		Session("Usuario")=Req.Form("usuario")
	end if
	rst.Close
	
	if Req.Form("password")="" then
		Mail "diego@domoh.com", "Password vacia en InsertUsuario", " Usuario " & Session("Usuario")
		Response.End
	end if

	sQuery = "INSERT INTO Usuarios (tipo, esagencia, activo, numanuncios, nombre, dir1, ciudaddir, dir2, cp, fuente, tel1, tel2, tel3, tel4, email, usuario, "
	sQuery = sQuery & " password, mostrartel1, mostrartel2, mostrartel3, instrucciones, mostrartel4, mostraremail, fechaalta, ultimavisita) "
	sQuery = sQuery & "VALUES ('"& Session("TipoUsuario") &"','"& Req.Form("EsAgencia") & "','Mail',1,'" & Replace(Req.Form("nombre"),"'","''") & "',"
	sQuery = sQuery & "'" & Replace(Req.Form("dir1"),"'","''")  & "', '" & Replace(Req.Form("ciudad"),"'","''") & "','" & Replace(Req.Form("dir2"),"'","''") & "','" & Req.Form("cp") & "',"
	sQuery = sQuery & Req.Form("fuente") & ",'" & LimpiaNum(Req.Form("tel1")) & "','" & LimpiaNum(Req.Form("tel2")) & "','" & LimpiaNum(Req.Form("tel3")) & "','" & LimpiaNum(Req.Form("tel4")) & "',"
	sQuery = sQuery & "'" & Req.Form("email") & "','" & Session("Usuario") & "','" & Req.Form("password") & "','" & Req.Form("mostrartel1") & "',"
	sQuery = sQuery & "'" & Req.Form("mostrartel2") & "','" & Req.Form("mostrartel3") & "',"
	sQuery = sQuery & "'" & Replace(Req.Form("instrucciones"),"'","''") & "','" & Req.Form("mostrartel4") & "','" & Req.Form("mostraremail") & "',GETDATE(),GETDATE())"

	Err.Clear
	rst.Open sQuery, sProvider
	if Err then
		Mail "diego@domoh.com", "Error en InsertUsuario", sQuery & " - " & Err.Description
		Response.End
	end if
	Session("EMail")=Req.Form("email")		
	Session("Password")=Req.Form("password")
	Session("Nombre")=Req.Form("nombre")
	Session("Id")=""
End Sub

'---- FormDatosPersonales pide nombre, teléfono, email y usuario
Sub FormDatosPersonales
%>
  	<div><p><%=MesgS("Datos Personales","Personal Details")%></p></div>
	<div>
		*<%=MesgS("Nombre de Contacto","Contact Name")%>
		<input title='<%=MesgS("Tu nombre y apellidos","First and Last Names")%>' maxlength=35 name=nombre size=20 autocomplete='name' > <input type=checkbox name=esagencia /><%=MesgS("Somos una Agencia","We are an Agency")%></div>
	<div><%=MesgS("Forma de Contacto","Contact Details")%></div>
	<div>
        <%=MesgS("Teléfono 1","Phone #1")%>:  <input title='<%=MesgS("Teléfono principal","Main Phone Number")%>' maxlength=30 name=tel1 size=20 autocomplete='tel'/>
        <input title='<%=MesgS("Pulsa si quieres que te llamen a este","Check if you want users to call you there")%>' type=checkbox name=mostrartel1 checked /><%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
	<div>
        <%=MesgS("Teléfono 2","Phone #2")%>:  <input title='<%=MesgS("Segundo teléfono","Second Phone Number")%>' maxlength=30 name=tel2 size=20 />
        <input title='<%=MesgS("Pulsa si quieres que te llamen a este","Check if you want users to call you there")%>' type=checkbox name=mostrartel2 checked /><%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
	<div>
        <%=MesgS("Teléfono 3","Phone #3")%>:  <input title='<%=MesgS("Tercer teléfono","Third Phone Number")%>' maxlength=30 name=tel3 size=20 />
        <input title='<%=MesgS("Pulsa si quieres que te llamen a este","Check if you want users to call you there")%>' type=checkbox name=mostrartel3 checked /><%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
	<div>
        <%=MesgS("Teléfono 4","Phone #4")%>:  <input title='<%=MesgS("Cuarto teléfono","Fourth Phone Number")%>' maxlength=30 name=tel4 size=20 />
        <input title='<%=MesgS("Pulsa si quieres que te llamen a este","Check if you want users to call you there")%>' type=checkbox name=mostrartel4 checked /><%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
	<div><%=MesgS("Instrucciones para llamarte","Calling Preferences")%>:  <input title='<%=MesgS("Hora y modo de contacto","Time and means of contact")%>' name=instrucciones size=56 /></div>
	<div>*  e-Mail:  <input maxlength=60 name=email size=20 autocomplete='email' /><input type=checkbox name=mostraremail checked /><%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
	<div>
        <p>
            <%=MesgS("ADVERTENCIA: POR FAVOR RELLENA CORRECTAMENTE LA DIRECCIÓN DE CORREO. ","WARNING: IT IS IMPORTANT THAT YOU FILL IN CORRECTLY YOUR E-MAIL ADDRESS.")%>
			<%=MesgS("TE ENVIAREMOS UN MENSAJE CON INSTRUCCIONES PARA ACTIVAR TU ANUNCIO. ","WE WILL SEND YOU THERE AN E-MAIL WITH INDICATIONS FOR CONFIRMING YOUR ADVERT.")%>
			<%=MesgS("EN CASO DE QUE NO LO HAGAS NO PODREMOS PUBLICAR TU ANUNCIO","OTHERWISE WE WON'T BE ABLE TO PUBLISH YOUR ADVERT")%>. 
			<% Response.Write MesgS("En caso de no recibir nuestro correo, por favor, comprueba tu correo Spam, Junk Mail o Correo no deseado", _
                "If you don't get our e-mail please check your 'Spam', 'Junk Mail' or 'Unwanted' folders")%></p></div>
    <div><%=MesgS("Identificación en domoh","Domoh Identification")%></div>
    <div>* <%=MesgS("Nombre de Usuario","User Name")%>:  <input title='<%=MesgS("Nombre que te identificará","Id to identify yourself in domoh")%>' maxlength=60 name=usuario size=20 autocomplete='username' /></div>
    <div>* <%=MesgS("Elige una Clave","Choose a Password")%>:       <input type=password maxlength=20 size=10 name=password autocomplete="new-password"/></div>
    <div><div>* <%=MesgS("Tecléala de Nuevo","Type it again")%>: <input type=password maxlength=20 size=10 name=password2 autocomplete="new-password"/></div><div>*<%=MesgS("¿Cómo nos conociste?","How did you get to Domoh?")%><% SelectFuente 0 %></div></div>
<%
End Sub

'---- NuevoAnuncio Contabiliza el nuevo anuncio en el usuario
Sub NuevoAnuncio
	sQuery="UPDATE Usuarios set numanuncios=numanuncios+1 WHERE usuario='" & Session("Usuario") & "'"
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en NuevoAnuncio", sQuery & " - " & Err.Description
End Sub

'---- Contenido muestra un banner de vRst
Sub Contenido(vRst)
	Response.Write rst("texto")
%>
<%	if right(rst("banner"),3) = "oldswf" then %>
	<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" id=banner>
		<param name=movie value='<%=rst("banner")%>'/><param name=quality value=high /><param name=bgcolor value=#FFFFFF />
		<embed src='<%=rst("banner")%>' type=application/x-shockwave-flash pluginspace='http://www.macromedia.com/go/getflashplayer'/></object>
<%	else %>
	<a href='<%=rst("html")%>' target=_new>
<%	    if right(rst("banner"),3) = "swf" then %>
	    <img src='<%=Replace(rst("banner"),".","") & ".jpg"%>' />
<%	    elseif rst("banner") <> "" then %>
		<img src='<%=rst("banner")%>' />
<%		end if %>
		</a>
<%	
	end if 
End Sub

'---- Fecha formatea [vFecha]
Sub Fecha(vFecha)
    dim sHora

	sHora=Hour(vFecha) & ":" & Minute(vFecha)
	if sHora="0:0" then
		Response.Write FormatDateTime(vFecha,2)
	else
		Response.Write FormatDateTime(vFecha,2) & " " & FormatDateTime(vFecha,4)
	end if
End Sub

'---- Mascota saca el icono correspondiente a [vMascota]
Sub Mascota(vMascota)
    if UCase(vMascota) = "ON" then 
%>
    <img src=images/Perrillos.gif alt='Admiten Mascotas'/>
<%	else %>
    <img src=images/NoPerrillos.gif alt='No Admiten Mascotas'/>
<%		        
    end if 
End Sub
%>