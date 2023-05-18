<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vtipo, vdir1, vciudadnombre, vdir2, vcp, vzona, vprecio, vdescripcion, vfoto, vidiomaes, vidiomaen

	vtipo="Piso"
	vfoto=" checked"
	if Session("Idioma")="En" then
		vidiomaen=" checked"
		vidiomaes=""
	else
		vidiomaes=" checked"
		vidiomaen=""
	end if

	if Request("id")<>"" then
		if Request("op")="Borrar" then
		 	sQuery = "UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") &")"
			rst.Close
			Mail "hector@domoh.com", "Piso Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=House+Hidden."
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el piso (" & rst("cabecera") &")"
			rst.Close
			Mail "hector@domoh.com", "Piso Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=House+Republished."
		elseif Request("op")="Modificar" then
			sQuery="SELECT * FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.id=" & Request("id")
			rst.Open sQuery, sProvider
			vtipo=rst("tipo")
			vdir1=rst("dir1")
			vdir2=rst("dir2")
			vcp=rst("cp")
			vciudadnombre=rst("ciudadnombre")
			vzona=rst("zona")
			if rst("precio") > 0 then vprecio=rst("precio")
			vdescripcion=rst("descripcion")
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen=""
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes="" else vidiomaes=" checked" 
			if IsNull(rst("foto")) or rst("foto")="" then vfoto=""
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrCasaCompraOfrezco.asp?op=FotoBorrada"
		end if
	end if
%>
<!--#include file="IncTrCabecera.asp"-->
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.ciudadnombre, "Ciudad")) return false; if(!check(document.frm.zona, "Zona")) return false;
	    if(!check(document.frm.precio, "Precio")) return false;	if(!check(document.frm.descripcion, "Descripción")) return false;
	    if(!check(document.frm.nombre, "Nombre")) return false;
<%	if Request("id")="" then %>
	    if(!check(document.frm.email, "E-mail")) return false; if(!check(document.frm.usuario, "Nombre de Usuario")) return false;
	    if(!checkSelect(document.frm.fuente, "Cómo nos Conociste")) return false;
	    if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if ((document.frm.email.value.length && document.frm.mostraremail.checked) || (document.frm.tel1.value.length && document.frm.mostrartel1.checked) ||
		    (document.frm.tel2.value.length && document.frm.mostrartel2.checked) ||	(document.frm.tel3.value.length && document.frm.mostrartel3.checked) || (document.frm.tel4.value.length && document.frm.mostrartel4.checked))
	        return true;
	    alert("Debes activar alguna forma de contacto"); return false;
<%	else %>
	    return true;
<%	end if %>
    }
</script>
<body onload="window.parent.location.hash='top';"> 
<form name=frm onsubmit="return checkForm();" action=NuCasaCompraOfertaEn.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<nav class=logo><!--#include file="IncNuSubMenu.asp"--></nav>
<%	if Request("op")="Destacar" then %>
    <!--#include file="IncTrDestacado.asp"-->         
<%	
		Response.End
	end if 
%>
	<div class=banner>
		<h1>Post an apartment, house, office FOR SALE</h1></div>	
	<div class=main>
		<p>
    		Posting your advert means that you agree to our 
			<a title='Read Terms and Conditions' href=CondUsoEn.htm>Terms and Conditions</a> and on our <a title='Read Privacy Policy' href=ProtDatosEn.asp target=_blank>Privacy and Safety</a> policy</p></div></div>
<div>*are required fields</div>
<div><p>How is the property you want to sell?</p></div>
<div>What is it?</div>
<div>
	<input title="Check if it's an apartment" type=radio value=Piso name=tipo <% if vtipo="Piso" then Response.Write " checked"%> /> Apartment
	<input title="Check if it's a whole house" type=radio value=Casa name=tipo <% if vtipo="Casa" then Response.Write " checked"%> /> House 
	<input title='Check if none of the above' type=radio value=Local name=tipo <% if vtipo="Local" then Response.Write " checked"%> /> Office, parking... </div>
<div>Where is the property?</div>
<div>
	* Region <input title='City or Town' maxlength=50 size=19 name=ciudadnombre value='<%=vciudadnombre%>'/>
	* Area <input title='Neighbourhood/District' maxlength=45 size=30 name=zona value="<%=vzona%>"/>
	* Price (in thousands of €): <input title="Price in 000's Euros" maxlength=30 name=precio value='<%=vprecio%>' size=8 /></div>
<div>
	Address (optional) More accuracy will produce a better map <input title='First line of the address' maxlength=100 name=dir1 value="<%=vdir1%>" size=34 />
    Address 1 <input title='Second line of the address' maxlength=100 name=dir2 value="<%=vdir2%>" size=34 />
    ZIP Code/P.O. Box <input title='Used for the map' maxlength=20 size=10 name=cp value='<%=vcp%>'/></div>
<div>
	* Describe the property: <textarea title='Rooms, size, location...' cols=36 rows=7 name=descripcion><%=vdescripcion%></textarea>
	<input title="Check if the advert's text is in Spanish" type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Advert in Spanish
	<input title="Check if the advert's text is in English" type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Advert in English
	<input title='If you check we will ask you to upload pictures' type=checkbox name=foto <%=vfoto%> />Do you want to include images? Nobody will contact you if they can't see pictures!</div>
<%	if Request("id")="" then %>
<div><p>Personal Details</p></div>
<div>* Contact Name <input title='Your name' maxlength=35 name=nombre size=20 /><input type=checkbox name=esagencia value=ON />Estate Agency</div>
<div>Contact details</div>
<div><div>Phone 1:       <input maxlength=30 name=tel1 size=20 /></div><div><input type=checkbox name=mostrartel1 checked value=ON />Allow view</div></div>
<div><div>Phone 2:       <input maxlength=30 name=tel2 size=20 /></div><div><input type=checkbox name=mostrartel2 checked value=ON />Allow view</div></div>
<div><div>Phone 3:       <input name=tel3 size=20 /></div><div><input type=checkbox name=mostrartel3 checked value=ON />Allow view</div></div>
<div><div>Phone 4:       <input maxlength=30 name=tel4 size=20 /></div><div><input type=checkbox name=mostrartel4 checked value=ON />Allow view</div></div>
<div>Contact Instructions: <input name=instrucciones size=56 /></div>
<div><div>*  e-Mail:  <input maxlength=60 name=email size=20 /></div><div><input type=checkbox name=mostraremail checked value=ON />Allow view</div></div>
<div>
    <p>
        WARNING: PLEASE CHECK YOUR EMAIL ADDRESS. WE WILL SEND YOU AN EMAIL WITH INSTRUCTIONS TO ACTIVATE YOUR ADVERT. YOU WILL HAVE TO ANSWER THIS EMAIL.  
        IF YOU DON'T FOLLOW THE INSTRUCTIONS WE WILL NOT POST YOUR ADVERT. HOTMAIL USERS AND OTHERS: PLEASE CHECK OUT YOUR "Junk Mail" or "Spam" </p></div>
    <div>Logon details</div>
    <div><div>* User name:  </div><div><input maxlength=60 name=usuario size=20 /></div></div>
    <div><div>*Password:       </div><div><input type=password maxlength=20 size=10 name=password /></div><div>* How did you know domoh.com? <% SelectFuente 0 %></div></div>
    <div><div>* Write the password      again:  </div><div><input type=password maxlength=20 size=10 name=password2 /></div></div>
	<div><p><input type=submit value='Post House' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
	<div><p><input type=submit value='Add House' class=btnForm /></p></div>
<%	else %>
	<div><p><input type=submit value='Update House' class=btnForm /></p></div>
<%	end if %>
</form>
<!-- #include file="IncPie.asp" -->
</body>