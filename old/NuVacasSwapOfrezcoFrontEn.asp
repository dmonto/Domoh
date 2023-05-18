<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vtipo, vdir1, vciudadnombre, vdir2, vcp, vzona, vdescripcion, vmascota, vidiomaes, vidiomaen, vfoto, sBody

	vmascota=" checked"
	vidiomaes=""
	vidiomaen=" checked"
	vfoto=" checked"

	if Request("id")<>"" then
		if Request("op")="Borrar" then
		 	sQuery = "UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Piso Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=House+Hidden."
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el piso (" & rst("cabecera") & ")"
			rst.Close
			Mail "hector@domoh.com", "Piso Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=House+Republished."
		elseif Request("op")="Modificar" then
			sQuery="SELECT * FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.id=" & Request("id")
			rst.Open sQuery, sProvider
			vciudadnombre=rst("ciudadnombre")
			vzona=rst("zona")
			if rst("rentaviv") > 0 then vrentaviv=rst("rentaviv")
			vdescripcion=rst("descripcion")
			vdir1=rst("dir1")
			vdir2=rst("dir2")
			vcp=rst("cp")			
			if UCase(rst("mascota"))<>"ON" then vmascota=""
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes="" else vidiomaes=" checked"
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen=""
			if IsNull(rst("foto")) or rst("foto")="" then vfoto=""
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrVacasSwapOferta.asp?op=FotoBorrada"
		end if
	end if
%>
<head>
    <!-- #include file="IncTrCabecera.asp" -->
    <script type=text/javascript>
        function checkForm() {
	        if(!check(document.frm.ciudadnombre, "City")) return false;	if(!check(document.frm.zona, "Area")) return false;
	        if(!check(document.frm.descripcion, "Description")) return false;
<%	if Request("id")="" then %>
	        if(!check(document.frm.nombre, "Contact Name")) return false; if(!check(document.frm.email, "e-mail address")) return false;
	        if(!check(document.frm.usuario, "User Name")) return false; if(!checkSelect(document.frm.fuente, "Where did you hear about us?")) return false;
	        if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	        if ((document.frm.email.value.length && document.frm.mostraremail.checked) || (document.frm.tel1.value.length && document.frm.mostrartel1.checked) ||
    		    (document.frm.tel2.value.length && document.frm.mostrartel2.checked) ||	(document.frm.tel3.value.length && document.frm.mostrartel3.checked) ||	(document.frm.tel4.value.length && document.frm.mostrartel4.checked))
	            return true;
	        alert("We need to display some kind of contact information"); return false;
<%	else %>
	        return true;
<%	end if %>
        }
    </script>
    <title>Holidays - New Advert for Swapping</title>
</head>
<body onload="window.parent.location.hash='top';">
<form name=frm onsubmit='return checkForm();' action=TrVacasSwapOferta.asp method=post>
	<input type=hidden name=id value=<%=Request("id")%> />
<div>
	<nav><!--#include file="IncNuSubMenu.asp"--></nav>
<%	if Request("op")="Destacar" then %>
    <!--#include file="IncTrDestacado.asp"-->
<%	
		Response.End
	end if 
%>
	<div>Post an Apartment for Swapping</div>
	<div>
		<p>
            Posting your advert means that you agree on our <a title='Read Terms & Conditions' href=CondUsoEn.htm>Terms and Conditions</a> and on our 
		    <a title='Read Privacy Policy' href=ProtDatosEn.asp target=_blank>Privacy and Safety</a> policy</p></div>
	<div>* are required fields</div
	<div><p>How is the apartment you want to swap?</p></div>
	<div>Where is it?</div>
	<div>
		<div>
			* City: <input title='Town or City of the Lodging' maxlength=50 size=19 name=ciudadnombre value='<%=vciudadnombre%>'/>
			* Area:	<input title='Area or District' maxlength=45 size=30 name=zona value='<%=vzona%>'/></div>
		<div>
			Address More accuracy will produce a better map <input title='Type in the address' maxlength=100 name=dir1 value="<%=vdir1%>" size=34 />
			Address (2) <input title='Type in the address (continued)' maxlength=100 name=dir2 value="<%=vdir2%>" size=34 /> ZIP Code <input title='Important for the Map' maxlength=20 size=10 name=cp value='<%=vcp%>'/></div></div>
	<div>
		* Please describe the property: <textarea title='How big and accesible it is...' cols=35 rows=7 name=descripcion><%=vdescripcion%></textarea>
		<input title='Check if the text is in English' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Ad in English
		<input title='Check if the text is in Spanish' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Ad in Spanish
		<input title='If checked you will be able to upload pictures' type=checkbox name=foto <%=vfoto%> />Do you want to upload pictures of the property? Nobody will contact you if they can't see pictures!</div will allow people to come with pets' type=checkbox name=mascota <%=vmascota%> value=ON />Pets Allowed</div>
<%	if Request("id")="" then %>
  	<div><p>Personal Details</p></div>
	<div>*Contact Name <input maxlength=35 name=nombre size=20 /><input title='Check if you are a professional' type=checkbox name=esagencia />Agency?</div>
	<div>Contact Details</div>
	<div><div>Phone 1:       <input title='Type in the complete number' maxlength=30 name=tel1 size=20 /></div><div><input type=checkbox name=mostrartel1 checked value=ON />Allow view</div></div>
	<div><div>Phone 2:       <input maxlength=30 name=tel2 size=20 /></div><div><input type=checkbox name=mostrartel2 checked value=ON />Allow view</div></div>
	<div><div>Phone 3:       <input name=tel3 size=20 /></div><div><input type=checkbox name=mostrartel3 checked value=ON />Allow view</div></div>
	<div><div>Phone 4:       <input maxlength=30 name=tel4 size=20 /></div><div><input type=checkbox name=mostrartel4 checked value=ON />Allow view</div></div>
	<div>Contact Instructions: <input name=instrucciones size=56 /></div>
	<div><div>*   e-Mail:  <input maxlength=60 name=email size=20 /></div><div><input type=checkbox name=mostraremail checked value=ON />Allow view</div></div>
	<div>
        <p>
            WARNING: PLEASE CHECK YOUR EMAIL ADDRESS. WE WILL SEND YOU AN EMAIL WITH INSTRUCTIONS TO ACTIVATE YOUR ADVERT. YOU WILL HAVE TO ANSWER THIS EMAIL.  
            IF YOU DON'T FOLLOW THE INSTRUCTIONS WE WILL NOT POST YOUR ADVERT. HOTMAIL USERS AND OTHERS: PLEASE CHECK OUT YOUR "Junk Mail" or "Spam" </p></div<div><div>* User name:  </div><div><input maxlength=60 name=usuario size=20 /></div></div>
    <div><div>* Password:       </div><div><input type=password maxlength=20 size=10 name=password /></div><div>* How did you know domoh.com? <% SelectFuente 0 %></div></div>
    <div><div>* Write the password      again:  </div><div><input type=password maxlength=20 size=10 name=password2 /></div></div>
	<div><p><input type=submit value='Post Apartament'/></p></div>
<%	elseif Request("id")="nuevo" then %>
	<div><p><input type=submit value='Add Apartment' /></p></div>
<%	else %>
	<div><p><input type=submit value='Update Apartment' /></p></div>
<%	end if %>
    </div file="IncPie.asp" -->
</body>
