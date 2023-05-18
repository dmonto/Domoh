<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.cabecera, "Heading")) return false; if(!check(document.frm.anuncio, "What are you looking for?")) return false;
	    if(!check(document.frm.ciudad, "Region")) return false;
<%  if Request("id")="" then %>
	    if(!check(document.frm.nombre, "Name")) return false; if(!check(document.frm.email, "e-mail")) return false;
	    if(!check(document.frm.usuario, "User Name")) return false;	if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if(!checkSelect(document.frm.fuente, "Where did you hear about us?")) return false;
<%  end if %>
   	    return true;
    }
</script>
<%
	dim vcabecera, vanuncio, vfuente, vprovincia, vciudad, vmaximo, vidiomaes, vidiomaen, sBody

	vprovincia=0
	vidiomaes=""
	vidiomaen=" checked"

	if Request("id")<>"" then
		if Request("op")="Borrar" then
		 	sQuery = "UPDATE Anuncios SET activo='No',fechaultimamodificacion=#" & Format(Now) & "# WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody = Session("Usuario") & " ha desactivado el anuncio #" & Request("id")
			Mail "hector@domoh.com", "Comprador Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario.asp?msg=Ad+Hidden."
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si',fechaultimamodificacion=#" & Format(Now) & "# WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el anuncio #" & Request("id")
			Mail "hector@domoh.com", "Comprador Reactivado", sBody
			Session("id")=""
			Response.Redirect "NuCasaInquilino.asp?msg=Ad+Activated."
		elseif Request("op")="Modificar" then
			sQuery="SELECT * FROM Inquilinos i INNER JOIN Anuncios a ON i.id=a.id WHERE a.id=" & Request("id")
			rst.Open sQuery, sProvider
			vcabecera=rst("cabecera")
			vanuncio=rst("descripcion")
			vprovincia=rst("provincia")
			vciudad=rst("ciudad")
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes="" else vidiomaes=" checked"
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen=""
			vmaximo=rst("maximo")
			rst.Close
			Session("Id")=Request("id")
		end if
	end if
%>
<body onload="window.parent.location.hash='top';"> 
<form method=post name=frm onsubmit='return checkForm();' action=NuCasaCompraDemandaEn.asp>
	<input type=hidden name=id value=<%=Request("id")%> />
<div class=container>
	<nav><!--#include file="IncNuSubMenu.asp"--></nav>
	<div class="row">
<%	if Request("op")="Destacar" then %>
        <!--#include file="IncTrDestacado.asp"-->
<%	
		Response.End
	end if 
%>
		</div>
    <h1>Post a house/flat request for buying</h1>
    <!--#include file="IncNuCondUso.asp"-->
	<h2 class=seccion>* Ad Heading: <input title='This is what users will see first' type=text name=cabecera size=50 value="<%=vcabecera%>"/></h2>
	<div><p> * How is the property you are looking for?</p></div>
	<div>
		<div class=seccion>
			<textarea title='Text of the Advert' cols=35 rows=7 name=anuncio><%=vanuncio%></textarea>
			<input title='Check if the advert is in Spanish' type=checkbox name=idiomaes <%=vidiomaes%> value=ON />Texto en Español
			<input title='Check if the advert is in English' type=checkbox name=idiomaen <%=vidiomaen%> value=ON />Texto en Inglés</div>
		<div>
            * Preferred City/Region:<input title='Name of the city' type=text name=ciudad size=25 value="<%=vciudad%>"/>
            Maximum price you are willing to pay <input title='Price in thousand Euros' size=5 name=maximo value="<%=vmaximo%>"/>  thousand Euros</div></div>
<% if Request("id")="" then %>
	<div>Personal Details</div>
	<div><p>* Contact Name:  <input maxlength=25 name=nombre size=20 /></p></div>
    <div>Contact Details</div>
	<div>
		<div>Phone 1:       <input title='Type in all the digits' maxlength=30 name=tel1 size=20 /></div>
		<div><input title='Check if you want other users to see this number' type=checkbox name=mostrartel1 checked value=ON />Allow view</div></div>
	<div>
		<div>Phone 2:       <input title='Type in all the digits' maxlength=30 name=tel2 size=20 /></div>
		<div><input title='Check if you want other users to see this number' type=checkbox name=mostrartel2 checked value=ON />Allow view</div></div>
	<div><div>Phone 3:       <input title='Type in all the digits' name=tel3 size=20 /></div><div><input type=checkbox name=mostrartel3 checked value=ON />Allow view</div></div>
	<div><div>Phone 4:       <input maxlength=30 name=tel4 size=20 /></div><div><input type=checkbox name=mostrartel4 checked value=ON />Allow view</div></div>
	<div>Contact Instructions: <input name=instrucciones size=56 /></div>
	<div><div>*  e-Mail:  <input maxlength=60 name=email size=20 /></div><div><input type=checkbox name=mostraremail checked value=ON />Allow view</div></div>
	<div class="seccion">
        <p>
            WARNING: PLEASE CHECK YOUR EMAIL ADDRESS. WE WILL SEND YOU AN EMAIL WITH INSTRUCTIONS TO ACTIVATE YOUR ADVERT. YOU WILL HAVE TO ANSWER THIS EMAIL.  
            IF YOU DON'T FOLLOW THE INSTRUCTIONS WE WILL NOT POST YOUR ADVERT. HOTMAIL USERS AND OTHERS: PLEASE CHECK OUT YOUR "Junk Mail" or "Spam" </p></div>
    <div>Logon details</div>
    <div><div>* User name:  </div><div><input maxlength=60 name=usuario size=20 /></div></div>
    <div><div>* Password:       </div><div><input type=password maxlength=20 size=10 name=password /></div><div>*How did you know domoh.com? <% SelectFuente 0 %></div></div>
    <div><div>* Write the password      again: </div><div><input type=password maxlength=20 size=10 name=password2 /></div></div>
	<div><p><input type=submit value='Post Advert' class=btnForm /></p></div>
<%	elseif Request("id")="nuevo" then %>
	<div><p><input type=submit value='Add Advert' class=btnForm /></p></div>
<%	else %>
	<div><p><input type=submit value='Update Advert' class=btnForm /></p></div>
<%	end if %>
</div></form>
<!-- #include file="IncPie.asp" -->
</body>
