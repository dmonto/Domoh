<!-- #include file="IncNuBD.asp" -->
<%
	dim vnombre, vtel1, vtel2, vtel3, vtel4, vemail, vusuario, vpassword, vesagencia, vfoto
	dim vmostrartel1, vmostrartel2, vmostrartel3, vmostrartel4, vmostraremail, vinstrucciones, vperfil, vmostrarperfil, vtipo

	if Request("usuario")<>"" then Session("Usuario")=Request("usuario")
	if Request("op")="Test" then Session("Usuario")="dmonto"
	if Session("Usuario")="" then Response.Redirect "TrLogOn.asp"

	rst.Open "SELECT * FROM Usuarios WHERE usuario='" & Session("Usuario") & "'", sProvider
	vtipo=rst("tipo")
	vnombre=rst("nombre")
	if rst("tel1") <> 0 then vtel1=rst("tel1")
	if rst("tel2") <> 0 then vtel2=rst("tel2")
	if rst("tel3") <> 0 then vtel3=rst("tel3")
	if rst("tel4") <> 0 then vtel4=rst("tel4")
	vemail=rst("email")
	vusuario=rst("usuario")
	vfoto=rst("foto")
	vpassword=rst("password")
	vperfil=rst("perfil")
	if UCase(rst("mostrartel1"))="ON" then vmostrartel1=" checked"
	if UCase(rst("mostrartel2"))="ON" then vmostrartel2=" checked"
	if UCase(rst("mostrartel3"))="ON" then vmostrartel3=" checked"
	if UCase(rst("mostrartel4"))="ON" then vmostrartel4=" checked"
	if UCase(rst("mostraremail"))="ON" then vmostraremail=" checked"
	if UCase(rst("esagencia"))="ON" then vesagencia=" checked"
	if UCase(rst("mostrarperfil"))="ON" then vmostrarperfil=" checked"
	vinstrucciones=rst("instrucciones")
	rst.Close
%>
<!--#include file="IncTrCabecera.asp"-->
<script type=text/javascript>
function checkForm() {
	if(!check(document.frm.nombre, "<%=MesgS("Nombre","Name")%>")) return false; if(!check(document.frm.email, "E-mail")) return false;
<% if vtipo<>"Janrain" then %>
	if(!check(document.frm.usuario, "<%=MesgS("Nombre de Usuario","Username")%>")) return false;
	if(!checkPassword(document.frm.password, document.frm.password2)) return false;
<% end if %>
	if ((document.frm.email.value.length && document.frm.mostraremail.checked) || (document.frm.tel1.value.length && document.frm.mostrartel1.checked) ||
		(document.frm.tel2.value.length && document.frm.mostrartel2.checked) ||	(document.frm.tel3.value.length && document.frm.mostrartel3.checked) ||
		(document.frm.tel4.value.length && document.frm.mostrartel4.checked)) return true;
	alert("<% Response.Write MesgS("Debes activar alguna forma de contacto", "We need you to activate one way of contacting you")%>");
	return false;
}
</script>
<body onload="window.parent.location.hash='top';"> 
<form name=frm onsubmit='return checkForm();' action=TrUsuarioGraba.asp method=post enctype=multipart/form-data>
	<input type=hidden name=origen value="<%=Request("origen")%>"/>
<div class=container>
	<div class=row>
		<div class=col-6><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
		<div class=col-6><a href=QuHomeUsuario.asp class=linkutils><%=MesgS("Tu Home","Your Home Page")%> &gt;</a><%=MesgS("Modificar Perfil","Update Profile")%></div></div>
    <h1>
		<%=MesgS("Datos de Usuario","User Profile")%>
<% if vtipo="Janrain" then %>
		Tu usuario de domoh.com es el de <%=vpassword%>. Por favor en futuras visitas haz login mediante el botón correspondiente. Puedes completar tus datos abajo.
<% end if %>
		</h1>
    <h2>* <%=MesgS("son campos obligatorios","are required fields")%></h2>
    <div>
	    <div>
		    *<%=MesgS("Nombre de Contacto","Contact Name")%> <input title='<%=MesgS("Nombre y Apellidos","Complete Name")%>' maxlength=35 name=nombre value='<%=vnombre%>' size=20 autocomplete='name' />
            <input title='<% Response.Write MesgS("Pulsa si no eres un particular", "Check if not an individual")%>' type=checkbox name=esagencia <%=vesagencia%> /><%=MesgS("Somos una Agencia","We are an Agency")%></div>
        <div>
<% if vfoto="" or IsNull(vfoto) then %>
		    <%=MesgS("Deja aquí tu foto ","Upload your picture")%><input title='<% Response.Write MesgS("Elige el fichero con la foto", "Browse for the file with the picture")%>' type=file name=foto />
<% else %>
            <img alt='Foto' title='<% Response.Write MesgS("Para cambiar la foto pulsa en el botón", "Press the button for changing the picture")%>' width=100 class=bordeazul src="http://domoh.com/<%=vfoto%>" />
            <%=MesgS("Cambiar foto","Change Picture")%> <input title='<% Response.Write MesgS("Para cambiar la foto pulsa a la derecha","Press the button on the right for changing the picture")%>' type=file name=foto />
<% end if %>
			</div></div>
    <div><%=MesgS("Forma de Contacto","Contact Information")%></div>
    <div>
        <%=MesgS("Teléfono 1","Phone #1")%>: <input title='<% Response.Write MesgS("Teclea todos los números del teléfono", "Type in all the digits")%>' maxlength=30 name=tel1 value='<%=vtel1%>' size=20 autocomplete='tel' />
		<input title='<% Response.Write MesgS("Pulsa si este es el teléfono que quieres mostrar","Check if you want other users to call you there")%>' type=checkbox name=mostrartel1 <%=vmostrartel1%> />
		<%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
    <div>
        <%=MesgS("Teléfono 2","Phone #2")%>: <input title='<% Response.Write MesgS("Teclea todos los números del teléfono", "Type in all the digits")%>' maxlength=30 name=tel2 value='<%=vtel2%>' size=20 />
	    <input title='<% Response.Write MesgS("Pulsa si este es el teléfono que quieres mostrar", "Check if you want other users to call you there")%>' type=checkbox name=mostrartel2 <%=vmostrartel2%> />
		<%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
    <div>
        <%=MesgS("Teléfono 3","Phone #3")%>: <input title='<% Response.Write MesgS("Teclea todos los números del teléfono", "Type in all the digits")%>' maxlength=30 name=tel3 value='<%=vtel3%>' size=20 />
    	<input title='<% Response.Write MesgS("Pulsa si este es el teléfono que quieres mostrar","Check if you want other users to call you there")%>' type=checkbox name=mostrartel3 <%=vmostrartel3%> />
		<%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
    <div>
        <%=MesgS("Teléfono 4","Phone #4")%>: <input title='<% Response.Write MesgS("Teclea todos los números del teléfono", "Type in all the digits")%>' name=tel4 value='<%=vtel4%>' size=20 />
        <input type=checkbox name=mostrartel4 <%=vmostrartel4%> /> <%=MesgS("Visible para Usuarios","Show in Adverts")%></div>
    <div><%=MesgS("Instrucciones para llamarte","Special indications for calling you")%>: <input name=instrucciones value="<%=vinstrucciones%>" size=56 /></div></div>
<div>
    * e-Mail: <input maxlength=60 name=email value='<%=vemail%>' size=20 autocomplete="email" /><input type=checkbox name=mostraremail <%=vmostraremail%>/><%=MesgS("Visible para Usuarios","Show in Adverts")%>
<% if vtipo="Janrain" then %>
    <input type=hidden name=usuario value='<%=vusuario%>' /><input type=hidden name=password value="<%=vpassword%>" />
<% else %>
	</div>
<div><%=MesgS("Identificación en domoh","Login Information")%></div>
<div>* <%=MesgS("Nombre de Usuario","Username")%>: <input maxlength=60 name=usuario value='<%=vusuario%>' size=20 autocomplete="username"/></div>
<div>* <%=MesgS("Elige una Clave","Choose a Password")%>: <input type=password maxlength=20 size=10 name=password value='<%=vpassword%>' autocomplete="new-password" /></div>
<div>* <%=MesgS("Tecléala de Nuevo","Type it again")%>: <input type=password maxlength=20 size=10 name=password2 value='<%=vpassword%>' /></div>
<% end if %>
<div><p><input type=submit value='<%=MesgS("Modificar Datos","Update Info")%>' class=btnForm /></p></div></form>
<!-- #include file="IncPie.asp" -->
</body>
