<!-- #include file="IncNuBD.asp" -->
<% if Session("Usuario")<>"" and Session("Activo")="Si" then Response.Redirect "QuHomeUsuario.asp" %>
<% Response.Charset = "windows-1252" %>
<!doctype html>
<html>
<head>
    <title>Domoh</title>
    <meta name='viewport' content='width=device-width, initial-scale=1'/>
	<link rel='stylesheet' href='https://code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.css' />
	<script src='https://code.jquery.com/jquery-1.11.1.min.js'></script>
	<script src='https://code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.js'></script>
    <script src='https://code.jquery.com/jquery-1.12.4.js'></script>
    <script src='https://code.jquery.com/ui/1.12.1/jquery-ui.js'></script>
    <script type=text/javascript src=forms<%=Session("Idioma")%>.js></script>
    <link rel=stylesheet href='//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css'>
    <link rel=stylesheet href=JQList.css>
    <link rel=stylesheet href=TrDomoh.css>
</head>
<body>
    <div class=container>
	    <div class=main>
                <form action=TrLogOn.asp method=post target=principal>
                        <% Response.Write MesgS("Tenias ya usuario de domoh? Introdúcelo debajo.", "Already enrolled? Sign In")%>
                        <p><%=MesgS("Usuario","Username")%>:</p>
                        <p><input title='<% Response.Write MesgS("Nombre de usuario que nos diste al registrarte", "User id with which you registered on domoh")%>' size=18 name=id autocomplete="username" /></p>
                        <p><%=MesgS("Contraseña","Password")%>:</p>
                        <p><input title='<% Response.Write MesgS("Clave que nos diste al registrarte", "Password with which you registered on domoh")%>' type=password size=18 value='' name=password autocomplete="current-password" /></p>
                        <p><input name=Submit type=submit class=btnLogin value='<%=MesgS("Entrar","Enter")%>'/></p>
                        <p><a href=NuOlvidoFront.asp title="<%=MesgS("Te la recordamos por e-mail","We'll send you an e-mail with your password")%>" class=linkTxtSub><%=MesgS("Olvidé mi Contraseña","Forgot my Password")%></a></p></form></div>
</div>
</body>
