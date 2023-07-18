<!-- #include file="IncNuBD.asp" -->
<% if Session("Usuario")<>"" and Session("Activo")="Si" then Response.Redirect "QuHomeUsuario.asp" %>
<html>
<head><link href=TrDomoh.css rel=stylesheet type=text/css><title>Domoh - LogIn</title></head>
<body onload="window.parent.location.hash='top';">
<div class=container>
    <div class=logo><a title='Página Principal' href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%></a></div>
	<div class=main>
        <div class="col-10">
            <form action=TrLogOn.asp method=post target=principal>
            <div><p><% Response.Write MesgS("ENTRA en tu Perfil. Si tienes usuario de uno de estos tipos, púlsalo.", "Sign In")%></p></div>
            <div>Tenias ya usuario de domoh? Introdúcelo debajo.</div>
            <div>
    	        <div><p><%=MesgS("Usuario","Username")%>:</p></div>
                <div><p><input title='<% Response.Write MesgS("Nombre de usuario que nos diste al registrarte", "User id with which you registered on domoh")%>' size=18 name=id autocomplete="username" /></p></div></div>
	        <div>
    	        <div><p><%=MesgS("Contraseña","Password")%>:</p></div>
                <div><p><input title='<% Response.Write MesgS("Clave que nos diste al registrarte", "Password with which you registered on domoh")%>' type=password size=18 value='' name=password autocomplete="current-password" /></p></div></div>
            <div><input name=Submit type=submit class=btnLogin value='<%=MesgS("Entrar","Enter")%>'/></div>
            <div><a href=NuOlvidoFront.asp title='<%=MesgS("Te la recordamos por e-mail","We'll send you an e-mail with your password")%>' class=linkTxtSub><%=MesgS("Olvidé mi Contraseña","Forgot my Password")%></a> </div></form></div>
    <div>
        <div><a href=http://www.communitas.es target=_blank><img src=conts/Communitas.gif /></a></div>
        <div>
            <script type=text/javascript>google_ad_client = "pub-0985599315730502"; google_ad_slot = "4588720766"; google_ad_width = 120; google_ad_height = 240;</script>
            <script type=text/javascript src='http://pagead2.googlesyndication.com/pagead/show_ads.js'></script></div>
        <div><a title='Aula Creactiva' href=http://www.aulacreactiva.com/ target=_blank><img src=conts/Creactiva.gif /></a></div>
        <div><a href=http://www.materiagris.es target=_blank><img src=conts/MatGris.gif id=banner_4 /></a></div></div></div>
    <div><!-- #include file="IncPie.asp" --></div></div></body></html>