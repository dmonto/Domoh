<!-- #include file="IncNuBD.asp" -->
<% if Session("Usuario")<>"" and Session("Activo")="Si" then Response.Redirect "QuHomeUsuario.asp" %>
<!-- #include file="IncTrCabecera.asp" -->
<body onload="window.parent.location.hash='top';" class="fondo">
<%  if Request("head")="si" then %>
<!--#include file="IncFrHead.asp"-->
<%  end if %>
<div class=container>
	<div class=logo><a title='P�gina Principal' href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
    <div class=row>
        <div class=col-1>
            <a title='Aula Creactiva' href=http://www.aulacreactiva.com/ target=_blank><img alt='Creactiva' src=conts/Creactiva.gif /></a>
            <a title='Materia Gris' href=http://www.materiagris.es target=_blank><img alt='Materia Gris' src=conts/MatGris.gif /></a></div>
        <div>
            <form action=TrLogOn.asp method=post target=principal>
            <div><p><%=MesgS("ENTRA en tu Perfil","Sign In")%></p></div>
            <div>
                <div><p><%=MesgS("Usuario","Username")%>:</p></div>
                <div><p><input title='<% Response.Write MesgS("Nombre de usuario que nos diste al registrarte", "User id with which you registered on domoh")%>' size=18 name=id /></p></div></div>
                <div>
                    <div><p><%=MesgS("Contrase�a","Password")%>:</p></div>
                    <div><p><input title='<% Response.Write MesgS("Clave que nos diste al registrarte",	"Password with which you registered on domoh")%>' type=password autocomplete="current-password" size=18 value='' name=password /></p></div></div>
                <div><input name=Submit type=submit class=btnLogin value='<%=MesgS("Entrar","Enter")%>'/></div>
                <div>
    			    <a href=NuOlvidoFront.asp title='<% Response.Write MesgS("Te la recordamos por e-mail","We'll send you an e-mail with your password")%>' class=linkTxtSub>
                        <%=MesgS("Olvid� mi Contrase�a","Forgot my Password")%></a></div></form></div>
            <div id=banners_lat>
                <div><a href=http://www.communitas.es target=_blank><img src=conts/Communitas.gif /></a></div>
                <div>
                    <script type=text/javascript>google_ad_client = "pub-0985599315730502";google_ad_slot = "4588720766"; google_ad_width = 120; google_ad_height = 240;</script>
                    <script type=text/javascript src='http://pagead2.googlesyndication.com/pagead/show_ads.js'></script></div>
                <div><a title='Aula Creactiva' href=http://www.aulacreactiva.com/ target=_blank><img src=conts/Creactiva.gif /></a></div>
                <div><a href=http://www.materiagris.es target=_blank><img src=conts/MatGris.gif id=banner_4 /></a></div></div></div></div>
<div><!-- #include file="IncPie.asp" --></div></body>