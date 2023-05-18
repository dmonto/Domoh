<% if Request("head")<>"no" then %>
<head data-hayidioma='<%=Session("HayIdioma")%>' data-idioma='<%=Session("Idioma")%>' data-req='<%=Request("idioma")%>'>
    <link rel=stylesheet type=text/css href=TrDomoh.css>
    <link rel='alternate' href='http://domoh.com/PeMenu.asp?nuevoidioma=En' hreflang='en'/>
	<script type=text/javascript src=forms<%=Session("Idioma")%>.js></script>
</head>
<% end if %>