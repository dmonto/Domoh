<!-- #include file="IncNuBD.asp" -->
<!doctype html>
<html>
<head>
    <title>Domoh</title>
    <meta charset='utf-8'/>
    <meta name='viewport' content='width=device-width, initial-scale=1'/>
    <link rel=stylesheet href='http://code.jquery.com/mobile/1.1.0/jquery.mobile-1.1.0.min.css'/>
    <script src='http://code.jquery.com/jquery-1.7.1.min.js'></script>
    <script src='https://code.jquery.com/jquery-1.12.4.js'></script>
    <script src='http://code.jquery.com/mobile/1.1.0/jquery.mobile-1.1.0.min.js'></script>
    <script src='https://code.jquery.com/ui/1.12.0/jquery-ui.js'></script>
    <script>
        <!-- $.get("http://domoh.com/JQViviendaBusco.asp?destino=/JQViviendaBusco.asp&tipo=alquiler&pais=34&provincia=0&precio=&Submit=Buscar", function( data ) {$("#contentdomoh").html(data);}); -->

        $(function() {
            $("#butmenu").button({icon: "star"});
            $("#butuser").click(function() {$.get("http://domoh.com/QuDomohLogOn.asp", function(data) {$("#contentdomoh").html(data);});});
            $("#aViv").button(); $("#aVaca").button(); $("#tabs").tabs(); $("#menu").menu();
            $("#navbar").click(function() {alert("navbar");});
        });
    </script>
    <style>.ui-menu {width: 150px;}</style>
</head>
<body>
<!-- Start of first page: #one -->
<div data-role=page id=one>
    <div data-role=header>
        <h1>Domoh</h1>
        <div data-role=navbar>
            <ul>
                <li><a href=JQViviendaBusco.asp>Vivienda</a></li><li><a href=JQViviendaBusco.asp>Vacas</a></li><li><a href=PeDomohLogOn.asp>Usuarios</a></li>
                <li><a href=TrVivienda.asp>Menu</a></li></ul></div><!-- /navbar --></div><!-- /header -->
    <div data-role=content id=contentdomoh>
        <div id=tabs><ul><li><a href=#tabs-1>Vivienda</a></li><li><a href=#tabs-2>Vacas</a></li><li><a href=#tabs-3>Usuarios</a></li><li><a href=#tabs-4>Menu</a></li></ul>
            <div id=tabs-1>
                <p>
                    Proin elit arcu, rutrum commodo, vehicula tempus, commodo a, risus. Curabitur nec arcu. Donec sollicitudin mi sit amet mauris. Nam elementum quam ullamcorper ante. 
                    Etiam aliquet massa et lorem. Mauris dapibus lacus auctor risus. Aenean tempor ullamcorper leo. Vivamus sed magna quis ligula eleifend adipiscing. Duis orci. 
                    Aliquam sodales tortor vitae ipsum. Aliquam nulla. Duis aliquam molestie erat. Ut et mauris vel pede varius sollicitudin. Sed ut dolor nec orci tincidunt interdum. Phasellus ipsum. 
                    Nunc tristique tempus lectus.</p></div>
            <div id=tabs-2>
                <p>
                    Morbi tincidunt, dui sit amet facilisis feugiat, odio metus gravida ante, ut pharetra massa metus id nunc. Duis scelerisque molestie turpis. 
                    Sed fringilla, massa eget luctus malesuada, metus eros molestie lectus, ut tempus eros massa ut dolor. Aenean aliquet fringilla sem. Suspendisse sed ligula in ligula suscipit aliquam. 
                    Praesent in eros vestibulum mi adipiscing adipiscing. Morbi facilisis. Curabitur ornare consequat nunc. Aenean vel metus. Ut posuere viverra nulla. Aliquam erat volutpat. Pellentesque convallis. 
                    Maecenas feugiat, tellus pellentesque pretium posuere, felis lorem euismod felis, eu ornare leo nisi vel felis. Mauris consectetur tortor et purus.</p></div>
            <div id=tabs-3>
                <p>
                    Mauris eleifend est et turpis. Duis id erat. Suspendisse potenti. Aliquam vulputate, pede vel vehicula accumsan, mi neque rutrum erat, eu congue orci lorem eget lorem. Vestibulum non ante. 
                    Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Fusce sodales. Quisque eu urna vel enim commodo pellentesque. Praesent eu risus hendrerit ligula tempus pretium. 
                    Curabitur lorem enim, pretium nec, feugiat nec, luctus a, lacus.</p>
                    <p>Duis cursus. Maecenas ligula eros, blandit nec, pharetra at, semper at, magna. Nullam ac lacus. Nulla facilisi. Praesent viverra justo vitae neque. Praesent blandit adipiscing velit. Suspendisse potenti. 
                    Donec mattis, pede vel pharetra blandit, magna ligula faucibus eros, id euismod lacus dolor eget odio. Nam scelerisque. Donec non libero sed nulla mattis commodo. Ut sagittis. 
                    Donec nisi lectus, feugiat porttitor, tempor ac, tempor vitae, pede. Aenean vehicula velit eu tellus interdum rutrum. Maecenas commodo. Pellentesque nec elit. Fusce in lacus. 
                    Vivamus a libero vitae lectus hendrerit hendrerit.</p></div>
                <div id=tabs-4>
                    <p>Mauris eleifend est et turpis. Duis id erat. Suspendisse potenti. Aliquam vulputate, pede vel vehicula accumsan, mi neque rutrum erat, eu congue orci lorem eget lorem. Vestibulum non ante. 
                    Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Fusce sodales. Quisque eu urna vel enim commodo pellentesque. Praesent eu risus hendrerit ligula tempus pretium. 
                    Curabitur lorem enim, pretium nec, feugiat nec, luctus a, lacus.</p>
                    <p>Duis cursus. Maecenas ligula eros, blandit nec, pharetra at, semper at, magna. Nullam ac lacus. Nulla facilisi. Praesent viverra justo vitae neque. Praesent blandit adipiscing velit. Suspendisse potenti. 
                    Donec mattis, pede vel pharetra blandit, magna ligula faucibus eros, id euismod lacus dolor eget odio. Nam scelerisque. Donec non libero sed nulla mattis commodo. Ut sagittis. 
                    Donec nisi lectus, feugiat porttitor, tempor ac, tempor vitae, pede. Aenean vehicula velit eu tellus interdum rutrum. Maecenas commodo. Pellentesque nec elit. Fusce in lacus. 
                    Vivamus a libero vitae lectus hendrerit hendrerit.</p></div></div><!-- /tabs -->
            <input type=button data-role=button id=aVaca />Vacaciones
            <h3>Show internal pages:</h3>
            <p><a href=#popup data-role=button data-rel=dialog data-transition=pop>Show page "popup" (as a dialog)</a></p></div><!-- /content -->
        <div data-role=footer data-theme=d><h4>Page Footer</h4></div><!-- /footer --></div><!-- /page one -->

        <!-- Start of second page: #two -->
        <div data-role=page id=two data-theme=a>
            <div data-role=header><h1>Two</h1></div><!-- /header -->
            <div data-role=content id=userdomoh data-theme=a>
                <a href='http://www.domoh.com/TrDomohLogOn.asp?idioma=<%=MesgS("Es","En")%>'>Users</a>            
                <h2>Two</h2>
                <p>I have an id of "two" on my page container. I'm the second page container in this multi-page template.</p>
                <p>
                    Notice that the theme is different for this page because we've added a few <code>data-theme</code> swatch assigments here to show off how flexible it is. You can add any content or widget 
                    to these pages, but we're keeping these simple.</p>
                <p><a href=#one data-direction=reverse data-role=button data-theme=b>Back to page "one"</a></p></div><!-- /content -->
            <div data-role=footer><h4>Page Footer</h4></div><!-- /footer --></div><!-- /page two -->
        
        <!-- Start of third page: #popup -->
        <div data-role=page id=popup>
            <div data-role=header data-theme=e><h1>Dialog</h1></div><!-- /header -->
            <div data-role=content data-theme=d>
                <h2>Popup</h2>
                <p>I have an id of "popup" on my page container and only look like a dialog because the link to me had a <code>data-rel="dialog"</code> attribute which gives me this inset look 
                and a <code>data-transition="pop"</code> attribute to change the transition to pop. Without this, I'd be styled as a normal page.</p>
                <p><a href=#one data-rel=back data-role=button data-inline=true data-icon=back>Back to page "one"</a></p></div><!-- /content -->
            <div data-role=footer><h4>Page Footer</h4></div><!-- /footer --></div><!-- /page popup -->
                <ul id=menu><li class=ui-state-disabled><div>Toys (n/a)</div></li><li><div>Books</div></li><li><div>Clothing</div></li><li><div>Electronics</div>
                <ul><li class=ui-state-disabled><div>Home Entertainment</div></li><li><div>Car Hifi</div></li><li><div>Utilities</div></li></ul></li><li><div>Movies</div></li><li><div>Music</div>
                <ul><li><div>Rock</div><ul><li><div>Alternative</div></li><li><div>Classic</div></li></ul></li><li><div>Jazz</div><ul><li><div>Freejazz</div></li><li><div>Big Band</div></li><li><div>Modern</div></li></ul></li>
                <li><div>Pop</div></li></ul></li><li class=ui-state-disabled><div>Specials (n/a)</div></li></ul>
</body></html>