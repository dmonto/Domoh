<!-- #include file="IncNuBD.asp" -->
<% 
	if Request("idioma")="Es" then 
		Session("Idioma")=""
	elseif Request("idioma")="En" then 
		Session("Idioma")="En"
	end if

	if Request("nuevoidioma")="Es" then 
		Session("Idioma")=""
		Response.Cookies("Idioma")="Es"
		Response.Cookies("Idioma").Expires=Now+365
	elseif Request("nuevoidioma")="En" then 
		Session("Idioma")="En"
		Response.Cookies("Idioma")="En"
		Response.Cookies("Idioma").Expires=Now+365
	end if
%>
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
    <script>
        $(document).ready(function() {
            $("#contentviv").html("Cargando...");
            $.ajax({
                url: "JQViviendaBusco.asp",
                success: function(output) {
                   $("#contentviv").html(output);
                   $("#accordion").accordion({heightStyle: "content"});
                },
                datatype: 'html',         
                contentType: "text/html;charset=windows-1252"
            });

            $("#contentvaca").html("Cargando...");
            $.ajax({
                url: "JQViviendaBusco.asp?tipo=vacas",
                success: function(output) {
                   $("#contentvaca").html(output);
                   $("#accordion").accordion({heightStyle: "content"});
                },
                datatype: 'html',         
                contentType: "text/html;charset=windows-1252"
            });

            $("#contentusu").html("Cargando...");
            $.ajax({
                url: "TabLogOn.asp",
                success: function(output) {
                   $("#contentusu").html(output);
                },
                datatype: 'html',         
                contentType: "text/html;charset=windows-1252"
            });

            $("#contentmen").html("Cargando...");
            $.ajax({
                url: "TabMenu.asp",
                success: function(output) {
                   $("#contentmen").html(output);
                },
                datatype: 'html',         
                contentType: "text/html;charset=windows-1252"
            });
        });
    </script>
    <style>
        li#tab1{margin: 0px 0px 0px 0px;} 
        li#tab2{margin: 0px 0px 0px 0px;} 
        li#tab3{margin: 0px 0px 0px 0px;} 
        li#tab4{margin: 0px 0px 0px 0px;} 
        #navbarviv, #navbarvac, #navbarusu, #navbarmen {float: none;}
    </style>
</head>
<body>
<div data-role=page>
	<div data-role=header>
        <h1>Domoh</h1>
	</div><!-- /header -->
	<div data-role=tabs>
        <div data-role=navbar id=domnavbar class='navbar navbar-fixed-top'>
            <ul class=nav>
                <li class=ui-tabs id="tab1"><a id=navbarviv href=#contentviv class='ui-tabs ui-btn-active ui-state-persist'><%=MesgS("Vivienda","Home")%></a></li>
                <li class=ui-tabs id="tab2"><a id=navbarvac href=#contentvaca class=ui-tabs><%=MesgS("Vacaciones","Holidays")%></a></li>
                <li class=ui-tabs id="tab3"><a id=navbarusu href=#contentusu class=ui-tabs><%=MesgS("Usuarios","Users")%></a></li>
                <li class=ui-tabs id="tab4"><a id=navbarmen href=#contentmen class=ui-tabs><%=MesgS("Menú","Menu")%></a></li></ul></div><!-- /navbar -->
    <div class=ui-content id=contentviv>Vivienda</div><!-- Cargar solo el contenido del div -->
    <div class=ui-content id=contentvaca>Vacas</div><!-- /content -->
    <div class=ui-content id=contentusu>Usuario</div><!-- /content -->
    <div class=ui-content id=contentmen>Menu</div><!-- /content -->
	</div><!-- /content -->
	<div data-role=footer>
		<h4><!-- #include file="IncPie.asp" --></h4>
	</div><!-- /footer -->
</div><!-- /page -->

</body></html>
