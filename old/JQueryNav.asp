<!-- #include file="IncNuBD.asp" -->
<!doctype html>
<html>
<head>
    <title>Domoh</title>
    <meta charset='utf-8'/>
    <meta name='viewport' content='width=device-width, initial-scale=1'/>
    <link rel='stylesheet' href='http://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css'>
    <link rel='stylesheet' href='http://code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.css' />
    <script src='https://code.jquery.com/jquery-1.12.4.js'></script>  
    <script src='https://code.jquery.com/ui/1.12.1/jquery-ui.js'></script>
	<script src='http://code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.js'></script>
    <script>
        $(function() {
            $("#domohnavbar").navbar();
            $("#domohnavbar").click(function() {});
        $("#livivienda").on('click', function(){

   $.ajax({
              url: "JQViviendaBusco.asp",
              success: function(output){

                   $("#contentdomoh").html(output);

              }
         });

});
        });
    </script>
    <style>.ui-menu {width: 150px;}</style>
</head>
<body>
<!-- Start of first page: #one -->
<div data-role=page id=one>
    <div data-role=header data-position=fixed>
        <h1>Domoh</h1>
        <div id=domohnavbar data-role=navbar>
            <ul>
                <li id=livivienda><a>Vivienda</a></li>
                <li><a href=JQViviendaBusco.asp>Vacas</a></li>
                <li><a href=JQViviendaBusco.asp>Usuarios</a></li>
                <li><a href=JQViviendaBusco.asp>Menu</a></li></ul></div><!-- /navbar --></div><!-- /header -->
        <div data-role=content id=contentdomoh></div><!-- /content -->
        <div data-role=footer data-theme=d><h4>Page Footer</h4></div><!-- /footer --></div><!-- /page one -->
</body></html>
