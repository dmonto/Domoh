<!doctype html>
<html>
<head>
    <title>Domoh</title>
    <meta charset='utf-8'/>
    <meta name='viewport' content='width=device-width, initial-scale=1'/>
    <link rel='stylesheet' href='//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css'>
    <script src='https://code.jquery.com/jquery-1.12.4.js'></script>  
    <script src='https://code.jquery.com/ui/1.12.1/jquery-ui.js'></script>
    <script>
        $(function() {
            $("#tabs").tabs(); 
            $("#tabs").click(function() {});
            $("#tab-1").click(function() {});
        });
    </script>
    <style>.ui-menu {width: 150px;}</style>
</head>
<body>
<!-- Start of first page: #one -->
<div data-role=page id=one>
    <div data-role=header><h1>Domoh</h1></div><!-- /navbar -->
    <div data-role=content id=contentdomoh>
        <div data-role=tabs id=tabs>
            <ul><li><a href=JQViviendaBusco.asp>Vivienda</a></li><li><a href=JQViviendaBusco.asp>Vacas</a></li><li><a href=QuHomeUsuario.asp>Usuarios</a></li><li><a href=TrVivienda.asp>Menu</a></li></ul></div>
        </div><!-- /content -->
    <div data-role=footer data-theme=d><h4>Page Footer</h4></div><!-- /footer --></div><!-- /page one -->
</body></html>