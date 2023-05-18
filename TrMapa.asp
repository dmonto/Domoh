<!-- #include file="IncNuBD.asp" -->
<% 	
	dim sDir, sId
	
	sId=Request("id")
    if sId="" then
        sQuery="SELECT max(id) as maxId FROM Pisos WHERE dir1<>''" 
    	rst.Open sQuery, sProvider
    	if not rst.Eof then	sId=rst("maxId")
        rst.Close
    end if

    sQuery="SELECT * FROM Pisos WHERE id=" & sId
	rst.Open sQuery, sProvider
	if not rst.Eof then	sDir=rst("dir1") & " " & rst("dir2") & " " &  rst("cp") & " " & rst("ciudadnombre")
%>
<html>
<head>
	<title>Domoh - Mapa</title>
    <script src='http://maps.google.com/maps?file=api&v=2&key=ABQIAAAALJ_MoMrY7OsX8_SSig2N0hQcVdSNv70PRsmToXpRep_2R5ABmxQPrrtpNp8CKFsuKeMrNIaqJCVEfw' type=text/javascript></script>
    <script type=text/javascript>
	    function load(address) {
	        var map = new GMap2(document.getElementById("map")); map.addControl(new GSmallMapControl()); map.addControl(new GMapTypeControl()); map.width=600; map.height=400;
		    var geocoder = new GClientGeocoder();
        
		    if (geocoder) {
    			geocoder.getLatLng(address,
				    function(point) {
    					if (!point) {
	    					alert("<%=MesgS("No hay mapa para ","No map for ")%>" + address); window.close();
					    } else {
    						map.setCenter(point, 13);
						    var marker = new GMarker(point); map.addOverlay(marker); marker.openInfoWindowHtml('<font size=2 face=Arial><br/>'+address+'</font>');
						    map.zoomIn(); map.zoomIn();}}
			    );
		    }
	    }
    </script></head>
<body onload="load('<%=sDir%>')" onunload='GUnload()'><div id=map></div></body>
</html>
