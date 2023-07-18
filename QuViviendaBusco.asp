<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncPaginacion.asp" -->
<!-- #include file="IncQuDetalles.asp" -->
<%
	dim sWhere, sOrder, i, j, sNombreProvincia, numProvincia, sHora, sReq, numPais, numBueno, numRegular, numMalo, numVotos, numPisos, sBack, sIdioma, sCampoPrecio, sTabla, sCampoTipo, sCampoFoto, sDesc, sClass, sInvDesc
	on error goto 0
				 	
	if Request("tipo")="inquilino" or Request("tipo")="comprador" then
		sTabla="Inquilinos"
		sCampoPrecio="maximo"
		sCampoTipo="tipoviv"
		sCampoFoto="u.foto"
	else
		sTabla="Pisos"
		sCampoTipo="tipo"
		sCampoFoto="a.foto"
		if Request("tipo")="alquiler" or Request("op")="Test" then
			sCampoPrecio="rentaviv"
		elseif Request("tipo")="venta" then
			sCampoPrecio="precio"
		elseif Request("tipo")="vacas" or Request("tipo")="vacasswap" then
			sCampoPrecio="rentavacas"
		else
			Response.Write "Por favor intentelo de nuevo."
			Response.End
		end if
	end if
	sReq="tipo=" & Request("tipo") & "&"
			
	numProvincia=0
	if Request.Form("provincia")<>"" or Request.QueryString("provincia")<>"" then
		numProvincia=CLng(Request("provincia"))
		Response.Cookies("ProvinciaViv")=numProvincia
		Response.Cookies("ProvinciaViv").Expires=Now+30
		sReq=sReq & "provincia=" & Request("provincia") & "&"
	end if

	if Request("pais")<>"" then
		numPais=CLng(Request("pais"))
		sReq=sReq & "pais=" & Request("pais") & "&"
	end if

	sWhere= "WHERE a.activo='Si' AND u.activo='Si' "
	if Request("tipo")="vacasswap" then
		sWhere= sWhere & "AND rentavacas=0 AND rentaviv=0 AND precio=0 " 
	elseif Request("tipo")="inquilino" then
		sWhere= sWhere & "AND p.tipoviv <> 'Compra' " 
	elseif Request("tipo")="comprador" then		
		sWhere= sWhere & "AND p.tipoviv = 'Compra' " 
	else
		sWhere= sWhere & "AND p." & sCampoPrecio & " <> 0 " 
	end if
		
	if numProvincia<>0 then 
		sWhere= sWhere & " AND a.provincia = " & numProvincia
	elseif numPais<>0 then 
		sWhere= sWhere & " AND a.provincia IN (SELECT id FROM Provincias WHERE pais= " & numPais & ")"
	else
		sWhere= sWhere & " AND a.provincia <> 0 "
	end if
		
	if Request("precio") <> "" then 
		sWhere= sWhere & " AND p." & sCampoPrecio & " <= " & Request("precio")
		sReq=sReq & "&precio=" & Request("precio")
	end if
	
	sQuery="SELECT COUNT(*) AS numPisos FROM (" & sTabla & " p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.Usuario " & sWhere
	rst.Open sQuery, sProvider
	if Err then Mail "diego@domoh.com", "Error en TrCasaBusco", sQuery & " - " & Err.Description
	numPisos=rst("numPisos")
	rst.Close

	if Request("orden")="fecha" then
		if Request("desc")="" then sInvDesc="" else sInvDesc="DESC"
		sOrder= "a.fechaultimamodificacion " & sInvDesc & ", a.idiomaes DESC, p." & sCampoPrecio 
	elseif Request("orden")="tipo" then
		sOrder= "p.tipo " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, nummalo "
	elseif Request("orden")="ciudad" then
		sOrder= "ciudadnombre " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="foto" then
		sOrder= sCampoFoto & " " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="fuma" then
		sOrder= "p.fuma " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="mascota" then
		sOrder= "p.mascota " & Request("desc") & ", 2 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	elseif Request("orden")="precio" then
		sOrder= "p." & sCampoPrecio & " " & Request("desc") & ", a.idiomaes DESC, a.fechaultimamodificacion DESC"
	elseif Request("orden")="idioma" then
		sOrder= "a.idiomaes " & Request("desc") & ", a.fechaultimamodificacion DESC, p." & sCampoPrecio 
	elseif Request("orden")="valor" then
		sOrder= "2 " & Request("desc") & ", 1 " & Request("desc") & ", numbueno DESC, numregular DESC, NumMalo "
	else
		sOrder= "UPPER(ciudadnombre) " & Request("desc") & ", 8, 7 DESC, 1 DESC, numbueno DESC, numregular DESC, NumMalo "
	end if
	
	if Request("desc")="" then 
		sDesc="DESC" 
	else 
		sDesc=""
	end if
	
	sQuery="SELECT FLOOR(SUM(av.visitas)/(DATEDIFF(DAY,MIN(av.fecha),GETDATE())+1)*7/50) AS VistoSemana, FLOOR(u.numanuncios/5) as NumAnuncios, a.id AS aId, p.mascota AS pMascota, p.fuma AS pFuma, "
	sQuery=sQuery & " CONVERT(VARCHAR, a.fechaultimamodificacion, 3) AS aFechaUltimaModificacion, p." & sCampoPrecio & " AS pPrecio, p." & sCampoTipo & " AS pTipo, "
	sQuery=sQuery & sCampoFoto & " AS aFoto, a.destacado, a.idiomaes, a.idiomaen, esagencia, cabecera, "
	if sTabla="Pisos" then sQuery=sQuery & " zona, ciudadnombre, "
	sQuery=sQuery & " a.numbueno, a.nummalo, a.numregular "
	sQuery=sQuery & " FROM ((" & sTabla & " p INNER JOIN Anuncios a ON p.id=a.id) INNER JOIN Usuarios u ON a.usuario=u.Usuario) LEFT JOIN AnunciosVistos av ON a.id=av.anuncio " 
	sQuery=sQuery & sWhere & " GROUP BY a.id, p.mascota, p.fuma, a.fechaultimamodificacion, p." & sCampoPrecio & ", p." & sCampoTipo & ", " & sCampoFoto & ", a.destacado, a.idiomaes, a.idiomaen, esagencia, cabecera, "
	if sTabla="Pisos" then sQuery=sQuery & " zona, ciudadnombre, "
	sQuery=sQuery & " a.numbueno, a.nummalo, a.numregular, numanuncios ORDER BY " & sOrder
%>
<head>
    <title>Domoh - Habitaciones y Pisos Gay-Friendly</title>
    <link rel=stylesheet type=text/css href=TrDomoh.css>
    <style type=text/css>
        body {min-width:980px; min-height:800px; text-align:center;}
        nav {text-align: center; position:absolute; top:60px; left:120px; width:400px; z-index:1000;}
        /* Page-Structure rules */
        #idPage {/* Layout */ position: relative; width: 980px; min-height:800px; margin: 0 auto; padding: 0; text-align:left;}
        #idColumns {/* Layout */ position: relative; width:980px; height:3000px; overflow:hidden; margin: 0 0; padding: 0; /* Style */ background-color: transparent;}
        #idMainColumn {/* Layout */ margin: 0 0 0 150px; /* Style */ padding: 1px; background-color: #FFF;}
        #idSideColumn {/* Layout */ position:absolute; overflow:hidden; left:0; top:0; width: 150px; margin: 0 0; /* Style */ padding: 2px; min-height:600px;}
        #idFooter {/* Layout */ margin: 0; text-align:left; /* Style */ padding: 10px; background-color: #9CF; font-size: smaller;}
        .panel {display:block; border-bottom: white 3px solid; height: 40px;}
        #gallery {height: 320px; display:none;}
        #slideshow {display:none;}
        .gcap {font-size:xx-small; margin:3px; padding:0;}
        .scap {font-size:xx-small; margin:3px; padding:0;}
        .gcon { /* gallery image/caption container */ width:114px; height:120px; margin:5px; background:#CFD4E6; border:1px solid #BF8660; float:left;}
        .scon { /* slideshow image/caption container */ width:580px; margin:10px; background:#CFD4E6; border:1px solid #BF8660; position: absolute; width: 200px; left: 100px; top: 50px; padding: 4px; z-index: 101;}
        #gallery img {margin:6px 6px 3px 6px; background:#CFD4E6;}
        #slideshow img {color: #BF8660; background-color: #FFF; border: 2px solid #BF8660; margin:6px 6px 3px 6px; background:#CFD4E6;}
        #navigation {text-align:center; width: 200px; text-align: center;}
        #prev, #next, #back, #auto, #time {color:#BF8660; cursor:pointer; font-weight:bold; margin-right:20px;}
        .clearAll {clear:both; margin:0; padding:0;}
        .slide {position:absolute; top:10px; left:50px;}
    </style>
    <script type=text/javascript src=/x/xmenu6.js></script><script type=text/javascript src=/x/xaddeventlistener.js></script><script type=text/javascript src=/x/xgetelementbyid.js></script>
    <script type=text/javascript src=/x/xwalkul.js></script><script type=text/javascript src=/x/xfirstchild.js></script><script type=text/javascript src=/x/xnextsib.js></script>
    <script type=text/javascript src=/x/xgetcomputedstyle.js></script><script type=text/javascript src=/x/xtabpanelgroup.js></script><script type=text/javascript src=/x/x.js></script>
    <script type=text/javascript src=/x/xhttprequest.js></script>
    <script type=text/javascript>
    window.onload = setup;

    function setup() {
    	var ele, i = 1;

    	for (i=i;i<=<%=numPisos%>;i++) {new xTabPanelGroup('tpg'+i, 770, 310, 22, 'tabPanel', 'tabGroup', 'tabDefault', 'tabSelected');}
    	i=1; do {ele = xGetElementById('trigger' + i++);if (ele) {ele.onclick = tOnClick;}} while(ele);
        i=1; do {ele = xGetElementById('anuncio' + i);if (ele) {ele.panelId = 'tpg' + i++;}} while(ele);
    	i=1; do {ele = xGetElementById('tpgcab' + i);if (ele) {ele.panelId = 'tpg' + i++;}} while(ele);
    	i=1; do {ele = xGetElementById('cab' + i);if (ele) {ele.anuncioId = xGetElementById('fotos' + i).tag; ele.panelId = 'tpg' + i++; ele.onmouseover = tOnMouseover;}} while(ele);
    	i=1; do {ele = xGetElementById('bloquea' + i);if (ele) {ele.panelId = i++;ele.onclick = bloqueaEle;}} while(ele);
    	i=1; do {ele = xGetElementById('quita' + i);if (ele) {ele.panelId = i++;ele.onclick = quitaEle;}} while(ele);
    	i=1;
    	do {ele = xGetElementById('anuncio' + i);
		    if (ele) {
//  			xImgGallery(i++);
			    i++;
		    }
	    } while(ele);

	    i=1; do {ele = xGetElementById('tpg' + i++);if (ele) {ele.style.display='none';}} while(ele);
    }

    var bloqueaTodos=false; var galAnuncio;

    function tOnMouseover() {
    	var ele, eletr, i=1;
    	if (!bloqueaTodos) {
		    do {
    			eletr = xGetElementById('cabtr' + i); ele = xGetElementById('tpg' + i++);
			    if (ele) {
    				console.log(ele.id);
				    if (ele.id == this.panelId) {console.log(ele.id + ' block'); ele.style.display='block'; eletr.style.backgroundColor='blue'; if (!(ele.fotosCargadas)) {xImgGallery(i-1); ele.fotosCargadas=true;}}
				    else if (!(ele.bloquear)) {console.log(ele.id + ' borrar'); ele.style.display='none'; eletr.style.backgroundColor='';}
			    }
		    } while(ele);
	    }
    }

    function bloqueaEle() {var ele = xGetElementById('tpg' + this.panelId); ele.bloquear=true; ele.style.display='block';}
    function quitaEle() {var ele = xGetElementById('tpg' + this.panelId); ele.style.display='none'; ele = xGetElementById('cab' + this.panelId); ele.style.display='none';}

    var noClick=false;

    function tOnClick() {
	    var el, eletag, eleint, ele, stag, i=1, chk, chTodos;
	
	    chk=this.checked; chTodos=xGetElementById('trigger1');
	
	    if (this.id==chTodos.id) {
	        i=2; ele = xGetElementById('trigger' + i++);
		    do {
    			if (chTodos.checked && !ele.checked) {noClick=true; ele.click(); noClick=false;} else if (!chTodos.checked && ele.checked) {noClick=true; ele.click(); noClick=false;}
			    ele = xGetElementById('trigger' + i++);
		    } while(ele);
	    }
	    else if (!noClick) {
    		chTodos.checked=true; i=2; ele = xGetElementById('trigger' + i++);
		    do {
    			if (!ele.checked) {chTodos.checked=false;}
			    ele = xGetElementById('trigger' + i++);
		    } while(ele);
	    }
	
	    i=1; stag=this.name.substring(0,1); + '='; el = xGetElementById('anunciotag' + i);
	    do {
		    eletag = el.value; eleint = eletag.substring(eletag.indexOf(stag)+stag.length, eletag.length); ele = eleint.substring(1,eleint.indexOf('&'));
		    if (this.name.substring(1,this.name.length)==ele) {
    			if (chk) {console.log('anuncio' + i + ' mostrar'); xGetElementById('anuncio' + i).style.display='block';}
			    else {console.log('anuncio' + i + ' no mostrar'); xGetElementById('anuncio' + i).style.display='none';}
		    }
		    i++; el = xGetElementById('anunciotag' + i);
	    } while(el);
    }

    function changeTab(demo, tab) {xTabPanelGroup.instances['tpg' + demo].select(tab); if (tab==2) requestTable(); return false;}

    var imgsPerPg = 10; // number of img elements in the html
    var slideTimeout = 7; // seconds before loading the next slide

    var gPath = '/mini';  // gallery files (thumbnails) path, include trailing slash
    var sPath = '/'; // slideshow files (big imgs) path, include trailing slash

    var fotos = new Array(); var imgsMax = new Array(); var galMode = true;

    // There must be (imgsMax + 1) captions.
<%	
	set rstDet = Server.CreateObject("ADODB.recordset")
	rst.Open sQuery, sProvider
    for i=1 to numPisos
	    rstDet.Open "SELECT * FROM Fotos WHERE piso=" & rst("aId"), sProvider 
	    j=1		
%>
    fotos[<%=i%>] = new Array();
<%	    while not rstDet.Eof %>
    fotos[<%=i%>][<%=j%>] = '<%=rstDet("foto")%>';
<%	
	        j=j+1
		    rstDet.movenext
	    wend
	    rstDet.Close
	    rst.Movenext
%>
    imgsMax[<%=i%>] = <%=j-1%>;  // total number of images
<%
    next
    rst.Close
%>
    var autoSlide = false; var slideTimer = null; var slideCounter = 0; var currentSlide = 1;

    function xImgGallery(anuncio) {if (document.getElementById && document.getElementById('navigation').style) {gal_paint(anuncio, 1); document.getElementById('time').style.display = 'none';}}
    function gal_paint(anuncio, n) {gal_setImgs(anuncio, n); gal_setNav(anuncio, n);}

    function gal_setImgs(anuncio, n) {
    	var ssEle = document.getElementById('slideshow'); var galEle = document.getElementById('gallery'+anuncio);
    	var i, imgTitle, pth, img, id, src, ipp, idPrefix, imgSuffix, imgPrefix, zeros, digits, capEle, capPar;

    	if (galMode) {ipp = imgsPerPg; idPrefix = anuncio + 'g'; imgTitle = 'Click to view large image'; ssEle.style.display = 'none'; galEle.style.display = 'block'; pth = gPath;} 
	        else {currentSlide = n; ipp = 1; idPrefix = 's'; imgTitle = ''; ssEle.style.display = 'block'; galEle.style.display = 'none'; pth = sPath;}

	    for (i = 0; i < ipp; ++i) {
    		id = idPrefix + (i + 1); img = document.getElementById(id);
		
    		if ((n + i) <= imgsMax[anuncio]) {
	    		img.title = imgTitle; img.src = pth + fotos[anuncio][n+i]; // img to load now
			    img.onerror = imgOnError; img.onload = imgOnLoad; if (galMode) {img.style.width=105;} else {img.style.width=500;}
			    img.style.margin=5;
	
			    if (galMode) {
    				img.style.cursor = 'pointer'; img.slideNum = n + i; // slide img to load onclick
				    img.anuncio=anuncio; img.onclick = imgOnClick; img.style.display = 'inline';
			    }
		    } else {img.style.display = 'none'; img.parentNode.style.display = 'none';}
	    }
    }

    function imgOnClick() {galMode = false; bloqueaTodos=true; galAnuncio=this.anuncio; gal_paint(this.anuncio, this.slideNum);}
    function imgOnError() {var p = this.parentNode; if (p) p.style.display = 'none';}

    function imgOnLoad() {
    	var p = this.parentNode;
    	if (p) p.style.height=this.height*(105/this.width)+8; p = document.getElementById("scon1");
    	p.style.height=this.height*(500/this.width)+22; p.style.width=519; p = document.getElementById("s1"); p.style.margin=10;
    }

    function gal_setNav(anuncio, n) {
    	var ipp = galMode ? imgsPerPg : 1;
    	var e = document.getElementById('next');

    	// Next
    	if (n + ipp <= imgsMax[anuncio]) {e.nextNum = n + ipp; e.onclick = next_onClick; e.style.display = 'inline';} else {e.nextNum = 1;}

    	// Previous
    	e = document.getElementById('prev'); e.style.display = 'inline'; e.onclick = prev_onClick; if (n > ipp) {e.prevNum = n - ipp;} else {e.prevNum = galMode ? normalize(imgsMax[anuncio]) : imgsMax[anuncio];}

    	// Back
    	e = document.getElementById('back'); if (!galMode) {e.onclick = back_onClick; e.style.display = 'inline'; e.tag=anuncio; e.backNum = normalize(n);} else {e.style.display = 'none';}

    	// Auto Slide
    	e = document.getElementById('auto'); if (!galMode) {e.onclick = auto_onClick; e.style.display = 'inline'; e.tag = anuncio;} else {e.style.display = 'none';}
    }   

    function normalize(n) {return 1 + imgsPerPg * (Math.ceil(n / imgsPerPg) - 1);}
    function next_onClick(e) {gal_paint(galAnuncio, this.nextNum);}
    function prev_onClick(e) {gal_paint(galAnuncio, this.prevNum);}

    function back_onClick(e) {
    	galMode = true; bloqueaTodos=false;
    	if (slideTimer) {clearTimeout(slideTimer);}
    	autoSlide = false; gal_paint(this.tag, this.backNum); document.getElementById('time').style.display = 'none';
    }

    function auto_onClick(e) {
        var ele = document.getElementById('time'); 
        autoSlide = !autoSlide;
    	if (autoSlide) {slideCounter = 0; slideTimer = setTimeout("slide_OnTimeout("+ this.tag +")", slideTimeout); ele.style.display = 'inline';} else if (slideTimer) {clearTimeout(slideTimer); ele.style.display = 'none';}
    }

    function slide_OnTimeout(anuncio) {
    	slideTimer = setTimeout("slide_OnTimeout("+ anuncio +")", 1000); ++slideCounter;
    	document.getElementById('time').innerHTML = slideCounter + '/' + slideTimeout;
    	if (slideCounter == slideTimeout) {if (++currentSlide > imgsMax[anuncio]) currentSlide = 1; gal_paint(anuncio, currentSlide); slideCounter = 0;}
    }
    
    /* xGetURLArguments and xPad are part of the X library,
    distributed under the terms of the GNU LGPL,
    and maintained at Cross-Browser.com.
    */
    function xGetURLArguments() {
    	var idx = location.href.indexOf('?');
    	var params = new Array();
    	if (idx != -1) {
		    var pairs = location.href.substring(idx+1, location.href.length).split('&');
		    for (var i=0; i<pairs.length; i++) {nameVal = pairs[i].split('='); params[i] = nameVal[1]; params[nameVal[0]] = nameVal[1];}
	    }
	    return params;
    }
    
    function xPad(str, finalLen, padChar, left) {
    	if (typeof str != 'string') str = str + '';
    	if (left) {for (var i=str.length; i<finalLen; ++i) str = padChar + str;} else {for (var i=str.length; i<finalLen; ++i) str += padChar;}
    	return str;
    }
    </script>
</head>
<body> 
<div class=container id=idPage>
    <div class='row slide' id=slideshow>
        <nav id=navigation>
            <span id=prev>&#171; Previous</span>  <span id=next>Next &#187;</span> <span id=back>Back to the Gallery</span> <span id=auto title='Toggle Auto-Slide'>Auto-Slide</span> <span id=time>&nbsp;</span></nav>
        <div id=scon1 class=scon><img src=x id='s1'/><div class='scap' id='sc1'>&nbsp;</div></div></div>
    <div id=idHeader>
<%  if Request("msg")<>"" then %>
        <p><%=Request("msg")%></p>
<%  end if %>
        <div class=row>
            <div class=left><a href=PeDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
		    <div class=right>
<%  if Request("tipo")="alquiler" or Request("tipo")="venta" or Request("tipo")="inquilino" or Request("tipo")="comprador" then %>
    	        <a href=QuVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a>
<%	else %>
			    <a href=TrVacas.asp class=linkutils><%=MesgS("Anuncios Vacaciones","Holiday Adverts")%> &gt; </a>
<%	
    end if
	if Request("tipo")="alquiler" then 
%>
			    <a href=QuViviendaBuscoFront.asp class=linkutils></a>
<%	elseif Request("tipo")="vacas" then %>
			    <a href=TrVacasBuscoFront.asp class=linkutils></a>
<%	end if  %>
                <a href=QuViviendaBusco.asp>QuViviendaBusco.asp</a>
<%	if Request("tipo")="alquiler" then %>
			    <%=MesgS("Pisos y Habitaciones en ALQUILER","Flats and Rooms for RENT")%> 
<%	elseif Request("tipo")="venta" then %>
			    <%=MesgS("Pisos y Casas en VENTA","Flats and Houses for SALE")%> 
<%	elseif Request("tipo")="vacas" then %>
			    <%=MesgS("Apartamentos y Hoteles","Flats and Hotels for VACATIONS")%>
<%	elseif Request("tipo")="vacasswap" then %>
			    <% Response.Write MesgS("Apartamentos y Hoteles para INTERCAMBIO","Flats and Hotels for SWAPPING")%>
<%	elseif Request("tipo")="inquilino" then %>
			    <%=MesgS("Ofertas de INQUILINOS","People looking for Flats and Houses for RENT")%>
<%	end if  %>
                </div>
<%	if Request("tipo")="comprador" then %>
	        <h1>
    		    <a href=mailto:atencioncliente@domoh.com>atencioncliente@domoh.com</a>
<%	    if numPisos=0 then %>
                <%=MesgS("Pulsa en 'Publicar Piso o Casa' para registrar tu anuncio","Go to 'Post New House for Sale' for selling your property")%>.
<%		
		    Response.End
	    end if 
%>
                 </h1></div>
<%	end if %>
        </div></div>
<div id=idColumns>
    <div id=idSideColumn>
        <a href='javascript:xImgGallery(1);'>Cargar</a>
	    <input type=checkbox checked id=trigger1 /> Ver Todos
<%		
	dim sTramoPrecio, sKeyPrecio, colCiudad, colTipo, colPrecio
	set colCiudad = Server.CreateObject("Scripting.Dictionary")
	set colTipo = Server.CreateObject("Scripting.Dictionary")
	set colPrecio = Server.CreateObject("Scripting.Dictionary")
	rst.Open sQuery, sProvider

	for i=1 to numPisos
		if rst.Eof then exit for
		if not(colCiudad.Exists(UCase(Trim(rst("ciudadnombre"))))) then colCiudad.Add UCase(Trim(rst("ciudadnombre"))), UCase(Trim(rst("ciudadnombre")))
		if not(colTipo.Exists(UCase(Trim(rst("pTipo"))))) then colTipo.Add UCase(Trim(rst("pTipo"))), UCase(Trim(rst("pTipo")))
		if rst("pPrecio") < 300 then 
			sTramoPrecio="Menos que 300€"
			sKeyPrecio="<300"
		elseif rst("pPrecio") < 400 then 
			sTramoPrecio="300€-400€"
			sKeyPrecio="300-400"
		else
			sTramoPrecio="Más que 400€"
			sKeyPrecio=">400"
		end if			
		if not(colPrecio.Exists(sTramoPrecio)) then colPrecio.Add sTramoPrecio, sKeyPrecio
		rst.Movenext
	next
	rst.Close

	dim sItem
	i=2
	for each sItem in colCiudad.Items
		Response.Write "<input id=trigger" & i & " name=c" & sItem &" checked type=checkbox> " & sItem
		i=i+1
	next
	Response.Write sItem
	for each sItem in colTipo
		Response.Write "<input id=trigger" & i & " name=t" & sItem &" checked type=checkbox> " & sItem
		i=i+1
	next
	Response.Write "<p>" & sItem & "</p>"

	j=0
	for each sItem in colPrecio
			if sItem= "Menos que 300€" then 
				sKeyPrecio="<300"
			elseif sItem="300€-400€" then
				sKeyPrecio="300-400"
			else
				sKeyPrecio=">400"
			end if			
		Response.Write "<br><input id=trigger" & i & " name='p" & sKeyPrecio &"' checked type=checkbox>" & sItem
		i=i+1
		j=j+1
	next
%>
    </div> <!-- end idSideColumn -->
    <div id=idMainColumn>
<%		
	rst.Open sQuery, sProvider

	for i=1 to numPisos
	    if rst.Eof then exit for
		if rst("destacado") then sClass="anunciosDest" else sClass="anuncios2"
		if rst("pPrecio") < 300 then 
			sKeyPrecio="<300"
		elseif rst("pPrecio") < 400 then 
			sKeyPrecio="300-400"
		else
			sKeyPrecio=">400"
		end if			
%>
	<div id='anuncio<%=i%>' class=panel>
	    <input type=hidden id='anunciotag<%=i%>' value='c=<%=UCase(rst("ciudadnombre"))%>&t=<%=UCase(rst("pTipo"))%>&p=<%=sKeyPrecio%>&' />
<%
	    dim numVisitas, sBody, sQueryDet
	    on error goto 0

    	if sTabla="Pisos" then
	    	Response.Write CasaDetalles(Session("Idioma"), rst("aId"), "Inquilino", "", i)
	    elseif sTabla="Inquilinos" then
    		Response.Write InquilinoDetalles(Session("Idioma"), rst("aId"))
    	end if
%>
	</div>
<%
		rst.movenext
	next
	rst.Close
%>
    </div> <!-- end idMainColumn --></div> <!-- end idColumns -->
</body>
