<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncNuMail.asp" -->
<%
	dim vtipo, vdir1, vciudadnombre, vdir2, vcp, vzona, vrentaviv, vdescripcion, vgente, vfuma, vmascota, vfoto, vidiomaes, vidiomaen, sBody

	vfuma=" checked"
	vmascota=" checked"
	if Session("Idioma")="En" then 
		vidiomaes="" 
		vidiomaes=" checked" 
	else 
		vidiomaes=" checked"
		vidiomaen=""
	end if
	vtipo="Piso"
	vfoto=" checked"

	if Request("id")<>"" then
		if Request("op")="Borrar" then
		 	sQuery = "UPDATE Anuncios SET activo='No', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha desactivado el piso (" & rst("cabecera") &")"
			rst.Close
			Mail "hector@domoh.com", "Piso Desactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Ocultado.","House+Hidden.")
		elseif Request("op")="Delete" then
		 	sQuery = "DELETE FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "DELETE FROM Pisos WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha borrado el piso (" & Request("id") &")"
			Mail "hector@domoh.com", "Piso Borrado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Borrado.","House+Deleted.")
		elseif Request("op")="Activar" then
		 	sQuery = "UPDATE Anuncios SET activo='Si', fechaultimamodificacion=GETDATE() WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
		 	sQuery = "SELECT * FROM Anuncios WHERE id=" & Request("id")
			rst.Open sQuery, sProvider
			sBody=Session("Usuario") & " ha reactivado el piso (" & rst("cabecera") &")"
			rst.Close
			Mail "hector@domoh.com", "Piso Reactivado", sBody
			Session("Id")=""
			Response.Redirect "QuHomeUsuario" & Session("Idioma") & ".asp?msg=" & MesgS("Piso+Activado.","House+Republished.")
		elseif Request("op")="Modificar" then
			sQuery="SELECT * FROM Pisos p INNER JOIN Anuncios a ON p.id=a.id WHERE a.Id=" & Request("id")
			rst.Open sQuery, sProvider
			vtipo=rst("tipo")
			vdir1=rst("dir1")
			vdir2=rst("dir2")
			vcp=rst("cp")
			vciudadnombre=rst("ciudadnombre")
			vzona=rst("zona")
			if rst("rentaviv") > 0 then vrentaviv=rst("rentaviv")
			vdescripcion=rst("descripcion")
			vgente=rst("gente")
			if UCase(rst("fuma"))<>"ON" then vfuma=""
			if UCase(rst("mascota"))<>"ON" then vmascota=""
			if UCase(rst("idiomaes"))<>"ON" then vidiomaes=""
			if UCase(rst("idiomaen"))<>"ON" then vidiomaen="" else vidiomaen=" checked"
			if IsNull(rst("foto")) or rst("foto")="" then vfoto=""
			Session("Id")=Request("id")
		elseif Request("op")="Fotos" then
			Session("Id")=Request("id")
			Response.Redirect "TrCasaRegOferta.asp?op=FotoBorrada"
		end if
	end if
%>
<!-- #include file="IncTrCabecera.asp" -->
<style type=text/css>textarea { margin: 2em;}</style>
<script type=text/javascript>
    function checkForm() {
	    if(!check(document.frm.ciudadnombre, "<%=MesgS("Ciudad","City")%>")) return false;
	    if(!check(document.frm.zona, "<%=MesgS("Zona","Area")%>")) return false;
	    if(!checkNumber(document.frm.rentaviv, "<%=MesgS("Renta Mensual","Monthly Rent")%>")) return false;
	    if(!check(document.frm.descripcion, "<%=MesgS("Descripción","Description")%>")) return false;
	    if(document.frm.tipo[1].checked) {if (!check(document.frm.gente, "<%=MesgS("Cómo son las personas","People living there")%>")) return false};
	    if(!check(document.frm.nombre, "<%=MesgS("Nombre","Name")%>")) return false;
<%	if Request("id")="" then %>
	    if(!check(document.frm.email, "E-mail")) return false;
	    if(!check(document.frm.usuario, "<%=MesgS("Nombre de Usuario","Username")%>")) return false;
	    if(!checkSelect(document.frm.fuente, "<%=MesgS("Cómo nos Conociste","How did you get to us")%>")) return false;
	    if(!checkPassword(document.frm.password, document.frm.password2)) return false;
	    if ((document.frm.email.value.length && document.frm.mostraremail.checked) || (document.frm.tel1.value.length && document.frm.mostrartel1.checked) ||
		    (document.frm.tel2.value.length && document.frm.mostrartel2.checked) || (document.frm.tel3.value.length && document.frm.mostrartel3.checked) ||	(document.frm.tel4.value.length && document.frm.mostrartel4.checked))
	        return true;
	    alert("<%=MesgS("Debes activar alguna forma de contacto","You must check one way of contact")%>");
	    return false;
<%	else %>
	    return true;
<%	end if %>
    }
    window.onload = function() {taInit(false);}
    // Initialize all textareas.
    // If bCols is false then columns will not be resized.
    function taInit(bCols) {
    	var i, ta = document.getElementsByTagName('textarea');
    	for (i = 0; i < ta.length; ++i) {
		    ta[i]._ta_resize_cols_ = bCols;	ta[i]._ta_default_rows_ = ta[i].rows; ta[i]._ta_default_cols_ = ta[i].cols;
		    ta[i].onkeyup = taExpand; ta[i].onmouseover = taExpand; ta[i].onmouseout = taRestore; ta[i].onfocus = taOnFocus; ta[i].onblur = taOnBlur;
	    }
    }

    function taOnFocus(e) {this._ta_is_focused_ = true; this.onmouseover();}
    function taOnBlur() {this._ta_is_focused_ = false; this.onmouseout();}

    // Set to default size if not focused.
    function taRestore() {if (!this._ta_is_focused_) {this.rows = this._ta_default_rows_; if (this._ta_resize_cols_) {this.cols = this._ta_default_cols_;}}}

    // Resize rows and cols to fit text.
    function taExpand() {
        var a, i, r, c = 0;
        a = this.value.split('\n');
    
        if (this._ta_resize_cols_) {
            for (i = 0; i < a.length; i++) // find max line length
            {
                if (a[i].length > c) {c = a[i].length;}}
                if (c < this._ta_default_cols_) {c = this._ta_default_cols_;}
                this.cols = c; r = a.length;
            }
        else {
            for (i = 0; i < a.length; i++) {
                if (a[i].length > this.cols) // find number of wrapped lines
                {
                    c += Math.floor(a[i].length / this.cols);
                }
            }
            r = c + a.length; // add number of wrapped lines to number of lines
        }  
        
        if (r < this._ta_default_rows_) {r = this._ta_default_rows_;}
        this.rows = r;
    }
</script>
<body>
<div class=container>
    <form name=frm onsubmit="return checkForm();" action=TrCasaRegOferta.asp method=post>
        <input type=hidden name=id value=<%=Request("id")%> />
    <div class=logo><a href=QuDomoh.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
    <div class=topmenu><a href=TrVivienda.asp class=linkutils><%=MesgS("Anuncios Vivienda","Property Adverts")%> &gt; </a><% Response.Write MesgS("Publicar Nuevo Lugar para Alquilar", "Post New Advert for Rent")%></div>
	<div class=main>
    	<% Response.Write MesgS("Publicación de Piso o Habitación para Vivir", "Post New Advert for Rent")%></div>
    <div>
        </div>
	<div>* <%=MesgS("son campos obligatorios","are required fields")%></div>
    <div><% Response.Write MesgS("¿Cómo es la vivienda que quieres alquilar?","How is the property you want to rent?")%></div>
    <div>
        <input title='<%=MesgS("Pulsa aquí si alquilas todo el piso","Check if it is a complete apartment")%>' type=radio '<% if vtipo="Piso" then Response.Write " checked"%>' value='Piso' name=tipo />
        <%=MesgS("Es un piso completo ","Complete Apartment")%>
        <input title='<%=MesgS("Pulsa aquí es para compartir","Check if it is for sharing")%>' type=radio value='Habitación' <% if vtipo="Habitación" then Response.Write " checked"%> name=tipo />
        <%=MesgS("Es una habitación en casa compartida","Is a room in a shared apartment")%>
        <input title='<%=MesgS("Pulsa aquí si no es una vivienda","Check if not for living")%>' type=radio value='Local' <% if vtipo="Local" then Response.Write " checked"%> name=tipo />
        <%=MesgS("Es un local, garaje","It's commercial property, garage")%>...</div>
    <div><%=MesgS("¿Dónde está el piso/habitación?","Where is the property?")%></div>
    <div>
        <div>
            * <%=MesgS("Ciudad","City")%> <input title='<%=MesgS("Nombre de la Población","Name of the City/Town")%>' maxlength=50 size=19 name=ciudadnombre value="<%=vciudadnombre%>" />
            * <%=MesgS("Zona","Area")%> <input title='<%=MesgS("Barrio, Distrito...","Name of the Neighbourhood")%>' maxlength=45 size=30 name=zona value="<%=vzona%>"/>
            * <%=MesgS("Renta Mensual","Monthly Rent")%>: <input title='<%=MesgS("Renta Mensual en euros","Monthly Rent in Euros")%>' maxlength=30 size=8 name=rentaviv value='<%=vrentaviv%>'/></div>
        <div>
            <%=MesgS("Dirección ","Address ")%>(<%=MesgS("Cuanto más preciso seas, mejor será el mapa. Ej.: Calle Preciados 5","Please be accurate so we can produce a better map. For instance: Calle Preciados 5")%>)
            <input title='<%=MesgS("Calle, número, piso","Street, number, floor...")%>' maxlength=100 size=34 name=dir1 value="<%=vdir1%>"  autocomplete="address-level1" />
            <%=MesgS("Dirección (sigue)","Address (cont.)")%> <input maxlength=100 size=34 name=dir2 value="<%=vdir2%>" autocomplete="address-level2" /> <%=MesgS("Código Postal","Postal Code")%><input maxlength=20 size=10 name=cp value='<%=vcp%>'/></div></div>
        <div>
            <div>
                * <%=MesgS("Descríbenos la vivienda","Describe the Property")%>: <div><textarea name=descripcion rows=1 cols=36 wrap=soft><%=vdescripcion%></textarea></div>
                <input type=checkbox <%=vidiomaes%> value='ON' name=idiomaes /><%=MesgS("Texto en Español","Text in Spanish")%>
                <input type=checkbox <%=vidiomaen%> value='ON' name=idiomaen /><%=MesgS("Texto en Inglés ","Text in English")%>
                <input type=checkbox <%=vfoto%> name=foto /><%=MesgS("¿Quieres subir fotos de la casa?","Do you want to upload pictures?")%>
                <%=MesgS("¿No te llamarán si no pones fotos!","People won't call if you don't upload pictures!")%></div>
            <div>
                <%=MesgS("¿Cómo son las personas que viven en la casa?","How are the people that are living there?")%> * <%=MesgS("Obligatorio en Casas Compartidas ","Required for Shared Houses ")%>
                <div><textarea name=gente rows=1 cols=34 wrap=soft><%=vgente%></textarea></div>
                <input type=checkbox <%=vfuma%> value='ON' name=fuma /><%=MesgS("Admitís Fumadores ","Smokers Allowed ")%>
                <input type=checkbox <%=vmascota%> value='ON' name=mascota /><%=MesgS("Admitís Mascotas","Pets Allowed ")%></div></div>                            
<%	if Request("id")="" and Session("Usuario")="" then FormDatosPersonales %>
            <div>
                <input name=submit type=submit class=btnLogin id=submit 
<%  if Request("id")="" then %>
                    value='<%=MesgS("Publicar anuncio","Post Advert")%>' />
<%	elseif Request("id")="nuevo" then %>
                    value='Añadir piso'/>                                         
<%	else %>
                    value='Modificar piso'/>
<%	end if %>
                </div></form>
            <div id=banners_lat>
                <div><a href=http://www.communitas.es target=_blank><img src=conts/Communitas.gif /></a></div><div><a href=http://www.aulacreactiva.com/ target=_blank><img src=conts/Creactiva.gif /></a></div>
                <div><a href=http://www.materiagris.es target=_blank><img src=conts/MatGris.gif id=banner_4 /></a></div></div></div>
        <div>
            <a href=http://europegay.eu title='Web Europea' class=linkTxtSub>EuropeGay</a>  |  <a href=http://domoh.eu title='Web Europea' class=linkTxtSub> DomohEu</a>
            Unipersonal, C/Mayor, 12-3-A 28013 Madrid R.M. de Madrid Tº 20.536 Libro 0 Folio 48 Sección 8 Hoja M-363453 CIF: B-84096262</div>
</body>
