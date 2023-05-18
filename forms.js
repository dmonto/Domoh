//--- Comprueba que un campo esté relleno
function check(fld, fieldname) {
	k=fld.value.length;
	if(k<1) {alert("Por favor rellena el campo "+fieldname); fld.focus(); return false;} 
	return true;
}

//--- Comprueba que un combo esté relleno
function checkSelect(fld, fieldname) {
	if(fld.options[0].selected) {alert("Por favor selecciona "+fieldname); fld.focus(); return false;}
	return true; 
}

//--- Comprueba dos campos de password
function checkPassword(fld1, fld2) {
	k=fld1.value.length;

	if(k<1) {alert("Escribe una clave por favor"); return false;} 
	if(k<4) {alert("Escribe una clave más larga"); return false;} 
	if(k>10) {alert("Escribe una clave más corta"); return false;} 
	if(fld1.value!=fld2.value) {alert("Las claves deben ser iguales"); return false;}
	
	return true;
}

//--- Comprueba que se aceptan las condiciones de uso
function checkCond(fld, fieldname) {
	if(!fld.checked) {alert("Debes leer y aceptar las "+fieldname); fld.focus(); return false;}
	return true; 
}

//--- Comprueba que un número es correcto
function checkNumber(x){
	if (x.value.length<1) return true;
	var anum=/(^\d+$)|(^\d+\.\d+$)/
	if (anum.test(x.value)) { return true; } else { alert(x.value + " no es un número correcto"); return false; }
}

//--- Abre una ventana de pop-up
function detalle(url){
	sw1=window.open(url,'searchpage','toolbar=no,location=no,directories=no,status=no,scrollbars=yes,resizable=no,copyhistory=no,width=700,height=550');
	sw1.focus();
}

//--- Abre un pop-up de fotos
function foto(id) {
    url = "TrFotos.asp?id=" + id + "&";
    sw1 = window.open(url, 'searchpage', 'toolbar=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,copyhistory=no,width=655,height=500');
    sw1.focus();
}