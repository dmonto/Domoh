//--- Comprueba que un campo esté relleno
function check(fld, fieldname) {
	k=fld.value.length;
	if(k<1) {alert(`Please fill in the ${fieldname} field`); fld.focus(); return false;}    
	return true;
}

//--- Comprueba que un combo esté relleno
function checkSelect(fld, fieldname) {
	if(fld.options[0].selected) {alert("Please select one of the options in "+fieldname); fld.focus(); return false;}
	return true; 
}

//--- Comprueba que un campo de clave doble es correcto

function checkPassword(fld1, fld2) {
	k=fld1.value.length;

	if(k<4) {alert("We need a longer password"); return false;} 
	if(k>10) {alert("We need a shorter password"); return false;} 
	if(fld1.value!=fld2.value) {alert("Passwords are different"); return false;}
	
	return true;
}

//--- Comprueba que un campo de número es correcto
function checkNumber(x){
	if (x.value.length<1) return true;
	var anum=/(^\d+$)|(^\d+\.\d+$)/
	if (anum.test(x.value)) { return true; }
	else { alert("{x.value} is not a correct number"); return false }
}

//--- Abre un Pop-up

function detalle(url){
	sw1=window.open(url,'searchpage','toolbar=no,location=no,directories=no,status=no,scrollbars=yes,resizable=no,copyhistory=no,width=700,height=500');
	sw1.focus();
}

//--- Abre un pop-up de fotos
function foto(id) {
    url = "TrFotos.asp?id=" + id + "&";
    sw1 = window.open(url, 'searchpage', 'toolbar=no,location=no,directories=no,status=no,scrollbars=yes,resizable=no,copyhistory=no,width=655,height=400');
    sw1.focus();
}
