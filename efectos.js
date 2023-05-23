//--- Repone la imagen al salir el raton
function ImgRestore() {	document.imgs[0].src=document.imgs[0].oSrc; }

//--- Busca un objeto por nombre
function FindObj(n, d) {
	var p, i, x;  
	
	if (!d) d=document; 
	
	if ((p=n.indexOf("?"))>0 && parent.frames.length) {d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
	
	if (!(x=d[n]) && d.all) 
		x=d.all[n]; 

	for (i=0; !x && i<d.forms.length; i++) 
		x=d.forms[i][n];

	for (i=0; !x && d.layers && i<d.layers.length; i++) 
		x=FindObj(n, d.layers[i].document); 

	return x;
}

//--- Cambia la imagen de un objeto
function ImgSwap(nombre, imagen) {
	var x; 

	document.imgs=new Array; 
	
	if ((x=FindObj(nombre)) != null) {
		document.imgs[0]=x; 
		if(!x.oSrc) x.oSrc=x.src; 
		x.src=imagen;
	}
}