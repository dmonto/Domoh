<!-- #include file="IncNuBD.asp" -->
<!-- #include file="IncTrCabecera.asp" -->
<%	
	dim numFot, i
	
	rst.Open "SELECT COUNT(*) AS numFot FROM Fotos WHERE piso=" & Request("id"), sProvider 
	numFot=CLng(rst("numFot"))
	rst.Close	
    if numFot=0 then
        Response.Write "<h1>No hay fotos</h1>"
        Response.End
    end if
	rst.Open "SELECT * FROM Fotos WHERE piso=" & Request("id"), sProvider 
%>
<script type=text/javascript>
	pos=1;

	function foco(){
	if (pos==0) pos=<%=numFot%>; if (pos><%=numFot%>) pos=1; ImgSwap('fototop',FindObj('foto'+pos).src); cuad=FindObj('imgazul');
	cuad.parentNode.style.top=FindObj('foto'+pos).offsetParent.offsetParent.offsetTop+FindObj('foto'+pos).offsetParent.offsetTop+FindObj('foto'+pos).offsetTop-4;
	cuad.parentNode.style.left=
	FindObj('foto'+pos).offsetParent.offsetParent.offsetLeft+FindObj('foto'+pos).offsetParent.offsetLeft+FindObj('foto'+pos).offsetLeft-2;
	FindObj('imgazul').height=FindObj('foto'+pos).offsetHeight;
	FindObj('imgazul').width=FindObj('foto'+pos).offsetWidth-2;
}
</script>
<body onload='foco();'>
<a title='Foto Anterior' onclick='pos=pos-1;foco();'><img src=images/FlechaPrevio.gif /></a> <a title='Foto Siguiente' onclick='pos=pos+1;foco();'><img src=images/FlechaProx.gif /></a>
<div>
    <div class="fotogran"><img alt='Foto' title='Foto de la Vivienda' src="http://domoh.com/<%=rst("foto")%>" /></div>
    <div>
        <span title='Foto Ampliada' id=cuadrado class="absoluta"><img alt='Blanco' id=imgazul class=absoluta src=images/spacer.gif /></span>
<%
	i=1
	do while not rst.Eof
%>
        <img id=foto<%=i%> title='Foto de la Vivienda' src="http://domoh.com/<%=rst("foto")%>" onmouseover="pos=<%=i%>;foco();"/>
<%		
		if (i mod 3)=0 then Response.Write "</div><div>"
		i=i+1
		rst.Movenext
	loop
%>
</div></div><!-- #include file="IncTrGoogleAn.asp" --></body>
