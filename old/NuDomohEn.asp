<!--#include file="IncNuBD.asp"-->
<%	Session("Idioma")="En" %>
<!--#include file="IncTrCabecera.asp"-->
<body style = "background-image:url('images/NuDomohFondo4.jpg');background-repeat: no-repeat; background-attachment: fixed; background-position: -0px -200px" onload="window.parent.location.hash='top';">
<a title='Go to Miscellaneous' href=TrVarios.asp onmouseout="ImgRestore();document.getElementById('bolatexto').innerText='';"
    onmouseover="ImgSwap('Misc','images/MiscFoto.gif');	document.getElementById('bolatexto').innerText='Miscellaneous';">
    <img alt=Varios id=Misc src=images/MiscBolaEn.gif/></a>
<a title='Go to My Holidays' href=TrVacas.asp onmouseout="ImgRestore();document.getElementById('bolatexto').innerText='';"
	onmouseover="ImgSwap('Vacas','images/VacasFoto.gif');document.getElementById('bolatexto').innerText='My Holidays';">
	<img alt=Vacaciones id=Vacas src=images/VacasBolaEn.gif /></a>
<a title='Go to My Job' href=TrTrabajo.asp 
	onmouseout="ImgRestore();document.getElementById('bolatexto').innerText='';" onmouseover="ImgSwap('Curro','images/CurroFoto.gif'); document.getElementById('bolatexto').innerText='My Job';">
	<img alt='Trabajo' id=Curro src=images/CurroBolaEn.gif /></a>
<a title='Go to My Home' href=TrVivienda.asp onmouseout="ImgRestore();document.getElementById('bolatexto').innerText='';"
    onmouseover="ImgSwap('Casa','images/CasaFoto.gif');	document.getElementById('bolatexto').innerText='My Home';">
	<img alt='Casa' id=Casa src=images/CasaBolaEn.gif /></a>
<span id=bolatexto></span>
<div>
	<div>
<%  if Request("msg")<>"" then %>
		<p><%=Request("msg")%></p>
<%  end if %>
		</div>
<%	rst.Open "SELECT * FROM Contenidos WHERE pagina='ModaEn' and Posicion=0", sProvider %>
	<div><a title='Go to Health&Beauty' href='old/TrSeccion.asp?seccion=Moda'><img src=images/ModaEn.jpg /></a></div></div>
<div>
    <div><a title='Go to Health&Beauty' href='old/NuSeccion.asp?seccion=Moda'><img src="<%=rst("banner")%>" /></a></div>
    <div><p><%=rst("texto")%> <a title='Go to Health&Beauty' href='old/NuSeccion.asp?seccion=Moda'>More</a></p></div></div>
<%	
	rst.Close 
	rst.Open "SELECT * FROM Contenidos where Pagina='BellezaEn' and Posicion=0", sProvider 
%>
<div><a title='Go to Trends&Fashion' href='old/NuSeccion.asp?seccion=Belleza'><img src=images/BellezaEn.jpg /></a></div>
<div>
	<div><a title='Go to Trends&Fashion' href='old/NuSeccion.asp?seccion=Belleza'><img src='<%=rst("banner")%>' /></a></div>
	<div><p><%=rst("texto")%><a title='Go to Trends&Fashion' href='old/NuSeccion.asp?seccion=Belleza'>More</a></p></div></div>
<%	
	rst.Close
	rst.Open "SELECT * FROM Contenidos where Pagina='MusicaEn' and Posicion=0", sProvider 
%>
<div><a title='Go to Music' href='old/NuSeccion.asp?seccion=Musica'><img src=images/MusicaEn.jpg /></a></div>
<div>
	<div><a title='Go to Music' href='old/NuSeccion.asp?seccion=Musica'><img src='<%=rst("banner")%>' /></a></div=rst("texto")%> <a title='Go to Music' href='old/NuSeccion.asp?seccion=Musica'>More</a></p></div></div>
<!-- #include file="IncPie.asp" -->
</body>
