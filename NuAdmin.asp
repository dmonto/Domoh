<!--#include file="IncNuBD.asp" -->
<!--#include file="IncQuCabecera.asp"-->
<div class=container>
	<div class=main>
		<p>Buenas Carretier. Este es tu Panel de Control. <a title='Termina Sesión' href=NuLogOff.asp>Salir</a></p>
		<h1>Pisos</h1>
		<h2><a title='Pantalla de publicar pisos' href='NuAdminPisos.asp?tipo=nuevos' target=pisos>Ver nuevos Pisos</a></h2>
		<h2><a title='Búsqueda de un piso concreto' href='TrAdminAnunciosGrid.asp?tipo=menu target=pisos'>Ver todos los Anuncios</a></h2>
		<h2><a title='Limpiar Pisos' href=TrAdminPisosViejos.asp target=pisos>Dar de baja Pisos Viejos</a></h2>
		<h2><a title='Limpiar Pisos Viejos' href=NuAdminPisosInactivos.asp target=pisos>Borrar Pisos Inactivos</a></h2>
		<h2><a title='Publicar anuncios de Inquilinos' href='NuAdminInquilinos.asp?tipo=nuevos' target=pisos>Ver Inquilinos</a></h2>
		<h1>Usuarios</h1>
		<h2><a title='Buscar un usuario' href='NuAdminUsuarios.asp?tipo=menu' target=usuarios>Ver Usuarios</a></h2>
		<h2><a title='Ver los que aun no han confirmado' href='NuAdminUsuarios.asp?tipo=nuevos' target=usuarios>Ver nuevos Usuarios</a></h2>
		<h2><a title='Ver de donde vienen los nuevos usuarios' href=NuAdminFuentes.asp target=usuarios>¿Cómo nos conocieron?</a></h2>
		<h2><a title='Nueva home de usuario' href='QuHomeUsuario.asp?usuario=hector' target=usuarios>Pruebas de 'Mi Domoh'</a></h2>
		<h1>Estadísticas</h1>
		<h2><a title='Abrir Analytics para ver Estadisticas' href='https://www.google.com/analytics/reporting/?reset=1&id=7351112' target=stats>Estadísticas</a></h2>
		<h2><a title='Estadisticas de Pisos Vistos por Dia' href=NuAdminAnunciosVistos.asp target=stats>Pisos Vistos por Día</a></h2>
		<h2><a title='Estadisticas de Disponibilidad' href=http://www.pingdom.com/reports/wzqhlazq78rx/ target=stats>Pingdom</a></h2>
		<h1>Contenidos</h1>
		<h2><a title='Actualizar artículos, banners...' href=NuAdminContenidos.asp target=conts>Contenidos de la Web</a></h2>
		<h1>Tablas</h1>
		<h2><a title='Añadir o cambiar provincias' href='NuAdminTabla.asp?tabla=Provincias&nombretabla=Provincias&auto=no' target=tablas>Provincias</a></h2>
		<h2><a title='Añadir o cambiar paises' href='NuAdminTabla.asp?tabla=Paises&nombretabla=Paises' target=tablas>Paises</a></h2>
		<h1>Otros</h1>
		<h2><a title='Datos de Usuario' href='TrUsuario.asp?usuario=hector' target=tablas>Cambiar Password etc</a><a title='Cuidaditorrrl!!' href=NuAdminBD.asp target=bd>Acceso Directo a BD</a></h2></div>
	<div class=foot>
		<!-- #include file="IncPie.asp" --></div></div>
