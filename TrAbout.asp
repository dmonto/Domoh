<!--#include file="IncNuBD.asp"-->
<!--#include file="IncTrCabecera.asp"-->
<body onload="window.parent.location.hash='top';">
<div class=container>
<%  if Request("head")="si" then %>
    <!--#include file="IncFrHead.asp"-->
<%  end if %>
    <div class="logo"><a title='<%=MesgS("Ir a la p�gina principal","Go to Main Page")%>' href=PeMenu.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
    <div class=banner><h1><%=MesgS("Con�cenos","About Us")%></h1></div>
<%  if Request("idioma")="" or Request("idioma")="Es" then %>
    <div class=main>
        La mejor manera para conocernos, pensamos, es a trav�s de las opiniones y comentarios que nos han enviado nuestr@s usuari@s. Aqu� van algunos como ejemplo:
        <h3>
            !Hola! Me llamo V�leri y quer�a deciros que entrar en vuestra p�gina es un placer. Me encanta el lay-out y la idea detr�s de la p�gina en s�. Ya he encontrado casa 
            y os escribo para pediros que expand�is los servicios y las posibilidades de la web, porque me gustar�a seguir visit�ndola. Me alegro mucho de vuestra
            iniciativa y espero vaya a m�s. Un saludo muy fuertote!</h3>
        <h4>V�leri</h4>
        <p>
            Al principio ten�a un poco de miedo, yo viv�a s�lo y estaba muy c�modo La verdad es que hemos sido muy independientes el uno con el otro... Supongo que cuando se vaya 
            volver� a repetir porque compartir la casa con alguien gay como yo, ha hecho enriquecerme y aprender a�n m�s.</p>
        <p>Pablo.</p>
        <p>Me registr� como usuario anteriormente para curiosear en busca de habitaci�n en Barcelona, y vuestra web me pareci� �til. Ahora quer�a alquilar una habitaci�n en mi casa de Madrid</p>
        <p>Antonio.</p>
        <p>Soy un ingeniero franc�s y mi novio es periodista de la Vanguardia. Hemos encontrado vuestro sitio y os felicitamos. Estamos buscando un piso en Barcelona, m�s o menos c�ntrico, lo m�s grande</p>
        <p>Thomas.</p>
        <p>Primero muchas gracias por vuestra p�gina, y el servicio que nos prest�is!!! ... Gracias de nuevo por todo y enhorabuena por vuestra p�gina y por su dise�o...bicos desde Galicia!!!.</p>
        <p>Roc�o.</p>
        <p>Hola, antes de nada daros la enhorabuena por la idea tan genial de hacer una p�gina como esta. Acabo de registrarme y me parece que est� muy bien hecha</p>
        <p>Ra�l.</p>
        <p>Gracias por crear una pagina para poder ayudarnos a los m�s j�venes a encontrar un lugar donde poder vivir tranquilamente...</p>
        <p>Eglar�thiel</p>
        <p>Me acabo de registrar, y creo que la web es absolutamente �til. Creo que la voy a recomendar m�s de una vez</p>
        <p>Antonio.</p>
        <p>Estos son s�lo una muestra de lo que piensan nuestr@s usuarios. Pero si a�n as� tienes dudas de qu� es lo que hacemos y para qu�, te lo explicamos muy rapidito:</p>
        <p>Por qu� complicarnos viviendo en un sitio en el que no podamos ser nosotr@s mism@s con gente que no nos gusta o que sencillamente no es lo que queremos.</p>
        <p>Si podemos elegirlo TODO �por qu� no hacerlo?</p>
        <p>Todos los que entienden de verdad y los que no, no lo hacen de otra manera: ELIGEN.</p>
        <p>Eligen por cu�nto quieren vivir: un piso enorme de 5 habitaciones para t� solit@ car���simo, o una habitaci�n en una casa compartida barat���sima o algo intermeeedio.</p>
        <p>Eligen d�nde quieren vivir: en el centro, no tan en el centro, un poco m�s de no tan en el centro o me gusta el campo.</p>
        <p>Eligen c�mo quieren vivir: eligen si lo quieren hacer s�l@s, con sus parejas, con sus amig@s, con sus os@s, con sus boll@s, con sus perr@s, y en cualquier caso con chic@s respetuos@s con tod@s.</p>
        <p>
            En domoh.com encontrar�s tu casa y tus vacaciones mucho m�s f�cilmente, m�s barato y por supuesto sin ning�n problema sea cual sea tu orientaci�n sexual: te lo garantizamos. 
            Nuestros ofertantes son personas que quieren alquilar sus casas a personas como nosotr@s: no nos importa tu orientaci�n sexual. Es asunto tuyo.</p>
        <p>Elige en qu� lugar quieres pasar tus vacaciones y encuentra pisos y casas compartidas a tu alcance: para un fin de semana, semana o lo que te apetezca sin gastarte lo que no te apetece.</p>
        <p>Entra y mira de forma gratuita las casas, pisos o habitaciones en casa compartidas que est�n disponibles: !Date prisa porque la cosa est� que arde!</p>
        <p>Visita nuestras <a title='Ver Links Recomendados' href=NuRecomendamos.asp>recomendaciones</a>. Siempre hay algo nuevo que no hab�as visto antes sin lo que no podr�s vivir a partir de ahora.</p>
        <p>Si tienes alg�n amig@ que est� buscando un sitio para vivir h�blale de domoh, seguro que te lo agradecer�.</p>
        <p>En domoh queremos llegar a todas las ciudades lo antes posible y estamos trabajando duro para conseguirlo. Visita peri�dicamente nuestra p�gina y pronto ver�s como ya hay ofertas tambi�n en tu ciudad.</p>
        <p>Si tienes poco dinero y quieres independizarte y tienes amig@s que est�n igual que t�, en domoh.com encontrar�s la soluci�n: comparte un piso con ell@s a un precio m�s que razonable.</p>		
        <p>Si tienes un piso/habitaci�n para alquilar o conoces a alguien que quiera alquilar un piso/habitaci�n en tu ciudad dale nuestra direcci�n y le ayudar�s a ella/�l y a nosotros.</p>
        <p>No dudes en visitar domoh.com tantas veces como necesites.</p>
        <p>
            No dudes en contarnos qu� cambiar�as de nuestra p�gina en <a title='Mandar mail de sugerencias' href='mailto:sugerencias@domoh.com'>sugerencias@domoh.com</a>: 
            seguro que tienes buenas ideas que nos ayudar�n a hacer mejor nuestro trabajo. Y si quieres hacer un comentario sobre nosotr@s, el tiempo, 
            el finde pasado... tienes nuestro correo disponible para ti: <a title='Mandar mail de comentarios' href='mailto:comentarios@domoh.com'>comentarios@domoh.com</a></p>
        <p>Vuelve pronto y recuerda que estamos empezando. Tendremos novedades que te gustar�n.</p></div>
<%  else %>
    <div class=main>
        <p>The access to the international community is basic for us. The improvement and implementation in the contents and services for English-speaking users is very important in domoh.com.</p>
        <p>
            The foreign visitors will find in domoh.com their lodging while visiting Spain and Europe. Besides somewhere to stay they will find what Spanish and Europeans like to do on their 
            leisure time: music, beauty, culture, discos, bars, restaurants</p>
        <p><a title='Home Page' href=http://www.domoh.com/ >www.domoh.com</a></p>
        <p>he website number 1 in gay friendly lodgings and holidays in Spain and Europe. New contents now!</p>
        <p>1.- Create and maintain an open, tolerant and respectful place with everyone.</p>
        <p>2.- To offer a quick and easy website.</p>
        <p>3.- Quality information = quality service</p>
        <p>
            My home This section is for those who are looking for their residence either in a flat on their own or in a shared flat.You can advertise a 
            flat for rent or for sale.</p>
        <p>
            My holidays Our users can choose where to go, how, who with and how much money they will spend. From rooms in a shared flat to resorts, luxurious hotels, hostels 
            And, of course, the swapping of flats, apartments and villas for holidays. </p>
        <p>My job Job positions for everyone. It is important the no discrimination within our private life, as well as the no discrimination in our professional environment. </p>
        <p>
            Miscellaneous Every sort of advert will be published in this section: sale, rent, exchange of cars, computers, musical instruments, parking
            The section will be defined by our users.</p>
        <p>Health & beauty - (new)</p>
        <p>Domoh.com will include a place where the latest in Health and Beauty will play the main roll. Information very demanded by our public.</p>
        <p>Fashion & Trends - (new)</p>
        <p>The trendiest, the latest trends in fashion and culture will be here. Avant-garde. Pioneering. Opened to all the creators who have something new to say, to feel. </p>
        <p>Music & Entertainment - (new)</p>
        <p>
            The latest albums, concerts, artists. Theatre, cinema, spectacles, bars, pubs, restaurants. In fact, everything that we can do on our leisure time will be @domoh.com.
            The restaurants, pubs and bars more exclusive will have a very special place @domoh.com.</p></div>
<%  end if %>
<!-- #include file="IncPie.asp" --></div></body>
