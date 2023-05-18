<!--#include file="IncNuBD.asp"-->
<!--#include file="IncTrCabecera.asp"-->
<body onload="window.parent.location.hash='top';">
<div class=container>
<%  if Request("head")="si" then %>
    <!--#include file="IncFrHead.asp"-->
<%  end if %>
    <div class="logo"><a title='<%=MesgS("Ir a la página principal","Go to Main Page")%>' href=PeMenu.asp class=linkutils>&lt;&lt; <%=MesgS("Volver a Home","Back to Home Page")%> </a></div>
    <div class=banner><h1><%=MesgS("Conócenos","About Us")%></h1></div>
<%  if Request("idioma")="" or Request("idioma")="Es" then %>
    <div class=main>
        La mejor manera para conocernos, pensamos, es a través de las opiniones y comentarios que nos han enviado nuestr@s usuari@s. Aquí van algunos como ejemplo:
        <h3>
            !Hola! Me llamo Váleri y quería deciros que entrar en vuestra página es un placer. Me encanta el lay-out y la idea detrás de la página en sí. Ya he encontrado casa 
            y os escribo para pediros que expandáis los servicios y las posibilidades de la web, porque me gustaría seguir visitándola. Me alegro mucho de vuestra
            iniciativa y espero vaya a más. Un saludo muy fuertote!</h3>
        <h4>Váleri</h4>
        <p>
            Al principio tenía un poco de miedo, yo vivía sólo y estaba muy cómodo La verdad es que hemos sido muy independientes el uno con el otro... Supongo que cuando se vaya 
            volverá a repetir porque compartir la casa con alguien gay como yo, ha hecho enriquecerme y aprender aún más.</p>
        <p>Pablo.</p>
        <p>Me registré como usuario anteriormente para curiosear en busca de habitación en Barcelona, y vuestra web me pareció útil. Ahora quería alquilar una habitación en mi casa de Madrid</p>
        <p>Antonio.</p>
        <p>Soy un ingeniero francés y mi novio es periodista de la Vanguardia. Hemos encontrado vuestro sitio y os felicitamos. Estamos buscando un piso en Barcelona, más o menos céntrico, lo más grande</p>
        <p>Thomas.</p>
        <p>Primero muchas gracias por vuestra página, y el servicio que nos prestáis!!! ... Gracias de nuevo por todo y enhorabuena por vuestra página y por su diseño...bicos desde Galicia!!!.</p>
        <p>Rocío.</p>
        <p>Hola, antes de nada daros la enhorabuena por la idea tan genial de hacer una página como esta. Acabo de registrarme y me parece que está muy bien hecha</p>
        <p>Raúl.</p>
        <p>Gracias por crear una pagina para poder ayudarnos a los más jóvenes a encontrar un lugar donde poder vivir tranquilamente...</p>
        <p>Eglarüthiel</p>
        <p>Me acabo de registrar, y creo que la web es absolutamente útil. Creo que la voy a recomendar más de una vez</p>
        <p>Antonio.</p>
        <p>Estos son sólo una muestra de lo que piensan nuestr@s usuarios. Pero si aún así tienes dudas de qué es lo que hacemos y para qué, te lo explicamos muy rapidito:</p>
        <p>Por qué complicarnos viviendo en un sitio en el que no podamos ser nosotr@s mism@s con gente que no nos gusta o que sencillamente no es lo que queremos.</p>
        <p>Si podemos elegirlo TODO ¿por qué no hacerlo?</p>
        <p>Todos los que entienden de verdad y los que no, no lo hacen de otra manera: ELIGEN.</p>
        <p>Eligen por cuánto quieren vivir: un piso enorme de 5 habitaciones para tí solit@ carííísimo, o una habitación en una casa compartida baratííísima o algo intermeeedio.</p>
        <p>Eligen dónde quieren vivir: en el centro, no tan en el centro, un poco más de no tan en el centro o me gusta el campo.</p>
        <p>Eligen cómo quieren vivir: eligen si lo quieren hacer sól@s, con sus parejas, con sus amig@s, con sus os@s, con sus boll@s, con sus perr@s, y en cualquier caso con chic@s respetuos@s con tod@s.</p>
        <p>
            En domoh.com encontrarás tu casa y tus vacaciones mucho más fácilmente, más barato y por supuesto sin ningún problema sea cual sea tu orientación sexual: te lo garantizamos. 
            Nuestros ofertantes son personas que quieren alquilar sus casas a personas como nosotr@s: no nos importa tu orientación sexual. Es asunto tuyo.</p>
        <p>Elige en qué lugar quieres pasar tus vacaciones y encuentra pisos y casas compartidas a tu alcance: para un fin de semana, semana o lo que te apetezca sin gastarte lo que no te apetece.</p>
        <p>Entra y mira de forma gratuita las casas, pisos o habitaciones en casa compartidas que están disponibles: !Date prisa porque la cosa está que arde!</p>
        <p>Visita nuestras <a title='Ver Links Recomendados' href=NuRecomendamos.asp>recomendaciones</a>. Siempre hay algo nuevo que no habías visto antes sin lo que no podrás vivir a partir de ahora.</p>
        <p>Si tienes algún amig@ que está buscando un sitio para vivir háblale de domoh, seguro que te lo agradecerá.</p>
        <p>En domoh queremos llegar a todas las ciudades lo antes posible y estamos trabajando duro para conseguirlo. Visita periódicamente nuestra página y pronto verás como ya hay ofertas también en tu ciudad.</p>
        <p>Si tienes poco dinero y quieres independizarte y tienes amig@s que están igual que tú, en domoh.com encontrarás la solución: comparte un piso con ell@s a un precio más que razonable.</p>		
        <p>Si tienes un piso/habitación para alquilar o conoces a alguien que quiera alquilar un piso/habitación en tu ciudad dale nuestra dirección y le ayudarás a ella/él y a nosotros.</p>
        <p>No dudes en visitar domoh.com tantas veces como necesites.</p>
        <p>
            No dudes en contarnos qué cambiarías de nuestra página en <a title='Mandar mail de sugerencias' href='mailto:sugerencias@domoh.com'>sugerencias@domoh.com</a>: 
            seguro que tienes buenas ideas que nos ayudarán a hacer mejor nuestro trabajo. Y si quieres hacer un comentario sobre nosotr@s, el tiempo, 
            el finde pasado... tienes nuestro correo disponible para ti: <a title='Mandar mail de comentarios' href='mailto:comentarios@domoh.com'>comentarios@domoh.com</a></p>
        <p>Vuelve pronto y recuerda que estamos empezando. Tendremos novedades que te gustarán.</p></div>
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
