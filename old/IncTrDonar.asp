<div class=aviso>
	<div><h1><% Response.Write MesgS("Domoh es una web gratuita por lo que agradecemos tu ayuda para poder seguir dando servicio", "Domoh is a free service. We appreciate your economic support.")%></h1></div>
    <div>
        <h2><% Response.Write MesgS("Puedes apoyarnos simplemente enviando un SMS con el texto DOMOH al ", "To help us, please send a SMS with the text DOMOH to ")%>25588 </h2>
        <p>
            (<% Response.Write MesgS("Coste del SMS 1,42eur (IVA inc). Titular 25588","SMS Cost 1.42Eur (VAT Inc). Register #25588")%>: 
            Sit Consulting S.L. CIF B59585935. <% Response.Write MesgS("At.Cliente","Cust.Service")%>:902116106 info@sitmobile.com)</p>
		<h2><% Response.Write MesgS("También puedes usar tu tarjeta de crédito. Pulsa 'Donar' y sigue las instrucciones.", "You can also use your credit card. Press 'Donate' and follow instructions.")%></h2>
<form target=_new action=https://www.paypal.com/cgi-bin/webscr method=post>
	<input type=hidden name=cmd value=_s-xclick />
	<input type=hidden name=hosted_button_id value=9714016 />
	<input type=image src=https://www.paypal.com/es_ES/ES/i/btn/btn_donateCC_LG.gif name=submit alt='PayPal. La forma rápida y segura de pagar en Internet.' class=noborde/>
	<img src=https://www.paypal.com/es_ES/i/scr/pixel.gif />
</form>	
<form name=frm2 action=<%=Request("destino")%>?<%=Request("QUERY_STRING")%> method=post>
    <input type=hidden name=donar value=Si />
    <p><input name=Submit type=submit class=btnLogin value='<%=MesgS("Seguir y Ver Anuncios","Watch Adverts")%>'/></p></form>
</div></div>