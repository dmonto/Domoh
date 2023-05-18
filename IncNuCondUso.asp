<div class=aviso>
	<p>
        <% Response.Write MesgS("Al publicar este anuncio estás aceptando estas", "Posting your advert means that you agree to our")%>
	    <a title='<%=MesgS("Leer las Condiciones de Uso","Read Terms & Conditions")%>' href='CondUso<%=Session("Idioma")%>.htm' target=_blank>
	    <%=MesgS("Condiciones Generales de Uso","Terms and Conditions")%></a> <%=MesgS("y nuestra ","and our ")%>
	    <a title='<%=MesgS("Leer la Política de Protección de Datos","Read Privacy Policy")%>' href='ProtDatos<%=Session("Idioma")%>.asp' target=_blank>
	    <%=MesgS("Política de Protección de Datos","Privacy and Safety policy")%></a></p></div>
<div class=seccion><p>* <%=MesgS("son campos obligatorios","are required fields")%></p></div>