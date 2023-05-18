<% 
    dim blk, loc
    uss=Request("user")
    loc=Request("locale")
    op=Request("Op")
    fn=Request("first name")
    ln=Request("last name")    

    if op="Oneliner" then
%>
{
 "messages": [
   {"text": "<%=fn%>, Vas al 85% de tu objetivo mensual. Te faltan 36k por ganar este mes, y 20 dias"},
   {"text": "Llegaras al objetivo con un 99% de probabilidad."},
   {"text": "Vas al 56% en Engagement de Clientes."},
   {"text": "Tu pipeline tiene un avance del 75%."},
   {"text": "Tus clientes tienen satisfaccion del 45%."},
   {"text": "Tu consumo de capital es adecuado. (Source: Dummy)"}
 ]
}
<p>
<% else %>
    <%=uss%> %>&#39;s Dashboard</p>
<p>
    Appointments</p>
<p>
    Main deals in Pipeline</p>
<p>
    Main Clients</p>
<p>
    Recommendations</p>
<p>
    &nbsp;</p>
<% end if %>