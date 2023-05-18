<% 
    dim blk, loc
    blk=Request("last visited block name")
    loc=Request("locale")
    op=Request("op")
    cli=Request("client")
    uss=Request("chatfuel user id")
    
  if op="Oneliner" then 
%>    
{
 "messages": [
   {"text": "Ok let's see <%=cli %>"},
   {"text": "The most remarkable thing is...."}
 ]
}
<%  elseif cli<>"" then %>
<h1><%=cli%></h1>
<p>
    Estructura</p>
<p>
    Noticias</p>
<p>
    Ultimas Interacciones</p>
<p>
    Wallet y Oportunidades</p>
<p>
    Indicadores</p>
<%  end if %>
