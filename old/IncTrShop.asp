<%
	dim sNombreTienda, sColorBG, sColorText, sColorDark, sColorLight, sColorTLight, sColorTDark, sColorLink, sMoneda, sFont

	sMoneda="€"
	sNombreTienda="EShop"
	sFont="arial"

	sColorBG="#FFFFFF"
	sColorText="#000000"
	sColorLink="#999999"

	sColorLight="#006bce"
	sColorDark="#003366"
	sColorTLight="#FFFFFF"
	sColorTDark="#FFFFFF"

	Function valid_sql(s)
		dim i, temp

		for i = 1 To Len(s)
			if Mid(s, i, 1) = "'" then temp = temp + "'"
			temp = temp + Mid(s, i, 1)
		next
		valid_sql = Trim(temp) 
	End Function

	Sub Header
%>
<div><img alt='Shopping' src=images/TrSecShopping.gif /><img alt='Blanco' src=images/spacer.gif /></div>
<%
	End Sub

    Sub categorymenu
	    dim showcart, numitems, i
	
	    showcart=false
	    '--- work out contents of shopping cart
	    numitems=0

	    if IsArray(Session("cart")) = false then
		    dim acart(19,1)
		    Session("cart") = acart
		    showcart=false
	    else
		    acart=Session("cart")
		    for i=LBound(acart) to UBound(acart)
		        if acart(i,0)<>"" and acart(i,1)<>"" then
			        numitems=numitems+acart(i,1)
			        showcart=true
		        end if
		    next
	    end if
%>
<div>
<%      if showcart then %>
	<div><%= numitems %>&nbsp;<%=MesgS("artículos en la cesta","items in shopping cart")%></div>
<%      end if %>
		
	<nav>
		<b>
		<a title='<%=MesgS("Volver a Domoh","Back to Domoh")%>' href=QuDomoh.asp ><%=MesgS("Inicio","Home")%></a> |
		<a title='<%=MesgS("Mandar mail a Soporte","Send e-mail to Support")%>' href=mailto:hector@domoh.com ><%=MesgS("Ayuda","Support")%></a>

<%      if showcart then %>
		| <a title='<%=MesgS("Ver Cesta","Review Cart")%>' href=old/TrShopRevisa.asp ><%=MesgS("Mi Cesta","My Cart")%></a>
<%  
        end if 
	
        '--- Check if user is  signed in
	    if Session("custid")<>""  then 
%>
		| <a href=signout.asp ><%=MesgS("Salir","Sign Out")%></a>
<%      else %>
		| <%=MesgS("Entrar","Sign In")%>
<%      end if  %>
        </b></nav></div>
<div>
<%  
        '--- Display list of categories
	    rst.Open "SELECT * FROM CatShop", sProvider

	    if not rst.Eof then
		    while not rst.Eof
%>
    <a href='TrShopCategoria.asp?catcode=<%= rst("id") %>' ><%=rst("nombre")%></a>
<%
		        rst.Movenext
		        if not rst.Eof then
%>
	&nbsp;|&nbsp;
<%
		        end if
		    wend
	    end if
	    rst.Close
%>
    </div>
<%  End Sub %>