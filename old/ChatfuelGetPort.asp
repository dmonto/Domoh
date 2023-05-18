<% 
    dim blk, loc, uss, op
    uss=Request("chatfuel user id")
    loc=Request("locale")
    op=Request("op")

    if uss="" then 
        uss="User"
    elseif uss="1773283139382994" then
        uss="Diego"
    end if

    if op="NewTop" then
%>
<%
    elseif op="Top" then
%>
{
  "messages": [
    {
      "attachment": {
        "type": "template",
        "payload": {
          "template_type": "button",
          "text": "Choose one or type in",
          "buttons": [
            {
              "type": "show_block",
              "block_names": ["Todos"],
              "title": "All"
            },
            {
              "type": "show_block",
              "block_names": ["UnCliente"],
              "title": "Anglo American"
            },
            {
              "type": "show_block",
              "block_names": ["UnCliente"],
              "title": "Rolls Royce"
            }
          ]
        }
      }
    }
  ]
}    
<%  else %>
<head>
    <link rel=stylesheet type=text/css href=TrDomoh.css>
    <style type="text/css">
        .auto-style1 {
            width: 100%;
        }
        .auto-style2 {
            width: 165px;
        }
        .auto-style3 {
            width: 205px;
        }
        .auto-style4 {
            width: 162px;
        }
    </style>
</head>
<p>
    <%=uss%>&#39;s Portfolio</p>
<p>
    <br />
    <table class="auto-style1">
        <tr>
            <td class="auto-style2">350 Clients</td>
            <td>GBP 30.000 YTD</td>
        </tr>
        <tr>
            <td class="auto-style2">Pipeline Value</td>
            <td>GBP 50.000 </td>
        </tr>
        <tr>
            <td class="auto-style2">Last Deal: Rolls Royce</td>
            <td>GPB 10.000</td>
        </tr>
    </table>
</p>
<p>
    Top Clients</p>
<table class="auto-style1">
    <tr>
        <td class="auto-style2">Client</td>
        <td>Revenues YTD</td>
    </tr>
    <tr>
        <td class="auto-style2">Anglo American</td>
        <td>GBP 20.000</td>
    </tr>
    <tr>
        <td class="auto-style2">Rolls Royce</td>
        <td>GPB 10.000</td>
    </tr>
</table>
<p>
    Main deals in Pipeline</p>
<table class="auto-style1">
    <tr>
        <td class="auto-style2">Client</td>
        <td class="auto-style3">Product</td>
        <td class="auto-style4">Revenue</td>
        <td>Prob</td>
    </tr>
    <tr>
        <td class="auto-style2">Anglo American</td>
        <td class="auto-style3">Loan</td>
        <td class="auto-style4">60.000 GBP</td>
        <td>50%</td>
    </tr>
    <tr>
        <td class="auto-style2">Rolls Royce</td>
        <td class="auto-style3">Deposit</td>
        <td class="auto-style4">60.000 GBP</td>
        <td>20%</td>
    </tr>
</table>
<p>
    Top News</p>
<table class="auto-style1">
    <tr>
        <td class="auto-style2">Client</td>
        <td>Story</td>
    </tr>
    <tr>
        <td class="auto-style2">Anglo American</td>
        <td>
            <h1 class="title entry-title" ><a href="http://www.miningne.ws/2018/03/12/anglo-american-mining-could-be-double-its-size/#">Anglo American mining could be double its size</a></h1>
        </td>
    </tr>
    <tr>
        <td class="auto-style2">Rolls Royce</td>
        <td>
            <h1 itemprop="mainEntityOfPage" ><a href="https://www.moneyjournals.com/rolls-royce-increases-job-cuts-simplify-staff-structure/">Rolls Royce Increases Job Cuts to Simplify Staff Structure</a></h1>
        </td>
    </tr>
</table>
<p>
    Recommendations</p>
<table class="auto-style1">
    <tr>
        <td class="auto-style2">Client</td>
        <td>Recommendation</td>
    </tr>
    <tr>
        <td class="auto-style2">Anglo American</td>
        <td>Follow Up from Last Meeting on 3/Feb</td>
    </tr>
    <tr>
        <td class="auto-style2">Rolls Royce</td>
        <td>Review Pipeline </td>
    </tr>
</table>

<%  end if %>
