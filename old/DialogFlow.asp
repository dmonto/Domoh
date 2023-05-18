<% 
    
Function BytesToStr(bytes)
    Dim Stream
    Set Stream = Server.CreateObject("Adodb.Stream")
        Stream.Type = 1 'adTypeBinary
        Stream.Open
        Stream.Write bytes
        Stream.Position = 0
        Stream.Type = 2 'adTypeText
        Stream.Charset = "iso-8859-1"
        BytesToStr = Stream.ReadText
        Stream.Close
    Set Stream = Nothing
End Function

    dim blk, loc
    blk=Request.Item(0)
    blk=BytesToStr(Request.BinaryRead(Request.TotalBytes))
%>
{
"displayText": "",
      "messages": [
        {
          "type": 0,
          "speech": "Cheers <%=replace(blk,"""","'")%> it's great to see you back!"
        },
        {
          "type": 0,
          "speech": "I can help you understand your clients and plan your next actions. I am the webhook"
        }
      ]
    }
  