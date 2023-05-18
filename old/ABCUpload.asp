<% @Language="VBScript" %>
<% 
Set theForm = Server.CreateObject("ABCUpload4.XForm")
theForm.AbsolutePath = True
'Response.Write(theForm.Files.Count)

If theForm.Files("Filedata").FileExists Then 
  theForm.Files("Filedata").Save Server.MapPath("UploadedFiles") & "\" & theForm.Files("Filedata").SafeFileName
  Response.Write "File uploaded: " & theForm.Files("Filedata").SafeFileName
Else
  Response.Write "<eaferror>No file uploaded...</eaferror>"
End If
%>