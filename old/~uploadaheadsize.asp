<%
    dim WebServer, UploadReadAheadSize, INSTANCE_ID
    '  Set WebServer = GetObject("IIS://localhost/W3SVC[/sitenum[/folder[/uploadscript.asp]]]")
	INSTANCE_ID = request.ServerVariables("INSTANCE_ID")
	on error resume next
    set WebServer = GetObject("IIS://localhost/W3SVC/" & INSTANCE_ID)

	if Err <> 0 then
	    Response.Status = "401 access denied"
		Response.Write "Please log on as an IIS administrator to manage IIS properties" & Replace(Space(1000), " ", "&nbsp;") ' why MS means friendly error messages are friendly??!!
		Response.End
	end if
    on error goto 0
  
	if Request("iUploadReadAheadSize")<>"" and IsNumeric(Request("iUploadReadAheadSize")) then
		WebServer.UploadReadAheadSize = Request("iUploadReadAheadSize")
		WebServer.SetInfo 
	end if

    UploadReadAheadSize = WebServer.UploadReadAheadSize
%>
Please see notes at <a href=http://www.motobit.com/help/scptutl/pa35.htm>Upload - Monitor and handle upload state/result</a> page.
UploadReadAheadSize for this instance (no. <%=INSTANCE_ID%>  & ") is <%=UploadReadAheadSize%>B.
<form><input name=iUploadReadAheadSize value=<%=UploadReadAheadSize%> /><input type=submit value="Set UploadReadAheadSize for Instance <%=INSTANCE_ID%> &gt;&gt;" /></form>
<%
    Response.End
    dim scriptName, scriptPath, pos

	scriptName = Request.ServerVariables("SCRIPT_NAME")
	pos = InStr(1, scriptName , "/")
	do while pos>0 
        scriptPath = Left(scriptName , pos-1)
		Response.Write "IIS://localhost/W3SVC/" & INSTANCE_ID & "/Root" & scriptPath
        set WebServer = GetObject("IIS://localhost/W3SVC/" & INSTANCE_ID & "/Root" & scriptPath)

%>
UploadReadAheadSize for this instance (no. <%=INSTANCE_ID%>  & "), folder <%=ScriptPath%> is <%=UploadReadAheadSize%>B.
<form><input name=iUploadReadAheadSize value=<%=UploadReadAheadSize%> /><input type=submit value="Set UploadReadAheadSize for Instance <%=INSTANCE_ID%> &gt;&gt;"/></form>
<%
    	pos = InStr(pos+1,scriptName , "/")
	loop
%>