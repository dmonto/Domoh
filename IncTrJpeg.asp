<%
Class DigJpeg
	public File, PreserveAspectRatio, Width, Height
	
	Public Sub Open(sFileName)
		File=sFileName
	End Sub
	
	Public Sub Save(vPath)
		dim strImageName, strThumbName, sURL, objXML
		
		strImageName = File
		strThumbName = vPath
		sURL="http://domoh.com/ImageResizer.aspx?image=" & strImageName & "&thumb=" & strThumbName & "&width=-1&height=" & Height
		set objXML = Server.CreateObject("Microsoft.XMLHTTP")
		objXML.Open "GET", sURL, true
		objXML.Send
		set objXML=Nothing
	End Sub
End Class
%>
