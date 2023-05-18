<%
'---  For examples, documentation, and your own free copy, go to: http://www.freeaspupload.net
'---  Note: You can copy and use this script for free and you can make changes to the code, but you cannot remove the above comment.

'--- Changes:
'--- Aug 2, 2005: Add support for checkboxes and other input elements with multiple values
'--- Jan 6, 2009: Lars added ASP_CHUNK_SIZE
'--- Sep 3, 2010: Enforce UTF-8 everywhere; new function to convert byte array to unicode string
const DEFAULT_ASP_CHUNK_SIZE = 200000
const adModeReadWrite = 3
const adTypeBinary = 1
const adTypeText = 2
const adSaveCreateOverWrite = 2

Class FreeASPUpload
	Public UploadedFiles, FormElements
    private VarArrayBinRequest, StreamRequest, uploadedYet
	private internalChunkSize

	private Sub Class_Initialize()
		set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
		set FormElements = Server.CreateObject("Scripting.Dictionary")
		set StreamRequest = Server.CreateObject("ADODB.Stream")
		StreamRequest.Type = adTypeText
		StreamRequest.Open
		uploadedYet = False
		internalChunkSize = DEFAULT_ASP_CHUNK_SIZE
	end Sub
	
	private Sub Class_Terminate()
		if IsObject(UploadedFiles) then
			UploadedFiles.RemoveAll()
			set UploadedFiles = Nothing
		end if
		if IsObject(FormElements) then
			FormElements.RemoveAll()
			set FormElements = Nothing
		end if
		StreamRequest.Close
		set StreamRequest = Nothing
	end Sub

	Public Property Get Form(sIndex)
		Form = ""
		if FormElements.Exists(LCase(sIndex)) then Form = FormElements.Item(LCase(sIndex))
	End Property

	Public Property Get Files()
		Files = UploadedFiles.Items
	End Property
	
    Public Property Get Exists(sIndex)
        Exists = False
        if FormElements.Exists(LCase(sIndex)) then Exists = True
    End Property
        
    Public Property Get FileExists(sIndex)
        FileExists = False
        if UploadedFiles.Exists(LCase(sIndex)) then FileExists = True
    End Property
        
    Public Property Get chunkSize()
		chunkSize = internalChunkSize
	End Property

	Public Property Let chunkSize(sz)
		internalChunkSize = sz
	End Property

	'--- Calls Upload to extract the data from the binary request and then saves the uploaded files
	Public Sub Save(path)
		dim streamFile, fileItem, filePath

		if Right(path, 1) <> "\" then path = path & "\"
		if not uploadedYet then Upload

		for each fileItem in UploadedFiles.Items
			filePath = path & fileItem.FileName
			set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = adTypeBinary
			streamFile.Open
			StreamRequest.Position=fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile filePath, adSaveCreateOverWrite
			streamFile.Close
			set streamFile = Nothing
			fileItem.Path = filePath
		 next
	End Sub
	
	Public Sub SaveOne(path, num, ByRef outFileName, ByRef outLocalFileName)
		dim streamFile, fileItems, fileItem, fs

        set fs = Server.CreateObject("Scripting.FileSystemObject")
		if Right(path, 1) <> "\" then path = path & "\"

		if not uploadedYet then Upload
		if UploadedFiles.Count > 0 then
			fileItems = UploadedFiles.Items
			set fileItem = fileItems(num)
		
			outFileName = fileItem.FileName
			outLocalFileName = GetFileName(path, outFileName)
		
			set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = adTypeBinary
			streamFile.Open
			StreamRequest.Position = fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile path & outLocalFileName, adSaveCreateOverWrite
			streamFile.Close
			set streamFile = Nothing
			fileItem.Path = path & filename
		end if
	End Sub

	Public Function SaveBinRequest(path) '--- For debugging purposes
		StreamRequest.SaveToFile path & "\debugStream.bin", 2
	End Function

	Public Sub DumpData() '--- only works if files are plain text
		dim i, aKeys, f

		Response.Write "Form Items:"
		aKeys = FormElements.Keys
		for i = 0 to FormElements.Count -1 '--- Iterate the array
			Response.Write aKeys(i) & " = " & FormElements.Item(aKeys(i)) 
		next
		Response.Write "Uploaded Files:"
		for each f in UploadedFiles.Items
			Response.Write "Name: " & f.FileName & "Type: " & f.ContentType & "Start: " & f.Start & "Size: " & f.Length 
		next
   	End Sub

	Public Sub Upload()
		dim nCurPos, nDataBoundPos, nLastSepPos, nPosFile, nPosBound, sFieldName, osPathSep, auxStr, readBytes, readLoop, tmpBinRequest
		
		'--- RFC1867 Tokens
		dim vDataSep, tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
		tNewLine = String2Byte(Chr(13))
		tDoubleQuotes = String2Byte(Chr(34))
		tTerm = String2Byte("--")
		tFilename = String2Byte("filename=""")
		tName = String2Byte("name=""")
		tContentDisp = String2Byte("Content-Disposition")
		tContentType = String2Byte("Content-Type:")

		uploadedYet = true

		on error resume next
        '--- Copy binary request to a byte array, on which functions like InstrB and others can be used to search for separation tokens
		readBytes = internalChunkSize
		VarArrayBinRequest = Request.BinaryRead(readBytes)
		VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
		do until readBytes < 1
            tmpBinRequest = Request.BinaryRead(readBytes)
			if readBytes > 0 then
			    VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
			end if
        loop
		StreamRequest.WriteText(VarArrayBinRequest)
		StreamRequest.Flush()
	    if Err.Number <> 0 then 
            Response.Write "System reported this error:<p>"
            Response.Write Err.Description
            Response.Write "<p>The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. Please see instructions in the "
            Response.Write "<a href='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</a>.<p>"
			Exit Sub
		end if
		
        on error goto 0 '--- reset error handling
        nCurPos = FindToken(tNewLine,1) '--- Note: nCurPos is 1-based (and so is InstrB, MidB, etc)

	    if nCurPos <= 1 then Exit Sub
		 
		'--- vDataSep is a separator like -----------------------------21763138716045
		vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)

		'--- Start of current separator
		nDataBoundPos = 1

		'--- Beginning of last line
		nLastSepPos = FindToken(vDataSep & tTerm, 1)

		do until nDataBoundPos = nLastSepPos
			nCurPos = SkipToken(tContentDisp, nDataBoundPos)
			nCurPos = SkipToken(tName, nCurPos)
			sFieldName = ExtractField(tDoubleQuotes, nCurPos)

			nPosFile = FindToken(tFilename, nCurPos)
			nPosBound = FindToken(vDataSep, nCurPos)
			
			if nPosFile <> 0 and nPosFile < nPosBound then
				dim oUploadFile

				set oUploadFile = New UploadedFile			
				nCurPos = SkipToken(tFilename, nCurPos)
				auxStr = ExtractField(tDoubleQuotes, nCurPos)
                '--- We are interested only in the name of the file, not the whole path
                '--- Path separator is \ in windows, / in UNIX
                '--- While IE seems to put the whole pathname in the stream, Mozilla seem to only put the actual file name, so UNIX paths may be rare. But not impossible.
                osPathSep = "\"
                if InStr(auxStr, osPathSep) = 0 then osPathSep = "/"
				oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))

				if (Len(oUploadFile.FileName) > 0) then '--- File field not left empty
					nCurPos = SkipToken(tContentType, nCurPos)
					
                    auxStr = ExtractField(tNewLine, nCurPos)
                    '--- NN on UNIX puts things like this in the stream: ?? python py type=?? python application/x-python
					oUploadFile.ContentType = Right(auxStr, Len(auxStr)-InStrRev(auxStr, " "))
					nCurPos = FindToken(tNewLine, nCurPos) + 4 '--- skip empty line
					
					oUploadFile.Start = nCurPos+1
					oUploadFile.Length = FindToken(vDataSep, nCurPos) - 2 - nCurPos
					
					if oUploadFile.Length > 0 then UploadedFiles.Add LCase(sFieldName), oUploadFile
				end if
			else
				dim nEndOfData, fieldValueUniStr

				nCurPos = FindToken(tNewLine, nCurPos) + 4 '--- skip empty line
				nEndOfData = FindToken(vDataSep, nCurPos) - 2
				fieldValueuniStr = ConvertUtf8BytesToString(nCurPos, nEndOfData-nCurPos)
				if not FormElements.Exists(LCase(sFieldName)) then 
					FormElements.Add LCase(sFieldName), fieldValueuniStr
				else
                    FormElements.Item(LCase(sFieldName))= FormElements.Item(LCase(sFieldName)) & ", " & fieldValueuniStr
                end if 
			end if

			'--- Advance to next separator
			nDataBoundPos = FindToken(vDataSep, nCurPos)
		loop
	End Sub

	private Function SkipToken(sToken, nStart)
		SkipToken = InStrB(nStart, VarArrayBinRequest, sToken)
		if SkipToken = 0 then
			Response.Write "Error in parsing uploaded binary request. The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. "
            Response.Write "Please see instructions in the <a href='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</a>.<p>"
			Response.End
		end if
		SkipToken = SkipToken + LenB(sToken)
	end Function

	private Function FindToken(sToken, nStart)
		FindToken = InStrB(nStart, VarArrayBinRequest, sToken)
	end Function

	private Function ExtractField(sToken, nStart)
		dim nEnd

		nEnd = InStrB(nStart, VarArrayBinRequest, sToken)
		if nEnd = 0 then
			Response.Write "Error in parsing uploaded binary request."
			Response.End
		end if
		ExtractField = ConvertUtf8BytesToString(nStart, nEnd-nStart)
	end Function

	'--- String to byte string conversion
	private Function String2Byte(sString)
		dim i

		for i = 1 to Len(sString)
		   String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
		next
	end Function

	private Function ConvertUtf8BytesToString(start, length)	
		StreamRequest.Position = 0
	
	    dim objStream, strTmp
	    
	    '--- init stream
	    set objStream = Server.CreateObject("ADODB.Stream")
	    objStream.Charset = "utf-8"
	    objStream.Mode = adModeReadWrite
	    objStream.Type = adTypeBinary
	    objStream.Open
	    
	    '--- write bytes into stream
	    StreamRequest.Position = start+1
	    StreamRequest.CopyTo objStream, length
	    objStream.Flush
	    
	    '--- rewind stream and read text
	    objStream.Position = 0
	    objStream.Type = adTypeText
	    strTmp = objStream.ReadText
	    
	    '--- close up and return
	    objStream.Close
	    set objStream = Nothing
	    ConvertUtf8BytesToString = strTmp	
	end Function
End Class

Class UploadedFile
	Public ContentType, Start, Length, Path
	private nameOfFile

    '--- Need to remove characters that are valid in UNIX, but not in Windows
    Public Property Let FileName(fN)
        nameOfFile = fN
        nameOfFile = SubstNoReg(nameOfFile, "\", "_")
        nameOfFile = SubstNoReg(nameOfFile, "/", "_")
        nameOfFile = SubstNoReg(nameOfFile, ":", "_")
        nameOfFile = SubstNoReg(nameOfFile, "*", "_")
        nameOfFile = SubstNoReg(nameOfFile, "?", "_")
        nameOfFile = SubstNoReg(nameOfFile, """", "_")
        nameOfFile = SubstNoReg(nameOfFile, "<", "_")
        nameOfFile = SubstNoReg(nameOfFile, ">", "_")
        nameOfFile = SubstNoReg(nameOfFile, "|", "_")
    End Property

    Public Property Get FileName()
        FileName = nameOfFile
    End Property

    '--- Public Property Get FileN()ame
End Class

'--- Does not depend on RegEx, which is not available on older VBScript
'--- Is not recursive, which means it will not run out of stack space
Function SubstNoReg(initialStr, oldStr, newStr)
    dim currentPos, oldStrPos, skip

    if IsNull(initialStr) or Len(initialStr) = 0 then
        SubstNoReg = ""
    elseif IsNull(oldStr) or Len(oldStr) = 0 then
        SubstNoReg = initialStr
    else
        if IsNull(newStr) then newStr = ""
        currentPos = 1
        oldStrPos = 0
        SubstNoReg = ""
        skip = Len(oldStr)
        do while currentPos <= Len(initialStr)
            oldStrPos = InStr(currentPos, initialStr, oldStr)
            if oldStrPos = 0 then
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, Len(initialStr) - currentPos + 1)
                currentPos = Len(initialStr) + 1
            else
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, oldStrPos - currentPos) & newStr
                currentPos = oldStrPos + skip
            end if
        loop
    end if
End Function

'--- This function is used when saving a file to check there is not already a file with the same name so that you don't overwrite it.
'--- It adds numbers to the filename e.g. file.gif becomes file1.gif becomes file2.gif and so on.
'--- It keeps going until it returns a filename that does not exist.
'--- You could just create a filename from the ID field but that means writing the record - and it still might exist!
'--- N.B. Requires strSaveToPath variable to be available - and containing the path to save to
Function GetFileName(strSaveToPath, FileName)
    dim Counter, Flag, strTempFileName, FileExt, NewFullPath, objFSO, p

    set objFSO = CreateObject("Scripting.FileSystemObject")
    Counter = 0
    p = instrrev(FileName, ".")
    FileExt = mid(FileName, p+1)
    strTempFileName = left(FileName, p-1)
    NewFullPath = strSaveToPath & "\" & FileName
    Flag = false
    
    do until Flag = true
        if objFSO.FileExists(NewFullPath) = false then
            Flag = true
            GetFileName = Mid(NewFullPath, InstrRev(NewFullPath, "\") + 1)
        else
            Counter = Counter + 1
            NewFullPath = strSaveToPath & "\" & strTempFileName & Counter & "." & FileExt
        end if
    loop
End Function 
%>
