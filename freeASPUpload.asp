<%
'--- For examples, documentation, and your own free copy, go to: http://www.freeaspupload.net
'--- Note: You can copy and use this script for free and you can make changes to the code, but you cannot remove the above comment.

'--- Changes:
'--- Aug 2, 2005: Add support for checkboxes and other input elements with multiple values
'--- Jan 6, 2009: Lars added ASP_CHUNK_SIZE
const DEFAULT_ASP_CHUNK_SIZE = 200000

Class FreeASPUpload
	Public UploadedFiles, FormElements
	private VarArrayBinRequest, StreamRequest, uploadedYet, internalChunkSize

    '--- Inicializa UploadedFiles, FormElements, StreamRequest vacio
    private Sub Class_Initialize()
		set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
		set FormElements = Server.CreateObject("Scripting.Dictionary")
		set StreamRequest = Server.CreateObject("ADODB.Stream")
		StreamRequest.Type = 2	'--- adTypeText
		StreamRequest.Open
		uploadedYet = false
		internalChunkSize = DEFAULT_ASP_CHUNK_SIZE
	end Sub
	
	'--- Libera todo
    private Sub Class_Terminate()
		if IsObject(UploadedFiles) then
			UploadedFiles.RemoveAll()
			set UploadedFiles = Nothing
		end if
		if IsObject(FormElements) then
			FormElements.RemoveAll()
			set FormElements = Nothing
		end If
		StreamRequest.Close
		set StreamRequest = Nothing
	end Sub

	'--- Contenido del campo
    Public Property Get Form(sIndex)
		Form = ""
		if FormElements.Exists(LCase(sIndex)) then Form = FormElements.Item(LCase(sIndex))
	End Property

	'--- Contenedor de los Ficheros
    Public Property Get Files()
		Files = UploadedFiles.Items
	End Property
	
	'--- Devuelve si un campo de Form esta presente
    Public Property Get Exists(sIndex)
    	Exists = False
		if FormElements.Exists(LCase(sIndex)) then Exists = true
    End Property
        
	'--- Devuelve si un Fichero esta presente
    Public Property Get FileExists(sIndex)
        FileExists = False
		if UploadedFiles.Exists(LCase(sIndex)) then FileExists = true
    End Property
        
    '--- Devuelve el tamaño de bloque
    Public Property Get chunkSize()
		chunkSize = internalChunkSize
	End Property

	'--- Fija el tamaño de bloque
    Public Property Let chunkSize(sz)
		internalChunkSize = sz
	End Property

	'--- Calls Upload to extract the data from the binary request and then saves the uploaded files
	Public Sub Save(path)
		dim streamFile, fileItem

		if Right(path, 1) <> "\" then path = path & "\"
		if not uploadedYet then Upload

		for each fileItem In UploadedFiles.Items
			set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = 1
			streamFile.Open
			StreamRequest.Position=fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile path & fileItem.FileName, 2
			streamFile.close
			set streamFile = Nothing
			fileItem.Path = path & fileItem.FileName
		 next
	End Sub
	
	'--- Calls Upload to extract the data from the binary request and then saves file number [num] on [outLocalFileName]
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
			streamFile.Type = 1
			streamFile.Open
			StreamRequest.Position = fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile path & outLocalFileName, 2
			streamFile.Close
			set streamFile = Nothing
			fileItem.Path = path & filename
		end if
	End Sub

	'--- Graba el contenido en debugStream.bin
    Public Function SaveBinRequest(path) '--- For debugging purposes
		StreamRequest.SaveToFile path & "\debugStream.bin", 2
	End Function

	'--- Vuelca el contenido del form en pantalla
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

	'--- Guarda el contenido del form en las estructuras
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
	    readBytes = internalChunkSize
		VarArrayBinRequest = Request.BinaryRead(readBytes)
        VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
        for readLoop = 0 to 300000
            tmpBinRequest = Request.BinaryRead(readBytes)
            if readBytes < 1 then exit for
		    VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
		next
		if Err.Number <> 0 then 
            Response.Write "System reported this error:<p>" & Err.Description & "</p>"
			Response.Write "<p>The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. "
            Response.Write "Please see instructions in the <a href='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</a>.</p>"
			Exit Sub
		end if
		on error goto 0 '--- reset error handling
    	nCurPos = FindToken(tNewLine,1) '--- Note: nCurPos is 1-based (and so is InstrB, MidB, etc)
    	if nCurPos <= 1  then Exit Sub
		 
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

			    set oUploadFile = new UploadedFile			
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
			    dim nEndOfData
				nCurPos = FindToken(tNewLine, nCurPos) + 4 '--- skip empty line
				nEndOfData = FindToken(vDataSep, nCurPos) - 2
				if not FormElements.Exists(LCase(sFieldName)) then 
				    FormElements.Add LCase(sFieldName), Byte2String(MidB(VarArrayBinRequest, nCurPos, nEndOfData-nCurPos))
				else
                    FormElements.Item(LCase(sFieldName))= FormElements.Item(LCase(sFieldName)) & ", " & Byte2String(MidB(VarArrayBinRequest, nCurPos, nEndOfData-nCurPos)) 
                end if 
			end if

			'--- Advance to next separator
			nDataBoundPos = FindToken(vDataSep, nCurPos)
		loop
		
        StreamRequest.WriteText(VarArrayBinRequest)
	End Sub

    '--- Parsea VarArrayBinRequest segun sToken a partir de nStart
    private Function SkipToken(sToken, nStart)
		SkipToken = InStrB(nStart, VarArrayBinRequest, sToken)
		if SkipToken = 0 then
			Response.Write "Error in parsing uploaded binary request. The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. "
            Response.Write "Please see instructions in the <a href='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</a>."
			Response.End
		end if
		SkipToken = SkipToken + LenB(sToken)
	end Function

	'--- Busca sToken en VarArrayBinRequest
    private Function FindToken(sToken, nStart)
		FindToken = InStrB(nStart, VarArrayBinRequest, sToken)
	end Function

	'--- ExtractField Extrae un campo de VarArrayBinRequest separado por sToken
    private Function ExtractField(sToken, nStart)
		dim nEnd

		nEnd = InstrB(nStart, VarArrayBinRequest, sToken)
		if nEnd = 0 then
			Response.Write "Error in parsing uploaded binary request."
			Response.End
		end if
		ExtractField = Byte2String(MidB(VarArrayBinRequest, nStart, nEnd-nStart))
	end Function

	'--- String2Byte does String to byte string conversion
	private Function String2Byte(sString)
		dim i
		for i = 1 to Len(sString)
		   String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
		next
	end Function

	'--- Byte2String does Byte string to string conversion
	private Function Byte2String(bsString)
		dim i, b1, b2, b3, b4
		Byte2String =""
		for i = 1 to LenB(bsString)
		    if AscB(MidB(bsString,i,1)) < 128 then
		        '--- One byte
    		    Byte2String = Byte2String & ChrW(AscB(MidB(bsString,i,1)))
    		elseif AscB(MidB(bsString,i,1)) < 224 then
    		    '--- Two bytes
    		    b1 = AscB(MidB(bsString,i,1))
    		    b2 = AscB(MidB(bsString,i+1,1))
    		    Byte2String = Byte2String & ChrW((((b1 and 28) / 4) * 256 + (b1 and 3) * 64 + (b2 and 63)))
    		    i = i + 1
    		elseif AscB(MidB(bsString,i,1)) < 240 then
    		    '--- Three bytes
    		    b1 = AscB(MidB(bsString,i,1))
    		    b2 = AscB(MidB(bsString,i+1,1))
    		    b3 = AscB(MidB(bsString,i+2,1))
    		    Byte2String = Byte2String & ChrW(((b1 and 15) * 16 + (b2 and 60)) * 256 + (b2 and 3) * 64 + (b3 and 63))
    		    i = i + 2
    		else
    		    '--- Four bytes
    		    b1 = AscB(MidB(bsString,i,1))
    		    b2 = AscB(MidB(bsString,i+1,1))
    		    b3 = AscB(MidB(bsString,i+2,1))
    		    b4 = AscB(MidB(bsString,i+3,1))
    		    '--- Don't know how to handle this, I believe Microsoft doesn't support these characters so I replace them with a "^"
    		    '--- Byte2String = Byte2String & ChrW(((b1 AND 3) * 64 + (b2 AND 63)) & "," & (((b1 AND 28) / 4) * 256 + (b1 AND 3) * 64 + (b2 AND 63)))
    		    Byte2String = Byte2String & "^"
    		    i = i + 3
		    end if
		next
	end Function
End Class

'--- UploadedFile supports each of the uploaded files
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

    '--- Returns file name
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

'--- GetFileName devuelve el nombre a grabar el fichero sin machacar otro
'--- This function is used when saving a file to check there is not already a file with the same name so that you don't overwrite it.
'--- It adds numbers to the filename e.g. file.gif becomes file1.gif becomes file2.gif and so on.
'--- It keeps going until it returns a filename that does not exist.
'--- You could just create a filename from the ID field but that means writing the record - and it still might exist!
'--- N.B. Requires strSaveToPath variable to be available - and containing the path to save to
Function GetFileName(strSaveToPath, FileName)
    dim Counter, Flag, strTempFileName, FileExt, NewFullPath, objFSO, p

    set objFSO = CreateObject("Scripting.FileSystemObject")
    Counter = 0
    p = InStrRev(FileName, ".")
    FileExt = Mid(FileName, p+1)
    strTempFileName = left(FileName, p-1)
    NewFullPath = strSaveToPath & "\" & FileName
    Flag = false
    
    do until Flag = true
        if objFSO.FileExists(NewFullPath) = false then
            Flag = true
            GetFileName = Mid(NewFullPath, InStrRev(NewFullPath, "\") + 1)
        else
            Counter = Counter + 1
            NewFullPath = strSaveToPath & "\" & strTempFileName & Counter & "." & FileExt
        end if
    loop
End Function
%>