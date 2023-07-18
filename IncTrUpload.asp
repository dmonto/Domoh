<%
const DEFAULT_ASP_CHUNK_SIZE = 200000

'---- FileUploader wrapper para subir ficheros a Web
Class FileUploader
	Public maxSize, fileExt, error, errorDesc, UploadedFiles, FormElements
	private mcolFormElem, intMaxFileSize, VarArrayBinRequest, StreamRequest, uploadedYet, internalChunkSize

	'---- Class_Initialize Creates the UploadedFiles, FormElements empty objects 
    private Sub Class_Initialize()
		set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
		set FormElements = Server.CreateObject("Scripting.Dictionary")
		set StreamRequest = Server.CreateObject("ADODB.Stream")
		StreamRequest.Type = 2	'---- adTypeText
		StreamRequest.Open
		uploadedYet = false
		internalChunkSize = DEFAULT_ASP_CHUNK_SIZE
	end Sub
	
	'---- Class_Terminate libera todo
    private Sub Class_Terminate()
		if IsObject(UploadedFiles) then
			UploadedFiles.RemoveAll()
			set UploadedFiles = Nothing
		end If
		
		if IsObject(FormElements) then
			FormElements.RemoveAll()
			set FormElements = Nothing
		end If
		
		StreamRequest.Close
		set StreamRequest = Nothing
	end Sub

	'---- Recupera el valor de un campo del form
    Public Property Get Form(sIndex)
		form = ""
		if FormElements.Exists(LCase(sIndex)) then Form = FormElements.Item(LCase(sIndex))
	End Property

	'---- Recupera la coleccion de ficheros
    Public Property Get Files()
		Files = UploadedFiles.Items
	End Property
	
	'---- Devuelve si existe un campo del form
    Public Property Get Exists(sIndex)
		Exists = false
		if FormElements.Exists(LCase(sIndex)) then Exists = true
	End Property
        
	'---- Devuelve si existe un fichero
    Public Property Get FileExists(sIndex)
		FileExists = false
		if UploadedFiles.Exists(LCase(sIndex)) then FileExists = true
	End Property
        
	'---- Devuelve el tamaño del bloque
    Public Property Get chunkSize()
		chunkSize = internalChunkSize
	End Property

	'---- Fija el tamaño del bloque
    Public Property Let chunkSize(sz)
		internalChunkSize = sz
	End Property

	'---- Devuelve el tamaño máximo de fichero
    Public Property Get maxFileSize()
		maxFileSize = intMaxFileSize
	End Property

	'---- Fija el tamaño máximo de fichero
    Public Property Let maxFileSize(sz)
		intMaxFileSize = sz
	End Property

	'---- Calls Upload to extract the data from the binary request and then saves the uploaded files
	Public Sub Save(path)
		dim streamFile, fileItem

		if Right(path, 1) <> "\" then path = path & "\"
		if not uploadedYet then Upload

		for each fileItem in UploadedFiles.Items
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
	
	'---- Saves one of the uploaded files in [path]
    Public Sub SaveOne(path, num, byref outFileName, byref outLocalFileName)
		dim streamFile, fileItems, fileItem, fs

		set fs = Server.CreateObject("Scripting.FileSystemObject")
		if Right(path, 1) <> "\" then path = path & "\"
		if not uploadedYet then Upload
		
		if UploadedFiles.Count > 0 then
			fileItems = UploadedFiles.Items
			set fileItem = fileItems(num)
		
			if outFileName="" then outFileName = fileItem.FileName
			outLocalFileName = GetFileName(path, outFileName)
			set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = 1
			streamFile.Open
			StreamRequest.Position = fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile path & outLocalFileName, 2
			streamFile.close
			set streamFile = Nothing
			fileItem.Path = path & outFileName
		end if
	End Sub

	'---- Saves uploaded files to debugStream.bin
    Public Function SaveBinRequest(path) '--- For debugging purposes
		StreamRequest.SaveToFile path & "\debugStream.bin", 2
	End Function

	'---- Dumps all form contents on-screen
    Public Sub DumpData() '---- only works if files are plain text
		dim i, aKeys, f
		
		Response.Write "Form Items:"
		aKeys = FormElements.Keys
		for i = 0 To FormElements.Count -1 '---- Iterate the array
			Response.Write aKeys(i) & " = " & FormElements.Item(aKeys(i)) 
		next
		Response.Write "Uploaded Files:"
		for each f In UploadedFiles.Items
			Response.Write "Name: " & f.FileName 
			Response.Write "Type: " & f.ContentType 
			Response.Write "Start: " & f.Start 
			Response.Write "Size: " & f.Length 
        next
   	End Sub

	'---- Dumps all form contents on internal structures
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
		VarArrayBinRequest = midb(VarArrayBinRequest, 1, lenb(VarArrayBinRequest))
		for readLoop = 0 to 300000
			tmpBinRequest = Request.BinaryRead(readBytes)
			if readBytes < 1 then exit for
			VarArrayBinRequest = VarArrayBinRequest & midb(tmpBinRequest, 1, lenb(tmpBinRequest))
		next
		if Err.Number <> 0 then 
            Response.Write "System reported this error:"
			Response.Write "<p>" & Err.Description & "</p>"
		    Response.Write "<p>The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. "
			Response.Write "Please see instructions in the <a href='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</a>.<p>"
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

			    set oUploadFile = new UploadedFile			
			    nCurPos = SkipToken(tFilename, nCurPos)
			    auxStr = ExtractField(tDoubleQuotes, nCurPos)
                '--- We are interested only in the name of the file, not the whole path. Path separator is \ in windows, / in UNIX
                '--- While IE seems to put the whole pathname in the stream, Mozilla seem to 
                '--- only put the actual file name, so UNIX paths may be rare. But not impossible.
                osPathSep = "\"
                if InStr(auxStr, osPathSep) = 0 then osPathSep = "/"
        	    oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))
			    if (Len(oUploadFile.FileName) > 0) then '--- File field not left empty
    				nCurPos = SkipToken(tContentType, nCurPos)
					
                    auxStr = ExtractField(tNewLine, nCurPos)
                    '--- NN on UNIX puts things like this in the stream:
                    '--- ?? python py type=?? python application/x-python
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

	'--- Salta un token sToken en VarArrayBinRequest
    private Function SkipToken(sToken, nStart)
		SkipToken = InStrB(nStart, VarArrayBinRequest, sToken)
		if SkipToken = 0 then
			Response.Write "Error in parsing uploaded binary request. The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. "
            Response.Write "Please see instructions in the <a href='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</a>.<p>"
			Response.End
		end if
		SkipToken = SkipToken + LenB(sToken)
	end Function

	'--- FindToken busca sToken en VarArrayBinRequest
    private Function FindToken(sToken, nStart)
		FindToken = InStrB(nStart, VarArrayBinRequest, sToken)
	end Function

	'--- ExtractField devuelve el siguiente campo separado por sToken en VarArrayBinRequest
    private Function ExtractField(sToken, nStart)
		dim nEnd

		nEnd = InstrB(nStart, VarArrayBinRequest, sToken)
		if nEnd = 0 then
			Response.Write "Error in parsing uploaded binary request."
			Response.End
		end if
		ExtractField = Byte2String(MidB(VarArrayBinRequest, nStart, nEnd-nStart))
	End Function

	'--- String2Byte does String to byte string conversion
	Private Function String2Byte(sString)
		dim i
		for i = 1 to Len(sString)
		   String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
		next
	End Function

	'--- Byte2String does Byte string to string conversion
	Private Function Byte2String(bsString)
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
	End Function

	'--- UploadViejo is the old upload function
    Public Default Sub UploadViejo()
		dim biData, sInputName, nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos, nPosFile, nPosBound, counter, sFileName1

		error = false
		errorDesc = ""
		
		on error resume next
		Err.Clear
		biData = Request.BinaryRead(Request.TotalBytes)
		if Err then 
			Mail "diego@domoh.com", "Error en IncTrUpload", "BinaryRead - " & Request.TotalBytes & ": " & Err.Description
		  	Error=true
		  	errorDesc=MesgS("Por favor sube menos ficheros o más pequeños", "Please upload less or smaller files") 
		  	Exit Sub
		end if		
		nPosBegin = 1
		nPosEnd = InStrB(nPosBegin, biData, CByteString(Chr(13)))
		if (nPosEnd-nPosBegin) <= 0 then Exit Sub		 
		vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
		nDataBoundPos = InStrB(1, biData, vDataBounds)
		counter = 0
		
		do until nDataBoundPos = InStrB(biData, vDataBounds & CByteString("--"))		
			counter = counter + 1	
			nPos = InStrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InStrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InStrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InStrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InStrB(nPosEnd, biData, vDataBounds)
			
			if nPosFile <> 0 and nPosFile < nPosBound then
			    dim oUploadFile, sFileName, sFileExt
				set oUploadFile = new UploadedFile
				
				nPosBegin = nPosFile + 10
				nPosEnd = InStrB(nPosBegin, biData, CByteString(Chr(34)))				
				sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				sFileName1 = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))
				sfileExt = Right(sFileName1, Len(sFileName1) - InStrRev(sFileName1, "."))
				
				if not IsNull(fileExt) then
				    if InStr(UCase(fileExt), UCase(sFileExt)) = 0 then
				  	    Error = true
				  	    errorDesc = errorDesc & "File " & oUploadFile.FileName & ":" & MesgS("Tipo de fichero no permitido (solo ","The file type is not allowed (only ") & fileExt & ")" 
				    else
    				    nPos = InStrB(nPosEnd, biData, CByteString("Content-Type:"))
    				    nPosBegin = nPos + 14
    				    nPosEnd = InStrB(nPosBegin, biData, CByteString(Chr(13)))
    				    oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
    				    nPosBegin = nPosEnd+4
    				    nPosEnd = InStrB(nPosBegin, biData, vDataBounds) - 2
    				    oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
    				    if (oUploadFile.FileSize > 0) and (oUploadFile.FileSize <= maxFileSize) then Files.Add LCase(sInputName), oUploadFile
    
    				    if not IsNull(maxFileSize) then
      				        if oUploadFile.FileSize > maxFileSize then
    	 			  	        Error = true
    				  	        errorDesc = errorDesc & MesgS("Fichero <b>","File <b>") & oUploadFile.FileName & "</b>: " & MesgS("El fichero es demasiado grande (max: 300 kb)","The file is too large (max: 300kb)") & "<br/>"
    				  	        'Exit Sub
    				        end if				
    				    end if
				    end if
				end if
			else
				nPos = InStrB(nPos, biData, CByteString(Chr(13)))
				nPosBegin = nPos + 4
				nPosEnd = InStrB(nPosBegin, biData, vDataBounds) - 2
				if not mcolFormElem.Exists(LCase(sInputName)) then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			end if

			nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
		loop
	End Sub

	'--- Upload2 uses ADO to save to SaveFile.txt
    Sub Upload2()
	    dim TotalSize, ChunkSize, BytesLeft, CurrentBytes, objADO

	    TotalSize = Request.TotalBytes
	    ChunkSize = 64*1024
	    if ChunkSize > TotalSize Then ChunkSize = TotalSize
	    BytesLeft = TotalSize
	    set objADO = CreateObject("ADODB.Stream")
	    objADO.Type = 1
	    objADO.Open

	    do while BytesLeft > 0
		    if BytesLeft < ChunkSize then
			    ChunkSize = BytesLeft
		    end if
		    objADO.Write Request.BinaryRead(ChunkSize)
		    BytesLeft = BytesLeft - ChunkSize
	    loop

	    objADO.SaveToFile "C:\WebUploads\SaveFile.txt"
	    objADO.Close
    End Sub

    '--- Upload2 uses ADO to save Request in one chunk
    Sub Upload3()
	    dim recordSet
	    recordSet=Server.CreateObject("ADODB.Recordset")
	    recordSet.fields.append 0, 201, Request.TotalBytes
	    recordSet.open()
	    recordSet.addNew()
	    recordSet.fields(0).appendChunk(Request.binaryRead(Request.TotalBytes))

	    dim data
	    data=cWideString(recordSet.Fields(0))
    End Sub

	'--- CByteString does String to byte string conversion
	Private Function CByteString(sString)
		dim nIndex
		for nIndex = 1 to Len(sString)
		    CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
		next
	End Function

	'--- CWideString does Byte string to string conversion
	Private Function CWideString(bsString)
		dim nIndex
		CWideString =""
		for nIndex = 1 to LenB(bsString)
		    CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
		next
	End Function
End Class

'--- UploadedFile supports each Uploaded File
Class UploadedFile
	Public ContentType, Start, Length, Path
	Private nameOfFile

    '--- Need to remove characters before storing FileName that are valid in UNIX, but not in Windows
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

    '--- FileName returns the name of the file
    Public Property Get FileName()
        FileName = nameOfFile
    End Property

	'--- FileSize returns the size of the file
    Public Property Get FileSize()
		FileSize = Length
	End Property

	'--- Ext returns the extension of the file
    Public Property Get Ext()
		Ext = Right(nameOfFile,4)
	End Property

	'--- SaveToDisk stores the file in [sPath]
	Public Sub SaveToDisk(sPath)
		dim oFS, oFile, nIndex
	
		if sPath = "" or FileName = "" then Exit Sub
		if Mid(sPath, Len(sPath)) <> "\" then sPath = sPath & "\"
		set oFS = Server.CreateObject("Scripting.FileSystemObject")
		if not oFS.FolderExists(sPath) then Exit Sub
		
		if oFS.FileExists(sPath & FileName) then
       	    on error resume next
       	    oFS.MoveFile sPath & FileName, sPath & "old-" & FileName
       	    on error goto 0
		end if

		set oFile = oFS.CreateTextFile(sPath & FileName, true)
		
		for nIndex = 1 to LenB(FileData)
	        oFile.Write Chr(AscB(MidB(FileData,nIndex,1)))
		next
		oFile.Close
		set oFS = nothing
	End Sub
	
	'--- SaveToDatabase stores the file in [oField]
	Public Sub SaveToDatabase(ByRef oField)
		if LenB(FileData) = 0 then Exit Sub
		if IsObject(oField) then oField.AppendChunk FileData
	End Sub
End Class

'--- SubstNoReg Does not depend on RegEx, which is not available on older VBScript. Is not recursive, which means it will not run out of stack space
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

'--- GetFileName is used when saving a file to check there is not already a file with the same name so that you don't overwrite it.
'--- It adds numbers to the filename e.g. file.gif becomes file1.gif becomes file2.gif and so on.
'--- It keeps going until it returns a filename that does not exist.
'--- You could just create a filename from the ID field but that means writing the record - and it still might exist!
'--- N.B. Requires strSaveToPath variable to be available - and containing the path to save to
Function GetFileName(strSaveToPath, FileName)
    dim Counter, Flag, strTempFileName, FileExt, NewFullPath, objFSO, p

    set objFSO = CreateObject("Scripting.FileSystemObject")
    Counter = 0
    p = instrrev(FileName, ".")
    FileExt = Mid(FileName, p+1)
    strTempFileName = Left(FileName, p-1)
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

'--- IsImage creates an ImageSize and returns whether [strImageName] has image format
Function IsImage(strImageName)
	dim objImageSize
	
    set objImageSize = new ImageSize
    with objImageSize
        .ImageFile = Server.MapPath(strImageName)
        if .IsImage then IsImage = true else IsImage = false
    end with
    set objImageSize = nothing
End Function

'--- ImageSize holds the file type & features
Class ImageSize
    Public Dimensioni(2), bolIsImage, strPathFile, strImageType
  
    Private Sub Class_Initialize()
    End Sub
  
    Private Sub Class_Terminate()
    End Sub

    '--- Returns the name of file
    Public Property Get ImageName
        ImageName = doName(ImageFile)
    End Property
  
    '--- Returns the path
    Public Property Get ImageFile
  	    ImageFile = strPathFile
    End Property
  
    '--- Sets the path
    Public Property Let ImageFile(strImageFile)
  	    strPathFile = strImageFile
  	    gfxSpex(strPathFile)
    End Property
  
    '--- ImageWidth returns the width
    Public Property Get ImageWidth
        ImageWidth = Dimensioni(1)
    End Property
  
    '--- ImageHeight returns the height
    Public Property Get ImageHeight
  	    ImageHeight = Dimensioni(0)
    End Property
  
    '--- ImageDepth returns the depth
    Public Property Get ImageDepth
  	    ImageDepth = Dimensioni(2)
    End Property 
  
    '--- ImageType returns the extension
    Public Property Get ImageType
  	    ImageType = strImageType
    End Property      
  
    '--- IsImage returns whether the file has image format
    Public Property Get IsImage
  	    IsImage = bolIsImage
    End Property     

    '--- GetBytes returns [bytes] from [flnm] after [offset]
    Public Function GetBytes(flnm, offset, bytes)
        dim objFSO, objFTemp, objTextStream, lngSize
        on error resume next
    
        set objFSO = CreateObject("Scripting.FileSystemObject")
    
        '--- First, we get the filesize
        set objFTemp = objFSO.GetFile(flnm)
        lngSize = objFTemp.Size
        set objFTemp = nothing
        fsoForReading = 1
        set objTextStream = objFSO.OpenTextFile(flnm, fsoForReading)  
        if offset > 0 then strBuff = objTextStream.Read(offset - 1)
    
        if bytes = -1 then		'--- Get All!
            GetBytes = objTextStream.Read(lngSize)  '--- ReadAll
        else
            GetBytes = objTextStream.Read(bytes)
        end if
    
        objTextStream.Close
        set objTextStream = nothing
        set objFSO = nothing
    End Function

    '--- lngConvert converts the bigendian ascii codes in [strTemp] to a long
    Public Function lngConvert(strTemp)
        lngConvert = CLng(Asc(Left(strTemp, 1)) + ((Asc(Right(strTemp, 1)) * 256)))
    End Function

    '--- lngConvert converts the littleendian ascii codes in [strTemp] to a long
    Public Function lngConvert2(strTemp)
        lngConvert2 = CLng(Asc(Right(strTemp, 1)) + ((Asc(Left(strTemp, 1)) * 256)))
    End Function
  
    '--- gfxSpex gets the details from the image file [flnm]
    Public Sub gfxSpex(flnm)
        dim strPNG, strBuff, lngSize, flgFound, strTarget, strGIF, strBMP, strType

        strType = ""
        strImageType = "(unknown)"
        bolIsImage = False
        strPNG = Chr(137) & Chr(80) & Chr(78)
        strGIF = "GIF"
        strBMP = Chr(66) & Chr(77)
        strType = GetBytes(flnm, 0, 3)

        if strType = strGIF then				'--- is GIF
            strImageType = "GIF"
            lngWidth = lngConvert(GetBytes(flnm, 7, 2))
            lngHeight = lngConvert(GetBytes(flnm, 9, 2))
            lngDepth = 2 ^ ((Asc(GetBytes(flnm, 11, 1)) and 7) + 1)
            bolIsImage = true
        elseif Left(strType, 2) = strBMP then		'--- is BMP
            strImageType = "BMP"
            lngWidth = lngConvert(GetBytes(flnm, 19, 2))
            lngHeight = lngConvert(GetBytes(flnm, 23, 2))
            lngDepth = 2 ^ (Asc(GetBytes(flnm, 29, 1)))
            bolIsImage = true
        elseif strType = strPNG then			'--- Is PNG
            strImageType = "PNG"
            lngWidth = lngConvert2(GetBytes(flnm, 19, 2))
            lngHeight = lngConvert2(GetBytes(flnm, 23, 2))
            lngDepth = getBytes(flnm, 25, 2)
            
            select case Asc(Right(lngDepth,1))
            case 0
                lngDepth = 2 ^ (Asc(Left(lngDepth, 1)))
                bolIsImage = true
            case 2
                lngDepth = 2 ^ (Asc(Left(lngDepth, 1)) * 3)
                bolIsImage = true
            case 3
                lngDepth = 2 ^ (Asc(Left(lngDepth, 1)))  '8
                bolIsImage = true
            case 4
                lngDepth = 2 ^ (Asc(Left(lngDepth, 1)) * 2)
                bolIsImage = true
            case 6
                lngDepth = 2 ^ (Asc(Left(lngDepth, 1)) * 4)
                bolIsImage = true
            case else
                lngDepth = -1
            end select
        else
            strBuff = GetBytes(flnm, 0, -1)		'--- Get all bytes from file
            lngSize = Len(strBuff)
            flgFound = 0
            strTarget = Chr(255) & Chr(216) & Chr(255)
            flgFound = InStr(strBuff, strTarget)
            if flgFound = 0 then Exit Sub
            strImageType = "JPG"
            lngPos = flgFound + 2
            ExitLoop = false
 
            do while ExitLoop = false and lngPos < lngSize
                do while Asc(Mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
                    lngPos = lngPos + 1
                loop
 
                if Asc(Mid(strBuff, lngPos, 1)) < 192 or Asc(Mid(strBuff, lngPos, 1)) > 195 then
                    lngMarkerSize = lngConvert2(Mid(strBuff, lngPos + 1, 2))
                    lngPos = lngPos + lngMarkerSize  + 1
                else
                    ExitLoop = true
                end if
            loop
 
            if ExitLoop = false then
                lngWidth = -1
                lngHeight = -1
                lngDepth = -1
            else
                lngHeight = lngConvert2(Mid(strBuff, lngPos + 4, 2))
                lngWidth = lngConvert2(Mid(strBuff, lngPos + 6, 2))
                lngDepth = 2 ^ (Asc(Mid(strBuff, lngPos + 8, 1)) * 8)
                bolIsImage = true
            end if                   
        end if

        Dimensioni(0) = lngHeight
        Dimensioni(1) = lngWidth
        Dimensioni(2) = lngDepth
    End Sub  

    '--- doName returns the file name from the path [strPath]
    Public Function doName(strPath)
        dim arrSplit
        arrSplit = Split(strPath, "\")
        doName = arrSplit(UBound(arrSplit))
    End Function
End Class

'--- GrabaFoto dumps the picture in vUpload using fotos/[vPrefijo] and a minified version in minifotos/[vPrefijo]
Function GrabaFoto(vUpload, vPrefijo)
	dim fich, Jpeg, fName, sNombreFinal, i, j
	on error resume next
	
	Err.Clear
	if vUpload.FileExists("foto") then
		on error goto 0
		set fich=vUpload.UploadedFiles.Item("foto")
		fName="fotos/" & vPrefijo & Session("Usuario") & fich.Ext
		Upload.SaveOne Server.MapPath("/"), j, fName, sNombreFinal
		set Jpeg = new DigJpeg
		Jpeg.Open fName
		Jpeg.Height = 100
		Jpeg.Save "mini" & fName
		GrabaFoto="'" & fName & "'"
	else
		GrabaFoto="null"
	end if
End Function
%>
