<%
'--- Pure-ASP upload v. 2.06 (with progress bar) This software is a FreeWare with limitted use.
'--- 1. You can use this software to upload files with size up to 10MB for free if you want to upload bigger files, please register ScriptUtilities and Huge-ASP upload
'--- See license info at http://asp-upload.borec.net
'--- 2. I'll be glad if you include <a href="http://asp-upload.borec.net">Pure ASP file upload</a> or similar text link to http://www.motobit.com on the web site using Pure-ASP upload.
'--- Feel free to send comments/suggestions to help@pstruh.cz
const adTypeBinary = 1
const adTypeText = 2

const xfsCompleted    = &H0 '--- 0  Form was successfully processed. 
const xfsNotPost      = &H1 '--- 1  Request method is NOT post 
const xfsZeroLength   = &H2 '--- 2  Zero length request (there are no data in a source form) 
const xfsInProgress   = &H3 '--- 3  Form is in a middle of process. 
const xfsNone         = &H5 '--- 5  Initial form state 
const xfsError        = &HA '--- 10  
const xfsNoBoundary   = &HB '--- 11  Boundary of multipart/form-data is not specified. 
const xfsUnknownType  = &HC '--- 12  Unknown source form (Content-type must be multipart/form-data) 
const xfsSizeLimit    = &HD '--- 13  Form size exceeds allowed limit (ScriptUtils.ASPForm.SizeLimit) 
const xfsTimeOut      = &HE '--- 14  Upload time exceeds allowed limit (ScriptUtils.ASPForm.ReadTimeout) 
const xfsNoConnected  = &HF '--- 15  Client was disconnected before upload was completted.
const xfsErrorBinaryRead = &H10 '--- 16  Unexpected error from Request.BinaryRead method (ASP error).

'--- This class emulates ASPForm Class of Huge-ASP upload, http://upload.Motobit.cz
'--- See ScriptUtilities and Huge-ASP file upload help (ScptUtl.chm)
Class ASPForm
	private m_ReadTime
	Public ChunkReadSize, BytesRead, TotalBytes, UploadID

	'--- non-used properties.
	Public TempPath, MaxMemoryStorage, CharSet, FormType, SourceData, ReadTimeout

	Public Default Property Get Item(Key)
		Read
		set Item = m_Items.Item(Key)
	End Property

	Public Property Get Items
		Read
		set Items = m_Items
	End Property

	Public Property Get Files
		Read
		set Files = m_Items.Files
	End Property

	Public Property Get Texts
		Read
		set Texts = m_Items.Texts
	End Property
	
	Public Property Get NewUploadID
		Randomize
		NewUploadID = CLng(Rnd * &H7FFFFFFF)
	End Property

	Public Property Get ReadTime
		if isempty(m_ReadTime) then
			if not IsEmpty(StartUploadTime) then ReadTime = CLng((Now() - StartUploadTime) * 86400 * 1000)
		else '--- For progress window.
			ReadTime = m_ReadTime
		end if
	End Property

	Public Property Get State
		if m_State = xfsNone then Read
		State = m_State
	End Property
    
	private Function CheckRequestProperties
        '--- Wscript.Echo "**CheckRequestProperties"
	    if UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" then '--- Request method must be "POST"
			m_State = xfsNotPost 
			exit Function
		end if '--- If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 
	
		dim CT
		CT = Request.ServerVariables("HTTP_Content_Type") '--- reads Content-Type header
		if Len(CT) = 0 then CT = Request.ServerVariables("CONTENT_TYPE") '--- reads Content-Type header UNIX/Linux 
	    if LCase(Left(CT, 19)) <> "multipart/form-data" then '--- Content-Type header must be "multipart/form-data"
			m_State = xfsUnknownType 
			exit Function
		end if '--- If LCase(Left(CT, 19)) <> "multipart/form-data" Then 

		dim PosB '--- help position variable
		'--- This is upload request. Get the boundary and length from Content-Type header
		PosB = InStr(LCase(CT), "boundary=") '--- Finds boundary
		if PosB = 0 then
			m_State = xfsNoBoundary
			exit Function
		end if 'If PosB = 0 Then
		if PosB > 0 then Boundary = Mid(CT, PosB + 9) '--- Separetes boundary

		'--- ****** Error of IE5.01 - doubbles http header
		PosB = InStr(LCase(CT), "boundary=") 
		if PosB > 0 then '--- Patch for the IE error
			PosB = InStr(Boundary, ",")
			if PosB > 0 then Boundary = Left(Boundary, PosB - 1)
		end if
		
		'--- ****** Error of IE5.01 - doubbles http header
		on error resume next
		TotalBytes = Request.TotalBytes
		if Err<>0 then
			'--- For UNIX/Linux 
			TotalBytes = CLng(Request.ServerVariables("HTTP_Content_Length")) '--- Get Content-Length header
			if Len(TotalBytes)=0 then TotalBytes = CLng(Request.ServerVariables("CONTENT_LENGTH")) '--- Get Content-Length header
		end if
		
		if TotalBytes = 0 then
			m_State = xfsZeroLength 
			exit Function
		end if

		if IsInSizeLimit(TotalBytes) then '--- Form data are in allowed limit
			CheckRequestProperties = true
			m_State = xfsInProgress 
		else   '--- Form data are in allowed limit
			'--- Form data exceeds the limit.
			m_State = xfsSizeLimit	
		end if '--- Form data are in allowed limit
	end Function
    
	'--- reads source data using BinaryRead and store them in SourceData stream
	Public Sub Read()
		if m_State <> xfsNone then Exit Sub

		'--- Wscript.Echo "**Read"
		if not CheckRequestProperties then 
			WriteProgressInfo
			Exit Sub
		end if

		'--- Initialize binary store stream
		if IsEmpty(bSourceData) then set bSourceData = CreateObject("ADODB.Stream")
		bSourceData.Open
		bSourceData.Type = 1 '--- Binary

		'--- Initialize Read variables.
		dim DataPart, PartSize
		BytesRead = 0
		StartUploadTime = Now

		'--- read source data stream in chunks of ChunkReadSize
		do while BytesRead < TotalBytes
			'--- Read chunk of data
			PartSize = ChunkReadSize
			if PartSize + BytesRead > TotalBytes then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize
			'--- Wscript.Echo PartSize

			'--- Store the part size in our stream
			bSourceData.Write DataPart

			'--- Write progress info for secondary window.
			WriteProgressInfo

			'--- Check if the client is still connected
			if not Response.IsClientConnected then
				m_State = xfsNoConnected  
				Exit Sub
			end if
		loop
		m_State = xfsCompleted

		'--- We have all source data in bSourceData stream
		ParseFormData
	End Sub

	private Sub ParseFormData
		dim Binary

		bSourceData.Position = 0
		Binary = bSourceData.Read
		'--- wscript.echo "Binary", LenB(Binary)
		m_Items.mpSeparateFields Binary, Boundary
	end Sub

	'--- This function reads progress info data from a temporary file.
	Public Function getForm(FormID)
		if IsEmpty(ProgressFile.UploadID) then '--- Was UploadID of ProgressFile set?
			ProgressFile.UploadID = FormID
		end if

		'--- Get progress data
		dim ProgressData
		
		ProgressData = ProgressFile
		
		if Len(ProgressData) > 0 then '--- There are some progress data
			if ProgressData = "DONE" then '--- Upload was done.
				ProgressFile.Done
				Err.Raise 1, "getForm", "Upload was done"
			else '--- if ProgressData = "DONE" Then 'Upload was done.
				'--- m_State & vbCrLf & TotalBytes & vbCrLf & BytesRead & vbCrLf & ReadTime
				ProgressData = Split (ProgressData, vbCrLf)
				if UBound(ProgressData) = 3 then
					m_State = CLng(ProgressData(0))
					TotalBytes = CLng(ProgressData(1))
					BytesRead = CLng(ProgressData(2))
					m_ReadTime = CLng(ProgressData(3))
				end if '--- if ubound(ProgressData) = 3 Then
			end if '--- if ProgressData = "DONE" Then 'Upload was done.
		end if '--- if len(ProgressData) > 0 then 'There are some progress data
		set getForm = Me
	End Function

	'--- This function writes progress info data to a temporary file.
	Private Sub WriteProgressInfo
		if UploadID > 0 then '--- Is the upload ID defined? (Upload is using progress)
			if IsEmpty(ProgressFile.UploadID) then '--- Was UploadID of ProgressFile set?
				ProgressFile.UploadID = UploadID
			end if

			dim ProgressData, FileName
			ProgressData = m_State & vbCrLf & TotalBytes & vbCrLf & BytesRead & vbCrLf & ReadTime
			ProgressFile.Contents = ProgressData
		end if
	End Sub

	'--- ASPForm Constructor 
	Private Sub Class_Initialize()
		ChunkReadSize = &H10000 '--- 64 kB
		SizeLimit = &H100000 '1MB
		BytesRead = 0
		m_State = xfsNone
		
		TotalBytes = Request.TotalBytes

		set ProgressFile = New cProgressFile
		set m_Items = New cFormFields
	End Sub

	'--- ASPForm Destructor
	Private Sub Class_Terminate()
		if UploadID > 0 then '--- Is the upload ID defined? (Upload is using progress)
			'--- We have to close info about upload.
			ProgressFile.Contents = "DONE"
		end if
	End Sub

	Private Function IsInSizeLimit(TotalBytes)
		IsInSizeLimit = (m_SizeLimit = 0 or m_SizeLimit > TotalBytes) and (MaxLicensedLimit > TotalBytes)
	End Function

	Public Property Get SizeLimit
		SizeLimit = m_SizeLimit
	End Property 

	'--- Pure - ASP upload is a free script, but with 10MB upload limit if you want to upload bigger files, please register ScriptUtilities and Huge-ASP upload at http://www.motobit.com/help/scptutl/lc2.htm
	Public Property Let SizeLimit(NewLimit)
	    if NewLimit > MaxLicensedLimit Then
		    Err.Raise 1, "ASPForm - SizeLimit", "This version of Pure-ASP upload is licensed with maximum limit of 10MB (" & MaxLicensedLimit & "B)"
		    m_SizeLimit = MaxLicensedLimit
	    else
		    m_SizeLimit = NewLimit
		end if
	End Property 

	Public Boundary
	private m_Items, m_State, m_SizeLimit '--- Defined form size limit.
	private bSourceData '--- ADODB.Stream
	private StartUploadTime, TempFiolder, ProgressFile '--- File with info about current progress
End Class '--- ASPForm

const MaxLicensedLimit = &HA00000

'************************************************************************
'--- Emulates ScriptUtilities FormFields object
'--- We must have such class because of multiselect fields. See http://www.motobit.com
Class cFormFields
	dim m_Keys(), m_Items(), m_Count
	
	Public Default Property Get Item(ByVal Key)
        if vartype(Key) = vbInteger or vartype(Key) = vbLong then
			'--- Numeric index
			if Key<1 or Key>m_Count then Err.Raise "Index out of bounds"
			set Item = m_Items(Key-1)
			Exit Property
		end if

		'--- wscript.echo "**Item", Key
		dim Count
		Count = ItemCount(Key)
		Key = LCase(Key)
		
		if Count>0 then
			if Count>1 then
				'--- More items. Get them All as an cFormFields
				dim OutItem, ItemCounter
				set OutItem = New cFormFields
				ItemCounter = 0
				
				for ItemCounter = 0 to Ubound(m_Keys)
					if LCase(m_Keys(ItemCounter)) = Key then OutItem.Add Key, m_Items(ItemCounter)
				next
				set Item = OutItem
				'--- wscript.echo "***Item-More", Key
			else 
				for ItemCounter = 0 to Ubound(m_Keys)
					if LCase(m_Keys(ItemCounter)) = Key then exit for
				Next

				if isobject (m_Items(ItemCounter)) then
					Set Item = m_Items(ItemCounter)
				else
					Item = m_Items(ItemCounter)
				end if
				'--- wscript.echo "***Item-One", Key
			end if
		else '--- No item 
			set Item = New cFormField
		end if
	End Property

	Public Property Get MultiItem(ByVal Key)
		'--- returns an array of items with the same Key
		dim Out: set Out = Newn cFormFields
		dim I, vItem, Count
		Count = ItemCount(Key)
		
		if Count = 1 then
			'--- one key - get it from Item
			Out.Add Key, Item(Key)
		elseif Count > 1 then
			'--- more keys - enumerate them using Items
			For Each I In Item(Key).Items
				Out.Add Key, I
			Next
		End If

		set MultiItem = Out
	End Property

	'--- For multiitem (I'm sorry, VBS does not support optional parameters for Item property)
	Public Property Get Value
		dim I, V
		for rach I in m_Items
			V = V & ", " & I 
		next
		V = Mid(V, 3)
		Value = V
	End Property

	Public Property Get xA_NewEnum
		Set xA_NewEnum = m_Items
	End Property

	Public Property Get Items()
		'--- Wscript.Echo "**cFormFields-Items"		
		Items = m_Items
	End Property

	Public Property Get Keys()
		Keys = m_Keys
	End Property

	Public Property Get Files
		dim cItem, OutItem, ItemCounter
		set OutItem = New cFormFields 
		ItemCounter = 0
		if m_Count > 0 then '--- Enumerate only non-empty form
			for ItemCounter = 0 To Ubound(m_Keys)
				set cItem = m_Items(ItemCounter)
				if cItem.IsFile then
					OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter)
				end if
			Next
		end if
		set Files = OutItem 
	End Property

	Public Property Get Texts
		dim cItem, OutItem, ItemCounter
		set OutItem = New cFormFields 
		ItemCounter = 0
		
		for ItemCounter = 0 to Ubound(m_Keys)
			set cItem = m_Items(ItemCounter)
			if Not cItem.IsFile then
				OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter)
			end if
		next
		set Texts = OutItem
	End Property

	Public Sub Save(Path)
		dim Item
		for each Item in m_Items
			if Item.isFile then
				Item.Save Path
			end if
		next
	End Sub

	'--- Count of dictionary items within specified key
	Public Property Get ItemCount(ByVal Key)
		'--- wscript.echo "ItemCount"
		dim cKey, Counter
		
		Counter = 0
		Key = LCase(Key)
		for each cKey in m_Keys
			'--- wscript.echo "ItemCount", "cKey"
			if LCase(cKey) = Key then Counter = Counter + 1
		next
		ItemCount = Counter
	End Property

	'--- Count of all dictionary items
	Public Property Get Count()
		Count = m_Count
	End Property

	Public Sub Add(byval Key, Item)
		Key = "" & Key
		redim preserve m_Items(m_Count)
		redim preserve m_Keys(m_Count)
		m_Keys(m_Count) = Key
		set m_Items(m_Count) = Item
		m_Count = m_Count + 1
	End Sub

	Private Sub Class_Initialize()
		dim vHelp()
		'--- I do not know why, but some of VBS verrsions declares m_Items and m_Keys as Empty, not as Variant() - see class variables. vHelp eliminates this problem. V. 2.03, 2.04
		on error resume next
		m_Items = vHelp
		m_Keys = vHelp
		m_Count = 0
	End Sub


	'********************************** mpSeparateFields **********************************
	'--- This method retrieves the upload fields from binary data. Binary is safearray ( VT_UI1 | VT_ARRAY ) of all multipart document raw binary data from input.
	Public Sub mpSeparateFields(Binary, ByVal Boundary)
		dim PosOpenBoundary, PosCloseBoundary, PosEndOfHeader, isLastBoundary

		Boundary = "--" & Boundary			
		Boundary = StringToBinary(Boundary)

		PosOpenBoundary = InStrB(Binary, Boundary)
		PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary, 0)

		do while (PosOpenBoundary > 0 and PosCloseBoundary > 0 and not isLastBoundary)
			'--- Header and file/source field data
			dim HeaderContent, bFieldContent
			
			'--- Header fields
			dim Content_Disposition, FormFieldName, SourceFileName, Content_Type
			'--- Helping variables
			dim TwoCharsAfterEndBoundary
			'--- Get end of header
			PosEndOfHeader = InStrB(PosOpenBoundary + Len(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))

			'--- Separates field header
			HeaderContent = MidB(Binary, PosOpenBoundary + LenB(Boundary) + 2, PosEndOfHeader - PosOpenBoundary - LenB(Boundary) - 2)
    
			'--- Separates field content
			bFieldContent = MidB(Binary, (PosEndOfHeader + 4), PosCloseBoundary - (PosEndOfHeader + 4) - 2)
			
			'---- Separates header fields from header
			GetHeadFields BinaryToString(HeaderContent), FormFieldName, SourceFileName, Content_Disposition, Content_Type

			'--- Create one field and assign parameters
			dim Field        '--- All field values.
			set Field = new cFormField

			Field.ByteArray = MultiByteToBinary(bFieldContent)

			Field.Name = FormFieldName
			Field.ContentDisposition = Content_Disposition
			if not isempty(SourceFileName) then
				Field.FilePath = SourceFileName
				Field.FileName = GetFileName(SourceFileName)
				Field.FileExt = GetFileExt(SourceFileName)
			else '--- if not isempty(SourceFileName) then
			end if '--- if not isempty(SourceFileName) then
			Field.ContentType = Content_Type
			
			Add FormFieldName, Field

			'--- Is this last boundary ?
			TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, PosCloseBoundary + LenB(Boundary), 2))
			isLastBoundary = TwoCharsAfterEndBoundary = "--"

			if not isLastBoundary then '--- This is not last boundary - go to next form field.
				PosOpenBoundary = PosCloseBoundary
				PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary)
			end if
		loop
	End Sub
End Class '--- cFormFields


'--- This class transfers data between primary (upload) and secondary (progress) window.
Class cProgressFile
	private fs
	Public TempFolder, m_UploadID, TempFileName

	Public Default Property Get Contents()
		Contents = GetFile(TempFileName)
	End Property

	Public Property Let Contents(inContents)
		WriteFile TempFileName, inContents
	End Property

	Public Sub Done 'Delete temporary file when upload was done.
		FS.DeleteFile TempFileName
	End Sub

	Public Property Get UploadID()
		UploadID = m_UploadID
	End Property

	Public Property Let UploadID(inUploadID)
		if isempty(FS) then Set fs = CreateObject("Scripting.FileSystemObject")
		TempFolder = fs.GetSpecialFolder(2)

		m_UploadID = inUploadID
		TempFileName = TempFolder & "\pu" & m_UploadID & ".~tmp"
		
		dim DateLastModified
		on error resume next
		DateLastModified = fs.GetFile(TempFileName).DateLastModified
		on error goto 0
		if isempty(DateLastModified) then '--- OK
		elseif Now-DateLastModified>1 Then 'I think upload duration will be less than one day
			FS.DeleteFile TempFileName
		end if
	End Property

	Private Function GetFile(Byref FileName)
		dim InStream
		on error resume next
		set InStream = fs.OpenTextFile(FileName, 1)
		GetFile = InStream.ReadAll
		on error goto 0
	End Function

	Private Function WriteFile(Byref FileName, Byref Contents)
		'--- wscript.echo "WriteFile", FileName, Contents
		dim OutStream
		on error resume next
		set OutStream = fs.OpenTextFile(FileName, 2, True)
		OutStream.Write Contents
	End Function

	Private Sub Class_Initialize()
	End Sub
End Class '--- cProgressFile


'--- ******************************************************************************
'-- Emulates ScriptUtilities FormField object. See http://www.motobit.com
Class cFormField
	'--- Used properties
	Public ContentDisposition, ContentType, FileName, FilePath, FileExt, Name, ByteArray

	'--- non-used properties.
	Public CharSet, HexString, InProgress, SourceLength, RAWHeader, Index, ContentTransferEncoding
 
	Public Default Property Get String()
		'--- wscript.echo "**Field-String", Name, LenB(ByteArray)
		String = BinaryToString(ByteArray)
	End Property 

	Public Property Get IsFile()
		IsFile = not isempty(FileName)
	End Property

	Public Property Get Length()
		Length = LenB(ByteArray)
	End Property

	Public Property Get Value()
		set Value = Me
	End Property

	Public Sub Save(Path)
	  	'--- 2.06 - and len(FileName)>0
		if IsFile and Len(FileName)>0 then
			dim fullFileName
			fullFileName = Path & "\" & FileName
			SaveAs fullFileName
		else
			'--- response.write "<br/>" & typename(Name)
			'--- Err.Raise 1, "Text field " & Name & " does not have a file name"
		end if
	End Sub

	Public Sub SaveAs(newFileName)
		'--- 2.06 - removed if len(ByteArray)>0 then 
		SaveBinaryData newFileName, ByteArray
	End Sub
End Class


Function StringToBinary(String)
  	dim I, B
  	for I=1 to Len(String)
    	B = B & ChrB(Asc(Mid(String,I,1)))
  	next
  	StringToBinary = B
End Function

Function BinaryToString(Binary)
  	'--- 2001 Antonin Foller, Motobit Software. Optimized version of PureASP conversion function. Selects the best algorithm to convert binary data to String data
  	dim TempString 

  	on error resume next
 	'--- Recordset conversion has a best functionality
  	TempString = RSBinaryToString(Binary)
  	if Len(TempString) <> LenB(Binary) then '---Conversion error
    	'--- We have to use multibyte version of BinaryToString
    	TempString = MBBinaryToString(Binary)
  	end if
  	BinaryToString = TempString
End Function

Function MBBinaryToString(Binary)
  	'--- 1999 Antonin Foller, Motobit Software. MultiByte version of BinaryToString function. Optimized version of simple BinaryToString algorithm.
  	dim cl1, cl2, cl3, pl1, pl2, pl3, L', nullchar
  	cl1 = 1
  	cl2 = 1
  	cl3 = 1
  	L = LenB(Binary)
  
  	do while cl1<=L
		pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		cl1 = cl1 + 1
		cl3 = cl3 + 1
		if cl3>300 then
			pl2 = pl2 & pl3
			pl3 = ""
			cl3 = 1
			cl2 = cl2 + 1
			if cl2>200 then
				pl1 = pl1 & pl2
				pl2 = ""
				cl2 = 1
			end if
		end if
  	loop
  	MBBinaryToString = pl1 & pl2 & pl3
End Function


Function RSBinaryToString(xBinary)
  	'--- 1999 Antonin Foller, Motobit Software. This function converts binary data (VT_UI1 | VT_ARRAY or MultiByte string) to string (BSTR) using ADO recordset
	'--- The fastest way - requires ADODB.Recordset. Use this function instead of MBBinaryToString if you have ADODB.Recordset installed to eliminate problem with PureASP performance

	dim Binary
	'--- MultiByte data must be converted to VT_UI1 | VT_ARRAY first.
	if vartype(xBinary) = 8 then Binary = MultiByteToBinary(xBinary) else Binary = xBinary
	
  	dim RS, LBinary
  	const adLongVarChar = 201
  	set RS = CreateObject("ADODB.Recordset")
  	LBinary = LenB(Binary)
	
	if LBinary>0 then
		RS.Fields.Append "mBinary", adLongVarChar, LBinary
		RS.Open
		RS.AddNew
		RS("mBinary").AppendChunk Binary 
		RS.Update
		RSBinaryToString = RS("mBinary")
	else
		RSBinaryToString = ""
	end if
End Function


Function MultiByteToBinary(MultiByte)
  	'--- This function converts multibyte string to real binary data (VT_UI1 | VT_ARRAY). Using recordset
 	dim RS, LMultiByte, Binary
  	const adLongVarBinary = 205
  	set RS = CreateObject("ADODB.Recordset")
  	LMultiByte = LenB(MultiByte)
	if LMultiByte>0 then
		RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
		RS.Open
		RS.AddNew
		RS("mBinary").AppendChunk MultiByte & ChrB(0)
		RS.Update
		Binary = RS("mBinary").GetChunk(LMultiByte)
	end if
  	MultiByteToBinary = Binary
End Function


'************** Upload Utilities. Separates header fields from upload header
Function GetHeadFields(ByVal Head, Name, FileName, Content_Disposition, Content_Type)
  	'--- Get name of the field. Name is separated by name= and ;
  	Name = (SeparateField(Head, "name=", ";")) '--- ltrim
  	'--- Remove quotes (if the field name is quoted)
  	if Left(Name, 1) = """" then Name = Mid(Name, 2, Len(Name) - 2)

  	'--- Same for source filename
 	FileName = (SeparateField(Head, "filename=", ";")) '--- ltrim
	if Left(FileName, 1) = """" then FileName = Mid(FileName, 2, Len(FileName) - 2)

  	'--- Separate content-disposition and content-type header fields
  	Content_Disposition = LTrim(SeparateField(Head, "content-disposition:", ";"))
  	Content_Type = LTrim(SeparateField(Head, "content-type:", ";"))
End Function

'--- Separates one field between sStart and sEnd
Function SeparateField(From, ByVal sStart, ByVal sEnd)
  	dim PosB, PosE, sFrom
  	sFrom = LCase(From)
  	PosB = InStr(sFrom, sStart)
  	if PosB > 0 then
    	PosB = PosB + Len(sStart)
    	PosE = InStr(PosB, sFrom, sEnd)
    	if PosE = 0 then PosE = InStr(PosB, sFrom, vbCrLf)
    	if PosE = 0 then PosE = Len(sFrom) + 1
    	SeparateField = Mid(From, PosB, PosE - PosB)
  	else
    	SeparateField = Empty
  	end if
End Function

Function SplitFileName(FullPath)
  	dim Pos, PosF
  	PosF = 0
  	for Pos = Len(FullPath) to 1 Step -1
    	select case Mid(FullPath, Pos, 1)
     		case ":", "/", "\": PosF = Pos + 1: Pos = 0
    	end select
  	next
  	if PosF = 0 then PosF = 1
	SplitFileName = PosF
End Function

Function GetPath(FullPath)
  	GetPath = left(FullPath, SplitFileName(FullPath)-1)
End Function

'--- Separetes file name from the full path of file
Function GetFileName(FullPath)
  	GetFileName = Mid(FullPath, SplitFileName(FullPath))
End Function

'--- Separetes file name from the full path of file
Function GetFileExt(FullPath)
	dim Pos: Pos = InStrRev(FullPath,".")
	if Pos>0 then GetFileExt = Mid(FullPath, Pos)
End Function

Function RecurseMKDir(ByVal Path)
  	dim FS: set FS = CreateObject("Scripting.FileSystemObject")
	
  	Path = Replace(Path, "/", "\")
  	if Right(Path, 1) <> "\" then Path = Path & "\"   '"
  	dim Pos, n
  	Pos = 0: n = 0
  	Pos = InStr(Pos + 1, Path, "\")   '"
  	do while Pos > 0
    	on error resume next
    	FS.CreateFolder Left(Path, Pos - 1)
    	if Err = 0 then n = n + 1
    	Pos = InStr(Pos + 1, Path, "\")   '"
  	loop
  	RecurseMKDir = n
End Function

Function SaveBinaryData(FileName, ByteArray)
	SaveBinaryData = SaveBinaryDataStream(FileName, ByteArray)
End Function

Function SaveBinaryDataTextStream(FileName, ByteArray)
  	dim FS : set FS = CreateObject("Scripting.FileSystemObject")
	on error resume next
  	dim TextStream 
	set TextStream = FS.CreateTextFile(FileName)
	if Err = &H4c then '--- Path not found.
		on error goto 0
		RecurseMKDir GetPath(FileName)
		on error resume next
		set TextStream = FS.CreateTextFile(FileName)
	end if
  	TextStream.Write BinaryToString(ByteArray) '--- BinaryToString is in upload.inc.
  	TextStream.Close

	dim ErrMessage, ErrNumber
	ErrMessage = Err.Description
	ErrNumber = Err

	on error goto 0
	if ErrNumber<>0 then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage 
End Function

Function SaveBinaryDataStream(FileName, ByteArray)
	dim BinaryStream
	set BinaryStream = CreateObject("ADODB.Stream")
	BinaryStream.Type = 1 '--- Binary
	BinaryStream.Open
	
	'--- 2.06 - zero byte file is legal
	if LenB(ByteArray)>0 then BinaryStream.Write ByteArray
	on error resume next
	
	BinaryStream.SaveToFile FileName, 2 '--- Overwrite

	if Err = &Hbbc then '--- Path not found.
		on error goto 0
		RecurseMKDir GetPath(FileName)
		on error resume next
		BinaryStream.SaveToFile FileName, 2 '--- Overwrite
	end if
	dim ErrMessage, ErrNumber
	
	ErrMessage = Err.Description
	ErrNumber = Err

	on error goto 0
	if ErrNumber<>0 then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage 
End Function
'--- ************** Upload Utilities - end

'--- Emulates response object
Class cResponse
	Public Property Get IsClientConnected
		Randomize
		IsClientConnected = CBool(CLng(Rnd * 4))
		IsClientConnected = True
	End Property 
End Class 

Class cRequest
	Private Readed, BinaryStream

	Public Function ServerVariables(Name)	
		select case UCase(Name) 
			case "CONTENT_TYPE": 
			case "HTTP_CONTENT_TYPE": 
				ServerVariables = "multipart/form-data; boundary=---------------------------7d21960404e2"
			case "CONTENT_LENGTH": 
			case "HTTP_CONTENT_LENGTH": 
				ServerVariables = "" & TotalBytes
			case "REQUEST_METHOD": 
				ServerVariables = "POST"
		end select
	End Function

	Public Function BinaryRead(ByRef Bytes)
		if Bytes<=0 then Exit Function

		if Readed + Bytes > TotalBytes Then Bytes = TotalBytes - Readed
		BinaryRead = BinaryStream.Read(Bytes)
	End Function

	Public Property Get TotalBytes
		TotalBytes = BinaryStream.Size
	End Property

	Private Sub Class_Initialize()
		set BinaryStream = CreateObject("ADODB.Stream")
		BinaryStream.Type = 1 '--- Binary
		BinaryStream.Open
		BinaryStream.LoadFromFile "F:\InetPub\Motobit\pureupload\2.txt"
		BinaryStream.Position = 0
		Readed = 0
	End Sub
End Class
%>
