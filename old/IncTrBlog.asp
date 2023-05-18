<%
'---- NameFromMonth() devuelve el nombre del mes [vMonth]
Function NameFromMonth(vMonth)
	select case cint(vMonth)
		case 1
			NameFromMonth = MesgS("Ene","Jan")
		case 2
			NameFromMonth = "Feb"
		case 3
			NameFromMonth = "Mar"
		case 4
			NameFromMonth = MesgS("Abr","Apr")
		case 5
			NameFromMonth = "May"
		case 6
			NameFromMonth = "Jun"
		case 7
			NameFromMonth = "Jul"
		case 8
			NameFromMonth = MesgS("Ago","Aug")
		case 9 
			NameFromMonth = "Sep"
		case 10
			NameFromMonth = "Oct"
		case 11
			NameFromMonth = "Nov"
		case 12
			NameFromMonth = MesgS("Dic","Dec")
	end select
End Function

'---- LastDay() devuelve el último dia del mes [testMonth] del año [testYear]
Function LastDay(testYear, testMonth)
    LastDay = Day(DateSerial(testYear, testMonth + 1, 0))
End Function

'---- GetPrevMonth() devuelve el mes anterior a [iThisMonth] de [iThisYear]
Function GetPrevMonth(iThisMonth,iThisYear)
	GetPrevMonth=Month(DateSerial(iThisYear,iThisMonth,1)-1)
End Function

'---- GetPrevMonth() devuelve el año anterior a [iThisMonth] de [iThisYear]
Function GetPrevMonthYear(iThisMonth,iThisYear)
	GetPrevMonthYear=Year(DateSerial(iThisYear,iThisMonth,1)-1)
End Function

'---- GetNextMonth() devuelve el mes siguiente a [iThisMonth] de [iThisYear]
Function GetNextMonth(iThisMonth,iThisYear)
	GetNextMonth=Month(DateSerial(iThisYear,iThisMonth+1,1))
End Function

'---- GetNextMonth() devuelve el año siguiente a [iThisMonth] de [iThisYear]
Function GetNextMonthYear(iThisMonth,iThisYear)
	GetNextMonthYear=Year(DateSerial(iThisYear,iThisMonth+1,1))
End Function

'---- DisplaySmallCalendar() muestra un calendario del mes llamado [sMonth] número [numMes] del año [iYear]
Sub DisplaySmallCalendar(sMonth, numMes, iYear)
	dim dDay, iLastDay, iFirstDay, i, j, k, sDate, sEvent
%>	
    <div class=row>
	    <a title='<%=MesgS("Mes Anterior","Previous Month")%>' href='TrBlogFront.asp?mes=<%=GetPrevMonth(numMes,iYear)%>&ano=<%=GetPrevMonthYear(numMes,iYear)%>'>
		<img alt='Previo' src=images/FlechaPrevio.gif /><img src=images/FlechaPrevio.gif /></a></div>
	<div class=row><%=UCase(sMonth)%> <%=iYear%></div>
	<div class="row">
	    <a title='<%=MesgS("Mes Siguiente","Following Month")%>' href='TrBlogFront.asp?mes=<%=GetNextMonth(numMes,iYear)%>&ano=<%=GetNextMonthYear(numMes,iYear)%>'>
		<img alt='Proximo' src=images/FlechaProx.gif /><img alt='Proximo' src=images/FlechaProx.gif /></a></div>
	<div>
	    <div><%=MesgS("Dom","Sun")%></div><div><%=MesgS("Lun","Mon")%></div><div><%=MesgS("Mar","Tue")%></div><div><%=MesgS("Mie","Wed")%></div>
		<div><%=MesgS("Jue","Thu")%></div><div><%=MesgS("Vie","Fri")%></div><div><%=MesgS("Sab","Sat")%></div></div>
	<div class=hrgreen><img alt='Blanco' src=images/spacer.gif /></div>
<%
	i = 0
	dDay=DateSerial(iYear,numMes,1)
	numMes = Month(dDay)
	iYear = Year(dDay)
	iLastDay = LastDay(iYear, numMes)
	iFirstDay=  Weekday(dDay)
	iLastDay = iFirstDay + iLastDay -1

	do while i<= iLastDay
		if i <> iLastDay then
%>
	<div>
<%
		else
			exit do
		end if
		
		for j=1 to 7
			if (j < iFirstDay and i = 0) or (i + j > iLastDay) then
%>
		<div>&nbsp;</div>
<%
			else
				k =  k + 1
				sDate = "'" & numMes & "/" & k & "/" & iYear & "'"
				sQuery = "SELECT * FROM BlogDetalle WHERE data=" & sDate
				if rst.State <> 0 then rst.close
				rst.Open sQuery, sProvider
				if rst.Eof then
				    if (k = Day(now)) and (numMes = Month(now)) and (iYear = Year(now)) then sEvent = "<div class=giornocorrente>" & k & " </div>" else sEvent =  k 
				else 
				    if (k = Day(now)) and (numMes = Month(now)) and (iYear = Year(now)) then
						sEvent = "<div class=back><a href='TrBlogFront.asp?dia=" & k & "&mes=" & numMes & "&ano=" & iYear & "' title='" & MesgS("Mostrar Blog","Show Blog") & "'>" & k & "</a></div>"
					else
						sEvent = "<div class=back2><a href='TrBlogFront.asp?dia=" & k & "&mes=" & numMes & "&ano=" & iYear & "' title='Mostrar Blog'>" & k & "</a></div>"
					end if
				end if	
%>
		<div><%=sEvent%></div>
<%
			end if								
		next
		i=i+7
	loop
%>
	    </div>
<%
End Sub

'---- Anteprima() muestra las primeras [nParole] palabras del texto [sText]
Function Anteprima(sText, nParole)
	dim nTemp, nVolte

	'--- Eliminiamo gli eventuiali caratteri di CR ed LF
	sText = Replace(sText, vbCrLf, "")

	'--- Cerca la fine della prima parola
	nTemp = InStr(sText, " ")

	if nTemp <> 0 then
        nVolte = 1
        '--- Finche` non abbiamo finito le parole o abbiamo raggiunto quelle massime
        while nTemp <> 0 and nVolte < nParole 
            '--- Incrementiamo il numero delle parole trovate
            nVolte = nVolte + 1

            '--- Cerchiamo la fine della parola successiva
            nTemp = InStr(nTemp + 1, sText, " ")
        wend
    end if

    '--- Se abbiamo trovato qualche parola
    if nVolte > 0 then
        '--- Se La variabile nTemp > 0 allora significa che abbiamo trovato le n parole che ci serivivano
        if nTemp > 0 then
            '--- Le stampiamo insieme ai puntini
            Anteprima = Mid(sText, 1, nTemp - 1) & " ..."
        else
            '--- Altrimenti abbiamo trovato meno delle n parole. Stampiamo la frase intera
            Anteprima = sText
        end if
    else
        '--- una sola parola
        if Len(sText) > 0 then
            Anteprima = sText
        else
            '--- La frase passata ha lunghezza 0
            Anteprima = "" 
        end If
    end If
End Function

'---- removeAllTags() escapa los caracteres especiales de [strInputEntry]
private Function removeAllTags(ByVal strInputEntry)
	strInputEntry = Replace(strInputEntry, "&", "&amp;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<", "&lt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ">", "&gt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "#", "&#035;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "%", "&#037;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "*", "&#042;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "\", "&#092;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "'", "&#146;", 1, -1, 1)
	removeAllTags = strInputEntry	
end Function

'---- Objeto para subir Ficheros del Blog
Class FileUploader
	Public Files, maxSize, maxFileSize, fileExt, error, errorDesc
	private mcolFormElem
    	
    '---- Crea Files y mcolFormElem
	private Sub Class_Initialize()
		set Files = Server.CreateObject("Scripting.Dictionary")
		set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
    '---- Libera Files y mcolFormElem
    private Sub Class_Terminate()
		if IsObject(Files) then
			Files.RemoveAll()
			set Files = Nothing
		end if
		if IsObject(mcolFormElem) then
			mcolFormElem.RemoveAll()
			set mcolFormElem = Nothing
		end if
	End Sub

	'---- Devuelve el contenido del campo [sIndex]
    Public Property Get Form(sIndex)
		Form = ""
		if mcolFormElem.Exists(LCase(sIndex)) then Form = mcolFormElem.Item(LCase(sIndex))
	End Property
	
	'---- Lee y almacena en el Blog los contenidos del Form
    Public Default Sub Upload()
		dim biData, sInputName, nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos, nPosFile, nPosBound, counter, sFileName1

		error = false
		errorDesc = ""
		
		if not IsNull(maxSize) and not IsNumeric(maxSize) then
			if Request.TotalBytes > maxSize then
				Error = true
				errorDesc = errorDesc & MesgS("El fichero es más grande que ","The file is bigger than ") & maxSize & " byte" 
				Exit Sub
			end if
		end if

		biData = Request.BinaryRead(Request.TotalBytes)
		nPosBegin = 1
		nPosEnd = InStrB(nPosBegin, biData, CByteString(Chr(13)))
		if (nPosEnd-nPosBegin) <= 0 then Exit Sub
		vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
		nDataBoundPos = InstrB(1, biData, vDataBounds)
		counter = 0
		
		do until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))
		    counter = counter + 1	
			nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InstrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InstrB(nPosEnd, biData, vDataBounds)
			
			if nPosFile <> 0 and nPosFile < nPosBound then
			    dim oUploadFile, sFileName, sFileExt

				set oUploadFile = New UploadedFile			
				nPosBegin = nPosFile + 10
				nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))				
				sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				sFileName1 = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))
				
				sfileExt = Right(sFileName1, Len(sFileName1) - InStrRev(sFileName1, "."))
				if not IsNull(fileExt) then
				    if Instr(fileExt, sFileExt) = 0 then
				  	    Error = true
				  	    errorDesc = errorDesc & "File " & oUploadFile.FileName  & ":" & MesgS("Tipo de fichero no permitido (solo ","The file type is not allowed (only ") & fileExt & ")"
				    else
    				    nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
    				    nPosBegin = nPos + 14
    				    nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
    				    oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
    				    nPosBegin = nPosEnd+4
    				    nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
    				    oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
    				
    				    if (oUploadFile.FileSize > 0) and (oUploadFile.FileSize <= maxFileSize) then Files.Add LCase(sInputName), oUploadFile
    
    				    if not IsNull(maxFileSize) then
      				        if oUploadFile.FileSize > maxFileSize then
    	 			  	        Error = true
    				  	        errorDesc = errorDesc & Mesg("Fichero ","File ") & oUploadFile.FileName & ": " & Mesg("El fichero es demasiado grande (max: 50 kb)","The file is too large (max: 50kb)") 
    				  	        '--- Exit Sub
    				        end if				
    				    end if
				    end if
				end if
			else
				nPos = InstrB(nPos, biData, CByteString(Chr(13)))
				nPosBegin = nPos + 4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				if not mcolFormElem.Exists(LCase(sInputName)) then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			end if

			nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
		loop
	End Sub

	'--- String to byte string conversion
	private Function CByteString(sString)
		dim nIndex

		for nIndex = 1 to Len(sString)
		   CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
		next
	end Function

	'--- Byte string to string conversion
	private Function CWideString(bsString)
		dim nIndex

		CWideString =""
		for nIndex = 1 to LenB(bsString)
		   CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
		next
	end Function
End Class

'--- UploadedFile soporta los ficheros subidos de Blog
Class UploadedFile
	Public ContentType, FileName, FileData
	
    '--- Tamaño del fichero
	Public Property Get FileSize()
		FileSize = LenB(FileData)
	End Property

    '--- SaveToDisk vuelca lo subido al fichero sPath
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
		
		set oFS = Nothing
	End Sub
	
    '--- SaveToDatabase graba el contenido binario en un campo oField
	Public Sub SaveToDatabase(ByRef oField)
		if LenB(FileData) = 0 then Exit Sub
		if IsObject(oField) then oField.AppendChunk FileData
	End Sub
End Class

'--- IsImage devuelve si strImageName es una imagen
Function IsImage(strImageName)
	dim objImageSize
	
    set objImageSize = New ImageSize
    with objImageSize
        .ImageFile = Server.MapPath(strImageName)
        if .IsImage then IsImage = true else IsImage = false
    end with
    set objImageSize = nothing
End Function

'--- IsImage alberga caracteristicas de una imagen
Class ImageSize
    Public Dimensioni(2), bolIsImage, strPathFile, strImageType
  
    private Sub Class_Initialize()
    end Sub
  
    private Sub Class_Terminate()
    end Sub

    '--- ImageName devuelve el nombre del fichero
    Public Property Get ImageName
  	    ImageName = doName(ImageFile)
    End Property
  
    '--- ImageFile devuelve el path del fichero
    Public Property Get ImageFile
  	    ImageFile = strPathFile
    End Property
  
    '--- ImageFile asigna el path al fichero
    Public Property Let ImageFile(strImageFile)
  	    strPathFile = strImageFile
  	    gfxSpex(strPathFile)
    End Property
  
    '--- ImageWidth devuelve la anchura de la imagen
    Public Property Get ImageWidth
  	    ImageWidth = Dimensioni(1)
    End Property
  
    '--- ImageHeight devuelve la altura de la imagen
    Public Property Get ImageHeight
  	    ImageHeight = Dimensioni(0)
    End Property
  
    '--- ImageDepth devuelve la profundidad de la imagen
    Public Property Get ImageDepth
      	ImageDepth = Dimensioni(2)
    End Property 
  
    '--- ImageDepth devuelve la extension de la imagen
    Public Property Get ImageType
  	    ImageType = strImageType
    End Property      
  
    '--- IsImage devuelve si el fichero es una imagen
    Public Property Get IsImage
  	    IsImage = bolIsImage
    End Property     

    '--- GetBytes lee y devuelve el stream de bytes
    Public Function GetBytes(flnm, offset, bytes)
        dim objFSO, objFTemp, objTextStream, lngSize
    
        on error resume next
        set objFSO = CreateObject("Scripting.FileSystemObject")
    
        '--- First, we get the filesize
        set objFTemp = objFSO.GetFile(flnm)
        lngSize = objFTemp.Size
        set objFTemp = Nothing
    
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

    '--- lngConvert convierte un ASCII codificado (bigendian) a Long
    Public Function lngConvert(strTemp)
        lngConvert = CLng(Asc(Left(strTemp, 1)) + ((Asc(Right(strTemp, 1)) * 256)))
    End Function

    '--- lngConvert2 convierte un ASCII codificado (littleendian) a Long
    Public Function lngConvert2(strTemp)
        lngConvert2 = CLng(Asc(Right(strTemp, 1)) + ((Asc(Left(strTemp, 1)) * 256)))
    End Function
  
    '--- gfxSpex clasifica el fichero según su tipo
    Public Sub gfxSpex(flnm)
        dim strPNG, strBuff, lngSize, flgFound, strTarget, strGIF, strBMP, strType
     
        strType = ""   
        strImageType = "(unknown)"
        bolIsImage = False
        strPNG = Chr(137) & Chr(80) & Chr(78)
        strGIF = "GIF"
        strBMP = Chr(66) & Chr(77)
        strType = GetBytes(flnm, 0, 3)

        if strType = strGIF then				    '--- is GIF
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
        elseif strType = strPNG then			    '--- is PNG
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

    '--- doName strips the file name from path
    Public Function doName(strPath)
        dim arrSplit
        arrSplit = Split(strPath, "\")
        doName = arrSplit(UBound(arrSplit))
    End Function
End Class

'--- removemaligno escapes the weird characters
Private Function removemaligno(ByVal strInput)
    strInput = Replace(strInput, "&", "&amp;", 1, -1, 1)
	strInput = Replace(strInput, "#", "&#035;", 1, -1, 1)
	strInput = Replace(strInput, "%", "&#037;", 1, -1, 1)
	strInput = Replace(strInput, "*", "&#042;", 1, -1, 1)
	strInput = Replace(strInput, "\", "&#092;", 1, -1, 1)
	strInput = Replace(strInput, "'", "&#146;", 1, -1, 1)
	
	removemaligno = strInput
End Function

'--- FormBlog in the main form
Sub FormBlog
    if strMode="new" then 
%>
<script type=text/javascript>
<!-- 
//funzione per  aggiungere i tag grassetto, sottolineato e corsivo
function AddMessageCode(code, promptText, InsertText) {
	if (code != "") {
	    insertCode = prompt(promptText + "\n[" + code + "]xxx[/" + code + "]", InsertText);
		if ((insertCode != null) && (insertCode != "")){document.formblog.blog_testo.value += "[" + code + "]" + insertCode + "[/" + code + "]";}
	}
	document.formblog.blog_testo.focus();
}

// Funzione per modificare carattere, colore e dimensione del testo
function FontCode(code) {
	if (code != "") {
	    insertCode = prompt("<%=MesgS("Teclea el texto en ","Type in the text in ")%> + code + "\n[" + code + "]xxx[/font]", '');
		if ((insertCode != null) && (insertCode != "")){document.formblog.blog_testo.value += "[" + code + "]" + insertCode + "[/font]";}
	}	
	document.formblog.blog_testo.focus();
}

//Funzione per aggiungere url, email, immagine ed elenco
function AddCode(code) {
	//per url
	if ((code != "") && (code == "URL")) {
	    insertText = prompt("<%=MesgS("Teclea el texto que se verá en el enlace","Type in the text to show in the link")%>", "");	
		if ((insertText != null) && (insertText != "") && (code == "URL")){
		    insertCode = prompt("<%=MesgS("Teclea la dirección a la que quieres apuntar","Type in the target URL address")%>", "http://");		
			if ((insertCode != null) && (insertCode != "") && (insertCode != "http://")){document.formblog.blog_testo.value += "[" + code + "=" + insertCode + "]" + insertText + "[/" + code + "]";}}
	}
	
	//per email
	if ((code != "") && (code == "EMAIL")) {
		insertText = prompt("<%=MesgS("Teclea el texto con el que se verá la dirección de e-mail","Type in the text to show with the e-mail address")%>", "");	
	    if ((insertText != null) && (insertText != "")){
			insertCode = prompt("<%=MesgS("Teclea la dirección de e-mail","Type in the e-mail address")%>", "");
			if ((insertCode != null) && (insertCode != "")){document.formblog.blog_testo.value += "[" + code + "=" + insertCode + "]" + insertText + "[/" + code + "]";}
			}
	}
	
	//per immagine
	if ((code != "") && (code == "IMG")) {	
		insertCode = prompt("<%=MesgS("Teclea la ruta de la imagen","Type in the image URL address")%>", "http://");
		if ((insertCode != null) && (insertCode != "") && (insertCode != "http://")) {document.formblog.blog_testo.value += "[" + code + "]" + insertCode + "[/" + code + "]";}			
	}
	
	//per elenco
	if ((code != "") && (code == "LIST")) {
		listType = prompt("<%=MesgS("Tipo de lista\nteclea \'1\' para números o vacío para puntos","Type of list\ntype in \'1\' for numbers or empty for bullets")%>", "");
			
		while ((listType != null) && (listType != "") && (listType != "1")) {
			listType = prompt("<%=MesgS("Error. Por favor teclea \'1\' para números o vacío para puntos","Error. Please type of list\ntype in \'1\' for numbers or empty for bullets")%>","");               
			}
			
			if (listType != null) {			
				var listItem = "1"; var insertCode = "";
				
				while ((listItem != "") && (listItem != null)) {
					listItem = prompt("<%=MesgS("Teclea el texto de la lista o nada para acabarla","Type in the text of the list or empty for ending it")%>",""); 
					if (listItem != "") {insertCode += "[LI]" + listItem + "[/LI]";}                   
				} 
				
				if (listType == "") {document.formblog.blog_testo.value += "[" + code + "]" + insertCode + "[/" + code + "]";} 
				else {document.formblog.blog_testo.value += "[" + code + "=" + listType + "]" + insertCode + "[/" + code + "=" + listType + "]";} 
			}
	    }

	    document.formblog.blog_testo.focus();
    }
//-->
</script>
<%      else %>
<script type=text/javascript>
<!-- 
//funzione per  aggiungere i tag grassetto, sottolineato e corsivo
function AddMessageCode(code, promptText, InsertText) {
	if (code != "") {
	    insertCode = prompt(promptText + "\n<" + code + ">xxx</" + code + ">", InsertText);
		if ((insertCode != null) && (insertCode != "")){document.formblog.blog_testo.value += "<" + code + ">" + insertCode + "</" + code + ">";}
	}
	document.formblog.blog_testo.focus();
}

// Funzione per modificare carattere, colore e dimensione del testo
function FontCode(code) {
	if (code != "") {
		insertCode = prompt("<%=MesgS("Teclea el texto en ","Type in the text in ")%>" + code + '\n<' + code + '>xxx</font>', '');
		if ((insertCode != null) && (insertCode != "")){document.formblog.blog_testo.value += '<' + code + '>' + insertCode + '</font>';}
	}
	document.formblog.blog_testo.focus();
}

//Funzione per aggiungere url, email, immagine ed elenco
function AddCode(code) {
	//per url
	if ((code != "") && (code == "URL")) {
		insertText = prompt("<%=MesgS("Teclea el texto que se verá en el enlace","Type in the text to show with the link")%>", "");
			
		if ((insertText != null) && (insertText != "") && (code == "URL")) {
			insertCode = prompt("<%=MesgS("Teclea la dirección a la que quieres apuntar","Type in the target address")%>", "http://");
		    if ((insertCode != null) && (insertCode != "") && (insertCode != "http://")){document.formblog.blog_testo.value += '<a href="' + insertCode + '" target="_blank">' + insertText + '</a>';}
	    }
    }
	
	//per email
	if ((code != "") && (code == "EMAIL")) {
		insertText = prompt("<%=MesgS("Teclea el texto con el que se verá la dirección de e-mail","Type in the text to show with the e-mail address")%>", "");

		if ((insertText != null) && (insertText != "")) {
			insertCode = prompt("<%=MesgS("Teclea la dirección de e-mail","Type in the e-mail address")%>", "");				
			if ((insertCode != null) && (insertCode != "")) {document.formblog.blog_testo.value += '<a href="mailto:' + insertCode + '">' + insertText + '</a>';}
		}
	}
	
	//per immagine
	if ((code != "") && (code == "IMG")) {	
		insertCode = prompt("<%=MesgS("Teclea la ruta de la imagen","Type in the path to the image")%>", "http://");
		if ((insertCode != null) && (insertCode != "") && (insertCode != "http://")) {document.formblog.blog_testo.value += '<img src="' + insertCode + '" border="0">';}			
	}
	
	//per elenco
	if ((code != "") && (code == "LIST")) {
		listType = prompt("<%=MesgS("Tipo de lista\nteclea \'1\' para números o vacío para puntos","Type of list\ntype \'1\' for numbers or empty for bullets")%>", "");
			
		while ((listType != null) && (listType != "") && (listType != "1")) {
			listType = prompt("<%=MesgS("Error. Por favor teclea \'1\' para números o vacío para puntos","Error. Please type \'1\' for numbers or empty for bullets")%>","");               
		}
			
		if (listType != null) {			
			var listItem = "1"; var insertCode = "";
				
			while ((listItem != "") && (listItem != null)) {
				listItem = prompt("<%=MesgS("Teclea el texto de la lista o nada para acabarla","Type in the text of the list or empty for ending it")%>",""); 
				if (listItem != "") {insertCode += "<li>" + listItem + "</li>";}                   
			} 
				
			if (listType == "") {document.formblog.blog_testo.value += "<ul>" + insertCode + "</ul>";} else {document.formblog.blog_testo.value += "<ol>" + insertCode + "</ol>";} 		
		}
	}
	document.formblog.blog_testo.focus();
}
//-->
</script>
<%      end if %>
<script type=text/javascript>
<!-- 
function CheckForm() {
    if (document.formblog.blog_autore.value=="") {alert("<%=MesgS("Por favor teclea tu nombre","Please type in your name")%>"); return false;}
    if (document.formblog.email.value!="") {
		if (document.formblog.email.value.indexOf("@")==-1) {alert("<%=MesgS("Dirección de e-mail incorrecta","Wrong e-mail address")%>"); return false;}	
	    if (document.formblog.email.value.indexOf(".")==-1) {alert("<%=MesgS("Dirección de e-mail incorrecta","Wrong e-mail address")%>"); return false;}	
	}
    if (document.formblog.blog_titolo.value=="") {alert("<%=MesgS("Por favor teclea el título","Please type in the title")%>"); return false;}
    if (document.formblog.blog_testo.value=="") {alert("<%=MesgS("Por favor teclea el texto","Please type in the text")%>"); return false;}
    return true;
}

// preview blog
function OpenPreviewWindow(){
	Autore = escape("<%=Session("Usuario")%>"); Titolo = escape(document.formblog.blog_titolo.value); Testo = escape(document.formblog.blog_testo.value); Modo = escape(document.formblog.strMode.value);
	Data = escape(document.formblog.data.value); Ora = escape(document.formblog.ora.value); 
	document.cookie = "Blog_Autore=" + Autore 
	document.cookie = "Blog_Titolo=" + Titolo
   	document.cookie = "Blog_Testo=" + Testo
   	document.cookie = "Blog_Modo=" + Modo
   	document.cookie = "Blog_Data=" + Data
   	document.cookie = "Blog_Ora=" + Ora
	
   	openWin('TrBlogPreview.asp','preview','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=500,height=300')
}

function openWin(theURL,winName,features) {window.open(theURL,winName,features);}

// rollover 	
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
    if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
    else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_swapImgRestore() { //v3.0
    var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
    
function MM_findObj(n, d) { //v4.0
    var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
    if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
    for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
    if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
    var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
    if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<form method=post name=formblog action=old/TrBlogGraba.asp onsubmit='return CheckForm();' onreset="return confirm('<%=MesgS("¿Seguro que quieres limpiar los campos?","Sure you want to wipe all the fields?")%>');">
<div class=tablemenu>
    <div class=hrarancio><img src=../images/spacer.gif /></div><div>* <%=MesgS("son campos obligatorios","are required fields")%></div>
    <div><%=MesgS("Título","Title")%> (*) :</div><div> <input title='<%=MesgS("Título del Blog","Blog Title")%>' type=text name=blog_titolo size=25 maxlength=80 class=form value="<%= blog_titolo %>"/> </div>
    <div> 
        <div><%=MesgS("Formato ","Format ")%>:</div>
        <div>
	  	    <select title='<%=MesgS("Elige el aspecto de la letra","Choose the font face")%>' name=selectFont 
                onchange='FontCode(selectFont.options[selectFont.selectedIndex].value);document.formblog.selectFont.options[0].selected = true;' class=form>
		        <option selected>-- <%=MesgS("Tipo de Letra","Font Face")%> --</option><option value='font face=Arial'>Arial</option><option value='font face=Courier'>Courier New</option>
		        <option value='font face=Times'>Times New Roman</option><option value='font face=Verdana'>Verdana</option></select>
	        <select name=selectColour onchange='FontCode(selectColour.options[selectColour.selectedIndex].value);document.formblog.selectColour.options[0].selected = true;' class=form>
                <option value=0 selected>-- Color --</option>
                <option value='font color=black'><%=MesgS("Negro","Black")%></option><option value='font color=white'><%=MesgS("Blanco","White")%></option>
                <option value='font color=blue'><%=MesgS("Azul","Blue")%></option><option value='font color=red'><%=MesgS("Rojo","Red")%></option>
                <option value='font color=orange'><%=MesgS("Naranja","Orange")%></option><option value='font color=green'><%=MesgS("Verde","Green")%></option>
                <option value='font color=yellow'><%=MesgS("Amarillo","Yellow")%></option><option value='font color=gray'><%=MesgS("Gris","Grey")%></option></select> 
		    <select	name=selectSize onchange='FontCode(selectSize.options[selectSize.selectedIndex].value);document.formblog.selectSize.options[0].selected = true;' class=form>
                <option selected>-- Tamaño --</option><option value='font size=1'>1</option><option value='font size=2'>2</option><option value='font size=3'>3</option>
		        <option value='font size=4'>4</option><option value='font size=5'>5</option><option value='font size=6'>6</option></select></div></div>
    <div> 
        <a href="javascript:AddMessageCode('B','<%=MesgS("Teclea el texto en Negrita","Type in Text in Bold")%>', '')" onmouseout='MM_swapImgRestore()' onmouseover="MM_swapImage('Image1','','images/BlogNegrita_O.gif',1)">
            <img id=Image1 border=0 src=images/BlogNegrita.gif alt='<%=MesgS("Negrita","Bold")%>'/></a> 
        <a href="javascript:AddMessageCode('I','<%=MesgS("Teclea el texto en Cursiva","Type in Text in Italics")%>', '')" onmouseout='MM_swapImgRestore()' 
            onmouseover="MM_swapImage('Image2','','images/BlogCursiva_O.gif',1)"><img id=Image2 border=0 src=images/BlogCursiva.gif alt='<%=MesgS("Cursiva","Italics")%>'/></a> 
        <a href="javascript:AddMessageCode('U','<%=MesgS("Teclea el texto Subrayado","Type in Text to Underline")%>', '')" onmouseout='MM_swapImgRestore()' 
            onmouseover="MM_swapImage('Image3','','images/BlogSubrayar_O.gif',1)">
            <img id=Image3 border=0 src=images/BlogSubrayar.gif alt='<%=MesgS("Subrayado","Underline")%>'/></a> 
        <a href="javascript:AddCode('URL')" onmouseout='MM_swapImgRestore()' onmouseover="MM_swapImage('Image5','','images/BlogLink_O.gif',1)">
            <img border=0 src=images/BlogLink.gif alt='<%=MesgS("Añadir Link a otra Web","Add link to another site")%>'/></a> 
        <a href="javascript:AddCode('EMAIL')" onmouseout='MM_swapImgRestore()' onmouseover="MM_swapImage('Image9','','images/BlogEmail_O.gif',1)">
            <img border=0 src=old/BlogEmail.gif alt='<%=MesgS("Añadir Link a un e-mail","Add link to e-mail address")%>'/></a> 
        <a href="javascript:AddMessageCode('CENTER','<%=MesgS("Teclea el texto Centrado","Type in Text to Center")%>', '')" onmouseout='MM_swapImgRestore()' 
            onmouseover="MM_swapImage('Image6','','images/BlogCentra_O.gif',1)"><img border=0 src=images/BlogCentra.gif alt=Centrado /></a> 
        <a href="javascript:AddCode('LIST')" onmouseout='MM_swapImgRestore()' onmouseover="MM_swapImage('Image8','','images/BlogLista_O.gif',1)"><img border=0 src=images/BlogLista.gif alt='<%=MesgS("Lista","List")%>'/></a> 
        <a href="javascript:AddCode('IMG')" onmouseout='MM_swapImgRestore()' onmouseover="MM_swapImage('Image7','','images/BlogImage_O.gif',1)"><img border=0 src=images/BlogImage.gif alt='<%=MesgS("Imagen","Image")%>'/></a> 
        <a href="javascript:openWin('TrBlogUpload.asp?strMode=<%=strMode%>','images','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=400,height=330')" 
            onmouseout='MM_swapImgRestore()' onmouseover="MM_swapImage('Image12','','images/BlogImage_O.gif',1)">
            <img src=images/BlogUpload.gif alt='<%=MesgS("Subir imagen","Upload Picture")%>' border=0 /></a></div></div>
<div><div><%=MesgS("Texto ","Text ")%>(*) :</div><div> <textarea name=blog_testo cols=55 rows=10 class=form><%= blog_testo %></textarea> </div></div>
<div> <input name=comments type=checkbox value=True <% if blnComments = true then Response.Write(" checked") %> /><%=MesgS("Permitir que los usuarios dejen comentarios","Allow users to comment")%></div>
<div> 
    <div> 
        <input type=hidden name=strMode value='<%= strMode %>'/> <input type=hidden name=blog_id value='<%= CLng(Request.QueryString("blog_id")) %>'/> 
        <input type=hidden name=data value='<%= data %>'/><input type=hidden name=ora value='<%= ora %>'/><input type=hidden name=mese value='<%= mese %>'/> 
        <input type=hidden name=anno value='<%= anno %>'/></div>
    <div>
        <input type=button name=Preview value='<%=MesgS("Vista Previa","Preview")%>' onclick='OpenPreviewWindow()' class=form /> 
<%      if strMode="edit" then %>
        <input type=hidden name=page value='<%= page %>'/> <input type=hidden name=Block value='<%= Block %>'/> <input type=submit name=Submit value='<%=MesgS("Editar Blog","Edit Blog")%>' class=form /> 
<%      else %>
        <input type=submit name=Submit value='<%=MesgS("Añadir Blog","Add Blog")%>' class=form /> <input type=reset name=Reset value='<%=MesgS("Limpiar Campos","Clear Fields")%>' class=form /> 
<%      end if %>
        </div></div></form>	
<% End Sub %>