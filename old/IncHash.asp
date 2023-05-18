<%
'---- Funciones para TPV --------------------------------------------------------
Class clsSHA1

'---- AndW devuelve el And palabra a palabra de [pBytWord1Ary] y [pBytWord2Ary]
Private Function AndW(ByRef pBytWord1Ary, ByRef pBytWord2Ary)
	dim lBytWordAry(3), lLngIndex

	for lLngIndex = 0 To 3
		lBytWordAry(lLngIndex) = CByte(pBytWord1Ary(lLngIndex) And pBytWord2Ary(lLngIndex))
	next
	AndW = lBytWordAry
End Function

'---- OrW devuelve el Or palabra a palabra
Private Function OrW(ByRef pBytWord1Ary, ByRef pBytWord2Ary)
	dim lBytWordAry(3), lLngIndex

	for lLngIndex = 0 To 3
		lBytWordAry(lLngIndex) = CByte(pBytWord1Ary(lLngIndex) Or pBytWord2Ary(lLngIndex))
	next
	OrW = lBytWordAry
End Function

'---- XOrW devuelve el Or palabra a palabra
Private Function XorW(ByRef pBytWord1Ary, ByRef pBytWord2Ary)
	dim lBytWordAry(3), lLngIndex
	
	for lLngIndex = 0 To 3
		lBytWordAry(lLngIndex) = CByte(pBytWord1Ary(lLngIndex) Xor pBytWord2Ary(lLngIndex))
	next
	XorW = lBytWordAry
End Function

'---- NotW devuelve el Not palabra a palabra
Private Function NotW(ByRef pBytWordAry)
	dim lBytWordAry(3), lLngIndex
	
	for lLngIndex = 0 To 3
		lBytWordAry(lLngIndex) = Not CByte(pBytWordAry(lLngIndex))
	next
	
	NotW = lBytWordAry
End Function

'---- AddW suma 2 palabras
Private Function AddW(ByRef pBytWord1Ary, ByRef pBytWord2Ary)
	dim lLngIndex, lIntTotal
	dim lBytWordAry(3)
	
	for lLngIndex = 3 To 0 Step -1
		if lLngIndex = 3 Then
			lIntTotal = CInt(pBytWord1Ary(lLngIndex)) + pBytWord2Ary(lLngIndex)
			lBytWordAry(lLngIndex) = lIntTotal Mod 256
		else
			lIntTotal = CInt(pBytWord1Ary(lLngIndex)) + pBytWord2Ary(lLngIndex) + (lIntTotal \ 256)
			lBytWordAry(lLngIndex) = lIntTotal Mod 256
		end If
	next
	AddW = lBytWordAry
End Function

'---- CircShiftLeftW hace un shift a la izquierda de [pBytWordAry], [pLngShift] posiciones
Private Function CircShiftLeftW(ByRef pBytWordAry, ByRef pLngShift)
	dim lDbl1, lDbl2
	
	lDbl1 = WordToDouble(pBytWordAry)
	lDbl2 = lDbl1
	lDbl1 = CDbl(lDbl1 * (2 ^ pLngShift))
	lDbl2 = CDbl(lDbl2 / (2 ^ (32 - pLngShift)))
	CircShiftLeftW = OrW(DoubleToWord(lDbl1), DoubleToWord(lDbl2))
End Function

'---- WordToHex pasa [pBytWordAry] a hexadecimal
Private Function WordToHex(ByRef pBytWordAry)
	dim lLngIndex

	for lLngIndex = 0 To 3
		WordToHex = WordToHex & Right("0" & Hex(pBytWordAry(lLngIndex)), 2)
	next
End Function

'---- HexToWord pasa [pStrHex] a número
Private Function HexToWord(ByRef pStrHex)
	HexToWord = DoubleToWord(CDbl("&h" & pStrHex)) ' needs "#" at end for VB?
End Function

'---- DoubleToWord pasa [pDblValue] a Array
Private Function DoubleToWord(ByRef pDblValue)
	dim lBytWordAry(3)
	
    lBytWordAry(0) = Int(DMod(pDblValue, 2 ^ 32) / (2 ^ 24))
	lBytWordAry(1) = Int(DMod(pDblValue, 2 ^ 24) / (2 ^ 16))
	lBytWordAry(2) = Int(DMod(pDblValue, 2 ^ 16) / (2 ^ 8))
	lBytWordAry(3) = Int(DMod(pDblValue, 2 ^ 8))
	DoubleToWord = lBytWordAry
End Function

'---- WordToDouble pasa [pBytWordAry] a Double
Private Function WordToDouble(ByRef pBytWordAry)
	WordToDouble = CDbl((pBytWordAry(0) * (2 ^ 24)) + (pBytWordAry(1) * (2 ^ 16)) + (pBytWordAry(2) * (2 ^ 8)) + pBytWordAry(3))
End Function

'---- DMod hace el Modulo de [pDblValue] entre [pDblDivisor]
Private Function DMod(ByRef pDblValue, ByRef pDblDivisor)
	dim lDblMod

	lDblMod = CDbl(CDbl(pDblValue) - (Int(CDbl(pDblValue) / CDbl(pDblDivisor)) * CDbl(pDblDivisor)))
	if lDblMod < 0 then lDblMod = CDbl(lDblMod + pDblDivisor)
	DMod = lDblMod
End Function

'---- F Pasa a Flotante
Private Function F(ByRef lIntT, ByRef pBytWordBAry, ByRef pBytWordCAry, ByRef pBytWordDAry)
	if lIntT <= 19 then
		F = OrW(AndW(pBytWordBAry, pBytWordCAry), AndW((NotW(pBytWordBAry)), pBytWordDAry))
	elseif lIntT <= 39 then
		F = XorW(XorW(pBytWordBAry, pBytWordCAry), pBytWordDAry)
	elseIf lIntT <= 59 Then
		F = OrW(OrW(AndW(pBytWordBAry, pBytWordCAry), AndW(pBytWordBAry, pBytWordDAry)), AndW(pBytWordCAry, pBytWordDAry))
	else
		F = XorW(XorW(pBytWordBAry, pBytWordCAry), pBytWordDAry)
	end If
End Function

'---- SecureHash codifica [pStrMessage]
Public Function SecureHash(ByVal pStrMessage)
	dim lLngLen, lBytLenW, lStrPadMessage, lLngNumBlocks, lVarWordWAry(79), lLngTempWordWAry, lStrBlockText, lStrWordText, lLngBlock, lIntT, lBytTempAry, lVarWordKAry(3)
    dim lBytWordH0Ary, lBytWordH1Ary, lBytWordH2Ary, lBytWordH3Ary, lBytWordH4Ary, lBytWordAAry, lBytWordBAry, lBytWordCAry, lBytWordDAry, lBytWordEAry, lBytWordFAry

	lLngLen = Len(pStrMessage)
	lBytLenW = DoubleToWord(CDbl(lLngLen) * 8)
	lStrPadMessage = pStrMessage & Chr(128) & String((128 - (lLngLen Mod 64) - 9) Mod 64, Chr(0)) & String(4, Chr(0)) & Chr(lBytLenW(0)) & Chr(lBytLenW(1)) & Chr(lBytLenW(2)) & Chr(lBytLenW(3))

	lLngNumBlocks = Len(lStrPadMessage) / 64
	lVarWordKAry(0) = HexToWord("5A827999")
	lVarWordKAry(1) = HexToWord("6ED9EBA1")
	lVarWordKAry(2) = HexToWord("8F1BBCDC")
	lVarWordKAry(3) = HexToWord("CA62C1D6")
	lBytWordH0Ary = HexToWord("67452301")
	lBytWordH1Ary = HexToWord("EFCDAB89")
	lBytWordH2Ary = HexToWord("98BADCFE")
	lBytWordH3Ary = HexToWord("10325476")
	lBytWordH4Ary = HexToWord("C3D2E1F0")

	for lLngBlock = 0 To lLngNumBlocks - 1
		lStrBlockText = Mid(lStrPadMessage, (lLngBlock * 64) + 1, 64)

		for lIntT = 0 To 15
			lStrWordText = Mid(lStrBlockText, (lIntT * 4) + 1, 4)
			lVarWordWAry(lIntT)=Array(Asc(Mid(lStrWordText, 1, 1)), Asc(Mid(lStrWordText, 2, 1)), Asc(Mid(lStrWordText, 3, 1)), Asc(Mid(lStrWordText, 4, 1)))
		next

		for lIntT = 16 To 79
			lVarWordWAry(lIntT) = CircShiftLeftW(XorW(XorW(XorW(lVarWordWAry(lIntT - 3), lVarWordWAry(lIntT - 8)), lVarWordWAry(lIntT - 14)), lVarWordWAry(lIntT - 16)), 1)
		next
		
		lBytWordAAry = lBytWordH0Ary
		lBytWordBAry = lBytWordH1Ary
		lBytWordCAry = lBytWordH2Ary
		lBytWordDAry = lBytWordH3Ary
		lBytWordEAry = lBytWordH4Ary
		
		for lIntT = 0 To 79
			lBytWordFAry = F(lIntT, lBytWordBAry, lBytWordCAry, lBytWordDAry)
			lBytTempAry=AddW(AddW(AddW(AddW(CircShiftLeftW(lBytWordAAry, 5), lBytWordFAry), lBytWordEAry), lVarWordWAry(lIntT)), lVarWordKAry(lIntT \ 20))
			lBytWordEAry = lBytWordDAry
			lBytWordDAry = lBytWordCAry
			lBytWordCAry = CircShiftLeftW(lBytWordBAry, 30)
			lBytWordBAry = lBytWordAAry
			lBytWordAAry = lBytTempAry
		next
		lBytWordH0Ary = AddW(lBytWordH0Ary, lBytWordAAry)
		lBytWordH1Ary = AddW(lBytWordH1Ary, lBytWordBAry)
		lBytWordH2Ary = AddW(lBytWordH2Ary, lBytWordCAry)
		lBytWordH3Ary = AddW(lBytWordH3Ary, lBytWordDAry)
		lBytWordH4Ary = AddW(lBytWordH4Ary, lBytWordEAry)
	Next
	SecureHash = WordToHex(lBytWordH0Ary) & WordToHex(lBytWordH1Ary) & WordToHex(lBytWordH2Ary) & WordToHex(lBytWordH3Ary) & WordToHex(lBytWordH4Ary)
End Function
' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
%>