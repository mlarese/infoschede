<% 
'classe per l'integrazione di funzioni per la crittografia, generazione di UID e checksum

class CryptographyManager
	private oMD5
	private oSHA256
	private oAES
	
	
	'**************************************************************************************************************
	'FUNZIONI DI INIZZIALIZZAZIONE
	'**************************************************************************************************************
	
	Private Sub Class_Initialize()
		set oMD5 = new Algorithm_MD5
		set oSHA256 = new Algorithm_SHA256
		set oAES = new Algorithm_AES
	end sub
		
	Private Sub Class_Terminate()
		set oMD5 = nothing
		set oSHA256 = nothing
		set oAES = nothing
	end sub
	
	
	'**************************************************************************************************************
	'FUNZIONI PUBBLICHE
	'**************************************************************************************************************
	
	'...........................................................................
	'MD5
	'...........................................................................
	
	'ritorna checksum md5 della stringa
	public function md5_of_string(str)
		md5_of_string = oMD5.MD5(str)
	end function
	
	
	'ritorna checksum md5 per il contenuto testuale del file
	public function md5_of_file(path)
		md5_of_file = md5_of_string(GetFileContent(path))
	end function
	
	
	'...........................................................................
	'sha256
	'...........................................................................
	
	'ritorna checksum sha256 della stringa
	public function sha256_of_string(str)
		sha256_of_string = oSHA256.SHA256(str)
	end function
	
	'ritorna checksum sha256 per il contenuto testuale del file
	public function sha256_of_file(path)
		sha256_of_file = sha256_of_string(GetFileContent(path))
	end function
	
	'...........................................................................
	'AES
	'...........................................................................
	
	'ritorna codifica AES della stringa
	public function aes_of_string(str, key)
		aes_of_string = oAES.AESEncrypt(str, key)
	end function
	
	
	'ritorna codifica AES del file
	public function aes_of_file(path, key)
		aes_of_file = aes_of_string(GetFileContent(path), key)
	end function
	
	
	'ritorna decodifica AES della stringa
	public function string_from_aes(aes_string, key)
		string_from_aes = oAES.AESDecrypt(aes_string, key)
	end function
	
		
	'**************************************************************************************************************
	'FUNZIONI PRIVATE
	'**************************************************************************************************************
	
	'recupera il contenuto testuale di un file
	private function GetFileContent(path)
		dim fso, file
		set fso = CreateObject("Scripting.FileSystemObject")
		
		'verifica se il file esiste
		if fso.FileExists(path) then
			'recupera file
			set file = fso.OpenTextFile(path, 1, 0)
			GetFileContent = file.ReadAll()
			CALL file.close()
			set file = nothing
		else
			GetFileContent = ""
		end if

		set fso = nothing
	end function
	
end class



'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************


'**************************************************************************************************************
'**************************************************************************************************************
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32
	

class Algorithm_MD5

	Private m_lOnBits(30)
	Private m_l2Power(30)
	 
	Private Sub Class_Initialize()
		m_lOnBits(0) = CLng(1)
	    m_lOnBits(1) = CLng(3)
	    m_lOnBits(2) = CLng(7)
	    m_lOnBits(3) = CLng(15)
	    m_lOnBits(4) = CLng(31)
	    m_lOnBits(5) = CLng(63)
	    m_lOnBits(6) = CLng(127)
	    m_lOnBits(7) = CLng(255)
	    m_lOnBits(8) = CLng(511)
	    m_lOnBits(9) = CLng(1023)
	    m_lOnBits(10) = CLng(2047)
	    m_lOnBits(11) = CLng(4095)
	    m_lOnBits(12) = CLng(8191)
	    m_lOnBits(13) = CLng(16383)
	    m_lOnBits(14) = CLng(32767)
	    m_lOnBits(15) = CLng(65535)
	    m_lOnBits(16) = CLng(131071)
	    m_lOnBits(17) = CLng(262143)
	    m_lOnBits(18) = CLng(524287)
	    m_lOnBits(19) = CLng(1048575)
	    m_lOnBits(20) = CLng(2097151)
	    m_lOnBits(21) = CLng(4194303)
	    m_lOnBits(22) = CLng(8388607)
	    m_lOnBits(23) = CLng(16777215)
	    m_lOnBits(24) = CLng(33554431)
	    m_lOnBits(25) = CLng(67108863)
	    m_lOnBits(26) = CLng(134217727)
	    m_lOnBits(27) = CLng(268435455)
	    m_lOnBits(28) = CLng(536870911)
	    m_lOnBits(29) = CLng(1073741823)
	    m_lOnBits(30) = CLng(2147483647)
	    
	    m_l2Power(0) = CLng(1)
	    m_l2Power(1) = CLng(2)
	    m_l2Power(2) = CLng(4)
	    m_l2Power(3) = CLng(8)
	    m_l2Power(4) = CLng(16)
	    m_l2Power(5) = CLng(32)
	    m_l2Power(6) = CLng(64)
	    m_l2Power(7) = CLng(128)
	    m_l2Power(8) = CLng(256)
	    m_l2Power(9) = CLng(512)
	    m_l2Power(10) = CLng(1024)
	    m_l2Power(11) = CLng(2048)
	    m_l2Power(12) = CLng(4096)
	    m_l2Power(13) = CLng(8192)
	    m_l2Power(14) = CLng(16384)
	    m_l2Power(15) = CLng(32768)
	    m_l2Power(16) = CLng(65536)
	    m_l2Power(17) = CLng(131072)
	    m_l2Power(18) = CLng(262144)
	    m_l2Power(19) = CLng(524288)
	    m_l2Power(20) = CLng(1048576)
	    m_l2Power(21) = CLng(2097152)
	    m_l2Power(22) = CLng(4194304)
	    m_l2Power(23) = CLng(8388608)
	    m_l2Power(24) = CLng(16777216)
	    m_l2Power(25) = CLng(33554432)
	    m_l2Power(26) = CLng(67108864)
	    m_l2Power(27) = CLng(134217728)
	    m_l2Power(28) = CLng(268435456)
	    m_l2Power(29) = CLng(536870912)
	    m_l2Power(30) = CLng(1073741824)
	end sub
	 
	    	
	Private Function LShift(lValue, iShiftBits)
	    If iShiftBits = 0 Then
	        LShift = lValue
	        Exit Function
	    ElseIf iShiftBits = 31 Then
	        If lValue And 1 Then
	            LShift = &H80000000
	        Else
	            LShift = 0
	        End If
	        Exit Function
	    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
	        Err.Raise 6
	    End If
	
	    If (lValue And m_l2Power(31 - iShiftBits)) Then
	        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
	    Else
	        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
	    End If
	End Function
	
	Private Function RShift(lValue, iShiftBits)
	    If iShiftBits = 0 Then
	        RShift = lValue
	        Exit Function
	    ElseIf iShiftBits = 31 Then
	        If lValue And &H80000000 Then
	            RShift = 1
	        Else
	            RShift = 0
	        End If
	        Exit Function
	    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
	        Err.Raise 6
	    End If
	    
	    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
	
	    If (lValue And &H80000000) Then
	        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
	    End If
	End Function
	
	Private Function RotateLeft(lValue, iShiftBits)
	    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
	End Function
	
	Private Function AddUnsigned(lX, lY)
	    Dim lX4
	    Dim lY4
	    Dim lX8
	    Dim lY8
	    Dim lResult
	 
	    lX8 = lX And &H80000000
	    lY8 = lY And &H80000000
	    lX4 = lX And &H40000000
	    lY4 = lY And &H40000000
	 
	    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
	 
	    If lX4 And lY4 Then
	        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
	    ElseIf lX4 Or lY4 Then
	        If lResult And &H40000000 Then
	            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
	        Else
	            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
	        End If
	    Else
	        lResult = lResult Xor lX8 Xor lY8
	    End If
	 
	    AddUnsigned = lResult
	End Function
	
	Private Function F(x, y, z)
	    F = (x And y) Or ((Not x) And z)
	End Function
	
	Private Function G(x, y, z)
	    G = (x And z) Or (y And (Not z))
	End Function
	
	Private Function H(x, y, z)
	    H = (x Xor y Xor z)
	End Function
	
	Private Function I(x, y, z)
	    I = (y Xor (x Or (Not z)))
	End Function
	
	Private Sub FF(a, b, c, d, x, s, ac)
	    a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
	    a = RotateLeft(a, s)
	    a = AddUnsigned(a, b)
	End Sub
	
	Private Sub GG(a, b, c, d, x, s, ac)
	    a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
	    a = RotateLeft(a, s)
	    a = AddUnsigned(a, b)
	End Sub
	
	Private Sub HH(a, b, c, d, x, s, ac)
	    a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
	    a = RotateLeft(a, s)
	    a = AddUnsigned(a, b)
	End Sub
	
	Private Sub II(a, b, c, d, x, s, ac)
	    a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
	    a = RotateLeft(a, s)
	    a = AddUnsigned(a, b)
	End Sub
	
	Private Function ConvertToWordArray(sMessage)
	    Dim lMessageLength
	    Dim lNumberOfWords
	    Dim lWordArray()
	    Dim lBytePosition
	    Dim lByteCount
	    Dim lWordCount
	    
	    Const MODULUS_BITS = 512
	    Const CONGRUENT_BITS = 448
	    
	    lMessageLength = Len(sMessage)
	    
	    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
	    ReDim lWordArray(lNumberOfWords - 1)
	    
	    lBytePosition = 0
	    lByteCount = 0
	    Do Until lByteCount >= lMessageLength
	        lWordCount = lByteCount \ BYTES_TO_A_WORD
	        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
	        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
	        lByteCount = lByteCount + 1
	    Loop
	
	    lWordCount = lByteCount \ BYTES_TO_A_WORD
	    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
	
	    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
	
	    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
	    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
	    
	    ConvertToWordArray = lWordArray
	End Function
	
	Private Function WordToHex(lValue)
	    Dim lByte
	    Dim lCount
	    
	    For lCount = 0 To 3
	        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
	        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
	    Next
	End Function
	
	Public Function MD5(sMessage)
	    Dim x
	    Dim k
	    Dim AA
	    Dim BB
	    Dim CC
	    Dim DD
	    Dim a
	    Dim b
	    Dim c
	    Dim d
	    
	    Const S11 = 7
	    Const S12 = 12
	    Const S13 = 17
	    Const S14 = 22
	    Const S21 = 5
	    Const S22 = 9
	    Const S23 = 14
	    Const S24 = 20
	    Const S31 = 4
	    Const S32 = 11
	    Const S33 = 16
	    Const S34 = 23
	    Const S41 = 6
	    Const S42 = 10
	    Const S43 = 15
	    Const S44 = 21
	
	    x = ConvertToWordArray(sMessage)
	    
	    a = &H67452301
	    b = &HEFCDAB89
	    c = &H98BADCFE
	    d = &H10325476
	
	    For k = 0 To UBound(x) Step 16
	        AA = a
	        BB = b
	        CC = c
	        DD = d
	    
	        FF a, b, c, d, x(k + 0), S11, &HD76AA478
	        FF d, a, b, c, x(k + 1), S12, &HE8C7B756
	        FF c, d, a, b, x(k + 2), S13, &H242070DB
	        FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
	        FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
	        FF d, a, b, c, x(k + 5), S12, &H4787C62A
	        FF c, d, a, b, x(k + 6), S13, &HA8304613
	        FF b, c, d, a, x(k + 7), S14, &HFD469501
	        FF a, b, c, d, x(k + 8), S11, &H698098D8
	        FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
	        FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
	        FF b, c, d, a, x(k + 11), S14, &H895CD7BE
	        FF a, b, c, d, x(k + 12), S11, &H6B901122
	        FF d, a, b, c, x(k + 13), S12, &HFD987193
	        FF c, d, a, b, x(k + 14), S13, &HA679438E
	        FF b, c, d, a, x(k + 15), S14, &H49B40821
	    
	        GG a, b, c, d, x(k + 1), S21, &HF61E2562
	        GG d, a, b, c, x(k + 6), S22, &HC040B340
	        GG c, d, a, b, x(k + 11), S23, &H265E5A51
	        GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
	        GG a, b, c, d, x(k + 5), S21, &HD62F105D
	        GG d, a, b, c, x(k + 10), S22, &H2441453
	        GG c, d, a, b, x(k + 15), S23, &HD8A1E681
	        GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
	        GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
	        GG d, a, b, c, x(k + 14), S22, &HC33707D6
	        GG c, d, a, b, x(k + 3), S23, &HF4D50D87
	        GG b, c, d, a, x(k + 8), S24, &H455A14ED
	        GG a, b, c, d, x(k + 13), S21, &HA9E3E905
	        GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
	        GG c, d, a, b, x(k + 7), S23, &H676F02D9
	        GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
	            
	        HH a, b, c, d, x(k + 5), S31, &HFFFA3942
	        HH d, a, b, c, x(k + 8), S32, &H8771F681
	        HH c, d, a, b, x(k + 11), S33, &H6D9D6122
	        HH b, c, d, a, x(k + 14), S34, &HFDE5380C
	        HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
	        HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
	        HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
	        HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
	        HH a, b, c, d, x(k + 13), S31, &H289B7EC6
	        HH d, a, b, c, x(k + 0), S32, &HEAA127FA
	        HH c, d, a, b, x(k + 3), S33, &HD4EF3085
	        HH b, c, d, a, x(k + 6), S34, &H4881D05
	        HH a, b, c, d, x(k + 9), S31, &HD9D4D039
	        HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
	        HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
	        HH b, c, d, a, x(k + 2), S34, &HC4AC5665
	    
	        II a, b, c, d, x(k + 0), S41, &HF4292244
	        II d, a, b, c, x(k + 7), S42, &H432AFF97
	        II c, d, a, b, x(k + 14), S43, &HAB9423A7
	        II b, c, d, a, x(k + 5), S44, &HFC93A039
	        II a, b, c, d, x(k + 12), S41, &H655B59C3
	        II d, a, b, c, x(k + 3), S42, &H8F0CCC92
	        II c, d, a, b, x(k + 10), S43, &HFFEFF47D
	        II b, c, d, a, x(k + 1), S44, &H85845DD1
	        II a, b, c, d, x(k + 8), S41, &H6FA87E4F
	        II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
	        II c, d, a, b, x(k + 6), S43, &HA3014314
	        II b, c, d, a, x(k + 13), S44, &H4E0811A1
	        II a, b, c, d, x(k + 4), S41, &HF7537E82
	        II d, a, b, c, x(k + 11), S42, &HBD3AF235
	        II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
	        II b, c, d, a, x(k + 9), S44, &HEB86D391
	    
	        a = AddUnsigned(a, AA)
	        b = AddUnsigned(b, BB)
	        c = AddUnsigned(c, CC)
	        d = AddUnsigned(d, DD)
	    Next
	    
	    MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
	End Function
end class
'**************************************************************************************************************
'**************************************************************************************************************

'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************

'**************************************************************************************************************
'**************************************************************************************************************
class Algorithm_sha256

	Private m_lOnBits(30)
	Private m_l2Power(30)
	Private K(63)
	
	Private Sub Class_Initialize()
		m_lOnBits(0) = CLng(1)
		m_lOnBits(1) = CLng(3)
		m_lOnBits(2) = CLng(7)
		m_lOnBits(3) = CLng(15)
		m_lOnBits(4) = CLng(31)
		m_lOnBits(5) = CLng(63)
		m_lOnBits(6) = CLng(127)
		m_lOnBits(7) = CLng(255)
		m_lOnBits(8) = CLng(511)
		m_lOnBits(9) = CLng(1023)
		m_lOnBits(10) = CLng(2047)
		m_lOnBits(11) = CLng(4095)
		m_lOnBits(12) = CLng(8191)
		m_lOnBits(13) = CLng(16383)
		m_lOnBits(14) = CLng(32767)
		m_lOnBits(15) = CLng(65535)
		m_lOnBits(16) = CLng(131071)
		m_lOnBits(17) = CLng(262143)
		m_lOnBits(18) = CLng(524287)
		m_lOnBits(19) = CLng(1048575)
		m_lOnBits(20) = CLng(2097151)
		m_lOnBits(21) = CLng(4194303)
		m_lOnBits(22) = CLng(8388607)
		m_lOnBits(23) = CLng(16777215)
		m_lOnBits(24) = CLng(33554431)
		m_lOnBits(25) = CLng(67108863)
		m_lOnBits(26) = CLng(134217727)
		m_lOnBits(27) = CLng(268435455)
		m_lOnBits(28) = CLng(536870911)
		m_lOnBits(29) = CLng(1073741823)
		m_lOnBits(30) = CLng(2147483647)
		
		m_l2Power(0) = CLng(1)
		m_l2Power(1) = CLng(2)
		m_l2Power(2) = CLng(4)
		m_l2Power(3) = CLng(8)
		m_l2Power(4) = CLng(16)
		m_l2Power(5) = CLng(32)
		m_l2Power(6) = CLng(64)
		m_l2Power(7) = CLng(128)
		m_l2Power(8) = CLng(256)
		m_l2Power(9) = CLng(512)
		m_l2Power(10) = CLng(1024)
		m_l2Power(11) = CLng(2048)
		m_l2Power(12) = CLng(4096)
		m_l2Power(13) = CLng(8192)
		m_l2Power(14) = CLng(16384)
		m_l2Power(15) = CLng(32768)
		m_l2Power(16) = CLng(65536)
		m_l2Power(17) = CLng(131072)
		m_l2Power(18) = CLng(262144)
		m_l2Power(19) = CLng(524288)
		m_l2Power(20) = CLng(1048576)
		m_l2Power(21) = CLng(2097152)
		m_l2Power(22) = CLng(4194304)
		m_l2Power(23) = CLng(8388608)
		m_l2Power(24) = CLng(16777216)
		m_l2Power(25) = CLng(33554432)
		m_l2Power(26) = CLng(67108864)
		m_l2Power(27) = CLng(134217728)
		m_l2Power(28) = CLng(268435456)
		m_l2Power(29) = CLng(536870912)
		m_l2Power(30) = CLng(1073741824)
		    
		K(0) = &H428A2F98
		K(1) = &H71374491
		K(2) = &HB5C0FBCF
		K(3) = &HE9B5DBA5
		K(4) = &H3956C25B
		K(5) = &H59F111F1
		K(6) = &H923F82A4
		K(7) = &HAB1C5ED5
		K(8) = &HD807AA98
		K(9) = &H12835B01
		K(10) = &H243185BE
		K(11) = &H550C7DC3
		K(12) = &H72BE5D74
		K(13) = &H80DEB1FE
		K(14) = &H9BDC06A7
		K(15) = &HC19BF174
		K(16) = &HE49B69C1
		K(17) = &HEFBE4786
		K(18) = &HFC19DC6
		K(19) = &H240CA1CC
		K(20) = &H2DE92C6F
		K(21) = &H4A7484AA
		K(22) = &H5CB0A9DC
		K(23) = &H76F988DA
		K(24) = &H983E5152
		K(25) = &HA831C66D
		K(26) = &HB00327C8
		K(27) = &HBF597FC7
		K(28) = &HC6E00BF3
		K(29) = &HD5A79147
		K(30) = &H6CA6351
		K(31) = &H14292967
		K(32) = &H27B70A85
		K(33) = &H2E1B2138
		K(34) = &H4D2C6DFC
		K(35) = &H53380D13
		K(36) = &H650A7354
		K(37) = &H766A0ABB
		K(38) = &H81C2C92E
		K(39) = &H92722C85
		K(40) = &HA2BFE8A1
		K(41) = &HA81A664B
		K(42) = &HC24B8B70
		K(43) = &HC76C51A3
		K(44) = &HD192E819
		K(45) = &HD6990624
		K(46) = &HF40E3585
		K(47) = &H106AA070
		K(48) = &H19A4C116
		K(49) = &H1E376C08
		K(50) = &H2748774C
		K(51) = &H34B0BCB5
		K(52) = &H391C0CB3
		K(53) = &H4ED8AA4A
		K(54) = &H5B9CCA4F
		K(55) = &H682E6FF3
		K(56) = &H748F82EE
		K(57) = &H78A5636F
		K(58) = &H84C87814
		K(59) = &H8CC70208
		K(60) = &H90BEFFFA
		K(61) = &HA4506CEB
		K(62) = &HBEF9A3F7
		K(63) = &HC67178F2
	end sub
	
	Private Function LShift(lValue, iShiftBits)
	    If iShiftBits = 0 Then
	        LShift = lValue
	        Exit Function
	    ElseIf iShiftBits = 31 Then
	        If lValue And 1 Then
	            LShift = &H80000000
	        Else
	            LShift = 0
	        End If
	        Exit Function
	    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
	        Err.Raise 6
	    End If
	    
	    If (lValue And m_l2Power(31 - iShiftBits)) Then
	        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
	    Else
	        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
	    End If
	End Function
	
	Private Function RShift(lValue, iShiftBits)
	    If iShiftBits = 0 Then
	        RShift = lValue
	        Exit Function
	    ElseIf iShiftBits = 31 Then
	        If lValue And &H80000000 Then
	            RShift = 1
	        Else
	            RShift = 0
	        End If
	        Exit Function
	    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
	        Err.Raise 6
	    End If
	    
	    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
	    
	    If (lValue And &H80000000) Then
	        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
	    End If
	End Function
	
	Private Function AddUnsigned(lX, lY)
	    Dim lX4
	    Dim lY4
	    Dim lX8
	    Dim lY8
	    Dim lResult
	 
	    lX8 = lX And &H80000000
	    lY8 = lY And &H80000000
	    lX4 = lX And &H40000000
	    lY4 = lY And &H40000000
	 
	    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
	 
	    If lX4 And lY4 Then
	        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
	    ElseIf lX4 Or lY4 Then
	        If lResult And &H40000000 Then
	            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
	        Else
	            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
	        End If
	    Else
	        lResult = lResult Xor lX8 Xor lY8
	    End If
	 
	    AddUnsigned = lResult
	End Function
	
	Private Function Ch(x, y, z)
	    Ch = ((x And y) Xor ((Not x) And z))
	End Function
	
	Private Function Maj(x, y, z)
	    Maj = ((x And y) Xor (x And z) Xor (y And z))
	End Function
	
	Private Function S(x, n)
	    S = (RShift(x, (n And m_lOnBits(4))) Or LShift(x, (32 - (n And m_lOnBits(4)))))
	End Function
	
	Private Function R(x, n)
	    R = RShift(x, CInt(n And m_lOnBits(4)))
	End Function
	
	Private Function Sigma0(x)
	    Sigma0 = (S(x, 2) Xor S(x, 13) Xor S(x, 22))
	End Function
	
	Private Function Sigma1(x)
	    Sigma1 = (S(x, 6) Xor S(x, 11) Xor S(x, 25))
	End Function
	
	Private Function Gamma0(x)
	    Gamma0 = (S(x, 7) Xor S(x, 18) Xor R(x, 3))
	End Function
	
	Private Function Gamma1(x)
	    Gamma1 = (S(x, 17) Xor S(x, 19) Xor R(x, 10))
	End Function
	
	Private Function ConvertToWordArray(sMessage)
	    Dim lMessageLength
	    Dim lNumberOfWords
	    Dim lWordArray()
	    Dim lBytePosition
	    Dim lByteCount
	    Dim lWordCount
	    Dim lByte
	    
	    Const MODULUS_BITS = 512
	    Const CONGRUENT_BITS = 448
	    
	    lMessageLength = Len(sMessage)
	    
	    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
	    ReDim lWordArray(lNumberOfWords - 1)
	    
	    lBytePosition = 0
	    lByteCount = 0
	    Do Until lByteCount >= lMessageLength
	        lWordCount = lByteCount \ BYTES_TO_A_WORD
	        
	        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
	        
	        lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
	        
	        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
	        lByteCount = lByteCount + 1
	    Loop
	
	    lWordCount = lByteCount \ BYTES_TO_A_WORD
	    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
	
	    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
	
	    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
	    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
	    
	    ConvertToWordArray = lWordArray
	End Function
	
	Public Function SHA256(sMessage)
	    Dim HASH(7)
	    Dim M
	    Dim W(63)
	    Dim a
	    Dim b
	    Dim c
	    Dim d
	    Dim e
	    Dim f
	    Dim g
	    Dim h
	    Dim i
	    Dim j
	    Dim T1
	    Dim T2
	    
	    HASH(0) = &H6A09E667
	    HASH(1) = &HBB67AE85
	    HASH(2) = &H3C6EF372
	    HASH(3) = &HA54FF53A
	    HASH(4) = &H510E527F
	    HASH(5) = &H9B05688C
	    HASH(6) = &H1F83D9AB
	    HASH(7) = &H5BE0CD19
	    
	    M = ConvertToWordArray(sMessage)
	    
	    For i = 0 To UBound(M) Step 16
	        a = HASH(0)
	        b = HASH(1)
	        c = HASH(2)
	        d = HASH(3)
	        e = HASH(4)
	        f = HASH(5)
	        g = HASH(6)
	        h = HASH(7)
	        
	        For j = 0 To 63
	            If j < 16 Then
	                W(j) = M(j + i)
	            Else
	                W(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j - 2)), W(j - 7)), Gamma0(W(j - 15))), W(j - 16))
	            End If
	                
	            T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, Sigma1(e)), Ch(e, f, g)), K(j)), W(j))
	            T2 = AddUnsigned(Sigma0(a), Maj(a, b, c))
	            
	            h = g
	            g = f
	            f = e
	            e = AddUnsigned(d, T1)
	            d = c
	            c = b
	            b = a
	            a = AddUnsigned(T1, T2)
	        Next
	        
	        HASH(0) = AddUnsigned(a, HASH(0))
	        HASH(1) = AddUnsigned(b, HASH(1))
	        HASH(2) = AddUnsigned(c, HASH(2))
	        HASH(3) = AddUnsigned(d, HASH(3))
	        HASH(4) = AddUnsigned(e, HASH(4))
	        HASH(5) = AddUnsigned(f, HASH(5))
	        HASH(6) = AddUnsigned(g, HASH(6))
	        HASH(7) = AddUnsigned(h, HASH(7))
	    Next
	    
	    SHA256 = LCase(Right("00000000" & Hex(HASH(0)), 8) & Right("00000000" & Hex(HASH(1)), 8) & Right("00000000" & Hex(HASH(2)), 8) & Right("00000000" & Hex(HASH(3)), 8) & Right("00000000" & Hex(HASH(4)), 8) & Right("00000000" & Hex(HASH(5)), 8) & Right("00000000" & Hex(HASH(6)), 8) & Right("00000000" & Hex(HASH(7)), 8))
	End Function
end class

'**************************************************************************************************************
'**************************************************************************************************************

'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
class Algorithm_AES

	Private m_lOnBits(30)
	Private m_l2Power(30)
	Private m_bytOnBits(7)
	Private m_byt2Power(7)

	Private m_InCo(3)

	Private m_fbsub(255)
	Private m_rbsub(255)
	Private m_ptab(255)
	Private m_ltab(255)
	Private m_ftable(255)
	Private m_rtable(255)
	Private m_rco(29)

	Private m_Nk
	Private m_Nb
	Private m_Nr
	Private m_fi(23)
	Private m_ri(23)
	Private m_fkey(119)
	Private m_rkey(119)

	Private Sub Class_Initialize()
		m_InCo(0) = &HB
		m_InCo(1) = &HD
		m_InCo(2) = &H9
		m_InCo(3) = &HE
			
		m_bytOnBits(0) = 1
		m_bytOnBits(1) = 3
		m_bytOnBits(2) = 7
		m_bytOnBits(3) = 15
		m_bytOnBits(4) = 31
		m_bytOnBits(5) = 63
		m_bytOnBits(6) = 127
		m_bytOnBits(7) = 255
			
		m_byt2Power(0) = 1
		m_byt2Power(1) = 2
		m_byt2Power(2) = 4
		m_byt2Power(3) = 8
		m_byt2Power(4) = 16
		m_byt2Power(5) = 32
		m_byt2Power(6) = 64
		m_byt2Power(7) = 128
			
		m_lOnBits(0) = 1
		m_lOnBits(1) = 3
		m_lOnBits(2) = 7
		m_lOnBits(3) = 15
		m_lOnBits(4) = 31
		m_lOnBits(5) = 63
		m_lOnBits(6) = 127
		m_lOnBits(7) = 255
		m_lOnBits(8) = 511
		m_lOnBits(9) = 1023
		m_lOnBits(10) = 2047
		m_lOnBits(11) = 4095
		m_lOnBits(12) = 8191
		m_lOnBits(13) = 16383
		m_lOnBits(14) = 32767
		m_lOnBits(15) = 65535
		m_lOnBits(16) = 131071
		m_lOnBits(17) = 262143
		m_lOnBits(18) = 524287
		m_lOnBits(19) = 1048575
		m_lOnBits(20) = 2097151
		m_lOnBits(21) = 4194303
		m_lOnBits(22) = 8388607
		m_lOnBits(23) = 16777215
		m_lOnBits(24) = 33554431
		m_lOnBits(25) = 67108863
		m_lOnBits(26) = 134217727
		m_lOnBits(27) = 268435455
		m_lOnBits(28) = 536870911
		m_lOnBits(29) = 1073741823
		m_lOnBits(30) = 2147483647
			
		m_l2Power(0) = 1
		m_l2Power(1) = 2
		m_l2Power(2) = 4
		m_l2Power(3) = 8
		m_l2Power(4) = 16
		m_l2Power(5) = 32
		m_l2Power(6) = 64
		m_l2Power(7) = 128
		m_l2Power(8) = 256
		m_l2Power(9) = 512
		m_l2Power(10) = 1024
		m_l2Power(11) = 2048
		m_l2Power(12) = 4096
		m_l2Power(13) = 8192
		m_l2Power(14) = 16384
		m_l2Power(15) = 32768
		m_l2Power(16) = 65536
		m_l2Power(17) = 131072
		m_l2Power(18) = 262144
		m_l2Power(19) = 524288
		m_l2Power(20) = 1048576
		m_l2Power(21) = 2097152
		m_l2Power(22) = 4194304
		m_l2Power(23) = 8388608
		m_l2Power(24) = 16777216
		m_l2Power(25) = 33554432
		m_l2Power(26) = 67108864
		m_l2Power(27) = 134217728
		m_l2Power(28) = 268435456
		m_l2Power(29) = 536870912
		m_l2Power(30) = 1073741824
	end sub

	Private Function LShift(lValue, iShiftBits)
		If iShiftBits = 0 Then
			LShift = lValue
			Exit Function
		ElseIf iShiftBits = 31 Then
			If lValue And 1 Then
				LShift = &H80000000
			Else
				LShift = 0
			End If
			Exit Function
		ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
			Err.Raise 6
		End If
		
		If (lValue And m_l2Power(31 - iShiftBits)) Then
			LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
		Else
			LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
		End If
	End Function

	Private Function RShift(lValue, iShiftBits)
		If iShiftBits = 0 Then
			RShift = lValue
			Exit Function
		ElseIf iShiftBits = 31 Then
			If lValue And &H80000000 Then
				RShift = 1
			Else
				RShift = 0
			End If
			Exit Function
		ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
			Err.Raise 6
		End If
		
		RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
		
		If (lValue And &H80000000) Then
			RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
		End If
	End Function

	Private Function LShiftByte(bytValue, bytShiftBits)
		If bytShiftBits = 0 Then
			LShiftByte = bytValue
			Exit Function
		ElseIf bytShiftBits = 7 Then
			If bytValue And 1 Then
				LShiftByte = &H80
			Else
				LShiftByte = 0
			End If
			Exit Function
		ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
			Err.Raise 6
		End If
		
		LShiftByte = ((bytValue And m_bytOnBits(7 - bytShiftBits)) * m_byt2Power(bytShiftBits))
	End Function

	Private Function RShiftByte(bytValue, bytShiftBits)
		If bytShiftBits = 0 Then
			RShiftByte = bytValue
			Exit Function
		ElseIf bytShiftBits = 7 Then
			If bytValue And &H80 Then
				RShiftByte = 1
			Else
				RShiftByte = 0
			End If
			Exit Function
		ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
			Err.Raise 6
		End If
		
		RShiftByte = bytValue \ m_byt2Power(bytShiftBits)
	End Function

	Private Function RotateLeft(lValue, iShiftBits)
		RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
	End Function

	Private Function RotateLeftByte(bytValue, bytShiftBits)
		RotateLeftByte = LShiftByte(bytValue, bytShiftBits) Or RShiftByte(bytValue, (8 - bytShiftBits))
	End Function

	Private Function Pack(b())
		Dim lCount
		Dim lTemp
		
		For lCount = 0 To 3
			lTemp = b(lCount)
			Pack = Pack Or LShift(lTemp, (lCount * 8))
		Next
	End Function

	Private Function PackFrom(b(), k)
		Dim lCount
		Dim lTemp
		
		For lCount = 0 To 3
			lTemp = b(lCount + k)
			PackFrom = PackFrom Or LShift(lTemp, (lCount * 8))
		Next
	End Function

	Private Sub Unpack(a, b())
		b(0) = a And m_lOnBits(7)
		b(1) = RShift(a, 8) And m_lOnBits(7)
		b(2) = RShift(a, 16) And m_lOnBits(7)
		b(3) = RShift(a, 24) And m_lOnBits(7)
	End Sub

	Private Sub UnpackFrom(a, b(), k)
		b(0 + k) = a And m_lOnBits(7)
		b(1 + k) = RShift(a, 8) And m_lOnBits(7)
		b(2 + k) = RShift(a, 16) And m_lOnBits(7)
		b(3 + k) = RShift(a, 24) And m_lOnBits(7)
	End Sub

	Private Function xtime(a)
		Dim b
		
		If (a And &H80) Then
			b = &H1B
		Else
			b = 0
		End If
		
		xtime = LShiftByte(a, 1)
		xtime = xtime Xor b
	End Function

	Private Function bmul(x, y)
		If x <> 0 And y <> 0 Then
			bmul = m_ptab((CLng(m_ltab(x)) + CLng(m_ltab(y))) Mod 255)
		Else
			bmul = 0
		End If
	End Function

	Private Function SubByte(a)
		Dim b(3)
		
		Unpack a, b
		b(0) = m_fbsub(b(0))
		b(1) = m_fbsub(b(1))
		b(2) = m_fbsub(b(2))
		b(3) = m_fbsub(b(3))
		
		SubByte = Pack(b)
	End Function

	Private Function product(x, y)
		Dim xb(3)
		Dim yb(3)
		
		Unpack x, xb
		Unpack y, yb
		product = bmul(xb(0), yb(0)) Xor bmul(xb(1), yb(1)) Xor bmul(xb(2), yb(2)) Xor bmul(xb(3), yb(3))
	End Function

	Private Function InvMixCol(x)
		Dim y
		Dim m
		Dim b(3)
		
		m = Pack(m_InCo)
		b(3) = product(m, x)
		m = RotateLeft(m, 24)
		b(2) = product(m, x)
		m = RotateLeft(m, 24)
		b(1) = product(m, x)
		m = RotateLeft(m, 24)
		b(0) = product(m, x)
		y = Pack(b)
		
		InvMixCol = y
	End Function

	Private Function ByteSub(x)
		Dim y
		Dim z
		
		z = x
		y = m_ptab(255 - m_ltab(z))
		z = y
		z = RotateLeftByte(z, 1)
		y = y Xor z
		z = RotateLeftByte(z, 1)
		y = y Xor z
		z = RotateLeftByte(z, 1)
		y = y Xor z
		z = RotateLeftByte(z, 1)
		y = y Xor z
		y = y Xor &H63
		
		ByteSub = y
	End Function

	Private Sub gentables()
		Dim i
		Dim y
		Dim b(3)
		Dim ib
		
		m_ltab(0) = 0
		m_ptab(0) = 1
		m_ltab(1) = 0
		m_ptab(1) = 3
		m_ltab(3) = 1
		
		For i = 2 To 255
			m_ptab(i) = m_ptab(i - 1) Xor xtime(m_ptab(i - 1))
			m_ltab(m_ptab(i)) = i
		Next
		
		m_fbsub(0) = &H63
		m_rbsub(&H63) = 0
		
		For i = 1 To 255
			ib = i
			y = ByteSub(ib)
			m_fbsub(i) = y
			m_rbsub(y) = i
		Next
		
		y = 1
		For i = 0 To 29
			m_rco(i) = y
			y = xtime(y)
		Next
		
		For i = 0 To 255
			y = m_fbsub(i)
			b(3) = y Xor xtime(y)
			b(2) = y
			b(1) = y
			b(0) = xtime(y)
			m_ftable(i) = Pack(b)
			
			y = m_rbsub(i)
			b(3) = bmul(m_InCo(0), y)
			b(2) = bmul(m_InCo(1), y)
			b(1) = bmul(m_InCo(2), y)
			b(0) = bmul(m_InCo(3), y)
			m_rtable(i) = Pack(b)
		Next
	End Sub

	Private Sub gkey(nb, nk, key())                
		Dim i
		Dim j
		Dim k
		Dim m
		Dim N
		Dim C1
		Dim C2
		Dim C3
		Dim CipherKey(7)
		
		m_Nb = nb
		m_Nk = nk
		
		If m_Nb >= m_Nk Then
			m_Nr = 6 + m_Nb
		Else
			m_Nr = 6 + m_Nk
		End If
		
		C1 = 1
		If m_Nb < 8 Then
			C2 = 2
			C3 = 3
		Else
			C2 = 3
			C3 = 4
		End If
		
		For j = 0 To nb - 1
			m = j * 3
			
			m_fi(m) = (j + C1) Mod nb
			m_fi(m + 1) = (j + C2) Mod nb
			m_fi(m + 2) = (j + C3) Mod nb
			m_ri(m) = (nb + j - C1) Mod nb
			m_ri(m + 1) = (nb + j - C2) Mod nb
			m_ri(m + 2) = (nb + j - C3) Mod nb
		Next
		
		N = m_Nb * (m_Nr + 1)
		
		For i = 0 To m_Nk - 1
			j = i * 4
			CipherKey(i) = PackFrom(key, j)
		Next
		
		For i = 0 To m_Nk - 1
			m_fkey(i) = CipherKey(i)
		Next
		
		j = m_Nk
		k = 0
		Do While j < N
			m_fkey(j) = m_fkey(j - m_Nk) Xor _
				SubByte(RotateLeft(m_fkey(j - 1), 24)) Xor m_rco(k)
			If m_Nk <= 6 Then
				i = 1
				Do While i < m_Nk And (i + j) < N
					m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
						m_fkey(i + j - 1)
					i = i + 1
				Loop
			Else
				i = 1
				Do While i < 4 And (i + j) < N
					m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
						m_fkey(i + j - 1)
					i = i + 1
				Loop
				If j + 4 < N Then
					m_fkey(j + 4) = m_fkey(j + 4 - m_Nk) Xor _
						SubByte(m_fkey(j + 3))
				End If
				i = 5
				Do While i < m_Nk And (i + j) < N
					m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
						m_fkey(i + j - 1)
					i = i + 1
				Loop
			End If
			
			j = j + m_Nk
			k = k + 1
		Loop
		
		For j = 0 To m_Nb - 1
			m_rkey(j + N - nb) = m_fkey(j)
		Next
		
		i = m_Nb
		Do While i < N - m_Nb
			k = N - m_Nb - i
			For j = 0 To m_Nb - 1
				m_rkey(k + j) = InvMixCol(m_fkey(i + j))
			Next
			i = i + m_Nb
		Loop
		
		j = N - m_Nb
		Do While j < N
			m_rkey(j - N + m_Nb) = m_fkey(j)
			j = j + 1
		Loop
	End Sub

	Private Sub encrypt(buff())
		Dim i
		Dim j
		Dim k
		Dim m
		Dim a(7)
		Dim b(7)
		Dim x
		Dim y
		Dim t
		
		For i = 0 To m_Nb - 1
			j = i * 4
			
			a(i) = PackFrom(buff, j)
			a(i) = a(i) Xor m_fkey(i)
		Next
		
		k = m_Nb
		x = a
		y = b
		
		For i = 1 To m_Nr - 1
			For j = 0 To m_Nb - 1
				m = j * 3
				y(j) = m_fkey(k) Xor m_ftable(x(j) And m_lOnBits(7)) Xor _
					RotateLeft(m_ftable(RShift(x(m_fi(m)), 8) And m_lOnBits(7)), 8) Xor _
					RotateLeft(m_ftable(RShift(x(m_fi(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
					RotateLeft(m_ftable(RShift(x(m_fi(m + 2)), 24) And m_lOnBits(7)), 24)
				k = k + 1
			Next
			t = x
			x = y
			y = t
		Next
		
		For j = 0 To m_Nb - 1
			m = j * 3
			y(j) = m_fkey(k) Xor m_fbsub(x(j) And m_lOnBits(7)) Xor _
				RotateLeft(m_fbsub(RShift(x(m_fi(m)), 8) And m_lOnBits(7)), 8) Xor _
				RotateLeft(m_fbsub(RShift(x(m_fi(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
				RotateLeft(m_fbsub(RShift(x(m_fi(m + 2)), 24) And m_lOnBits(7)), 24)
			k = k + 1
		Next
		
		For i = 0 To m_Nb - 1
			j = i * 4
			UnpackFrom y(i), buff, j
			x(i) = 0
			y(i) = 0
		Next
	End Sub

	Private Sub decrypt(buff())
		Dim i
		Dim j
		Dim k
		Dim m
		Dim a(7)
		Dim b(7)
		Dim x
		Dim y
		Dim t
		
		For i = 0 To m_Nb - 1
			j = i * 4
			a(i) = PackFrom(buff, j)
			a(i) = a(i) Xor m_rkey(i)
		Next
		
		k = m_Nb
		x = a
		y = b
		
		For i = 1 To m_Nr - 1
			For j = 0 To m_Nb - 1
				m = j * 3
				y(j) = m_rkey(k) Xor m_rtable(x(j) And m_lOnBits(7)) Xor _
					RotateLeft(m_rtable(RShift(x(m_ri(m)), 8) And m_lOnBits(7)), 8) Xor _
					RotateLeft(m_rtable(RShift(x(m_ri(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
					RotateLeft(m_rtable(RShift(x(m_ri(m + 2)), 24) And m_lOnBits(7)), 24)
				k = k + 1
			Next
			t = x
			x = y
			y = t
		Next
		
		For j = 0 To m_Nb - 1
			m = j * 3
			
			y(j) = m_rkey(k) Xor m_rbsub(x(j) And m_lOnBits(7)) Xor _
				RotateLeft(m_rbsub(RShift(x(m_ri(m)), 8) And m_lOnBits(7)), 8) Xor _
				RotateLeft(m_rbsub(RShift(x(m_ri(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
				RotateLeft(m_rbsub(RShift(x(m_ri(m + 2)), 24) And m_lOnBits(7)), 24)
			k = k + 1
		Next
		
		For i = 0 To m_Nb - 1
			j = i * 4
			
			UnpackFrom y(i), buff, j
			x(i) = 0
			y(i) = 0
		Next
	End Sub

	Private Function IsInitialized(vArray)
		On Error Resume Next
		
		IsInitialized = IsNumeric(UBound(vArray))
	End Function

	Private Sub CopyBytesASP(bytDest, lDestStart, bytSource(), lSourceStart, lLength)
		Dim lCount
		
		lCount = 0
		Do
			bytDest(lDestStart + lCount) = bytSource(lSourceStart + lCount)
			lCount = lCount + 1
		Loop Until lCount = lLength
	End Sub

	Private Function EncryptData(bytMessage, bytPassword)
		Dim bytKey(31)
		Dim bytIn()
		Dim bytOut()
		Dim bytTemp(31)
		Dim lCount
		Dim lLength
		Dim lEncodedLength
		Dim bytLen(3)
		Dim lPosition
		
		If Not IsInitialized(bytMessage) Then
			Exit Function
		End If
		If Not IsInitialized(bytPassword) Then
			Exit Function
		End If
		
		For lCount = 0 To UBound(bytPassword)
			bytKey(lCount) = bytPassword(lCount)
			If lCount = 31 Then
				Exit For
			End If
		Next
		
		gentables
		gkey 8, 8, bytKey
		
		lLength = UBound(bytMessage) + 1
		lEncodedLength = lLength + 4
		
		If lEncodedLength Mod 32 <> 0 Then
			lEncodedLength = lEncodedLength + 32 - (lEncodedLength Mod 32)
		End If
		ReDim bytIn(lEncodedLength - 1)
		ReDim bytOut(lEncodedLength - 1)
		
		Unpack lLength, bytIn
		CopyBytesASP bytIn, 4, bytMessage, 0, lLength

		For lCount = 0 To lEncodedLength - 1 Step 32
			CopyBytesASP bytTemp, 0, bytIn, lCount, 32
			Encrypt bytTemp
			CopyBytesASP bytOut, lCount, bytTemp, 0, 32
		Next
		
		EncryptData = bytOut
	End Function

	Private Function DecryptData(bytIn, bytPassword)
		Dim bytMessage()
		Dim bytKey(31)
		Dim bytOut()
		Dim bytTemp(31)
		Dim lCount
		Dim lLength
		Dim lEncodedLength
		Dim bytLen(3)
		Dim lPosition
		
		If Not IsInitialized(bytIn) Then
			Exit Function
		End If
		If Not IsInitialized(bytPassword) Then
			Exit Function
		End If
		
		lEncodedLength = UBound(bytIn) + 1
		
		If lEncodedLength Mod 32 <> 0 Then
			Exit Function
		End If
		
		For lCount = 0 To UBound(bytPassword)
			bytKey(lCount) = bytPassword(lCount)
			If lCount = 31 Then
				Exit For
			End If
		Next
		
		gentables
		gkey 8, 8, bytKey

		ReDim bytOut(lEncodedLength - 1)
		
		For lCount = 0 To lEncodedLength - 1 Step 32
			CopyBytesASP bytTemp, 0, bytIn, lCount, 32
			Decrypt bytTemp
			CopyBytesASP bytOut, lCount, bytTemp, 0, 32
		Next

		lLength = Pack(bytOut)
		
		If lLength > lEncodedLength - 4 Then
			Exit Function
		End If
		
		ReDim bytMessage(lLength - 1)
		CopyBytesASP bytMessage, 0, bytOut, 4, lLength
		
		DecryptData = bytMessage
	End Function

	
	
	
	Public Function AESEncrypt(sPlain, sPassword)
		Dim bytIn()
		Dim bytOut
		Dim bytPassword()
		Dim lCount
		Dim lLength
		Dim sTemp
		
		lLength = Len(sPlain)
		ReDim bytIn(lLength-1)
		For lCount = 1 To lLength
			bytIn(lCount-1) = CByte(AscB(Mid(sPlain,lCount,1)))
		Next
		lLength = Len(sPassword)
		ReDim bytPassword(lLength-1)
		For lCount = 1 To lLength
			bytPassword(lCount-1) = CByte(AscB(Mid(sPassword,lCount,1)))
		Next

		bytOut = EncryptData(bytIn, bytPassword)

		sTemp = ""
		For lCount = 0 To UBound(bytOut)
			sTemp = sTemp & Right("0" & Hex(bytOut(lCount)), 2)
		Next

		AESEncrypt = sTemp
	End Function

	Public Function AESDecrypt(sCypher, sPassword)
		Dim bytIn()
		Dim bytOut
		Dim bytPassword()
		Dim lCount
		Dim lLength
		Dim sTemp
		
		lLength = Len(sCypher)
		ReDim bytIn(lLength/2-1)
		For lCount = 0 To lLength/2-1
			bytIn(lCount) = CByte("&H" & Mid(sCypher,lCount*2+1,2))
		Next
		lLength = Len(sPassword)
		ReDim bytPassword(lLength-1)
		For lCount = 1 To lLength
			bytPassword(lCount-1) = CByte(AscB(Mid(sPassword,lCount,1)))
		Next

		bytOut = DecryptData(bytIn, bytPassword)

		lLength = UBound(bytOut) + 1
		sTemp = ""
		For lCount = 0 To lLength - 1
			sTemp = sTemp & Chr(bytOut(lCount))
		Next

		AESDecrypt = sTemp
	End Function

end class
'**************************************************************************************************************
'**************************************************************************************************************


'**************************************************************************************************************
'**************************************************************************************************************
%>