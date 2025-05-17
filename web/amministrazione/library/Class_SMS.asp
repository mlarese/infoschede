<%

class SMSSender
	
	public Messaggio
	public SmsID
	public SenderID
	
	private SMSuser, SMSpass, CELLmittente
	
	
	Private Sub Class_Initialize()
		CELLmittente=""
		Messaggio = ""
		SenderID=0
		
		if Session("SMS_LOGIN")="" OR Session("SMS_PASSWORD")="" then
			'recupera login e password per spedizione.
			dim conn
			set conn = Server.CreateObject("ADODB.Connection")
			conn.open Application("DATA_ConnectionString")
			SMSuser = GetModuleParam(conn, "SMS_LOGIN")
			SMSpass = GetModuleParam(conn, "SMS_PASSWORD")
			conn.close
			set conn = nothing
		else
			SMSuser = Session("SMS_LOGIN")
			SMSpass = Session("SMS_PASSWORD")
		end if
	End Sub
	
	
	Private Sub Class_Terminate()
	
	End Sub
	
	
'DEFINIZIONE PRORIETA' STATICHE DEL MESSAGGIO *******************
	
	Public Property Get MimeType()
		MimeType = MIME_TEXT
	End Property
	
	Public Property Get MessageType()
    	MessageType = MSG_SMS
	End Property
	
	Public Property Get RecipientType()
    	RecipientType = VAL_CELLULARE
	End Property
	
	
'DEFINIZIONE PRORIETA'*******************
	'oggetto
	Public Property Get Subject()
    	Subject = ""
	End Property
	
	'testo del messaggio
	Public Property Get Body()
    	Body = Messaggio
	End Property
	
	Public Property Let Body(text)
		Messaggio = cString(text)
	End Property
	
	'testo html messaggio
	Public Property Get HtmlBody()
    	HtmlBody = ""
	End Property
	
	'id del messaggio salvato
	Public Property Get MessageId()
    	MessageId = SmsId
	End Property
	
	Public Property Let MessageId(id)
		SmsId = id
	End Property
	
'DEFINIZIONE METODI:*********************
	
	
	Public Sub Send(cellulare)
		
		dim strPostData, ritorno, xml
		strPostData = "user=" & Server.URLEncode(SMSuser) & _
					  "&pass=" & Server.URLEncode(SMSpass) & _
					  "&rcpt=" & Server.URLEncode(FormatMobilePhone(cellulare)) & _
					  "&data=" & Server.URLEncode(Messaggio) & _
					  "&sender=" & Server.URLEncode(FormatMobilePhone(CELLmittente))
		Set xml = Server.CreateObject("Microsoft.XMLHTTP")
		xml.Open "POST", "http://sms.tol.it/sms/send.php", False
		xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xml.Send strPostData
'		response.write "<pre>"+vbCRLF
'		response.write strPostData
'		response.write "</pre>"+vbCRLF
		ritorno = xml.responseText
'		response.write "<pre>"+vbCRLF
'		response.write ritorno
'		response.write "</pre>"+vbCRLF
'		response.end
		if instr(1, ritorno, "KO", vbTextCompare)>0 then
			'errore di spedizione: genera errore
			CALL Err.raise(10000, "Class_SMS.asp", "Errore nella spedizione dell'sms: " + ritorno, "")
		end if
		Set xml = Nothing
		
	end Sub
	
	
	Public Function SendOnError(cellulare)
		if IsPhoneNumber( cellulare )then
			On Error Resume Next
			CALL Send(cellulare)
			SendOnError = err.number
			On error goto 0
		else
			SendOnError = 1
		end if
	End Function
	
	
	Public sub LoadHTML(byval URL, contentBaseURL)
		
	end sub
	
	
	'carica indirizzo di spedizione da record dell'utente
	Public Sub LoadSenderByID(conn, rs, admin_id)	
			dim sql
			sql = "SELECT admin_cell FROM tb_admin WHERE id_admin=" & cInteger(admin_id)
			CELLmittente = GetValueList(conn, rs, sql)
			SenderID = admin_id
	end sub
	
	'*******************************************************************************
	'INCLUSIONE METODI COMUNI A TUTTE LE CLASI
	'*******************************************************************************
	%>
	<!--#INCLUDE FILE="Class_Messages_CommonParts.asp" -->
	<%	
	'*******************************************************************************
	
end class

%>