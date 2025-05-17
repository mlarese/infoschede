<%

class Faxer
	
	Public Configuration
	
	Private senderID
	Private faxID
	Private faxMittente
	Private faxFilePath
	Private faxHtml
	
	'email mittente con cui vengono inviati i fax
	Private faxSenderEmail
	'dominio a cui spedire le email dei fax
	Private faxSenderDomain
	
	Private Sub Class_Initialize()
		
		'crea oggetti per gestione posta
		CALL EmailMessage_LoadConfiguration()
		
		if Session("FAX_SENDER_EMAIL")="" OR Session("FAX_SENDER_DOMAIN")="" then
			'recupera login e password per spedizione.
			dim conn
			set conn = Server.CreateObject("ADODB.Connection")
			conn.open Application("DATA_ConnectionString")
			faxSenderEmail = GetModuleParam(conn, "FAX_SENDER_EMAIL")
			faxSenderDomain = GetModuleParam(conn, "FAX_SENDER_DOMAIN")
			conn.close
			set conn = nothing
		else
			faxSenderEmail = Session("FAX_SENDER_EMAIL")
			faxSenderDomain = Session("FAX_SENDER_DOMAIN")
		end if
		
	End Sub
	
		
	Private Sub Class_Terminate()
		set Configuration = nothing
	End Sub
	
'DEFINIZIONE PRORIETA' STATICHE DEL MESSAGGIO *******************
	Public Property Get MimeType()
		if faxHtml<>"" then
			MimeType = MIME_HTML
		else
			MimeType = MIME_TEXT
		end if
	End Property
	
	Public Property Get MessageType()
    	MessageType = MSG_FAX
	End Property
	
	Public Property Get RecipientType()
    	RecipientType = VAL_FAX
	End Property
	
	
'DEFINIZIONE PRORIETA'*******************
	'oggetto
	Public Property Get Subject()
    	Subject = ""
	End Property
	
	'testo del messaggio
	Public Property Get Body()
    	Body = faxFilePath
	End Property
	
	'testo del messaggio
	Public Property Let Body(value)
    	faxFilePath = value
		if faxFilePath <> "" then
			faxHtml = ""
		end if
	End Property
	
	'testo html messaggio
	Public Property Get HtmlBody()
    	HtmlBody = faxHtml
	End Property
	
	'testo del messaggio
	Public Property Let HtmlBody(value)
    	faxHtml = value
		if faxHtml <> "" then
			faxFilePath = ""
		end if
	End Property
	
	'id del messaggio salvato
	Public Property Get MessageId()
    	MessageId = faxId
	End Property
	
	Public Property Let MessageId(id)
		FaxId = id
	End Property
	
	
'DEFINIZIONE METODI:*********************
	
	
	Public sub Send(faxNumber)
		dim message
		Set message = Server.CreateObject("CDO.Message")
		'imposta configurazione messaggio di posta
		set message.Configuration = Configuration
		
		'imposta corpo del messaggio per TAUFI
		message.subject = "invio fax  - " + GetCurrentBaseUrl()
		message.sender = faxSenderEmail
		message.to = FormatFax(faxNumber) + "@" + faxSenderDomain
		
		if MimeType = MIME_HTML then
			dim fso, tmpfile, tmpfilepath
			Set Fso = CreateObject("Scripting.FileSystemObject")
			tmpfilepath = Application("IMAGE_PATH") & "/temp/" & Session.SessionId & "_" & cIntero(faxId) & ".htm"
			set tmpfile = fso.CreateTextFile(tmpfilepath, True)
			tmpfile.write(faxHtml)
			tmpfile.close
			Message.AddAttachment tmpfilepath
			fso.Deletefile tmpfilepath, true
		else
			Message.AddAttachment faxFilePath
		end if
		
		'spedisce messaggio
		message.Send
		set message = nothing
	end sub
	
	
	Public Function SendOnError(faxNumber)
		if IsPhoneNumber( faxNumber )then
			On Error Resume Next
			CALL Send(faxNumber)
			SendOnError = err.number
			On error goto 0
		else
			SendOnError = 1
		end if
	End Function
	
	
	'carica indirizzo di spedizione da record dell'utente
	Public Sub LoadSenderByID(conn, rs, admin_id)	
		dim sql
		sql = "SELECT admin_fax FROM tb_admin WHERE id_admin=" & cInteger(admin_id)
		faxMittente = GetValueList(conn, rs, sql)
		senderID = admin_id
	end sub
	
	
	Public Function FormatFax(faxNumber)
		faxNumber = RemoveInvalidChar(faxNumber, "0123456789")
		if left(faxNumber, 2) = "00" then
			FormatFax = right(faxNumber, len(faxNumber)-2)
		else
			FormatFax = "39" + faxNumber
		end if
	end function
	
	
	'carica il corpo del messaggio da un url
	Public sub LoadHTML(byval URL, contentBaseURL)
		dim oMessage
		Set oMessage = Server.CreateObject("CDO.Message")
		set oMessage.Configuration = Configuration
		CALL EmailMessage_LoadHTML(oMessage, URL, contentBaseURL)
		faxHtml = oMessage.HtmlBody
		set oMessage = nothing
	end sub
	
	
	Public sub SaveAttachments(faxId)
		
		if faxFilePath<>"" then
			
			dim docsPath, fso
			docsPath = Application("IMAGE_PATH") & "\docs\eml_" & faxId & "\"
		
			'verifica se esiste la directory, eventualmente la crea
			Set Fso = CreateObject("Scripting.FileSystemObject")
			if not Fso.FolderExists(docsPath) then
				CALL Fso.CreateFolder(docsPath)
			end if
			CALL Fso.CopyFile(faxFilePath, docsPath, true)
			set Fso = nothing
			
		end if
		
	end sub
	
	'*******************************************************************************
	'INCLUSIONE METODI COMUNI A TUTTE LE CLASI
	'*******************************************************************************
	%>
	<!--#INCLUDE FILE="Class_Messages_CommonParts.asp" -->
<%	
	'*******************************************************************************
	
	'*******************************************************************************
	'INCLUSIONE METODI COMUNI ALLE CLASSI CHE USANO LE EMAIL
	'*******************************************************************************
	%>
	<!--#INCLUDE FILE="Class_Messages_EmailsParts.asp" -->
<%	
	'*******************************************************************************
	
end class

%>