
<!--#INCLUDE FILE="../nextWeb5/Tools_ClassCssManager.asp" -->

<%
class Mailer
	
	Public Configuration
	Public Message
	
	Private SenderID
	Private EmailID
	
	Private Sub Class_Initialize()
		
		Set Message = Server.CreateObject("CDO.Message")
		CALL EmailMessage_LoadConfiguration()
		
	End Sub
		
	Private Sub Class_Terminate()
		set Configuration = nothing
		set Message = nothing
	End Sub
	
	
'DEFINIZIONE METODI:*********************
	
	Public sub Send()
		dim debug_mail
		dim to_email
		debug_mail = false
		'imposta configurazione messaggio
		set Message.Configuration = Configuration
		
		if debug_mail then
			
			to_email = Trim(Message.To) 
			to_email = RemoveInvalidChar(to_email,ALPHANUMERIC_CHARSET)
			Message.To = to_email & "@vmmultimedia.com"
			Message.CC = ""
			Message.BCC = ""
		end if
		'spedisce messaggio
		Message.Send
		
		'Response.Write "To: " + Message.To + "<br>"
		'Response.write "CC: "+ Message.CC + "<br>"
		'Response.write "BCC: " + Message.BCC + "<br>"
		'Response.end
	end sub
	
	
	Public Function SendOnError(email)
		message.to = email
		if IsEmail(email) then
			On Error Resume Next
			CALL Send()
			SendOnError = err.number
			
			CALL WriteLogAdmin(null, "tb_email", 0, "SendOnError", "Spedizione fallita" +vbCrLf + LastErrorDump())
																		
			On error goto 0
		else
			SendOnError = 1
		end if
	End Function
	
	
	'costruisce il corpo HTML completo, dato il codice HTML da inserire nel body
	Public function LoadHTMLCode(conn, title, HTMLBody, tagBody)
		dim html
		html = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""> " & vbCRLF & _
			   "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCRLF & _
			   "	<head lang=""it"">" & vbCRLF & _
			   "		<title>" & title & "</title>" & vbCRLF & _
			   "		<meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" />" & vbCRLF & _
			   "		<meta name=""robots"" content=""noindex,nofollow"" />" & vbCRLF & _
			   " 		<meta name=""copyright"" content=""Copyright © " & Year(Date()) & " - Next-aim"" />" & vbCRLF & _
			   "		<style>" & vbCRLF
				dim stili
				set stili = new CssManager
				html = html & Replace(stili.GetCssStandard(conn, Application("AZ_ID")), """", "'")		
				html = html & "</style>" & vbCRLF & _
			   "	</head>" & vbCRLF
		if tagBody <> "" then
			html = html & tagBody & vbCRLF
		else
			html = html & "	<body>" & vbCRLF
		end if
		html =  html & "		<div class=""layers_text_s"">" & vbCRLF & _
			   "			" & HTMLBody & vbCRLF & _
			   "		</div>" & vbCRLF & _
			   "	</body>" & vbCRLF & _
			   "</html>"	
		LoadHTMLCode = html
	end function
	
	'carica il corpo del messaggio da un url
	Public sub LoadHTML(byval URL, contentBaseURL)
		CALL EmailMessage_LoadHTML(Message, URL, contentBaseURL)
	end sub
	
	'carica il corpo del messaggio da una pagina del next-web
	Public sub PageLoadHTML(pageId, QueryString)
		CALL EmailMessage_PageLoadHtml(Message, pageId, QueryString)
	end sub
	
	'carica il corpo del messaggio da una pagina sito del next-web
	Public sub PageSiteLoadHTML(pageSiteId, QueryString, lingua)
		CALL EmailMessage_PageSiteLoadHtml(Message, pageSiteId, QueryString, Lingua)
	end sub
	
	
	'carica gli allegati da una lista separata da ";"
	Public sub LoadAttachments(attachmentsList, emailId)
		dim docsPath, docsList, i
		
		if cIntero(emailId)>0 then
			docsPath = Application("IMAGE_PATH") & "\docs\eml_" & emailId & "\"
		else
			docsPath = Application("IMAGE_PATH") & "\temp\"
		end if
		
		if cString(attachmentsList)<>"" then
			docsList = split(attachmentsList, ";")
			for i = lbound(docsList) to ubound(docsList)
				docsList(i) = Trim(docsList(i))
				if docsList(i)<>"" then
					Message.AddAttachment docsPath & docsList(i)
				end if
			next
		end if
		
	end sub
	
	
	Public sub SaveAttachments(emailId)
		dim docsPath, Attachment, fso
		docsPath = Application("IMAGE_PATH") & "\docs\eml_" & emailId & "\"
		
		'verifica se esiste la directory, eventualmente la crea
		Set Fso = CreateObject("Scripting.FileSystemObject")
		if not Fso.FolderExists(docsPath) then
			CALL Fso.CreateFolder(docsPath)
		end if
		set Fso = nothing
		
		'salva gli allegati nella corretta cartella dell'email
		for each Attachment in message.Attachments
			CALL Attachment.SaveToFile(docsPath + Attachment.FileName)
		next
	end sub
	
	
	'carica indirizzo di spedizione da record dell'utente
	Public Sub LoadSenderByID(conn, rs, admin_id)	
		dim sql
		sql = "SELECT admin_email, admin_email_newsletter FROM tb_admin WHERE id_admin=" & cInteger(admin_id)
		rs.open sql, conn, adopenstatic, adLockOptimistic, adCmdText
		if not rs.eof then
			if IsEmail(rs("admin_email_newsletter")) then
				Message.From = cString(rs("admin_email_newsletter"))
			else
				Message.From = cString(rs("admin_email"))
			end if
		end if
		rs.close
		SenderID = admin_id
	end sub
	
	
'DEFINIZIONE PRORIETA' STATICHE DEL MESSAGGIO *******************
	Public Property Get MimeType()
		if Message.HTMLbody<>"" then
	    	MimeType = MIME_HTML
		else
	    	MimeType = MIME_TEXT
		end if
	End Property
	
	Public Property Get MessageType()
    	MessageType = MSG_EMAIL
	End Property
	
	Public Property Get RecipientType()
    	RecipientType = VAL_EMAIL
	End Property
	
	
'DEFINIZIONE PRORIETA'*******************
	'Mittente
	Public Property Get Sender()
    	Sender = Message.From
	End Property
	
	Public Property Let Sender(email)
		Message.From = email
	End Property
		
	'destinatario
	Public Property Get Dest()
    	Dest = Message.To
	End Property
	
	Public Property Let Dest(email)
		Message.To = email
	End Property
	
	'copia carbone
	Public Property Get CC()
    	CC = Message.CC
	End Property
	
	Public Property Let CC(email)
		Message.CC = email
	End Property
	
	'copia carbone nascosta
	Public Property Get bCC()
    	bCC = Message.CC
	End Property
	
	Public Property Let bCC(email)
		Message.bCC = email
	End Property
	
	'oggetto
	Public Property Get Subject()
    	Subject = Message.Subject
	End Property
	
	Public Property Let Subject(sbj)
		Message.Subject = sbj
	End Property
	
	'testo del messaggio
	Public Property Get Body()
    	Body = Message.TextBody
	End Property
	
	Public Property Let Body(text)
		Message.TextBody = text
	End Property
	
	'testo html messaggio
	Public Property Get HtmlBody()
    	HtmlBody = Message.HtmlBody
	End Property
	
	'id del messaggio salvato
	Public Property Get MessageId()
    	MessageId = EmailId
	End Property
	
	Public Property Let MessageId(id)
		EmailId = id
	End Property
	
	
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
'*******************************************************************************


'*******************************************************************************
'vedere SendPageFromAdminToContactExtended per il funzionamento
'	conn:			connessione aperta sul database
'	rs:				oggetto recordset creato
'	Config:			oggetto configurazione del plug in in cui la pagina &egrave; caricata
'	Subject:		Oggetto dell'email che verr&agrave; spedita
'   URL:			indirizzo relativo della pagina da caricare
'	SenderAdminID	ID dell'utente amministratore che spedisce l'email
'	Dest_NextCom_ID	ID del contatto destinatario dell'email
'*******************************************************************************
function SendPageFromAdminToContact(conn, rs, Config, Subject, URL, SenderAdminID, Dest_NextCom_ID, SenderBCC)
    CALL SendPageFromAdminToContactExtended(conn, rs, Config.Lingua, Subject, URL, Config.BaseUrl, SenderAdminID, Dest_NextCom_ID, SenderBCC)
end function
'*******************************************************************************


'*******************************************************************************
'funzione che invia una email ad un contatto da parte di un utente amministratore
'componendo l'email da una pagina e salvando poi la spedizione delle email nel NEXT-com
'	conn:			connessione aperta sul database
'	rs:				oggetto recordset creato
'	lingua:			lingua corrente in cui generare la pagina
'	Subject:		Oggetto dell'email che verr&agrave; spedita
'   URL:			indirizzo relativo della pagina da caricare
'   BaseUrl:        url di base del sito
'	SenderAdminID	ID dell'utente amministratore che spedisce l'email
'	Dest_NextCom_ID	ID del contatto destinatario dell'email
'*******************************************************************************
function SendPageFromAdminToContactExtended(conn, rs, Lingua, Subject, URL, BaseUrl, SenderAdminID, Dest_NextCom_ID, SenderBCC)
    if cInteger(SenderAdminID)<>0 AND cInteger(Dest_NextCom_ID)<>0 then
		
		dim OBJ_email
		set OBJ_email = new mailer
		OBJ_email.subject = Subject
		'se nell'url di generazione della pagina non c'e' la lingua la aggiunge (solo per pagine dell'editor)
		if instr(1, URL, "lingua=", vbTextCompare)<1 AND instr(1, URL, GetNextWebPageFileName(NULL), vbTextCompare)>0 then
			URL = URL + IIF(instr(1, URL, "?", vbTextCompare), "&", "?") + "lingua=" & Lingua
		end if
		
		'response.write URL & "<br>" 
		'response.write BaseUrl & "<br>" 
		'response.end
		'genera HTML da url
		OBJ_email.LoadHTML URL , BaseUrl
		
		'carica id ed email dell'utente amministratore per la spedizione
		CALL OBJ_email.LoadSenderByID(conn, rs, SenderAdminID)
		
		'salva l'email generata con relativo mittente nel NEXT-com
		CALL OBJ_email.Save(conn, rs)
		
		if SenderBCC then
			OBJ_email.Message.BCC = OBJ_email.Message.From
		end if
		
		'salva nel NEXT-com la spedizione dell'email al contatto e la spedisce
		CALL OBJ_email.SendToContact(conn, rs, Dest_NextCom_ID)
		
		set OBJ_email = nothing
		
	end if
end function



'*******************************************************************************
'funzione che invia una email testuale ad un contatto da parte di un utente amministratore
'	conn:			connessione aperta sul database
'	rs:				oggetto recordset creato
'	Subject:		Oggetto dell'email che verr&agrave; spedita
'   Text:			Corpo dell'email da inviare
'	SenderAdminID	ID dell'utente amministratore che spedisce l'email
'	Dest_NextCom_ID	ID del contatto destinatario dell'email
'*******************************************************************************
sub SendEmailTextFromAdminToContact(conn, rs, Subject, Text, SenderAdminID, Dest_NextCom_ID, SenderBCC)
	dim OBJ_email
	set OBJ_email = new mailer
	OBJ_email.subject = Subject
	OBJ_email.body = Text
	
	'carica id ed email dell'utente amministratore per la spedizione
	CALL OBJ_email.LoadSenderByID(conn, rs, SenderAdminID)
		
	'salva l'email generata con relativo mittente nel NEXT-com
	CALL OBJ_email.Save(conn, rs)
		
	if SenderBCC then
		OBJ_email.Message.BCC = OBJ_email.Message.From
	end if
		
	'salva nel NEXT-com la spedizione dell'email al contatto e la spedisce
	CALL OBJ_email.SendToContact(conn, rs, Dest_NextCom_ID)
		
	set OBJ_email = nothing
end sub


'*******************************************************************************
'funzione che invia una email testuale ad un contatto da parte di un utente amministratore
'	conn:			connessione aperta sul database
'	rs:				oggetto recordset creato
'	Subject:		Oggetto dell'email che verr&agrave; spedita
'   Text:			Corpo dell'email da inviare
'	SenderAdminID	ID dell'utente amministratore che spedisce l'email
'	Dest_Admin_IDS	Lista degli utenti amministratori che ricevono l'email
'*******************************************************************************
sub SendEmailTextFromAdminToAdmins(conn, rs, Subject, Text, SenderAdminID, Dest_Admin_IDS)
	dim OBJ_email
	set OBJ_email = new mailer
	OBJ_email.subject = Subject
	OBJ_email.body = Text
	
	'carica id ed email dell'utente amministratore per la spedizione
	CALL OBJ_email.LoadSenderByID(conn, rs, SenderAdminID)
		
	'salva l'email generata con relativo mittente nel NEXT-com
	CALL OBJ_email.Save(conn, rs)
		
	'salva nel NEXT-com la spedizione dell'email all'amministratore
	dim rsa, sql
	set rsa = Server.CreateObject("ADODB.Recordset")
	
	sql = "SELECT id_admin FROM tb_admin WHERE id_admin IN (0, " & Dest_Admin_IDS & ") "
	rsa.open sql, conn, adOpenStatic, adLockOptimistic
	
	while not rsa.eof
		
		CALL OBJ_email.SendToAdmin(conn, rs, cIntero(rsa("id_admin")))
		
		rsa.movenext
	wend
	
	rsa.close
		
	set OBJ_email = nothing
end sub




'*******************************************************************************
'funzione che dati i parametri per la generazione, spedizione e redirect dei form
'Config:		oggetto configurazione del plug in in cui la pagina &egrave; caricata
'PageEmail		numero della PaginaSito da utilizzare per la generazione dell'email
'				(opzionale se manca usa, nell'ordine, PageRedirect o ConfirmUrl)
'PageRedirect 	numero della PaginaSito da utilizzare per il redirect del browser 
'				nella pagina di conferma (opzionale, se manca usa ConfirmUrl)
'ConfirmUlr		URL base calcolato sulla pagina corrente
'EmailUrl		Parametro che restituisce l'url corretto per la generazione delle email
'RedirectUrl	Parametro che restituisce l'url corretto per il redirect del browser
'*******************************************************************************
function PrepareFormUrls(Config, PageEmail, PageRedirect, byRef EmailUrl, byRef RedirectUrl)
	'calcola url per email
	if cInteger(PageEmail)>0 then
		'c'e' la pagina specifica per l'email: compone l'url
        EmailUrl = GetPageSiteUrl(NULL, PageEmail, Config.lingua)
	else
		EmailUrl = ""
	end if

	'calcola url per redirect del browser
	if cInteger(PageRedirect)>0 then
		'c'e la pagina specifica per il browser compone l'url
        RedirectUrl = GetPageSiteUrl(NULL, PageRedirect, Config.lingua)
	else
		'non c'&egrave; la pagina specifica: rimanda alla stessa pagina
        RedirectUrl = GetPageURL(NULL, request("PAGINA"))
	end if
	
	'verifica url per email:
	if EmailUrl = "" then
		EmailUrl = RedirectUrl
	end if
	
end function
'*******************************************************************************

%>

