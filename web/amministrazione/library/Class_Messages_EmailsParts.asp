
<%

'dichiara o imposta la configurazione base per la spedizione dei messaggi di posta.
Private Sub EmailMessage_LoadConfiguration()
	
	'crea oggetti per gestione posta
	if IsObject(Application("Class_mailer_Configuration")) then
		Set Configuration = Application("class_mailer_configuration")
	else
		Set Configuration = Server.CreateObject("CDO.Configuration")
		'configurazione di base messaggio
		with Configuration.Fields
			.Item(cdoSMTPServer) = Request.ServerVariables("SERVER_NAME")
			.Item(cdoNNTPAuthenticate) = cdoAnonymous
			.Item(cdoSendUsingMethod) = cdoSendUsingPort
			.Item(cdoURLGetLatestVersion) = true
			.update
		end with
	end if
	
end sub

'carica il corpo del messaggio direttamente dalla pagina sito
Private Sub EmailMessage_PageLoadHtml(objMessage, pageId, QueryString)
	
	dim PageUrl, SiteUrl
	PageUrl = GetPageURL(NULL, pageId)
	SiteUrl = ExtractPageBaseUrl(PageUrl)
	
	if cString(QueryString)<>"" then
		if left(Trim(QueryString), 1) <> "&" then
			PageUrl = PageUrl & "&"
		end if
		PageUrl = PageUrl & QueryString
	end if
	
	CALL EmailMessage_LoadHTML(objMessage, PageUrl, SiteUrl)
	
end sub


'carica il corpo del messaggio direttamente dalla pagina sito
Private Sub EmailMessage_PageSiteLoadHtml(objMessage, pageSiteId, QueryString, Lingua)
	
	CALL EmailMessage_PageLoadHtml(objMessage, GetPageByLanguage(NULL, NULL, pageSiteId, LINGUA), QueryString)

end sub


Function ReplaceAllByExpression(StringToExtract, MatchPattern, ReplacementText)
	Dim regEx, CurrentMatch, CurrentMatches
	Set regEx = New RegExp
	regEx.Pattern = MatchPattern
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True
	StringToExtract = regEx.Replace(StringToExtract, ReplacementText)
	Set regEx = Nothing
	ReplaceAllByExpression = StringToExtract
End Function



'carica il corpo del messaggio da un url ( FUNZIONA SOLO PER EMAIL E FAX )
Private Sub EmailMessage_LoadHTML(objMessage, byval URL, contentBaseURL)
	dim HTML

	'prepara url richiesto per gestione interna
	URL = replace(URL, "https", "http")
	if instr(1, URL, "dynalay", vbTextCompare)>0 or _
	   instr(1, URL, "mail", vbTextCompare)>0 or _
	   instr(1, URL, "default.asp", vbTextCompare)>0 or _
	   instr(1, URL, "default.aspx", vbTextCompare)>0 then
		URL = EncodeUrlForEmail(URL)
	end if
	if IsPreMailerRenderingActive() then
		'codice con premailer attivo
		HTML = ExecuteHttpUrl(EncodeCssInlinedUrl(URL))
		HTML = ABSOLUTIZE_LINKS(ContentBaseURL, HTML, 1)
		objMessage.HTMLBody = HTML
	else
		
		'codice standard framework
		'dim sBody

		'objMessage.CreateMHTMLBody URL, CdoSuppressAll  'Giacomo - commentato il 20/03/2013 e sostituito con le due righe successive
		HTML = ExecuteHttpUrl(URL)
		'objMessage.HTMLBody = HTML

		'recupera corpo in HTML del messaggio
		'set sBody = objMessage.HTMLBodyPart.GetDecodedContentStream
		'aggiorna i link e rende assoluti i dati
		'HTML = sBody.readText
		HTML = ABSOLUTIZE_LINKS(ContentBaseURL, HTML, 1)

		'ricerco la width del tag form per applicarlo al tag table che aggiungerò
		dim style_width, style_body
		style_width = Right(HTML, Len(HTML) - InStr(HTML, "<form"))
		style_width = Left(style_width, InStr(style_width, ">"))
		if inStr(style_width, "style=")>0 then
			style_width = Right(style_width, Len(style_width) - (InStr(style_width, "style=")+5))
			style_width = Left(style_width, InStr(2, style_width, """", 0))
			style_width = "style="&style_width
		else
			style_width = ""
		end if

		'tolgo il tag form
		CALL ReplaceAllByExpression(HTML, "<form[^<]*?>", "<table align=""left"" cellspacing=""0"" cellspadding=""0""><tr><td class=""nextform_div"" "&style_width&">")
		HTML = Replace(HTML, "</form>", "</td></tr></table>")
		CALL ReplaceAllByExpression(HTML, "<input type=""hidden"" name=""__VIEWSTATE""[^<]*?>", "<input type=""hidden"" value="""">")

		'svuota lo stream che punta al conenuto HTML del messaggio
		'sbody.position = 0
		'sbody.setEOS
		'riscrive il contenuto della parte di messaggio HTML
		'sbody.Type = adTypeText
		'sbody.charset = "UTF-8" 'Giacomo - aggiunto il 20/03/2013
		'sbody.WriteText HTML
		'sbody.flush

		objMessage.HTMLBody = HTML
	end if
	
end sub


'FUNZIONI SPECIFICHE PER FAX ED EMAIL ************************
	
'.................................................................................................
'	che rende assoluti tutti gli url dell'HTML passato in base al parametro url_base
'	level deve essere impostata a 1
'.................................................................................................
private function ABSOLUTIZE_LINKS(url_base, HTML, level)
       dim char, char_to_skip, skipped_charset
	if level = 0 then level = 1
	select case level
		case 1
			char_to_skip = "H"
			skipped_charset = ""
               HTML = replace(HTML, "href=""?", "href=""" & GetNextWebPageFileName(NULL) & "?",1,-1,vbTextCompare)
               HTML = IntelligentReplace(HTML, "<script", "/script>", "")
			HTML = AttributeReplace(HTML, "onsubmit", "")
			HTML = AttributeReplace(HTML, "onclick", "")
			HTML = AttributeReplace(HTML, "onchange", "")
			HTML = AttributeReplace(HTML, "onload", "")
		case 2
			char_to_skip = "T"
			skipped_charset = "H"
		case 3
			char_to_skip = "T"
			skipped_charset = "HT"
		case 4
			char_to_skip = "P"
			skipped_charset = "HTT"
		case 5
			char_to_skip = ":"
			skipped_charset = "HTTP"
		case else
			ABSOLUTIZE_LINKS = HTML
			Exit function
	end select
	for char= Asc("/") to Asc("Z")
		if char<>Asc(char_to_skip) and (char<=asc("9") or char >= asc("A")) then
			'sostituzione href dei link
			HTML = replace(HTML, "href=" & skipped_charset & "" & chr(char), "href=" & url_base & skipped_charset & chr(char),1,-1,vbTextCompare)
			HTML = replace(HTML, "href='" & skipped_charset & "" & chr(char), "href='" & url_base & skipped_charset & chr(char),1,-1,vbTextCompare)
			HTML = replace(HTML, "href=""" & skipped_charset & "" & chr(char), "href=""" & url_base & skipped_charset & chr(char),1,-1,vbTextCompare)
			
			'sostituzione sorgenti di immagine
			HTML = replace(HTML, "src=" & skipped_charset & "" & chr(char), "src=" & url_base & skipped_charset & chr(char),1,-1,vbTextCompare)
			HTML = replace(HTML, "src='" & skipped_charset & "" & chr(char), "src='" & url_base & skipped_charset & chr(char),1,-1,vbTextCompare)
			HTML = replace(HTML, "src=""" & skipped_charset & "" & chr(char), "src=""" & url_base & skipped_charset & chr(char),1,-1,vbTextCompare)
		end if
	next
	
	'corregge errore nel link delle email
	HTML = replace(HTML, "href=" & url_base & "mailto:", "href=mailto:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "href='" & url_base & "mailto:", "href='mailto:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "href=""" & url_base & "mailto:", "href=""mailto:", 1, -1, vbTextCompare)	
	'corregge errore nei link ai javascript
	HTML = replace(HTML, "href=" & url_base & "javascript:", "href=javascript:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "href='" & url_base & "javascript:", "href='javascript:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "href=""" & url_base & "javascript:", "href=""javascript:", 1, -1, vbTextCompare)
	'corregge errore nel link con https
	HTML = replace(HTML, "href=" & url_base & "https:", "href=https:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "href='" & url_base & "https:", "href='https:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "href=""" & url_base & "https:", "href=""https:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "src=" & url_base & "https:", "src=https:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "src='" & url_base & "https:", "src='https:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "src=""" & url_base & "https:", "src=""https:", 1, -1, vbTextCompare)
	'corregge errore nel link con callto
	HTML = replace(HTML, "href=" & url_base & "callto:", "href=callto:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "href='" & url_base & "callto:", "href='callto:", 1, -1, vbTextCompare)
	HTML = replace(HTML, "href=""" & url_base & "callto:", "href=""callto:", 1, -1, vbTextCompare)
	
	ABSOLUTIZE_LINKS = ABSOLUTIZE_LINKS(url_base, HTML, level+1)
	
end function
	
	
'.................................................................................................
'	funzione che esegue un replace all'interno di una stringa della porzione di testo compreso tra 
'	la stringa s_begin e l's_end successivo.
'.................................................................................................
private function IntelligentReplace(str, s_begin, s_end, replacer)
	dim begins, ends
	begins = 1
	do while begins > 0
		begins = instr(1, str, s_begin, vbTextCompare)
		if begins > 0 then
			ends = instr(begins, str, s_end, vbTextCompare)
			str = left(str, begins -1) + _
				  replacer + _
				  right(str, len(str) - (ends + len(s_end) - 1 ))
		else
			exit do
		end if
	loop
	IntelligentReplace = str
end function


'.................................................................................................
'	ripulisce la stringa dall'attributo ed i relativi valori.
'.................................................................................................
private function AttributeReplace(str, AttributeName, replacer)
	dim begins, ends, delimiter
	begins = 1
	AttributeName = AttributeName & "="
	do while begins > 0
		begins = instr(begins + 1, str, AttributeName, vbTextCompare)
		if begins > 0 then
			delimiter = Mid(str, begins + len(AttributeName), 1)
			ends = instr(begins + len(AttributeName) + 1, str, delimiter, vbTextCompare)
			str = left(str, begins -1) + _
				  replacer + _
				  right(str, len(str) - ends)
		else
			exit do
		end if
	loop
	AttributeReplace = str
end function

%>