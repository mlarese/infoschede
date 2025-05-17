<% 

'**************************************************************************************************************************************
'DEFINIZIONE COSTANTI GLOBALI
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'funzione che resetta lo stato della lingua e relative variabili di sessione
'**************************************************************************************************************************************
sub Applicazione_RESET(lingua)
	Session.Contents.Remove("PAGINE")
	Session("LINGUA") = Left(LCase(lingua),2)
end sub


'**************************************************************************************************************************************
'funzione che verifica se lo stato della sessione e' corretto
'**************************************************************************************************************************************
sub Applicatione_CHECK(conn, rs)
	if isEmpty(session("PAGINE")) then
		CALL Applicazione_INIT(conn, rs)
	end if
end sub

'**************************************************************************************************************************************
'funzione che scrive il file di stili generali nello stream response.
'**************************************************************************************************************************************
sub WriteStili()
    dim fso, CssPath, CssFile
	set fso = Server.CreateObject("scripting.filesystemobject")
	CssPath = request.ServerVariables("APPL_PHYSICAL_PATH") & "stili.css"
	if fso.FileExists(CssPath) then
		set CssFile = fso.OpenTextFile(CssPath, 1, false)%>
			<style type="text/css">
				<%= lcase(cString(CssFile.ReadAll)) %>
			</style>
		<%CssFile.close
		set CssFile = nothing
	end if
	set fso = nothing
end sub


'**************************************************************************************************************************************
'funzione che ritorna l'id della pagina corrispondente alla pagina sito
'**************************************************************************************************************************************
function DecodePaginaSito(conn, rs, id_paginasito)
	dim sql
	'recupera record della pagina sito richiesta
	sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & cIntero(id_paginasito)
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if rs.eof then
		'non trovato: manda a pagina di errore
		DecodePaginaSito = 0
	else
		DecodePaginaSito = rs("id_pagDyn_" & IIF(Session("LINGUA") = "", LINGUA_ITALIANO, Session("LINGUA")))
	end if
	rs.close
end function


'**************************************************************************************************************************************
'procedura che aggiorna i contatori dell'applicazione
'**************************************************************************************************************************************
sub LogVisit_Application(conn)	
	dim sql, UserType
	'esegue il log solo se :		non e' gia' stato eseguito 
	'								non ha gia' i cookie attivi (vuol dire che arriva dallo stesso sito)
	'								test generali validi
	if not Session("VISIT_LOGGED") AND _
	   instr(1, Request.ServerVariables("HTTP_COOKIE"), "ASPSESSION", vbTextCompare)<1 AND _
	   ActionLoggable(NULL) then
		
		'recupera tipo di UserAgent
		UserType = GetUserAgentType()
		
		'esegue aggiornamento contatori
		sql = " UPDATE tb_webs SET " + UserType + "= (" + UserType + " + 1), contatore = (contatore + 1) " + _
			  " WHERE id_webs=" & Session("AZ_ID")
		CALL conn.execute(sql, , adExecuteNoRecords)
	end if
	
	Session("VISIT_LOGGED") = true
end sub


'**************************************************************************************************************************************
'procedura che aggiorna i contatori della pagina
'**************************************************************************************************************************************
sub LogVisit_Page(conn)
	dim sql, UserType
	
	'esegue il log solo se: test generali validi
	if ActionLoggable(NULL) then
		
		'recupera tipo di UserAgent
		UserType = GetUserAgentType()
		
		'esegue aggiornamento contatori
		sql = " UPDATE tb_pages SET " + UserType + "= (" + UserType + " + 1), contatore = (contatore + 1) " + _
			  " WHERE id_page=" & "0" & cIntero(request("PAGINA"))
		CALL conn.execute(sql, , adExecuteNoRecords)
	end if
end sub


'**************************************************************************************************************************************
'funzione che verifica la validita' della richiesta proteggendo da attacchi sql injection
'**************************************************************************************************************************************
function Security_RequestIsValid()
	
	dim Word, Field, IsValid, InvalidatedFrom
	dim ForbiddenWords_Querystring, ForbiddenWords_Form
	'rispetto all'insieme iniziale ho rimosso le chiavi ";" e "@"
	ForbiddenWords_Querystring = split("--, ;--, /*, */, @@, char, nchar, varchar, nvarchar, alter, begin, " + _
									   "cast, create, cursor, declare, delete, drop, end, exec, execute, fetch, insert, " + _
									   "kill, open, select, sys, sysobjects, syscolumns, table, update", _
									   ", ")
	ForbiddenWords_Form = split("--, ;--, /*, */, @@", ", ")
	
	IsValid = true
	InvalidatedFrom = ""
	
	for each Word in ForbiddenWords_Querystring
		Word = Trim(Word)
		
		if isValid then
			'verifica querystring
			for each Field in request.querystring
				if instr(1, cString(request.Querystring(Field)), Word, vbTextCompare) > 0 then
					isValid = false
					InvalidatedFrom = " ( querystring: """ & Field & """=""" & request.Querystring(Field) & """ )"
					exit for
				end if
			next
		else
			exit for
		end if
		
	next
	
	for each Word in ForbiddenWords_Form
		Word = Trim(Word)
		
		if isValid then
			'verifica form
			for each Field in request.Form
				if instr(1, cString(request.Form(Field)), Word, vbTextCompare) > 0 then
					isValid = false
					InvalidatedFrom = " ( form: """ & Field & """=""" & request.Form(Field) & """ )"
					exit for
				end if
			next
		else
			exit for
		end if
		
	next
	
	if not IsValid then
		Application("INVALID_REQUESTS_COUNT") = cIntero(Application("INVALID_REQUESTS_COUNT")) + 1
		Application("INVALID_REQUESTS_LIST") = cString(Application("INVALID_REQUESTS_LIST")) + _
											   GetCurrentUrl() + "?" + Request.ServerVariables("QUERY_STRING") + vbCrLf + _
											   vbTab + InvalidatedFrom + vbCrLf + vbCrLf
		
		if (Application("INVALID_REQUESTS_COUNT") mod 10) = 0 then
			'invia avviso di richieste non valide a support.
			CALL SendEmailSupport("Raggiunto limite di nuove richieste non valide per controlli contro SQL INJECTION." & vbCrLF & _
								  "Conteggio attuale: " & Application("INVALID_REQUESTS_COUNT") & vbCrLF & _
								  "Elenco richieste: " & vbCrLF & _
								  Application("INVALID_REQUESTS_LIST") )
			Application("INVALID_REQUESTS_LIST") = ""
		end if
	end if
	
	Security_RequestIsValid = IsValid
end function
 %>
 