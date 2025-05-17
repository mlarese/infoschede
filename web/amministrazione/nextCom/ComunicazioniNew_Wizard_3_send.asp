<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1000000 %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->
<% 
dim MessageType
MessageType = cIntero(request("type"))

dim conn, conn_a, rs, rse, rst, sql, ID, sql_where
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set conn_a = Server.CreateObject("ADODB.Connection")
if Application("DATA_ARCHIVE_ConnectionString")<>"" then
	conn_a.open Application("DATA_ARCHIVE_ConnectionString")
end if
set rs = Server.CreateObject("ADODB.RecordSet")
set rse = Server.CreateObject("ADODB.RecordSet")
set rst = Server.CreateObject("ADODB.RecordSet")


dim bozzaSalva, bozzaCarica, Inoltro, Messaggio, Url

if cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, EMAIL_TYPE_NEWSLETTER))>0 then	
	dim tipo_newsletter
	tipo_newsletter = ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo")
	tipo_newsletter = Replace(tipo_newsletter, EMAIL_TYPE_NEWSLETTER&"_", "")
end if


bozzaSalva = request("SALVA") <> ""

if not bozzaSalva then
	dim controllo_invio_doppio
	controllo_invio_doppio = requesT("codice_verifica_invio")
	if Session("controllo_invio_doppio") = controllo_invio_doppio then
		Session("ERRORE") = "Invio doppio della newsletter correttamente bloccato."
		CALL SendEmailSupportEX("Newsletter da:" & Request.ServerVariables("SERVER_NAME"), "Errore di invio doppio della newsletter sventato.")
		Response.redirect "Comunicazioni.asp"
	end if
	Session("controllo_invio_doppio") = controllo_invio_doppio
end if 

Select case ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")
	case EMAIL_TYPE_BOZZA, EMAIL_TYPE_BOZZA_HTML, FAX_TYPE_BOZZA, SMS_TYPE_BOZZA
		bozzaCarica = true
		
	case EMAIL_TYPE_ERROR, FAX_TYPE_ERROR, SMS_TYPE_ERROR
		
		'spedizione di un messaggio con errori: completamento della spedizione.
		ID = ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_id")
		
		'impostazione testata messaggio
		sql = "SELECT * FROM tb_email WHERE email_id = "& ID
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		if rs("email_archiviata") then
			rse.open sql, conn_a, adOpenStatic, adLockOptimistic, adCmdText
		else
			set rse = rs
		end if
		
		
		Select case MessageType
			case MSG_EMAIL
				'carica dati dell'email ed inmposta oggetto per spedizione
				set Messaggio = new Mailer
				Messaggio.subject = rs("email_object")
				if CString(rs("email_mime")) = "" then
					Messaggio.body = rse("email_text")
				else
					Messaggio.message.HTMLBody = rse("email_text")
				end if
				'carica allegati
				CALL Messaggio.LoadAttachments(rs("email_docs"), ID)
				
			case MSG_FAX
				set Messaggio = new Faxer
				if cString(rs("email_mime")) = MIME_HTML then
					Messaggio.HtmlBody = rse("email_text")
				else
					Messaggio.Body = rse("email_text")
				end if
				
			case MSG_SMS
				set Messaggio = new SMSSender
				Messaggio.body = rse("email_text")
				
		end select
		CALL Messaggio.LoadSenderById(conn, rst, ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_dipgenera"))
		rs.close
		
		'ripete spedizione per elementi con errori
		dim update
		if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) AND cBoolean(ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti_email_newsletter"), false) then
			sql_where = SQL_isTrue(conn, "email_newsletter")
		else
			sql_where = SQL_IsTrue(conn, "email_default")
		end if
		sql = " SELECT * FROM log_cnt_email l"& _
			  " INNER JOIN tb_valoriNumeri v ON (l.log_cnt_id = v.id_Indirizzario"& _
			  " 								 AND id_TipoNumero=" & Messaggio.RecipientType & " AND " & sql_where &")"& _
			  " WHERE log_email_id = "& ID & _
			  " AND NOT "& SQL_IsTrue(conn, "log_inviato_ok")
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		while not rs.eof
			update = false
			if Messaggio.SendOnError(rs("valoreNumero")) = 0 then
				rs("log_inviato_ok") = true
				update = true
			end if
			if UCase(CString("log_email")) <> UCase(CString(rs("valoreNumero"))) then
				rs("log_email") = rs("valoreNumero")
				update = true
			end if
			if update then
				rs.update
			end if
			
			rs.movenext
		wend
		rs.close
		conn.close
		if Application("DATA_ARCHIVE_ConnectionString")<>"" then
			conn_a.close
		end if
		set rs = nothing
		set rse = nothing
		set rst = nothing
		set conn = nothing
		set conn_a = nothing
		response.redirect "ComunicazioniView.asp?FromSend=" & ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") & "&ID="& ID
	case EMAIL_TYPE_INOLTRO, EMAIL_TYPE_INOLTRO_HTML
		Inoltro = true
		bozzaCarica = false
	case else
		bozzaCarica = false
end select 


'salva messaggio dai dati presenti in sessione
CALL CheckMessage(MessageType)


if session("ERRORE") = "" then
	
	conn.BeginTrans
	
	'salva messaggio nella tabella email.
	sql = "SELECT * FROM tb_email WHERE (email_tipi_messaggi_id = " & MessageType & ") "
	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tfn_email_tipi_messaggi_id", MessageType)	
	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_control_key", GetRandomString(ALPHANUMERIC_CHARSET, 8))
	ID = SalvaCampiEsterniUltra(conn, rs, sql, "email_id", ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_id"), "", 0, "", NULL, session.contents, "COM_NEW_WIZARD_" & MessageType & "_tf")
	
	if session("ERRORE") = "" then
		'salvataggio andato a buon fine: genera messaggi effettivi e prepara spedizione
		
		Select case MessageType
			case MSG_EMAIL
				'carica dati dell'email
				set Messaggio = new Mailer
				'compone corpo del messaggio a seconda del tipo:
				
				if cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_pagina_esistente"))>0 OR _
				   cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_nuova_pagina"))>0 then
				   
				   	'invio di una NEXT_MAIL nuova o selezionata dalle pagine esistenti
					
					if cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_nuova_pagina"))>0 then
            			URL = GetPageUrl(NULL, ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_nuova_pagina"))
					elseif cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_pagina_esistente"))>0 then
						URL = GetPageUrl(NULL, ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_pagina_esistente"))
					end if
					
					'carica HTML su corpo email
					Messaggio.LoadHTML URL, ExtractPageBaseUrl(URL)
				
				elseif CString(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_file")) <> "" then				
					'invio di un file HTML come corpo del messaggio
					
					'carica HTML su corpo email
					Messaggio.message.CreateMHTMLBody "http://"& Application("IMAGE_SERVER") &"/"& Application("AZ_ID") &"/images/"& ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_file"), CdoSuppressAll
				
				elseif CIntero(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_pagina_link")) > 0 then			
					'invio testo + link
					Messaggio.body = IIF(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_text1") <> "", ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_text1") + vbCrLf, "") + _
					 			   	 GetPageUrl(NULL, ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_pagina_link")) + _
									 IIF(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_text2") <> "", vbCrLf + ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_text2"), "")
				
				elseif ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_tipo") = EMAIL_TYPE_HTML OR _
					cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, EMAIL_TYPE_NEWSLETTER)) > 0 OR _
					ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = EMAIL_TYPE_INOLTRO_HTML OR _
					ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = EMAIL_TYPE_BOZZA_HTML then
					'email HTML, email NEWSLETTER oppure INOLTRO di una mail html			
			
					CALL SetLinkViewContentWithBrowser(conn, rs, ID)
					
					'carico l'html da inviare dal file temporaneo salvato in precedenza
					URL = ComunicazioniNew_Wizard_Session_GetField(messageType, "url_bozza")
					Messaggio.LoadHTML URL, ExtractPageBaseUrl(URL)
					
				else
					'invio di testo semplice o bozza o inoltro di una email
					if ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_mime") = "text/html" then
						Messaggio.message.HTMLbody = ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_text")
					else
						Messaggio.body = ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_text")
					end if 
				end if
				
				Messaggio.Subject = ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_object")
				'carica allegati email
				if BozzaCarica then
					CALL Messaggio.LoadAttachments(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_docs"), ID)
				elseif Inoltro then
					CALL Messaggio.LoadAttachments(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_docs"), ComunicazioniNew_Wizard_Session_GetField(messageType, "inoltro_email_id"))
				else
					CALL Messaggio.LoadAttachments(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_docs"), 0)
				end if
			
			case MSG_FAX
				set Messaggio = new Faxer
				if cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_pagina_esistente"))>0 OR _
				   cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_nuova_pagina"))>0 OR _
				   cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, EMAIL_TYPE_NEWSLETTER))>0 then
				   	'invio di un next-fax nuovo o selezionato dalle pagine esistenti
					
					if cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_nuova_pagina"))>0 then
            			URL = GetPageUrl(NULL, ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_nuova_pagina"))
					elseif cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_pagina_esistente"))>0 then
						URL = GetPageUrl(NULL, ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_pagina_esistente"))
					else
						URL = GetPageUrl(NULL, ComunicazioniNew_Wizard_Session_GetField(MessageType, EMAIL_TYPE_NEWSLETTER))&"&TIPO_NEWSLETTER="&tipo_newsletter
					end if
        			
					'carica HTML su corpo email
					Messaggio.LoadHTML URL, ExtractPageBaseUrl(URL)
					
				else
					'invio di un fax semplice
					Messaggio.Body = Application("IMAGE_PATH") &"/"& Application("AZ_ID") &"/images/"& ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_file")
				end if
			case MSG_SMS
				
				set Messaggio = new SMSSender
				Messaggio.body = ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_text")
				
		end select
		
		Messaggio.MessageId = ID

		CALL Messaggio.LoadSenderById(conn, rst, ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_dipgenera"))
	
		'costruzione query per indirizzi
		sql = ""
		'carica elenco contatti
		if ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti")<>"" then
			sql = sql & " id_Indirizzario IN ("& GetListPVSql(ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti")) & ") "
		end if

		'gestisce rubriche
		if ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche")<>"" then
			dim rubrica, rubriche
			rubriche = GetListPVSql(ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche"))
			
			if NOT bozzaSalva then
				'prepara query dei contatti associati alle rubriche indicate
				if ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti")<>"" then
					sql = sql & " OR "
				end if
				sql = sql & " EXISTS (SELECT 1 FROM rel_rub_ind "& _
							" 		  WHERE id_rubrica IN (" & rubriche & ")"& _
							"		  AND (id_indirizzo = idElencoIndirizzi"
				if cIntero(ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche_interni"))>0 then
					sql = sql &" OR id_indirizzo = cntRel"
				end if
				sql = sql &")"
				if ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche_lingua") <> "" then
					sql = sql &" AND lingua = '"& ParseSQL(ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche_lingua"), adChar) &"'"
				end if
				sql = sql &")"
			end if
		
			'salvo il log delle rubriche
			if bozzaCarica then
				'cancella eventuale log precedente
				conn.Execute("DELETE FROM log_rubriche_email WHERE log_email_id = "& ID)
			end if
			'salva log di spedizione alle rubriche
			for each rubrica in Split(rubriche, ",")
				if CIntero(rubrica) > 0 then
					conn.Execute("INSERT INTO log_rubriche_email (log_rubrica_id, log_email_id) VALUES ("& rubrica &", "& ID &")")
				end if
			next
		end if		
		
		'carica elenco contatti e spedisce o salva messaggio
		if sql <> "" then
			if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) AND cBoolean(ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti_email_newsletter"), false) then
				sql_where = SQL_isTrue(conn, "email_newsletter")
			else
				sql_where = SQL_IsTrue(conn, "email_default")
			end if
								
			sql = " SELECT * FROM tb_ValoriNumeri v" &_
				  " INNER JOIN tb_indirizzario i ON v.id_indirizzario = i.idElencoIndirizzi"& _
				  " WHERE id_TipoNumero=" & Messaggio.RecipientType & " AND " & sql_where & _
				  " AND (" & sql & ") ORDER BY ValoreNumero"
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			if bozzaSalva then
				'salva come bozza: salva solo log di spedizione,
				while not rs.eof
					sql = " INSERT INTO log_cnt_email (log_cnt_id, log_email, log_email_id, log_cnt_nominativo, log_inviato_ok)"& _
					  	  " VALUES ("& rs("idElencoIndirizzi") &", '"& rs("valoreNumero") &"', "& ID &", '"& ParseSQL(ContactFullName(rs), adChar) &"', 0)"
					CALL conn.Execute(sql)
					rs.movenext
				wend
			else
				'spedizione messaggi e registrazione log
				if bozzaCarica then
					'ripulisce log da eventuale caricamento bozza
					CALL conn.Execute("DELETE FROM log_cnt_email WHERE log_email_id = "& ID)
				end if
				
				sql = ""
				while not rs.eof
					'spedizione messaggio
					CALL Messaggio.SendSave(conn, rs("idElencoIndirizzi"), rs("ValoreNumero"), ContactFullName(rs))
				
					rs.movenext
				wend
			end if
			rs.close
		end if
	end if
	
	
	'sposta gli allegati nella cartella docs e pulisce cartella temporanea
	if NOT bozzaCarica then
		if messageType = MSG_EMAIL then
			CALL Messaggio.SaveAttachments(ID)
		end if
	end if

	'modifica dati email
	sql = "SELECT * FROM tb_email WHERE email_ID=" & ID
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	if bozzaCarica then
		rs("email_isBozza") = false
	end if
	if bozzaSalva then
		rs("email_isBozza") = true
	end if	
	rs("email_mime") = Messaggio.MimeType
	rs("email_text") = IIF(Messaggio.MimeType = MIME_HTML, Messaggio.HtmlBody, Messaggio.Body)
	rs.update
	rs.close
		
	if cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_nuova_pagina"))>0 then
		'invio di una next-mail con nuova pagina creata al volo
		dim nextWeb_Conn
		'cancella pagina creata temporaneamente
		set nextWeb_Conn = Server.CreateObject("ADODB.Connection")
		nextWeb_Conn.open Application("l_conn_ConnectionString"),"",""
		
		sql = "DELETE FROM tb_pages WHERE id_page=" & cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, "email_nuova_pagina"))
		CALL nextWeb_Conn.execute(sql, 0, adExecuteNoRecords)
		
		nextWeb_Conn.close
		set nextWeb_Conn = nothing
	end if
	
	if cInteger(ComunicazioniNew_Wizard_Session_GetField(MessageType, EMAIL_TYPE_NEWSLETTER))>0 then
		sql = " UPDATE tb_newsletters_contents SET nlc_data_invio="&SQL_date(conn, Now())&", nlc_email_inviata_id="&Messaggio.MessageId& _
			  " WHERE ISNULL(nlc_data_invio,0)=0 AND nlc_tipo_id = " & tipo_newsletter
		conn.execute(sql)
		sql = " UPDATE tb_email SET email_newsletter_tipo_id = "& tipo_newsletter & _
			  " WHERE email_id = " & Messaggio.MessageId 
		conn.execute(sql)
	end if
	
	
	conn.CommitTrans
	conn.close
	set rs = nothing
	set rst = nothing
	set conn = nothing
		
	if bozzaSalva then
		response.redirect "Comunicazioni.asp"
	else
		response.redirect "ComunicazioniView.asp?FromSend=" & ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") & "&ID="& ID
	end if
end if

%>