<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->
<%
dim conn, rs, sql, MessageType, MittenteValido, ConfigurazioneValida
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

MessageType = cIntero(request("type"))

'carica form nel sistema di recupero.
if request.ServerVariables("REQUEST_METHOD") = "POST" then
	
	'salvo in sessione
	dim field
	for each field in request.form
		field = LCase(field)
		if field <> "indietro" OR field <> "avanti" then
		   	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, field, request.form(field))
		end if
	next

	
	'Salva messaggio in file html temporaneo
	CALL WriteBozzaHtml(ComunicazioniNew_Wizard_Session_GetField(messageType, "filepath_bozza"))


	if request.form("indietro") <> "" then
		if ComunicazioniNew_Wizard_Session_GetField(MessageType, "newsletter_scelta_contenuti") <> "" then
			response.redirect "ComunicazioniNew_Wizard_2_newsletter.asp?type=" & MessageType
		else
			CALL ComunicazioniNew_Wizard_Session_Reset(MessageType)
			response.redirect "ComunicazioniNew_Wizard_1.asp?type=" & MessageType
		end if
	elseif request.form("salva") <> "" OR request.form("invia") <> "" then
		server.execute("ComunicazioniNew_Wizard_3_Send.asp")
	else
		'controllo campi obbligatori
		CALL CheckMessage(MessageType)
		if session("ERRORE") = "" then
			'messaggio riempito correttamente.
			response.redirect "ComunicazioniNew_Wizard_3.asp?type=" & MessageType
		end if
	end if

elseif cIntero(request.querystring("ID")) > 0 OR _
	   cIntero(request.querystring("Inoltra_id"))>0 OR _
	   cIntero(request.querystring("RitentaErrati_id"))>0 then
	'caricamento bozza o inoltro dell'email
	dim ID, str
	if CIntero(request.querystring("ID")) > 0 then
		'caricamento bozza
		ID = request.querystring("ID")
	elseif cIntero(request.querystring("Inoltra_id"))>0 then
		'inoltro
		ID = request.querystring("Inoltra_id")
	else
		'reinvio per errori di spedizione
		ID = request.querystring("RitentaErrati_id")
	end if
	
	'carica dati messaggio
	sql = "SELECT * FROM tb_email WHERE email_id = "& cIntero(ID)
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	
	'----- Giacomo - 02/01/2012 - gestione e-mail archiviate 
	if rs("email_archiviata") AND Application("DATA_ARCHIVE_ConnectionString")<>"" then
		dim ARCHIVE_new_connection_string, dbName
		dbName = rs("email_name_database")
		rs.close
		dim Aconn
		set Aconn = Server.CreateObject("ADODB.Connection")
		Aconn.open Application("DATA_ARCHIVE_ConnectionString"),"",""
		ARCHIVE_new_connection_string = Replace(Application("DATA_ARCHIVE_ConnectionString"), cString(Aconn.DefaultDatabase), dbName)
		Aconn.close
		Aconn.open ARCHIVE_new_connection_string,"",""
		rs.open sql, Aconn, adOpenForwardOnly, adLockReadOnly, adCmdText
	end if
	'-------
	
	MessageType = rs("email_tipi_messaggi_id")
	CALL ComunicazioniNew_Wizard_Session_Reset(MessageType)
	if cIntero(request.querystring("Inoltra_id"))>0 then
		'inoltro del messaggio
		Select case MessageType
			case MSG_EMAIL
				if rs("email_mime") = MIME_HTML then
					CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", EMAIL_TYPE_INOLTRO_HTML)
				else
					CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", EMAIL_TYPE_INOLTRO)
				end if
			case MSG_FAX
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", FAX_TYPE_INOLTRO)
			case MSG_SMS
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", SMS_TYPE_INOLTRO)
		end select
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "inoltro_email_id", rs("email_id"))
	
	elseif cIntero(request.querystring("ID"))>0 then
		'caricamento bozza
		Select case MessageType
			case MSG_EMAIL
				if rs("email_mime") = MIME_HTML then
					CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", EMAIL_TYPE_BOZZA_HTML)
				else
					CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", EMAIL_TYPE_BOZZA)
				end if
			case MSG_FAX
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", FAX_TYPE_BOZZA)
			case MSG_SMS
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", SMS_TYPE_BOZZA)
		end select
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "email_id", rs("email_id"))
	
	else
		'caricamento email con errori
		Select case MessageType
			case MSG_EMAIL
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", EMAIL_TYPE_ERROR)
			case MSG_FAX
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", FAX_TYPE_ERROR)
			case MSG_SMS
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", SMS_TYPE_ERROR)
		end select
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "email_id", rs("email_id"))
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_dipgenera", rs("email_dipgenera"))
	end if
	
	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "inoltro_email_archived", IIF(rs("email_archiviata"), 1, 0))
	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_object", rs("email_object"))
	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_mime", rs("email_mime"))
	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_text", Replace(rs("email_text"), "font-family: &quot;", "")) 
								'Giacomo, 04/04/2013 - il replace server per l'html delle vecchie email, nei casi: style=" font-family:"Trebuchet MS" ", per evitare che si pianti CKEditor
	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_docs", rs("email_docs"))
	
	if rs("email_docs") <> "" AND ID > 0 then
		dim files, fileName, i
		files = Split(rs("email_docs"), ";")
		for i = lbound(files) to ubound(files)
			fileName = Trim(files(i))
			if fileName <> "" then
				'se ci sono allegati li copio nella cartella \temp
				CALL File_Copy(Application("IMAGE_PATH") & "\docs\eml_"&ID&"\" & fileName, Application("IMAGE_PATH") & "\temp\" & fileName)
			end if
		next
	end if
	
	'Giacomo - commentato il 02/01/2012
	' if rs("email_archiviata") then
		''recupera contenuto email dall'archivio
		' rs.close
		' dim Aconn
		' set Aconn = Server.CreateObject("ADODB.Connection")
		' Aconn.open Application("DATA_ARCHIVE_ConnectionString"),"",""
		' rs.open sql, Aconn, adOpenStatic, adLockOptimistic
		' CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_text", rs("email_text"))
		' rs.close
		' Aconn.close
		' set Aconn = nothing
	' else
		' rs.close
	' end if
	
	'----- Giacomo - aggiunto il 02/01/2012
	if rs("email_archiviata") then
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_text", rs("email_text"))
		rs.close
		Aconn.close
		set Aconn = nothing
	else
		rs.close
	end if
	'-----
	
	Select case ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")
		
		case EMAIL_TYPE_INOLTRO, EMAIL_TYPE_INOLTRO_HTML
			CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_object", "I: " + ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_object"))
		
		case EMAIL_TYPE_BOZZA, EMAIL_TYPE_BOZZA_HTML, FAX_TYPE_BOZZA, SMS_TYPE_BOZZA
			'carica destinatari per la bozza
			'rubriche
			sql = "SELECT log_rubrica_id FROM log_rubriche_email WHERE log_email_id = "& ID
			CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "rubriche", Replace(GetValueList(conn, rs, sql), ",", ";"))
			
			'contatti
			sql = "SELECT log_cnt_id FROM log_cnt_email WHERE log_email_id = "& ID
			CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "contatti", Replace(GetValueList(conn, rs, sql), ",", ";"))
		
		case EMAIL_TYPE_ERROR, FAX_TYPE_ERROR, SMS_TYPE_ERROR
			'carica destinatari con errori per cui ripetere l'invio
			sql = GetQuery_LogContatti(conn, ID, true)
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			if rs.eof then
				Session("ERRORE") = "Nessun errore di invio da ritentare per l'email selezionata."
				response.redirect "Comunicazioni.asp"
			else
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "contatti", ValueList(rs, "IdElencoIndirizzi"))
			end if
			rs.close
			
			response.redirect "ComunicazioniNew_Wizard_3.asp?Type=" & MessageType
	end select
end if

'response.write ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_text")
'response.end

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = ComunicazioniNew_Wizard_Titolo("Comunicazioni in uscita - ", 2, ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo"))
'Indirizzo pagina per link su sezione 
	HREF = "Comunicazioni.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->
<!--#INCLUDE FILE="../library/editorHTML/ckeditor/Tools_CKEditor.asp" -->
<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<div id="content">
<form action="<%= GetPageName() %>?type=<%= MessageType %>" method="post" id="form1" name="form1">
	<input type="hidden" name="email_tipo" value="<%= ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") %>">
	<input type="hidden" name="tfd_email_Data" value="NOW">
	<input type="hidden" name="tft_email_dipgenera" value="<%= Session("ID_ADMIN") %>">
	<input type="hidden" name="tfn_email_isBozza" value="<%= IIF(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = EMAIL_TYPE_BOZZA OR ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = EMAIL_TYPE_BOZZA_HTML OR ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = FAX_TYPE_BOZZA OR ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = SMS_TYPE_BOZZA, "1", "0") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption <%= Comunicazioni_CssByType(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo"), true) %>>
			<% CALL Comunicazioni_Icona(MessageType) %>&nbsp;
			<%= ComunicazioniNew_Wizard_Titolo("", 2, ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")) %> -
			<% if ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") <> EMAIL_TYPE_BOZZA OR _
				  ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") <> EMAIL_TYPE_BOZZA_HTML OR _
				  ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") <> SMS_TYPE_BOZZA OR _
				  ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") <> FAX_TYPE_BOZZA then %>
				composizione messaggio e 
			<% end if %>
			selezione destinatari 
		</caption>
		<% 
		MittenteValido = Write_Mittente(conn, rs, 0, MessageType) 

		if inStr(ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo"), EMAIL_TYPE_NEWSLETTER) > 0 then
			dim idNewsletter
			idNewsletter = Replace(ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo"), EMAIL_TYPE_NEWSLETTER, "")
			idNewsletter = Replace(idNewsletter, "_", "")
			idNewsletter = cIntero(idNewsletter)
			if idNewsletter > 0 then
				dim listaRubriche, listaContatti
				
				listaRubriche = GetValueList(conn, NULL, "SELECT nl_rubriche_default FROM tb_newsletters WHERE nl_id = " & idNewsletter)
				if cString(listaRubriche) <> "" AND cString(ComunicazioniNew_Wizard_Session_GetField(messageType, "rubriche")) = "" then
					CALL ComunicazioniNew_Wizard_Session_AddField(messageType, "rubriche", listaRubriche)
				end if
				
				listaContatti = GetValueList(conn, NULL, "SELECT nl_contatti_default FROM tb_newsletters WHERE nl_id = " & idNewsletter)
				if cString(listaContatti) <> "" AND cString(ComunicazioniNew_Wizard_Session_GetField(messageType, "contatti")) = "" then
					CALL ComunicazioniNew_Wizard_Session_AddField(messageType, "contatti", listaContatti)
				end if
			end if
		end if
		
		CALL Write_SelezioneDestinatari(conn, rs, ComunicazioniNew_Wizard_Session_GetField(messageType, "rubriche"), _
										ComunicazioniNew_Wizard_Session_GetField(messageType, "rubriche_interni"), _
										ComunicazioniNew_Wizard_Session_GetField(messageType, "rubriche_lingua"), _
										ComunicazioniNew_Wizard_Session_GetField(messageType, "contatti"), messageType)
		
		if messageType = MSG_EMAIL then %>
			<tr><th colspan="2">OGGETTO</th></tr>
			<tr>
				<td class="content" colspan="2">
					<input type="text" name="tft_email_object" value="<%= ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_object")%>" style="width:100%;">
				</td>
			</tr>
		<% end if

		select case ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo")
			case EMAIL_TYPE_TEXT, EMAIL_TYPE_INOLTRO, EMAIL_TYPE_BOZZA %>
				<tr><th colspan="2">TESTO DEL MESSAGGIO</th></tr>
				<tr>
					<td class="content" colspan="2"><textarea style="width:100%;" rows="15" name="tft_email_text"><%= ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_text") %></textarea></td>
				</tr>
			<% case EMAIL_TYPE_HTML, EMAIL_TYPE_INOLTRO_HTML, EMAIL_TYPE_BOZZA_HTML %>
				<%
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_mime", "text/html") 
				dim htmlCode
				htmlCode = CleanHtmlForCKEditor(ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_text"))
				%>
				<tr><th colspan="2">CORPO DEL MESSAGGIO</th></tr>
				<tr>
					<td class="content" colspan="2"><textarea style="width:100%;" rows="15" name="tft_email_text"><%= htmlCode %></textarea></td>
				</tr>			
				<%
				CALL activateCKEditorComplete("tft_email_text", "400px")
				CALL importStiliStandardCSS("tft_email_text") 
				%>
			<% case EMAIL_TYPE_TEXTLINK %>
				<tr><th colspan="2">CORPO DEL MESSAGGIO</th></tr>
				<tr><th class="L2" colspan="2">TESTO PRECEDENTE AL LINK</th></tr>
				<tr>
					<td class="content" colspan="2"><textarea style="width:100%;" rows="8" name="email_text1"><%= ComunicazioniNew_Wizard_Session_GetField(messageType, "email_text1") %></textarea></td>
				</tr>
				<tr><th class="L2" colspan="2">LINK ALLA PAGINA DI APPROFONDIMENTO</th></tr>
				<tr>
					<td class="label" rowspan="2">pagina</td>
					<td class="content">
						<% CALL DropDownPages(NULL, "form1", 540, 0, "email_pagina_link", ComunicazioniNew_Wizard_Session_GetField(messageType, "email_pagina_link"), false, true)%>
					</td>
				</tr>
				<tr>
					<td class="note">Scegli una pagina del NextWeb precedentemente realizzata come link da includere nel messaggio.</td>
				</tr>
				<tr><th class="L2" colspan="2">TESTO SUCCESSIVO AL LINK</th></tr>
				<tr>
					<td class="content" colspan="2"><textarea style="width:100%;" rows="8" name="email_text2"><%= ComunicazioniNew_Wizard_Session_GetField(messageType, "email_text2") %></textarea></td>
				</tr>
			<% case EMAIL_TYPE_NEXTMAIL, FAX_TYPE_NEXTMAIL %>
				<tr><th colspan="2">CORPO DEL MESSAGGIO</th></tr>
				<tr><th class="L2" colspan="2">PAGINA DEL NEXT-web</th></tr>
				<tr>
					<td class="label" rowspan="2">pagina</td>
					<td class="content">
						<% CALL DropDownPages(NULL, "form1", 540, 0, "email_pagina_esistente", ComunicazioniNew_Wizard_Session_GetField(messageType, "email_pagina_esistente"), false, true)%>
					</td>
				</tr>
				<tr>
					<td class="note">Scegli una pagina del NextWeb precedentemente realizzata come link da includere nel messaggio.</td>
				</tr>
			<% case EMAIL_TYPE_NEWNEXTMAIL, FAX_TYPE_NEWNEXTMAIL %>
				<tr><th colspan="2">CORPO DEL MESSAGGIO</th></tr> 
				<input type="hidden" name="email_nuova_pagina" id="email_nuova_pagina" value="<%= ComunicazioniNew_Wizard_Session_GetField(messageType, "email_nuova_pagina") %>">
				<tr><th class="L2" colspan="2">NUOVA PAGINA CREATA CON IL NEXT-web</th></tr>
				<tr>
					<td class="content_right" colspan="2">
						<script language="JavaScript" type="text/javascript">
							function OpenLoadShock(operazione){
								var width, height;
								if (operazione == 'modifica'){
									width = document.body.clientWidth;
									height = screen.height;
								}
								else{
									width = 500;
									height = 250;
								}

								OpenAutoPositionedScrollWindow('ComunicazioniNew_Wizard_2_loadShock.asp?type=<%= MessageType %>&PAGINA=' + document.form1.email_nuova_pagina.value + '&operazione=' + operazione, 
																operazione, width, height, true);
							}
														
							function SetPreview(pagina){
								if (pagina > 0) {
									//imposta visualizzazione della pagina nel frame e numero di pagina nell'input
									form1.email_nuova_pagina.value = pagina;
									document.email_nuova_pagina_view.document.location ="<%= GetLibraryPath() %>site/PageView.asp?PAGINA=" + pagina;
								} else
									document.email_nuova_pagina_view.document.location.reload(true);
							}
						</script>
						<a class="button_L2" href="javascript:void(0);" onclick="OpenLoadShock('modifica')">
							MODIFICA IL CORPO DEL MESSAGGIO
						</a>
					</td>
				</tr>
				<tr>
					<td class="content" colspan="2">
						<iframe src="" name="email_nuova_pagina_view" width="100%" height="300" id="email_nuova_pagina_view"></iframe>
					</td>
				</tr>
				<% if GetNextWebCurrentVersion(NULL, rs) > 4 then %>
					<tr><th class="L2" colspan="2">OPERAZIONI SULLA PAGINA</th></tr>
					<tr>
						<td class="content_right" colspan="2">
							<span class="note">
								Lingua di spedizione della pagina:
							</span>
							<a class="button_L2" href="javascript:void(0);" onclick="OpenLoadShock('lingua')" title="Apri gestione della lingua di spedizione della pagina." <%= ACTIVE_STATUS %>>
								LINGUA
							</a>
						</td>
					</tr>
					<tr>
						<td class="content_right" colspan="2">
							<span class="note">
								Associa e gestisci il template della pagina:
							</span>
							<a class="button_L2" href="javascript:void(0);" onclick="OpenLoadShock('template')" title="Apri lo strumento per l'associazione del template." <%= ACTIVE_STATUS %>>
								TEMPLATE
							</a>
						</td>
					</tr>
					<tr>
						<td class="content_right" colspan="2">
							<span class="note">
								&Egrave; possibile copiare il corpo dell'email da una pagina o da un template esistente:&nbsp;
							</span>
							<a class="button_L2" href="javascript:void(0);" onclick="OpenLoadShock('copia')" title="Apri lo strumento per la selezione della pagina o del template da cui copiare." <%= ACTIVE_STATUS %>>
								COPIA
							</a>
						</td>
					</tr>
				<% end if
				if cIntero(ComunicazioniNew_Wizard_Session_GetField(messageType, "email_nuova_pagina"))>0 then 
					%>
					<script language="JavaScript" type="text/javascript">
						//imposta visualizzazione della pagina nel frame e numero di pagina nell'input
						SetPreview( <%= ComunicazioniNew_Wizard_Session_GetField(messageType, "email_nuova_pagina") %>);
					</script>
				<% end if
			case EMAIL_TYPE_FILE, FAX_TYPE_FILE %>
				<tr><th colspan="2">CORPO DEL MESSAGGIO</th></tr> 
				<% if MessageType = EMAIL_TYPE_FILE then %>
					<tr><th class="L2" colspan="2">File HTML</th></tr>
				<% end if %>
				<tr>
					<td class="label" rowspan="2">file da inviare:</td>
					<td class="content">
						<% CALL WriteFileSystemPicker_Input(Application("AZ_ID"), FILE_SYSTEM_FILE, "images", IIF(MessageType = MSG_FAX, EXTENSION_FAX, EXTENSION_HTML), "form1", "email_file", ComunicazioniNew_Wizard_Session_GetField(messageType, "email_file"),  "width: 510px;", true, true) %>
					</td>
				</tr>
				<tr>
					<td class="note">Scegli un file HTML con il file manager.</td>
				</tr>
			<% CASE SMS_TYPE_TEXT %>
				<tr><th colspan="2">TESTO DEL MESSAGGIO (max 160 caratteri)</th></tr>
				<tr>
					<td class="content" colspan="2"><textarea style="width:100%;" rows="2" name="tft_email_text"><%= ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_text") %></textarea></td>
				</tr>
			<%	case SMS_TYPE_BOZZA, FAX_TYPE_BOZZA 'EMAIL_TYPE_BOZZA
				CALL Write_MessageViewFrame(MessageType, "ComunicazioniViewMessage.asp?ID=" & ComunicazioniNew_Wizard_Session_GetField(messageType, "email_id"))
			case SMS_TYPE_INOLTRO, FAX_TYPE_INOLTRO
				CALL Write_MessageViewFrame(MessageType, "ComunicazioniViewMessage.asp?ID=" & ComunicazioniNew_Wizard_Session_GetField(messageType, "inoltro_email_id"))
			case else
				'NEWSLETTER
				CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_mime", "text/html") %>
				<tr><th colspan="2">CORPO DEL MESSAGGIO</th></tr>
				<tr>
					<td class="content" colspan="2"><textarea style="width:100%; height:100%;" rows="15" name="tft_email_text"><%= ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_text") %></textarea></td>
				</tr>			
				<%
				CALL activateCKEditorComplete("tft_email_text", "400px")
				CALL importStiliStandardCSS("tft_email_text") 
		end select
		
		if MessageType = MSG_EMAIL then
			'if (ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = EMAIL_TYPE_BOZZA OR _
			'	ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = EMAIL_TYPE_BOZZA_HTML OR _
			'   ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = EMAIL_TYPE_INOLTRO OR _
			'   ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = EMAIL_TYPE_INOLTRO_HTML) AND _
			'   AllegatiPresenti(ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_docs")) then 
			%>
			<!--	<tr><th colspan="2">ALLEGATI</th></tr> 
				<tr><td class="content" colspan="2">
				<%
				'CALL Write_Allegati(ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_docs"), ComunicazioniNew_Wizard_Session_GetField(messageType, "email_id"))
				%>
				</td></tr>	-->
			<%
			'else
				CALL Write_SelezioneAllegati(conn, rs, ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_docs"))
			'end if
		end if
		
		'verifica configurazione
		Select case MessageType
			case MSG_EMAIL
				ConfigurazioneValida = true
			case MSG_FAX
				if Session("FAX_ABILITATI") then
					if Session("FAX_SENDER_EMAIL")<>"" AND _
					   Session("FAX_SENDER_DOMAIN")<>"" then
						ConfigurazioneValida = true
					else
						ConfigurazioneValida = false
					end if
				else
					ConfigurazioneValida = false
				end if
			case MSG_SMS
				if Session("SMS_ABILITATI") then
					if Session("SMS_LOGIN")<>"" AND _
					   Session("SMS_PASSWORD")<>"" then
						ConfigurazioneValida = true
					else
						ConfigurazioneValida = false
					end if
				else
					ConfigurazioneValida = false
				end if
				
		end select
		
	 	
		if MittenteValido AND ConfigurazioneValida then %>
			<tr>
				<td class="footer" colspan="2">
					<span style="float:left;">Opzioni di invio:</span>
					<span style="float:right;">
						<table cellpadding="0" cellspacing="0">
							<script language="JavaScript" type="text/javascript">
								function ConfermaInvio(){
									if (window.confirm("Procedere con l'invio <%= Comunicazioni_LabelByType(MessageType, "dell'email", "del fax", "del sms") %> ai destinatari selezionati?")){
										form1.invia.disabled=false;
										form1.submit();
									}
										
								}
								
								function ConfermaSalva(){
									//if (window.confirm(<% response.write Comunicazioni_LabelByType(MessageType, _
																		"""ATTENZIONE:\nSalvando l'email non sara' piu' possibile modificare il corpo del messaggio e gli allegati.\nSalvare l'email?""", _
																		"""ATTENZIONE:\nSalvando il fax non sara' piu' possibile modificare il corpo del messaggio.\nSalvare il fax?""", _
																		"""ATTENZIONE:\nSalvando non sara' piu' possibile modificare il corpo del sms.\nSalvare sms?""")%>
									//									))
									//{
										form1.salva.disabled=false;
										form1.submit();
									//}
								}
							</script>
							<tr>
								<td class="footer_content" align="right">
									Salva <%= Comunicazioni_LabelByType(MessageType, "l'email ed inviala", "il fax ed invialo", "l'sms ed invialo") %>  in un altro momento:
									<input style="width:190px;" type="button" class="button_L2" onclick="ConfermaSalva()" name="btn_salva" value="SALVA ED INVIA SUCCESSIVAMENTE">
									<input type="hidden" name="salva" value="SALVA ED INVIA SUCCESSIVAMENTE" disabled>
								</td>
							</tr>
						</table>
					</span>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="2">
					(*) Campi obbligatori.
					<% if not ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = EMAIL_TYPE_BOZZA OR _
						  not ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = EMAIL_TYPE_BOZZA_HTML OR _
					      not ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = SMS_TYPE_BOZZA OR _
					      not ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = FAX_TYPE_BOZZA then %>
						<input style="width:10%;" type="submit" class="button" name="indietro" value="&laquo; INDIETRO" title="torna alla selezione del formato.">
					<% end if %>
					<input style="width:10%;" type="submit" class="button" name="avanti" value="AVANTI &raquo;" title="vai all'anteprima di spedizione <%= Comunicazioni_LabelByType(MessageType, "dell'email", "del fax", "del sms") %>.">
				</td>
			</tr>
		<% else %>
			<tr>
				<td class="errore" colspan="2">
					ERRORI DI CONFIGURAZIONE:<br>
					<% if not MittenteValido then %>
						Parametri di spedizione dell'utente non validi (indirizzo mittente non valido).<br>
					<% end if
					if not ConfigurazioneValida then %>
						Parametri di sistema per la spedizione del messaggio non validi. Contattare l'amministratore.<br>
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="2">
					<a href="Comunicazioni.asp" class="button">ANNULLA</a>
				</td>
			</tr>	
		<% end if %>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
conn.close 
set rs = nothing
set conn = nothing%>