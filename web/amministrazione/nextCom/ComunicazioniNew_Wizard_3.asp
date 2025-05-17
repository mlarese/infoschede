<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1000000 %>
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->
<%
dim MessageType 
MessageType = cIntero(request("type"))

if request.form("indietro") <> "" then
	'if ComunicazioniNew_Wizard_Session_GetField(MessageType, "newsletter_scelta_contenuti") <> "" then
	'	response.redirect "ComunicazioniNew_Wizard_2_newsletter.asp?type=" & messageType
	'else
		response.redirect "ComunicazioniNew_Wizard_2.asp?type=" & messageType
	'end if
elseif request.form("avanti") <> "" then
	server.execute("ComunicazioniNew_Wizard_3_Send.asp")
end if

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action,value
'Titolo della pagina
	Titolo_sezione = lcase(ComunicazioniNew_Wizard_Titolo("Comunicazioni in uscita - ", 3, ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")))
'Indirizzo pagina per link su sezione
	HREF = "Comunicazioni.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->
<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<div id="content">
<form action="<%= GetPageName() %>?type=<%= MessageType %>" method="post" id="form1" name="form1">
	<% Select case ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")
		case EMAIL_TYPE_BOZZA, EMAIL_TYPE_ERROR, SMS_TYPE_BOZZA, SMS_TYPE_ERROR, FAX_TYPE_BOZZA, FAX_TYPE_ERROR %>
		<input type="hidden" name="ID" value="<%= ComunicazioniNew_Wizard_Session_GetField(messageType, "email_id") %>">
	<% end select %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption <%= Comunicazioni_CssByType(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo"), true) %>>
			<% CALL Comunicazioni_Icona(MessageType) %>&nbsp;
			<%= ComunicazioniNew_Wizard_Titolo("", 3, ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")) %> - anteprima spedizione
		</caption>
		<% CALL Write_Mittente(conn, rs, 	ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_dipgenera"), MEssageTYpe) %>
		<tr><th colspan="2">DESTINATARI <%= Comunicazioni_LabelByType(messageType, "DELL'EMAIL", "DEL FAX", "DEL SMS") %></th></tr>
		<% Select case ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")
			case EMAIL_TYPE_ERROR, SMS_TYPE_ERROR, FAX_TYPE_ERROR
			case else
				CALL Write_ElencoRubriche(conn, rs, ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche"), "RUBRICHE SELEZIONATE", ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche_interni"), ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche_lingua"))
		end select
		
		dim email_for_newsletter
		if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then
			email_for_newsletter = cBoolean(ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti_email_newsletter"), false)
		else
			email_for_newsletter = false
		end if
		
		CALL Write_ElencoContatti(conn, ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti"), _
								GetListPVSql(ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche")), _
								ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche_interni"), _
								ComunicazioniNew_Wizard_Session_GetField(MessageType, "rubriche_lingua"), _
								"ELENCO COMPLETO DESTINATARI", messageType, _
								email_for_newsletter)
		if messageType = MSG_EMAIL then%>
			<tr><th colspan="2">OGGETTO</th></tr>
			<tr>
				<td class="content" colspan="2">
					<%= ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_object")%>
				</td>
			</tr>
		<% end if %>
		<tr><th colspan="2">CORPO DEL MESSAGGIO</th></tr>
		<% select case ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo")
			case EMAIL_TYPE_TEXT, EMAIL_TYPE_INOLTRO, SMS_TYPE_TEXT %>
				<tr>
					<td class="content" colspan="2">
						<%= TextEncode(ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_text")) %>
					</td>
				</tr>
			<% case EMAIL_TYPE_HTML, EMAIL_TYPE_INOLTRO_HTML %>
				<tr>
					<td colspan="2">
						<%
						dim iframe_url
						iframe_url = ComunicazioniNew_Wizard_Session_GetField(messageType, "url_bozza") & "?key=" & uCase(GetRandomString(DOCUMENTS_FILES_CHARSET, 5))
						%>
						<iframe id="IFrameHTML" name="" style="border:0px; width:100%; height:400px;" src="<%= iframe_url %>"  frameborder="0" scrolling="auto">
						</iframe>
					</td>
				</tr>
			<% case EMAIL_TYPE_TEXTLINK %>
				<tr>
					<td class="content" colspan="2">
						<div title="Testo che precede il link">
							<%= TextEncode(ComunicazioniNew_Wizard_Session_GetField(messageType, "email_text1")) %>
						</div>
						<div title="Link alla pagina di approfondimento del NEXT-web">
						<%value = GetPageURL(NULL, ComunicazioniNew_Wizard_Session_GetField(messageType, "email_pagina_link")) %>
							<a href="<%= value %>" title="Link alla pagina del NEXT-web"><%= value %></a>
						</div>
						<div title="Testo successivo al link">
							<%= ComunicazioniNew_Wizard_Session_GetField(messageType, "email_text2") %>
						</div>
					</td>
				</tr>
			<% case EMAIL_TYPE_NEXTMAIL, FAX_TYPE_NEXTMAIL
				CALL Write_MessageViewFrame(messageType, EncodeCssInlinedUrl(EncodeUrlForEmail(GetPageURL(NULL, ComunicazioniNew_Wizard_Session_GetField(messageType, "email_pagina_esistente")))))
				
			case EMAIL_TYPE_NEWNEXTMAIL, FAX_TYPE_NEWNEXTMAIL
				CALL Write_MessageViewFrame(messageType, EncodeCssInlinedUrl(EncodeUrlForEmail(GetPageURL(NULL, ComunicazioniNew_Wizard_Session_GetField(messageType, "email_nuova_pagina")))))
				
			case EMAIL_TYPE_FILE
				CALL Write_MessageViewFrame(messageType, GetUrlImage(ComunicazioniNew_Wizard_Session_GetField(messageType, "email_file"), 0))
			case FAX_TYPE_FILE %>
				<tr>
					<td class="label">File da inviare</td>
					<td class="content">
						<% FileLink(GetUrlImage(ComunicazioniNew_Wizard_Session_GetField(messageType, "email_file"), 0)) %>
					</td>
				</tr>
			<% case EMAIL_TYPE_BOZZA, EMAIL_TYPE_ERROR, SMS_TYPE_BOZZA, SMS_TYPE_ERROR, FAX_TYPE_BOZZA, FAX_TYPE_ERROR
				CALL Write_MessageViewFrame(messageType, "ComunicazioniViewMessage.asp?ID=" & ComunicazioniNew_Wizard_Session_GetField(messageType, "email_id"))
				
			case SMS_TYPE_INOLTRO, FAX_TYPE_INOLTRO
				CALL Write_MessageViewFrame(MessageType, "ComunicazioniViewMessage.asp?ID=" & ComunicazioniNew_Wizard_Session_GetField(messageType, "inoltro_email_id"))
		 
			case else
				'NEWSLETTER
				%>
				<tr>
					<td colspan="2">
						<%
						iframe_url = ComunicazioniNew_Wizard_Session_GetField(messageType, "url_bozza") & "?key=" & uCase(GetRandomString(DOCUMENTS_FILES_CHARSET, 5))
						%>
						<iframe id="IFrameHTML" name="" style="border:0px; width:100%; height:400px;" src="<%= iframe_url %>"  frameborder="0" scrolling="auto">
						</iframe>
					</td>
				</tr>
				<%
		end select 

		if messageType = MSG_EMAIL then %>
		 	<tr><th colspan="2">ALLEGATI</th></tr>
			<tr><td class="content" colspan="2"><% CALL Write_Allegati(ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_docs"), _
													IIF((ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = EMAIL_TYPE_INOLTRO OR _
														ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo") = EMAIL_TYPE_INOLTRO_HTML), _
														ComunicazioniNew_Wizard_Session_GetField(messageType, "inoltro_email_id"), _
														ComunicazioniNew_Wizard_Session_GetField(messageType, "email_id"))) %></td></tr>	
		<% end if %>
		<tr>
			<td class="footer" colspan="2">
				<% Select case ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo")
					case EMAIL_TYPE_ERROR, SMS_TYPE_ERROR, FAX_TYPE_ERROR %>
					<% case else %>
						<input style="width:10%;" type="submit" class="button" name="indietro" value="&laquo; INDIETRO" title="torna alla selezione dei destinatari e del corpo del messaggio.">
				<% end select %>
				<input style="width:10%;" type="submit" class="button" name="avanti" value="INVIA" title="spedisci <%= Comunicazioni_LabelByType(messageType, "l'email", "il fax", "l'sms") %>">
			</td>
		</tr>
		<input type="hidden" name="codice_verifica_invio" value="<%=GetRandomString(DOCUMENTS_FILES_CHARSET, 10)%>">
	</table>
	&nbsp;
</div>
</body>
</html>
<%= IsPreMailerRenderingActive() %>
<%
conn.close
set rs = nothing
set conn = nothing
%>
