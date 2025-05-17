<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->
<%
dim conn, rs, sql, colspan
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
			
dim MessageType
MessageType = cIntero(request.querystring("type"))

CALL ComunicazioniNew_Wizard_Session_Reset(MessageType)

if request.querystring("new")<>"" OR _
   request.querystring("contatti")<>"" OR _ 
   request.querystring("rubriche")<>"" then
	
	'carica elenco di destinatari (contatti e rubriche)
	if request.querystring("contatti")<>"" then
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "contatti", request.querystring("contatti"))
	end if
	if request.querystring("rubriche")<>"" then
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "rubriche", request.querystring("rubriche"))
	end if
	
elseif request.ServerVariables("REQUEST_METHOD") = "POST" then
	if inStr(request("tipo"), EMAIL_TYPE_NEWSLETTER) > 0 then
		if inStr(request("tipo"), "_scelta_contenuti") > 0 then
			CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "newsletter_scelta_contenuti", "true")
		end if
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", Replace(request("tipo"), "_scelta_contenuti", ""))
		response.redirect "ComunicazioniNew_Wizard_2_newsletter.asp?type=" & MessageType
	else
		CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tipo", request("tipo"))
		response.redirect "ComunicazioniNew_Wizard_2.asp?type=" & MessageType
	end if
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Comunicazioni in uscita - " & Comunicazioni_LabelByType(MessageType, "nuova email", "nuovo fax", "nuovo sms") & " - passo 1 di 4"
'Indirizzo pagina per link su sezione 
	HREF = "Comunicazioni.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<%
'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

'verifica se nextweb attivo (SOLO PER EMAIL E FAX)
dim NextWebAbilited
NextWebAbilited = false
if MessageType<>MSG_SMS then
	sql = " SELECT COUNT(*) FROM rel_admin_sito INNER JOIN tb_siti ON rel_admin_sito.sito_id = tb_siti.id_sito"& _
		  " WHERE admin_id = "& Session("ID_ADMIN") & " AND sito_dir LIKE '%NextWeb%'"
	if CIntero(GetValueList(conn, NULL, sql)) > 0 then
		NextWebAbilited = true
	end if
end if
%>
<script type="text/javascript">
	function submit()
	{
		document.form1.submit();
	}
</script>
<div id="content">
	<form action="<%= GetPageName() %>?type=<%= MessageType %>" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption><% CALL Comunicazioni_Icona(MessageType) %>&nbsp;<%= Comunicazioni_LabelByType(MessageType, "Invio di una nuova email", "Invio di un nuovo fax", "Invio di un nuovo sms") %> - passo 1 di 4 - scelta del formato</caption>
		<tr><th colspan="3">SELEZIONA IL TIPO ED IL FORMATO <%= Comunicazioni_LabelByType(MessageType, "dell'email", "del fax", "del sms") %></th></tr>
		<% select case MessageType 
			case MSG_EMAIL%>
				<tr>
					<td class="label" style="width:17%;" rowspan="2">email in formato testo:</td>
					<td class="content" style="width:20%;">
						<input type="radio" name="tipo" value="<%= EMAIL_TYPE_TEXT %>" class="noborder" onclick="submit();" <%= chk(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")="" OR ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = EMAIL_TYPE_TEXT) %>>
						testo semplice
					</td>
					<td class="note">
						Componi ed invia una email testuale.
					</td>
				</tr>
				<tr>
					<td colspan="2" class="note">
						Spedire una email in formato testo garantisce la compatibilit&agrave; con ogni tipo di client per la lettura della posta,
						a differenza di tutti gli altri formati che sono assoggettati alle regole di visualizzazione e di filtro anti-spam proprie di ogni programma 
						(Outlook, Outlook express ...) e servizio di webmail (Google, Yahoo, Hotmail, Libero...).
						<br><br>
					</td>
				</tr>
				<% if NextWebAbilited then %>
					<tr>
						<td class="label" rowspan="2">email in formato HTML:</td>
						<td class="content">
							<input type="radio" name="tipo" value="<%= EMAIL_TYPE_HTML %>" class="noborder" onclick="submit();" <%= chk(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = EMAIL_TYPE_HTML) %>>
							Email HTML
						</td>
						<td class="note">Componi ed invia una email utilizzando un editor HTML.</td>
					</tr>
					<tr>
						<td colspan="2" class="note">
							Utilizzare il formato HTML permette di creare email personalizzate graficamente, ma espone a dei rischi di "eterogeneit&agrave;" della visualizzazione 
							dell'email ai destinatari a causa delle differenti regole adottate da ogni programma o webmail per la lettura della posta.
							<br><br>
						</td>
					</tr>
				<% end if %>
			<% case MSG_FAX %>
				<tr>
					<td class="label" style="width:17%;">fax da file:</td>
					<td class="content" style="width:28%;">
						<input type="radio" name="tipo" value="<%= FAX_TYPE_FILE %>" class="noborder" <%= chk(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")="" OR ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = FAX_TYPE_FILE) %>>
						Invia file
					</td>
					<td class="note">
						Carica un file in formato 
						<img src="../grafica/filemanager/FileIcon_PDF.gif" width="16" height="16" alt=""> Pdf, 
						<img src="../grafica/filemanager/FileIcon_DOC.gif" width="16" height="16" alt=""> Word, 
						<img src="../grafica/filemanager/FileIcon_XLS.gif" width="16" height="16" alt=""> Excel, 
						<img src="../grafica/filemanager/FileIcon_TIF.gif" width="14" height="16" alt=""> Tiff o
						<img src="../grafica/filemanager/FileIcon_JPG.gif" width="14" height="16" alt=""> Jpeg ed invialo come fax.
					</td>
				</tr>
				<tr>
					<td class="label" rowspan="<%= IIF(NextWebAbilited, 2, 1) %>">fax in formato HTML:</td>
					<td class="content">
						<input type="radio" name="tipo" value="<%= FAX_TYPE_NEXTMAIL %>" class="noborder" <%= chk(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = FAX_TYPE_NEXTMAIL) %>>
						NEXT-fax
					</td>
					<td class="note">Seleziona una pagina del NEXT-web tra quelle esistenti e spediscila direttamente come fax.</td>
				</tr>
				<% if NextWebAbilited then %>
					<tr>
						<td class="content">
							<input type="radio" name="tipo" value="<%= FAX_TYPE_NEWNEXTMAIL %>" class="noborder" <%= chk(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = FAX_TYPE_NEWNEXTMAIL) %>>
							nuovo NEXT-fax
						</td>
						<td class="note">Crea una nuova pagina con il NEXT-web e spediscila direttamente come fax.</td>
					</tr>
				<% end if 
			case MSG_SMS 
				'verifica della configurazione
				
				%>
				<tr>
					<td class="label" style="width:17%;">sms in formato testo:</td>
					<td class="content" style="width:28%;">
						<input type="radio" name="tipo" value="<%= SMS_TYPE_TEXT %>" class="noborder" <%= chk(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")="" OR ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = SMS_TYPE_TEXT) %>>
						testo semplice
					</td>
					<td class="note">
						Componi ed invia un sms testuale.
					</td>
				</tr>
		<% end select %>
		
		<%
		set rs = Server.CreateObject("ADODB.RecordSet")
		sql = "SELECT * FROM tb_newsletters WHERE nl_gestione_dinamica_contenuti = 0 ORDER BY nl_gestione_dinamica_contenuti, nl_nome_it"
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		%>
		<tr>
			<td colspan="3">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% if not rs.eof then %>
						<tr>
							<td class="label" style="width:17%;" rowspan="<%=rs.recordCount + 1%>">newsletter:</td>
							<th class="L2" colspan="3">NEWSLETTER CON TEMPLATE PREDEFINITI</th>
						</tr>
						<% CALL WriteElencoNewsletter(conn, rs)
					end if
					rs.close
					
					sql = "SELECT * FROM tb_newsletters WHERE nl_gestione_dinamica_contenuti = 1 ORDER BY nl_gestione_dinamica_contenuti, nl_nome_it"
					rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
					if not rs.eof then
						%>
						<tr>
							<td class="label" style="width:17%;" rowspan="<%=rs.recordCount + 1%>">&nbsp;</td>
							<th class="L2" colspan="3">NEWSLETTER CON SCELTA MANUALE DEI CONTENUTI</th>
						</tr>
						<% 
					end if
					CALL WriteElencoNewsletter(conn, rs)
					rs.close
					%>
				</table>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
				<input style="width:10%;" type="submit" class="button" id="avanti" name="avanti" value="AVANTI &raquo;" title="vai alla scelta destinatari e composizione del corpo del messaggio">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
	<% if MessageType = MSG_SMS then %>
		<script language="JavaScript" type="text/javascript">
			var btnSubmit = document.getElementById("avanti");
			btnSubmit.click();
		</script>
	<% end if %>
</div>
</body>
</html>

<% 

sub WriteElencoNewsletter(conn, rs)
	while not rs.eof %>
		<tr>
			<td class="content" style="width:46%;">
				<% if Trim(cString(rs("nl_lingua")))<>"" then %>
					<span>
						<img src="../grafica/flag_mini_<%= rs("nl_lingua") %>.jpg">
					</span>
				<% end if %>
				<% dim input_value
				input_value = EMAIL_TYPE_NEWSLETTER&"_"&rs("nl_id")&IIF(rs("nl_gestione_dinamica_contenuti"), "_scelta_contenuti", "") 
				%>
				<input type="radio" name="tipo" value="<%= input_value %>" class="noborder" onclick="submit();" <%= chk(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = input_value) %>>
				<span>
					<%=rs("nl_nome_it")%>
				</span>
			</td>
			<td class="note">
				<% if rs("nl_gestione_dinamica_contenuti") then %>
					n. contenuti da spedire: 
					<% sql = "SELECT COUNT(*) FROM tb_newsletters_contents WHERE nlc_tipo_id = "&rs("nl_id")&" AND ISNULL(nlc_data_invio,0)=0 "
					response.write cIntero(GetValueList(conn, NULL, sql)) %>
					<br>
				<% end if %>
				<% dim data_ultimo_invio
				sql = "SELECT TOP 1 nlc_data_invio FROM tb_newsletters_contents WHERE nlc_tipo_id = "&rs("nl_id")&" AND NOT ISNULL(nlc_data_invio,0)=0 ORDER BY nlc_data_invio DESC"
				data_ultimo_invio = GetValueList(conn, NULL, sql) 
				if data_ultimo_invio <> "" then %>
					data ultimo invio: <%= data_ultimo_invio %>
				<% end if %>
			</td>
			<td class="content_right" style="width:14%">
				<a class="button_L2" id="anteprima_<%=rs("nl_id")%>" target="_blank" href="<%=GetPageSiteUrl(conn, rs("nl_pagina_id"), rs("nl_lingua"))%>&TIPO_NEWSLETTER=<%=rs("nl_id")%>&HTML_FOR_EMAIL=1">
					VEDI ANTEPRIMA
				</a>
			</td>
		</tr>
		<%
		rs.moveNext 
	wend
end sub


conn.close()
set conn = nothing
%>