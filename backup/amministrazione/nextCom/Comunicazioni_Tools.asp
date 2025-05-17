<!--#INCLUDE FILE="../library/Class_Mailer.asp" -->
<!--#INCLUDE FILE="../library/Class_Fax.asp" -->
<!--#INCLUDE FILE="../library/Class_Sms.asp" -->
<% 
'.................................................................................................
'.................................................................................................
'.................................................................................................
'FUNZIONI E COSTANTI PER LA GENERAZIONE DELLE EMAIL
'.................................................................................................
'.................................................................................................

'definizione tipi dei messaggi
const EMAIL_TYPE_TEXT = "email_testo"
const EMAIL_TYPE_TEXTLINK = "email_testolink"
const EMAIL_TYPE_NEXTMAIL = "email_next"
const EMAIL_TYPE_NEWNEXTMAIL = "email_newnext"
const EMAIL_TYPE_FILE = "email_file"
const EMAIL_TYPE_BOZZA = "email_bozza"
const EMAIL_TYPE_BOZZA_HTML = "email_bozza_html"
const EMAIL_TYPE_ERROR = "email_errore"
const EMAIL_TYPE_INOLTRO = "email_inoltra"
const EMAIL_TYPE_NEWSLETTER = "email_newsletter"
const EMAIL_TYPE_HTML = "email_html"
const EMAIL_TYPE_INOLTRO_HTML = "email_inoltra_html"

const SMS_TYPE_TEXT = "sms"
const SMS_TYPE_BOZZA = "sms_bozza"
const SMS_TYPE_ERROR = "sms_errore"
const SMS_TYPE_INOLTRO = "sms_inoltra"

const FAX_TYPE_FILE = "fax_file"
const FAX_TYPE_NEXTMAIL = "fax_next"
const FAX_TYPE_NEWNEXTMAIL = "fax_newnext"
const FAX_TYPE_BOZZA = "fax_bozza"
const FAX_TYPE_ERROR = "fax_errore"
const FAX_TYPE_INOLTRO = "fax_inoltra"

'.................................................................................................
'FUNZIONI di supporto
'.................................................................................................
sub Comunicazioni_Icona(messageType)
	select case cIntero(messageType)
	case MSG_SMS %>
		<img src="../grafica/icona_sms.gif" alt="SMS">
	<% case MSG_FAX %>
		<img src="../grafica/icona_fax.gif" alt="FAX">
	<% case MSG_EMAIL %>
		<img src="../grafica/icona_email.gif" alt="EMAIL">
	<%end select
end sub


function Comunicazioni_LabelByType(messageType, emailLabel, faxLabel, smsLabel)
	select case cIntero(messageType)
		case MSG_SMS 
			Comunicazioni_LabelByType = smsLabel
		case MSG_FAX
			Comunicazioni_LabelByType = faxLabel
		case MSG_EMAIL
			Comunicazioni_LabelByType = emailLabel
	end select
end function


function Comunicazioni_CssByType(tipo, writeAttribute)
	Select case tipo
		case EMAIL_TYPE_BOZZA, EMAIL_TYPE_BOZZA_HTML, SMS_TYPE_BOZZA, FAX_TYPE_BOZZA
			Comunicazioni_CssByType = "sticker"
		case EMAIL_TYPE_ERROR, FAX_TYPE_ERROR, SMS_TYPE_ERROR
			Comunicazioni_CssByType = "alert"
		case else
			Comunicazioni_CssByType = ""
	end select
	if writeAttribute AND Comunicazioni_CssByType <> "" then
		Comunicazioni_CssByType = " class=""" & Comunicazioni_CssByType & """ "
	end if
end function 


function ComunicazioniNew_Wizard_Titolo(BaseLabel, Passo, tipo)
	dim label
	Select case tipo
		case EMAIL_TYPE_BOZZA, EMAIL_TYPE_BOZZA_HTML
			label = "Invio email salvata - passo " & (Passo - 1) & " di 3"
		case EMAIL_TYPE_ERROR
			label = "Nuovo invio email con errori - passo " & (Passo - 2) & " di 2"
		case EMAIL_TYPE_INOLTRO 
			label = "Inoltro dell'email - passo " & (Passo - 1) & " di 3"
		case EMAIL_TYPE_TEXT, EMAIL_TYPE_TEXTLINK, EMAIL_TYPE_NEXTMAIL, EMAIL_TYPE_NEWNEXTMAIL, EMAIL_TYPE_FILE, EMAIL_TYPE_HTML
			label = "Nuova email - passo " & Passo & " di 4"
		
		case SMS_TYPE_TEXT
			label = "Nuovo sms - passo " & (Passo - 1) & " di 3"
		case SMS_TYPE_BOZZA
			label = "Invio sms salvato - passo " & (Passo - 1) & " di 3"
		case SMS_TYPE_ERROR 
			label = "Nuovo invio sms con errori - passo " & (Passo - 2) & " di 2"
		case SMS_TYPE_INOLTRO 
			label = "Inoltro del sms - passo " & (Passo - 1) & " di 3"
		
		case FAX_TYPE_FILE, FAX_TYPE_NEXTMAIL, FAX_TYPE_NEWNEXTMAIL
			label = "Nuovo fax - passo " & Passo & " di 4"
		case FAX_TYPE_BOZZA
			label = "Invio fax salvato - passo " & (Passo - 1) & " di 3"
		case FAX_TYPE_ERROR 
			label = "Nuovo invio fax con errori - passo " & (Passo - 2) & " di 2"
		case FAX_TYPE_INOLTRO 
			label = "Inoltro del fax - passo " & (Passo - 1) & " di 3"
		
		case else
			if instr(tipo, EMAIL_TYPE_NEWSLETTER) > 0 then
				label = "Nuova newsletter - passo " & Passo & " di 4"
			end if
			if ComunicazioniNew_Wizard_Session_GetField(request("type"), "newsletter_scelta_contenuti") <> "" then
				label = "Nuova newsletter - passo " & (Passo + 1) & " di 5"
			end if
	end select

	ComunicazioniNew_Wizard_Titolo = BaseLabel + label
end function


'.................................................................................................
'FUNZIONI di gestione
'.................................................................................................

'resetta tutti i campi dell'email
sub ComunicazioniNew_Wizard_Session_Reset(messageType)
	dim VarName, VarNameBegin, LenVarNameBegin
	
	CALL DeleteBozzaHtml()
	
	VarNameBegin = lcase("COM_NEW_WIZARD_" & messageType)
	LenVarNameBegin = len(VarNameBegin)
	for each VarName in Session.Contents
		if lcase(left(VarName, LenVarNameBegin)) = VarNameBegin then
			Session(VarName) = ""
		end if
	next

end sub

'aggiunge un campo da "mantenere" nella creazione dell'email
sub ComunicazioniNew_Wizard_Session_AddField(messageType, field, value)
	Session("COM_NEW_WIZARD_" & messageType & "_" + field) = value
end sub

'restituisce il valore recuperato dalla sessione
function ComunicazioniNew_Wizard_Session_GetField(messageType, field)
	ComunicazioniNew_Wizard_Session_GetField = Session("COM_NEW_WIZARD_" & messageType & "_" + field)
end function

'controlla i dati dell'email in sessione e in caso di fallimento imposta il session("ERRORE")
Sub CheckMessage(messageType)
	if ComunicazioniNew_Wizard_Session_GetField(messageType, "contatti")="" AND ComunicazioniNew_Wizard_Session_GetField(messageType, "rubriche")="" then
		'controlla destinatari per tutti i tipi di messaggi
		Session("ERRORE") = "Selezionare almeno un destinatario."
	elseif messageType = MSG_EMAIL AND ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_object") = "" then
		'controlla oggetto dell'email
		Session("ERRORE") = "Oggetto del messaggio vuoto."
	elseif ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_text") = "" AND _
		   ComunicazioniNew_Wizard_Session_GetField(messageType, "email_text1") = "" AND _
		   ComunicazioniNew_Wizard_Session_GetField(messageType, "email_text2") = "" AND _
		   ComunicazioniNew_Wizard_Session_GetField(messageType, "email_pagina_esistente") = "" AND _
		   ComunicazioniNew_Wizard_Session_GetField(messageType, "email_nuova_pagina") = "" AND _
		   ComunicazioniNew_Wizard_Session_GetField(messageType, "email_newsletter") = "" AND _
		   ComunicazioniNew_Wizard_Session_GetField(messageType, "email_file") = "" then
		   	'controlla corpo del messaggio per tutti i tipi
			Session("ERRORE") = "Corpo del messaggio vuoto."
	end if
End Sub


'.................................................................................................
'FUNZIONI PER LA GENERAZIONE DELL'INTERFACCIA
'.................................................................................................
function Write_Mittente(conn, rs, DipId, messageType)
	dim sql, valid%>
	<tr><th colspan="2">MITTENTE</th></tr>
	<tr>
		<td class="label">mittente:</td>
		<td class="content">
			<% sql = "SELECT * FROM tb_admin WHERE id_admin=" & IIF(cIntero(DipId)>0, DipId, Session("ID_ADMIN"))
			rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
			<%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %> 
			<span class="note"> ( 
				<% Select case messageType
					case MSG_EMAIL
						if IsEmail(rs("admin_email_newsletter")) then
							Write_Mittente = true%>
							<%= rs("admin_email_newsletter") %>
						<% elseif IsEmail(rs("admin_email")) then 
							Write_Mittente = true%>
							<%= rs("admin_email") %>
						<% else 
							Write_Mittente = false%>	
							<span class="alert">Indirizzo email non valido.</span>
						<% end if 
					case MSG_SMS
						if IsPhoneNumber(rs("admin_cell")) then 
							Write_Mittente = true%>
							<%= rs("admin_cell") %>
						<% else 
							Write_Mittente = false%>	
							<span class="alert">Numero di cellulare non valido.</span>
						<% end if 
					case MSG_FAX
						if IsPhoneNumber(rs("admin_fax")) then 
							Write_Mittente = true%>
							<%= rs("admin_fax") %>
						<% else 
							Write_Mittente = false%>	
							<span class="alert">Numero di fax non valido.</span>
						<% end if 
				end select %>
			 ) </span>
			<% rs.close %>
		</td>
	</tr>
<% End function 


'restituisce la lista di ID corretta da inserire nell'SQL a partire dalla lista delle rubriche o dei contatti
Function GetListPVSql(rubricheIdList)
	if CString(rubricheIdList) = "" then
		GetListPVSql = "0"
	else
		GetListPVSql = Trim(Replace(rubricheIdList, ";", ","))
		if Right(GetListPVSql, 1) = "," then
			GetListPVSql = GetListPVSql + "0"
		end if
		if left(GetListPVSql, 1) = "," then
			GetListPVSql = "0" + GetListPVSql
		end if
	end if
End Function


sub Write_SelezioneDestinatari(conn, rs, rubricheIdList, rubricheInterni, rubricheLingua, contattiIdList, messageType)
	dim sql, rubricheNameList
	
	if cString(rubricheIdList) <> "" then
		
		sql = "SELECT * FROM tb_rubriche WHERE id_rubrica IN (" & GetListPVSql(rubricheIdList) & ")"
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		rubricheIdList = ""
		rubricheNameList = ""
		
		while not rs.eof
			rubricheIdList = rubricheIdList & " " & rs("id_rubrica") &";"
			rubricheNameList = rubricheNameList & " " & JSReplacerEncode(rs("nome_Rubrica")) &";"
			rs.movenext
		wend
		
		rs.close
	end if
	
	
	%>
	<tr><th colspan="2">DESTINATARI</th></tr>
	<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then %>
		<tr>
			<td class="label_no_width" style="width:17%;">scegli i recapiti:</td>
			<td class="content">
				<table border="0" cellspacing="0" cellpadding="0" style="width:100%">
					<tr>
						<td class="content" style="width:36%;">
							<input type="radio" class="checkbox" name="contatti_email_newsletter" id="contatti_email_newsletter_true" value="true" <%=chk(cBoolean(ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti_email_newsletter"), true))%>>
							recapiti per NEWSLETTER&nbsp;<%CALL write_icona_newsletter(true)%>
						</td>
						<td class="content">
							<input type="radio" class="checkbox" name="contatti_email_newsletter" value="false" <%=chk(not cBoolean(ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti_email_newsletter"), true))%>>
							recapiti di default
						</td>
					</tr>
				</table>
			</td>
		</tr>
	<% end if %>
	<tr>
		<td class="label">
			intere rubriche:
			<input type="Hidden" name="rubriche" value="<%= rubricheIdList %>">
		</td>
		<td class="content">
			<table border="0" cellspacing="0" cellpadding="0" style="width:100%">
				<tr>
					<td style="width:83%;" colspan="3">
						<textarea READONLY style="width:100%; height:30px;" rows="2" name="visRubriche"><%= rubricheNameList %></textarea>
					</td>
					<td style="width:120px; vertical-align: top; padding-top: 1px;">
						<a class="button_textarea" style="line-height:30px; float:left;"
							href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ComunicazioniNew_Wizard_2_Rubriche.asp?messageType=<%= messageType %>&page_No=1&elenco='+form1.rubriche.value, 'selezione_rubriche', 600, 400, true);">
							SCEGLI
						</a>
						<a class="button_textarea" style="line-height:30px; float:left;"
							href="javascript:void(0)" onclick="form1.visRubriche.value='';form1.rubriche.value=''">
							RESET
						</a>
					</td>
				</tr>
				<% sql = "SELECT COUNT(*) FROM tb_indirizzario WHERE " & SQL_IfIsNull(conn, "CntRel", "0") & "<>0 "
				if cIntero(GetValueList(conn, rs, sql))>0 then  
					dim rubint_value
					if cString(rubricheInterni) <> "" then
						rubint_value = cIntero(rubricheInterni)
					elseif cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) AND _
						   cBoolean(ComunicazioniNew_Wizard_Session_GetField(MessageType, "contatti_email_newsletter"), true) then
						rubint_value = 1
					else
						rubint_value = 0
					end if
					%>
					<tr>
						<td class="label_no_width" style="width:17%;" rowspan="2">modalit&agrave; di invio:</td>
						<td class="content">
							<input type="radio" class="checkbox" name="rubriche_interni" value="0" <%= Chk(rubint_value = 0) %>>
							invia solo ai contatti "principali"
						</td>
						<td class="label_no_width" style="width:28%;" rowspan="2">filtro lingua per i contatti delle rubriche scelte:</td>
						<td class="content" rowspan="2">
							<% CALL DropLingue(conn, rs, "rubriche_lingua", rubricheLingua, true, true, "") %>
						</td>
					</tr>
					<tr>
						<td class="content">
							<input type="radio" class="checkbox" name="rubriche_interni" id="invia_anche_contatti_interni" value="1" <%= Chk(rubint_value = 1) %>>
							invia anche ai contatti interni
						</td>
					</tr>
				<% else %>
					<tr>
						<td class="label_no_width" style="width:30%;" colspan="2">filtro lingua per i contatti delle rubriche scelte:</td>
						<td class="content" colspan="2">
							<% CALL DropLingue(conn, rs, "rubriche_lingua", rubricheLingua, true, true, "") %>
						</td>
					</tr>
				<% end if %>
			</table>
		</td>
	</tr>
	<tr>
		<td class="label">singoli contatti:</td>
		<td class="content">
			<% 
			if messageType = MSG_EMAIL then
				CALL WriteContactPicker_Input_Option(conn, rs, "", "", "form1", "contatti", contattiIdList, "EMAILMANDATORY;CNTREL", true, false, false, "", true) 
			elseif messageType = MSG_SMS then
				CALL WriteContactPicker_Input(conn, rs, "", "", "form1", "contatti", contattiIdList, "CELLMANDATORY;CNTREL", true, false, false, "")
			elseif messageType = MSG_FAX then
				CALL WriteContactPicker_Input(conn, rs, "", "", "form1", "contatti", contattiIdList, "FAXMANDATORY;CNTREL", true, false, false, "")
			end if	
			%>
		</td>
	</tr>
<% end sub


sub Write_SelezioneAllegati(conn, rs, docsList)%>
	<tr><th colspan="2">ALLEGATI</th></tr>
	<tr>
		<td class="content" colspan="2">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
				<tr>
					<td width="92%"><input READONLY type="text" name="tft_email_docs" value="<%= docsList%>" style="width:100%;" onclick="OpenAutoPositionedWindow('ComunicazioniNew_Wizard_2_Allegati.asp?docs=' + form1.tft_email_docs.value, 'allegati', 450, 300)"></td>
					<td><a class="button_input" href="javascript:void(0)" onclick="form1.tft_email_docs.onclick()">ALLEGATI</a></td>
				</tr>
			</table>
		</td>
	</tr>
<% end sub


function AllegatiPresenti(allegati)
	allegati = cString(allegati)
	AllegatiPresenti = (allegati<>"")
end function

sub Write_Allegati(allegati, emailId)
	allegati = cString(allegati)
	if cString(allegati) <> "" then
		dim allegato, basePath, baseUrl
		
		if cIntero(emailId)>0 then
			basePath = application("IMAGE_PATH") & "\docs\eml_" & emailId & "\"
			baseUrl = "http://" & Application("IMAGE_SERVER") & "/docs/eml_" & emailId & "/"
		else
			basePath = application("IMAGE_PATH") & "/temp/"
			baseUrl = "http://" & Application("IMAGE_SERVER") & "/temp/"
		end if
				
		for each allegato in Split(allegati, ";")
			allegato = Trim(allegato)
			if allegato <> "" then
				CALL Write_Allegati_FileLink(basePath, baseUrl, allegato)
			end if
		next
	else%>
		<span class="notes">Nessun allegato selezionato.</span>
	<% end if
end sub


sub Write_Allegati_FileLink(basePath, baseUrl, fileName)
	dim FileExtension
	FileExtension = File_Extension( FileName ) %>
	<table cellpadding="0" cellspacing="0">
		<tr>
			<td class="content_image">
				<a title="visualizza il file '<%= FileName %>' in una nuova finestra" onclick="<%= File_OpenInNewWindow(baseUrl + FileName) %>" <%= ACTIVE_STATUS %> href="javascript:void(0)">
	   				<img src="../grafica/filemanager/<%= File_Icon( FileExtension ) %>" alt="visualizza il file '<%= FileName %>' in una nuova finestra" border="0">
				</a>
			</td>
			<td class="content">
				<a title="visualizza il file '<%= FileName %>' in una nuova finestra" onclick="<%= File_OpenInNewWindow(baseUrl + FileName) %>" <%= ACTIVE_STATUS %> href="javascript:void(0)">
	   				<%= File_Name(FileName) %>&nbsp;(<%= File_Dimension( File_Size(basePath + fileName) ) %>)
				</a>
			</td>
		</tr>
	</table>
<% end sub


Sub Write_ElencoContatti(conn, idList, rubricheIdList, rubricheInterni, rubricheLingua, thLabel, MessageType, email_for_newsletter)
	dim rsd, rse, sql, TypeOfVal
	
	TypeOfVal = Comunicazioni_LabelByType(messageType, VAL_EMAIL, VAL_FAX, VAL_CELLULARE)
	
	set rsd = server.createobject("adodb.recordset")
	set rse = server.createobject("adodb.recordset")
	sql = " SELECT * FROM tb_indirizzario i"& _
		  " WHERE idElencoIndirizzi IN ("& GetListPVSql(idList) &")"& _
		  " OR EXISTS (SELECT 1 FROM rel_rub_ind"& _
		  " 		   WHERE id_rubrica IN ("& GetListPVSql(rubricheIdList) &")"& _
		  "			   AND (id_indirizzo = idElencoIndirizzi"
	if cIntero(rubricheInterni)>0 then
		sql = sql &" OR id_indirizzo = cntRel"
	end if
	sql = sql &")"
	if rubricheLingua <> "" then
		sql = sql &" AND lingua = '"& ParseSql(rubricheLingua, adChar) &"'"
	end if
	sql = sql &") ORDER BY modoRegistra"
	'response.write sql
	rsd.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	if thLabel<>"" then %>
		<tr>
			<th class="L2" colspan="2">
				<%= thLabel %>
				<span class="smaller"> ( n&ordm; <%= rsd.recordcount %> contatti ) </span>
				<% if email_for_newsletter then %>
					<span class="smaller" style="float:right;padding-right:5px;">e-mail per newsletter&nbsp;<img src="../grafica/i.p.new.gif"/></span>
				<% end if %>
			</th>
		</tr>
	<% end if
	if not rsd.eof then %>
		<tr>
			<td colspan="2">
				<% 	if rsd.recordcount > 8 then %>
					<span class="overflow" style="height:120px;">
				<% end if %>
				<table cellpadding="0" cellspacing="1" style="width: 100%;">
					<% dim emails, cIndex
					cIndex = 0
					
					while not rsd.eof 
						
						cIndex = cIndex + 1
						if (cIndex Mod 500)=0 then
							Response.Flush()
						end if %>
						<tr>
							<td class="content">
								<% 
								dim sql_where
								if (TypeOfVal=VAL_EMAIL) then
									if cBoolean(email_for_newsletter, false) then
										sql_where = SQL_isTrue(conn, "email_newsletter")
									else
										sql_where = SQL_isTrue(conn, "email_Default")
									end if
								end if
								
								'recupera email o altro riferimento (tel,fax,email) del contatto
			                    sql = " SELECT ValoreNumero FROM tb_ValoriNumeri " &_
								      " WHERE id_TipoNumero=" & TypeOfVal & " AND " + sql_where + _
									  "	AND NOT " + SQL_IsNull(conn, "ValoreNumero") + " AND ValoreNumero<>'' " + _
									  " AND id_Indirizzario=" & rsd("IdElencoIndirizzi")
			                    rsE.open sql , conn, adOpenStatic, adLockOptimistic, adCmdText
	                    
								emails = ""
								if not rsE.eof then
				                    while not rsE.eof
									    if (TypeOfVal=VAL_EMAIL) and IsEmail(cString(rsE("ValoreNumero")))  then
											Emails = Emails & rsE("ValoreNumero")
										elseif (TypeOfVal=VAL_CELLULARE) AND IsPhoneNumber(rsE("ValoreNumero")) then
											Emails = Emails & rsE("ValoreNumero")
										elseif (TypeOfVal=VAL_FAX) AND IsPhoneNumber(rsE("ValoreNumero")) then
											Emails = Emails & rsE("ValoreNumero")
										else
										    Emails = Emails & "<span class=""alert"">" & rsE("ValoreNumero") & " NON VALIDA!</span>"
										end if
									    rsE.movenext
									    if not rsE.eof then Emails = Emails & ", "
				                    wend
								else
									emails = "<span class=""warning"">NON PRESENTE</span>"
								end if
								rsE.close %>
								<span class="note" style="float:right;"><%= Emails %></span>
								<%= ContactLinkedName(rsd) %>
								<% if cIntero(rsd("cntRel")) > 0 then 
									sql = " SELECT IDElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, CognomeElencoIndirizzi, NomeElencoIndirizzi, isSocieta " & _
										  " FROM tb_indirizzario WHERE IDElencoIndirizzi = " & cIntero(rsd("cntRel"))
									rsE.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
									<span class="note">(contatto interno di: <%=ContactFullName(rsE)%>)</span>
									<% rsE.close %>
								<% end if %>
							</td>
						</tr>
						<% rsd.movenext
					wend %>
				</table>
				<% if rsd.recordcount > 8 then %>
					</span>
				<% end if %>
			</td>
		</tr>
	<% end if
	rsd.close
	set rsd = nothing
	set rse = nothing
End Sub



Sub Write_ElencoRubriche(conn, rsd, rubricheIdList, thLabel, CntInterni, lingua)
	
	if  DB_Type(conn) <> DB_ACCESS then
		sql = " SELECT id_rubrica, nome_rubrica, " &_
			  " (SELECT COUNT(*) FROM (rel_rub_ind r"& _
			  "  INNER JOIN tb_indirizzario i ON r.id_indirizzo = i.idElencoIndirizzi ) "& _
			  "  WHERE r.id_rubrica= tb_rubriche.id_rubrica"& _
			  IIF(lingua <> "", " AND i.lingua = '"& ParseSQL(lingua, adChar) &"'", "") & _
			  "  ) AS COUNT_CNT " & _
			  " FROM tb_rubriche" &_
			  " WHERE id_rubrica IN ("& GetListPVSql(rubricheIdList) &")"& _
			  " ORDER BY nome_rubrica"
	else
		sql = " SELECT tb_rubriche.id_rubrica, nome_rubrica, COUNT(id_indirizzo) AS COUNT_CNT " + _
			  " FROM ((tb_rubriche INNER JOIN rel_rub_ind ON tb_rubriche.id_rubrica = rel_rub_ind.id_rubrica) " + _
			  " 	   INNER JOIN tb_indirizzario ON rel_rub_ind.id_indirizzo = tb_indirizzario.idelencoindirizzi) " + _
			  " WHERE tb_rubriche.id_rubrica IN ("& GetListPVSql(rubricheIdList) &")"& _
			  IIF(lingua <> "", " AND tb_indirizzario.lingua = '"& ParseSQL(lingua, adChar) &"'", "") & _
			  " GROUP BY tb_rubriche.id_rubrica, nome_rubrica " + _
			  " ORDER BY nome_rubrica "
	end if

	rsd.open sql, conn, adOpenStatic, adLockReadOnly
	if thLabel<>"" then %>
		<tr>
			<th class="L2" colspan="2">
				<%= thLabel %>
				<span class="smaller"> ( n&ordm; <%= rsd.recordcount %> rubriche ) </span>
			</th>
		</tr>
	<% end if %>
	<tr>
		<td colspan="2">
			<% 	if rsd.recordcount > 8 then %>
				<span class="overflow" style="height:120px;">
			<% end if %>
			<table cellpadding="0" cellspacing="1" style="width:100%;">
				<% if rsd.eof then %>
					<tr><td class="label">Nessuna rubrica selezionata.</td></tr>
				<% else 
					if cIntero(CntInterni)>0 then%>
						<tr>
							<td class="content notes">
								Invia anche ai contatti interni associati alle rubriche selezionate.
							</td>
						</tr>
					<% end if 
					if lingua <> "" then%>
						<tr>
							<td class="content notes">
								La lingua dei contatti delle rubriche &egrave; "<%= GetNomeLingua(lingua) %>".
							</td>
						</tr>
					<% end if 
					while not rsd.eof %>
						<tr>
							<td class="content">
								<span class="note" style="float:right;">n&ordm; <%= cIntero(rsd("COUNT_CNT")) %> contatti associati </span>
								<%= rsd("nome_rubrica") %>
							</td>
						</tr>
						<% rsd.movenext
					wend
				end if %>
			</table>
			<% if rsd.recordcount > 8 then %>
				</span>
			<% end if %>
		</td>
	</tr>
	<% rsd.close
	
end sub


Sub Write_MessageViewFrame(messageType, URL) %>
	<tr>
		<td colspan="2">
			<iframe src="<%= URL %>" name="preview" width="100%" height="<%= IIF(messageType = MSG_SMS, "100", "300") %>" id="preview"></iframe>
		</td>
	</tr>
	<% if messageType <> MSG_SMS then %>
		<tr>
			<td colspan="2" class="label_right">
				<a target="_blank" href="<%= URL %>" class="button_L2">VISUALIZZA IL CORPO <%= Comunicazioni_LabelByType(messageType, "dell'email", "del fax", "del sms") %> IN UNA NUOVA FINESTRA</a>
			</td>
		</tr>
	<% end if
end sub


sub Write_LogCompleto_Destinatari(rs, rse, rsr, rsd, preview)
	dim sql, rsl %>
	<table cellpadding="0" cellspacing="1" width="100%">
		<% if rsd.recordcount = 0 AND rsr.recordcount = 0 then %>
			<tr>
				<td class="note">Nessun destinatario selezionato.</td>
			</tr>
		<% else %>
			<% if rse.recordcount > 0 AND not CBoolean(rs("email_isBozza"), false) then%>
				<tr>
					<th class="l2 alert">
						<span style="float:right;">
							<a class="button_L2" title="Ritenta l'inivio dell'email ai destinatari ai quali non &egrave; stata ancora recapitata."
							   href="ComunicazioniNew_Wizard_2.asp?RitentaErrati_id=<%= rs("email_id") %>">
							   RITENTA INVIO <%= IIF(rse.recordcount=1, "AL CONTATTO", "AI " & rse.recordcount & " CONTATTI") %>
							</a>
						</span>
						ERRORI DI INVIO <span class="smaller"> ( n&ordm; <%= rse.recordcount %> errori ) </span>
					</th>
				</tr>
				<tr>
					<td>
						<% CALL Write_Log_Contatti(rse, false, 4) %>
					</td>
				</tr>
			<% end if 
			if rsr.recordcount > 0 then%>
				<tr>
					<th class="l2">rubriche <%= IIF(cBoolean(rs("email_isBozza"), false), "selezionate", "") %> <span class="smaller"> ( n&ordm; <%= rsr.recordcount %> rubriche ) </span></th>
				</tr>
				<tr>
					<td>
						<% CALL Write_Log_Rubriche(rsr, IIF(not preview, 10, 4)) %>
					</td>
				</tr>
			<% end if
			
			if rsd.recordcount > 0 then %>
				<tr>
					<th class="l2">
						contatti <%= IIF(cBoolean(rs("email_isBozza"), false), "selezionati", "") %>
						<% sql = "SELECT DISTINCT log_email FROM log_cnt_email WHERE log_email_id =" & rs("email_id")
						set rsl = Server.CreateObject("ADODB.recordset")
						rsl.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
						<span class="smaller" title="Inviato a n&ordm; <%= rsl.recordcount %> indirizzi diversi">
							(
							n&ordm; <%= rsd.recordcount %> contatti
							
							<% if rsd.recordcount > 4 AND not cBoolean(rs("email_isBozza"), false) AND preview then %>
								 - <a href="ComunicazioniView.asp?ID=<%= rs("email_id") %>">elenco completo all'interno</a>
							<% end if %>
							)
						</span>
						<% rsl.close
						set rsl = nothing %>
					</th>
				</tr>
				</tr>
					<td>
						<% CALL Write_Log_Contatti(rsd, not cBoolean(rs("email_isBozza"), false) AND preview, IIF(not preview, 10, 4)) %>
					</td>
				</tr>
			<% end if
		end if%>
	</table>
<% end sub


function GetQuery_LogContatti(conn, emailId, soloErrori)
	GetQuery_LogContatti = " SELECT * FROM log_cnt_email " + _
						   " LEFT JOIN tb_indirizzario ON log_cnt_email.log_cnt_id = tb_indirizzario.IdElencoIndirizzi " + _
						   " WHERE log_email_id=" & cIntero(emailId) & _
						   IIF(soloErrori, " AND NOT " & SQL_IsTrue(conn, "log_inviato_ok"), "")
	if soloErrori then
		GetQuery_LogContatti = GetQuery_LogContatti + " AND NOT " & SQL_IsNull(conn, "log_cnt_id")
	end if
	GetQuery_LogContatti = GetQuery_LogContatti + " ORDER BY ISNULL(ModoRegistra, log_cnt_nominativo) "
end function


function GetQuery_LogRubriche(conn, emailId)
	GetQuery_LogRubriche = " SELECT * FROM log_rubriche_email " + _
						   " LEFT JOIN tb_rubriche ON log_rubriche_email.log_rubrica_id = tb_rubriche.id_rubrica " + _
						   " WHERE log_email_id=" & cIntero(emailId) & " ORDER BY log_rubrica_nome "
end function


Sub Write_Log_Contatti(rs, preview, maxRow)
	dim rowCount
	if preview then
		rowCount = maxRow
	else
		rowCount = rs.recordcount + 1
	end if
	
	if not preview OR rs.recordcount < rowCount then
		if rs.recordcount > (maxRow + 1) then %>
		<span class="overflow" <%if maxRow > 5 then %>style="height:<%= 15 * maxRow %>px;"<% end if %>>
		<% end if
	end if%>
	<table cellpadding="0" cellspacing="1" style="width:100%;">
		<tbody style="display: block; max-height: 150px; overflow: auto; width: 100%;">
		<% if rs.eof then %>
			<tr style="display: table; width:100%;">
				<td class="label">Nessun contatto trovato.</td>
			</tr>
		<% else 
			if not preview OR rs.recordcount < rowCount then
				dim cIndex
				cIndex = 0
				while not rs.eof 
					cIndex = cIndex + 1
					if (cIndex Mod 500)=0 then
						Response.Flush()
					end if%>
					<tr style="display: table; width:100%;">
						<td class="content">
							<% if cIntero(rs("IdElencoIndirizzi"))>0 then %>
								<%= ContactLinkedName(rs) %>
							<% else %>
								<%= rs("log_cnt_nominativo") %>
							<% end if %>
							<span class="notes"> ( <%= rs("log_email") %> )</span>
						</td>
							<td style="width:24px;" class="content_right<%= IIF(rs("log_inviato_ok"), " ok", " alert") %>" title="<%= IIF(rs("log_inviato_ok"), "Email inviata correttamente.", "Email non inviata.")  %>">
								<input type="checkbox" class="checkbox" disabled <%= Chk(rs("log_inviato_ok")) %>>
							</td>
					</tr>
					<% rs.movenext
				wend
			end if
			if preview AND not rs.eof then
				dim sql, rsl
				sql = "SELECT DISTINCT log_email FROM log_cnt_email WHERE log_email_id =" & rs("log_email_id")
				set rsl = Server.CreateObject("ADODB.recordset")
				rsl.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
				<tr style="display: table; width:100%;">
					<td colspan="2" class="content" title="Inviato a n&ordm; <%= rsl.recordcount %> indirizzi diversi">
						<span style="float:right;">
							<a class="button_L2" href="ComunicazioniView.asp?ID=<%= rs("log_email_id") %>">
								visualizza elenco completo
							</a>
							<br><br>
						</span>
						email inviata a n&ordm; <%= rs.recordcount %> contatti
					</td>
				</tr>
				<% rsl.close
				set rsl = nothing
			end if
		end if %>
		</tbody>
	</table>
			
	<% if rs.recordcount > (maxRow + 1) AND not preview then %>
		</span>
	<% end if
	
end sub


Sub Write_Log_Rubriche(rs, maxRow)
	
	if rs.recordcount > (maxRow + 1) then %>
		<span class="overflow" <%if maxRow > 5 then %>style="height:<%= 15 * maxRow %>px;"<% end if %>>
	<% end if %>
	
	<table cellpadding="0" cellspacing="1" style="width:100%;">
		<tbody style="display: block; max-height: 150px; overflow: auto; width: 100%;">
		<% if rs.eof then %>
			<tr style="display: table; width:100%;">
				<td class="label">Nessuna rubrica selezionata.</td>
			</tr>
		<% else 
			while not rs.eof %>
				<tr style="display: table; width:100%;">
					<td class="content" style="">
						<% if cIntero(rs("id_rubrica"))>0 then %>
							<%= rs("nome_rubrica") %>
						<% else %>
							<%= rs("log_rubrica_nome") %>
						<% end if %>
					</td>
				</tr>
				<% rs.movenext
			wend
		end if %>
		</tbody>
	</table>
	
	<% if rs.recordcount > (maxRow + 1) then %>
		</span>
	<% end if
	
end sub



Sub ViewMessage()

	dim conn, rs, sql, dbName, ARCHIVE_new_connection_string

	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString"),"",""
	set rs = Server.CreateObject("ADODB.Recordset")

	sql = "SELECT email_mime, email_text, email_object, email_archiviata, email_name_database FROM tb_email WHERE email_id=" & cIntero(request("ID"))
	if Session("LOGIN_4_LOG")="" then
		'utente non loggato: mette il controllo sul querystring con la chiave
		sql = sql & " AND "&SQL_IfIsNull(conn, "email_control_key", "''")&" LIKE '" & ParseSql(cString(request("KEY")), adChar) & "'"
	end if
	'vecchio codice di filtro delle newsletter: usato solo per sicurezza (ora sostituito dal codice di verifica)
	'cambiato il 14/03/2013 da Nicola
	'if isNewsletter then
	'	sql = sql & "AND ISNULL(email_newsletter_tipo_id, 0) > 0 "
	'end if
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if not rs.eof then
		if rs("email_archiviata") AND Application("DATA_ARCHIVE_ConnectionString")<>"" then
			dbName = rs("email_name_database")
			rs.close
			conn.close
			conn.open Application("DATA_ARCHIVE_ConnectionString"),"",""
			ARCHIVE_new_connection_string = Replace(Application("DATA_ARCHIVE_ConnectionString"), cString(conn.DefaultDatabase), dbName)
			conn.close
			conn.open ARCHIVE_new_connection_string,"",""
			rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		end if
		if cString(rs("email_mime")) <> MIME_HTML then
			'scrive email in solo testo
			%>
			<html>
				<head>
					<title><%= rs("email_object") %></title>
					<meta name="robots" content="noindex,nofollow" />
					<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
				</head>
				<body style="font-family:Arial; font-size:11px;">
					<%=TextEncode(rs("email_text"))%>
				</body>
			</html>
		<%else
			'scrive email in formato html
			%>
			<%= rs("email_text") %>
		<%end if
	else %>
		<html>
			<head>
				<title>Email non trovata</title>
				<meta name="robots" content="noindex,nofollow" />
				<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
			</head>
			<body style="font-family:Arial; font-size:11px;">
				Email non trovata.
			</body>
		</html>
	<% end if


	conn.close
	set rs = nothing
	set conn = nothing
	
end sub


'funzione che aggiunge una stringa "stringForWrap" dopo ogni tag html di chiusura e prima di ogni tag d'apertura
Function ReplaceHtmlForWrap(html, stringForWrap)
	Dim regEx, CurrentMatch, CurrentMatches, i, tagList, tagArray, tag, htmlResult, charForSplit
	Set regEx = New RegExp
	regEx.Pattern = "<[^>]*>"
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True
	Set CurrentMatches = regEx.Execute(html)
	
	charForSplit = "§"
	tagList = charForSplit
	If CurrentMatches.Count >= 1 Then
		for i = 0 to CurrentMatches.Count -1
			Set CurrentMatch = CurrentMatches(i)
			if inStr(tagList, charForSplit & CurrentMatch & charForSplit) = 0 then
				tagList = tagList & CurrentMatch & charForSplit
			end if
		next
	End if
	
	htmlResult = html
	tagArray = Split(tagList, charForSplit)
	for each tag in tagArray
		if tag <> "" then
			if inStr(tag, "</") > 0 then
				'tag di chiusura
				htmlResult = Replace(htmlResult, tag, tag & stringForWrap)
			else
				'tag di apertura
				htmlResult = Replace(htmlResult, tag, stringForWrap & tag)
			end if
		end if
	next

	Set regEx = Nothing
	ReplaceHtmlForWrap = htmlResult
End Function



function CleanHtmlForCKEditor(htmlToClean)
	dim html, html_to_replace, i, f, html_pre_body, html_post_body
	html = htmlToClean
	if Len(html) > 0 then
		if inStr(html, "<body") > 0 then
			i = inStr(html, "<body")
			f = inStr(i, html, ">")
			'salvo in sessione il codice html che tolgo, per riutilizzarlo quando ricostruirà la pagina
			CALL ComunicazioniNew_Wizard_Session_AddField(messageType, "HTMLBodyTag", Mid(html, i, (f - i + 1)))
			html = Right(html, Len(html) - f)
			html = Left(html, InstrRev(html, "</body") - 1)
		end if
		
		'ripulisco il codice html dagli input hidden
		CALL ReplaceRexEx(html, "<input type=""hidden[^>]*>", "")
		
		'ripulisce i font size dichiarati in percentuale
		dim cssObj, fontSize, fontReplace
		set cssObj = new cssManager
		for each fontSize in cssObj.FONT_SIZE.keys
			fontReplace = "font-size: " & replace(cString(fontSize), ",", ".") & "%;"
			CALL ReplaceRexEx(html, fontReplace, cssObj.FontSize_px_CSS(fontSize))
		next
		
		'rendo assoluti i link relativi (delle immagini)
		html = Replace(html, "src=""/", "src=""" & GetSiteUrl(null, 0, 0) & "/")
		html = Replace(html, "src='/", "src='" & GetSiteUrl(null, 0, 0) & "/")
		
	end if
	CleanHtmlForCKEditor = html
end function


'salva il corpo dell'email che si sta componendo in un file html
sub WriteBozzaHtml(old_file_path)
	dim fso, bozza, path, filename, url, html_code, Messaggio
	dim stringForSplit, htmlArray, htmlPart
	set fso = CreateObject("Scripting.FileSystemObject")
	
	'path
	path = Application("IMAGE_PATH") & Application("AZ_ID") & "\bozze\"
	if not fso.FolderExists(path) then
		fso.CreateFolder(path)
	end if
	
	'creazione file
	if old_file_path <> "" then
		if inStr(old_file_path, path) > 0 then
			filename = Replace(old_file_path, path, "")
		end if
	else
		filename = Session.SessionId
		filename = filename & "_" & uCase(GetRandomString(DOCUMENTS_FILES_CHARSET, 4)) & ".htm"
	end if
	
	set bozza = fso.CreateTextFile(path & filename, True, True) '(filename[,overwrite[,unicode]])
	
	'html con corpo email
	set Messaggio = new Mailer

	html_code = Messaggio.LoadHTMLCode(conn, ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_object"), _
											ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_text"), _
											ComunicazioniNew_Wizard_Session_GetField(messageType, "HTMLBodyTag"))

	'sostituisco i paragrafi vuoti (usati da CKEditor per andare a capo) con paragrafi contenenti <br>
	CALL ReplaceRexEx(html_code, "<p[^>]*> </p>", "<p> <br /></p>")

	stringForSplit = "§€£$%§"
	html_code = ReplaceHtmlForWrap(html_code, stringForSplit)	
	htmlArray = Split(html_code, stringForSplit)
	for each htmlPart in htmlArray
		if htmlPart <> "" then
			bozza.WriteLine htmlPart
		end if
	next
	bozza.Close
	
	CALL ComunicazioniNew_Wizard_Session_AddField(request("type"), "filepath_bozza", path & filename)
	url = "http://" & Application("IMAGE_SERVER") & "/" & Application("AZ_ID") & "/bozze/" & filename
	CALL ComunicazioniNew_Wizard_Session_AddField(request("type"), "url_bozza", url)
end sub


sub DeleteBozzaHtml()
	dim VarNameBegin, LenVarNameBegin, VarName	
	VarNameBegin = lcase("COM_NEW_WIZARD_" & messageType)
	LenVarNameBegin = len(VarNameBegin)
	for each VarName in Session.Contents
		if lcase(left(VarName, LenVarNameBegin)) = VarNameBegin AND inStr(VarName, "filepath_bozza") > 0 then
			dim fs
			Set fs=Server.CreateObject("Scripting.FileSystemObject")	
			if fs.FileExists(Session(VarName)) then
				fs.DeleteFile(Session(VarName))
			end if
			set fs=nothing
			Session(VarName) = ""
		end if
	next
end sub


sub SetLinkViewContentWithBrowser(conn, rs, ID_email)
	dim fso, path, file, TextStream, html, substitute, url
	set fso = CreateObject("Scripting.FileSystemObject")
	path = ComunicazioniNew_Wizard_Session_GetField(messageType, "filepath_bozza")
	url = ComunicazioniNew_Wizard_Session_GetField(messageType, "url_bozza")

	if (fso.FileExists(path)) = true then
		set file = fso.GetFile(path)
		 ' Open the file
		Set TextStream = file.OpenAsTextStream(1, -1)
		' Read the file line by line
		Do While Not TextStream.AtEndOfStream
			html = html & TextStream.readline & vbCRLF
		Loop
	end if
	
	if html <> "" AND inStr(html, "<!-- begin maillink -->")>0 AND inStr(html, "<!-- end maillink -->")>0 then
		sql = "SELECT email_control_key FROM tb_email WHERE email_id = " & ID_email
		rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
		substitute = "<!-- begin maillink --><a href="""&GetSiteUrl(conn, 0, 0)&"/amministrazione/nextCom/ComunicazioniViewMessage.asp?ID="&ID_email&"&KEY="&rs("email_control_key")&"""><!-- end maillink -->"
		rs.close
		html = ReplaceRexEx(html, "<!-- begin maillink -->[^*]*<!-- end maillink -->", substitute)
		
		Set TextStream = file.OpenAsTextStream(2, -1)
		TextStream.Write html
		
	end if
	set file = nothing
end sub

'sub SetLinkViewContentWithBrowser(conn, rs, ID_email)
'	dim html, sql, substitute
'	sql = "SELECT email_text, email_control_key FROM tb_email WHERE email_id = " & ID_email
'	rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
'	if not rs.eof then
'		html = rs("email_text")
'		substitute = "<!-- begin maillink --><a href="""&GetSiteUrl(conn, 0, 0)&"/amministrazione/nextCom/ComunicazioniViewMessage.asp?ID="&ID_email&"&KEY="&rs("email_control_key")&"""><!-- end maillink -->"
'		html = ReplaceRexEx(html, "<!-- begin maillink -->[^*]*<!-- end maillink -->", substitute)
'		rs("email_text") = html
'		rs.Update
'	end if
'	rs.close
'	response.end
'end sub

%>