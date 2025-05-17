
<!--#INCLUDE FILE="../library/Class_Mailer.asp" -->

<% 

const DATA_SENZA_FINE = "31/12/9999"



'.................................................................................................
'..		Restituisce il drop down con i giorni della settimana
'..		inputName		nome del drop down
'..		selectedValue	valore selezionato
'..		obbligatorio	se deve per forza essere selezionato un valore
'.................................................................................................
sub WriteDropDownGiorno(inputName,selectedValue,obbligatorio)
	dim i
	%>
	<select name="<%=inputName%>">
		<% if not obbligatorio then %>
			<option value="" <%=IIF(selectedValue="", "selected", "")%>>scegli...</option>
		<% end if %>
		<% for i=1 to 7 %>
			<option value="<%=i%>" <%=IIF(selectedValue=i, "selected", "")%>><%=NomeGiorno(i, LINGUA_ITALIANO)%></option>
		<% next %>
	</select>
	<%
end sub



'.................................................................................................
'..		Restituisce il drop down
'..		inputName			nome del drop down
'.................................................................................................
sub WriteDropDownMinutiAnticipo(inputName,selectedValue)
	%>
	<select name="<%=inputName%>" id="<%=inputName%>">
		<option value="15" <%= IIF(cIntero(selectedValue) = 15, "selected", "")%>>15 min</option>
		<option value="30" <%= IIF(cIntero(selectedValue) = 30, "selected", "")%>>30 min</option>
		<option value="60" <%= IIF(cIntero(selectedValue) = 60, "selected", "")%>>60 min</option>
		<option value="120" <%= IIF(cIntero(selectedValue) = 120, "selected", "")%>>2 ore</option>
		<option value="1440" <%= IIF(cIntero(selectedValue) = 1440, "selected", "")%>>1 giorno</option>
		<option value="10080" <%= IIF(cIntero(selectedValue) = 10080, "selected", "")%>>1 settimana</option>
	</select>
	<%
end sub 



'.................................................................................................
'..		Restituisce il codice colore (formato #FFFFFF) della tipologia di impegni con id = idTipologia
'.................................................................................................
function GetColorTipologia(idTipologia)
	dim conn, sql
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString"),"",""
	sql = "SELECT tim_colore FROM mtb_tipi_impegni WHERE tim_id = " & idTipologia
	GetColorTipologia = GetValueList(conn, NULL, sql)
end function



'.................................................................................................
'..		Restituisce il nome dell'utente dato l'id 
'.................................................................................................
function GetNomeUtente(id_utente)
	dim conn, sql, rs
	set conn = Server.CreateObject("ADODB.Connection")
	set rs = server.createobject("adodb.recordset")
	conn.open Application("DATA_ConnectionString"),"",""
	sql = "SELECT * FROM tb_Indirizzario WHERE IdElencoIndirizzi IN (SELECT ut_NextCom_ID FROM tb_Utenti WHERE ut_ID = " & id_utente & ")"
	rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
	GetNomeUtente = ContactFullName(rs)
end function



'.................................................................................................
'..		Restituisce il nome del profilo dato l'id 
'.................................................................................................
function GetNomeProfilo(id_profilo, lingua)
	dim conn, sql, risultato
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString"),"",""
	sql = "SELECT pro_nome_"&lingua&" FROM mtb_profili WHERE pro_id = " & id_profilo
	risultato = GetValueList(conn, NULL, sql)
	if cString(risultato) = "" then
		sql = "SELECT pro_nome_it FROM mtb_profili WHERE pro_id = " & id_profilo
		risultato = GetValueList(conn, NULL, sql)
	end if
	GetNomeProfilo = risultato
end function




'.................................................................................................
'..		Spedisce un e-mail a tutti gli utenti collegati all'impegno
'..		idImpegno			id dell'impegno
'..		idPaginaSito		id della pagina da inviare via mail
'.................................................................................................
function SendAvvisoImpegno(conn,idImpegno,idPaginaSito)
	dim rs, sql, listaUserId, UserIdOk, userId, lingua_Dest, cntId_Dest, senderAdminID, titolo_impegno
	set rs = server.createobject("adodb.recordset")

	'conn.beginTrans
		
	if cIntero(idImpegno) > 0 then
		'recupero gli utenti collegati direttamente all'impegno
		sql = "SELECT riu_utente_id FROM mrel_impegni_utenti WHERE riu_impegno_id = " & idImpegno
		listaUserId = GetValueList(conn, NULL, sql)
		
		'recupero gli utenti collegati all'impegno tramite il profilo
		sql = "SELECT rpu_utenti_id FROM mrel_profili_utenti WHERE " & _
			  "		rpu_profilo_id IN (SELECT rip_profilo_id FROM mrel_impegni_profili WHERE rip_impegno_id = " & idImpegno & ")"
		listaUserId = listaUserId & ", " & GetValueList(conn, NULL, sql) & ","

		'recupero l'id del contatto mittente
		sql = "SELECT imp_modAdmin_id FROM mtb_impegni WHERE imp_id = " & idImpegno
		senderAdminID = GetValueList(conn, NULL, sql)
		
		listaUserId = Replace(listaUserId," ","")
		listaUserId = Split(listaUserId,",")
		UserIdOk = ","
		for each userId in listaUserId
			'se non ho appena spedito a questo contatto...
			if inStr(UserIdOk,","&userId&",") = 0 AND cIntero(userId) > 0 then
				'recupero l'id del contatto destinatario
				sql = "SELECT ut_NextCom_ID FROM tb_utenti WHERE ut_ID = " & userId
				cntId_Dest = GetValueList(conn, NULL, sql)
				
				'recupero la lingua del destinatario
				sql = "SELECT lingua FROM tb_Indirizzario WHERE IDElencoIndirizzi = " & cntId_Dest
				lingua_Dest = GetValueList(conn, NULL, sql)
				
				sql = "SELECT imp_titolo_" & lingua_Dest & " FROM mtb_impegni WHERE imp_id = " & idImpegno
				titolo_impegno = GetValueList(conn, NULL, sql)
				if Trim(cString(titolo_impegno)) = "" then
					sql = "SELECT imp_titolo_it FROM mtb_impegni WHERE imp_id = " & idImpegno
					titolo_impegno = GetValueList(conn, NULL, sql)
				end if
				
				'invio l'e-mail
				CALL SendPageFromAdminToContactExtended(conn, rs, lingua_Dest, _
														GetModuleParamExtended(conn, "OGGETTO_PAGINA_AVVISO", lingua_Dest, false) & titolo_impegno, _
														GetPageSiteUrl(conn, idPaginaSito, lingua_Dest) & "&ID="&idImpegno&"&HTML_FOR_EMAIL=1", _
														GetSiteBaseUrl(conn, idPaginaSito), _
														senderAdminID, _
														cntId_Dest, _
														false)
				'scrivo sul log degli avvisi
				CALL WriteLogAvvisoSpedito(conn, idImpegno, userId, senderAdminID)
				
				UserIdOk = UserIdOk & userId & ","
			end if
		next
	end if
	
	'conn.commitTrans
end function



'.................................................................................................
'..		Scrive sul log l'avviso spedito
'..		idImpegno			id dell'impegno
'.................................................................................................
function WriteLogAvvisoSpedito(conn, idImpegno, idUtenteDestinatario, idAdminMittente)
	dim sql
	sql = "INSERT INTO mtb_log_avvisi_spediti(las_impegno_id,las_data_spedizione,las_id_utente_destinatario,las_id_admin_mittente) " & _
		  " SELECT imp_id, "&SQL_Now(conn)&", "&idUtenteDestinatario&", "&idAdminMittente&" FROM mtb_impegni WHERE imp_id = " & idImpegno
	conn.Execute(sql)
end function









%>