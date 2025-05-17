
<%
dim cnt_id, ut_id, admin_id, gruppo_id, ag_id, rubrica_id, nominativo, permesso_amm_agente, applicativo_amm_agente, permesso_area_riservata
dim write_log, log_desc
OBJ_contatto.LoadFromForm("isSocieta;chk_abilitato")

'controllo per login e password per utente area risevata
if OBJ_contatto.ValidateLoginAndPassword(request("old_login"), request("conferma_password")) then
	'controllo per login e password per utente amministratore: imposta session errore se login gia' presente
	if request("ag_admin_id")<>"" then
		ut_id = request("ag_admin_id")
	elseif request("IDCNT")<>"" then
		sql = "SELECT ut_id FROM tb_utenti WHERE ut_nextCom_id=" & cIntero(request("IDCNT"))
		ut_id = GetValueList(OBJ_contatto.conn, rs, sql)
	else
		ut_id = 0
	end if
	CALL Check_login(OBJ_contatto.conn, rs, true, ut_id, OBJ_contatto("login"))
	
	if session("ERRORE") = "" then
		'imposta campi obbligatori a seconda del tipo di contatto
		'login ok sia per area amministrativa che per area riservata
		if request("IsSocieta")<>"" then
			fields = "NomeOrganizzazioneElencoIndirizzi"
		else
			fields = "CognomeElencoIndirizzi;NomeElencoIndirizzi"
		end if
		if OBJ_contatto.ValidateFields(fields, true)	then
			'controllo esito positivo: dati validi
			OBJ_contatto.conn.beginTrans
		
			if request("IDCNT") <> "" OR request("ID") <> "" then 'sono in modifica o ho scelto dal menu
				OBJ_contatto.UpdateDB()
				cnt_id = OBJ_contatto("IDElencoIndirizzi")
			else
				cnt_id = OBJ_contatto.InsertIntoDB()
			end if
			
			'calcola nominativo agente per registrazione rubrica e gruppo di lavoro
			if OBJ_Contatto("IsSocieta") then
				nominativo =  OBJ_contatto("NomeOrganizzazioneElencoIndirizzi")
			else
				nominativo =  OBJ_contatto("CognomeElencoIndirizzi") + " " + OBJ_contatto("NomeElencoIndirizzi")
			end if
			
			if cIntero(rubrica_id) > 0 then
				CALL OBJ_contatto.AddToRubrica(cnt_id, rubrica_id)
			end if
			CALL OBJ_contatto.AddToRubrica(cnt_id, session("RUBRICA_AGENTI"))
			CALL OBJ_contatto.RemoveFromRubrica(cnt_id, session("RUBRICA_EX_AGENTI"))
			CALL OBJ_contatto.RemoveFromRubrica(cnt_id, session("RUBRICA_CLIENTI"))
			rubrica_id = 0
			
			'registrazione utente area riservata
			if Trim(cString(permesso_area_riservata)) <> "" then
				ut_id = OBJ_Contatto.UserFromContact(cnt_id, permesso_area_riservata)
			end	if
			ut_id = OBJ_Contatto.UserFromContact(cnt_id, UTENTE_PERMESSO_AGENTE)
			
			'salva campi esterni per tabella agenti (commissione)
			ag_id = SalvaCampiEsterni(OBJ_contatto.conn, rs, "SELECT * FROM gtb_agenti", "ag_id", IIF(request("ID")<>"", ut_id, 0), "ag_id", ut_id)
			
			'gestione utente area amministrativa
			sql = "SELECT ag_admin_id FROM gtb_agenti WHERE ag_id=" & ag_id
			sql = "SELECT * FROM tb_admin WHERE id_admin=" & cInteger(GetValueList(OBJ_contatto.conn, rs, sql))
			rs.open sql, OBJ_contatto.conn, adOpenKeySet, adLockOptimistic, adCmdText
			if rs.eof then
				rs.AddNew
			end if
			if OBJ_Contatto("IsSocieta") then
				rs("admin_cognome") =  OBJ_contatto("NomeOrganizzazioneElencoIndirizzi")
			else
				rs("admin_nome") = OBJ_contatto("NomeElencoIndirizzi")
				rs("admin_cognome") = OBJ_contatto("CognomeElencoIndirizzi")
			end if
			rs("admin_email") = OBJ_contatto("email")
			rs("admin_login") = OBJ_contatto("login")
			rs("admin_password") = EncryptPassword(OBJ_contatto("password"))
			if isDate(OBJ_contatto("Scandenza")) then
				rs("admin_scadenza") = NULL
			else
				rs("admin_scadenza") = OBJ_contatto("Scandenza")
			end if
			rs.update
			admin_id = rs("id_admin")
			rs.close
			
			'crea cartella documenti collegata all'utente
			if uCase(cString(request("old_login"))) <> uCase(OBJ_contatto("login")) then
				CALL CreateTemporaryDir(OBJ_contatto("login"), request("old_login"))
			end if
			
			'inserisce permesso dell'agente per accesso area amministrativa
			if Trim(cString(permesso_amm_agente)) = "" then
				permesso_amm_agente = POS_PERMESSO_AGENTE
			end if
			if Trim(cString(applicativo_amm_agente)) = "" then
				applicativo_amm_agente = NEXTB2B
			end if
			sql = " SELECT * FROM rel_admin_sito WHERE admin_id=" & admin_id & " AND sito_id=" & applicativo_amm_agente & _
				  " AND rel_as_permesso=" & permesso_amm_agente
			rs.open sql, OBJ_contatto.conn, adOpenStatic, adLockOptimistic, adCmdText
			if rs.eof then
				rs.AddNew
				rs("admin_id") = admin_id
				rs("sito_id") = applicativo_amm_agente
				rs("rel_as_permesso") = permesso_amm_agente
				rs.update
			end if  
			rs.close
			
			'iniserisce / modifica gruppo di lavoro dell'agente
			sql = "SELECT ag_gruppo_id FROM gtb_agenti WHERE ag_id=" & ag_id
			sql = "SELECT * FROM tb_gruppi WHERE id_gruppo=" & cInteger(GetValueList(OBJ_contatto.conn, rs, sql))
			rs.open sql, OBJ_contatto.conn, adOpenKeySet, adLockOptimistic, adCmdText
			if rs.eof then
				rs.AddNew
			end if
			rs("nome_gruppo") = "Gruppo di " + nominativo
			rs.update
			gruppo_id = rs("id_gruppo")
			rs.close
			
			'inserisce l'agente nel proprio gruppo di lavoro
			sql = "SELECT * FROM tb_rel_dipGruppi WHERE id_impiegato=" & admin_id & " AND id_gruppo=" & gruppo_id
			rs.open sql, OBJ_contatto.conn, adOpenStatic, adLockOptimistic, adCmdText
			if rs.eof then
				rs.AddNew
				rs("id_impiegato") = admin_id
				rs("id_gruppo") = gruppo_id
				rs.update
			end if  
			rs.close
			
			'inserisce / modifica rubrica dei clienti dell'agente e la collega al gruppo di lavoro
			rubrica_id = UpdateSyncroRubricaGruppo(OBJ_contatto.conn, rs, "Clienti di " + nominativo, "", "gtb_agenti", ag_id, gruppo_id)
			
			'imposta chiavi esterne per gruppo agente ed admin agente
			sql = "SELECT * FROM gtb_agenti WHERE ag_id=" & ag_id
			rs.open sql, OBJ_contatto.conn, adOpenStatic, adLockOptimistic, adCmdText
			rs("ag_admin_id") = admin_id
			rs("ag_gruppo_id") = gruppo_id
			rs.update
			rs.close
			
			'scrive sul log
			if cBoolean(write_log, false) then
				CALL WriteLogAdmin(OBJ_contatto.conn,"tb_Indirizzario",cnt_id,_
									IIF(request("IDCNT")<>"" OR request("ID")<>"","modifica","inserimento"),log_desc)
				CALL WriteLogAdmin(OBJ_contatto.conn,"gtb_agenti",ag_id,_
									IIF(request("IDCNT")<>"" OR request("ID")<>"","modifica","inserimento"),log_desc)									
			end if
			
			'chiude transazione e conferma dati
			OBJ_contatto.conn.commitTrans
		
			response.redirect "Agenti.asp"
		end if
	end if
end if
%>
