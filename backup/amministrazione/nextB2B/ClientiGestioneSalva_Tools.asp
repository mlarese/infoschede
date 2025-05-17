
<%

dim cnt_id, ut_id, mandatory_flds, id_rubrica, permesso_area_riservata
dim write_log, log_desc
OBJ_contatto.LoadFromForm("extC_riv_attivo;isSocieta;chk_abilitato")

'controllo per login e password
if OBJ_contatto.ValidateLoginAndPassword(request("old_login"), request("conferma_password")) then
	'imposta campi obbligatori a seconda del tipo di contatto
	if Obj_contatto("isSocieta") then
		mandatory_flds = "NomeOrganizzazioneElencoIndirizzi;"
	else
		mandatory_flds = "CognomeElencoIndirizzi;NomeElencoIndirizzi;"
	end if
	'controlla altri campi obbligatori
	if OBJ_contatto.ValidateFields(mandatory_flds, true)	then
		'controllo esito positivo: dati validi
		OBJ_contatto.conn.beginTrans
		
		OBJ_contatto("PraticaPrefisso") = request("extT_riv_codice")
		
		if request("IDCNT") <> "" OR request("ID") <> "" then			'sono in modifica o ho scelto dal menu
			OBJ_contatto.UpdateDB()
			cnt_id = OBJ_contatto("IDElencoIndirizzi")
		else
			cnt_id = OBJ_contatto.InsertIntoDB()
		end if
		
		
		if cIntero(id_rubrica) > 0 then
			CALL OBJ_contatto.AddToRubrica(cnt_id, id_rubrica)
		end if
		CALL OBJ_contatto.AddToRubrica(cnt_id, session("RUBRICA_CLIENTI"))
		CALL OBJ_contatto.RemoveFromRubrica(cnt_id, session("RUBRICA_EX_CLIENTI"))
		CALL OBJ_contatto.RemoveFromRubrica(cnt_id, session("RUBRICA_AGENTI"))
		
		'registrazione utente
		if Trim(cString(permesso_area_riservata)) <> "" then
			ut_id = OBJ_Contatto.UserFromContact(cnt_id, permesso_area_riservata)
		end	if
		ut_id = OBJ_Contatto.UserFromContact(cnt_id, UTENTE_PERMESSO_CLIENTE)
		
		
		if request.form("extN_riv_agente_id") <> "" then
			'cancella l'associazione con precedenti agenti
			sql = " DELETE FROM rel_rub_ind WHERE id_indirizzo=" & cnt_id & " AND " + _
				  " id_rubrica IN (SELECT id_rubrica FROM tb_rubriche WHERE SyncroFilterTable LIKE 'gtb_agenti')"
			CALL OBJ_contatto.conn.Execute(sql, , adExecuteNoRecords)
			
			sql = " INSERT INTO rel_rub_ind(id_indirizzo, id_rubrica) " + _
				  " SELECT " & cnt_id & ", id_rubrica FROM tb_rubriche WHERE SyncroFilterTable LIKE 'gtb_agenti' " + _
				  " AND SyncroFilterKey=" & cIntero(request.form("extN_riv_agente_id"))
			CALL OBJ_contatto.conn.Execute(sql, , adExecuteNoRecords)
		end if
		
		'salva campi esterni
		SalvaCampiEsterniChk OBJ_contatto.conn, rs, "SELECT * FROM gtb_rivenditori", "riv_id", id_ext, "riv_id", ut_id, "extC_riv_attivo"
		
		if(Session("maschera_codice_cl")<>"" and request("extT_riv_codice")="") then
			GeneraCodiceRivenditore OBJ_contatto.conn,ut_id
		end if

		'verifica se il listino deve essere creato:
		if cInteger(request("extn_riv_listino_id")) = 0 then
			dim nuovo_listino
			'deve inserire il nuovo listino collegato direttamente al cliente
			sql = "SELECT * FROM gtb_listini"
			rs.open sql, OBJ_contatto.conn, adOpenKeySet, adLockOptimistic, adCmdText
			rs.AddNew
			if request("listino_codice")<>"" then
				rs("listino_codice") = request("listino_codice")
			elseif request("IsSocieta")<>"" then
				rs("listino_codice") = "Listino cliente - " & request("tft_NomeOrganizzazioneElencoIndirizzi")
			else
				rs("listino_codice") = "Listino cliente - " & request("tft_cognomeelencoindirizzi") & " " & request("tft_nomeelencoindirizzi")
			end if
			rs("listino_datacreazione") = NULL
			rs("listino_datascadenza") = NULL
			rs("listino_B2C") = false
			rs("listino_offerte") = false
			rs("listino_base") = false
			rs("listino_base_attuale") = false
			rs("listino_with_child") = false
			
			if cInteger(request("copia_da_ancestor_id"))>0 then
				rs("listino_ancestor_id") = cInteger(request("copia_da_ancestor_id"))
			else
				rs("listino_ancestor_id") = 0
			end if
			
			rs.Update
			nuovo_listino = rs("listino_id")
			rs.close
				
			'verifica se i prezzi sono da copiare da altri listini
			if cInteger(request("copia_da_altro_id"))>0 or cInteger(request("copia_da_ancestor_id"))>0 then
				
				'copia i prezzi da altro listino (listino semplice e non in offerta speciale)
				sql = " INSERT INTO gtb_prezzi (prz_iva_id, prz_prezzo, prz_var_sconto, prz_var_euro, prz_visibile, " + _
					  " prz_promozione, prz_scontoQ_id, prz_listino_id, prz_variante_id ) " + _
					  " SELECT prz_iva_id, prz_prezzo, prz_var_sconto, prz_var_euro, prz_visibile, " + _
					  " prz_promozione, prz_scontoQ_id, " & nuovo_listino & ", prz_variante_id " + _
					  " FROM gv_listini WHERE prz_listino_id=" & _
					  IIF(cInteger(request("copia_da_altro_id"))>0, cInteger(request("copia_da_altro_id")), cInteger(request("copia_da_ancestor_id")))
				CALL OBJ_contatto.conn.execute(sql, , adExecuteNoRecords)
			end if
			
			'imposta il listino al cliente
			sql = "UPDATE gtb_rivenditori SET riv_listino_id=" & nuovo_listino & " WHERE riv_id=" & ut_id
			CALL OBJ_contatto.conn.execute(sql, , adExecuteNoRecords)
		end if

		'corregge assegnazione agente
		if cInteger(request.form("extN_riv_agente_id"))=0 then
			sql = "UPDATE gtb_rivenditori SET riv_agente_id=NULL WHERE riv_agente_id=0"
			CALL OBJ_contatto.conn.execute(sql, , adExecuteNoRecords)
		end if
		
		'corregge assegnazione lista codici
		if cInteger(request.form("extN_riv_lstcod_id"))=0 then
			sql = "UPDATE gtb_rivenditori SET riv_lstcod_id=NULL WHERE riv_lstcod_id=0"
			CALL OBJ_contatto.conn.execute(sql, , adExecuteNoRecords)
		end if 

		'toglie associazione con rubrica degli agenti / clienti non attivi
		sql = " DELETE FROM rel_rub_ind WHERE id_indirizzo=" & cnt_id & " AND " + _
			  " (id_rubrica = " & Session("RUBRICA_EX_AGENTI") & " OR id_rubrica = " & Session("RUBRICA_EX_CLIENTI") & ") "
		CALL OBJ_contatto.conn.execute(sql, , adExecuteNoRecords)
	
		'......................................................................................................
		'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
		CALL ADDON__CLIENTI__form_salva(OBJ_contatto, rs, ut_id)
		'......................................................................................................
		
		Session("B2B_HTTP_RESULT_SPEDIZ_CREDENZ_ACCESSO") = ""
		'if cIntero(Session("B2B_ID_PAG_SPEDIZ_CREDENZ_ACCESSO")) > 0 then
			'se è cambiato lo stato di abilitazione, ed ora è abilitato, spedisco le credenziali di accesso
		'	if not cBoolean(Session("B2B_OLD_VALUE_CHK_ABILITATO"), true) AND request.form("chk_abilitato") = "on" then
		'		dim codiceInserimento, httpResult
		'		codiceInserimento = GetValueList(OBJ_contatto.conn, NULL, "SELECT codiceInserimento FROM tb_indirizzario WHERE IDElencoIndirizzi = " & cnt_id)
		'		httpResult = ExecuteHttpUrl(GetPageSiteUrl(conn, Session("B2B_ID_PAG_SPEDIZ_CREDENZ_ACCESSO"), OBJ_contatto("lingua"))&"&RIV_ID="&ut_id&"&ID_ADMIN="&Session("ID_ADMIN")&"&IDCNT="&cnt_id&"&KEY="&codiceInserimento&"&HTML_FOR_EMAIL=1")
		'		Session("B2B_HTTP_RESULT_SPEDIZ_CREDENZ_ACCESSO") = httpResult
		'	end if
		'end if

		
		if pagina_redirect = "" then 
			if request.form("salva_modifica")<>"" then
				pagina_redirect =  GetPageName() & "?ID=" & cnt_id
			else
				pagina_redirect = "Clienti.asp"
			end if
		else
			pagina_redirect = pagina_redirect & "&ID=" & ut_id
		end if

		'scrive sul log
		if cBoolean(write_log, false) then
			CALL WriteLogAdmin(OBJ_contatto.conn,"tb_Indirizzario",cnt_id,_
								IIF(request("IDCNT")<>"" OR request("ID")<>"","modifica","inserimento"),log_desc)
			CALL WriteLogAdmin(OBJ_contatto.conn,"gtb_rivenditori",ut_id,_
								IIF(request("IDCNT")<>"" OR request("ID")<>"","modifica","inserimento"),log_desc)									
		end if
		
		'verifica errori di salvataggio
		if Session("ERRORE")<>"" then
			OBJ_contatto.conn.rollbacktrans
		else
			'chiude transazione e conferma dati
			OBJ_contatto.conn.commitTrans
			response.redirect pagina_redirect
		end if
	end if
end if

%>