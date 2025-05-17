<%

'***************************************************************************************************************************
'***************************************************************************************************************************
'FUNZIONI PER L'IMPORT E L'EXPORT DELLE APPLICAZIONI
'***************************************************************************************************************************
'***************************************************************************************************************************




'copia i dati di tabella in join da un db ad un altro
'connSorg:				connessione a db sorgente
'rsSorg:				recordset creato
'sqlSorg:				sql che rappresenta le tabelle da copiare (prima le tabelle principali)
'connDest:				connessione a db destinazione
'nomeTabelle:			array contenente i nomi delle tabelle da copiare
'nomeChiavi:			array contenente i nomi, in minuscolo, delle chiavi delle rispettive tabelle
'chiaviContatori:		array contenente dei booleani che specificano se le chiavi sono dei contatori [null = tutti contatori]
'nomeChiaviEsterne:		array contenente i nomi, in minuscolo, delle chiavi esterne delle rispettive tabelle [puo' essere null]
'nomeCodici:			array contenente i nomi dei campi per il confronto dati tra i db in modo da non inserire ma modificare [puo' essere null]
'modifica:				true per modificare se record matchano via codice
Sub CopyTables(connSorg, rsSorg, sqlSorg, connDest, nomeTabelle, nomeChiavi, chiaviContatori, nomeChiaviEsterne, nomeCodici, modifica)
	dim campo, campoNome, i
	dim inserimento, sql, codice
	dim rss, keys, keysDest
	redim rss(UBound(nomeTabelle))		'contiene i rs aperti per le varie tabelle del db destinazione
	redim keys(UBound(nomeTabelle))		'contiene le chiavi correnti dei rs del db sorgente
	redim keysDest(UBound(nomeTabelle))	'contiene le chiavi correnti nel db destinazione
	'open
	for i = 0 to UBound(rss)
		set rss(i) = Server.CreateObject("ADODB.RecordSet")
		rss(i).open "SELECT * FROM "& nomeTabelle(i) &" WHERE 1=0", connDest, adOpenKeySet, adLockOptimistic
		
		keys(i) = 0
	next
	
	rsSorg.open sqlSorg, connSorg, adOpenStatic, adLockReadOnly
	while not rsSorg.eof
		for i = 0 to UBound(rss)										'per ogni tabella
			if rsSorg(nomeChiavi(i)) <> keys(i) then					'se inserimento o modifica
				inserimento = true
				keys(i) = rsSorg(nomeChiavi(i))
				if NOT IsNull(nomeCodici) then							'se controllo modifica
					if nomeCodici(i) <> "" then
						rss(i).close
						sql = "SELECT * FROM "& nomeTabelle(i) &" WHERE (1=1)"
						for each codice in Split(nomeCodici(i), ";")
							sql = sql &" AND "& codice &" = "
							if IsNumeric(rsSorg(codice)) then
								sql = sql & rsSorg(codice)
							else
								sql = sql &"'"& rsSorg(codice) &"'"
							end if
						next
						rss(i).open sql, connDest, adOpenKeySet, adLockOptimistic
						inserimento = false
					end if
				end if
				if inserimento OR rss(i).eof then
					if InStr(1, nomeTabelle(i), "JOIN", vbTextCompare) > 0 then		'se c'è una join la tolgo per inserire
						rss(i).close
						sql = "SELECT * FROM "& Left(nomeTabelle(i), InStr(nomeTabelle(i), " ")) &" WHERE 1=0"
						rss(i).open sql, connDest, adOpenKeySet, adLockOptimistic
					end if
					rss(i).AddNew
					inserimento = true
				end if
				
				'copio
				if inserimento OR modifica then
					for each campo in rss(i).Fields
						campoNome = LCase(campo.name)
						if campoNome = nomeChiavi(i) then		'se campo contatore
							if IsNull(chiaviContatori) then
								campoNome = ""
							elseif chiaviContatori(i) then
								campoNome = ""
							end if
						end if
						
						if campoNome <> "" then
							response.write "<!--" & campo.name & "<br>-->"
							if IsNull(nomeChiaviEsterne) then
								rss(i)(campo.name) = rsSorg(campo.name)
							else
								if campoNome = nomeChiaviEsterne(i) then		'se chiave esterna
									rss(i)(campo.name) = keysDest(i - 1)
								else
									if FieldExists(rsSorg, campo.name) then
										rss(i)(campo.name) = rsSorg(campo.name)
									end if
								end if
							end if
						end if
					next
					
					rss(i).Update
				end if
				
				if NOT rss(i).eof then
					keysDest(i) = rss(i)(nomeChiavi(i))
				end if
			end if
		next
		
		rsSorg.movenext
	wend
	rsSorg.close
	
	'close
	for i = 0 to UBound(rss)
		rss(i).close
		set rss(i) = nothing
	next
End Sub


'restituisce la stringa di connessione al db di configurazione
Function GetConfigurationConnectionstring()
	GetConfigurationConnectionstring = "Provider=SQLOLEDB.1;" &_			   
										"User ID=sa;" & _
										"Password=an739NA;" & _
										"Initial Catalog=Configuration;" & _
										"Data Source=192.168.20.20;"
										'"Data Source=nextsviluppo;"		
End Function		


'importa i parametri delle applicazioni installate nel NextPassport
'connSorg:			connessione a db sorgente
'connDest:			connessione a db destinazione
'sitiId:			ID dei siti di cui importare i parametri, "" per tutti i siti
Sub ConfigurationImport(connSorg, connDest, sitiId)
	if sitiId <> "" then
		dim rs, sql, selezione
		set rs = Server.CreateObject("ADODB.RecordSet")
		

		'import descrittori raggruppamenti
		sql = " SELECT * FROM (tb_siti_descrittori_raggruppamenti g"& _
			  " INNER JOIN tb_siti_descrittori d ON g.sdr_id = d.sid_raggruppamento_id)"& _
			  " INNER JOIN rel_siti_descrittori r ON d.sid_id = r.rsd_descrittore_id"& _
			  " WHERE rsd_sito_id IN ("& sitiId &")"
		CALL CopyTables(connSorg, rs, sql, connDest, _
						Array("tb_siti_descrittori_raggruppamenti", "tb_siti_descrittori", _
							  "rel_siti_descrittori r INNER JOIN tb_siti_descrittori d ON r.rsd_descrittore_id = d.sid_id"), _
						Array("sdr_id", "sid_id", "rsd_id"), null, _
						Array("", "sid_raggruppamento_id", "rsd_descrittore_id"), _
						Array("sdr_titolo_it", "sid_codice", "rsd_sito_id;sid_codice"), false)
						
						
		'import descrittori non appartenenti a raggruppamenti
		sql = " SELECT * FROM tb_siti_descrittori INNER JOIN rel_siti_descrittori " + _
			  " ON tb_siti_descrittori.sid_id = rel_siti_descrittori.rsd_descrittore_id " + _
			  " WHERE (tb_siti_descrittori.sid_raggruppamento_id=0 OR tb_siti_descrittori.sid_raggruppamento_id IS NULL) AND (rsd_sito_id IN ("& sitiId &"))"
		CALL CopyTables(connSorg, rs, sql, connDest, _
						Array("tb_siti_descrittori", _
							  "rel_siti_descrittori INNER JOIN tb_siti_descrittori ON rel_siti_descrittori.rsd_descrittore_id = tb_siti_descrittori.sid_id"), _
						Array("sid_id", "rsd_id"), null, _
						Array("", "rsd_descrittore_id"), _
						Array("sid_codice", "rsd_sito_id;sid_codice"), false)
						
		
		'cancello gruppi di descrittori vecchi non personalizzati
		sql = " DELETE FROM tb_siti_descrittori_raggruppamenti"& _
			  " WHERE NOT "& SQL_IsTrue(connDest, "sdr_personalizzato") & _
			  " AND sdr_titolo_it NOT IN ('"& Replace(GetValueList(connSorg, rs, "SELECT sdr_titolo_it FROM tb_siti_descrittori_raggruppamenti"), ",", "','") &"')"
		'connDest.Execute(sql)
		
		'cancello descrittori e relazioni vecchi non personalizzati
		sql = " DELETE FROM tb_siti_descrittori"& _
			  " WHERE NOT "& SQL_IsTrue(connDest, "sid_personalizzato") & _
			  " AND sid_codice NOT IN ('"& Replace(Replace(GetValueList(connSorg, rs, "SELECT sid_codice FROM tb_siti_descrittori"), ",", "','"), " ", "") &"')"
		connDest.Execute(sql)
		
	end if
End Sub





'importa i valori dei parametri vecchi in quelli nuovi
Sub ParametersImport(connContent, sitiId)
	if sitiId <> "" then
		dim rs, sql, valore, canc
		canc = "0"
		sql = " SELECT * FROM (tb_siti_parametri p"& _
			  " INNER JOIN tb_siti_descrittori d ON p.par_key = d.sid_codice)"& _
			  " INNER JOIN rel_siti_descrittori r ON d.sid_id = r.rsd_descrittore_id"& _
			  " WHERE par_sito_id IN ("& sitiId &")"
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open sql, connContent, adOpenStatic, adLockOptimistic
		
		while not rs.eof
			valore = DesValore(rs("sid_tipo"), rs("par_value"), rs("par_value"))
			if rs("sid_tipo") = adBoolean then		'il booleano false corrisponde al salvataggio di stringa vuota
				if NOT valore then
					valore = ""
				end if
			end if
			if rs("sid_tipo") = adLongVarChar then
				rs("rsd_memo_it") = valore
			else
				rs("rsd_valore_it") = valore
			end if
			
			canc = canc &","& rs("par_id")
			rs.update
			rs.movenext
		wend
		
		rs.close
		set rs = nothing
		
		sql = "DELETE FROM tb_siti_parametri WHERE par_id IN ("& canc &")"
		connContent.Execute(sql)
	end if
End Sub



%>