<% 
Private Const R_FIELD = 0		'posizione del campo di origine mappato nell'array multidimensionale dei recapiti
Private Const R_TYPE = 1		'posizione del tipo di valore nell'array multidimensionale dei recapiti
Private Const R_DEFAULT = 2		'posizione del valore se l'email e' utilizzata o meno nelle mailing list
Private Const R_PRIVACY = 3		'posizione del valore se il campo e' protetto da privacy o meno


class IndirizzarioSyncro
	
	Private rs, rsI
	Private RecapitiCount
	Private Recapiti()		'elenco di recapiti: ARRAY a 3 dimensioni: 		0 = campo di origine mappato
							'												1 = tipo di valore da inserire (Secondo costanti definite in Tools.asp)
								
	'variabili per la mappatura dei campi di origine
	Public F_Nome				'NomeElencoIndirizzi
	Public F_SecondoNome		'SecondoNomeElencoIndirizzi
	Public F_Cognome			'CognomeElencoIndirizzi
	Public F_Titolo				'TitoloElencoIndirizzi
	Public F_Organizzazione		'NomeOrganizzazioneElencoIndirizzi
	Public F_Qualifica			'QualificaElencoIndirizzi
	Public F_Indirizzo			'IndirizzoElencoIndirizzi
	Public F_Localita			'LocalitaElencoIndirizzi
	Public F_Citta				'CittaElencoIndirizzi
	Public F_Provincia			'StatoProvElencoIndirizzi
	Public F_Zona				'ZonaElencoIndirizzi
	Public F_CAP				'CAPElencoIndirizzi
	Public F_Country			'CountryElencoIndirizzi
	Public F_DataNascita		'DTNASCElencoIndirizzi
	Public F_LuogoNascita		'LuogoNascita
	Public F_CodiceFiscale		'CF
	Public F_lingua				'lingua
	Public F_CntRel				'collegamento indirizzo "padre"
	Public F_IsSocieta			'indica se il record e' una societa' o un contatto privato
	
	'variabili per impostare direttamente dei valori di default dei campi
	Public V_Nome				'valore per il campo NomeElencoIndirizzi
	Public V_SecondoNome		'valore per il campo SecondoNomeElencoIndirizzi
	Public V_Cognome			'valore per il campo CognomeElencoIndirizzi
	Public V_Titolo				'valore per il campo TitoloElencoIndirizzi
	Public V_Organizzazione		'valore per il campo NomeOrganizzazioneElencoIndirizzi
	Public V_Qualifica			'valore per il campo QualificaElencoIndirizzi
	Public V_Indirizzo			'valore per il campo IndirizzoElencoIndirizzi
	Public V_Localita			'valore per il campo LocalitaElencoIndirizzi
	Public V_Citta				'valore per il campo CittaElencoIndirizzi
	Public V_Provincia			'valore per il campo StatoProvElencoIndirizzi
	Public V_Zona				'valore per il campo ZonaElencoIndirizzi
	Public V_CAP				'valore per il campo CAPElencoIndirizzi
	Public V_Country			'valore per il campo CountryElencoIndirizzi
	Public V_DataNascita		'valore per il campo DTNASCElencoIndirizzi
	Public V_LuogoNascita		'valore per il campo LuogoNascita
	Public V_CodiceFiscale		'valore per il campo CF
	Public V_lingua				'valore per il campo lingua
	Public V_CntRel				'valore per il collegamento indirizzo "padre"
	Public V_IsSocieta			'valore che indica se il record e' una societa' o un contatto privato
	
	'variabili per il blocco del contatto
	Public Table				'tabella di origine dei dati
	Public BlockField			'campo per bloccare la sincronizzazione se falso (se  campo non impostato sincronizza sempre)
	Public FilterTable			'tabella di origine del filtro collegato alla rubrica
	Public FilterField			'campo della tabella di origine usato come filtro per inserire la rubrica figlia (tipo o categoria del record sorgente)
	Public KeyValue				'id del record origine dei dati
	Public ApplicationID		'id dell'applicazione che gestisce la tabella 
	
	Public CntId				'Id del contatto inserito
	
	'**************************************************************************************************************
	'FUNZIONI DI INIZZIALIZZAZIONE
	Private Sub Class_Initialize()
		'crea recordset
		set rs = server.CreateObject("ADODB.recordset")		'recordset sorgente dati
		set rsI = server.CreateObject("ADODB.recordset")	'recordset indirizzario
		
		'imposta le dimensioni dell'array dei recapiti (dimensione 3 colonne X n righe)
		ReDim Recapiti(3,0)
		RecapitiCount = 0
		
		'imposta valori di base per campi sensibili
		V_DataNascita = NULL
		V_IsSocieta = true
		V_lingua = LINGUA_ITALIANO
	end sub
	
	
	Private Sub Class_Terminate()
		set rs = nothing
		set rsI = nothing
	end sub
	
	'**************************************************************************************************************
	'FUNZIONI PUBBLICHE
	
	'............................................................
	'Aggiunge un collegamento per un recapito alla lista di recapiti
	Public sub Recapito(Field, Typ)
		CALL RecapitoPrivacy(Field, Typ, (Typ = VAL_EMAIL), false)
	end sub
	
	Public sub RecapitoPrivacy(Field, Typ, EmailDefault, Protetto)
		'aumenta la dimensione dell'array dei recapiti per aggiungerne uno
		Redim Preserve Recapiti(3, RecapitiCount + 1)
		
		Recapiti(R_FIELD, RecapitiCount) = Field			'campo collegato
		Recapiti(R_TYPE, RecapitiCount) = Typ				'tipo di valore
		Recapiti(R_DEFAULT, RecapitiCount) = EmailDefault	'email usata nelle mailing list
		Recapiti(R_PRIVACY, RecapitiCount) = Protetto		'valore protetto da privacy
		'aumenta conteggio dei recapiti collegati
		RecapitiCount = RecapitiCount + 1		
	end sub
	
	
	Public Sub Syncronize(sql, conn)
		dim sqlI, Adding, ActivateTransaction
		'attiva transazione
		if not isEmpty(conn) then
			ActivateTransaction = false
		else
			ActivateTransaction = true
			set conn = server.CreateObject("ADODB.connection")
			conn.open Application("DATA_ConnectionString"), "", ""
			conn.BeginTrans 
		end if
		
		'controlla se le proprieta' di base sono impostate correttamente
		CALL Validate(Table, "table")
		CALL Validate(KeyValue, "KeyValue")
		CALL Validate(ApplicationID, "ApplicationID")
		
		'apre recordset sorgente
		rs.open sql, conn, adOpenStatic, adLockOptimistic
		if rs.recordcount=1 then		'controlla che il record e' stato individuato univocamente
			if cString(BlockField)<>"" then
				if not rs(BLockField) OR ISNULL(rs(BLockField)) then
					'salta la sincornizzazione perche' il campo di blocco e' impostato su false
					rs.close
					if ActivateTransaction then
						conn.rollbacktrans
					end if
					exit sub
				end if
			end if
	
			'controlla se e' presente almeno una rubrica collegabile al contatto
			sqlI = "SELECT * FROM tb_rubriche WHERE SyncroTable LIKE '" & Table & "' " & _
				   " AND (" & SQL_IsNULL(conn, "SyncroFilterKey") & " OR SyncroFilterKey=0) AND " & SQL_isTrue(conn, "rubrica_esterna")
			rsI.open sqlI, conn, adOpenstatic, adLockOptimistic, adCmdText
			if cInteger(rsI.recordcount)<1 then
				'salta la sincronizzazione perche' la tabella non ha la rubrica collegata direttamente
				Session("ERRORE") = "Errore nella gestione della rubrica per l'inserimento nei contatti."
				rs.close
				if ActivateTransaction then
					conn.rollbacktrans
				end if
				exit sub
			end if
			rsI.close
			
			'compone query per indirizzario
			sqlI = "SELECT * FROM tb_Indirizzario WHERE SyncroKey LIKE '" & KeyValue & "' " &_
				  " AND SyncroApplication=" & ApplicationID & " AND SyncroTable LIKE '" & Table & "'"
			rsI.open sqlI, conn, adOpenKeySet, adLockOptimistic, adCmdText
			if rsI.recordcount<1 then
				'record collegato non presente : lo aggiunge
				rsi.AddNew
				rsI("SyncroKey") = KeyValue
				rsI("SyncroTable") = Table
				rsI("SyncroApplication") = ApplicationID
				Adding = true
			else
				Adding = false
			end if
			
			'imposta i valori della tabella
			CALL UpdateValue(rs, rsI, "NomeElencoIndirizzi", 				F_Nome, 			V_nome)
			CALL UpdateValue(rs, rsI, "SecondoNomeElencoIndirizzi", 		F_SecondoNome, 		V_SecondoNome)
			CALL UpdateValue(rs, rsI, "CognomeElencoIndirizzi", 			F_Cognome, 			V_Cognome)
			CALL UpdateValue(rs, rsI, "TitoloElencoIndirizzi", 				F_Titolo, 			V_Titolo)
			CALL UpdateValue(rs, rsI, "NomeOrganizzazioneElencoIndirizzi", 	F_Organizzazione,	V_Organizzazione)
			CALL UpdateValue(rs, rsI, "QualificaElencoIndirizzi", 			F_Qualifica,		V_Qualifica)
			CALL UpdateValue(rs, rsI, "IndirizzoElencoIndirizzi", 			F_Indirizzo,		V_Indirizzo)
			CALL UpdateValue(rs, rsI, "LocalitaElencoIndirizzi", 			F_Localita,			V_Localita)
			CALL UpdateValue(rs, rsI, "CittaElencoIndirizzi", 				F_Citta,			V_Citta)
			CALL UpdateValue(rs, rsI, "StatoProvElencoIndirizzi", 			F_Provincia,		V_Provincia)
			CALL UpdateValue(rs, rsI, "ZonaElencoIndirizzi", 				F_Zona,				V_Zona)
			CALL UpdateValue(rs, rsI, "CAPElencoIndirizzi", 				F_CAP,				V_CAP)
			CALL UpdateValue(rs, rsI, "CountryElencoIndirizzi", 			F_Country,			V_Country)
			CALL UpdateValue(rs, rsI, "DTNASCElencoIndirizzi", 				F_DataNascita,		V_DataNascita)
			CALL UpdateValue(rs, rsI, "LuogoNascita", 						F_LuogoNascita,		V_LuogoNascita)
			CALL UpdateValue(rs, rsI, "CF", 								F_CodiceFiscale,	V_CodiceFiscale)
			CALL UpdateValue(rs, rsI, "lingua", 							F_lingua,			V_lingua)
			CALL UpdateValue(rs, rsI, "CntRel", 							F_CntRel,			V_CntRel)
			CALL UpdateValue(rs, rsI, "IsSocieta",							F_IsSocieta,		V_IsSocieta)
	
			if rsI("IsSocieta") AND cString(rsI("NomeOrganizzazioneElencoIndirizzi"))<>"" then
				rsI("isSocieta") = true
			else
				rsI("isSocieta") = false
			end if
			
			if rsI("isSocieta") then
				rsI("ModoRegistra") = rsI("NomeOrganizzazioneElencoIndirizzi")
			else
				rsI("ModoRegistra") = rsI("CognomeElencoIndirizzi")
			end if

			CALL SetUpdateParamsRS(rsI, "cnt_", Adding)
			
			rsI.Update
			'recupera id del contatto
            
			CntId = rsI("IDElencoIndirizzi")
			rsI.close

			'gestione recapiti
			dim i, j
			for i=0 to (RecapitiCount-1)
				'apre recordset su recapito corrente
				sqlI = " SELECT * FROM tb_ValoriNumeri WHERE id_indirizzario=" & CntId & _
					   " AND id_TipoNumero=" & recapiti(R_TYPE, i) & " AND SyncroField LIKE '" & recapiti(R_FIELD, i) & "'"
				rsI.open sqlI, conn, adOpenStatic, adLockOptimistic, adCmdText

				if rsI.eof and Trim(cString(rs(recapiti(R_FIELD,i))))<>"" then
					if (recapiti(R_TYPE, i) <> VAL_EMAIL) OR _
						isEmail(rs(recapiti(R_FIELD,i))) then
						
						'inserisce il recapito
						rsI.AddNew
						rsI("id_indirizzario") = CntId
						rsI("id_TipoNumero") = recapiti(R_TYPE, i)
						rsI("SyncroField") = recapiti(R_FIELD, i)
						rsI("ValoreNumero") = rs(recapiti(R_FIELD,i))
						rsI("email_default") = recapiti(R_DEFAULT, i)
						rsI("protetto_privacy") = recapiti(R_PRIVACY, i)
						rsI.Update
					end if
				elseif not rsI.eof and cString(rs(recapiti(R_FIELD,i)))<>"" then
					'aggiorna il recapito
					rsI("ValoreNumero") = rs(recapiti(R_FIELD,i))
					rsI("email_default") = recapiti(R_DEFAULT, i)
					rsI("protetto_privacy") = recapiti(R_PRIVACY, i)
					rsI.Update
				 elseif not rsI.eof and cString(rs(recapiti(R_FIELD,i)))="" then
					'cancella il recapito
					rsI.Delete
				end if
				rsI.close
			next

			
			
			'aggiunge contatto alla rubrica principale per la tabella se in inserimento
			if Adding then
				sqlI = "INSERT INTO rel_rub_ind (id_indirizzo, id_rubrica) " &_
					  " SELECT " & cIntero(CntId) & ", id_rubrica FROM tb_rubriche WHERE SyncroTable LIKE '" & ParseSql(Table, adChar) & "' " &_
					  " AND (" & SQL_IsNULL(conn, "SyncroFilterKey") & " OR SyncroFilterKey=0) AND " & SQL_isTrue(conn, "rubrica_esterna")
				CALL conn.execute(sqlI, vbTrue, adExecuteNoRecords)
			end if

			'gestione rubrica figlia collegata alle categorie/tipi dei record (filterValue)
			'cancella eventuali collegamenti con rubriche figlie se non in modifica
			if not Adding then
				sqlI = "DELETE FROM rel_rub_ind WHERE id_indirizzo=" & cIntero(CntId) & " AND " &_ 
					   " id_rubrica IN (SELECT id_rubrica FROM tb_rubriche WHERE SyncroTable LIKE '" & ParseSql(Table, adchar) & "' " &_
					   " AND SyncroFilterTable LIKE '" & ParseSql(FilterTable, adChar) & "' " &_
					   " AND NOT(" & SQL_IsNULL(conn, "SyncroFilterKey") & ")  AND " & SQL_isTrue(conn, "rubrica_esterna") & ")"
				CALL conn.execute(sqlI, vbTrue, adExecuteNoRecords)
			end if
			
			if cString(FilterField)<>"" then
				if cInteger(rs(FilterField))>0 then
					'inserisce associazione con rubriche figlie
					sqlI = "INSERT INTO rel_rub_ind (id_indirizzo, id_rubrica) " &_
						  " SELECT " & cIntero(CntId) & ", id_rubrica FROM tb_rubriche WHERE SyncroTable LIKE '" & ParseSql(Table, adChar) & "' " &_
					   	  " AND SyncroFilterTable LIKE '" & ParseSql(FilterTable, adChar) & "' " &_
						  " AND SyncroFilterKey=" & rs(FilterField) & " AND " & SQL_isTrue(conn, "rubrica_esterna")
				CALL conn.execute(sqlI, vbTrue, adExecuteNoRecords)
				end if
			end if
			
		end if
		
		rs.close
	
		'chiude transazione
		if ActivateTransaction then
			conn.CommitTrans
			conn.close
			set conn = nothing
		end if

	end sub
	
	
	Public Sub DeleteSyncro(conn)
		dim sql
		sql = "DELETE FROM tb_indirizzario WHERE SyncroKey LIKE '" & KeyValue & "' " &_
				  " AND SyncroApplication=" & ApplicationID & " AND SyncroTable LIKE '" & Table & "'"
		CALL conn.execute(sql, vbTrue, adExecuteNoRecords)
	end sub
	
	'**************************************************************************************************************
	'FUNZIONI PRIVATE
	
	'............................................................
	'esegue un controllo se il parametro passato contiene un valore vuoto allora blocca l'esecuzione
	Private Sub Validate(value, var)
		if cString(Value)="" then
			response.write "manca valore necessario per la classe (" & var & ")!!"
			response.end
		end if
	end sub
	
	'............................................................
	'aggiorna il valore del campo: se il campo e' vuoto inserisce il valore di default
	Private Sub UpdateValue(rs, rsI, FieldI, Field, Default)
		if Field<>"" then
			if isNull(rs(Field)) OR cString(rs(Field))="" then
				rsI(FieldI) = Default
			else
				rsI(FieldI) = rs(Field)
			end if
		else
			rsI(FieldI) = Default
		end if
	end sub
	
end class



'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'FUNZIONI DI UTILITA PER RUBRICHE
'**************************************************************************************************************

'..............................................................................................................
' procedura che inserisce ed eventualmente aggiorna i dati della rubrica per la sincronizzazione tipizzata
' la terna di parametri SyncroTable, SyncroFilterTable, SyncroFilterKey individua univocamente la rubrica
'conn				connessione aperta sul database da utilizzare
'rs					oggetto recordset creato e chiuso
'NomeRubrica		nome della rubrica (Apparira' nel NEXT-COM)
'SyncroTable		Nome tabella principale del gruppo di sincronizzazione (TABELLA contenente i record syncronizzati)
'SyncroFilterTable	Nome tabella che categorizza i record sincronizzati
'SyncroFilterKey	ID della categoria che tipizza i dati
'..............................................................................................................
sub UpdateSyncroRubrica(conn, rs, NomeRubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey)
    CALL UpdateSyncroRubricaGruppo(conn, rs, NomeRubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, NULL)
end sub

function UpdateSyncroRubricaGruppo(conn, rs, NomeRubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, GruppoLavoro)
	dim sql, IdRubrica, isNew
	'cerca se esiste gi&agrave; la rubrica sincronizzata richiesta
	sql = " SELECT * FROM tb_rubriche WHERE SyncroFilterKey = " & SyncroFilterKey & _
          IIF(cString(SyncroTable)<>"", " AND SyncroTable LIKE '" & SyncroTable & "' ", "") & _
          IIF(cString(SyncroFilterTable)<>"", " AND SyncroFilterTable LIKE '" & SyncroFilterTable & "' ", "")
	rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	if rs.eof then
		'rubrica non trovata: la aggiunge impostando i campi di base
		isNew = true
		rs.AddNew
		rs("SyncroFilterTable") = SyncroFilterTable
		rs("SyncroTable") = SyncroTable
		rs("SyncroFilterKey") = SyncroFilterKey
		rs("locked_rubrica") = true
		rs("rubrica_esterna") = true
	else
		isNew = false
	end if
	'aggiorna i campi esterni
	rs("nome_Rubrica") = NomeRubrica
	rs.update
	'recupera id della rubrica
	IdRubrica = rs("id_Rubrica")
	rs.close
	
	if isNew AND cInteger(GruppoLavoro)=0 then
        'se la rubrica &egrave; appena stata inserita la collega a tutti i gruppi di lavoro
		sql = " INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & _
			  " SELECT " & IdRubrica & ", id_gruppo FROM tb_gruppi "
		CALL conn.execute(sql, , adExecuteNoRecords)
    else
        sql = "SELECT * FROM tb_rel_gruppiRubriche WHERE id_dellaRubrica=" & IdRubrica & " AND id_gruppo_assegnato=" & GruppoLavoro
	'response.write sql
        rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
        if rs.eof then
            rs.AddNew
            rs("id_dellaRubrica") = IdRubrica
            rs("id_gruppo_assegnato") = GruppoLavoro
            rs.Update
        end if
        rs.close
	end if
    
    UpdateSyncroRubricaGruppo = IdRubrica
end function


'..............................................................................................................
'procedura che rimuove la rubrica per la sincronizzazione tipizzata
' la terna di parametri SyncroTable, SyncroFilterTable, SyncroFilterKey individua univocamente la rubrica
'conn				connessione aperta sul database da utilizzare
'SyncroTable		Nome tabella principale del gruppo di sincronizzazione (TABELLA contenente i record syncronizzati)
'SyncroFilterTable	Nome tabella che categorizza i record sincronizzati
'SyncroFilterKey	ID della categoria che tipizza i dati
'..............................................................................................................
sub DeleteSyncroRubrica(conn, SyncroTable, SyncroFilterTable, SyncroFilterKey)
	dim sql
	sql = " DELETE FROM tb_rubriche WHERE SyncroFilterKey = " & SyncroFilterKey & _
          IIF(cString(SyncroTable)<>"", " AND SyncroTable LIKE '" & SyncroTable & "' ", "") & _
          IIF(cString(SyncroFilterTable)<>"", " AND SyncroFilterTable LIKE '" & SyncroFilterTable & "' ", "")
	CALL conn.execute(sql, , adExecuteNoRecords)
end sub 


'..............................................................................................................
'restituisce l'id della rubrica sincronizzata con i parametri indicati
'conn				connessione aperta sul database da utilizzare
'rs                 oggetto recordset per recuperare i dati (se NULL lo crea autonomamente)
'SyncroTable		Nome tabella principale del gruppo di sincronizzazione (TABELLA contenente i record syncronizzati)
'SyncroFilterTable	Nome tabella che categorizza i record sincronizzati
'SyncroFilterKey	ID della categoria che tipizza i dati
'..............................................................................................................
function GetSyncroRubrica(conn, rs, SyncroTable, SyncroFilterTable, SyncroFilterKey)
    
    dim sql
	sql = " SELECT id_rubrica FROM tb_rubriche WHERE SyncroFilterKey = " & SyncroFilterKey & _
          IIF(cString(SyncroTable)<>"", " AND SyncroTable LIKE '" & SyncroTable & "' ", "") & _
          IIF(cString(SyncroFilterTable)<>"", " AND SyncroFilterTable LIKE '" & SyncroFilterTable & "' ", "")
    GetSyncroRubrica = cInteger(GetValueList(conn, rs, sql))
    
end function

%>