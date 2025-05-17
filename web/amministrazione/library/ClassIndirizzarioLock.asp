
<%
const RANDOM_LOGIN_E_PASSWORD = "GENERA_RANDOM"


Class IndirizzarioLock
	Public Fields
	Public conn
	Private rs
	Public inInserimento
	
	'..................................................................................................................
	Private Sub Class_Initialize()
		'crea oggetto contenente i dati
		set Fields = Server.CreateObject("Scripting.Dictionary")
		Fields.CompareMode = vbTextCompare
		
		'connessione e recordset
		set conn = server.createobject("adodb.connection")
		conn.open Application("DATA_ConnectionString")
		set rs = server.createobject("adodb.recordset")
	end sub
	
	
	'..................................................................................................................
	Private Sub Class_Terminate()
		set Fields = nothing
		if IsObject(conn) then
			conn.close
			set conn = nothing
		end if
		set rs = nothing
	End Sub
	
	
	'..................................................................................................................
	Public sub ResyncConnection()
		dim ConnectionString
		ConnectionString = conn.ConnectionString
		conn.close
		conn.open ConnectionString
	end sub
	
	
	'..................................................................................................................
	public sub RemoveALL()
		'cancella valori precedenti
		Fields.RemoveAll
	end sub
		
	
	'..................................................................................................................
	'funzione che carica i dati di un contatto da database
	'..................................................................................................................
	public sub LoadFromDB(ID)
		dim sql, field
		sql = "SELECT * FROM tb_Indirizzario WHERE IDElencoIndirizzi=" & cIntero(ID)
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		
		if rs.eof then
			exit sub
		end if
		
		'cancella valori precedenti
		Fields.RemoveAll
		
		'inserisce valori da database
		for each field in rs.fields
			Fields.Add field.name, field.value
		next
		rs.close
		
		'inserisce dati aggiuntivi (valoriNumeri)
		sql = "SELECT TOP 1 ValoreNumero FROM tb_ValoriNumeri " &_
			  " WHERE id_Indirizzario=" & cIntero(ID) & " AND id_TipoNumero=<ID_TIPONUMERO> " + _
			  " ORDER BY "& SQL_OrderByBoolean(conn, "protetto_privacy", true) &", "& SQL_OrderByBoolean(conn, "email_default", true) & ", id_ValoreNumero"
		Fields.Add "telefono", GetValueList(conn, rs, replace(sql, "<ID_TIPONUMERO>", VAL_TELEFONO))
		Fields.Add "cellulare", GetValueList(conn, rs, replace(sql, "<ID_TIPONUMERO>", VAL_CELLULARE))
		Fields.Add "fax", GetValueList(conn, rs, replace(sql, "<ID_TIPONUMERO>", VAL_FAX))
		Fields.Add "email", GetValueList(conn, rs, replace(sql, "<ID_TIPONUMERO>", VAL_EMAIL))
		Fields.Add "web", GetValueList(conn, rs, replace(sql, "<ID_TIPONUMERO>", VAL_URL))
					
		'inserisce dati dell'eventuale utente associato
		sql = " SELECT * FROM tb_utenti " & _
			  " WHERE ut_nextCom_ID="& cIntero(ID)
			  '" LEFT JOIN gtb_rivenditori ON tb_utenti.ut_ID = gtb_rivenditori.riv_id " & _
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		if not rs.eof then
			Fields.Add "ut_id", rs("ut_id").value
			Fields.Add "login", rs("ut_login").value
			Fields.Add "password", rs("ut_password").value
			Fields.Add "abilitato", rs("ut_abilitato").value
			Fields.Add "scadenza", rs("ut_scadenzaAccesso").value
			'Fields.Add "attivo", CBoolean(rs("riv_attivo").value, false)
		else
			Fields.Add "ut_id", 0
			Fields.Add "login", ""
			Fields.Add "password", ""
			Fields.Add "abilitato", false
			Fields.Add "scadenza", ""
			'Fields.Add "attivo", false
		end if
		rs.close

		'inserisce rubrica
		sql = "SELECT TOP 1 id_rubrica FROM rel_rub_ind WHERE id_indirizzo=" & cIntero(ID)
		Fields.Add "rubrica", GetValueList(conn, rs, sql)

	end sub
	
	
	'..................................................................................................................
	'funzione che carica i dati di un contatto dal form
	'..................................................................................................................
	public sub LoadFromForm( checkbox_list )
		dim field, form_fields, i
		
		'cancella valori precedenti
		Fields.RemoveALL
		
		'inserisce valori da form
		for each field in request.form
			if lcase(left(field, 4))="tft_" OR _
				lcase(left(field, 4))="tfn_" OR _
				lcase(left(field, 4))="nfn_" OR _
				lcase(left(field, 4))="tfd_" then
				Fields.Add right(field, len(field)-4), request(field)
			end if
		next
		
		'carica campi checkbox
		form_fields = Split(checkbox_list, ";")
		for i=0 to uBound(form_fields)
			if instr(1, left(form_fields(i), 4), "chk_", vbTextCompare)>0 then
				field = right(Trim(form_fields(i)), len(Trim(form_fields(i)))-4)
			else
				field = Trim(form_fields(i))
			end if
			
			if field<>"" then
				if request(Trim(form_fields(i)))<>"" then
					Fields.Add field, true
				else
					Fields.Add field, false
				end if
			end if
		next
		
		'carica eventuale campo login
		if request.form("old_login") <> "" then
			Fields.Add "old_login", request.form("old_login")
		end if
	end sub
	
	
	'..................................................................................................................
	'controlla validita' campi
	'..................................................................................................................
	public function ValidateFields(Requested_Fields, CheckEmail)
		dim field_list,i, num_empty
		
		Session("ERRORE") = ""
		num_empty = 0
		field_list = Split(Requested_Fields, ";")
		for i=0 to uBound(field_list)
			if Trim(field_list(i))<>"" then
				if fields(Trim(field_list(i)))="" then
					num_empty = num_empty + 1
				end if
			end if
		next
		
		if num_empty>0 then
			'errore nella compilazione del form: mancano campi obbligatori
			Session("ERRORE") = ChooseValueByAllLanguages(session("lingua"), _
													  "Campi obbligatori non riempiti correttamente.", _
													  "Mandatory fields not filled correctly.", _
												 	  "Vorgeschriebene Felder nicht richtig gef&uuml;llt.", _
													  "Champs obligatoires non remplis correctement.", _
													  "Campos obligatorios no llenados correctamente", _
													  "Обязательные поля не заполнены правильно", _
													  "必填字段没有填写正确。", _
													  "O campo obrigatório não preenchido corretamente.")
		end if
		
		if CheckEmail then
			'verifica correttezza email solo se impostato errore relativo
			if not isEmail(Fields("email")) then
				'errore nell'email
				if Session("ERRORE")<>"" then
					Session("ERRORE")=Session("ERRORE") & "<br>"
				end if
				Session("ERRORE") = Session("ERRORE") & ChooseValueByAllLanguages(session("lingua"), _
													  	  "Errore nell'indirizzo email.", _
													  	  "Wrong email address.", _
												 		  "Falsches email adressen.", _
														  "Email adresse faux.", _
														  "Email direcci&oacute;n incorrecto.", _
														  "Ошибка в электронной почте.", _
														  "错误的电子邮件中。", _
														  "Erro no endereço de email.")
			end if
		end if
		
		'verifico correttezza data di nascita se inserita
		if request("tfd_DTNASCElencoIndirizzi") <> "" AND NOT IsDate(request("tfd_DTNASCElencoIndirizzi")) then
			if Session("ERRORE")<>"" then
				Session("ERRORE")=Session("ERRORE") & "<br>"
			end if
			Session("ERRORE") = Session("ERRORE") & ChooseValueByAllLanguages(session("lingua"), _
													  	  "Errore nella data di nascita.", _
													  	  "Wrong birth date format.", _
												 		  "Falsches Geburtsdatumformat", _
														  "Format de datede naissance faux.", _
														  "Formato incorrecto de fecha del nacimiento.", _
														  "Ошибка в дате рождения.", _
														  "错误的出生日期。", _
														  "Erro na data de nascimento.")
		end if
		
		ValidateFields = (Session("ERRORE") = "")
		
	end function
	
	
	'..................................................................................................................
	'funzione interna utilizzata per salvare i campi nel database
	'	ATTENZIONE: salva solo i campi che come gestione non differiscono tra inserimento e modifica del contatto
	'..................................................................................................................
	private sub SaveFields(rs)
		dim field
		
		for each field in rs.fields
'response.write "field: " & field & "<br>"
			if Fields.Exists(field.name) then
				if instr(1, Field.name, "IDElencoIndirizzi", vbTextCompare)<1 AND _
				   instr(1, Field.name, "isSocieta", vbTextCompare)<1 AND _
				   instr(1, Field.name, "ModoRegistra", vbTextCompare)<1 then 'salta la chiave primaria
					if instr(1, Field.name, "DTNASCElencoIndirizzi", vbTextCompare) > 0 OR _
					   instr(1, Field.name, "DataIscrizione", vbTextCompare) > 0 then
					   	'gestione campi di tipo data
						rs(field.name) = ConvertForSave_Date(Fields(field.name))
					elseif instr(1, Field.name, "google_maps_l", vbTextCompare) > 0 OR _
						   LCase(field.name) = "cntsede" then
						'gestione campi di tipo numerico
						rs(field.name) = ConvertForSave_Number(Fields(field.name), NULL)
					else
						'response.write "<br>fn:" & field.name & ":" & Fields(field.name)
						rs(field.name) = Fields(field.name)
					end if
				end if
			end if
		next
		
		if Fields.Exists("isSocieta") then
			if VarType(Fields("isSocieta")) = vbBoolean then
				rs("isSocieta") = Fields("isSocieta")
			else
				if cInteger(Fields("isSocieta")) = 0 OR (not isNumeric(Fields("isSocieta")) AND Fields("isSocieta")="") then
					rs("isSocieta") = false
				else
					rs("isSocieta") = true
				end if
			end if
		else
			if cInteger(request("isSocieta")) = 0 or instr(1, request("isSocieta"), "fal", vbTextCompare)>0 then
				rs("isSocieta") = false
			else
				rs("isSocieta") = true
			end if
		end if
		
		
		if rs("issocieta") then
			rs("ModoRegistra") = rs("NomeOrganizzazioneElencoIndirizzi")
		else
			rs("ModoRegistra") = rs("cognomeelencoindirizzi")
		end if

	end sub
	
	
	'..................................................................................................................
	'funzione che salva i dati di un nuovo contatto nel database
	'..................................................................................................................
	public function InsertIntoDB()
		dim sql, ID
		
		'salva dati della tabella principale
		sql = "SELECT * FROM tb_Indirizzario WHERE IDElencoIndirizzi=0"
		rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
		
		rs.AddNew
		
		rs("DataIscrizione") = Now
		
		CALL SetUpdateParamsRS(rs, "cnt_", true)
		
		CALL SaveFields(rs)
		
		if Fields("lingua")="" then
			if Session("LINGUA") <> "" then
				rs("lingua") = Session("LINGUA")
			else
				rs("lingua") = LINGUA_ITALIANO
			end if
		else
			rs("lingua") = Fields("lingua")
		end if
		
		rs.Update
		
		ID = rs("IDElencoIndirizzi")
		rs.close
		
		'aggiunge collegamento con rubica
		dim rubrica
		for each rubrica in split(replace(Fields("rubrica"), " ", ""), ",")
			CALL AddToRubrica(ID, rubrica)
		next
		
		if Fields("telefono")<>"" then
			CALL UpdateValoreNumero(ID, 1, false, "telefono")
		end if
		
		if Fields("cellulare")<>"" then
			CALL UpdateValoreNumero(ID, 3, false, "cellulare")
		end if
		
		if Fields("fax")<>"" then
			CALL UpdateValoreNumero(ID, 5, false, "fax")
		end if
		
		if Fields("email")<>"" then
			CALL UpdateValoreNumero(ID, 6, true, "email")
		end if
		
		if Fields("web")<>"" then
			CALL UpdateValoreNumero(ID, 7, false, "web")
		end if
		
		'rs.close
		
		'genero il codiceInserimento
		CALL SetCodiceInserimento(conn, ID)		
		
		InsertIntoDB = ID
	end function
	
	
	'..................................................................................................................
	'funzione che aggiunge il collegamento tra rubrica e contatto
	'..................................................................................................................
	public sub AddToRubrica(id_indirizzario, id_rubrica)
		dim sql
		if cInteger(id_indirizzario)>0 AND cInteger(id_rubrica)>0 then
			sql = " SELECT * FROM rel_rub_ind WHERE " + _
				  " id_indirizzo=" & id_indirizzario & _
				  " AND id_rubrica=" & id_rubrica
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
			if rs.eof then
				rs.AddNew
				rs("id_indirizzo") = id_indirizzario
				rs("id_rubrica") = id_rubrica
				rs.Update
			end if
			rs.close
		end if
	end sub
	
	
	'...................................................................................................................
	'	funzione che recupera l'id della rubrica con nome indicato, se non la trova la puo' inserire.
	'...................................................................................................................
	public function GetRubricaByName(rubricaNome, InsertIfNotFound)
		dim sql
		if cString(rubricaNome)<>"" then
			sql = "SELECT id_rubrica from tb_rubriche WHERE nome_Rubrica LIKE '" & ParseSQL(rubricaNome ,adChar) & "'"
			GetRubricaByName = cIntero(GetValueList(conn, rs, sql))
			if GetRubricaByName = 0 AND InsertIfNotFound then
				'non trova la rubrica: la inserisce
				sql = "SELECT * FROM tb_rubriche"
				rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
				rs.AddNew
				rs("nome_rubrica") = rubricaNome
				rs("rubrica_esterna") = false
				rs("locked_rubrica") = true
				rs("note_rubrica") = "Rubrica inserita automaticamente"
				rs.update
				GetRubricaByName = cIntero(rs("id_rubrica"))
				rs.close
			end if
		end if
	end function
	
	
	'..................................................................................................................
	'funzione che rimuove il collegamento tra rubrica e contatto
	'..................................................................................................................
    public sub RemoveFromRubrica(id_indirizzario, id_rubrica)
        dim sql
        if cInteger(id_indirizzario)>0 AND cInteger(id_rubrica)>0 then
            sql = "DELETE FROM rel_rub_ind WHERE " + _
                  " id_indirizzo=" & id_indirizzario & _
				  " AND id_rubrica=" & id_rubrica			  
            CALL conn.execute(sql, , adExecuteNoRecords)
        end if
    end sub
	
	
	'..................................................................................................................
	'funzione che aggiorna i dati di un contatto nel database
	'..................................................................................................................
	public function UpdateDB()
		dim sql
		
		'aggiorna dati principali del contatto
		sql = "SELECT * FROM tb_Indirizzario WHERE IDElencoIndirizzi=" & cIntero(Fields("IDElencoIndirizzi"))
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		CALL SaveFields(rs)
		
		CALL SetUpdateParamsRS(rs, "cnt_", false)
		
		rs.update
		rs.close
		
		'controlla ed eventualmente aggiunge collegamento con rubrica
		dim rubrica
		for each rubrica in split(replace(Fields("rubrica"), " ", ""), ",")
			if cInteger(rubrica)>0 then
                   CALL AddToRubrica(Fields("IDElencoindirizzi"), rubrica)
			end if
		next
		
		'aggiorna recapiti del contatto
		CALL UpdateValoreNumero(Fields("IDElencoIndirizzi"), 1, false, "telefono")
		CALL UpdateValoreNumero(Fields("IDElencoIndirizzi"), 3, false, "cellulare")
		CALL UpdateValoreNumero(Fields("IDElencoIndirizzi"), 5, false, "fax")
		CALL UpdateValoreNumero(Fields("IDElencoIndirizzi"), 6, true, "email")
		CALL UpdateValoreNumero(Fields("IDElencoIndirizzi"), 7, false, "web")
		
	end function
	
	
	'..................................................................................................................
	'funzione che aggiorna il valore.
	'..................................................................................................................
	Public sub UpdateValoreNumero(id_indirizzario, id_tipoNumero, email_Default, FieldName)
		dim sql
		if Fields.Exists(FieldName) then
			sql = "SELECT TOP 1 * FROM tb_ValoriNumeri WHERE id_indirizzario=" & cIntero(id_indirizzario) & _
				  " AND id_tipoNumero=" & cIntero(id_tipoNumero) & _
				  " ORDER BY "& SQL_OrderByBoolean(conn, "protetto_privacy", true) &", "& SQL_OrderByBoolean(conn, "email_default", true) & ", id_ValoreNumero"
				  'response.write sql & "<br>"
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	
			if rs.eof AND cString(Fields(FieldName))<>"" then
				'aggiunge nuovo record se valore <>"" 
				rs.AddNew
			elseif not rs.eof AND Fields(FieldName)="" then
				'cancella record se valore = ""
				rs.Delete adAffectCurrent
				rs.close
				exit sub
			elseif cString(Fields(FieldName))="" then
				rs.close
				exit sub
			end if
			
			'record presente ed univoco
			rs("id_indirizzario") = id_indirizzario
			rs("id_tipoNumero") = id_tipoNumero
			rs("ValoreNumero") = Left(Fields(FieldName), rs("ValoreNumero").DefinedSize)
			rs("protetto_privacy") = false
			rs("email_Default") = email_Default
			rs.update
			rs.close
		end if
	end sub
	
	
	'..................................................................................................................
	'procedura che aggiunge il reacapito direttamente al database
	'..................................................................................................................
	Public Sub AddValoreNumero(id_indirizzario, id_tipoNumero, email_Default, value)
		dim ok, sql
		if value<>"" then
			if id_tipoNumero=6 then		'verifica validità dell'email
				ok = isEmail(value)
			else
				ok = true
			end if
		else
			ok = false
		end if
		
		if ok then
			sql = "SELECT * FROM tb_ValoriNumeri WHERE id_indirizzario=" & cIntero(id_indirizzario) & _
				  " AND id_tipoNumero=" & cIntero(id_tipoNumero)
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			rs.addNew
			rs("id_indirizzario") = id_indirizzario
			rs("id_tipoNumero") = id_tipoNumero
			rs("ValoreNumero") = Left(value, rs("ValoreNumero").DefinedSize)
			rs("email_Default") = email_Default
			rs("protetto_privacy") = false
			rs.update
			rs.close
		end if
		
	end sub
	
	
	'..................................................................................................................
	'procedura che rimuove il reacapito
	'..................................................................................................................
	Public Sub RemoveValoreNumero(id_indirizzario, id_tipoNumero)
		sql = "DELETE FROM tb_valoriNumeri WHERE id_indirizzario=" & cIntero(id_indirizzario) & " AND id_tipoNumero=" & id_tipoNumero
		CALL conn.execute(sql)
	end sub
	
	
	'..................................................................................................................
	'blocca il contatto indicando che non puo' piu' essere cancellato
	'..................................................................................................................
	Public Sub LockContact(CntID, Abilitazione)
		dim sql, rsa
		
		'carica dati del contatto
		sql = " SELECT LockedByApplication, ApplicationsLocker FROM tb_Indirizzario WHERE IDElencoIndirizzi=" & cIntero(CntID)
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		if not rs.eof then
			
			if cIntero(Abilitazione)>0 then
				'lock con id di applicazione
				if not InIdList(cString(rs("ApplicationsLocker")), Abilitazione) then
					rs("LockedByApplication") = cInteger(rs("LockedByApplication")) + 1
					rs("ApplicationsLocker") = rs("ApplicationsLocker") & " " & Abilitazione & ","
					rs.update
				end if
			else
				'recupera lista delle applicazioni indicate dall'abilitazione
				set rsa = UserAbilitazione_Applicazioni(Abilitazione)
				
				while not rsa.eof
					if not InIdList(cString(rs("ApplicationsLocker")), rsa("id_sito")) then
						rs("LockedByApplication") = cInteger(rs("LockedByApplication")) + 1
						rs("ApplicationsLocker") = rs("ApplicationsLocker") & " " & rsa("id_sito") & ","
					end if
					rsa.movenext
				wend
				rs.update
			
				rsa.close
			end if
			
		end if
		rs.close
	end sub
	
	
	'..................................................................................................................
	'Sblocca il contatto dall'applicazione indicata
	'..................................................................................................................
	Public Sub UnLockContact(CntID, Abilitazione)
		dim sql, rsa
		set rsa = Server.CreateObject("ADODB.recordset")
		
		'carica dati del contatto
		sql = " SELECT LockedByApplication, ApplicationsLocker FROM tb_Indirizzario WHERE IDElencoIndirizzi=" & cIntero(CntID)
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		if not rs.eof then
			
			'recupera lista delle applicazioni indicate dall'abilitazione (nome del permesso)
			set rsa = UserAbilitazione_Applicazioni(Abilitazione)
			
			while not rsa.eof
				if InIdList(cString(rs("ApplicationsLocker")), rsa("id_sito")) then
					rs("LockedByApplication") = cInteger(rs("LockedByApplication")) - 1
					rs("ApplicationsLocker") = replace(cString(rs("ApplicationsLocker")), " " & rsa("id_sito") & ",", "")
				end if
				rsa.movenext
			wend
			rs.update
			
			rsa.close
			
		end if
		rs.close
	end sub
	
	
	'..................................................................................................................
	'valida login e pwd verificando anche se il login e' stato cambiato.
	'..................................................................................................................
	function ValidateLoginAndPassword(OldLogin, RetypedPassword)
		'controlla se password mancante
		if Fields("Password") = "" then
			Session("Errore") = ChooseValueByAllLanguages(session("lingua"), "Password mancante!", "Missing Password!", "Unzul&auml;ssiger Password!", "Password inadmissible!", "Password inv&aacute;lido", "Пароль хватает!", "密码丢失！", "Senha falta!")
			ValidateLoginAndPassword = false
			Exit function
		end if
		if Fields("Login") = "" then	
			Session("Errore") = ChooseValueByAllLanguages(session("lingua"), "Login mancante!", "Missing Login!", "Unzul&auml;ssiger Login!", "Login inadmissible!", "Login inv&aacute;lido", "Войти пропавших без вести!", "登录缺失！", "Log falta!")
			ValidateLoginAndPassword = false
			Exit function
		end if
		'controlla se la password e' uguale a quella di conferma
		if RetypedPassword = "" OR Fields("Password") <> RetypedPassword then
			Session("Errore") = ChooseValueByAllLanguages(session("lingua"), "Password non ridigitata correttamente!", _
													  "Password not retyped correctly!", _
												 	  "Kennwort nicht richtig neu getippt!", _
													  "Password non retap&eacute; correctement!", _
													  "Contrase&ntilde;a no escrita de nuevo correctamente!", _
													  "Пароль не перепечатанные правильно!", _
													  "重新输入密码不正确！", _
													  "A senha não retyped corretamente!")
			ValidateLoginAndPassword = false
			Exit function
		end if
		
		'controlla se la password e' composta di caratteri validi
		if not CheckChar(Fields("Password"), LOGIN_VALID_CHARSET) then
			Session("Errore") = ChooseValueByAllLanguages(session("lingua"), "Password non valida: usare solo lettere, numeri o &quot;_&quot;.", _
													  "Password not valid: use only letters, numbers or &quot;_&quot;.", _
													  "Password unzul&auml;ssig: benutzen Sie nur Buchstaben, Zahlen oder &quot;_&quot;", _
													  "Password inadmissible: employez seulement les lettres, nombres ou &quot;_&quot;", _
													  "Password inv&aacute;lido: utilice solamente las letras, n&uacute;meros o &quot;_&quot;", _
													  "Неверный пароль: Используйте только буквы, цифры или &quot;_&quot;", _
													  "无效的密码：仅可使用英文字母，数字或&quot;_&quot;", _
													  "Senha inválida: use somente letras, números ou &quot;_&quot;.")
			ValidateLoginAndPassword = false
			Exit function
		end if
		
		'esegue i controlli sul login solo se cambiato
		if uCase(Fields("Login")) <> Ucase(OldLogin) then
			'controlla che il login contenga solo caratteri validi
			if not CheckChar(Fields("Login"), LOGIN_VALID_CHARSET) then
				Session("Errore") = ChooseValueByAllLanguages(session("lingua"), "Login non valido: usare solo lettere, numeri o &quot;_&quot;.", _
														  "Login not valid: use only letters, numbers or &quot;_&quot;.", _
														  "Login unzul&auml;ssig: benutzen Sie nur Buchstaben, Zahlen oder &quot;_&quot;", _
														  "Login inadmissible: employez seulement les lettres, nombres ou &quot;_&quot;", _
														  "Login inv&aacute;lido: utilice solamente las letras, n&uacute;meros o &quot;_&quot;", _
														  "Неверное Логин: используйте только буквы, цифры или &quot;_&quot;", _
														  "无效的登入：仅使用字母，数字或&quot;_&quot;", _
														  "login inválido: use somente letras, números ou &quot;_&quot;.")
				ValidateLoginAndPassword = false
				Exit function
			end if
			
			'controlla che il login non sia gia' in uso da parte di un altro utente
			rs.open "SELECT (COUNT(*)) AS N_USERS FROM tb_utenti WHERE ut_login LIKE '" & ParseSql(Fields("login"), adChar) & "'", _
					conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			if rs("N_USERS") > 0 then
				Session("Errore") = ChooseValueByAllLanguages(session("lingua"), "Login gi&agrave; utilizzato da un altro utente, cambiare il login.", _
														  "Login already in use by another user, please change your login.", _
														  "Login bereits im Gebrauch durch einen anderen Benutzer, &auml;ndern Ihren login.", _
														  "Login d&eacute;j&agrave; en service d'un autre utilisateur, changent votre login", _
														  "La login ya en uso de otro usuario, cambia tu conexi&oacute;n.", _
														  "Логин уже используется другим пользователем, изменения Логин.", _
														  "由其他用户登录，已在利用变化的登录。", _
														  "Login já está em uso por outro usuário, mudar o login.")
				ValidateLoginAndPassword = false
				rs.close
				Exit function
			end if
			rs.close
		end if
		
		'se e' arrivato fino a qui la verifica ha dato esito positivo
		ValidateLoginAndPassword = (session("ERRORE") = "")
	end function
	
	
	'..................................................................................................................
	'imposta i dati dell'utente, eventualmente aggiornandoli
	'..................................................................................................................
	public function UserFromContact(CntID, Abilitazioni)
		dim ut_id, isNew, permesso, list, sql

		if Fields("Login") = RANDOM_LOGIN_E_PASSWORD AND Fields("Password") = RANDOM_LOGIN_E_PASSWORD AND Session("ERRORE") = "" then
		   	dim generati, count
			generati = false
			count = 0
			while not generati AND count < 100
				count = count + 1
				Fields("Login") = RemoveInvalidChar(left(cString(Fields("ModoRegistra")), 5), DOCUMENTS_FILES_CHARSET)
				Fields("Login") = uCase(Fields("Login") + right(FixLenght(cString(CntId), "0", 3), 3) + GetRandomString(DOCUMENTS_FILES_CHARSET, 7 - len(Fields("Login")) ) )
				Fields("Password") = uCase(GetRandomString(DOCUMENTS_FILES_CHARSET, 8))
				if ValidateLoginAndPassword("", Fields("Password")) then
					generati = true
				end if
				Session("ERRORE") = ""
			wend
			
		end if

		sql = "SELECT * FROM tb_utenti WHERE ut_nextCom_id=" & cIntero(CntID)
		rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
		if rs.eof then
			rs.AddNew
			isNew = true
		else
			isNew = false
		end if
		rs("ut_NextCom_ID") = CntID
		rs("ut_login") = Fields("Login")
		rs("ut_password") = Fields("Password")
		if Fields.Exists("Abilitato") then
			if VarType(Fields("Abilitato")) = vbBoolean then
				rs("ut_abilitato") = Fields("Abilitato")
			else
				if isNull(cInteger(Fields("Abilitato"))) then
					rs("ut_abilitato") = Fields("Abilitato")
				else
					rs("ut_abilitato") = cInteger(Fields("Abilitato")) > 0
				end if
			end if
		end if
		if Fields.Exists("Scadenza") then
			rs("ut_scadenzaAccesso") = ConvertForSave_Date(Fields("Scadenza"))
		end if
		
		rs.Update
		ut_ID = rs("ut_id")
		rs.close
		
		CALL UserAbilitazione_Add(CntId, ut_ID, abilitazioni)

		UserFromContact = ut_ID
	end function
	
	
	'..................................................................................................................
	'funzione che rimuove i dati utente di un contatto
	'..................................................................................................................
	sub RemoveUserFormContact(CntId, UtId, Applicazione)
		if cIntero(UtId) = 0 AND cIntero(CntId)>0 then
			sql = "SELECT ut_id FROM tb_utenti WHERE ut_nextcom_id=" & CntId
			UtId = cIntero(GetValueList(conn, rs, sql))
		elseif cIntero(UtId)>0 AND cIntero(CntId) = 0 then	
			sql = "SELECT ut_nextcom_id FROM tb_utenti WHERE ut_id=" & utId
			CntId = cIntero(GetValueList(conn, rs, sql))
		end if
		if cIntero(CntId) > 0 AND cIntero(UtId) > 0 then
			dim rsp
			set rsp = server.createObject("ADODB.recordset")
			
			'sblocca il contatto da tutte le applicazioni di area riservata per le quali e' abilitato
			sql = "SELECT * FROM rel_utenti_sito WHERE rel_ut_id=" & cIntero(UtId)
			rsp.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			if cIntero(Applicazione) > 0 then
				CALL UnLockContact(CntId, Applicazione)
			end if
			
			while not rsp.eof
				CALL UnLockContact(CntID, rsp("rel_sito_id"))
				rsp.movenext
			wend
			
			rsp.close
			set rsp = nothing
			
			'rimuove collegamenti a rubriche dell'area riservata
			sql = "DELETE FROM rel_rub_ind WHERE id_indirizzo = " & cIntero(CntId) & _
											"AND id_rubrica IN (SELECT sito_rubrica_area_riservata FROM tb_siti) "
			CALL conn.execute(sql, , adExecuteNoRecords)
				  
			'rimuove utente
			sql = "DELETE FROM tb_utenti WHERE ut_id=" & cIntero(UtId)
			CALL conn.execute(sql, , adExecuteNoRecords)
		end if
	end sub
	
	
	'sincronizza l'amministratore con i dati del contatto
	Function AdminFromContact(adminId, insIfDisable, sitoId, permessoNumero, dir)
		if adminId > 0 OR insIfDisable OR CBoolean(fields("abilitato"), false) then
			'controllo login
			if Cstring(fields("login")) <> "" then
				sql = "SELECT COUNT(*) FROM tb_admin WHERE admin_login LIKE '"& ParseSql(fields("login"), adChar) &"'"
				if adminId > 0 then
					sql = sql &" AND id_admin <> "& adminId
				end if
				if CIntero(GetValueList(conn, rs, sql)) > 0 then		
					session("ERRORE") = "Login gia' esistente."
				end if
			end if
			
			if session("ERRORE") = "" then
				rs.open "SELECT * FROM tb_admin WHERE id_admin = "& adminId, conn, adOpenKeySet, adLockOptimistic
				if adminId = 0 then
					rs.addnew
				end if
				
				if CBoolean(fields("isSocieta"), false) then
					rs("admin_cognome") = fields("nomeOrganizzazioneElencoIndirizzi")
				else
					rs("admin_cognome") = fields("cognomeElencoIndirizzi")
					rs("admin_nome") = fields("nomeElencoIndirizzi")
				end if
				if fields.Exists("email") then
					rs("admin_email") = fields("email")
				end if
				if fields.Exists("fax") then
					rs("admin_fax") = fields("fax")
				end if
				if fields.Exists("cellulare") then
					rs("admin_cell") = fields("cellulare")
				end if
				if fields.Exists("telefono") then
					rs("admin_telefono") = fields("telefono")
				end if
				rs("admin_dir") = dir
				
				'dati login
				if fields.Exists("login") then
					if fields("login") = "" then
						session("ERRORE") = "Login mancante!"
					end if
					rs("admin_login") = fields("login")
				end if
				if fields.Exists("password") then
					if fields("password") = "" then
						session("ERRORE") = "Password mancante!"
					end if
					rs("admin_password") = EncryptPassword(fields("password"))
				end if
				
				'abilitazione
				if IsDate(fields("scadenza")) then
					rs("admin_scadenza") = fields("scadenza")
				elseif cString(fields("scadenza"))="" then
					rs("admin_scadenza") = null
				end if
				if fields.Exists("abilitato") then
					if NOT CBoolean(fields("abilitato"), false) then
						rs("admin_scadenza") = DateIta(Date)
					end if
				end if
				
				'directory
				CALL CreateTemporaryDir(rs("admin_login"), IIF(adminId = 0, "", fields("old_login")))
				
				rs.update
				AdminFromContact = rs("id_admin")
				rs.close
				
				sql = " DELETE FROM rel_admin_sito WHERE admin_id = "& AdminFromContact & _
					  " AND sito_id = "& sitoId &" AND rel_as_permesso = "& permessoNumero
				conn.Execute(sql)
				sql = " INSERT INTO rel_admin_sito (admin_id, sito_id, rel_as_permesso)"& _
					  " VALUES ("& AdminFromContact &", "& sitoId &", "& permessoNumero &")"
				conn.Execute(sql)
			end if
		end if
	End Function
	
	
	'..................................................................................................................
	'restituisce il valore dell'elemento richiesto
	'..................................................................................................................
	Public Default Property Get Item(ByVal Key)
		Item = Fields(Key)
	end Property
	
	
	'..................................................................................................................
	'imposta il valore dell'elemento richiesto
	'..................................................................................................................
	Public Property Let Item(ByVal Key, ByVal Value)
		if Fields.Exists(Key) then
			Fields(Key) = Value
		else 
			Fields.Add Key, Value
		end if
	end Property
	
	
	'..................................................................................................................
	'..		procedura che aggiunge l'abilitazione all'utente
	'..................................................................................................................
	public sub UserAbilitazione_Add(CntId, UtId, abilitazioni)
		dim rsa, permessi, permesso, nAbilitazioni
		set rsa = Server.CreateObject("ADODB.recordset")
		
		set rsa = UserAbilitazione_Applicazioni(abilitazioni)

		while not rsa.eof
			permessi = UserAbilitazione_Permesso(rsa, abilitazioni)
			
			if permessi<>"" then
				nAbilitazioni = 0
				
				for each permesso in split(permessi, ",")
					if cIntero(permesso)>0 then
						sql = " SELECT * FROM rel_utenti_sito WHERE rel_ut_id=" & cIntero(UtId) & " AND rel_sito_id=" & rsa("id_sito") & _
							  " AND rel_permesso=" & permesso			
						rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
						nAbilitazioni = nAbilitazioni + 1			
						if rs.eof then
							rs.AddNew
							rs("rel_ut_id") = UtId
							rs("rel_sito_id") = rsa("id_sito")
							rs("rel_permesso") = permesso
							rs.Update
						end if
						rs.close
					end if
				next
				
				if nAbilitazioni > 0 then
					sql = "SELECT sito_rubrica_area_riservata FROM tb_siti WHERE id_sito = " & rsa("id_sito")
					CALL AddToRubrica(CntId, GetValueList(conn, rs, sql))
	
					'blocca il contatto
					CALL LockContact(CntId, rsa("id_sito"))
				end if
				
			end if
			
			rsa.movenext
		wend
	
		rsa.close
	end sub
	
	
	'..................................................................................................................
	'..		procedura che rimuove l'abilitazione all'utente
	'..................................................................................................................
	public sub UserAbilitazione_Remove(CntId, UtId, abilitazioni)
		dim rsa, permessi, permesso
		
		set rsa = UserAbilitazione_Applicazioni(abilitazioni)
		while not rsa.eof
			permessi = UserAbilitazione_Permesso(rsa, abilitazioni)
			
			if permessi<>"" then
				for each permesso in split(permessi, ",")
					if cIntero(permesso)>0 then
						sql = " DELETE FROM rel_utenti_sito " + _
							  " WHERE rel_ut_id=" & cIntero(UtId) & " AND rel_sito_id=" & rsa("id_sito") & _
							  " AND rel_permesso IN (" & permesso & ")"
						CALL conn.execute(sql, ,adCmdText)
					end if
				next
			end if
			
			'verifica se ci sono altre abilitazioni per quella applicazione.
			sql = "SELECT rel_id FROM rel_utenti_sito " + _
				  " WHERE rel_ut_id=" & cIntero(UtId) & " AND rel_sito_id=" & rsa("id_sito")
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
			if rs.eof then
				'rimuove rubrica collegata
				sql = "DELETE FROM rel_rub_ind WHERE id_indirizzo = " & cIntero(CntId) & _
											"AND id_rubrica IN (SELECT sito_rubrica_area_riservata FROM tb_siti WHERE id_sito=" & rsa("id_sito") & ") "
				CALL conn.execute(sql, ,adCmdText)
			end if
			rs.close
				
			CALL UnLockContact(CntID, rsa("id_sito"))
				  
			rsa.movenext
		wend
		
		rsa.close
		
	end sub
	
	
	'..................................................................................................................
	'..	recupera lista delle applicazioni abilitate
	'..................................................................................................................
	private function UserAbilitazione_Applicazioni(Abilitazioni)
		dim sql
		'carica dati delle applicazioni
		sql = "SELECT * FROM tb_siti WHERE "
		if cIntero(Abilitazioni) > 0 AND instr(1, cString(Abilitazioni), ",", vbTextCompare) < 1 then
			'id dell'applicazione da abilitare
			sql = sql + " id_sito = " & Abilitazioni
		else
			'permesso da ricercare
			dim lista, permesso
			lista = split(replace(Abilitazioni, " ", ","), ",")
			sql = sql + " NOT " & SQL_IsTrue(conn, "sito_amministrazione") & " AND ( (1=0) "
			for each permesso in lista 
				if permesso <> "" then
					sql = sql + " OR " + SQL_TextSearch(permesso, "sito_p1;sito_p2;sito_p3;sito_p4;sito_p5;sito_p6;sito_p7;sito_p8;sito_p9", false)
				end if
			next
			sql = sql + ")"
		end if
		
		set UserAbilitazione_Applicazioni =  server.createObject("ADODB.recordset")
		UserAbilitazione_Applicazioni.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		
	end function
	
	
	'..................................................................................................................
	'..	recupera lista dei permessi delle abilitazioni
	'..................................................................................................................
	private function UserAbilitazione_Permesso(rsa, Abilitazioni)
		
		if cIntero(Abilitazioni) > 0 AND instr(1, cString(Abilitazioni), ",", vbTextCompare) < 1 then
			UserAbilitazione_Permesso = "1"
		else
			dim field, listaAbilitazioni, valueAbilitazione
			listaAbilitazioni = split(replace(Abilitazioni, " ", ","), ",")
			UserAbilitazione_Permesso = ""
			
			for each field in rsa.fields
				if instr(1, field.name, "sito_p", vbTextCompare)>0 then
					for each valueAbilitazione in listaAbilitazioni
						if lcase(trim(valueAbilitazione)) = lcase(trim(cstring(field.value))) then
							UserAbilitazione_Permesso = IIF(cString(UserAbilitazione_Permesso)<>"", UserAbilitazione_Permesso & ",", "") & cIntero(replace(field.name, "sito_p", ""))
						end if
					next
				end if
			next
			if UserAbilitazione_Permesso = "" then
				UserAbilitazione_Permesso = "1"
			end if
		end if

	end function
	
	
end class



'..................................................................................................................
'gestisce i campi esterni a IndirizzarioLock come se fossero interni
'..................................................................................................................
Sub CaricaCampiEsterni(conn, rs, contatto, sql, id_nome, id_value)
	CALL CaricaAggiornaCampiEsterni(conn, rs, contatto, sql, id_nome, id_value, true)
end sub

Sub CaricaAggiornaCampiEsterni(conn, rs, contatto, sql, id_nome, id_value, CaricaTutti)
	dim campo, inp, nome
	if sql<>"" then
		if CInteger(id_value) > 0 then
			sql = sql & " WHERE "& id_nome &"="& cIntero(id_value)
		end if
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	end if
	for each campo in rs.Fields
		if CInteger(id_value) > 0 then	
			if CaricaTutti OR not contatto.Fields.exists(campo.Name) then
				contatto(campo.Name) = rs(campo.Name).Value
			end if
		else
			nome = ""
			for each inp in request.Form
				if InStr(1, inp, "ext", vbTextCompare) > 0 AND InStr(1, inp, campo.Name, vbTextCompare) > 0 then
					nome = inp
					exit for
				end if
			next
			contatto(campo.Name) = request.Form(nome)
		end if
	next
	if sql <>"" then
		rs.Close
	end if
End Sub


%>
