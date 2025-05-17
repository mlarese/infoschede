<% 

'.................................................................................................
'..                     Crea il pulsante di submit
'..                     tag_value	    = etichetta visualizzata sul submit
'..						tag_name		= nome del submit
'..						Image_SX		= immagine alla sinistra del submit
'..						Image_DX		= immagine alla destra del submit
'.................................................................................................
sub RenderSubmit(tag_value, tag_name, Image_SX, Image_DX, style)
	dim input_definition
	if tag_name="annulla" then
		input_definition = "<input type=""reset"" " 
	else
		input_definition = "<input type=""submit"" " 
	end if
	input_definition = input_definition & "name=""" & tag_name & """ value=""" & tag_value & """ class=""" & style & """>"
	if Image_SX & Image_DX<>"" AND right(image_sx, 1) <> "/" AND right(image_dx, 1) <> "/"  then%>
		<table border="0" cellspacing="0" cellpadding="0" align="right">
			<tr>
				<% if Image_SX<>"" then %>
					<td>
						<img src="<%= image_SX %>" alt="" border="0">
					</td>
				<% end if %>
				<td style="vertical-align:top;">
					<%= input_definition %>
				</td>
				<% if Image_DX<>"" then %>
					<td>
						<img src="<%= image_DX %>" alt="" border="0">
					</td>
				<% end if %>
			</tr>
		</table>
	<% else 
		response.write input_definition
	end if
end sub




'*****************************************************************************************************************
'..		funzione che esegue l'accesso all'area riservata impostando tutte le variabili di sessione per l'utente
'..		oppure, se login e/o password errati restituisce i valori di errore nella variabile di sessione("ERRORE")
'..			login				login dell'utente
'..			password			password dell'utente
'..			errore_login		valore da impostare in Session("ERRORE") in caso di login errato o non valido
'..			errore_password		valore da impostare in Session("ERRORE") in caso di password errata o non valida
'*****************************************************************************************************************
function CheckLogin(login, password)
	dim conn, rs, sql, field
	login = Ucase(Trim(login))
	password = Ucase(Trim(password))
	if login<>"" AND CheckChar(login, LOGIN_VALID_CHARSET) then
		if password<>"" AND CheckChar(password, LOGIN_VALID_CHARSET) then
			'Verifica permessi di accesso
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.open Application("DATA_ConnectionString"),"",""
			Set rs = Server.CreateObject("ADODB.RecordSet")
			'cerca utente con login richiesto
			sql = " SELECT * FROM tb_utenti WHERE ut_login LIKE '" & ParseSQL(login, adChar) & "' " &_
				  " AND UT_id IN (SELECT rel_ut_id FROM rel_utenti_sito) "
			rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			if rs.eof then
				Session("ERRORE") = ChooseByLanguage("Login non valido.", _
													 "Invalid login.", _
													 "Unzulässiger logon.", _
													 "Login inadmissible.", _
													 "Login Inválido.")
				Session("ERRORE_CODICE") = "LOGIN_ERROR"
			else
				'utente trovato: verifica password
				if UCASE(password) = UCASE(cString(rs("ut_password"))) then
					'password OK
					if rs("ut_abilitato") then
						'utente abilitato
						if isNULL(rs("ut_scadenzaAccesso")) OR rs("ut_ScadenzaAccesso") >= Date then
							'accesso valido
							Session("USER_4_LOG") = login
							Session("UT_ID") = rs("ut_ID")
							Session("COM_ID") = rs("ut_NextCom_ID")
							rs.close
							'legge permessi
							sql = "SELECT * FROM rel_utenti_sito INNER JOIN tb_siti ON rel_utenti_sito.rel_sito_id = tb_siti.id_sito " &_
								  " WHERE rel_utenti_sito.rel_ut_id=" & cIntero(Session("UT_ID"))
							rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
							'imposta tutte le variabili di sessione relative ai permessi per ogni applicazione abilitata
							while not rs.eof
								field = "sito_p" & rs("rel_permesso")
								Session(rs(field)) = login
								rs.MoveNext
							wend
							
							'registra accesso nel log
							sql = "INSERT INTO log_utenti (log_ut_id, log_sito_id, log_data, log_username) " &_
								  " SELECT " & cIntero(Session("UT_ID")) & ", id_sito, " & SQL_NOW(conn) & ", '" & ParseSql(Session("USER_4_LOG"), adChar) & "' FROM " &_
								  " tb_siti WHERE NOT " & SQL_IsTrue(conn, "sito_amministrazione")
							CALL conn.execute(sql, , adExecuteNoRecords)
						else
							Session("ERRORE") = ChooseByLanguage("Profilo di accesso scaduto.", _
																 "Account expired.", _
																 "Account ablaufen.", _
																 "Account expir&eacute;.", _
																 "Account expirado.")
							Session("ERRORE_CODICE") = "ACCOUNT_EXPIRED"
						end if
					else
						Session("ERRORE") = ChooseByLanguage("Accesso non consentito.", _
															 "Access denied.", _
															 "Zugang verweigert.", _
															 "L'acc&eacute;s a ni&egrave;.", _
															 "El acceso neg&oacute;.")
							Session("ERRORE_CODICE") = "ACCOUNT_DISABLED"
					end if
				else
					Session("ERRORE") = ChooseByLanguage("Password errata.", _
														 "Wrong password.", _
														 "Falsches kennwort.", _
														 "Password incorrecte.", _
														 "Password incorrecta.")
					Session("ERRORE_CODICE") = "PASSWORD_ERROR"
				end if
			end if
			rs.close
			conn.close
			set rs = nothing
			set conn = nothing
		else
			Session("ERRORE") = ChooseByLanguage("Password non valida.", _
												 "Incorrect password.", _
												 "Falscher kennwort.", _
												 "Password incorrecte.", _
												 "Password incorrecta.")
			Session("ERRORE_CODICE") = "PASSWORD_ERROR"
		end if
	else
		Session("ERRORE") = ChooseByLanguage("Login non valido.", _
											 "Incorrect login.", _
											 "Falscher login.", _
											 "Login incorrecte.", _
											 "Login incorrecto.")
		Session("ERRORE_CODICE") = "LOGIN_ERROR"
	end if
	
	'se ci sono stati errori toglie il login
	CheckLogin = not(Session("ERRORE")<>"")
end function


'**************************************************************************************
'funzione che verifica che la procedura di registrazione abbia dato esito positivo
'conn:				connessione aperta su dbContents
'rs:				oggetto recordset creato
'login:				login da verificare
'Password:			password da verificare e controllare
'RetypedPassword: 	password per confronto
'**************************************************************************************
function ValidateLoginAndPassword(conn, rs, Login, OldLogin, Password, RetypedPassword)
	'controlla se password mancante
	if Password = "" then
		Session("Errore") = ChooseByLanguage("Password mancante!", "Missing Password!", "Unzul&auml;ssiger Password!", "Password inadmissible!", "Password inv&aacute;lido")
		ValidateLoginAndPassword = false
		Exit function
	end if
	if Login = "" then	
		Session("Errore") = ChooseByLanguage("Login mancante!", "Missing Login!", "Unzul&auml;ssiger Login!", "Login inadmissible!", "Login inv&aacute;lido")
		ValidateLoginAndPassword = false
		Exit function
	end if
	'controlla se la password e' uguale a quella di conferma
	if RetypedPassword = "" OR Password <> RetypedPassword then
		Session("Errore") = ChooseByLanguage("Password non ridigitata correttamente!", _
											 "Password not retyped correctly!", _
											 "Kennwort nicht richtig neu getippt!", _
											 "Password non retapé correctement!", _
											 "Contraseña no escrita de nuevo correctamente!")
		ValidateLoginAndPassword = false
		Exit function
	end if
	
	'esegue i controlli sul login solo se cambiato
	if uCase(Login) <> Ucase(OldLogin) then
		'controlla che il login contenga solo caratteri validi
		if not CheckChar(Login, LOGIN_VALID_CHARSET) then
			Session("Errore") = ChooseByLanguage("Login non valido: usare solo lettere, numeri o &quot;_&quot;.", _
												 "Login not valid: use only letters, numbers or &quot;_&quot;.", _
												 "Login unzulässig: benutzen Sie nur Buchstaben, Zahlen oder &quot;_&quot;", _
												 "Login inadmissible: employez seulement les lettres, nombres ou &quot;_&quot;", _
												 "Login inválido: utilice solamente las letras, números o &quot;_&quot;")
			ValidateLoginAndPassword = false
			Exit function
		end if
		
		'controlla che il login non sia gia' in uso da parte di un altro utente
		sql = "SELECT (COUNT(*)) AS N_USERS FROM tb_utenti WHERE ut_login LIKE '" & ParseSQL(login, adChar) & "'"
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		if rs("N_USERS") > 0 then
			Session("Errore") = ChooseByLanguage("Login gi&agrave; utilizzato da un altro utente, cambiare il login.", _
												 "Login already in use by another user, please change your login.", _
												 "Login bereits im Gebrauch durch einen anderen Benutzer, ändern Ihren login.", _
												 "Login déjà en service d'un autre utilisateur, changent votre login", _
												 "La login ya en uso de otro usuario, cambia tu conexión.")
			ValidateLoginAndPassword = false
			rs.close
			Exit function
		end if
		rs.close
	end if
	
	'se e' arrivato fino a qui la verifica ha dato esito positivo
	ValidateLoginAndPassword = (session("ERRORE") = "")
end function


'**************************************************************************************
'funzione che seleziona il valore da restituire sulla base della lingua corrente (letta da sessione)
'valueIT:		valore in lingua italiana
'valueEN:		valore in lingua inglese
'valueDE:		valore in lingua tedesca
'valueFR:		valore in lingua francese
'valueES:		valore in lingua spagnola
'**************************************************************************************
function ChooseByLanguage(valueIT, valueEN, valueDE, valueFR, valueES)
	ChooseByLanguage = ChooseValueByLanguage(Session("LINGUA"), valueIT, valueEN, valueDE, valueFR, valueES)
	if CString(ChooseByLanguage) = "" AND _
	   LCase(Session("LINGUA")) <> LINGUA_ITALIANO AND LCase(Session("LINGUA")) <> LINGUA_INGLESE then
		ChooseByLanguage = valueEN
	end if
	if CString(ChooseByLanguage) = "" AND LCase(Session("LINGUA")) <> LINGUA_ITALIANO then
		ChooseByLanguage = valueIT
	end if
end function


'**************************************************************************************
'versione concisa e performante di ChooseByLanguage (occhio ai prerequisiti)
'dizionario:	oggetto che contiene i valori
'nome:			i nomi devono essere tutti nome_lingua
'SE SI PASSA UN RECORDSET COME DIZIONARIO OCCHIO ALL'EOF
'**************************************************************************************
Function CBL(ByRef dizionario, ByVal nome)
    CBL = CBLL(dizionario, nome, LCase(session("LINGUA")))
End Function


'**************************************************************************************
'restituisce la pagina da passare al dynalay data la paginaSito in base alla lingua corrente
'**************************************************************************************
Function GetPage(conn, rs, paginaSito)
	GetPage = GetPageByLanguage(conn, rs, PaginaSito, Session("lingua"))
End Function


'.........................................................................................
'verifica se l'utente e' abilitato per visualizzare il plug-in, altrimenti lo manda alla 
'pagina di login
'permission_list:		lista di permessi nei quali almeno uno deve essere attivo
'page_login:			id della pagina a cui viene fatto il redirect se l'utente non &egrave; loggato
'.........................................................................................
sub CheckPermission(permission_list, page_login)
	dim permissions, p, abilited
	permissions = split(replace(permission_list, " ", ""), ";")
	abilited = false
	for each p in permissions
		if not abilited then
			'verifica se e' abilitato per ogni permesso
			if isNumeric(Session(p)) then
				abilited = cInteger(Session(p))>0
			else
				abilited = Session(p)<>""
			end if
		end if
	next
	if not abilited then
		response.redirect "dynalay.asp?PAGINA=" & page_login
	end if
end sub

'.........................................................................................
'scrive il link "indietro"
'.........................................................................................
sub WriteBackURL(href) 
	dim label
	label = ChooseByLanguage("indietro", "back", "zur&uuml;ck", "retour", "atr&aacute;s")%>
	<div class="back">
		<a href="<%= href %>" title="<%= label %>"><%= label %></a></a>
	</div>
<% end sub


'scrive la versione di default
sub WriteBackLink() 
	CALL WriteBackURL("javascript:history.back();")
end sub



'.........................................................................................
'scrive "nessun record trovato"
'.........................................................................................
sub WriteNoRecord(message)
	if message="" then
		message = ChooseByLanguage("Nessun record disponibile.", _
								   "No record available.", _
								   "Keine Record vorhanden.", _
								   "Aucuns record disponibles.", _
								   "Ningunos record disponibles.")
	end if%>
	<div class="NoRecords"><%= message %></div>
<%end sub


'.........................................................................................
'funzione che restituisce la classe di stili da applicare in base alla posizione del 
'record corrente all'interno del recordset
'.........................................................................................
function CssRecordPosition(CssPrefix, AbsolutePosition, RecordCount)
	select case AbsolutePosition
		case 1
			CssRecordPosition = CssPrefix + "first"
		case RecordCount 
			CssRecordPosition = CssPrefix + "last"
		case else
			CssRecordPosition = CssPrefix + "middle"
	end select
end function
%>
