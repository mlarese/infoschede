<%
'.................................................................................................
'.................................................................................................
'COSTANTI
'.................................................................................................
'.................................................................................................

'.................................................................................................
'.................................................................................................
'FUNZIONI E PROCEDURE
'.................................................................................................
'.................................................................................................


'***************************************************************************************************************************
'***************************************************************************************************************************
'FUNZIONI DI GESTIONE DEI PERMESSI
'***************************************************************************************************************************
'***************************************************************************************************************************


'...............................................................................
'.. Verifica l'autenticazione a seconda della condizione passata per parametro.
'..	se l'utente non è autenticato o l'area amministrativa è disattivata lo manda alla default.asp.
'.. inoltre verifica se il dominio corrente in cui si sta usando l'amministrazione è quello corretto.
'...............................................................................
function CheckAutentication(cond)
	dim url
	url = GetCurrentFullUrl()
	
	if cString(application("AMMINISTRAZIONE_DISABLED")) <> "" then
		response.redirect GetAmministrazionePath()
	elseif not cond then
		url = Server.UrlEncode(url)
		
		if response.buffer then
			response.clear
			response.redirect "default.asp?RETURN_URL=" & url
		else%>
			<script language="JavaScript">
				document.location = "default.asp?RETURN_URL=<%= url %>"
			</script>
		<%end if
	
	else
	
		'verifica se il dominio dell'amministrazione è quello valido
		if cString(Application("AMMINISTRAZIONE_SERVER_NAME"))<>"" AND _
		   request.servervariables("REQUEST_METHOD")<>"POST" then
			'il sistema ha un dominio di amministrazione pricipale: se non sono in quel dominio, faccio il redirect.
			if instr(1, url, Application("AMMINISTRAZIONE_SERVER_NAME"), vbTextCompare) < 1 AND _
			   instr(1, url, Application("SERVER_NAME"), vbTextCompare)>0 then
				'non sono nel dominio corretto: devo fare il redirect
				response.redirect replace(url, Application("SERVER_NAME"), Application("AMMINISTRAZIONE_SERVER_NAME"))
			end if
		end if
	
	end if
	
end function


'...............................................................................
'.. Reindirizza all'URL specificato dal parametro RETURN_URL. 
'..	Se non è specificato nessun URL reindirizza alla pagina di default. 
'...............................................................................
function AutenticatedRedirect(defaultPage)
	if request.Querystring("RETURN_URL")<>"" then
		Response.Redirect(request.Querystring("RETURN_URL"))
	else
		Response.Redirect(defaultPage)
	end if
end function

'...............................................................................
'.. ritorna se il cookie permette un accesso valido o meno, verificando anche
'.. l'applicazione che ha emesso il cookie sia quella corrente
'.. se richiesto verifica permessi temporanei
'...............................................................................
function isCookieValid(CheckSession)
	if instr(1, request.cookies(CookieName)("applicazione"), CookieApplication, vbTextCompare)>0 then
		if request.Cookies(CookieName)("permessi")<>"" then
			isCookieValid = true
		elseif CheckSession AND Session("PERMESSI_TEMPORANEI")<>"" then
			isCookieValid = true
		else
			isCookieValid = false
		end if
	elseif CheckSession AND Session("PERMESSI_TEMPORANEI")<>"" then
		'verifica permessi temporanei
		isCookieValid = true
	else
		isCookieValid = false
	end if
end function


sub ReturnToLogin()
	if instr(1, Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME"), "amministrazione", vbTextCompare) then
		response.redirect("../../amministrazione")
	else
		Response.redirect("../amministrazione")
	end if
end sub


'...............................................................................
'..	Imposta il cookie per il nextPassport con il permesso indicato
'..	permessi:		stringa di permessi da impostare nel cookie
'...............................................................................
function SetCookie(permessi)
	if NOT cBoolean(Application("DO_NOT_SAVE_COOKIES"), false) then
		'imposta il cookie
		Response.Cookies(CookieName)("permessi") = permessi
		response.cookies(CookieName)("applicazione") = CookieApplication
		
		'imposta la data di scadenza
		Response.Cookies(CookieName).Expires = now + 0.5
	end if
end function



'...............................................................................
'..	Imposta il cookie per la lingua dell'area amministrativa
'...............................................................................
function SetCookieLingua()
	if NOT cBoolean(Application("DO_NOT_SAVE_COOKIES"), false) then
		'imposta il cookie
		if Session("LINGUA") <> "" then
			Response.Cookies(CookieName)("lingua") = Session("LINGUA")
		else
			Response.Cookies(CookieName)("lingua") = LINGUA_ITALIANO
		end if
		
		'imposta la data di scadenza
		Response.Cookies(CookieName).Expires = now + 0.5
	end if
end function



'...............................................................................
'.. ritorna il valore attuale del cookie per l'accesso via nextPassport
'...............................................................................
function GetCookie(CheckSession)
	GetCookie = request.Cookies(CookieName)("permessi")
	'recupera gli eventuali permessi temporanei (se abilitati)
	if GetCookie = "" AND CheckSession then
		GetCookie = Session("PERMESSI_TEMPORANEI")
	end if
end function



'...............................................................................
'.. ritorna il valore attuale del cookie per recuperare la lingua attiva dell'amministrazione
'...............................................................................
function GetCookieLingua()
	GetCookieLingua = request.Cookies(CookieName)("lingua")
	'recupera la lingua attiva dell'amministrazione
	if GetCookieLingua = "" then
		GetCookieLingua = Session("LINGUA")
	end if
end function


'...............................................................................
'.. cancella il cookie di accesso
'...............................................................................
function CookieLogout()
	Response.Cookies(CookieName)("permessi")= ""
	response.cookies(CookieName)("applicazione") = ""
   	Response.Cookies(CookieName) = ""
end function


'...............................................................................
'.. Imposta il nome del cookie in uso in base ai parametri di applicazione
'...............................................................................
sub InitializeCookie(byref CookieName, byref CookieApplication)
	dim Cripto
	set Cripto = new CryptographyManager
	
	CookieName = "NEXTframework"
	CookieApplication = "NEXTAIM_" & Cripto.md5_of_string(request.ServerVariables("INSTANCE_META_PATH"))
	set Cripto = nothing
end sub


'inizzializza variabili di sessione salvando i dati della sessione precedente (se richiesto)
Sub PreserveInitSex(byval id, PreserveOldSession)
	
'	PreserveOldSession = true
	
	dim permessi, i, punto, myVar, user
	'ripulisce la sessione da tutte le variabili presenti (mantenendo i permessi temporanei)
	if not PreserveOldSession then
		CALL ResetSession()
	end if
	
	if not isCookieValid(true) then
		CALL ReturnToLogin()
	end if
	 
	 permessi = GetCookie(true)
	 Session("LINGUA") = GetCookieLingua()
	 
	 punto = instr(1, permessi, ";;")
	 user = left(permessi, punto-1)
	 for i = 1 to id
	 	 punto = instr(1, permessi, ";;")
		 permessi = right(permessi, len(permessi)-punto-1)
	 next
	 while left(permessi, 1) <> ";" and permessi <> ""
 	     punto = instr(1, permessi, ",")
		 myVar = left(permessi, punto-1)
		 Session(myVar) = user
		 permessi = right(permessi, len(permessi)-punto)
	 wend
	
	dim conn, sql, rs
	set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	conn.open Application("DATA_ConnectionString"),"",""
	
	'controlla situazione parametri applicazione
	if not PreserveOldSession then
		CALL Parametri.Check(conn, rs, id)
	end if
	
	sql = "SELECT id_admin FROM tb_admin WHERE admin_login LIKE '" & ParseSQL(user, adChar) & "'"
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.eof then
		'utente non trovato: cancella cookie e rimanda al login
		CALL CookieLogout()
		if not PreserveOldSession then
			CALL ResetSession()
		end if
        Session("ERRORE") = "Utente non riconosciuto."
		CALL ReturnToLogin()
	else
		
		Session.Timeout = 60

		if not PreserveOldSession then
			Session("LOGIN_4_LOG") = user
		    Session("ID_SITO") = id
			Session("ID_ADMIN") = rs("id_Admin")
		end if
		
		'registra accesso su db
		sql = "INSERT INTO log_admin (log_admin_id, log_sito_id, log_data, log_username, log_http_raw) " & _
			  " VALUES (" & rs("id_Admin") & ", " & cIntero(id) & ", " & SQL_Now(conn) & ", '" & ParseSQL(user, adChar) & "', '"&ParseSql(GetRawHttp(), adChar)&"')"
		CALL conn.execute(sql,0, adExecuteNoRecords)
		
		rs.close
		
		'imposta titolo applicazione
        if not PreserveOldSession then
    		sql = "SELECT sito_nome FROM tb_siti WHERE id_sito=" & cIntero(id)
	    	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		    Session("NOME_APPLICAZIONE") = rs("sito_nome") & " - " & user
    		rs.close
        end if
		
		'imposta valori parametri dell'applicazione
		CALL Parametri.LoadAllParams(conn, rs, id)
		
	end if
	
	conn.close
	set conn = nothing
	set rs = nothing
end sub


'***************************************************************************************************************************
'***************************************************************************************************************************
'FUNZIONI GENERICHE
'***************************************************************************************************************************
'***************************************************************************************************************************


'.................................................................................................
'..			Restituisce l'elenco delle rubriche 
'..			conn:			connessione al database aperta
'..			rs:				oggetto recordset chiuso
'.................................................................................................
function GetList_Rubriche(conn, rs)
	dim sql, list
	if Session("COM_ADMIN")<>"" or _
	   (cIntero(Application("Site_id"))<>NEXTCOM AND cIntero(Session("id_sito"))<>NEXTCOM)then
		'amministratore: vede tutte le rubriche
		list = "SELECT id_rubrica FROM tb_rubriche"
	else
		sql = " SELECT (id_dellaRubrica) AS Id_rubrica FROM " &_
			  " (tb_gruppi INNER JOIN tb_rel_dipgruppi ON tb_gruppi.id_Gruppo = tb_rel_dipgruppi.id_gruppo) " &_
			  " INNER JOIN tb_rel_gruppirubriche ON tb_gruppi.id_Gruppo = tb_rel_gruppirubriche.id_Gruppo_assegnato " &_
			  " WHERE tb_rel_dipgruppi.id_impiegato=" & cInteger(Session("ID_ADMIN")) &_
			  " GROUP BY id_dellaRubrica"
	
	list = GetValueList(conn, rs, sql)
	end if
	GetList_Rubriche = list
end function


'.................................................................................................
'..                     Restituisce l'url dell'immagine in input
'.................................................................................................
Function GetUrlImage(image, AZ_ID)
	GetUrlImage = GetUrlFile(FILE_TYPE_IMAGE, image, AZ_ID)
End Function



'.................................................................................................
'..                     Restituisce l'url del file in input
'.................................................................................................
Function GetUrlFile(folder, image, AZ_ID)
	if cIntero(AZ_ID) = 0 then
		if cIntero(Session("AZ_ID")) = 0 then
			AZ_ID= Application("AZ_ID")
		else
			AZ_ID= Session("AZ_ID")
		end if
	end if
	
	GetUrlFile = "http://" & Application("IMAGE_SERVER") & "/" & AZ_ID & "/" & folder & "/" & image
End Function


'.................................................................................................
'.. 	procedura che scrive il link alla pagina indicata
'.................................................................................................
Sub WritePageLink(connWeb, rs, PaginaSitoId, lingua)
	dim Url, Name, sql
	Url = GetPageSiteUrl(connWeb, PaginaSitoId, lingua)
	if Url <> "" then
		
		sql = "SELECT (" & SQL_PaginaSitoNome(connWeb, "nome_ps_IT") & ") AS NOME FROM tb_paginesito WHERE id_paginesito=" & cIntero(PaginaSitoId)
		Name = cString(GetValueList(ConnWeb, rs, sql))
		
		if Name<>"" then %>
			<a href="<%= Url %>" title="apri il link &ldquo;<%= Url %>&rdquo; in una nuova finestra" target="_blank">
				<%= Name %> - <%= lcase(GetNomeLingua(lingua)) %>
			</a>
		<% end if
		
	end if
end sub


'.................................................................................................
'..                     Creazione drop-down  da lista file con pulsante per visualizzazione
'..						UTILIZZA FUNZIONE dropDown in TOOLS.asp
'.................................................................................................
sub dropDown_WithViewer(label, url, form, width, conn, sql, field_id, field_value, input, selected, obbligatorio)
	dim a_label, i, a_url%>
	<script language="JavaScript" type="text/javascript">
		function <%= form + input %>_View(url){
			var index = <%= form %>.<%= input %>.selectedIndex;
			<% if not obbligatorio then  %>
				if (index!=0)
			<% end if %>
					window.open(url + <%= form %>.<%= input %>.options[index].value , "<%= input %>_View",
									"left=50,top=50,width=<%= IIF(InStr(url, "dynalay.asp"), "800", "420") %>,height=350,scrollbars=yes,statusbar=yes,menubars=no,resizable=yes");
		}
	</script>
	<%
	if width<>"" then
		width = " style=""width:" & width & ";"" "
	else
		width = " style=""width:100%;"" "
	end if
	%>
	<table border="0" cellspacing="0" cellpadding="0" align="left" <%=width%>>
		<tr>
			<td style="padding-top:2px; padding-bottom:1px;">
				<% CALL dropDown(conn, sql, field_id, field_value, input, selected, obbligatorio, width, LINGUA_ITALIANO) %>
			</td>
			<% a_label = split(label, ";")
			a_url = split(url, ";")
			for i=lbound(a_label) to ubound(a_label) %>
				<td style="padding-top:1px;">
					<a href="javascript:void(0);" class="button_input" id="<%= input %>_link" style="padding-top:3px; padding-bottom:3px;" onclick="<%= form + input %>_View('<%= a_url(i) %>')" title="<%= a_label(i) %>" <%= ACTIVE_STATUS %>>
						<%= a_label(i) %>
					</a>
				</td>
			<% next %>
		</tr>
	</table>
<%end sub


'.................................................................................................
'..	Procedura che scrive il dropdown per elencare le pagine del nextweb
'.................................................................................................
sub DropDownPages(nextWeb_Conn, form, width, web_id, InputName, InputValue, obbligatorio, IsPageForEmail)
	CALL DropDownPagesAdvanced(nextWeb_Conn, form, width, web_id, InputName, InputValue, obbligatorio, IsPageForEmail, true, 0)
end sub

sub DropDownPagesAdvanced(nextWeb_Conn, form, width, web_id, InputName, InputValue, obbligatorio, IsPageForEmail, SelezionePaginaPubblica, paginaEsclusaID)

	dim ConnCreated, PreviewPath, sql, NextWebVersion, NomePs
    
	if not IsObjectCreated(nextWeb_Conn) then
		set nextWeb_Conn = Server.CreateObject("ADODB.Connection")
		nextWeb_Conn.open Application("l_conn_ConnectionString"),"",""
		ConnCreated = true
	end if
	
    'recupera versione next-web
    NextWebVersion = cInteger(GetNextWebCurrentVersion(NULL, NULL))
    
    if NextWebVersion < 5 then
        NomePs = "nome_ps_it"
    else
        NomePs = SQL_PaginaSitoNome(nextWeb_Conn, "nome_ps_it")
    end if
	
	if IsPageForEmail then
		
		sql = SQL_Pagine(nextWeb_Conn, NextWebVersion, web_id, "PAGINA_ID", "PAGINA_NOME", _
						 IIF(CIntero(paginaEsclusaID) = 0, "", " AND p.id_page <> "& paginaEsclusaID), _
						 SelezionePaginaPubblica)
		
		'link di visualizzazione effettivo
		PreviewPath = GetLibraryPath() + "site/PageView.asp?PAGINA="
	else
		'recupera query paginesito
		if cInteger(web_id)>0 then
			'pagine di un solo sito
			sql = " SELECT (id_pagineSito) AS PAGINA_ID, (" + NomePs + ") AS PAGINA_NOME " + _
				  " FROM tb_pagineSito"& _
				  " WHERE id_web=" & cIntero(web_id) & _
				  " AND id_pagineSito <> "& paginaEsclusaID & _
				  " ORDER BY " & IIF(NextWebVersion < 5, "nome_ps_it", "nome_ps_it, nome_ps_interno")
		else
			'pagine di tutti i siti
			sql = " SELECT (id_pagineSito) AS PAGINA_ID, ( " & SQL_IfIsNull(nextWeb_Conn, "nome_webs", "''") & SQL_concat(nextWeb_Conn) + "' - '" + SQL_concat(nextWeb_Conn) + NomePs + ") AS PAGINA_NOME " + _
				  " FROM tb_pagineSito INNER JOIN tb_webs ON tb_paginesito.id_web = tb_webs.id_webs " + _
				  " WHERE id_pagineSito <> "& paginaEsclusaID & _
				  " ORDER BY nome_webs, " & IIF(NextWebVersion < 5, "nome_ps_it", "nome_ps_it, nome_ps_interno")
		end if
		
		'link di visualizzazione anteprima in area amministrativa
		PreviewPath = GetAmministrazionePath() + GetNextWebDirectory(NextWebVersion) + "/SitoPagineView.asp?ID="
	end if
	
	CALL dropDown_WithViewer( ChooseValueByAllLanguages(Session("LINGUA"), "VISUALIZZA", "VIEW", "", "", "", "", "", ""), PreviewPath, _
							  form, width, nextWeb_Conn, sql, _
							  "PAGINA_ID", "PAGINA_NOME", InputName, InputValue, obbligatorio)
	
	if ConnCreated then
		nextWeb_Conn.close
		set nextWeb_Conn = nothing
	end if
end sub


'.................................................................................................
'.. 	Creazione di un imput con funzione di navigazione e scelta dei file in una nuova finestra
'.................................................................................................
sub WriteFilePicker_Input(FILEMAN_AZ_ID, file_type, form_name, field_name, field_value, field_style, obbligatorio)
    CALL WriteFileSystemPicker_Input(FILEMAN_AZ_ID, FILE_SYSTEM_FILE, file_type, "", form_name, field_name, field_value, field_style, false, obbligatorio)
end sub


'.................................................................................................
'.. 	Creazione di un imput con funzione di navigazione e scelta di un oggetto del file manager in una nuova finestra
'..     FILEMAN_AZ_ID       Id del sito del quale recuperare gli oggetti
'..     OBJECT_TYPE         Tipo di oggetto del file manager. valori : FILE_SYSTEM_FILE o FILE_SYSTEM_DIRECTORY
'..		directory			filtro sul tipo di oggetto next-web (directory in cui risiedono i files) da recuperare (images, objects, ecc..)
'..     file_type           eventuale filtro sul tipo di file ( lista di estensioni valide ) valido solo se il tipo di oggetto da selezionare e' FILE_SYSTEM_FILE
'..     form_name           nome del form constestuale agli input che verranno creati
'..     field_name          nome dell'input HTML che conterra' il nome dell'oggetto
'..     field_value         valore precaricato dell'input
'..     field_style         stile HTML dell'input
'..		showViewer			se true viene visualizzato il pulsante di visualizzazione del file
'..     obbligatorio        indica se la selezione e' obbligatoria e quindi il campo non puo' essere svuotato
'.................................................................................................
sub WriteFileSystemPicker_Input(FILEMAN_AZ_ID, OBJECT_TYPE, directory, file_type, form_name, field_name, field_value, field_style, showViewer, obbligatorio)
	dim FilteredPath, BaseType, LockPath, LibraryPath
	if instr(1, directory, "/", vbTextCompare)>0 then
		'se il tipo di file contiene un percorso blocca anche la navigazione nel percorso inserito
		FilteredPath = "&lock=" & directory
		BaseType = left(directory, instr(1, directory, "/", vbTextCompare)-1)
		LockPath = right(directory, len(directory) - instr(1, directory, "/", vbTextCompare)+1)
	else
		'navigazione per tipo di file
		FilteredPath = "&filter=" & directory
		BaseType = directory
		LockPath = ""
	end if

	%>
	<script language="JavaScript" type="text/javascript">
		function <%= field_name %>_select(obj_field){
			
			if (!document.getElementById('<%= field_name %>').disabled) {
				//in caso il file contenga gia un percorso fa aprire il filemanger direttamente nella directory
				var BasePath = obj_field.value;
				
				if (BasePath.indexOf("/") >= 0){
					var PathParam = BasePath.substring(0,BasePath.lastIndexOf("/"));
					var re = /\//;
					PathParam = PathParam.replace(re, "\\");
				}
				else
					var PathParam = "<%= LockPath %>";
				if (PathParam.substring(0,1) != "/" && PathParam.substring(0,1) != "\\")
					PathParam = "/" + PathParam;
	
				<% if automaticPathWriteFileSystemPicker_Input <> "" then 'automaticPathWriteFileSystemPicker_Input è dichiarata su Tools.asp
					dim fso, dir, segment, segmentPath, additionalPath
					set fso = Server.CreateObject("Scripting.FileSystemObject")						
					
					additionalPath = automaticPathWriteFileSystemPicker_Input
					additionalPath = Replace(additionalPath, "/", "\")
					additionalPath = Replace(additionalPath, "\\", "\")
								
					segmentPath = Split(additionalPath, "\")
					dir = Application("IMAGE_PATH") & Application("AZ_ID") & "\images\"
					for each segment in segmentPath
						dir = dir & "\" & segment
						dir = Replace(dir, "\\", "\")
						if NOT fso.FolderExists(dir) then 'se non c'è la directory, la creo
							fso.CreateFolder(dir)
						end if
					next
					set fso = nothing
					
					%>
					if (PathParam == "" || PathParam == "/")
						PathParam = '<%= Replace("\" & automaticPathWriteFileSystemPicker_Input, "\", "\\") %>'
				<% end if %>
				
				var url = "<%= GetLibraryPath() %>filemanager.asp?STANDALONE=1&OBJECT_TYPE=<%= OBJECT_TYPE %>&FILEMAN_AZ_ID=<%= FILEMAN_AZ_ID %>&form_name=<%= form_name %>&file_type_filter=<%= file_type %>&field_name=<%= field_name %>" +
						  "<%= FilteredPath %>&selected=" + obj_field.value + "&F=\\<%= BaseType %>" + PathParam;
				OpenPositionedScrollWindow(url, obj_field.name, window.screenLeft, window.screenTop, 770, 430, false);
			}
		}
		
		<% if OBJECT_TYPE = FILE_SYSTEM_FILE then %>
			function <%= form_name %>_<%= field_name %>_vedi_immagine(){
				var fileField = document.getElementById("<%= field_name %>");
				if (fileField.value != ""){
					var fileUrl = "http://<%= Application("IMAGE_SERVER") %>/<%= FILEMAN_AZ_ID %>/<%= directory %>/" + fileField.value.toLowerCase();
					var imagesExtensions = "<%= lcase(EXTENSION_IMAGES) %>";
					if ( imagesExtensions.indexOf(fileUrl.substr(fileUrl.lastIndexOf(".") + 1, 3)) >= 0 )
						opensmartimage(fileUrl);		//immagine
					else
						OpenWindow(fileUrl, '', '');	//altro file
				}
			}
		<% end if %>
	</script>
	<table cellspacing="0" cellpadding="0" class="PickerComponent">
		<tr>
			<td>
				<input type="text" id="<%= field_name %>" name="<%= field_name %>" value="<%= field_value %>" ondblclick="<%= field_name %>_select(this)" style="letter-spacing:1px;<%= field_style %>" <%= IIF(OBJECT_TYPE = FILE_SYSTEM_FILE, "", "readonly") %>>
			</td>
			<td>
				<a href="javascript:void(0);" id="link_scegli_<%= field_name %>" title="<% if OBJECT_TYPE = FILE_SYSTEM_FILE then %>seleziona il nome del file dall'elenco dei files caricati<% else %>seleziona il nome della directory dall'elenco delle directory presenti.<% end if %>" <%= ACTIVE_STATUS %> 
				   class="button_input" onclick="<%= field_name %>_select(<%= form_name %>.<%= field_name %>)">
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "SCEGLI", "CHOOSE", "", "", "", "", "", "")%>
				</a>
			</td>
			<% if OBJECT_TYPE = FILE_SYSTEM_FILE AND showViewer then %>
				<td>
					<a href="javascript:void(0);" id="link_vedi_<%= field_name %>" class="button_input" onclick="<%= form_name %>_<%= field_name %>_vedi_immagine()" <%= ACTIVE_STATUS %>>
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "VEDI", "VIEW", "", "", "", "", "", "")%>
					</a>
				</td>
			<% end if
			if not obbligatorio then %>
				<td>
					<a href="javascript:void(0);" id="link_reset_<%= field_name %>" class="button_input" onclick="if (!document.getElementById('<%= field_name %>').disabled) <%= form_name %>.<%= field_name %>.value='';" title="cancella la selezione" <%= ACTIVE_STATUS %>>
						RESET
					</a>
				</td>
			<% end if 
			if obbligatorio then %>
				<td>&nbsp;(*)</td>
			<% end if %>
		</tr>
	</table>
<%end sub


'.................................................................................................
'..                     procedura che prepara la tabella per la selezione via checkbox 
'..						dei valori di una relazione 
'..						conn		connessione al database aperta
'..						rs			oggetto recordset chiuso
'..                     sql		    Query di lettura della lista
'..						Cells4Row	Numero di colonne
'..						fieldID		campo id della tabella da relazionare
'..						FieldName	campo nome della tabella da relazionare
'..						FieldKey	campo id della relazione
'..						FormFieldName	nome della lista di checkbox sul form
'.................................................................................................
SUB Write_Relations_Checker(conn, rs, sql, Cells4Row, FieldID, FieldName, FieldKey, FormFieldName)
	dim FirstColumn_RowCount, Column_Offset, CurrentRow,i, offset, checked
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	
	if (rs.recordcount MOD Cells4Row) > 0 then
		FirstColumn_RowCount = (rs.recordcount \ Cells4Row) + 1
	else
		FirstColumn_RowCount = rs.recordcount \ Cells4Row
	end if
	Column_Offset = FirstColumn_RowCount %>
	<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
		<% while not rs.eof and  rs.AbsolutePosition <= FirstColumn_RowCount 
			CurrentRow = rs.AbsolutePosition%>
			<tr>
				<%for i=0 to Cells4Row-1
					if i=0 then 
						offset = 0
					else
						offset = Column_Offset
					end if
					
					if (rs.Absoluteposition + offset) <= rs.recordcount then
						rs.AbsolutePosition = rs.AbsolutePosition + offset
						
						if Request.ServerVariables("REQUEST_METHOD")="POST" then
							if inStr(","&Replace(request(FormFieldName), " ", "")&",", ","&rs(FieldID)&",") > 0 then
								checked = true
							else
								checked = false
							end if
						else
							if rs(FieldKey)>0 then
								checked = true
							else
								checked = false
							end if
						end if
						%>
						<td class="content<%=IIF(checked, " content_checked", "")%>" width="<%= 100 \ Cells4Row %>%">
							<table border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td valign="top"><input type="checkbox" name="<%= FormFieldName %>" value="<%= rs(FieldID) %>" <%= IIF(checked, "checked class=""checked"" ", " class=""checkbox"" ")%> ></td>
									<td class="content<%=IIF(checked, " content_checked", "")%>"><%= rs(FieldName) %></td>
								</tr>
							</table>
						</td>
					<% else %>
						<td class="content" width="<%= 100 \ Cells4Row %>%">&nbsp;</td>
					<%end if%>
				<%next%>
			</tr>
			<%rs.AbsolutePosition = CurrentRow
			rs.moveNext
		wend %>
		
		</table>
	<%rs.close
end sub


'.................................................................................................
'esegue una lista di istruzioni sql separate da ;
'.................................................................................................
sub ExecuteMultipleSql(conn, sql, logs)
	dim i, sql_list
	sql = replace(sql, ";;", "^^^")
	sql_list = split(sql, ";")
	for i= lbound(sql_list) to ubound(sql_list)
		if trim(replace(replace(sql_list(i), vbTab, ""), vbCrLF, ""))<>"" then
			sql_list(i) = replace(sql_list(i), "^^^", ";")
			'esegue aggiornamento
			'if logs then
				response.write "<!--" & sql_list(i) & "-->" & vbcrlf & vbcrlf 
			'end if	
			CALL conn.execute(sql_list(i), 0, adExecuteNoRecords)
		end if
	next
end sub


'.................................................................................................
'..				Spostamento al record richiesto in request("goto")
'..				conn		connessione al database aperta
'..				rs			oggetto recordset chiuso
'..             sql		    Query di lettura della lista
'..				FieldID		Campo indice per lo spostamento
'..				Page		Pagina di redirezione
'.................................................................................................
sub GotoRecord(conn, rs, sql, FieldID, Page)
	if sql<>"" AND request("goto")<>"" then
		dim ID
		ID = cIntero(request("ID"))
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdtext
		rs.Find FieldID & "=" & ID
		if not rs.eof then
			Select Case uCase(request("goto"))
				case "NEXT"
					rs.MoveNext
					if not rs.eof then
						ID = rs(FieldID)
					else
						Session("ERRORE") = "Record successivo non trovato: il record corrente &egrave; l'ultimo dell'elenco."
					end if
				case "PREVIOUS"
					rs.MovePrevious
					if not rs.bof then
						ID = rs(FieldID)
					else
						Session("ERRORE") = "Record precedente non trovato: il record corrente &egrave; il primo dell'elenco."
					end if
			end select
		else
			Session("ERRORE") = "Spostamento al record richiesto non eseguito. Record non trovato."
		end if
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
		if instr(1, page, "?", vbTextCompare)>0 then
			response.redirect Page & "&ID=" & ID
		else
			response.redirect Page & "?ID=" & ID
		end if
	end if
end sub



'.................................................................................................
'..				Pulizia cartella temporanea
'.................................................................................................
sub ClearTempDir(FSO)
	dim FSO_dir, FSO_file
	if fso.FolderExists(Application("IMAGE_PATH") & "\temp") then
		Set FSO_dir = fsO.GetFolder(Application("IMAGE_PATH") & "\temp")
		for each FSO_file in FSO_dir.Files
			if DateDiff("h", cDate(FSO_file.DateLastModified), Now)>2 then
				'cancella il file se non utilizzato da piu' di due ore.
				FSO.DeleteFile FSO_file.path, true
			end if
		next
		set FSO_dir = nothing
	else
		CALL fsO.CreateFolder(Application("IMAGE_PATH") & "\temp")
	end if
end sub



'.................................................................................................
'..			Cambia nome o crea la cartella temporanea per l'utente amministratore
'..			NewLogin:		Nuovo login al quale corrispondera' il nome della cartella
'..			OldLogin:		Eventuale login precedente alla variazione: cartella da rinominare
'.................................................................................................
sub CreateTemporaryDir(NewLogin, OldLogin)
	dim fso, path, base_path

	
	NewLogin = cString(NewLogin)
	OldLogin = cString(OldLogin)
	
	if uCase(NewLogin) <> uCase(OldLogin) AND NewLogin<>""  then
		
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		base_path = Application("IMAGE_PATH") & "temp\docs"

		'verifica path base, se manca lo crea
		if not fso.FolderExists(base_path) then
			fso.CreateFolder(base_path)
		end if
		
		'verifica se esiste il login precedente, eventualmente lo rinomina
		path = base_path & "\" & OldLogin
		if OldLogin <> "" AND fso.FolderExists(path) then
			'verifica se esiste il nuovo login
			if not fso.FolderExists(base_path & "\" & NewLogin) then
				'il login precedente viene rinominato
				dim folder
				set folder = fso.GetFolder(path)
				folder.name =  NewLogin
				set folder = nothing
			else
				'il nuovo login esiste gi&agrave;: copia la cartella del vecchio login sopra la cartella del nuovo
				CALL fso.CopyFolder(path, base_path & "\" & NewLogin, true)
				'cacella la vecchia directory
				CALL fso.DeleteFolder(path, true)
			end if
		else
			path = base_path & "\" & NewLogin
			if fso.FolderExists(path) then
				fso.DeleteFolder path, true
			end if
			fso.CreateFolder(path)
		end if
		
		set fso = nothing
	end if
end sub


'.................................................................................................
'..			Cancellazione contatti non collegati ad alcuna rubrica (next-com)
'..			conn		connessione aperta a database
'.................................................................................................
sub ClearNextCom(conn)
	dim sql
	sql = "DELETE FROM tb_Indirizzario WHERE " &_
		  " IDElencoIndirizzi NOT IN (SELECT id_indirizzo FROM rel_rub_ind) AND " &_
		  " IDElencoIndirizzi NOT IN (SELECT ut_NextCom_ID FROM tb_utenti) AND " & _
		  " ("& SQL_IsNULL(conn, "cntRel") &" OR cntRel = 0) AND " & _
		  " ("& SQL_IsNULL(conn, "SyncroApplication") &" OR SyncroApplication = 0) " & _
		  " AND ISNULL(LockedByApplication, 0) = 0 "
	CALL Conn.execute(sql, 0, adExecuteNorecords)
end sub


'.................................................................................................
'..          Scrive il recordset in formato XML su response
'.................................................................................................
sub Export_XML(rs, ClearLF)
	dim xml_stream
	set xml_stream = server.createobject("ADODB.Stream")
	
	'imposta proprieta' risposta
	'response.ContentType = "text/xml"
	response.ContentType = "text/xml"
	
	'salva proprieta' su recordset
	rs.Save xml_stream,adPersistXML
	
	
	xml_stream.Position=0
	
	'legge XML da stream
	response.clear
	response.write  "<?xml version=""1.0"" encoding=""UTF-8""?>" &_
		  			 replace(xml_stream.ReadText(xml_stream.size), vbLf, IIF(ClearLF, "", vbLf))
	
	'salva xml su file.
	'dim fso, file
	'set fso = server.createobject("Scripting.FileSystemObject")
	'set file = fso.CreateTextFile("d:\frameworks\agenziarallo.com\upload\loadxml.txt", true)
	'xml_stream.Position=0
	'file.write replace(xml_stream.ReadText(xml_stream.size), vbLf, IIF(ClearLF, "", vbLf))
	'file.close

	
	xml_stream.Close
	set xml_stream = nothing
end sub




'.................................................................................................
'..			genera il dropdown di scelta della lingua sulla base delle lingue installate
'..			conn:				connessione al database dbContent aperta
'..			rs:					oggetto recorrdset, se NULL lo crea e lo distrugge internamente
'..			SelectName:			nome dell'oggetto html creato
'..			Selected			valore selezionato
'..			UseItalianLabel		Indica se usare le etichette dei valori in italiano o le traduzioni in lingua
'..			style				eventuale stile HTML dell'input
'.................................................................................................
sub DropLingue(conn, rs, SelectName, Selected, UseItalianLabel, AllowNull, style)
	CALL DropLingueEx(conn, rs, SelectName, Selected, UseItalianLabel, AllowNull, style, false)
end sub

'.................................................................................................
'..			genera il dropdown di scelta della lingua sulla base delle lingue installate
'..			conn:				connessione al database dbContent aperta
'..			rs:					oggetto recorrdset, se NULL lo crea e lo distrugge internamente
'..			SelectName:			nome dell'oggetto html creato
'..			Selected			valore selezionato
'..			UseItalianLabel		Indica se usare le etichette dei valori in italiano o le traduzioni in lingua
'..			style				eventuale stile HTML dell'input
'..			disabled			se true il drop down risulterà disabilitato
'.................................................................................................
sub DropLingueEx(conn, rs, SelectName, Selected, UseItalianLabel, AllowNull, style, disabled)
	dim sql, rsCreated, NameField, ActiveLanguages
	'sceglie campo per generazione select
	if UseItalianLabel then
		NameField = "lingua_nome_it"
	else
		NameField = "lingua_nome"
	end if
	'calcola lista lingue installate
	ActiveLanguages = Join(Application("LINGUE"))
	
	'completa stile
	if style<>"" then
		if instr(1, style, "class", vbTextCompare)=0 AND instr(1, style, "style", vbTextCompare)=0 then
			style = " style=""" & style & """ "
		end if
	end if
	
	'se necessario crea recordset per lingue
	rsCreated = not IsObjectCreated(rs)
	if rsCreated then
		Set rs = Server.CreateObject("ADODB.RecordSet")
	end if
	
	'se nessuna lingua selezionata : sceglie italiano
	if cString(Selected) = "" AND NOT AllowNull then
		Selected = LINGUA_ITALIANO
	end if
	
	'ciclo per generazione select solo con lingue installate
	sql = "SELECT * FROM tb_cnt_lingue ORDER BY " & NameField
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	%>
	<select <%=IIF(disabled, "disabled", "")%> name="<%= SelectName %>" id="<%= SelectName %>" <%= Style %>>
		<% if AllowNull then %>	
			<option value="">Scegli..</option>
		<%end if
		while not rs.eof
			if instr(1, ActiveLanguages, rs("lingua_codice"), vbTextCompare)>0 then%>
				<option value="<%= rs("lingua_codice") %>" <%= IIF(instr(1, Selected, rs("lingua_codice"), vbTextCompare)>0, "selected", "")  %>>
					<%= rs(NameField) %>
				</option>
			<%end if
			rs.MoveNext
		wend%>
	</select>
	<%rs.close
	
	'se recordset creato internamente lo distrugge
	if rsCreated then
		set rs = nothing
	end if
end sub


'.................................................................................................
'..			simula ClassSalva per input con prefisso ext<TIPO>_
'..			extID_nome:		nome della chiave esterna
'..			extID_value:	valore della chiave esterna
'.................................................................................................
Function SalvaCampiEsterni(conn, rs, sql, id_nome, id_value, extID_nome, extID_value)
	salvaCampiEsterni = SalvaCampiEsterniUltra(conn, rs, sql, id_nome, id_value, extID_nome, extID_value, "", NULL, request.form, "ext")
End Function

Function SalvaCampiEsterniChk(conn, rs, sql, id_nome, id_value, extID_nome, extID_value, chkList)
    SalvaCampiEsterniChk = SalvaCampiEsterniUltra(conn, rs, sql, id_nome, id_value, extID_nome, extID_value, chkList, NULL, request.form, "ext")
end function

Function SalvaCampiEsterniAdvanced(conn, rs, sql, id_nome, id_value, extID_nome, extID_value, chkList, ParamsPrefix)
	SalvaCampiEsterniAdvanced = SalvaCampiEsterniUltra(conn, rs, sql, id_nome, id_value, extID_nome, extID_value, chkList, ParamsPrefix, request.form, "ext")
End Function

Function SalvaCampiEsterniUltra(conn, rs, sql, id_nome, id_value, extID_nome, extID_value, chkList, ParamsPrefix, collection, ByRef collectionPrefix)
    dim campo_nome, campo, chk, Inserimento, prefixLength
	id_value = CIntero(id_value)
	if id_value > 0 then
		sql = sql + IIF(instr(1, sql, "WHERE", vbTextCompare)>0, " AND ", " WHERE ") & _
			  id_nome &"="& cIntero(id_value)
	end if 
	rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	if rs.eof AND id_value > 0 then		'gestione relazione 1-1
		rs.AddNew
		rs(id_nome) = id_value
        Inserimento = true
	elseif id_value = 0 then
		rs.AddNew
        Inserimento = true
    else
        Inserimento = false
	end if
	
	prefixLength = Len(collectionPrefix) + 2
	collectionPrefix = LCase(collectionPrefix)
	for each campo in collection
		if len(campo) > prefixLength then
			campo_nome = right(campo, len(campo) - prefixLength)
		else
			campo_nome = ""
		end if
		'response.write "<br>CAMPO:"& campo &";NOME:"& campo_nome &";"
		if FieldExists(rs, campo_nome) then
			SELECT CASE lcase(left(campo, prefixLength))
				CASE collectionPrefix + "t_"
					rs(campo_nome) = collection(campo)
				CASE collectionPrefix + "n_"
                    rs(campo_nome) = ConvertForSave_Number(collection(campo), 0)
				CASE collectionPrefix + "v_"
                    rs(campo_nome) = ConvertForSave_Number(collection(campo), NULL)
				CASE collectionPrefix + "d_"
                    rs(campo_nome) = ConvertForSave_Date(collection(campo))
			END SELECT
			'response.write "VAL:"& rs(campo_nome) &";"
		end if
	next
	'response.end
	
	for each chk in Split(chkList, ";")
		if Left(chk, prefixLength) = collectionPrefix + "C_" then
			campo_nome = right(chk, len(chk) - prefixLength)
		else
			campo_nome = ""
		end if
		if FieldExists(rs, campo_nome) then
			'rs(campo_nome) = CIntero(collection(chk))
			if request(chk)<>"" then
				rs(campo_nome) = True
			else
				rs(campo_nome) = False
			end if
		end if
	next

	if id_value = 0 AND extID_nome <> "" then
		rs(extID_nome) = CInteger(extID_value)
	end if

    if not IsNull(ParamsPrefix) then
        CALL SetUpdateParamsRS(rs, ParamsPrefix, Inserimento)
    end if
	rs.Update
	SalvaCampiEsterniUltra = rs(id_nome)
	rs.Close
End Function


'.................................................................................................
'funzione che resetta la ricerca salvata nelle variabili di sessione
'VarNameBegin			Caratteri di inizio di tutte le variabili di sessione utilizzate per la ricerca
'.................................................................................................
sub SearchSession_Reset(VarNameBegin)
	dim VarName, LenVarNameBegin
	VarNameBegin = lcase(VarNameBegin)
	LenVarNameBegin = len(VarNameBegin)
	for each VarName in Session.Contents
		if lcase(left(VarName, LenVarNameBegin)) = VarNameBegin then
			'variabile di sessione di ricerca
			Session(VarName) = ""
		end if
	next
end sub


'.................................................................................................
'funzione che imposta le variabili di sessione per la ricerca alle variabili della richiesta
'VarNameBegin			Caratteri di inizio di tutte le variabili di sessione utilizzate per la ricerca
'.................................................................................................
sub SearchSession_Set(VarNameBegin)
	dim VarName, collection
	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		set collection = request.form
	else
		set collection = request.querystring
	end if
	for each VarName in collection
		if lcase(left(VarName, 7)) = "search_" then
			'variabile di sessione di ricerca
			Session(replace(VarName, "search_", VarNameBegin, 1, -1, vbTexTCompare)) = collection(varName)
		end if
	next
end sub


'.................................................................................................
'imposta il backgroud dell'intestazione della ricerca se i campi sono diversi da ""
'campi			campi in session da controllare (separati da ";")
'.................................................................................................
Function Search_Bg(campi)
dim campo
	Search_Bg = ""
	for each campo in split(campi, ";")
		if session(campo) <> "" then
			Search_Bg = " style=""color:white; background-color: #E99281;"" "
			exit for
		end if
	next
End Function


'.................................................................................................
'codifica la i parametri di accesso per permettere il login automatico ed il ritorno alla pagina 
'da cui viene lanciata la richiesta
'.................................................................................................
function LoginString_Encode(conn, rs, destination, site)
	if instr(1, request.ServerVariables("HTTPS"), "on", vbTextCompare) _
	   OR Application("SECURE_SERVER_NAME") = "" then
		LoginString_Encode = GetAmministrazionePath + destination
	else
		dim sql, Login, Password, Credentials, NextReferer
		Login = Session("LOGIN_4_LOG")
		sql = "SELECT admin_password FROM tb_admin WHERE admin_login LIKE '" & ParseSql(login, adChar) & "'"
		Password = GetValueList(Conn, rs, sql)
		
		Credentials = Ucase(login) & ";" & Ucase(Password) & ";" & destination & ";" & site
		NextReferer = "http://" & request.ServerVariables("SERVER_NAME") & request.ServerVariables("SCRIPT_NAME") & "?" & request.ServerVariables("QUERY_STRING")
		'Response.write Credentials
		'codifica credenziali di accesso
		dim i, Key, char, EncodedCredentials
		Key = LoginString_URL_KEY(NextReferer)
		EncodedCredentials = ""
		for i=1 to len(credentials)
			char = cString(Asc(Mid(credentials ,i, 1)) + Key)
			char = string(3 - len(char), "0") & char
			EncodedCredentials = EncodedCredentials & char
		next
		EncodedCredentials = Server.UrlEncode(EncodedCredentials)
		
		LoginString_Encode = "https://" & Application("SECURE_SERVER_NAME") & "/" & _
							 "amministrazione/default.asp?EXECUTE=ACCEDI&INFO=" & EncodedCredentials
	end if
end function


'.................................................................................................
'decodifica i dati per l'accesso su stringa codificata
'.................................................................................................
function LoginString_Decode(credentials, byref login, byref password, byref destination, byref site)
	dim i, char, chars, key, DecodedString
	key = LoginString_URL_KEY(request.ServerVariables("HTTP_REFERER"))
	DecodedString = ""
	while credentials <> ""
		chars = left(Credentials, 3)
		credentials = right(Credentials, len(Credentials) - 3)
		char = Chr(cInteger(chars) - Key)
		DecodedString = DecodedString & char
	wend
	for i=1 to len(credentials)
		char = Mid(credentials ,i, 1)
	next

	DecodedString = split(DecodedString, ";")
	login = DecodedString(0)
	password = DecodedString(1)
	destination = DecodedString(2)
	site = DecodedString(3)
end function


'.................................................................................................
'decodifica i dati per l'accesso su stringa codificata proveniente da .NET
'.................................................................................................
function LoginPassword_DecodeFromNET(byref login,byref password, credentials)

	dim chars, DecodedString
	DecodedString = ""
	while credentials <> ""		
		chars = left(credentials, 3)
		credentials = right(credentials, len(credentials) - 4)		
		DecodedString = DecodedString & Chr(cInteger(chars))
	wend

	DecodedString = split(DecodedString, ";")
	login = DecodedString(0)
	password = DecodedString(1)
	
end function

'.................................................................................................
'decodifica i dati per l'accesso su stringa codificata
'.................................................................................................
function LoginString_URL_KEY(URL)
	dim i, char, checksum
	if url <> "" then
		checksum = 0
		for i=1 to len(URL)
			checksum = checksum + Asc(Mid(URL ,i, 1))
		next
		checksum = checksum \ len(url)
		
		LoginString_URL_KEY = checksum
	end if
end function



'.................................................................................................
'funzione per la rimozione di una directory, forzando la rimozione se e' vuota o meno
'ATTENZIONE: restituisce true anche se non la trova!!
'.................................................................................................
function FolderRemove(fso, path, RemoveOnlyIfEmpty)
	dim Folder, ToRemove
	if fso.FolderExists(path) then
		set Folder = fso.GetFolder(path)
		if RemoveOnlyIfEmpty then
			if cInteger(Folder.Files.Count)<=0 AND cInteger(Folder.SubFolders.Count)<=0 then
				'directory vuota
				ToRemove = true
			else
				'directory non vuota
				ToRemove = false
			end if
		else
			'rimuovi sempre e comunque
			ToRemove = true
		end if
		if ToRemove then
			fso.DeleteFolder(path)
			FolderRemove = true
		else
			FolderRemove = false
		end if
	else
		'directory non trovata
		FolderRemove = true
	end if
end function


'.................................................................................................
'procedura per la cancellazione di un file. Se impostato a true il RecursiveMode esegue la cancellazione
'del file anche nelle eventuali sottodirectory
'.................................................................................................
sub FileRemove(fso, Path, FileToDelete, RecursiveMode)
	'verifica esistenza path
	if fso.FolderExists(Path) then
		'verifica esistenza file
		if fso.FileExists(Path + "\" + FileToDelete) then
			'cancella file
			fso.DeleteFile(Path + "\" + FileToDelete)
		end if
		'se in modalita' ricorsiva richiama la funzione per le sottodirectory
		if RecursiveMode then
			dim Folder, SubFolder
			'recupera sottodirectory
			set Folder = fso.GetFolder(path)
			for each SubFolder in Folder.SubFolders
				'richiama funzione per tutte le sottodirectory
				CALL FileRemove(fso, SubFolder.Path, FileToDelete, true)
			next
		end if
	end if
end sub



'.................................................................................................
'funzione che scrive la parte di interfaccia per la pubblicazione di un file
'.................................................................................................
sub FileLink(FileName)
	dim FileExtension
	FileExtension = File_Extension( FileName )
	if FileExtension<>"" then %>
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td class="content_image">
					<a title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "visualizza il file " & FileName & " in una nuova finestra", "view " & FileName & " in a new window", "", "", "", "", "", "")%>" onclick="<%= File_OpenInNewWindow(FileName) %>" <%= ACTIVE_STATUS %> href="javascript:void(0)">
		   				<img src="<%= GetAmministrazionePath() %>grafica/filemanager/<%= File_Icon( FileExtension ) %>" alt="visualizza il file '<%= FileName %>' in una nuova finestra" border="0">
					</a>
				</td>
				<td class="content">
					<a title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "visualizza il file " & FileName & " in una nuova finestra", "view " & FileName & " in a new window", "", "", "", "", "", "")%>" onclick="<%= File_OpenInNewWindow(FileName) %>" <%= ACTIVE_STATUS %> href="javascript:void(0)">
		   				<%= File_Name(FileName) %>
					</a>
				</td>
			</tr>
		</table>
	<% end if
end sub


'.................................................................................................
'funzione che verifica se un file e' utilizzato o meno
'.................................................................................................
function FileCanBeRemoved(conn, nextWebConn, rs, fileType, webId, fileName)
	
	FileCanBeRemoved = true
	if fileType = FILE_TYPE_IMAGE OR _
	   fileType = FILE_TYPE_FLASH OR _
	   fileType = FILE_TYPE_OBJECTS OR _
	   fileType = FILE_TYPE_TEXT then
   
		dim sql, FileName1, FileName2, NextWebVersion, ConnCreated
		
		if IsNull(NextWebConn) then
			ConnCreated = true
			set NextWebConn = Server.CreateObject("ADODB.Connection")
			NextWebConn.open Application("l_conn_ConnectionString"),"",""
		end if
		
		NextWebVersion = GetNextWebCurrentVersion(conn, rs)
		
		'calcola varianti del nome del fine : con e senza / all'inizio (mantiene compatibilita' con editor precedente)
		FileName1 = fileName
		FileName1 = replace(FileName1, "/", "", 1, 1, vbTextCompare)
		FileName1 = ParseSQL(lCase(FileName1), adChar)			'nome file neutro
		FileName2 = "/" & FileName1								'nome file con barra anteposta
	
		if fileType = FILE_TYPE_OBJECTS then
			'cancellazione immagine in cartella oggetti
			sql = "SELECT COUNT(*) FROM tb_objects WHERE img_objects LIKE '" & ParseSQL(FileName1, adChar) & "' OR img_objects LIKE '" & ParseSQL(FileName2, adChar) & "' "
			if NextWebVersion < 4 then
				if FileName1 = "obj_vuoto.jpg" OR FileName2 = "/obj_vuoto.jpg" then
					FileCanBeRemoved = false
				end if
			end if
		elseif fileType = FILE_TYPE_FLASH then
			'cancellazione di un file flash
			sql = "SELECT COUNT(*) FROM tb_layers WHERE " & _
				  " id_tipo=" & LAYER_FLASH & _
				  " AND (nome LIKE '" & ParseSQL(FileName1, adChar) & "' OR nome LIKE '" & ParseSQL(FileName2, adChar) & "') " & _
				  " AND id_pag IN (SELECT id_page from tb_pages where id_webs =" & cIntero(webId) & ")"
			'response.write sql
			'response.end
		else
			'cancellazione di una immagine
			if NextWebVersion = 5 then
				sql = "SELECT COUNT(*) FROM tb_layers INNER JOIN tb_pages ON tb_layers.id_pag=tb_pages.id_page WHERE " & _
					  " ((tb_pages.sfondoImmagine LIKE '" & ParseSQL(FileName1, adChar) & "' OR tb_pages.sfondoImmagine LIKE '" & ParseSQL(FileName2, adChar) & "') OR " & _
					  " (id_tipo=" & LAYER_IMAGE & " AND (tb_layers.nome LIKE '" & ParseSQL(FileName1, adChar) & "' OR tb_layers.nome LIKE '" & ParseSQL(FileName2, adChar) & "'))) " & _
					  " AND id_webs=" & cIntero(webId)
			else
				sql = "SELECT COUNT(*) FROM tb_layers INNER JOIN tb_pages ON tb_layers.id_pag=tb_pages.id_page WHERE " & _
					  " ((tb_pages.sfondo LIKE '" & ParseSQL(FileName1, adChar) & "' OR tb_pages.sfondo LIKE '" & ParseSQL(FileName2, adChar) & "') OR " & _
					  " (id_tipo=" & LAYER_IMAGE & " AND (tb_layers.nome LIKE '" & ParseSQL(FileName1, adChar) & "' OR tb_layers.nome LIKE '" & ParseSQL(FileName2, adChar) & "'))) " & _
					  " AND id_webs=" & cIntero(webId)
			end if
			if NextWebVersion < 4 then
				if FileName1 = "vuoto.jpg" OR FileName2 = "/vuoto.jpg" then
					'file di sistema per la creazione di nuove immagini
					FileCanBeRemoved = false
				end if
			end if
		end if
		 
		if FileCanBeRemoved then
			if GetValueList(NextWebConn, rs, sql) > 0 then
				'file utilizzato nella costruzione delle pagine
				FileCanBeRemoved = false
			end if
		end if
		
		if ConnCreated then
			NextWebConn.close
			set NextWebConn = nothing
		end if
	else
		FileCanBeRemoved = true
	end if
	
end function



'.................................................................................................
'procedura per la cancellazione delle cartelle vuote anche nelle eventuali sottodirectory
'.................................................................................................
sub RemoveEmptyFolders(fso, BasePath)
	Dim BaseFolder, SubFolders
	if fso.FolderExists(BasePath) then
		set BaseFolder = fso.GetFolder(BasePath)
		if cInteger(BaseFolder.SubFolders.Count)>0 then
			'verifica eventuale rimozione delle cartelle figlie
			for each SubFolder in BaseFolder.SubFolders
				'richiama funzione per tutte le sottodirectory
				CALL RemoveEmptyFolders(fso, SubFolder.Path)
			next
		end if
		'ricontrolla la directory per eventuali cancellazioni
		set BaseFolder = fso.GetFolder(BasePath)
		if cInteger(BaseFolder.Files.Count)<=0 AND cInteger(BaseFolder.SubFolders.Count)<=0 then
			'directory vuota: la cancella!
			CALL fso.DeleteFolder(BasePath)
		end if
	end if
end sub


'.................................................................................................
'procedura che restituisce il nome "corto" dell'applicazione
'.................................................................................................
function GetApplicationShortName(NomeApplicazione)
	GetApplicationShortName = uCase(Trim(left(NomeApplicazione, (instr(1, NomeApplicazione, "[", vbTextCompare)-1))))
end function


'.................................................................................................
'procedura che scrive il link di logout per le applicazioni
'.................................................................................................
sub WriteLogoutLink(NomeApplicazione)
	dim label, title
	
	NomeApplicazione = cString(NomeApplicazione)
	if instr(1, NomeApplicazione, "[", vbTextCompare)>0 then
		label = ChooseValueByAllLanguages(Session("LINGUA"), "Esci dal ", "Quit ", "", "", "", "", "", "") + GetApplicationShortName(NomeApplicazione)
		title = ChooseValueByAllLanguages(Session("LINGUA"), "Esci da &laquo;", "Quit &laquo;", "", "", "", "", "", "") & NomeApplicazione & ChooseValueByAllLanguages(Session("LINGUA"), "&raquo; e torna all'area di login", "&raquo; and return to the login area", "", "", "", "", "", "")
	else
		label = ChooseValueByAllLanguages(Session("LINGUA"), "Esci dall'applicazione", "Quit application", "", "", "", "", "", "")
		title = ChooseValueByAllLanguages(Session("LINGUA"), "Esci dall'applicazione e torna all'area di login", "Quit application and return to the login area", "", "", "", "", "", "")
	end if %>
	<a href="<%= GetLibraryPath() %>logout.asp" class="menu" title="<%= title %>" <%= ACTIVE_STATUS %>><%= label %></a>
<% end sub


 
'.................................................................................................
'funzione che ritorna la lista di colonne di una tabella (utilizzata per generare funzioni di 
'copia dei dati
'.................................................................................................
function TableFieldList(conn, rs, TableName, ExcludeFields)
	dim RsCreated, sql, field
	
    if not IsObjectCreated(rs) then
		set rs = Server.CreateObject("ADODB.Recordset")
		RsCreated = true
	end if
	
	sql = "SELECT TOP 1 * FROM " & TableName
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	TableFieldList = ""
	for each field in rs.fields
		if instr(1, " " + ExcludeFields + " ", " " + Trim(Field.name) + " ", vbTextCompare)=0 then
			'campo non escluso
			TableFieldList = TableFieldList + IIF(TableFieldList<>"", ", " , "") + Field.name
		end if
	next
	
	rs.close
	if RsCreated then
		set rs = nothing
	end if
	
end function


'.................................................................................................
'..			resetta la sessione salvando i permessi temporanei
'.................................................................................................
function ResetSession()
	'dim permessi
	'permessi = Session("PERMESSI_TEMPORANEI") 
	'Session.Contents.RemoveAll
	'Session("PERMESSI_TEMPORANEI") = permessi
end function


'.................................................................................................
'..			funzione che formatta l'id in modo coerente con la lista
'.................................................................................................
function IdList_ID(Id)
    Id = cString(Id)
    if Id<>"" then
        IdList_ID = "(" & ID & ")"
    else
        IdList_ID = ""
    end if
end function


'.................................................................................................
'..			funzione che aggiunge ad una lista di numeri in formato (n)(n) un numero
'.................................................................................................
function IdList_ADD(List, Id)
	List = Trim(cString(List))
	Id = cString(Id)
	if Id<>"" AND _
	   instr(1, List, IdList_ID(Id), vbTextCompare)<1 then
		List = List & IdList_ID(Id)
	end if
	IdList_ADD = IIF(List="", NULL, List)
end function


'.................................................................................................
'..			funzione che rimuove una lista di numeri in formato (n)(n) un numero
'.................................................................................................
function IdList_REMOVE(List, Id)
	List = Trim(cString(List))
	Id = cString(Id)
	if List <> "" then
		if instr(1, List, "(", vbTextCompare)<1 then
			if cInteger(List) = cInteger(Id) then
				List = ""
			end if
		else
			List = replace(List, IdList_ID(Id), "")
		end if
	end if
	IdList_REMOVE = IIF(List="", NULL, List)
end function


'.................................................................................................
'..			Restituisce l'sql da concatenare ad una query UPDATE
'.................................................................................................
Function SetUpdateParamsSQL(conn, prefix, inInserimento)
	
	SetUpdateParamsSQL = prefix &"modData = "& SQL_Now(conn) &", " & _
						 prefix &"modAdmin_id = "& CIntero(Session("ID_ADMIN"))

	if inInserimento then
		SetUpdateParamsSQL = SetUpdateParamsSQL + ", " & _
							 prefix &"insData = "& SQL_Now(conn) &", " & _
							 prefix &"insAdmin_id = "& CIntero(Session("ID_ADMIN")) & ","
	end if
	
End Function


'.................................................................................................
'..			Inserisce i dati di inserimento e modifica nel record specificato
'.................................................................................................
Sub UpdateParams(conn, tb, prefix, campoID, ID, inInserimento)
	dim sql, rs
	
    sql = " UPDATE "& tb & " SET "
	
	if inInserimento then
		sql = sql & prefix &"insData = "& SQL_Now(conn) & ", " & _
					prefix &"insAdmin_id = "& CIntero(Session("ID_ADMIN")) &","
	end if
	
	sql = sql & prefix &"modData = "& SQL_Now(conn) &", " & _
			    prefix &"modAdmin_id = "& CIntero(Session("ID_ADMIN")) & _
				" WHERE "& campoID &" = "
	
	if instr(1, ID, "SELECT", vbTextCompare) > 0 then
		sql = sql & ID
	else
		sql = sql & cIntero(ID)
	end if
	CALL conn.execute(sql, 0, adExecuteNoRecords)
End Sub


'.................................................................................................
'..			Inserisce i dati di inserimento e modifica nel record specificato
'.................................................................................................
Sub SetUpdateParamsRS(rs, prefix, inInserimento)
	dim tempo
	tempo = Now

	if inInserimento OR IsNull(rs(prefix &"insData")) then
		rs(prefix &"insData") = tempo
		rs(prefix &"insAdmin_id") = CIntero(Session("ID_ADMIN"))
	end if
	rs(prefix &"modData") = tempo
	rs(prefix &"modAdmin_id") = CIntero(Session("ID_ADMIN"))
End Sub


'.................................................................................................
'..			Visualizza cognome e nome dell'amministratore dato l'id
'.................................................................................................
Function GetAdminName(conn, ID)
	dim sql, connCreated
    if not IsObject(conn) then
        connCreated = true
        set conn = Server.CreateObject("ADODB.Connection")
	    conn.open Application("DATA_ConnectionString")
    else
        connCreated = false
    end if
    
	sql = " SELECT admin_cognome "& SQL_Concat(conn) &"' '"& SQL_Concat(conn) &" admin_nome"& _
		  " FROM tb_admin WHERE id_admin="& cIntero(ID)
	GetAdminName = GetValueList(conn, NULL, sql)
    
    if connCreated then
        conn.close
        set conn = nothing
    end if
End Function


'.................................................................................................
'..			Resituisce l'id dell'utente dato il login
'.................................................................................................
Function GetAdminId(conn, login)
	dim sql, connCreated
    if not IsObject(conn) then
        connCreated = true
        set conn = Server.CreateObject("ADODB.Connection")
	    conn.open Application("DATA_ConnectionString")
    else
        connCreated = false
    end if
    
	sql = " SELECT top 1 id_admin  FROM tb_admin WHERE admin_login LIKE '" & ParseSQL(login, adChar) & "'"
	GetAdminId = cIntero(GetValueList(conn, NULL, sql))
    
    if connCreated then
        conn.close
        set conn = nothing
    end if
End Function


'.................................................................................................
'..			Visualizza i dati di modifica e inserimento del record
'			rs:		recordset (o obj_contatto) aperto sul record contenente i dati di modifica
'			prefix:	prefisso da anteporre ai nomi dei campi standard
'.................................................................................................
Sub Form_DatiModifica(conn, rs, prefix)
	CALL Form_DatiModifica_EX(conn, rs, prefix, ChooseValueByAllLanguages(Session("LINGUA"), "DATI DEL RECORD", "RECORD INFORMATION", "", "", "", "", "", ""), "")
End Sub


'.................................................................................................
'..			Visualizza i dati di modifica e inserimento del record
'			rs:		recordset (o obj_contatto) aperto sul record contenente i dati di modifica
'.................................................................................................
Sub Form_DatiModifica_EX(conn, rs, prefix, label, ThClass) %>
	<tr><th <%= IIF(ThClass<>"", " class=""" & ThClass & """ ", "") %> colspan="4"><%= label %></th></tr>
	<tr>
		<td class="label" style="white-space: nowrap;"><%= ChooseValueByAllLanguages(Session("LINGUA"), "creato da:", "created by:", "", "", "", "", "", "")%></td>
		<td class="content"><%= GetAdminName(conn, cInteger(rs(prefix &"insAdmin_id"))) %></td>
		<td class="label" style="white-space: nowrap;"><%= ChooseValueByAllLanguages(Session("LINGUA"), "creato il:", "creation date:", "", "", "", "", "", "")%></td>
		<td class="content"><%= DateTimeLingua(rs(prefix &"insData"),Session("LINGUA")) %></td>
	</tr>
	<tr>
		<td class="label" style="white-space: nowrap;"><%= ChooseValueByAllLanguages(Session("LINGUA"), "ultima modifica da:", "last modify by:", "", "", "", "", "", "")%></td>
		<td class="content"><%= GetAdminName(conn, cInteger(rs(prefix &"modAdmin_id"))) %></td>
		<td class="label" style="white-space: nowrap;"><%= ChooseValueByAllLanguages(Session("LINGUA"), "ultima modifica il:", "last modify date:", "", "", "", "", "", "")%></td>
		<td class="content"><%= DateTimeLingua(rs(prefix &"modData"),Session("LINGUA")) %></td>
	</tr>
<%
End Sub


'.................................................................................................
'FUNZIONI SPECIALI PER LA RICERCA
'.................................................................................................

'funzione che restituisce la parte di sql da aggiungere alla clausola where per ricercare il contatto per nominativo full-text
function SQL_FullTextSearch_Contatto_Nominativo(conn, StrToSearch)
    SQL_FullTextSearch_Contatto_Nominativo = _
        SQL_FullTextSearch(StrToSearch, "NomeOrganizzazioneElencoIndirizzi;" + _
                                        "CognomeElencoIndirizzi;" + _
                                        "NomeElencoIndirizzi;" + _
                                        "SecondoNomeElencoIndirizzi;" + _
										SQL_ConcatFields(conn, IIF(DB_Type(conn) = DB_SQL, "ISNULL(CognomeElencoIndirizzi,'');ISNULL(NomeElencoIndirizzi,'');ISNULL(SecondoNomeElencoIndirizzi,'')", + _
										"CognomeElencoIndirizzi;NomeElencoIndirizzi;SecondoNomeElencoIndirizzi")))


end function								



'funzione che restituisce la parte di sql da aggiungere alla clausola where per ricercare il contatto per indirizzo full text
function SQL_FullTextSearch_Contatto_Indirizzo(conn, StrToSearch)
    SQL_FullTextSearch_Contatto_Indirizzo = _
        SQL_FullTextSearch(StrToSearch, "IndirizzoElencoIndirizzi;" + _
                                        "LocalitaElencoIndirizzi;" + _
										"StatoProvElencoIndirizzi;" + _
                                        "CittaElencoIndirizzi; " + _
                                        "CAPElencoIndirizzi; " + _
                                        "ZonaElencoIndirizzi; " + _
                                        SQL_ConcatFields(conn, IIF(DB_Type(conn) = DB_SQL, "ISNULL(IndirizzoElencoIndirizzi,'');ISNULL(LocalitaElencoIndirizzi,'');ISNULL(CittaElencoIndirizzi,'');ISNULL(CAPElencoIndirizzi,'');ISNULL(StatoProvElencoIndirizzi,'')", + _
										"IndirizzoElencoIndirizzi;LocalitaElencoIndirizzi;CittaElencoIndirizzi;CAPElencoIndirizzi;StatoProvElencoIndirizzi")))
end function



'.................................................................................................
'..     funzione che scrive la chiusura della testata(menu) uguale per tutti gli applicativi
'.................................................................................................
sub WriteChiusuraIntestazione(BarraApplicativo)

    if cString(BarraApplicativo)="" OR Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA")<>"" then
        BarraApplicativo = "barra_intestazione.jpg"
    end if
    BarraApplicativo = Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA") & BarraApplicativo %>
   <tr>
		<td style="background-image: url(<%= GetAmministrazionePath() + "grafica/" + BarraApplicativo %>);" class="barra_menu">
            
            <% if not Application("DISABLE_NEXTAIM_LINKS") AND cString(Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA"))="" then %>
                <a href="http://www.combinario.com" target="_blank" title="Supporto clienti su www.combinario.com" <%= ACTIVE_STATUS %>>
                    <img src="<%= GetAmministrazionePath() %>grafica/transp.gif" width="75" height="27" border="0" title="supporto clienti su www.combinario.com" alt="supporto clienti su www.combinario.com">
                </a>
            <% else %>
                <img src="<%= GetAmministrazionePath() %>grafica/transp.gif" width="75" height="27" border="0">
            <% end if %>
			<br>
			<%=DataEstesa(Date(), SESSION("LINGUA"))%>
		</td>
  	</tr>
  	<tr><td style="font-size:1px;">&nbsp;</td></tr>
<% end sub


'.................................................................................................
'..     funzione che scrive la chiusura della testata(menu) uguale per tutti gli applicativi con nuova grafica Combinario
'.................................................................................................
sub WriteChiusuraIntestazioneComb(BarraApplicativo, Sezione)
    if cString(BarraApplicativo)="" OR Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA")<>"" then
        BarraApplicativo = "combinario/sfondo-comb-grigio.jpg"
    end if
    BarraApplicativo = Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA") & BarraApplicativo %>
   <tr>
		<td class="barra_sezione" style="background-image: url(<%= GetAmministrazionePath() + "grafica/" + BarraApplicativo %>);">
			<span class="sezione"><%=Sezione%></span>
			<a href="http://www.combinario.com/" target="_blank">
				<img class="logo" src="<%= GetAmministrazionePath() + "grafica/combinario/logo-combinario-94x30.png" %>" alt="Combinario Logo" />
			</a>
		</td>
  	</tr>
<% end sub

'.................................................................................................
'..     procedura che scrive l'input per la selezione dei contatti
'..     conn:                   connessione aperta a database da dove vengono prelevati i nomi dei contatti gia' selezionati
'..     rs:                     recordset creato
'..     ContattisqlCondition    condizione di filtro per i contatti da caricare per la selezione
'..     RubrichesqlCondition    condizione di filtro per le rubriche da caricare per la selezione
'..     formName                nome del form della pagina chiamante
'..     fieldName               nome dell'input html che contiene il valore
'..     fieldValue              contatto o elenco di contatti selezionati (id)
'..     listType                tipologia di elenco dei contatti secondo le opzioni (possono essere combinate l'una con l'altra):
'..                                 ""                  elenco di contatti semplice
'..									"CNTREL"			elenco con i contatti interni
'..                                 "EMAIL"             elenco di contatti con email visualizzata
'..                                 "EMAILMANDATORY"    elenco di contatti con email visualizzata e selezionabili solo con email valida
'..                                 "LOGIN"             elenco di contatti con eventuale login
'..                                 "LOGINMANDATORY"    elenco di utenti
'..									"LOGINID"			l'ID passato e' quello dell'utente
'..     multipleSelection       se true genera un input textarea per la selezione di piu' contatti
'..     obbligatorio            indica se viene gestito il pulsante reset e visualizza (*)
'..     disabled                indica se l'input e' disabilitato o meno
'..     OnChangeSelectionMethod permette la gesione dell'evento di selezione indicando le operazioni da eseguire
'..                             secondo le opzioni:     ""              se vuoto non esegue nulla (Selezione semplice)
'..                                                     "REDIRECT"      esegue il redirect della pagina chiamante alla pagina stesasa con il parametro IDCNT impostato all'id del contatto
'..                                                     "SUBMIT"        esegue il submit del form che contiene l'input
'..                                                     "<nome metodo>" permette di eseguire la funzione <nome metodo> alla selezione del contatto
'.................................................................................................
sub WriteContactPicker_Input(conn, rs, ContattisqlCondition, RubrichesqlCondition, formName, fieldName, fieldValue, listType, multipleSelection, obbligatorio, disabled, OnChangeSelectionMethod)
	CALL WriteContactPicker_Input_Option(conn, rs, ContattisqlCondition, RubrichesqlCondition, formName, fieldName, fieldValue, listType, multipleSelection, obbligatorio, disabled, OnChangeSelectionMethod, false)
end sub

sub WriteContactPicker_Input_Option(conn, rs, ContattisqlCondition, RubrichesqlCondition, formName, fieldName, fieldValue, listType, multipleSelection, obbligatorio, disabled, OnChangeSelectionMethod, pulsantiOrizzontali)
    dim sql, AuxId, AuxValue, AuxHREF, AuxHTML
    'verifica oggetti database
    dim rsCreated, connCreated

    if not IsObjectCreated(rs) then
		rsCreated = true
		set rs = server.createobject("adodb.recordset")
	else
		rsCreated = false
	end if
    if not IsObjectCreated(conn) then
		connCreated = true
        set conn = Server.CreateObject("ADODB.Connection")
        conn.open Application("DATA_ConnectionString"),"",""
	else
		connCreated = false
	end if
    
    'imposta variabile per passaggio query
    session("CONDIZIONE_SELEZIONE_CONTATTI_" & formName & "_" & fieldName) = cString(ContattisqlCondition)
    session("CONDIZIONE_SELEZIONE_RUBRICHE_" & formName & "_" & fieldName) = cString(RubrichesqlCondition)
    
    'recupera valori selezionati visualizzati
	dim typeLoginId, ID
	typeLoginId = InStr(UCase(listType), "LOGINID")
    if cString(fieldValue)<>"" then
        sql = " SELECT * FROM tb_indirizzario i"& _
			  " LEFT JOIN tb_utenti u ON i.idElencoIndirizzi = u.ut_nextCom_id"& _
			  " WHERE "
		if typeLoginId then
			sql = sql &"ut_id "
		else
			sql = sql &"IdElencoIndirizzi "
		end if
		sql = sql & IIF(multipleSelection, "IN (" & replace(replace(fieldValue, "; ", ","), ";", "") & ") ", " = " & fieldValue) & _
        		    " ORDER BY ModoRegistra "

		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext

        if multipleSelection then
            while not rs.eof
				if typeLoginId then
					ID = rs("ut_id")
				else
					ID = rs("IdElencoIndirizzi")
				end if
                AuxId = AuxId & " " & ID & ";"
                AuxValue = AuxValue & " " & JSReplacerEncode(ContactFullName(rs)) & ";"
                rs.movenext
            wend
        elseif not rs.eof then
			if typeLoginId then
				ID = rs("ut_id")
			else
				ID = rs("IdElencoIndirizzi")
			end if
            AuxId = ID
            'AuxValue = ContactName(rs)
			AuxValue = ContactFullName(rs)
        end if
        rs.close
    else
        AuxValue = ""
    end if
    %>
	<script language="JavaScript" type="text/javascript">
		function <%= fieldName %>_seleziona_onClick(sender){
            if (<%= fieldName %>_is_input_enabled()){
                var url = "<%= GetLibraryPath() %>SelezioneContatti.asp?FieldName=<%= fieldName %>&FormName=<%= formName %>&ListType=<%= listType %>&MultipleSelection=<%= IIF(multipleSelection, 1, 0) %>";
                OpenAutoPositionedScrollWindow(url, "<%= formName & "_" & fieldName %>", 700, 400, true);
            }
		}
        
        function <%= fieldName %>_is_input_enabled(){
            return !<%= formName %>.<%= fieldName %>.disabled;
        }
        
        function <%= fieldName %>_input_onChangeSelection(){
             <% if OnChangeSelectionMethod <>"" then 
                Select case uCase(OnChangeSelectionMethod)
                    case "SUBMIT" 
                        'esegue submit del form dopo la selezione 
						%>
                        <%= formName %>.submit();
                    <% case "REDIRECT"
                        'esegue redirect dopo la selezione del contatto 
						%>
                        document.location = "<%= GetCurrentUrl() %>?IDCNT=" + <%= formName %>.<%= fieldName %>.value;
                    <% case else 
                        'esegue il metodo passato come parametro
						%>
                        <%= OnChangeSelectionMethod %>;
                <% end select
            end if %>
        }
		
        function <%= fieldName %>_is_selected_contatto(id){
            var value = <%= formName %>.<%= fieldName %>.value;
            return (value.indexOf(<%= fieldName %>_id_format(id)) != -1);
        }
        
        function <%= fieldName %>_id_format(id){
            <% if multipleSelection then %>
                return ' ' + id + ';';
            <% else %>
                return id;
            <% end if %>
        }
        
        function <%= fieldName %>_nome_format(nome){
            <% if multipleSelection then %>
                return ' ' + nome + ';';
            <% else %>
                return nome;
            <% end if %>
        }
                
        function <%= fieldName %>_selezione_contatto(sender, id, nome){
            <% if multipleSelection then 
                'selezione di un elenco di contatti
				%>
                if (!<%= fieldName %>_is_selected_contatto(id)){
                    //selezione del contatto
			        <%= formName %>.view_<%= fieldName %>.value += <%= fieldName %>_nome_format(nome);
                    <%= formName %>.<%= fieldName %>.value += <%= fieldName %>_id_format(id);
                }
                else {
                    //deselezione del contatto
                    var re = eval('/' + <%= fieldName %>_nome_format(nome) + '/g');
                    <%= formName %>.view_<%= fieldName %>.value = <%= formName %>.view_<%= fieldName %>.value.replace(re, '');
			        re = eval('/' + <%= fieldName %>_id_format(id) + '/g');
			        <%= formName %>.<%= fieldName %>.value = <%= formName %>.<%= fieldName %>.value.replace(re, '');
                }
            <% else
                'selezione diretta di un contatto
				%>
			    <%= formName %>.view_<%= fieldName %>.value = nome;
                <%= formName %>.<%= fieldName %>.value = id;
                <%= fieldName %>_input_onChangeSelection();
            <% end if %>
        }
        
		function <%= fieldName %>_reset_onClick(){
            if (<%= fieldName %>_is_input_enabled()){
	    		<%= formName %>.view_<%= fieldName %>.value = '';
                <%= formName %>.<%= fieldName %>.value = '';
                <%= fieldName %>_input_onChangeSelection();
            }
		}
	</script>
	<input type="hidden" name="<%= fieldName %>" id="<%= fieldName %>" value="<%= AuxId %>" <%= disable(disabled) %>>
	<table cellspacing="0" cellpadding="0" style="width:100%">
		<tr>
            <% AuxHREF = " onclick=""" & fieldName & "_seleziona_onClick(this)"" title=" & ChooseValueByAllLanguages(Session("LINGUA"), """click per aprire la finestra per la selezione dei contatti.""", """click to open the window for choosing the contacts.""", "", "", "", "", "", "")
            AuxHTML = " style=""width:100%;"" name=""view_" & fieldName & """ id=""view_" & fieldName & """ " & DisableClass(disabled, "")%>
			<% if cBoolean(pulsantiOrizzontali, false) then %>
				<td style="width:83%;">
			<% else %>
				<td style="width:95%;">
			<% end if %>
                <% if multipleSelection then %>
    				<textarea READONLY rows="3" 
					<% if cBoolean(pulsantiOrizzontali, false) then %> style="width:100%; height:30px;" <% end if %>
					<%= AuxHTML %> <%= AuxHREF %>><%= AuxValue %> </textarea>
                <% else %>
                    <input READONLY type="text" <%= AuxHTML %> <%= AuxHREF %> value="<%= AuxValue %>">
                <% end if %>
			</td>
			<% if cBoolean(pulsantiOrizzontali, false) then %>
				<td style="width:120px; vertical-align: top; padding-top: 1px;">
			<% else %>
				<td style="width:60px;">
			<% end if %>
                <a href="javascript:void(0);" 
				<% if cBoolean(pulsantiOrizzontali, false) then %> style="line-height:30px; float:left;" <% end if %>
				<%= ACTIVE_STATUS %> <%= AuxHREF %> id="link_scegli_<%= fieldName %>"
                    <% if multipleSelection then %> 
                        <%= DisableClass(disabled, "button_textarea") %> <% if obbligatorio then %> style="line-height:440%"<%end if 
                    else %>
                        <%= DisableClass(disabled, "button_input") %>
                    <% end if %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "SCEGLI", "CHOOSE", "", "", "", "", "", "")%>
				</a>
                <% if not obbligatorio then 
                    if not multipleSelection then%>
                        </td>
                        <td>
                    <% end if %>
                    <a href="javascript:void(0);" 
					<% if cBoolean(pulsantiOrizzontali, false) then %> style="line-height:30px; float:left;" <% end if %>
					<%= ACTIVE_STATUS %> onclick="<%= fieldName %>_reset_onClick()" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "cancella la selezione", "delete the selection", "", "", "", "", "", "")%>"
				       <%=DisableClass(disabled, IIF(multipleSelection, "button_textarea", "button_input"))%> id="link_reset_<%= fieldName %>">
					    RESET
                    </a>
                <% end if %>
            </td>
            <% if obbligatorio then %>
				<td style="width:15px;">&nbsp;(*)</td>
			<% end if %>
		</tr>
	</table>
	<%
    'verifica distruzione 
    if rsCreated then
		set rs = nothing
	end if
    if connCreated then
        conn.close()
        set conn = nothing
	end if
end sub


'.................................................................................................
'..     procedura che scrive l'input per la selezione degli amministratori
'..     conn:                   connessione aperta
'..     rs:                     recordset creato
'..     ContattisqlCondition    condizione di filtro per i contatti da caricare per la selezione
'..     formName                nome del form della pagina chiamante
'..     fieldName               nome dell'input html che contiene il valore
'..     fieldValue              contatto o elenco di contatti selezionati (id)
'..     listType                tipologia di elenco:
'..                                 ""                  elenco semplice
'..                                 "EMAILMANDATORY"    elenco con email visualizzata e selezionabili solo con email valida
'..                                 "FAXMANDATORY"      elenco con tel visualizzata e selezionabili solo con tel valida
'..									"CELLMANDATORY"     elenco con cell visualizzata e selezionabili solo con cell valida
'..     multipleSelection       se true genera un input textarea per la selezione di piu' contatti
'..     obbligatorio            indica se viene gestito il pulsante reset e visualizza (*)
'..     disabled                indica se l'input e' disabilitato o meno
'..     OnChangeSelectionMethod permette la gesione dell'evento di selezione indicando le operazioni da eseguire
'..                             secondo le opzioni:     ""              se vuoto non esegue nulla (Selezione semplice)
'..                                                     "REDIRECT"      esegue il redirect della pagina chiamante alla pagina stesasa con il parametro IDCNT impostato all'id del contatto
'..                                                     "SUBMIT"        esegue il submit del form che contiene l'input
'..                                                     "<nome metodo>" permette di eseguire la funzione <nome metodo> alla selezione del contatto
'.................................................................................................
sub WriteAdminPicker_Input(conn, rs, ContattisqlCondition, formName, fieldName, fieldValue, listType, multipleSelection, obbligatorio, disabled, OnChangeSelectionMethod)
    dim sql, AuxId, AuxValue, AuxHREF, AuxHTML
    'verifica oggetti database
    dim rsCreated, connCreated
    if not IsObjectCreated(rs) then
		rsCreated = true
		set rs = server.createobject("adodb.recordset")
	else
		rsCreated = false
	end if
    if not IsObjectCreated(conn) then
		connCreated = true
        set conn = Server.CreateObject("ADODB.Connection")
        conn.open Application("DATA_ConnectionString"),"",""
	else
		connCreated = false
	end if
    
    'imposta variabile per passaggio query
    session("CONDIZIONE_SELEZIONE_ADMIN_" & formName & "_" & fieldName) = cString(ContattisqlCondition)
    
    'recupera valori selezionati visualizzati
	dim ID
    if cString(fieldValue)<>"" then
        sql = " SELECT *, (admin_cognome "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) & SQL_IfIsNull(conn, "admin_nome", "''") &") AS NOME"& _
			  " FROM tb_admin"& _
			  " WHERE id_admin "& _
			  IIF(multipleSelection, "IN (" & replace(replace(fieldValue, "; ", ","), ";", "") & ") ", " = " & fieldValue) & _
        	  " ORDER BY admin_cognome "
		
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
        if multipleSelection then
            while not rs.eof
				ID = rs("id_admin")
                AuxId = AuxId & " " & ID & ";"
                AuxValue = AuxValue & " " & JSReplacerEncode(rs("admin_cognome") &" "& rs("admin_nome")) & ";"
                rs.movenext
            wend
        elseif not rs.eof then
			ID = rs("id_admin")
            AuxId = ID
            AuxValue = rs("admin_cognome") &" "& rs("admin_nome")
        end if
        rs.close
    else
        AuxValue = ""
    end if
    %>
	<script language="JavaScript" type="text/javascript">
		function <%= fieldName %>_seleziona_onClick(sender){
            if (<%= fieldName %>_is_input_enabled()){
                var url = "<%= GetLibraryPath() %>SelezioneAdmin.asp?FieldName=<%= fieldName %>&FormName=<%= formName %>&ListType=<%= listType %>&MultipleSelection=<%= IIF(multipleSelection, 1, 0) %>";
                OpenAutoPositionedScrollWindow(url, "<%= formName & "_" & fieldName %>", 600, 400, false);
            }
		}
        
        function <%= fieldName %>_is_input_enabled(){
            return !<%= formName %>.<%= fieldName %>.disabled;
        }
        
        function <%= fieldName %>_input_onChangeSelection(){
             <% if OnChangeSelectionMethod <>"" then 
                Select case uCase(OnChangeSelectionMethod)
                    case "SUBMIT" 
                        'esegue submit del form dopo la selezione 
						%>
                        <%= formName %>.submit();
                    <% case "REDIRECT"
                        'esegue redirect dopo la selezione del contatto 
						%>
                        document.location = "<%= GetCurrentUrl() %>?IDADM=" + <%= formName %>.<%= fieldName %>.value;
                    <% case else 
                        'esegue il metodo passato come parametro
						%>
                        <%= OnChangeSelectionMethod %>;
                <% end select
            end if %>
        }
		
        function <%= fieldName %>_is_selected_contatto(id){
            var value = <%= formName %>.<%= fieldName %>.value;
            return (value.indexOf(<%= fieldName %>_id_format(id)) != -1);
        }
        
        function <%= fieldName %>_id_format(id){
            <% if multipleSelection then %>
                return ' ' + id + ';';
            <% else %>
                return id;
            <% end if %>
        }
        
        function <%= fieldName %>_nome_format(nome){
            <% if multipleSelection then %>
                return ' ' + nome + ';';
            <% else %>
                return nome;
            <% end if %>
        }
                
        function <%= fieldName %>_selezione_contatto(sender, id, nome){
            <% if multipleSelection then 
                'selezione di un elenco di contatti
				%>
                if (!<%= fieldName %>_is_selected_contatto(id)){
                    //selezione del contatto
			        <%= formName %>.view_<%= fieldName %>.value += <%= fieldName %>_nome_format(nome);
                    <%= formName %>.<%= fieldName %>.value += <%= fieldName %>_id_format(id);
                }
                else {
                    //deselezione del contatto
                    var re = eval('/' + <%= fieldName %>_nome_format(nome) + '/g');
                    <%= formName %>.view_<%= fieldName %>.value = <%= formName %>.view_<%= fieldName %>.value.replace(re, '');
			        re = eval('/' + <%= fieldName %>_id_format(id) + '/g');
			        <%= formName %>.<%= fieldName %>.value = <%= formName %>.<%= fieldName %>.value.replace(re, '');
                }
            <% else
                'selezione diretta di un contatto
				%>
			    <%= formName %>.view_<%= fieldName %>.value = nome;
                <%= formName %>.<%= fieldName %>.value = id;
                <%= fieldName %>_input_onChangeSelection();
            <% end if %>
        }
        
		function <%= fieldName %>_reset_onClick(){
            if (<%= fieldName %>_is_input_enabled()){
	    		<%= formName %>.view_<%= fieldName %>.value = '';
                <%= formName %>.<%= fieldName %>.value = '';
                <%= fieldName %>_input_onChangeSelection();
            }
		}
	</script>
	<table cellspacing="0" cellpadding="0" style="width:100%">
	    <input type="hidden" name="<%= fieldName %>" id="<%= fieldName %>" value="<%= AuxId %>" <%= disable(disabled) %>>
		<tr>
            <% AuxHREF = " onclick=""" & fieldName & "_seleziona_onClick(this)"" title=""click per aprire la finestra per la selezione degli amministratori."" "
            AuxHTML = " style=""width:100%;"" name=""view_" & fieldName & """ id=""view_" & fieldName & """ " & DisableClass(disabled, "")%>
			<td style="width:95%;">
                <% if multipleSelection then %>
    				<textarea READONLY rows="3" <%= AuxHTML %> <%= AuxHREF %>><%= AuxValue %> </textarea>
                <% else %>
                    <input READONLY type="text" <%= AuxHTML %> <%= AuxHREF %> value="<%= AuxValue %>">
                <% end if %>
			</td>
			<td style="width:60px;">
                <a href="javascript:void(0);" <%= ACTIVE_STATUS %> <%= AuxHREF %> id="link_scegli_<%= fieldName %>"
                    <% if multipleSelection then %> 
                        <%= DisableClass(disabled, "button_textarea") %> <% if obbligatorio then %> style="line-height:440%"<%end if 
                    else %>
                        <%= DisableClass(disabled, "button_input") %>
                    <% end if %>>
					SCEGLI
				</a>
                <% if not obbligatorio then 
                    if not multipleSelection then%>
                        </td>
                        <td>
                    <% end if %>
                    <a href="javascript:void(0);" <%= ACTIVE_STATUS %> onclick="<%= fieldName %>_reset_onClick()" title="cancella la selezione"
				       <%=DisableClass(disabled, IIF(multipleSelection, "button_textarea", "button_input"))%> id="link_reset_<%= fieldName %>">
					    RESET
                    </a>
                <% end if %>
            </td>
            <% if obbligatorio then %>
				<td style="width:15px;">&nbsp;(*)</td>
			<% end if %>
		</tr>
	</table>
	<%
    'verifica distruzione 
    if rsCreated then
		set rs = nothing
	end if
    if connCreated then
        conn.close()
        set conn = nothing
	end if
end sub


'.................................................................................................
'..     genera la parte di form necessaria per selezionare ed impostare i dati da google maps
'..     conn:                   connessione aperta al database. Facoltativa
'..     rs:                   	recordset aperto in modifica. Se null in inserimento.
'..     prefix:                 prefisso compreso "_"
'..     addressFormFieldList	lista di campi del form utilizzati per calcolare l'indirizzo separati da ;
'.................................................................................................
Sub WriteGoogleMaps_Input(conn, rs, prefix, addressFormFieldList)
	call WriteGoogleMaps_Input_Ex(conn, rs, prefix, addressFormFieldList, "")
End Sub

Sub WriteGoogleMaps_Input_Ex(conn, rs, prefix, addressFormFieldList, suffix)
	dim lat, lng
	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		lat = request("nfn_"&prefix&"google_maps_latitudine"&suffix)
		lng = request("nfn_"&prefix&"google_maps_longitudine"&suffix)
	else
		lat = CBR(rs, prefix &"google_maps_latitudine"&suffix, "nfn_")
		lng = CBR(rs, prefix &"google_maps_longitudine"&suffix, "nfn_")
	end if
	%>
	<script language="JavaScript" type="text/javascript">
		var GOOGLE_MAPS_APPLICATION_PATH = '<%= GetLibraryPath() %>google_maps/';
	
		function <%= prefix %>_google_maps_SELECT<%=suffix%>(path){
			OpenAutoPositionedScrollWindow(<%= prefix %>_google_maps_GetHref<%=suffix%>(path, 'select.asp'), 'gmaps_select', 700, 500, true);
		}
		
		function <%= prefix %>_google_maps_GetHref<%=suffix%>(path, page){
			return path + page + '?prefix=<%= prefix %>' + 
				   '&lat=' + form1.nfn_<%= prefix %>google_maps_latitudine<%=suffix%>.value + 
				   '&lon=' + form1.nfn_<%= prefix %>google_maps_longitudine<%=suffix%>.value +
				   '&suffix=<%=suffix%>';
		}
		
		function <%= prefix %>_google_maps_SetCoords<%=suffix%>(lat, lng){
			form1.nfn_<%= prefix %>google_maps_latitudine<%=suffix%>.value = lat;
			form1.nfn_<%= prefix %>google_maps_longitudine<%=suffix%>.value = lng;
			
			var gmapframe = document.getElementById('gmaps<%=suffix%>');
			if (gmapframe){
				gmapframe.src = <%= prefix %>_google_maps_GetHref<%=suffix%>(GOOGLE_MAPS_APPLICATION_PATH, 'preview.asp');
			}
		}
		
		function <%= prefix %>_google_maps_LOCATE_BY_ADDRESS<%=suffix%>(){
			<% if cString(addressFormFieldList)<>"" then %>
				var oField;
				var address = '';
				var gmapframe = document.getElementById('gmaps<%=suffix%>');
				if (gmapframe){
					<% dim field
					for each field in split(addressFormFieldList, ";")
						field = trim(field)
						if field <> "" then %>
							oField = document.getElementById('<%= field %>');
							
							if (oField == null)
								oField = document.getElementsByName('<%= field %>')[0];
								
							if (oField){
								if(oField.value != ''){
									if(address != '')
										address += ', ';
									address += oField.value;
								}
							}
						<%end if
					next%>
					
					//richiama funzione su finestra di preview.
					if (address != '')
						//frames['gmaps<%=suffix%>'].LocateByAddress(address);
						//frames[0].LocateByAddress(address);
						for (var f = 0; f < frames.length; f++) {
							var frameId;
							frameId = frames[f].frameElement.id;
							if (frameId == "gmaps<%=suffix%>") {
								frames[f].LocateByAddress(address);
							}
						}
					else
						<%= prefix %>_google_maps_RESET<%=suffix%>();
				}
			<% end if %>
			return void(0);
		}
		
		function <%= prefix %>_google_maps_RESET<%=suffix%>(){	
			<%= prefix %>_google_maps_SetCoords<%=suffix%>('', '');
		}
		
	</script>
	<input type="hidden" name="blank" value="">
	<input type="hidden" name="nfn_<%= prefix %>google_maps_latitudine<%=suffix%>" id="nfn_<%= prefix %>google_maps_latitudine<%=suffix%>" value="<%= lat %>">
	<input type="hidden" name="nfn_<%= prefix %>google_maps_longitudine<%=suffix%>" id="nfn_<%= prefix %>google_maps_longitudine<%=suffix%>" value="<%= lng %>">
	<table cellspacing="0" cellpadding="0" width="100%" class="gmap<%=suffix%>_container">
		<tr>
			<td class="map" rowspan="<%= IIF(cString(addressFormFieldList)<>"", 3, 2) %>">
				<iframe id="gmaps<%=suffix%>" class="gmaps_iframe" frameborder="0" src="<%= GetLibraryPath() %>google_maps/preview.asp?prefix=<%= prefix %>&lat=<%= lat %>&lon=<%= lng %>&suffix=<%= suffix %>" scrolling="No"></iframe>
			</td>
			<td style="width:100px;">
				<a href="javascript:void(0)" class="button_textarea" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "click per selezionare un punto nella mappa con Google Maps", "click to select a point on the map with Google Maps", "", "", "", "", "", "")%>"
				   onclick="<%= prefix %>_google_maps_SELECT<%=suffix%>(GOOGLE_MAPS_APPLICATION_PATH)"
				   style="<% if cString(addressFormFieldList)<>"" then %>line-height:170%;padding:6px;<% else %>height:90px; padding-top:10px;<% end if %>">
				   	<%= IIF(cString(addressFormFieldList)<>"", ChooseValueByAllLanguages(Session("LINGUA"), "seleziona<br>sulla mappa", "select<br>on the map", "", "", "", "", "", ""), ChooseValueByAllLanguages(Session("LINGUA"), "seleziona<br> il punto<br> sulla mappa", "select<br> the point<br> on the map", "", "", "", "", "", "")) %></a>
			</td>
		</tr>
		<% if cString(addressFormFieldList)<>"" then %>
			<tr>
				<td>
					<a href="javascript: <%= prefix %>_google_maps_LOCATE_BY_ADDRESS<%=suffix%>()" class="button_textarea" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "click per impostare automaticamente il punto sulla mappa dall'indirizzo.", "click to automatically set the point on the map from the address.", "", "", "", "", "", "")%>"
					   style="padding:6px;line-height:170%;">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "calcola<br>da indirizzo</a>", "calculate<br>by address</a>", "", "", "", "", "", "")%>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td>
				<a href="javascript:void(0)" class="button_textarea" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "click per rimuovere la selezione", "click to remove the selection", "", "", "", "", "", "")%>"
				   onclick="<%= prefix %>_google_maps_RESET<%=suffix%>()"
				   style="<% if cString(addressFormFieldList)<>"" then %>line-height:170%;padding:4px; padding-bottom:3px;<% else %>height:32px; padding-top:4px;<% end if %>">
					reset</a>
			</td>
		</tr>
	</table>
<%
End Sub


'.................................................................................................
'scrive il pulsante per il post di una specifica ricerca in un'altra pagina
'	pagina:			pagina di destinazione del post di ricerca
'	nomeCampo:		nome del campo da impostare come filtro
'	value:			valore del filtro
'	label:			label del link che scatenta il post
'	cssClass:		stili del link visibile che scatena il post di ricerca
'.................................................................................................
Sub WriteCampoCerca(pagina, nomeCampo, value, label, cssClass) 
	dim FormName
	FormName = "ricerca_" & nomeCampo & "_" & RemoveInvalidChar(label, JS_OBJECTS_NAME_CHARSET) & "_" & RemoveInvalidChar(cString(rnd()), JS_OBJECTS_NAME_CHARSET)
	%>
	<form action="<%= pagina %>" method="post" name="<%= FormName %>" id="<%= FormName %>" style="display:none; visibility:hidden;">
		<input type="hidden" name="search_<%= nomeCampo %>" value="<%= value %>">
		<input type="hidden" name="cerca" value="<%= label %>">
	</form>
	<a class="<%= cssClass %>" href="javascript:<%= FormName %>.submit()" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Ricerca ", "Search ", "", "", "", "", "", "")%> <%= lcase(label) %>" <%= ACTIVE_STATUS %>>
		<%= label %>
	</a>
<%
End Sub



'.................................................................................................
'verifica se &egrave; attiva o meno l'area riservata.
'	conn:		connessione aperta su database.
'.................................................................................................
function IsAreaRiservataActive(conn)
	dim connCreated, sql
    if not IsObjectCreated(conn) then
		connCreated = true
        set conn = Server.CreateObject("ADODB.Connection")
        conn.open Application("DATA_ConnectionString"),"",""
	else
		connCreated = false
	end if
	
	sql = "SELECT COUNT(*) FROM tb_siti WHERE NOT "& SQL_isTrue(conn, "sito_amministrazione")
	IsAreaRiservataActive = CInt(GetValueList(conn, NULL, sql)) > 0 
	
	if connCreated then
        conn.close()
        set conn = nothing
	end if
end function




'restituisce true se l'utente corrente ha il permesso PASS_ADMIN del NextPassport
Function IsAdminCurrent(conn)
	IsAdminCurrent = CIntero(GetValueList(conn, NULL, " SELECT COUNT(*) FROM rel_admin_sito"& _
											   		  " WHERE rel_as_permesso = 1 AND sito_id = "& NEXTPASSPORT & _
											   		  " AND admin_id = "& session("ID_ADMIN"))) > 0
End Function


'scrive il contenuto del file nel percorso e con il nome indicato.
Sub WriteFileContent(filePath, fileContent, overWrite)
	dim fso, textFile
	
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	if overWrite then
		set textFile = fso.CreateTextFile(filePath, true, false)
	else
		set textFile = fso.OpenTextFile(filePath, 8, true)
	end if
	call textFile.write(fileContent)
	textFile.close
	
	set fso = nothing
	set textFile = nothing
end sub


'scrive il contenuto del file in modo binario nel percorso e con il nome indicato.
Sub WriteBinaryFileContent(filePath, fileContent, overWrite)
	dim oStream
	
	Set oStream = CreateObject("ADODB.Stream")
	oStream.Open
	oStream.Type = adTypeBinary
	
	oStream.Write fileContent
	oStream.SaveToFile filePath, adSaveCreateOverWrite
	
	oStream.Close
	
	set oStream = nothing
end sub


'legge il contenuto del file testuale indicato.
function ReadFileContent(filePath)
	dim fso, textFile
	
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set textFile = fso.OpenTextFile(filePath)
	
	ReadFileContent = textFile.ReadAll()
	textFile.close
	
	set textfile = nothing
	set fso = nothing
end function


'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
'CharSet - charset of the Text - default is "us-ascii"
Function Stream_StringToBinary(Text, CharSet)
  Const adTypeText = 2
  Const adTypeBinary = 1
  
  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText
  
  'Specify charset For the source text (unicode) data.
  If Len(CharSet) > 0 Then
    BinaryStream.CharSet = CharSet
  Else
    BinaryStream.CharSet = "us-ascii"
  End If
  
  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text
  
  
  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary
  
  'Ignore first two bytes - sign of
  BinaryStream.Position = 0
  
  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read
End Function


'**********************************************************************************************************
'integrazione con autocompletamento di lightbox

'..........................................................................................................
'scrive il div applicato all'elemento da "autocompletare"
'	HtmlFieldId			id del campo html
'	AutocompleteUrl		url da cui ricavare la lista per l'autocompletamento
'..........................................................................................................
sub Lightbox_Autocomplete_DIV(HtmlFieldId, AutocompleteUrl)
	if instr(1, AutoCompleteUrl, "?", vbTextCompare)>0 then
		AutocompleteUrl = AutocompleteUrl & "&"
	else
		AutocompleteUrl = AutocompleteUrl & "?"
	end if
	AutocompleteUrl = AutocompleteUrl & "input=" & HtmlFieldId
	
	%>
	<div id="div_<%= HtmlFieldId %>" style="display:none;" class="autocompletion"></div>
	<script type="text/javascript" language="javascript">
		// <![CDATA[
		new Ajax.Autocompleter('<%= HtmlFieldId %>','div_<%= HtmlFieldId %>','<%= AutoCompleteUrl %>',{});
		// ]]>
	</script>
	<%
end sub


'..........................................................................................................
'restituisce la lista html dei risultati da selezionare
'	query:			query che restituisce l'elenco di valori
'	valueToSet:		eventuale valore fisso da restituire
'..........................................................................................................
sub Lightbox_Autocomplete_QUERY(query, valueToSet)
	dim conn, rs
	set Conn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	Conn.Open Application("DATA_ConnectionString"), "", ""
	
	if query<>"" then
		rs.open query, conn, adOpenstatic, adLockOptimistic 
		%>
		<ul>
			<% while not rs.eof %>
				<li><%= valueToset %><%= rs(0) %></li>
				<% rs.movenext
			wend %>
		</ul>
		<%
		rs.close
	end if
	
	conn.close
	set rs = nothing
	set conn = nothing
end sub



'..........................................................................................................
'restituisce la password criptata
'	password:		password da criptare
'..........................................................................................................
Function EncryptPassword(password)
	dim Cripto
	set Cripto = new CryptographyManager
	EncryptPassword = Left(Cripto.aes_of_string(UCASE(password), UCASE(password)), 50)
	set Cripto = nothing
end function


'..........................................................................................................
'	copia della query dalla connessione di origine alla connessione di destinazione
'	le due tabelle di origine e destinazione devono essere identiche.
'	connSource:		connesssione da cui prendere i dati
'	connDest:		connessione in cui copiare i dati
'	sqlSource:		sql di origine
'	sqlDest:		sql di destinazione
'	primaryKey:		campo chiave primaria per verifica duplicazione in destinazione
'..........................................................................................................
sub CopyTable(connSource, connDest, sqlSource, sqlDest, primaryKey)
	dim rss, rsd, field, sql
	set rss = Server.CreateObject("ADODB.recordset")
	set rsd = Server.CreateObject("ADODB.recordset")
	
	rss.open sqlSource, connSource, adOpenStatic, adLockOptimistic, adCmdText
	
	while not rss.eof
		sql = sqlDest & IIF(instr(1, sqlDest, "WHERE", vbTextCompare)>0, " AND ", " WHERE ") & primaryKey & "=" & rss(primaryKey)
		rsd.open sql, connDest, adOpenStatic, adLockOptimistic, adCmdText
		if rsd.eof then
			rsd.addnew
			rsd(primaryKey) = rss(primaryKey)
		end if
		for each field in rss.fields
			if lCase(field.name) <> lCase(primaryKey) then
				rsd(field.name) = field.value
			end if
		next
		rsd.update
		rsd.close
		rss.movenext
	wend
	
	rss.close

end sub
%>