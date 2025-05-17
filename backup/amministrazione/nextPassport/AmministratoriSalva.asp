<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Passport.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/ClassCryptography.asp"-->
<%
dim Classe, conn, rs, sql

	Set Classe = New OBJ_Salva
	
	'controlli di correttezza:
	'email corretta
	if not isEmail(request("tft_admin_email")) then
		Session("ERRORE") = "Indirizzo email dell'utente non valido!"
	end if
	
	if cString(request("tft_admin_email_newsletter")) <> "" AND not IsEmail(request("tft_admin_email_newsletter")) then
		Session("ERRORE") = "Indirizzo email mittente per newsletter non valido!"
	end if
	
	if not CheckChar(request("tft_admin_password"), LOGIN_VALID_CHARSET) then
		Session("ERRORE") = "Password non valida! Utilizzare solo caratteri alfanumerici o &quot;_&quot;"
	end if
	
	'inserimento nuovo utente:
	'controllo per conferma password
	if request("ID")="" then
		if request("tft_admin_password")="" OR  (not CheckChar(request("tft_admin_password"), LOGIN_VALID_CHARSET)) then
			Session("ERRORE") = "Password mancante o non valida! Utilizzare solo caratteri alfanumerici o &quot;_&quot;"
		elseif uCase(request("tft_admin_password")) <> uCase(request("conferma_password")) then
			Session("ERRORE") = "Errore nella conferma della password!"
		end if
	end if
	
	'controllo per correttezza login
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString"),"",""
	set rs = Server.CreateObject("ADODB.RecordSet")
	CALL Check_login(conn, rs, true, request("ID"), request("tft_admin_login"))
	set rs = nothing
	conn.close
	set conn = nothing
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_admin_cognome; tft_admin_email; tft_admin_login"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_admin"
	Classe.id_Field					= "id_admin"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim i, sql, gruppi, gruppo, password
	
	'gestione del gruppo di lavoro
	sql = "DELETE FROM tb_rel_dipgruppi WHERE id_impiegato=" & ID
	CALL conn.execute(sql, ,adExecuteNoRecords)
	
	gruppi = replace(request("gruppi_di_lavoro"), " ", "")
	gruppi = split(gruppi, ",")
	for each gruppo in gruppi
		sql = " INSERT INTO tb_rel_dipgruppi (id_impiegato, id_gruppo) " + _
			  " VALUES( " & ID & ", " & gruppo & ")"
		CALL conn.execute(sql, , adExecuteNoRecords)
	next
	
	CALL save_permessi(conn, rs, true, ID)
	
	'se l'utente è collegato al next-com deve essere collegato ad un gruppo di lavoro
	sql = "SELECT COUNT(*) FROM rel_admin_sito WHERE admin_id=" & ID & " AND sito_id=" & NEXTCOM
	if cInteger(GetValueList(conn, rs, sql))>0 then
		'utente con accesso al next-com: verifica se presente gruppo di lavoro
		sql = "SELECT COUNT(*) FROM tb_rel_dipgruppi WHERE id_impiegato=" & ID
		if cInteger(GetValueList(conn, rs, sql))=0 then
			sql = "SELECT sito_nome FROM tb_siti WHERE id_sito=" & NEXTCOM
			Session("ERRORE") = "Per accedere a &quot;" & GetValueList(conn, rs, sql) & "&quot; è necessario il gruppo di lavoro."
		end if
	end if
	
	'generazione cartella di destinazione dei file utente (cartella temporanea)
	if (request("old_admin_login") <> request("tft_admin_login")) AND Session("ERRORE")="" then
		CALL CreateTemporaryDir(request("tft_admin_login"), request("old_admin_login"))
	end if
	
	if request("ID") = "" then
		'cripto la password: solo in inserimento.
		sql = "SELECT admin_password FROM tb_admin WHERE id_admin = " & ID
		password = GetValueList(conn, NULL, sql)
		password = EncryptPassword(password)
		sql = "UPDATE tb_admin SET admin_password = '" &password& "' WHERE id_admin = " & ID
		CALL conn.execute(sql, , adExecuteNoRecords)
	end if
	
	'scrivo sul log
	dim subject, text, code
	if request("ID")<>"" then
		CALL LogModificaUtente(conn, "AdminModifica", ID , request("tft_admin_login"), request("tft_admin_email"))
	else
		CALL LogModificaUtente(conn, "AdminInserimento", ID , request("tft_admin_login"), request("tft_admin_email"))
	end if
	
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	
    if request("salva_elenco")<>"" then
		Classe.Next_Page = "Amministratori.asp"
	else
        Classe.Next_Page = "AmministratoriMod.asp?ID="& ID
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>