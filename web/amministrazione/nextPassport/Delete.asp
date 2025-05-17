<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassDelete.asp" -->
<!--#INCLUDE FILE="../library/ClassIndirizzarioLock.asp" -->
<html>
<head>
	<title><%= Session("NOME_APPLICAZIONE") %></title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus()">
<%dim Class_delete, sql
set Class_delete = new OBJ_Delete

'parametri da impostare per sito
Class_delete.Section = request.Querystring("SEZIONE")
Class_delete.ID_Value = request.Querystring("ID")
Class_delete.PageName = "Delete.asp"

Class_delete.ReloadOpener = TRUE
Class_delete.ConnectionString = Application("DATA_ConnectionString")
Class_delete.LinkStyle = "class=""button"""
Class_delete.MessageStyle = ""
Class_delete.CaptionStyle = "style=""font-weight:bold;"""
Class_delete.CaptionColor = "#E6E6E6"
Class_delete.BorderDarkColor = "#919191"
Class_delete.BorderLightColor = "#FFFFFF"
Class_delete.BackgroundColor = "#F4F4F4"
Class_delete.DeleteRelations = FALSE
Class_delete.AfterDelete = FALSE

'parametri da impostare per ogni sezione
Select case request.Querystring("SEZIONE")
	case "APPLICAZIONI"
		Class_delete.Message = "Cancellare l'applicazione <RECORD>?"
		Class_delete.Name_Field = "sito_nome"
		Class_delete.ID_Field = "id_sito"
		Class_delete.Table = "tb_siti"
		Class_delete.Caption = "Gestione applicazioni"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = TRUE
	case "APPLICAZIONI_PARAMETRI"
		Class_delete.Message = "Cancellare il parametro <RECORD>?"
		Class_delete.Name_Field = "par_key"
		Class_delete.ID_Field = "par_id"
		Class_delete.Table = "tb_siti_parametri"
		Class_delete.Caption = "Gestione applicazioni - parametri"
	case "AMMINISTRATORI"
		Class_delete.Message = "Cancellare l'utente dell'area amministrativa <RECORD>?"
		Class_delete.Name_Field = "(admin_cognome + ' ' + admin_nome)"
		Class_delete.ID_Field = "id_admin"
		Class_delete.Table = "tb_admin"
		Class_delete.Caption = "Gestione utenti area amministrativa"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = FALSE
	case "GRUPPI"
		Class_delete.Message = "Cancellare il gruppo di lavoro <RECORD>?"
		Class_delete.Name_Field = "nome_gruppo"
		Class_delete.ID_Field = "id_gruppo"
		Class_delete.Table = "tb_gruppi"
		Class_delete.Caption = "Gruppi di lavoro"
	case "UTENTI"
		Class_delete.Message = "Cancellare l'utente dell'area riservata <RECORD>?"
        'permette la cancellazione del contatto se &egrave; bloccato solo dal next-Passport
        sql = "SELECT LockedByApplication FROM tb_indirizzario WHERE IDElencoIndirizzi IN (SELECT ut_nextCom_id FROM tb_utenti WHERE ut_id=" & Class_delete.ID_value & ")"
        if CInteger(GetValueList(Class_delete.conn, NULL, sql))=1 then
            'contatto bloccato solo dall'applicazione corrente: permette la cancellazione
            Class_delete.AddOption "delete_contatto", "cancella anche il contatto associato", true, ""
        else
            'contatto non cancellabile perch&egrave; bloccato da altre applicazioni
            Class_delete.Note = Class_delete.Note + " Non &egrave; possibile cancellare il contatto associato perch&egrave; utilizzato anche in altre applicazioni."
        end if
		Class_delete.Name_Field = ""
		Class_delete.ID_Field = "ut_id"
		Class_delete.Table = "tb_utenti"
		Class_delete.MsgSql = "SELECT (CognomeElencoIndirizzi + ' ' + NomeElencoIndirizzi + ' - ' + "& SQL_IfIsNull(Class_delete.conn, "NomeOrganizzazioneElencoIndirizzi", "''") &") " + _
							  " FROM tb_Indirizzario INNER JOIN tb_utenti" &_
							  " ON tb_indirizzario.IDElencoIndirizzi = tb_utenti.ut_NextCom_ID "
		Class_delete.Caption = "Gestione utenti area amministrativa"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = false
	case "SITI_TABELLE"
		Class_delete.Message = "Cancellare la tabella <RECORD>?"
		Class_delete.Name_Field = "tab_titolo"
		Class_delete.ID_Field = "tab_id"
		Class_delete.Table = "tb_siti_tabelle"
		Class_delete.Caption = "Tabelle dati"
	
	case "APPLICAZIONI_PARAMS_RAG"
		Class_delete.Message = "Cancellare il gruppo <RECORD>?"
		Class_delete.Name_Field = "sdr_titolo_it"
		Class_delete.ID_Field = "sdr_id"
		Class_delete.Table = "tb_siti_descrittori_raggruppamenti"
		Class_delete.Caption = "Gruppi di parametri"
	case "APPLICAZIONI_PARAMS"
		Class_delete.Message = "Cancellare il parametro <RECORD> e tutte le relazioni con le applicazioni?"
		Class_delete.Name_Field = "sid_nome_it"
		Class_delete.ID_Field = "sid_id"
		Class_delete.Table = "tb_siti_descrittori"
		Class_delete.Caption = "Parametri"
	
	case "ALERT"
		Class_delete.Message = "Cancellare l'alert <RECORD> e tutte le configurazioni?"
		Class_delete.Name_Field = "sev_nome_it"
		Class_delete.ID_Field = "sev_id"
		Class_delete.Table = "tb_siti_eventi"
		Class_delete.Caption = "Alert"
	case "ALERT_CONF"
		'controllo accesso
		CALL CheckAutentication(session("PASS_ADMIN") <> "")
		
		Class_delete.Message = "Cancellare la configurazione dell'alert <RECORD>?"
		Class_delete.Name_Field = "(SELECT sev_nome_it FROM tb_siti_eventi WHERE sev_id = rse_evento_id)"
		Class_delete.ID_Field = "rse_id"
		Class_delete.Table = "rel_siti_eventi"
		Class_delete.Caption = "Alert - configurazione"
	case "QUERY"
		Class_delete.Message = "Cancellare la query <RECORD>?"
		Class_delete.Name_Field = "tq_nome"
		Class_delete.ID_Field = "tq_id"
		Class_delete.Table = "tb_siti_tabelle_tag_query"
		Class_delete.Caption = "Gestione tabelle - query"
		Class_delete.DeleteRelations = FALSE
	case "AMMINISTRATORI_ORARIO"
		Class_delete.Message = "Cancellare l'orario di lavoro <RECORD>?"
		Class_delete.Name_Field = "'dal ' + ISNULL(CONVERT(VARCHAR, ao_data_dal, 120), '') + ' al ' + ISNULL(CONVERT(VARCHAR, ao_data_al, 120), '')"
		Class_delete.ID_Field = "ao_id"
		Class_delete.Table = "tb_admin_orario"
		Class_delete.Caption = "Gestione utenti area amministrativa - orario"
		Class_delete.DeleteRelations = FALSE
end select

'definizione eventuali operazioni su relazioni	
Sub Delete_Relazioni(conn, ID)
	dim sql, rs, cnt_id
	set rs = Server.CreateObject("ADODB.Recordset")
	
	Select case request.Querystring("SEZIONE")
		case "APPLICAZIONI"
			sql = "SELECT sito_amministrazione, sito_rubrica_area_riservata FROM tb_siti WHERE id_sito=" & ID
			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			if not rs("sito_amministrazione") then
				'applicazione dell'area riservata: cancella rubrica collegata
				sql = "DELETE FROM tb_rubriche WHERE id_rubrica=" & rs("sito_rubrica_area_riservata")
				CALL conn.execute(sql, 0, adExecuteNoRecords)
			end if
			rs.close
		case "AMMINISTRATORI"	
			'cancella cartella temporanea documenti
			dim fso
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			sql = "SELECT admin_login FROM tb_admin WHERE id_admin=" & ID
			rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
			if fso.FolderExists(Application("IMAGE_PATH") & "temp\docs\" & rs("admin_login")) then
				fso.DeleteFolder(Application("IMAGE_PATH") & "temp\docs\" & rs("admin_login"))
			end if
			rs.close
			set fso = nothing
		case "UTENTI"
			
            'recupera dati del contatto
    		sql = "SELECT ut_nextCom_ID FROM tb_utenti WHERE ut_id=" & ID
    		cnt_id = GetValueList(conn, rs, sql)
            
            'verifica opzioni scelte dall'utente per la cancellazione
			if request("delete_contatto")<>"" then
				'cancella contatto ed utente
				sql = "DELETE FROM tb_indirizzario WHERE IDElencoIndirizzi=" & cnt_id
				 CALL conn.execute(sql, , adExecuteNoRecords)
            else
                'cancella le associazioni con le rubriche di sistema per l'accesso all'area riservata
    			sql = " DELETE FROM rel_rub_ind WHERE id_indirizzo IN (SELECT ut_nextCom_ID FROM tb_utenti " &_
	    			  " WHERE ut_id=" & ID & ") AND id_rubrica IN (SELECT sito_rubrica_area_riservata FROM tb_siti)"
				CALL conn.execute(sql, 0, adExecuteNoRecords)
				
    			'sblocca il contatto
    			dim obj_contatto
    			set obj_contatto = new IndirizzarioLock
    			CALL obj_contatto.UnLockContact(cnt_id, NEXTPASSPORT)
    			
    			'sblocca anche le applicazione dell'area riservata che bloccano il contatto
    			sql = "SELECT id_sito FROM tb_siti WHERE not " + SQL_isTrue(conn, "sito_amministrazione")
    			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
				while not rs.eof
    				CALL obj_contatto.UnLockContact(cnt_id, rs("id_sito"))
    				rs.movenext
    			wend
    			rs.close
    			
                obj_contatto.conn = empty
    			set obj_contatto = nothing
            end if
	end select
	set rs = nothing
end Sub

Sub Operations_AfterDelete(conn, ID)
	dim sql, rs
	set rs = Server.CreateObject("ADODB.Recordset")
	
	Select case request.Querystring("SEZIONE")
		case "APPLICAZIONI"
			CALL ClearNextCom(conn)
	end select
	set rs = nothing
end Sub

Class_delete.Delete_Manager()
%>

</body>
</html>