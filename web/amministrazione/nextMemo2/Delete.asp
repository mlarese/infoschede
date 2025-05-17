<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassDelete.asp" -->
<!--#INCLUDE FILE="../library/ClassIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<!--#INCLUDE FILE="Tools4Issuu.asp" -->
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

'..............................................................................
'impostazione dei dati dell'indice
Class_delete.Index = Index
'..............................................................................

'parametri da impostare per ogni sezione
Select case request.Querystring("SEZIONE")
	case "DOCUMENTI"
		Class_delete.Message = "Cancellare il documento / circolare <RECORD>?"
		Class_delete.Name_Field = "doc_titolo_it"
		Class_delete.ID_Field = "doc_id"
		Class_delete.Table = "mtb_documenti"
		Class_delete.Caption = "Gestione documenti / circolari"
		
		'aggiungo l'opzione per la cancellazione dei file associati
		dim rs, rsa, check, file_path, i
		check = ";"
		set rs = Server.CreateObject("ADODB.Recordset")
		set rsa = Server.CreateObject("ADODB.Recordset")
		rs.open "SELECT * FROM tb_cnt_lingue", Class_delete.conn, adOpenKeySet, adLockOptimistic, adCmdText
		rsa.open "SELECT * FROM mtb_documenti WHERE doc_id="&Class_delete.ID_value, Class_delete.conn, adOpenKeySet, adLockOptimistic, adCmdText		
		while not rs.eof
			file_path = cString(rsa("doc_file_"&rs("lingua_codice")))
			if file_path <> "" then
				if inStr(check,";"&file_path&";") = 0 then
					check = check & file_path & ";"
					
					'costruisco la condizione per la query che controlla se il file è usato anche in altri documenti
					sql = ""
					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
						if i = 0 then
							sql = " ("
						else
							sql = sql & " OR "
						end if
						sql = sql & " doc_file_"&Application("LINGUE")(i)&" LIKE '"&file_path&"' "
					next
					sql = sql & ") "
					
					if GetValueList(Class_delete.conn,NULL,"SELECT COUNT(*) FROM mtb_documenti WHERE doc_id<>"&Class_delete.ID_value&" AND "&sql)>0 then
						Class_delete.Note = Class_delete.Note & "Impossibile cancellare il file associato <b>&quot;"&file_path&"</b>&quot; perch&egrave; utilizzato in altri documenti.<br>"
					else
						Class_delete.AddOption "delete_doc_file_"&rs("lingua_codice") , "cancella anche il file associato <b>&quot;"&file_path&"</b>&quot;", false, ""
					end if
				end if
			end if
			rs.moveNext
		wend
		rs.close
		
		'verifica presenza "sfoglia catalogo" con ISSUU
		dim lingua
		for each lingua in Application("LINGUE")
			
			if cString(rsa("doc_url_catalogo_" & lingua))<>"" then
				'avvia la cancellazione del documento
				dim postData, signatureData, httpRequest, httpResponse, docName
				
				Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
				httpRequest.Open "POST", ISSUU_url, False
				httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				
				docName = replace(replace(rsa("doc_url_catalogo_" & lingua), ISSUU_baseurl, ""), ISSUU_parameters, "")
				CALL AddIssuuPost("action", "issuu.document.delete", false, postData, signatureData)
				CALL AddIssuuPost("apiKey", ISSUU_apiKey, false, postData, signatureData)
				CALL AddIssuuPost("format", "xml", false, postData, signatureData)
				CALL AddIssuuPost("names", docName, false, postData, signatureData)
				
				CALL AddIssuuSignature(postData, signatureData)
					
				httpRequest.Send postData
				httpResponse = httpRequest.ResponseXML.xml
				
				if instr(1, httpResponse, "stat=""ok""", vbTextCompare)>0 then
					'cancellazione OK - invia email
					CALL SendEmailSupportEX("Issuu: "& Request.ServerVariables("SERVER_NAME") & " - cancellazione pdf da ISSUU", _
											"cancellazone file riuscita: " & rsa("doc_url_catalogo_" & lingua) & vbCrLF & _
											"Risultato cancellazone:" & vbCrLF & _
											httpResponse & vbCrLF & _
											"POST Inviato:" & vbCrLf & _
											postData)
				else
					'conversione fallita - manda avviso e scrive log
					CALL SendEmailSupportEX("Errore: "& Request.ServerVariables("SERVER_NAME") & " - cancellazione pdf da ISSUU", _
											"Errore nella cancellazone del file: " & rsa("doc_url_catalogo_" & lingua) & vbCrLF & _
											"Risultato cancellazone:" & vbCrLF & _
											httpResponse & vbCrLF & _
											"POST Inviato:" & vbCrLf & _
											postData)
				end if
				
			end if
			
		next
		
		rsa.close
		set rs = nothing
		set rsa = nothing
		set check = nothing
		set i = nothing
		
		Class_delete.AfterDelete = TRUE
		Class_delete.DeleteRelations = TRUE
	case "PROFILI"
		Class_delete.Message = "Cancellare il profilo <RECORD>?"
		Class_delete.Name_Field = "pro_nome_it"
		Class_delete.ID_Field = "pro_id"
		Class_delete.Table = "mtb_profili"
		Class_delete.Caption = "Gestione profili"
	case "CATEGORIE"
		Class_delete.Message = "Cancellare la categoria <RECORD>?"
		Class_delete.Name_Field = "catC_nome_it"
		Class_delete.ID_Field = "catC_id"
		Class_delete.Table = "mtb_documenti_categorie"
		Class_delete.Caption = "Categorie dei documenti"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "IMPEGNI"
		Class_delete.Message = "Cancellare l'impegno <RECORD>?"
		Class_delete.Name_Field = "imp_titolo_it"
		Class_delete.ID_Field = "imp_id"
		Class_delete.Table = "mtb_impegni"
		Class_delete.Caption = "Impegno/appuntamento"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = FALSE	
	case "AGENDA_CONFIGURA"
		Class_delete.Message = "Cancellare la configurazione <RECORD>?"
		Class_delete.Name_Field = "'giorno ' + CONVERT(nvarchar, coi_giorno) + ' - ' + CONVERT(nvarchar, coi_dal)"
		Class_delete.ID_Field = "coi_id"
		Class_delete.Table = "mtb_configurazione_impegni"
		Class_delete.Caption = "Configurazione agenda"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = FALSE
	case "IMPEGNI_TIPOLOGIA"
		Class_delete.Message = "Cancellare la tipologia <RECORD>?"
		Class_delete.Name_Field = "tim_nome_it"
		Class_delete.ID_Field = "tim_id"
		Class_delete.Table = "mtb_tipi_impegni"
		Class_delete.Caption = "Tipologia di impegnoi/appuntamento"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = FALSE	
	case "AMMINISTRATORI"
		Class_delete.Message = "Cancellare l'utente dell'area amministrativa <RECORD>?"
		Class_delete.Name_Field = "(admin_cognome + ' ' + admin_nome)"
		Class_delete.ID_Field = "id_admin"
		Class_delete.Table = "tb_admin"
		Class_delete.Caption = "Gestione utenti area amministrativa"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = FALSE
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
	case "CTECH"
		Class_delete.Message = "Cancellare la caratteristica <RECORD>?"
		Class_delete.Name_Field = "ct_nome_it"
		Class_delete.ID_Field = "ct_id"
		Class_delete.Table = "mtb_carattech"
		Class_delete.Caption = "CARATTERISTICHE"
		Class_delete.AfterDelete = FALSE
	case "CTECH_GRUPPI"
		Class_delete.Message = "Cancellare il gruppo di caratteristiche <RECORD>?"
		Class_delete.Name_Field = "ctr_titolo_it"
		Class_delete.ID_Field = "ctr_id"
		Class_delete.Table = "mtb_carattech_raggruppamenti"
		Class_delete.Caption = "GRUPPI DI CARATTERISTICHE"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
end select

'definizione eventuali operazioni su relazioni	
Sub Delete_Relazioni(conn, ID)
	dim sql, rs, cnt_id
	set rs = Server.CreateObject("ADODB.Recordset")
	Select case request.Querystring("SEZIONE")
		case "DOCUMENTI"
			sql = " DELETE FROM mrel_doc_ctech WHERE rdc_doc_id = " & ID
			CALL conn.execute(sql, 0, adExecuteNoRecords)
			
			dim var, campo
			Session("DELETE_DOC_FILES") = ""
			for each var in request.querystring
				if inStr(var, "delete_doc_file_") > 0 then
					campo = Replace(var, "delete_", "")
					Session("DELETE_DOC_FILES") = Session("DELETE_DOC_FILES") & GetValueList(conn,NULL,"SELECT "&campo&" FROM mtb_documenti WHERE doc_id="&ID)& ";"
				end if
			next
			set campo = nothing
			set var = nothing
			
		case "CATEGORIE"
			dim categorie
			set categorie = New objCategorie
				categorie.tabella = "mtb_documenti_categorie"
				categorie.prefisso = "catC"
			categorie.Delete(ID)
			set categorie = nothing
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
			
		case "CTECH_GRUPPI"
			sql = "UPDATE mtb_carattech SET ct_raggruppamento_id = NULL WHERE ct_raggruppamento_id=" & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
			
	end select
end Sub

Sub Operations_AfterDelete(conn, ID)	
	dim sql
	Select case request.Querystring("SEZIONE")
		case "DOCUMENTI"
			if Session("DELETE_DOC_FILES") <> "" then
				dim lista_file, file, path, fso
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				lista_file = split(Session("DELETE_DOC_FILES"), ";")
				
				for each path in lista_file
					if Trim(path) <> "" then
						file = Right(path, len(path) - inStrRev(path, "/"))
						path = Replace(path, file, "")
						CALL FileRemove(fso, Application("IMAGE_PATH")&Application("AZ_ID")&"\images\"&replace(path,"/","\") , file, false)				
					end if
				next
		
				set lista_file = nothing
				set path = nothing
				set file = nothing
				set fso = nothing
				Session("DELETE_DOC_FILES") = ""
			end if
			
	end select
end sub

Class_delete.Delete_Manager()
%>

</body>
</html>