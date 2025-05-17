<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassDelete.asp" -->
<!--#INCLUDE FILE="../library/ClassIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<html>
<head>
	<title><%= Session("NOME_APPLICAZIONE") %></title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus()">
<%dim Class_delete
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
	case "CONTATTI"
		Class_delete.Message = "Cancellare il contatto <RECORD>?"
		Class_delete.Name_Field = "ModoRegistra"
		Class_delete.ID_Field = "IDElencoIndirizzi"
		Class_delete.Table = "Tb_Indirizzario"
		Class_delete.Caption = "Anagrafica contatti"
		Class_delete.DeleteRelations = FALSE
		Class_delete.AfterDelete = TRUE
	case "ContattiCATEGORIE"
		Class_delete.Message = "Cancellare la categoria <RECORD>?"
		Class_delete.Name_Field = "icat_nome_it"
		Class_delete.ID_Field = "icat_id"
		Class_delete.Table = "tb_indirizzario_categorie"
		Class_delete.Caption = "Categorie dei contatti"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "ContattiCTECH"
		Class_delete.Message = "Cancellare la caratteristica <RECORD>?"
		Class_delete.Name_Field = "ict_nome_it"
		Class_delete.ID_Field = "ict_id"
		Class_delete.Table = "tb_indirizzario_carattech"
		Class_delete.Caption = "CARATTERISTICHE"
		Class_delete.AfterDelete = FALSE
	case "ContattiCTECH_GRUPPI"
		Class_delete.Message = "Cancellare il gruppo di caratteristiche <RECORD>?"
		Class_delete.Name_Field = "icr_titolo_it"
		Class_delete.ID_Field = "icr_id"
		Class_delete.Table = "tb_indirizzario_carattech_raggruppamenti"
		Class_delete.Caption = "GRUPPI DI CARATTERISTICHE"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "CONTATTI_ATTIVITA"
		Class_delete.Message = "Cancellare l'attivita' <RECORD>?"
		Class_delete.Name_Field = "ina_note"
		Class_delete.ID_Field = "ina_id"
		Class_delete.Table = "tb_indirizzario_attivita"
		Class_delete.Caption = "ATTIVITA' CON I CONTATTI"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
		if cIntero(request("CNT_ID")) > 0 then
			Session("DELETE_CONTATTI_ATTIVITA_CNT_ID") = cIntero(request("CNT_ID"))
		end if
	case "RECAPITI"
		Class_delete.Message = "Cancellare il recapito <RECORD>?"
		Class_delete.Name_Field = "ValoreNumero"
		Class_delete.ID_Field = "id_ValoreNumero"
		Class_delete.Table = "tb_ValoriNumeri"
		Class_delete.Caption = "Recapiti del contatto"
	case "RUBRICHE"
		Class_delete.Message = "Cancellare la rubrica <RECORD> ed i contatti associati <b>SOLO</b> ad essa?"
		Class_delete.Name_Field = "nome_rubrica"
		Class_delete.ID_Field = "id_rubrica"
		Class_delete.Table = "tb_rubriche"
		Class_delete.Caption = "Rubriche"
		Class_delete.DeleteRelations = FALSE
		Class_delete.AfterDelete = TRUE
	case "EMAIL"
		Class_delete.Message = "Cancellare l'email <RECORD>?"
		Class_delete.Name_Field = "email_object"
		Class_delete.ID_Field = "email_id"
		Class_delete.Table = "tb_email"
		Class_delete.Caption = "Mailing in uscita"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = FALSE
	case "SMS"
		Class_delete.Message = "Cancellare il messaggio <RECORD>?"
		Class_delete.Name_Field = "email_text"
		Class_delete.ID_Field = "email_id"
		Class_delete.Table = "tb_email"
		Class_delete.Caption = "SMS"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = FALSE
	case "PRATICHE"
		Class_delete.Message = "Cancellare la pratica <RECORD>?"
		Class_delete.Name_Field = "pra_nome"
		Class_delete.ID_Field = "pra_id"
		Class_delete.Table = "tb_pratiche"
		Class_delete.Caption = "Pratiche"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = FALSE
	case "DOCUMENTI"
		Class_delete.Message = "Cancellare il documento <RECORD>?"
		Class_delete.Name_Field = "doc_nome"
		Class_delete.ID_Field = "doc_id"
		Class_delete.Table = "tb_documenti"
		Class_delete.Caption = "Documenti"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = FALSE
	case "FILES"
		Class_delete.Message = "Cancellare il file <RECORD>?"
		Class_delete.Note = "ATTENZIONE: cancellando il file verranno eliminate anche tutte le associazioni con i documenti " + _
							"e l'operazione non sar&agrave; reversibile."
		Class_delete.Name_Field = "F_original_name"
		Class_delete.ID_Field = "F_ID"
		Class_delete.Table = "tb_files"
		Class_delete.Caption = "Files dei documenti"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = FALSE
	case "TIPI"
		Class_delete.Message = "Cancellare la tipologia <RECORD>?"
		Class_delete.Name_Field = "tipo_nome"
		Class_delete.ID_Field = "tipo_id"
		Class_delete.Table = "tb_tipologie"
		Class_delete.Caption = "Tipologie"
	case "DESCRITTORI"
		Class_delete.Message = "Cancellare il descrittore <RECORD>?"
		Class_delete.Name_Field = "descr_nome"
		Class_delete.ID_Field = "descr_id"
		Class_delete.Table = "tb_descrittori"
		Class_delete.Caption = "Descrittori"
	case "ATTIVITA"
		Class_delete.Message = "Cancellare l'attivita' <RECORD>?"
		Class_delete.Name_Field = "att_oggetto"
		Class_delete.ID_Field = "att_id"
		Class_delete.Table = "tb_attivita"
		Class_delete.Caption = "Attivita"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = FALSE
	case "NEWSLETTER_TIP"
		Class_delete.Message = "Cancellare la tipologia di newsletter <RECORD>?"
		Class_delete.Name_Field = "nl_nome_it"
		Class_delete.ID_Field = "nl_id"
		Class_delete.Table = "tb_newsletters"
		Class_delete.Caption = "Tipologia di Newsletter"
		Class_delete.DeleteRelations = FALSE
		Class_delete.AfterDelete = FALSE
	case "CAMPAGNE"
		Class_delete.Message = "Cancellare la campagna <RECORD>?"
		Class_delete.Name_Field = "inc_nome"
		Class_delete.ID_Field = "inc_id"
		Class_delete.Table = "tb_indirizzario_campagne"
		Class_delete.Caption = "Campagne marketing"
		Class_delete.DeleteRelations = FALSE
		Class_delete.AfterDelete = TRUE
	case "CAMPAGNA_CONTATTO"
		Class_delete.Message = "Rimuove il contatto <RECORD>?"
		Class_delete.Name_Field = "Cast(IsNull((SELECT top 1 CASE WHEN IsSocieta=1 THEN NomeORganizzazioneElencoIndirizzi ELSE CognomeElencoindirizzi + ' ' + NomeElencoIndirizzi END FROM tb_Indirizzario WHERE IdElencoIndirizzi = rel_cnt_campagne.rcc_cnt_id),'') AS nvarchar(250)) + " + _
								  "  + '""</b> dalla campagna <b>""' +  " + _
								  " Cast(IsNull((SELECT top 1 inc_nome FROM tb_indirizzario_campagne WHERE inc_id = rel_cnt_campagne.rcc_campagna_id),'') AS nvarchar(250)) "
		Class_delete.ID_Field = "rcc_id"
		Class_delete.Table = "rel_cnt_campagne"
		Class_delete.Caption = "Campagne marketing"
		Class_delete.DeleteRelations = FALSE
		Class_delete.AfterDelete = TRUE
end select

'definizione eventuali operazioni su relazioni	
Sub Delete_Relazioni(conn, ID)
	dim sql, rs
	dim fso, path
	set rs = Server.CreateObject("ADODB.RecordSet")
	
	Select case request.Querystring("SEZIONE")
		CASE "ContattiCATEGORIE"
			CatContatti.Delete(ID)
			
		case "ContattiCTECH_GRUPPI"
			sql = "UPDATE tb_indirizzario_carattech SET ict_raggruppamento_id = NULL WHERE ict_raggruppamento_id=" & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
			
		case "CONTATTI_ATTIVITA"
			dim campagna_id, cnt_id
			campagna_id = cIntero(GetValueList(conn, NULL, "SELECT ina_campagna_conclusa_id FROM tb_indirizzario_attivita WHERE ina_id = " & ID))
			cnt_id = Session("DELETE_CONTATTI_ATTIVITA_CNT_ID")
			if campagna_id > 0 AND cnt_id > 0 then
				sql = " UPDATE rel_cnt_campagne SET rcc_data_conclusione = NULL " & _
					  " WHERE rcc_campagna_id = " & campagna_id & " AND rcc_cnt_id = " & cnt_id
				response.write sql
				CALL conn.execute(sql, , adExecuteNoRecords)
				Session("DELETE_CONTATTI_ATTIVITA_CNT_ID") = NULL
			end if

		case "EMAIL"
			
			'cancella directory degli allegati
			CALL FolderRemove(	CreateObject("Scripting.FileSystemObject"), _
								Application("IMAGE_PATH") & "\docs\eml_" & ID, false)
			
		CASE "PRATICHE"
			
			'cancello attività associate e relazioni con allegati ma non i documenti
			'slego i documenti dalla pratica
			sql = "DELETE FROM tb_allegati WHERE "& _
				  "all_attivita_id IN (SELECT att_id FROM tb_attivita WHERE att_pratica_id="& ID &")"
			CALL conn.execute(sql, , adExecuteNoRecords)
			sql = "DELETE FROM tb_attivita WHERE att_pratica_id="& ID
			CALL conn.execute(sql, , adExecuteNoRecords)
			sql = "UPDATE tb_documenti SET doc_pratica_id=0 WHERE doc_pratica_id="& ID
			CALL conn.execute(sql, , adExecuteNoRecords)
			
		CASE "DOCUMENTI"
			
			'aggiorno AL della pratica del documento
			sql = "SELECT doc_pratica_id FROM tb_documenti WHERE doc_pratica_id <> 0 " + _
				  " AND NOT "& SQL_IsNull(conn, "doc_pratica_id") &" AND doc_id="& ID
			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdtext
			if not rs.eof then
				CALL AL_ins(conn, AL_PRATICHE, rs("doc_pratica_id"), false)
			end if
			rs.close
		
		CASE "FILES"
			
			'elimina il file fisico
			Set fso = CreateObject("Scripting.FileSystemObject")
			
			sql = "SELECT F_encoded_name, F_encoded_path FROM tb_files WHERE F_id="	& ID
			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			path = Application("IMAGE_PATH") & "\docs\" & rs("F_encoded_path") & "\" & rs("F_encoded_name")
			if fso.FileExists(path) then
				fso.DeleteFile(path)
			end if
			rs.close
			
			set FSO = nothing
			
		CASE "ATTIVITA"
		
			'aggiorno AL della pratica dell'attivita
			sql = "SELECT att_pratica_id FROM tb_attivita WHERE att_id="& ID
			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdtext
			if cInteger(rs("att_pratica_id"))>0 then
				CALL AL_ins(conn, AL_PRATICHE, rs("att_pratica_id"), false)
			end if
			rs.close
			
	end select
	
	set rs = nothing
end Sub

Sub Operations_AfterDelete(conn, ID)	
	dim sql
	
	Select case request.Querystring("SEZIONE")
		case "CONTATTI"
			'cancella contatti interni
			sql = "DELETE FROM tb_indirizzario WHERE cntRel=" & ID
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		
		case "CAMPAGNE"
			sql = "DELETE FROM rel_cnt_campagne WHERE rcc_campagna_id=" & ID
			CALL conn.execute(sql, 0, adExecuteNoRecords)

		case "RUBRICHE"
			'cancella i contatti non piu' associati ad alcuna rubrica
			CALL ClearNextCom(conn)
	end select
	
end sub

Class_delete.Delete_Manager()
%>

</body>
</html>