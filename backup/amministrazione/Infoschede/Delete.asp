<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 10000 %>
<% response.buffer = true %>

<!--#INCLUDE FILE="../library/ClassDelete.asp" -->
<!--#INCLUDE FILE="../library/ClassIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Const.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Categorie.asp" -->

<html>
<head>
	<title><%= Session("NOME_APPLICAZIONE") %></title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus()">
<%dim Class_delete, sql, conn
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
	case "PROBLEMI"
		Class_delete.Message = "Cancellare il problema <RECORD>?"
		Class_delete.Name_Field = "prb_nome_it"
		Class_delete.ID_Field = "prb_id"
		Class_delete.Table = "sgtb_problemi"
		Class_delete.Caption = "Gestione problemi"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	' case "PROFILI"
		' Class_delete.Message = "Cancellare il profilo <RECORD>?"
		' Class_delete.Name_Field = "pro_nome_it"
		' Class_delete.ID_Field = "pro_id"
		' Class_delete.Table = "mtb_profili"
		' Class_delete.Caption = "Gestione profili"
	case "PROBLEMI_ARTICOLI"
		Class_delete.Message = "Cancellare il collegamento con il modello <RECORD>?"
		Class_Delete.Note = "Verr&agrave; cancellata solo l'associazione e non il modello. Per cancellare completamente " + _
							"anche il modello utilizzare l'apposita sezione."
		Class_delete.MsgSql = " SELECT (art_cod_int " + SQL_concat(Class_delete.conn) + "' - '" + SQL_concat(Class_delete.conn) + " art_nome_it) AS NOMINATIVO FROM " + _
							  " gv_articoli INNER JOIN srel_problemi_articoli ON gv_articoli.rel_id = srel_problemi_articoli.rpa_articolo_rel_id "
		Class_delete.Name_Field = "NOMINATIVO"
		Class_delete.ID_Field = "rpa_id"
		Class_delete.Table = "srel_problemi_articoli"
		Class_delete.Caption = "Gestione articoli associati"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = FALSE
	case "PROBLEMI_MAR_TIP"
		Class_delete.Message = "Cancellare il collegamento tra il guasto e la marca/categoria <RECORD>?"
		Class_Delete.Note = "Verr&agrave; cancellata solo l'associazione e non la marca o la categoria. Per cancellare completamente " + _
							"anche la marca o la categoria utilizzare le apposite sezioni."
		Class_delete.MsgSql = " SELECT (ISNULL((SELECT mar_nome_it FROM gtb_marche WHERE mar_id=srel_problemi_mar_tip.rpm_marchio_id),'Tutte le marche') + ' / ' + " & _ 
							  " 	ISNULL((SELECT tip_nome_it FROM gtb_tipologie WHERE tip_id=srel_problemi_mar_tip.rpm_tipologia_id),'Tutti le categorie')) AS NOMINATIVO " & _
							  " FROM srel_problemi_mar_tip "
		Class_delete.Name_Field = "NOMINATIVO"
		Class_delete.ID_Field = "rpm_id"
		Class_delete.Table = "srel_problemi_mar_tip"
		Class_delete.Caption = "Gestione marche/categorie"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = FALSE
	case "ARTICOLI"
		sql = "SELECT COUNT(*) FROM gtb_articoli WHERE art_se_accessorio=1 AND art_id=" & Class_delete.ID_value
		Class_delete.Message = "Cancellare il modello <RECORD>?"
		Class_delete.Name_Field = " art_cod_int " + SQL_concat(Class_delete.conn) + "' - '" + SQL_concat(Class_delete.conn) + " art_nome_it "
		Class_delete.ID_Field = "art_id"
		Class_delete.Table = "gtb_articoli"
		Class_delete.Caption = "Gestione modelli"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "CATEGORIE"
		Class_delete.Message = "Cancellare la categoria <RECORD>?"
		Class_delete.Name_Field = "tip_nome_it"
		Class_delete.ID_Field = "tip_id"
		Class_delete.Table = "gtb_tipologie"
		Class_delete.Caption = "Categorie"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "CTECH"
		Class_delete.Message = "Cancellare la caratteristica <RECORD>?"
		Class_delete.Name_Field = "ct_nome_it"
		Class_delete.ID_Field = "ct_id"
		Class_delete.Table = "gtb_carattech"
		Class_delete.Caption = "CARATTERISTICHE"
		Class_delete.AfterDelete = FALSE
	case "CTECH_GRUPPI"
		Class_delete.Message = "Cancellare il gruppo di caratteristiche <RECORD>?"
		Class_delete.Name_Field = "ctr_titolo_it"
		Class_delete.ID_Field = "ctr_id"
		Class_delete.Table = "gtb_carattech_raggruppamenti"
		Class_delete.Caption = "GRUPPI DI CARATTERISTICHE"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "MARCHE"
		Class_delete.Message = "Cancellare la marca <RECORD>?"
		Class_delete.Name_Field = "mar_nome_it"
		Class_delete.ID_Field = "mar_id"
		Class_delete.Table = "gtb_marche"
		Class_delete.Caption = "Marchi"
		Class_delete.AfterDelete = FALSE
	case "STATIO"
		' Class_delete.Message = "Cancellare la causale <RECORD>?"
		' Class_delete.Name_Field = "cau_titolo_it"
		' Class_delete.ID_Field = "cau_id"
		' Class_delete.Table = "sgtb_ddt_causali"
		' Class_delete.Caption = "causale DDT"
		' Class_delete.AfterDelete = FALSE
	case "CLIENTI"
		dim rs, profilo, profiloS
		set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT pro_nome_it FROM gtb_profili WHERE pro_id = (SELECT riv_profilo_id FROM gtb_rivenditori WHERE riv_id = "&Class_delete.ID_value&")"
		profilo = GetValueList(Class_delete.conn, NULL, sql)
		sql = Replace(sql, "pro_nome_it", "pro_codice")
		profiloS = GetValueList(Class_delete.conn, NULL, sql)
		
		Class_delete.Message = "Cancellare il "&profiloS&" <RECORD>?"
		Class_delete.Name_Field = "SELECT ModoRegistra FROM tb_indirizzario INNER JOIN tb_utenti "& _
								  "ON tb_indirizzario.IDElencoIndirizzi = tb_utenti.ut_nextCom_ID "& _
								  "WHERE tb_utenti.ut_id=gtb_rivenditori.riv_id"
		Class_delete.ID_Field = "riv_id"
		Class_delete.Table = "gtb_rivenditori"
		Class_delete.Caption = "Gestione " & lCase(profilo)
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "AGENTI"
		Class_delete.Message = "Cancellare l'agente <RECORD>?"
		Class_delete.Note = "ATTENZIONE: Non verranno cancellati i clienti dell'agente, ne tutti i dati degli ordini e delle operazioni del cliente"
		'permette la cancellazione dell'utente dell'area amministrativa se non usa nessun'altra applicazione
		sql = "SELECT COUNT(*) FROM rel_admin_sito WHERE admin_id IN (SELECT ag_admin_id FROM gtb_agenti WHERE ag_id=" & Class_delete.ID_value & ") "
		if GetValueList(Class_delete.conn, NULL, sql)=1 then
			'utente admin con accesso solo all'applicazione B2B: permette la cancellazione
			Class_delete.AddOption "delete_admin", "cancella anche l'utente per l'accesso all'area amministrativa.", true, ""
		else
			'utente admin con accesso anche ad altre applicazioni: non cancellabile
			Class_delete.Note = Class_delete.Note + ", l'utente per l'accesso all'area amministrativa perch&egrave; utilizzato per l'accesso ad altre altre applicazioni"
		end if
		'permette la cancellazione dell'utente dell'area riservata se non usa nessun'altra applicazione
		sql = "SELECT COUNT(*) FROM rel_utenti_sito WHERE rel_ut_id = " & Class_delete.ID_value
		if GetValueList(Class_delete.conn, NULL, sql)=1 then
			'l'agente ha accesso solo all'area riservata per l'inserimento ordini del B2B: permette la cancellazione
			Class_delete.AddOption "delete_utente", "cancella anche l'utente per l'accesso all'area riservata.", true, ""
			
			'permette la cancellazione del contatto se &egrave; bloccato solo dall'applicazione corrente
			sql = "SELECT LockedByApplication FROM tb_indirizzario WHERE IDElencoIndirizzi IN (SELECT ut_nextCom_id FROM tb_utenti WHERE ut_id=" & Class_delete.ID_value & ")"
			'response.write sql
			if GetValueList(Class_delete.conn, NULL, sql)=1 then
				'contatto dell'agente bloccato solo dall'applicazione corrente: permette la cancellazione
				Class_delete.AddOption "delete_contatto", "cancella anche il contatto associato", true, ""
			else				
				'contatto non cancellabile perch&egrave; bloccato da altre applicazioni
				Class_delete.Note = Class_delete.Note + " ed il contatto associato perch&egrave; utilizzato anche in altre applicazioni."
			end if
		else
			'utente dell'area riservata non cancellabile perch&egrave; ha accesso ad altre sezioni dell'area riservata: blocca anche la cancellazione del contatto
			Class_delete.Note = Class_delete.Note + ", il contatto e l'utente dell'area riservata perch&egrave; utilizzato per l'accesso ad altre sezioni riservate."
		end if
		Class_delete.Note = Class_delete.Note + "<br>Per eliminare i dati residui fare riferimento alle relative aree amministrative."
		Class_delete.Name_Field = "SELECT ModoRegistra FROM tb_indirizzario INNER JOIN tb_utenti "& _
								  "ON tb_indirizzario.IDElencoIndirizzi = tb_utenti.ut_nextCom_ID "& _
								  "WHERE ut_id=ag_id" 
		Class_delete.ID_Field = "ag_id"
		Class_delete.Table = "gtb_agenti"
		Class_delete.Caption = "AGENTI"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "ACCESSORIO"
		Class_delete.Message = "Cancellare l'accessorio <RECORD>?"
		Class_delete.Name_Field = "acc_nome_it"
		Class_delete.ID_Field = "acc_id"
		Class_delete.Table = "sgtb_accessori"
		Class_delete.Caption = "Gestione accessori"
		Class_delete.AfterDelete = FALSE
	case "ESITO"
		Class_delete.Message = "Cancellare l'esito <RECORD>?"
		Class_delete.Name_Field = "esi_nome_it"
		Class_delete.ID_Field = "esi_id"
		Class_delete.Table = "sgtb_esiti"
		Class_delete.Caption = "Gestione esiti interventi"
		Class_delete.AfterDelete = FALSE
	case "STATIS"
		Class_delete.Message = "Cancellare lo stato <RECORD>?"
		Class_delete.Name_Field = "sts_nome_it"
		Class_delete.ID_Field = "sts_id"
		Class_delete.Table = "sgtb_stati_schede"
		Class_delete.Caption = "STATI SCHEDE"
		Class_delete.AfterDelete = FALSE
	case "SCHEDE"
		Class_delete.Message = "Cancellare la scheda n.<RECORD>?"
		Class_delete.Name_Field = " CONVERT(nvarchar(100), sc_numero) + ' del ' + CONVERT(nvarchar, sc_data_ricevimento) "
		Class_delete.ID_Field = "sc_id"
		Class_delete.Table = "sgtb_schede"
		Class_delete.Caption = "SCHEDE"
		Class_delete.AfterDelete = TRUE
	case "DETTAGLI_SCHEDE"
		Class_delete.Message = "Cancellare il dettaglio con riferimento al ricambio <RECORD>?"
		Class_delete.Name_Field = " dts_ricambio_nome + ' (cod. ' + dts_ricambio_codice + ')'"
		Class_delete.ID_Field = "dts_id"
		Class_delete.Table = "sgtb_dettagli_schede"
		Class_delete.Caption = "DETTAGLI SCHEDE"
		Class_delete.AfterDelete = FALSE
	case "DETTAGLI_DDT"
		Class_delete.Message = "Cancellare il dettaglio con riferimento all'articolo <RECORD>?"
		Class_delete.Name_Field = " dtd_articolo_nome + ' (cod. ' + dtd_articolo_codice + ')'"
		Class_delete.ID_Field = "dtd_id"
		Class_delete.Table = "sgtb_dettagli_ddt"
		Class_delete.Caption = "DETTAGLI SPEDIZIONE"
		Class_delete.AfterDelete = FALSE	
	case "DESCRITTORI"
		Class_delete.Message = "Cancellare il controllo <RECORD>?"
		Class_delete.Name_Field = "des_nome_it"
		Class_delete.ID_Field = "des_id"
		Class_delete.Table = "sgtb_descrittori"
		Class_delete.Caption = "Controlli schede"
		Class_delete.AfterDelete = FALSE
	case "DESCRAG"
		Class_delete.Message = "Cancellare il gruppo di controlli <RECORD>?"
		Class_delete.Name_Field = "rag_titolo_it"
		Class_delete.ID_Field = "rag_id"
		Class_delete.Table = "sgtb_descrittori_raggruppamenti"
		Class_delete.Caption = "Gruppi caratteristiche controlli"
		Class_delete.AfterDelete = FALSE
	case "SPEDIZIONI"
		Class_delete.Message = "Cancellare il record relativo alla spedizione <RECORD>?"
		Class_delete.Name_Field = " CONVERT(nvarchar, ddt_numero) + ' del ' + CONVERT(nvarchar, ddt_data) "
		Class_delete.ID_Field = "ddt_id"
		Class_delete.Table = "sgtb_ddt"
		Class_delete.Caption = "Spedizioni"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "RITIRI"
		Class_delete.Message = "Cancellare il record relativo al ritiro <RECORD>?"
		Class_delete.Name_Field = " CONVERT(nvarchar, ddt_numero) + ' del ' + CONVERT(nvarchar, ddt_data) "
		Class_delete.ID_Field = "ddt_id"
		Class_delete.Table = "sgtb_ddt"
		Class_delete.Caption = "Ritiri"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "CLIENTI_INDIRIZZI"
		Class_delete.Message = "Cancellare l'indirizzo <RECORD>?"
		Class_delete.Name_Field = "NomeOrganizzazioneElencoIndirizzi"
		Class_delete.ID_Field = "IDElencoIndirizzi"
		Class_delete.Table = "tb_indirizzario"
		Class_delete.Caption = "INDIRIZZI ALTERNATIVI"
		Class_delete.AfterDelete = FALSE
	case "OPERATORI_INT_CENTRI"
		Class_delete.Message = "Cancellare l'operatore <RECORD>?"
		Class_delete.Name_Field = "CognomeElencoIndirizzi + ' ' + NomeElencoIndirizzi "
		Class_delete.ID_Field = "IDElencoIndirizzi"
		Class_delete.Table = "tb_indirizzario"
		Class_delete.Caption = "OPERATORI"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
end select

'definizione eventuali operazioni su relazioni	
Sub Delete_Relazioni(conn, ID)
	dim sql, rs, rsr, cnt_id, ComId, ObjContatto
	set rs = Server.CreateObject("ADODB.Recordset")
	Select case request.Querystring("SEZIONE")
		case "PROBLEMI"
			sql = "DELETE FROM srel_problemi_mar_tip WHERE rpm_problema_id = " & ID
			conn.Execute(sql)

		case "ARTICOLI"
		
			set rs = Server.CreateObject("ADODB.Recordset")
			set rsr = Server.CreateObject("ADODB.Recordset")
			
			'se un articolo "confezione" o "bundle" libera gli articoli in bundle / confezione
			sql = "SELECT * FROM gv_articoli WHERE art_id=" & ID
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			if not rs.eof then	'se non e' un articolo padre
				if NOT rs("art_varianti") AND _
				   (rs("art_se_bundle") OR  rs("art_se_confezione")) then
					'l'articolo e' in bundle: sblocco i componenti
					sql = " SELECT * FROM gtb_bundle INNER JOIN gv_articoli ON gtb_bundle.bun_articolo_id = gv_articoli.rel_id " + _
						  " WHERE bun_bundle_id=" & rs("rel_id")
					rsr.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
					while not rsr.eof
						sql = " SELECT COUNT(*) FROM gtb_bundle " + _
							  " WHERE bun_articolo_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & rsr("art_id") & ") " & _
							  " AND bun_bundle_id<>" & rs("rel_id")
						if cInteger(GetValueList(conn, NULL, sql))=0 then
							if rs("art_se_bundle") then
								sql = "UPDATE gtb_Articoli SET art_in_bundle=0 WHERE art_id=" & rsr("art_id")
								CALL conn.execute(sql, , adExecuteNoRecords)
							elseif rs("art_se_confezione") then
								sql = "UPDATE gtb_Articoli SET art_in_confezione=0 WHERE art_id=" & rsr("art_id")
								CALL conn.execute(sql, , adExecuteNoRecords)
							end if
						end if
						rsr.movenext
					wend
					rsr.close
					
					sql = "DELETE FROM gtb_bundle WHERE bun_bundle_id=" & rs("rel_id")
					CALL conn.execute(sql, , adExecuteNoRecords)
				end if
			else
				rs.close
				sql = "SELECT * FROM gtb_articoli WHERE art_id=" & ID
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			end if
			
			if rs("art_ha_accessori") then
				'l'articolo ha accessori: sblocco gli accessori
				sql = " SELECT * FROM grel_art_acc WHERE aa_art_id=" & ID
				rsr.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				while not rsr.eof
					sql = "SELECT COUNT(*) FROM grel_art_acc WHERE aa_acc_id=" & rsr("aa_acc_id") & " AND aa_art_id<>" & ID
					if cInteger(GetValueList(conn, NULL, sql))=0 then
						sql = "UPDATE gtb_articoli SET art_se_accessorio=0 WHERE art_id=" & rsr("aa_acc_id")
						CALL conn.execute(sql, , adExecuteNoRecords) 
					end if
					rsr.movenext
				wend
				rsr.close
				
				sql = "DELETE FROM grel_art_acc WHERE aa_art_id=" & ID
				CALL conn.execute(sql, , adExecuteNoRecords)
			end if
			
			if rs("art_se_accessorio") then
				'l'articolo &egrave; accessorio di un altro articolo: libera relazione.
				sql = " SELECT * FROM grel_art_acc WHERE aa_acc_id=" & ID
				rsr.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				while not rsr.eof
					sql = "SELECT COUNT(*) FROM grel_art_acc WHERE aa_art_id=" & rsr("aa_art_id") & " AND aa_acc_id<>" & ID
					if cInteger(GetValueList(conn, NULL, sql))=0 then
						sql = "UPDATE gtb_articoli SET art_ha_accessori=0 WHERE art_id=" & rsr("aa_art_id")
						CALL conn.execute(sql, , adExecuteNoRecords) 
					end if
					rsr.movenext
				wend
				rsr.close
			end if
			rs.close
						
			set rs = nothing
			set rsr = nothing
			
		case "CATEGORIE"
			categorie.Delete(ID)
			
		case "CTECH_GRUPPI"
			sql = "UPDATE gtb_carattech SET ct_raggruppamento_id = NULL WHERE ct_raggruppamento_id=" & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
			
		case "CLIENTI"
			'cancella shopping cart collegate:
			sql = "DELETE FROM gtb_shopping_cart WHERE sc_riv_id=" & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
		
			'recupera id contatto
			ComID = GetValueList(conn, NULL, "SELECT ut_nextCom_id FROM tb_utenti WHERE ut_id="& ID)
			
			'verifica opzioni scelte dall'utente per la cancellazione
			sql = "SELECT COUNT(*) FROM rel_utenti_sito WHERE rel_ut_id = " & ID
			if CInteger(GetValueList(Class_delete.conn, NULL, sql))<=1 then
				'cancella contatto ed utente
				sql = "DELETE FROM tb_indirizzario WHERE IDElencoIndirizzi=" & ComID
				CALL conn.execute(sql, , adExecuteNoRecords)
			else
				set ObjContatto = new IndirizzarioLock
				set ObjContatto.conn = conn
				
				'il contatto rimane sempre registrato
				if request("delete_utente")<>"" then
					'rimuove utente
					CALL ObjContatto.RemoveUserFormContact(ComID, ID, NEXTPASSPORT)
				else
					'rimuove permesso di accesso all'utente.
					'CALL ObjContatto.UserAbilitazione_Remove(ComID, ID, UTENTE_PERMESSO_CLIENTE)
					
					'blocca il contatto dal next-passport per l'utente
					CALL ObjContatto.LockContact(ComID, NEXTPASSPORT)
				end if
				
				CALL ObjContatto.AddToRubrica(ComID, session("RUBRICA_EX_CLIENTI"))
				CALL ObjContatto.RemoveFromRubrica(ComID, session("RUBRICA_CLIENTI"))
				
                ObjContatto.conn = empty
				set ObjContatto = nothing
			end if
			
		case "AGENTI"
			set rs = Server.CreateObject("ADODB.Recordset")
			set rsr = Server.CreateObject("ADODB.Recordset")
			
			'recupera dati dell'agente
			sql = " SELECT * FROM gtb_agenti INNER JOIN tb_utenti ON gtb_agenti.ag_id = tb_utenti.ut_id " + _
				  " INNER JOIN tb_indirizzario ON tb_utenti.ut_NextCom_id = tb_indirizzario.IDElencoIndirizzi " + _
				  " WHERE ag_id=" & ID
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			
			'aggiorna dati della rubrica associata all'agente per rogliere l'associazione
			sql = "SELECT * FROM tb_rubriche WHERE syncroFilterTable LIKE 'gtb_agenti' AND syncroFilterKey="& ID
			rsr.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			rsr("nome_rubrica") = "Ex Clienti agente " + rs("CognomeElencoIndirizzi") + " " + rs("NomeElencoIndirizzi")
			rsr("locked_rubrica") = false
			rsr("rubrica_esterna") = false
			rsr("syncroFilterKey") = NULL
			rsr.update
			rsr.close
			
			'cancella gruppo di lavoro dell'agente
			sql = "DELETE FROM tb_gruppi WHERE id_gruppo="& rs("ag_gruppo_id")
			CALL conn.execute(sql, , adExecuteNoRecords)
			
			'cancellazione dell'utente amministratore dell'agente
			if request("delete_admin")<>"" then
				'cancella utente
				sql = "DELETE FROM tb_admin WHERE id_admin=" & rs("ag_admin_id")
				CALL conn.execute(sql, , adExecuteNoRecords)
			else
				'toglie permesso di accesso all'applicativo per l'utente amministratore
				sql = " DELETE FROM rel_admin_sito WHERE admin_id=" & rs("ag_admin_id") & " AND sito_id=" & NEXTB2B & _
					  " AND rel_as_permesso=" & POS_PERMESSO_AGENTE
				CALL conn.execute(sql, , adExecuteNoRecords)
			end if
			
			ComID = GetValueList(conn, NULL, "SELECT ut_nextCom_id FROM tb_utenti WHERE ut_id="& ID)
			'verifica opzioni scelte dall'utente per la cancellazione
			if request("delete_contatto")<>"" then
				'cancella contatto ed utente
				sql = "DELETE FROM tb_indirizzario WHERE IDElencoIndirizzi=" & ComID
				CALL conn.execute(sql, , adExecuteNoRecords)
			else
				set ObjContatto = new IndirizzarioLock
				set ObjContatto.conn = conn
				
				'il contatto rimane registrato
				if request("delete_utente")<>"" then
					'cancella solo l'utente ma lascia il contatto
					CALL ObjContatto.RemoveUserFormContact(ComID, ID, NEXTPASSPORT)
					
				else
					'rimuove permesso di accesso all'utente.
					'CALL ObjContatto.UserAbilitazione_Remove(ComID, ID, UTENTE_PERMESSO_AGENTE)
					
					'blocca il contatto dal next-passport per l'utente
					CALL ObjContatto.LockContact(ComID, NEXTPASSPORT)
				end if
				
				CALL ObjContatto.AddToRubrica(ComID, session("RUBRICA_EX_AGENTI"))
				CALL ObjContatto.RemoveFromRubrica(ComID, session("RUBRICA_AGENTI"))
				
                ObjContatto.conn = empty
				set ObjContatto = nothing
				
				'sgancia rivenditori da agente.
				sql = "UPDATE gtb_rivenditori SET riv_agente_id=NULL WHERE riv_agente_id=" & ID
				CALL conn.execute(sql, , adExecuteNoRecords)
			end if
			
			set rs = nothing
			set rsr = nothing
		case "SCHEDE"
			'cancella i dettagli
			sql = "DELETE FROM sgtb_dettagli_schede WHERE dts_scheda_id = " & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
		case "SPEDIZIONI"
			sql = "UPDATE sgtb_schede SET sc_costo_riconsegna = 0 WHERE sc_rif_DDT_di_resa_id = " & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
			sql = "UPDATE sgtb_schede SET sc_rif_DDT_di_resa_id = 0 WHERE sc_rif_DDT_di_resa_id = " & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
			'cancella i dettagli
			sql = "DELETE FROM sgtb_dettagli_ddt WHERE dtd_ddt_id = " & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
		case "RITIRI"
			sql = "UPDATE sgtb_schede SET sc_costo_presa = 0 WHERE sc_documento_ritiro_id = " & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
			sql = "UPDATE sgtb_schede SET sc_documento_ritiro_id = 0 WHERE sc_documento_ritiro_id = " & ID
			CALL conn.execute(sql, , adExecuteNoRecords)
		case "OPERATORI_INT_CENTRI"
			sql = "DELETE FROM tb_admin WHERE id_admin IN (SELECT ut_admin_id FROM tb_utenti WHERE ut_NextCom_id = " & ID & ")"
			CALL conn.execute(sql, , adExecuteNoRecords)
	end select
end Sub

Sub Operations_AfterDelete(conn, ID)	
	dim sql
	Select case request.Querystring("SEZIONE")
	end select
end sub

Class_delete.Delete_Manager()
%>

</body>
</html>