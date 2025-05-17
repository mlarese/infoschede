<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 10000 %>
<% response.buffer = true %>
<!--#INCLUDE FILE="../library/ClassDelete.asp" -->
<!--#INCLUDE FILE="../library/ClassIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<!--#INCLUDE FILE="Tools4Save_B2B.asp" -->
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
	case "SPESESPEDIZIONEARTICOLO"
		Class_delete.Message = "Cancellare la modalit&agrave; di spedizione dell'articolo <RECORD>?"
		Class_delete.Name_Field = "spa_nome_IT"
		Class_delete.ID_Field = "spa_id"
		Class_delete.Table = "gtb_spese_spedizione_articolo"
		Class_delete.Caption = "Gestione modalit&agrave; spedizione articolo"
		Class_delete.AfterDelete = FALSE
	case "ARTICOLICOMMENTI"
		Class_delete.Message = "Cancellare il commento <RECORD>?"
		Class_delete.Name_Field = "com_comment"
		Class_delete.ID_Field = "com_id"
		Class_delete.Table = "tb_comments"
		Class_delete.Caption = "Commenti articolo"
		Class_delete.AfterDelete = FALSE
	case "ARTICOLIFAQ"
		Class_delete.Message = "Cancellare il collegamento alla faq <RECORD> per questo articolo?"
		Class_delete.Name_Field = "faq_domanda_IT"
		Class_delete.Name_Field = "(SELECT faq_domanda_IT FROM tb_faq WHERE faq_id = grel_art_faq.raf_faq_id)"
		Class_delete.ID_Field = "raf_id"
		Class_delete.Table = "grel_art_faq"
		Class_delete.Caption = "FAQ articolo"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
		Class_delete.AddOption "delete_faq", "cancella anche la faq associata", false, ""
	case "MARCHE"
		Class_delete.Message = "Cancellare la marca <RECORD>?"
		Class_delete.Name_Field = "mar_nome_it"
		Class_delete.ID_Field = "mar_id"
		Class_delete.Table = "gtb_marche"
		Class_delete.Caption = "Marchi"
		Class_delete.AfterDelete = FALSE
	case "SCONTIQ"
		Class_delete.Message = "Cancellare la classe sconto quantità <RECORD>?"
		Class_delete.Name_Field = "scc_nome"
		Class_delete.ID_Field = "scc_id"
		Class_delete.Table = "gtb_scontiQ_classi"
		Class_delete.Caption = "Classi di sconto per quantit&agrave;"
		Class_delete.AfterDelete = FALSE
	case "SCONTIQ_D"
		Class_delete.Message = "Cancellare l'intervallo di sconto a partire da <RECORD>?"
		Class_delete.Name_Field = "sco_qta_da"
		Class_delete.ID_Field = "sco_id"
		Class_delete.Table = "gtb_scontiQ"
		Class_delete.Caption = "Intervalli di sconto"
		Class_delete.AfterDelete = TRUE
	case "VARIANTI"
		Class_delete.Message = "Cancellare la variante <RECORD>?"
		Class_delete.Name_Field = "var_nome_it"
		Class_delete.ID_Field = "var_id"
		Class_delete.Table = "gtb_varianti"
		Class_delete.Caption = "Varianti"
		Class_delete.AfterDelete = FALSE
	case "VALORI"
		Class_delete.Message = "Cancellare il valore <RECORD>?"
		Class_delete.Name_Field = "val_nome_it"
		Class_delete.ID_Field = "val_id"
		Class_delete.Table = "gtb_valori"
		Class_delete.Caption = "Valori delle varianti"
		Class_delete.AfterDelete = FALSE
	case "CATEGORIE"
		Class_delete.Message = "Cancellare la categoria <RECORD>?"
		Class_delete.Name_Field = "tip_nome_it"
		Class_delete.ID_Field = "tip_id"
		Class_delete.Table = "gtb_tipologie"
		Class_delete.Caption = "Categorie"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "RAGGRUPPAMENTO"
		Class_delete.Message = "Cancellare il raggruppamento <RECORD>?"
		Class_delete.Name_Field = "rag_nome_it"
		Class_delete.ID_Field = "rag_id"
		Class_delete.Table = "gtb_tipologie_raggruppamenti"
		Class_delete.Caption = "Raggruppamenti"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = FALSE
	case "VALUTE"
		Class_delete.Message = "Cancellare la valuta <RECORD>?"
		Class_delete.Name_Field = "valu_nome"
		Class_delete.ID_Field = "valu_id"
		Class_delete.Table = "gtb_valute"
		Class_delete.Caption = "Valute"
		Class_delete.AfterDelete = FALSE
	case "FATTURAZIONE"
		Class_delete.Message = "Cancellare la tipologia di fatturazione <RECORD>?"
		Class_delete.Name_Field = "fatt_codice"
		Class_delete.ID_Field = "fatt_id"
		Class_delete.Table = "gtb_fatturazioni"
		Class_delete.Caption = "Fatturazione"
		Class_delete.AfterDelete = FALSE
	case "CLIENTI"
		Class_delete.Message = "Cancellare il cliente <RECORD>?"
		Class_delete.Note = "ATTENZIONE: Non verranno cancellati i listini, i codici articolo personalizzati"
		'permette la cancellazione dell'utente se non usa nessun'altra applicazione
		sql = "SELECT COUNT(*) FROM rel_utenti_sito WHERE rel_ut_id = " & Class_delete.ID_value
		if CInteger(GetValueList(Class_delete.conn, NULL, sql))<=1 then
			'il cliente ha accesso solo all'area riservata del B2B: permette la cancellazione
			Class_delete.AddOption "delete_utente", "cancella anche l'utente associato", true, ""
			
			'permette la cancellazione del contatto se &egrave; bloccato solo dal next-Passport
			sql = "SELECT LockedByApplication FROM tb_indirizzario WHERE IDElencoIndirizzi IN (SELECT ut_nextCom_id FROM tb_utenti WHERE ut_id=" & Class_delete.ID_value & ")"
			if CInteger(GetValueList(Class_delete.conn, NULL, sql))=1 then
				'contatto bloccato solo dall'applicazione corrente: permette la cancellazione
				Class_delete.AddOption "delete_contatto", "cancella anche il contatto associato", true, ""
			else
				'contatto non cancellabile perch&egrave; bloccato da altre applicazioni
				Class_delete.Note = Class_delete.Note + " ed il contatto associato perch&egrave; utilizzato anche in altre applicazioni."
			end if
		else
			'utente non cancellabile perch&egrave; ha accesso ad altre applicazioni: blocca anche la cancellazione del contatto
			Class_delete.Note = Class_delete.Note + ", il contatto e l'utente dell'area riservata perch&egrave; utilizzato per l'accesso ad altre sezioni riservate."
		end if
		Class_delete.Note = Class_delete.Note + "<br>Per eliminare i dati residui fare riferimento alle relative aree amministrative."
		Class_delete.Name_Field = "SELECT ModoRegistra FROM tb_indirizzario INNER JOIN tb_utenti "& _
								  "ON tb_indirizzario.IDElencoIndirizzi = tb_utenti.ut_nextCom_ID "& _
								  "WHERE tb_utenti.ut_id=gtb_rivenditori.riv_id"
		Class_delete.ID_Field = "riv_id"
		Class_delete.Table = "gtb_rivenditori"
		Class_delete.Caption = "CLIENTI"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "CLIENTI_INDIRIZZI"
		Class_delete.Message = "Cancellare l'indirizzo <RECORD>?"
		Class_delete.Name_Field = "NomeOrganizzazioneElencoIndirizzi"
		Class_delete.ID_Field = "IDElencoIndirizzi"
		Class_delete.Table = "tb_indirizzario"
		Class_delete.Caption = "INDIRIZZI ALTERNATIVI"
		Class_delete.AfterDelete = FALSE
	case "CLIENTI_PROFILI"
		Class_delete.Message = "Cancellare il profilo <RECORD>?"
		Class_delete.Name_Field = "pro_nome_it"
		Class_delete.ID_Field = "pro_id"
		Class_delete.Table = "gtb_profili"
		Class_delete.Caption = "Profili"
		Class_delete.AfterDelete = TRUE
		Class_delete.DeleteRelations = FALSE
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
	case "LISTINO"
		Class_delete.Message = "Cancellare il listino <RECORD>?"
		Class_delete.Name_Field = "listino_codice"
		Class_delete.ID_Field = "listino_id"
		Class_delete.Table = "gtb_listini"
		Class_delete.Caption = "LISTINI"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "LISTA_CODICI"
		Class_delete.Message = "Cancellare la lista di codici <RECORD>?"
		Class_delete.Name_Field = "LstCod_nome"
		Class_delete.ID_Field = "LstCod_id"
		Class_delete.Table = "gtb_lista_codici"
		Class_delete.Caption = "liste codici"
		Class_delete.AfterDelete = FALSE
	case "ORDINI"
		Class_delete.Message = "Cancellare l'ordine <RECORD>?"
		Class_delete.Name_Field = "'n.' + CAST(ord_id AS nvarchar(20)) + ' del ' + CAST(CONVERT(DATETIME, ord_data, 102) AS nvarchar(11))"
		Class_delete.ID_Field = "ord_id"
		Class_delete.Table = "gtb_ordini"
		Class_delete.Caption = "ORDINI"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "STATIO"
		Class_delete.Message = "Cancellare lo stato <RECORD>?"
		Class_delete.Name_Field = "so_nome_it"
		Class_delete.ID_Field = "so_id"
		Class_delete.Table = "gtb_stati_ordine"
		Class_delete.Caption = "STATI ORDINE"
		Class_delete.AfterDelete = FALSE
	case "ARTICOLI"
		sql = "SELECT COUNT(*) FROM gtb_articoli WHERE art_se_accessorio=1 AND art_id=" & Class_delete.ID_value
		if cInteger(GetValueList(Class_delete.conn, NULL, sql))>0 then
			Class_delete.Message = "<strong>ATTENZIONE:</strong>L'articolo &egrave; accessorio o prodotto correlato di un altro articolo.<br>" + _
								   "Cancellare comunque l'articolo <RECORD>?"
		else
			Class_delete.Message = "Cancellare l'articolo <RECORD>?"
		end if
		Class_delete.Name_Field = " art_cod_int " + SQL_concat(Class_delete.conn) + "' - '" + SQL_concat(Class_delete.conn) + " art_nome_it "
		Class_delete.ID_Field = "art_id"
		Class_delete.Table = "gtb_articoli"
		Class_delete.Caption = "Gestione articoli"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "ARTICOLI_ACCESSORI"
		Class_delete.Message = "Cancellare il collegamento con l'articolo <RECORD>?"
		Class_Delete.Note = "Verr&agrave; cancellata solo collegamento e non l'articolo. Per cancellare completamente " + _
							"anche l'articolo utilizzare l'apposita sezione."
		Class_delete.MsgSql = " SELECT (art_cod_int " + SQL_concat(Class_delete.conn) + "' - '" + SQL_concat(Class_delete.conn) + " art_nome_it) AS NOMINATIVO FROM " + _
							  " gtb_articoli INNER JOIN grel_art_acc ON gtb_articoli.art_id = grel_art_acc.aa_acc_id "
		Class_delete.Name_Field = "NOMINATIVO"
		Class_delete.ID_Field = "aa_id"
		Class_delete.Table = "grel_art_acc"
		Class_delete.Caption = "Gestione articoli collegati"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "ARTICOLI_COMPONENTI"
		Class_delete.Message = "Cancellare il componente <RECORD>?"
		Class_Delete.Note = "Verr&agrave; cancellata solo l'associazione bundle-componente e non l'articolo. Per cancellare completamente " + _
							"anche l'articolo utilizzare l'apposita sezione."
		Class_delete.MsgSql = " SELECT (art_cod_int " + SQL_concat(Class_delete.conn) + "' - '" + SQL_concat(Class_delete.conn) + " art_nome_it) AS NOMINATIVO FROM " + _
							  " gv_articoli INNER JOIN gtb_bundle ON gv_articoli.rel_id = gtb_bundle.bun_articolo_id "
		Class_delete.Name_Field = "NOMINATIVO"
		Class_delete.ID_Field = "bun_id"
		Class_delete.Table = "gtb_bundle"
		Class_delete.Caption = "Gestione componenti "
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = TRUE
	case "ARTICOLI_FOTO"
		CALL oArticoliFoto.SetDeleteSettings(Class_delete, "Gestione articoli - foto")
	case "ARTICOLI_VARIANTI"
		sql = " SELECT (gtb_varianti.var_nome_it " + SQL_concat(Class_delete.conn) + "':<strong>'" + _
			  SQL_concat(Class_delete.conn) + "gtb_valori.val_nome_it " + SQL_concat(Class_delete.conn) + "'</strong>') AS NOMINATIVO " + _
			  " FROM grel_art_vv INNER JOIN gtb_valori ON grel_art_vv.rvv_val_id=gtb_valori.val_id " + _
			  " INNER JOIN gtb_varianti ON gtb_valori.val_var_id = gtb_varianti.var_id " + _
			  " WHERE rvv_art_var_id=" & cIntero(request("ID"))
		Class_delete.Message = "Cancellare la variante " + GetValueList(Class_delete.conn, NULL, sql) + " con codice <RECORD>?"
		Class_delete.Name_Field = "rel_cod_int"
		Class_delete.ID_Field = "rel_id"
		Class_delete.Table = "grel_art_valori"
		Class_delete.Caption = "Gestione varianti "
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "DETORD"
		Class_delete.Message = "Cancellare l'articolo <RECORD> dall'ordine?"
		Class_delete.Name_Field = "(CASE WHEN IsNull(det_art_var_id,0)>0 THEN (SELECT rel_cod_int FROM grel_art_valori WHERE rel_id=det_art_var_id) ELSE det_descr_it END)"
		Class_delete.ID_Field = "det_id"
		Class_delete.Table = "gtb_dettagli_ord"
		Class_delete.Caption = "ARTICOLI ORDINE"
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = TRUE
	case "CARICHI"
		Class_delete.Message = "Cancellare il carico non movimentato <RECORD>?"
		Class_delete.Name_Field = "car_fornitore_cod"
		Class_delete.ID_Field = "car_id"
		Class_delete.Table = "gtb_carichi"
		Class_delete.Caption = "Carichi"
		Class_delete.AfterDelete = FALSE
	case "MAGAZZINI"
		Class_delete.Message = "Cancellare il magazzino vuoto <RECORD>?"
		Class_delete.Name_Field = "mag_nome"
		Class_delete.ID_Field = "mag_id"
		Class_delete.Table = "gtb_magazzini"
		Class_delete.Caption = "Magazzino"
		Class_delete.AfterDelete = FALSE
	case "CAT_IVA"
		Class_delete.Message = "Cancellare la categoria I.V.A. <RECORD>?"
		Class_delete.Name_Field = "iva_nome + ' (' + CAST(iva_valore AS nvarchar(4)) + '%)'"
		Class_delete.ID_Field = "iva_id"
		Class_delete.Table = "gtb_iva"
		Class_delete.Caption = "categorie i.v.a."
		Class_delete.AfterDelete = FALSE
    case "ORDINI_TIPOLOGIE_RIGHE"
		Class_delete.Message = "Cancellare la tipologia <RECORD>?"
		Class_delete.Name_Field = "dot_nome_it"
		Class_delete.ID_Field = "dot_id"
		Class_delete.Table = "gtb_dettagli_ord_tipo"
		Class_delete.Caption = "tipologie di righe d'ordine"
		Class_delete.AfterDelete = FALSE
    case "ORDINI_INFO_RIGHE"
		Class_delete.Message = "Cancellare l'informazione <RECORD>?"
		Class_delete.Name_Field = "dod_nome_it"
		Class_delete.ID_Field = "dod_id"
		Class_delete.Table = "gtb_dettagli_ord_des"
		Class_delete.Caption = "informazioni per riga d'ordine"
		Class_delete.AfterDelete = FALSE
	case "SPESESPEDIZIONE"
		Class_delete.Message = "Cancellare la modalit&agrave; di spedizione <RECORD>?"
		Class_delete.Name_Field = "sp_area_nome_it"
		Class_delete.ID_Field = "sp_id"
		Class_delete.Table = "gtb_spese_spedizione"
		Class_delete.Caption = "Modalit&agrave; spedizione ordine"
		Class_delete.AfterDelete = FALSE
	case "MODPAGA"
		Class_delete.Message = "Cancellare la modalit&agrave; di pagamento <RECORD>?"
		Class_delete.Name_Field = "mosp_nome_it"
		Class_delete.ID_Field = "mosp_id"
		Class_delete.Table = "gtb_modipagamento"
		Class_delete.Caption = "Modalit&agrave; di pagamanto"
		Class_delete.AfterDelete = FALSE
	case "PORTI"
		Class_delete.Message = "Cancellare il porto <RECORD>?"
		Class_delete.Name_Field = "prt_nome_it"
		Class_delete.ID_Field = "prt_id"
		Class_delete.Table = "gtb_porti"
		Class_delete.Caption = "Porti"
		Class_delete.AfterDelete = FALSE
	case "TIPI_CONSEGNA"
		Class_delete.Message = "Cancellare il tipo consegna <RECORD>?"
		Class_delete.Name_Field = "tco_nome_it"
		Class_delete.ID_Field = "tco_id"
		Class_delete.Table = "gtb_tipo_consegna"
		Class_delete.Caption = "Tipi consegna"
		Class_delete.AfterDelete = FALSE
	case "TRASPORTATORI"
		Class_delete.Message = "Cancellare il trasportatore <RECORD>?"
		Class_delete.Name_Field = "tra_nome_it"
		Class_delete.ID_Field = "tra_id"
		Class_delete.Table = "gtb_trasportatori"
		Class_delete.Caption = "Trasportatori"
		Class_delete.AfterDelete = FALSE
end select

'definizione eventuali operazioni su relazioni	
Sub Delete_Relazioni(conn, ID)
	dim rs, rsr, sql, ComId, ObjContatto, Applicazione, Permesso
	select case request.Querystring("SEZIONE")
		case "ARTICOLIFAQ"
			'cancellazione faq collegata
			if request("delete_faq") <> "" then
				'cancellare la faq
				sql = "DELETE FROM tb_FAQ WHERE faq_id IN (SELECT raf_faq_id FROM grel_art_faq WHERE raf_id = " & ID & ")"
				CALL conn.execute(sql, ,adExecuteNoREcords)
			end if
			
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
			if request("delete_contatto")<>"" then
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
					CALL ObjContatto.UserAbilitazione_Remove(ComID, ID, UTENTE_PERMESSO_CLIENTE)
					
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
			if not rsr.eof then
				rsr("nome_rubrica") = "Ex Clienti agente " + rs("CognomeElencoIndirizzi") + " " + rs("NomeElencoIndirizzi")
				rsr("locked_rubrica") = false
				rsr("rubrica_esterna") = false
				rsr("syncroFilterKey") = NULL
				rsr.update
			end if
			rsr.close
			
			'cancella gruppo di lavoro dell'agente
			sql = "DELETE FROM tb_gruppi WHERE id_gruppo="& cIntero(rs("ag_gruppo_id"))
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
					CALL ObjContatto.UserAbilitazione_Remove(ComID, ID, UTENTE_PERMESSO_AGENTE)
					
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
		case "CATEGORIE"
		
			categorie.Delete(ID)
		
		case "ORDINI"
		
			'ricalcolo giacenza se impegna o movimenta
			set rs = conn.execute("SELECT ord_impegna, ord_movimenta, ord_magazzino_id FROM gtb_ordini WHERE ord_id="& ID)
			if rs("ord_impegna") then
				CALL SetGiacenza_ord(conn, ID, "O", "-", "I", rs("ord_magazzino_id"))
			elseif rs("ord_movimenta") then
				CALL SetGiacenza_ord(conn, ID, "O", "+", "M", rs("ord_magazzino_id"))
			end if
		
		case "DETORD"
		
			'ricalcolo giacenza se ordine impegna o movimenta
			set rs = conn.execute(" SELECT TOP 1 ord_impegna, ord_movimenta, ord_magazzino_id, det_art_var_id " + _
								  " FROM gtb_dettagli_ord d INNER JOIN gtb_ordini o ON d.det_ord_id=o.ord_id WHERE det_id="& ID)
			if rs("det_art_var_id")>0 then
				if rs("ord_impegna") then
					CALL SetGiacenza_ord(conn, ID, "D", "-", "I", rs("ord_magazzino_id"))
				elseif rs("ord_movimenta") then
					CALL SetGiacenza_ord(conn, ID, "D", "+", "M", rs("ord_magazzino_id"))
				end if
			end if
			set rs = nothing
			
		case "ARTICOLI_ACCESSORI"
			set rs = Server.CreateObject("ADODB.Recordset")
			set rsr = Server.CreateObject("ADODB.Recordset")
			
			sql = "SELECT * FROM grel_art_acc WHERE aa_id=" & ID
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			
			'ricalcola stato articolo di cui &egrave; accessorio
			sql = "SELECT COUNT(*) FROM grel_art_acc WHERE aa_art_id=" & rs("aa_art_id") & " AND aa_id<>" & ID
			if cInteger(GetValueList(conn, rsr, sql))=0 then
				'aggiorna stato articolo: non ha pi&ugrave; accessori
				sql = "UPDATE gtb_articoli SET art_ha_accessori=0 WHERE art_id=" & rs("aa_art_id")
				CALL conn.execute(sql, , adExecuteNoRecords)
			end if
			
			'ricalcola stato accessorio
			sql = "SELECT COUNT(*) FROM grel_art_acc WHERE aa_acc_id=" & rs("aa_acc_id") & " AND aa_id<>" & ID
			if cInteger(GetValueList(conn, rsr, sql))=0 then
				'aggiorna stato articolo: non &egrave; pi&ugrave; accessorio di nessun articolo
				sql = "UPDATE gtb_articoli SET art_se_accessorio=0 WHERE art_id=" & rs("aa_acc_id")
				CALL conn.execute(sql, , adExecuteNoRecords)
			end if
			
			rs.close
			set rs = nothing
			set rsr = nothing
		
		case "ARTICOLI_COMPONENTI"
			set rs = Server.CreateObject("ADODB.Recordset")
			set rsr = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT art_id, rel_id, bun_articolo_id, bun_bundle_id, art_se_bundle FROM gv_articoli INNER JOIN gtb_bundle ON gv_articoli.rel_id = gtb_bundle.bun_bundle_id WHERE bun_id=" & ID
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			
			'salva l'id del bundle di cui aggiornare le giacenze
			Session("DELETE_BUN_ID") = rs("bun_bundle_id")
			
			'disabilita bundle/confezione se non ha pi&ugrave; componenti
			sql = "SELECT COUNT(*) FROM gtb_bundle WHERE bun_bundle_id=" & rs("bun_bundle_id") & " AND bun_id<>" & ID
			if cInteger(GetValueList(conn, rsr, sql)) = 0 then	'non ha piu' componenti
				sql = "UPDATE gtb_articoli SET art_disabilitato=1 WHERE art_id=" & rs("art_id")
				CALL conn.execute(sql, , adExecuteNoRecords)
				'azzera le quantit&agrave;
				sql = "UPDATE grel_giacenze SET gia_qta=0, gia_impegnato=0, gia_ordinato=0 WHERE gia_art_var_id=" & rs("rel_id")
				CALL conn.execute(sql, , adExecuteNoRecords)
			end if
						
			'ricalcola stato dell'articolo componente.
			sql = "SELECT COUNT(*) FROM gtb_bundle WHERE bun_articolo_id=" & rs("bun_articolo_id") & " AND bun_id<>" & ID
			if cInteger(GetValuelist(conn, rsr, sql)) = 0 then	' non e' piu' componente di nessun bundle
				sql = " UPDATE gtb_articoli SET " + IIF(rs("art_se_bundle"), "art_in_bundle", "art_in_confezione") + "=0 WHERE " + _
					  " art_id IN (SELECT rel_art_id FROM grel_art_valori WHERE rel_id=" & rs("bun_articolo_id") & ")"
				CALL conn.execute(sql, , adExecuteNoRecords)
			end if
			rs.close
			set rs = nothing
			set rsr = nothing
		
		case "ARTICOLI_VARIANTI"
			
            dim oVar
            set oVar = new GestioneVariante
            set oVar.conn = conn
            CALL oVar.BeforeDelete(ID)
		
		case "ARTICOLI_FOTO"
			'effettua operazioni di base per l'oggetto di gestione della gallery
			CALL oArticoliFoto.OnDeleteRelazioni(conn, ID)
		
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
		case "LISTINO"
			set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT listino_with_child, listino_ancestor_id FROM gtb_listini WHERE listino_id=" & ID
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			if rs("listino_with_child") then
				'sgancia tutti i listini derivati dal listino corrente (listino principale)
				sql = "UPDATE gtb_listini SET listino_ancestor_id=NULL WHERE listino_ancestor_id=" & ID
				CALL conn.execute(Sql, , adCmdText)
			elseif cInteger(rs("listino_ancestor_id")) then
				'ricalcola l'evantuale listino padre
				sql = " UPDATE gtb_listini SET listino_with_child=CASE WHEN (SELECT COUNT(*) FROM gtb_listini L_child " + _
					  " WHERE L_child.listino_ancestor_id = gtb_listini.listino_id)>1 THEN 1 ELSE 0 END " + _
					  " WHERE gtb_listini.listino_id=" & rs("listino_ancestor_id")
				CALL conn.execute(sql, , adCmdText)
			end if
			rs.close
			
			set rs = nothing
	end select
end Sub



Sub Operations_AfterDelete(conn, ID)
	dim rs, sql
	select case request.Querystring("SEZIONE")
		case "ARTICOLI_COMPONENTI"
			  set rs = Server.CreateObject("ADODB.Recordset")
			  
			'ricalcola tutte le giacenze a magazzino
			sql = "SELECT mag_id FROM gtb_magazzini"
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			while not rs.eof
				CALL SetGiacenza_b(conn, Session("DELETE_BUN_ID"), QTA_IMPEGNATA, rs("mag_id"))
				CALL SetGiacenza_b(conn, Session("DELETE_BUN_ID"), QTA_ORDINATA, rs("mag_id"))
				CALL SetGiacenza_b(conn, Session("DELETE_BUN_ID"), QTA_GIACENZA, rs("mag_id"))
				rs.movenext
			wend
			rs.close
			set rs = nothing
		
		case "ARTICOLI_FOTO"
			'effettua operazioni di finali per l'oggetto di gestione della gallery
			CALL oArticoliFoto.OnAfterDelete(conn, ID)
		
		case "SCONTI_Q"
			'resetto le relazioni
			sql = "UPDATE grel_art_valori SET rel_scontoQ_id = NULL WHERE rel_scontoQ_id="& ID
			CALL conn.Execute(sql, , adExecuteNoRecords)
			
			sql = "UPDATE gtb_prezzi SET prz_scontoQ_id = NULL WHERE prz_scontoQ_id="& ID
			CALL conn.Execute(sql, , adExecuteNoRecords)
			
			sql = "UPDATE gtb_articoli SET art_scontoQ_id = NULL WHERE art_scontoQ_id=" & ID
			CALL conn.Execute(sql, , adExecuteNoRecords)
		
		case "CLIENTI_PROFILI"
            'rimuove rubrica collegata al profilo
            CALL DeleteSyncroRubrica(conn, "gtb_profili", "gtb_profili", ID)
	end select
end sub

Class_delete.Delete_Manager()
%>

</body>
</html>