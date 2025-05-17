<!--#INCLUDE FILE="Tools_ClassCssManager.asp"-->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp"-->
<%
'posizione dei parametri del NEXTweb nella tabella tb_siti
const WEB_ADMIN_POS = 1		'sito_p1
const WEB_POWER_POS = 2		'sito_p2
const WEB_USER_POS =  3		'sito_p3

'tag per la definizione dei path degli URL e delle immagini da passare alla dynalay
Private const tagPath = "<@PATH>"					'es.: http://sviluppo.next-aim.local/dynalay.asp
Private const tagPathResources = "<@PATH_RES>"		'es.: http://sviluppo.next-aim.local/upload/1
%>

<%
'Riporto su questo file alcune funzioni che erano presenti su Tools_NextWeb5.asp.
'Questo perchè le seguenti funzioni mi servono su Update__library__framework_core.asp, 
'e includendo Tools_NextWeb5.asp mi dava un errore dovuto alla "doppia inclusione" (il nome ... è ridefinito)
%>
<!--#INCLUDE FILE="Tools_NextWeb5_extension.asp"-->

<%
'***************************************************************************************************
'***************************************************************************************************
'		DEFINIZIONE FUNZIONI
'***************************************************************************************************
'***************************************************************************************************
'....................................................................................................
'		funzione che imposta le propriet&agrave; del sito richiesto per request(<ID del sito>)
'....................................................................................................
sub Imposta_Proprieta_Sito(parametroSito)
	dim conn, rs, sql
	if request(parametroSito)<>"" then
		set conn = Server.CreateObject("ADODB.Connection")
		set rs = Server.CreateObject("ADODB.recordSet")
		conn.open Application("DATA_ConnectionString"), "", ""
		sql = "SELECT * FROM tb_webs WHERE id_webs=" & cIntero(request(parametroSito))
		rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		Session("AZ_ID") = rs("id_webs")
		Session("NOME_WEBS") = rs("nome_webs")
		Session("SITO_MOBILE") = rs("sito_mobile")
		Session("LINGUA_IT") = true
		Session("LINGUA_EN") = rs("lingua_EN")
		Session("LINGUA_FR") = rs("lingua_FR")
		Session("LINGUA_DE") = rs("lingua_DE")
		Session("LINGUA_ES") = rs("lingua_ES")
		if FieldExists(rs, "lingua_RU") then
			Session("LINGUA_RU") = rs("lingua_RU")
		else
			Session("LINGUA_RU") = false
		end if
		if FieldExists(rs, "lingua_CN") then
			Session("LINGUA_CN") = rs("lingua_CN")
		else
			Session("lingua_CN") = false
		end if
		if FieldExists(rs, "LINGUA_PT") then
			Session("LINGUA_PT") = rs("LINGUA_PT")
		else
			Session("LINGUA_PT") = false
		end if
			
		'gestisce conteggio e lista di lingue attive
		Session("LINGUE_ATTIVE") = 1
		Session("LINGUE") = LINGUA_ITALIANO
		if rs("lingua_EN") then
			Session("LINGUE_ATTIVE") = Session("LINGUE_ATTIVE") + 1
			Session("LINGUE") = Session("LINGUE") + " " + LINGUA_INGLESE
		end if
		if rs("lingua_FR") then
			Session("LINGUE_ATTIVE") = Session("LINGUE_ATTIVE") + 1
			Session("LINGUE") = Session("LINGUE") + " " + LINGUA_FRANCESE
		end if
		if rs("lingua_DE") then
			Session("LINGUE_ATTIVE") = Session("LINGUE_ATTIVE") + 1
			Session("LINGUE") = Session("LINGUE") + " " + LINGUA_TEDESCO
		end if
		if rs("lingua_ES") then
			Session("LINGUE_ATTIVE") = Session("LINGUE_ATTIVE") + 1
			Session("LINGUE") = Session("LINGUE") + " " + LINGUA_SPAGNOLO
		end if
		if FieldExists(rs, "lingua_RU") then
			if rs("lingua_RU") then
				Session("LINGUE_ATTIVE") = Session("LINGUE_ATTIVE") + 1
				Session("LINGUE") = Session("LINGUE") + " " + LINGUA_RUSSO
			end if
		end if
		if FieldExists(rs, "lingua_CN") then
			if rs("lingua_CN") then
				Session("LINGUE_ATTIVE") = Session("LINGUE_ATTIVE") + 1
				Session("LINGUE") = Session("LINGUE") + " " + LINGUA_CINESE
			end if
		end if
		if FieldExists(rs, "lingua_PT") then
			if rs("lingua_PT") then
				Session("LINGUE_ATTIVE") = Session("LINGUE_ATTIVE") + 1
				Session("LINGUE") = Session("LINGUE") + " " + LINGUA_PORTOGHESE
			end if
		end if
		Session("LINGUE") = split(Session("LINGUE"), " ")
		
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	else
		if cIntero(Session("AZ_ID")) = 0 then
			response.redirect "Siti.asp"
		end if
	end if
end sub


sub Reset_Proprieta_Sito()
	Session("AZ_ID") = ""
	Session("NOME_WEBS") = ""
	Session("LINGUE_ATTIVE") = 1
	Session("LINGUA_IT") = true
	Session("LINGUA_EN") = false
	Session("LINGUA_FR") = false
	Session("LINGUA_DE") = false
	Session("LINGUA_ES") = false
	Session("LINGUA_RU") = false
	Session("LINGUA_CN") = false
	Session("LINGUA_PT") = false
	
	'resetto tutte le impostazioni di paginazione delle pagine interne
	dim pager
	set pager = new PageNavigator
	pager.ResetAll
	set pager = nothing
end sub




'....................................................................................................
'		funzione che copia una pagina in un'altra copiando anche i layer
'		conn:			connessione attiva al database
'		page_source: 	id della pagina sorgente
'		page_dest:		id della pagina di destinazione
'       add_layers:     se true indica che i layers della pagina di destinazione non devono essere
'                       cancellati, ma devono essere aggiunti solo i layers della pagina sorgente
'....................................................................................................
function Copy_page(conn, page_source, page_dest, add_layers)
	page_source = cIntero(page_source)
	page_dest = cIntero(page_dest)
	if page_source <> page_dest then
		dim rs_s, rs_d, sql, field, template
		set rs_s = server.createObject("ADODB.Recordset")
		set rs_d = server.createObject("ADODB.Recordset")
		template = CBoolean(GetValueList(conn, rs_s, "SELECT template FROM tb_pages WHERE id_page = "& page_dest), false)
		
		sql = "SELECT * FROM tb_pages WHERE id_page=" 
		
		'aggiorna record pagina
		rs_s.open sql & page_source, conn, adOpenStatic, adLockOptimistic, adCmdText
		'rs_d.open sql & page_dest, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		'rs_d("id_template") = rs_s("id_template")
		'rs_d("sfondoColore") = rs_s("sfondoColore")
		'rs_d("SfondoImmagine") = rs_s("SfondoImmagine")
		'rs_d.update
		
		sql = " UPDATE tb_pages SET " + _
			  " id_template=" & IIF(cInteger(rs_s("id_template"))>0 AND NOT template, rs_s("id_template"), "NULL") & ", " + _
			  " sfondoColore= '" & ParseSql(rs_s("sfondoColore"), adChar) & "', " + _
			  " SfondoImmagine= '" & ParseSql(rs_s("SfondoImmagine"), adChar) & "' " & _
			  " WHERE id_page=" & page_dest
		CALL conn.execute(sql, 0, adExecuteNoRecords)
		
		rs_s.close
		'rs_d.close
		
		'cancella layers pagina di destinazione
        if not add_layers then
    		sql = "DELETE FROM tb_layers WHERE id_pag=" & page_dest
	    	CALL conn.execute(sql, 0, adExecuteNoRecords)
        end if
			
		'aggiunge layers pagina di destinazione
		sql = "INSERT INTO tb_layers(id_pag, id_tipo, z_order, nome, visibile, x," &_
			  " y, largo, alto, html, format, testo, aspcode, id_objects, tipo_contenuto, em_x, em_y, em_largo, em_alto,"& _
			  " rtf, checksum_stili)" &_
			  " SELECT " & page_dest & ", id_tipo, z_order, nome, visibile, x, y, largo," &_
			  " alto, html, format, testo, aspcode, id_objects, tipo_contenuto, em_x, em_y, em_largo, em_alto,"& _
			  " rtf, checksum_stili"& _
			  " FROM tb_layers " &_
			  " WHERE id_pag=" & page_source
		CALL conn.execute(sql, 0, adExecuteNoRecords)
		
		'aggiorna data di modifica paginasito
		CALL UpdateDataModifica(conn, page_dest)
		
		'controllo plugin nella pagina copiata (page_dest) in caso di copia tra due pagine con diverso id_webs
		dim id_webs_source, id_webs_dest, sql_execute, id_plugin
		id_webs_source = cIntero(GetValueList(conn, NULL, "SELECT id_webs FROM tb_pages WHERE id_page=" & page_source))
		id_webs_dest = cIntero(GetValueList(conn, NULL, "SELECT id_webs FROM tb_pages WHERE id_page=" & page_dest))
		if id_webs_source <> id_webs_dest then
			dim rs
			set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM tb_layers WHERE id_tipo=" & LAYER_OBJECT & " AND id_pag=" & page_dest
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			'scorro tutti i layers di tipo object (i plugin) presenti nella pagina appena copiata
			sql_execute = ""
			id_plugin = 0
			while not rs.eof
				sql = "SELECT TOP 1 id_objects FROM tb_objects WHERE id_webs = " & id_webs_dest & " AND name_objects LIKE '" & rs("nome") & "'"
				id_plugin = cIntero(GetValueList(conn, NULL, sql))
				if id_plugin = 0 then
					'non è istanziato il plugin per l'id_webs_dest, quindi lo creo
					sql_execute = _
							" INSERT INTO tb_objects(id_webs,name_objects,identif_objects,param_list,obj_insData,obj_insAdmin_id,obj_type) " & _
							" SELECT "& id_webs_dest &",name_objects,identif_objects,param_list,"& SQL_Now(conn) &"," & Session("ID_ADMIN") & ",obj_type " & _
							" FROM tb_objects WHERE id_webs = " & id_webs_source & " AND id_objects = " & rs("id_objects")
					CALL conn.execute(sql_execute, 0, adExecuteNoRecords)
					id_plugin = cIntero(GetValueList(conn, NULL, sql))	
				end if
				'il plugin è istanziato, allora cambio l'id di riferimento sul layer della pagina appena creata
				sql_execute = " UPDATE tb_layers SET id_objects=" & id_plugin & " WHERE id_lay = " & rs("id_lay")
				CALL conn.execute(sql_execute, 0, adExecuteNoRecords)
				rs.moveNext
			wend
			rs.close
			set rs = nothing
			
			if not template then
				'inoltre se è una pagina cancello il riferimento al template sulla pagina copiata
				sql = "UPDATE tb_pages SET id_template = NULL WHERE id_webs = " & id_webs_dest & " AND id_page = " & page_dest
				CALL conn.execute(sql, 0, adExecuteNoRecords)
			end if
		end if
		
		set rs_s = nothing
		set rs_d = nothing
	end if
end function



'....................................................................................................
'		funzione che crea una nuova pagina copiandola da una già esistente
'		conn:				connessione attiva al database
'		PsSorgenteId: 		id della pagina sorgente
'		PsDestinazioneId:	id della pagina di destinazione
'       copiaTutto:     	se true copia tutti i campi della tabella tb_pagineSito di origine
'....................................................................................................
function CopiaPaginaSito(conn, PsSorgenteId, PsDestinazioneId, copiaTutto)
	dim rs_sorg, rs_dest, sql, field, i, lingua
	set rs_sorg = Server.CreateObject("ADODB.Recordset")
	set rs_dest = Server.CreateObject("ADODB.Recordset")


	'apre il record della pagina sorgente
	sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & PsSorgenteId
	rs_sorg.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	
	'apre il record della pagina di destinazione
	sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & PsDestinazioneId
	rs_dest.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	
	if not rs_sorg.eof then
		if rs_dest.eof then
			rs_dest.addnew
			CALL SetUpdateParamsRS(rs_dest, "ps_", true)   'inserisce la data di modifica
		else
			CALL SetUpdateParamsRS(rs_dest, "ps_", false)   'modifica data di modifica
		end if
	
		'copia i campi dalla pagina sorgente alla pagina di destinazione
		for each field in rs_sorg.fields
			if copiaTutto then   'campi copiati nel caso di un import dati
				if (InStr(1, field.name, "id_pag", 1)<1) then
					rs_dest(field.name)=rs_sorg(field.name)
				end if
			end if
			if InStr(1, field.name, "PAGE_", 1)<>0 then  'campi copiati in ogni caso
				rs_dest(field.name)=rs_sorg(field.name)
			end if
		next
	
		rs_dest.update   'aggiorna la tabella di destinazione
		PsDestinazioneId = rs_dest("id_pagineSito")
		rs_dest.close
		
		sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & PsDestinazioneId
		rs_dest.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		CALL Ceck_page_exists(conn, rs_dest)
	
		'copia la pagina di lavoro e la pagina pubblica per ogni lingua
		for i = 0 to uBound(application("LINGUE"))
			lingua = Application("LINGUE")(i)		
			CALL Copy_page(conn, rs_sorg("id_pagStage_"& lingua), rs_dest("id_pagStage_"& lingua), false)
			CALL Copy_page(conn, rs_sorg("id_pagDyn_"& lingua), rs_dest("id_pagDyn_"& lingua), false)
		next

		CopiaPaginaSito = PsDestinazioneId
	else
		CopiaPaginaSito = 0
	end if
	
	rs_sorg.close
	rs_dest.close
	
	set rs_sorg = nothing
	set rs_dest = nothing
end function




'....................................................................................................
'		funzione che copia un menu da un altro menu
'....................................................................................................
function Copy_MenuFromMenu(conn, menu_dest, menu_source)
	dim sql
	'copia del menu: recupera elenco colonne della tabella
	sql = TableFieldList(conn, NULL, "tb_menuItem", "mi_id mi_menu_id")
	
	'compone query ed esegue copia
	sql = "INSERT INTO tb_MenuItem (mi_menu_id, " + sql + ") " + _
		  "SELECT " & menu_dest & ", " + sql + " FROM tb_menuItem WHERE mi_menu_id=" & menu_source
	CALL conn.execute(sql, , adExecuteNoRecords)
End Function


'....................................................................................................
'		funzione che copia un menu da una voce dell'indice
'....................................................................................................
function Copy_MenuFromIndex(conn, menu_dest, index_source, abilitaFigli)
	dim sql
	sql = " INSERT INTO tb_menuItem (" + _
          "     mi_menu_id, " +_
          "     mi_titolo_it, mi_titolo_en, mi_titolo_fr, mi_titolo_es, mi_titolo_de, mi_titolo_ru, mi_titolo_cn, mi_titolo_pt," + _
          "     mi_ordine, " + _
          "     mi_index_id, " + _
          "     mi_attivo, " + _
		  " 	mi_figli )" + _
          " SELECT " & _
          "     " & menu_dest & ", " + _
          "     co_titolo_it, co_titolo_en, co_titolo_fr, co_titolo_es, co_titolo_de, co_titolo_ru, co_titolo_cn, co_titolo_pt," + _
          "     " + SQL_IfIsNull(conn, "idx_ordine", "co_ordine") + ", " + _
          "     idx_id, " + _
          "     " + SQL_IfIsNull(conn, "idx_visibile_assoluto", "co_visibile") + _
		  "		, " + IIF(abilitaFigli, "1", "0") + _
          " FROM v_indice WHERE idx_padre_id = "& index_source
	CALL conn.execute(sql, , adExecuteNoRecords)
End Function


'....................................................................................................
'		funzione che verifica se la pagina e' pubblicata o se deve essere aggiornata
'		conn:			connessione attiva al database
'		id_pagStage: 	id della pagina di lavoro
'		id_pagDyn:		id della pagina visibile al pubblico
'....................................................................................................
function must_be_published(conn, rs, id_pagStage, id_pagDyn)
	dim sql, template_diversi, min_lay_DYN, max_lay_STAGE, lay_STAGE, lay_DYN
	if cInteger(id_pagDyn) >0 then
		'recupera dati sulla situazione dei layer della pagina di lavoro
		sql = "SELECT (COUNT(id_lay)) AS N_LAY, (MAX(id_lay)) AS MAX_LAYER " &_
			  " FROM tb_layers WHERE id_pag=" & cIntero(id_pagStage)
		rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		if not rs.eof then
			lay_STAGE = rs("N_LAY")
			max_lay_STAGE = rs("MAX_LAYER")
		else
			lay_STAGE = 0
			max_lay_STAGE = 0
		end if
		rs.close

		'recupera dati sulla situazione dei layer della pagina pubblicata
		sql = "SELECT (COUNT(id_lay)) AS N_LAY, (MIN(id_lay)) AS MIN_LAYER " &_
			  " FROM tb_layers WHERE id_pag=" & id_pagDyn
		rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		if not rs.eof then
			lay_DYN = rs("N_LAY")
			min_lay_DYN = rs("MIN_LAYER")
		else
			lay_DYN = 0
			min_lay_DYN = 0
		end if
		rs.close
				
		'recupera dati sulla situazione dei template delle pagine di stage e pubblicate
		sql = " SELECT DISTINCT " & SQL_IfIsNull(conn, "id_template", 0) & " FROM tb_pages " + _
              " WHERE id_page=" & cIntero(id_pagStage) & " OR id_page=" & id_pagDyn
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText	'se trova piu' di un record i template sono diversi
		template_diversi = rs.recordcount>1
		rs.close
		
		'response.write "TEMPLATE:"& template_diversi &";MIN:"& min_lay_DYN &";MAX:"& max_lay_STAGE &";DYN:"& lay_DYN &";STAGE:"& lay_STAGE &";"
		must_be_published = (template_diversi) OR (min_lay_DYN < max_lay_STAGE) OR (lay_DYN <> lay_STAGE)
	else
		must_be_published = true
	end if
	
end function


'....................................................................................................
'		funzione che aggiorna la data di modifica della paginaSito a partire dalla pagina
'		conn:			connessione attiva al database
'		page_id			id della pagina modificata
'....................................................................................................
function UpdateDataModifica(conn, page_id)
	dim sql, i
	sql = "(SELECT id_paginaSito FROM tb_pages WHERE id_page=" & page_id & ")"
	CALL UpdateParams(conn, "tb_pagineSito", "ps_", "id_pagineSito", sql, false)
	
	CALL UpdateParams(conn, "tb_pages", "page_", "id_page", page_id, false)
end function



'....................................................................................................
'		funzione che aggiorna data di ultima modifica della struttura dei plugin del sito
'		conn:			connessione attiva al database
'		webs_id			id del sito i cui plugin sono stati variati
'....................................................................................................
sub UpdateSitoDataModificaPlugin(conn, webs_id)
    dim sql
    sql = "UPDATE tb_webs SET webs_modData_plugin = " & SQL_Now(conn) & " WHERE id_webs=" & webs_id
    CALL conn.execute(sql, ,adExecuteNoRecords)
end sub 


'....................................................................................................
'		funzione che aggiorna data di ultima modifica del sito
'		conn:			connessione attiva al database
'		webs_id			id del sito i cui plugin sono stati variati
'....................................................................................................
sub UpdateSitoDataModifica(conn, webs_id)
    dim sql
    sql = "UPDATE tb_webs SET webs_modData = " & SQL_Now(conn) & " WHERE id_webs=" & webs_id
    CALL conn.execute(sql, ,adExecuteNoRecords)
end sub 


'....................................................................................................
'		funzione che ripulisce la stringa secondo i vincoli di utilizzo dei caratteri del next-web
'		se preserveLenght mantiene la lunghezza totale della stringa sostituendo ai caratteri non validi
'		degli spazi.
'		ritorna la stringa ripulita.
'....................................................................................................
function ClearString(str, PreserveLenght)
	dim c, i, CheckedStr
	CheckedStr = ""	
	if cString(str)<>"" then
		for i=1 to len(str)
			c = Mid(str, i, 1)
			if isValidEditorChar(c) then
				CheckedStr = CheckedStr & c
			elseif PreserveLenght then
				CheckedStr = CheckedStr & " "
			end if
		next
	end if
	ClearString = CheckedStr
end function


'.................................................................................................
'				Controlla se il carattere &egrave; valido per la gestione dell'editor
'				c :		carattere da controllare
'.................................................................................................
function isValidEditorChar(c)
	dim Code
	isValidEditorChar = false
	if instr(1, EDITOR_BASE_CHARSET, c, vbTextCompare)>0 then	'caratteri di base
		isValidEditorChar = true
	else 
		Code = AscB(c)		'set esteso di caratteri
		if   Code = 10 OR Code = 13 OR _
		   ( Code > 31 AND Code < 127) OR _
		     Code = 128 OR Code = 142 OR _
		   ( Code > 129 AND Code < 141) OR _
		   ( Code > 144 AND Code < 157) OR _
		     Code > 157 then
			isValidEditorChar = true
		end if
	end if
end function


'restituisce la query per i dati della paginaSito data la pagina
'campi:				elenco di campi della tabella da includere nell'espressione SELECT
'pagID:				ID della pagina di cui si vuole la corrispettiva paginaSito
Function QryPagineSito(campi, pagID)
	dim i, lingua
	pagID = CIntero(pagID)
	QryPagineSito = " SELECT "& campi & " FROM tb_pagineSito WHERE id_pagDyn_it = "& pagID &" OR id_pagStage_it = "& pagID
	for i=lbound(Application("LINGUE"))+1 to ubound(Application("LINGUE"))
		lingua = Application("LINGUE")(i)
		QryPagineSito = QryPagineSito & " OR id_pagDyn_" & lingua & "=" & pagID
		QryPagineSito = QryPagineSito & " OR id_pagStage_" & lingua & "=" & pagID
	next
	'QryPagineSito = " SELECT "& campi &" FROM tb_pagineSito " & _
	'				" WHERE id_pagDyn_it = "& pagID &" OR id_pagDyn_en = "& pagID & _
	'				" OR id_pagDyn_fr = "& pagID &" OR id_pagDyn_es = "& pagID & _
	'				" OR id_pagDyn_de = "& pagID &" OR id_pagDyn_ru = "& pagID & _
	'				" OR id_pagDyn_cn = "& pagID &" OR id_pagDyn_pt = "& pagID &" OR id_pagStage_it = "& pagID & _
	'				" OR id_pagStage_en = "& pagID &" OR id_pagStage_fr = "& pagID & _
	'				" OR id_pagStage_es = "& pagID &" OR id_pagStage_de = "& pagID & _
	'				" OR id_pagStage_ru = "& pagID &" OR id_pagStage_cn = "& pagID &" OR id_pagStage_pt = "& pagID _
End Function


'restituisce la query per recuperare l'elenco completo dei template
'condition:		eventuale condizione aggiuntiva di filtro
Function QryElencoTemplate(condition, AddOrdine)
	dim copia_pagine_tra_siti
	copia_pagine_tra_siti = cBoolean(Session("COPIA_PAGINE_TRA_SITI"),false)
	QryElencoTemplate = _
		"SELECT id_page, " + _
			  " (" & IIF(copia_pagine_tra_siti, "tb_webs.nome_webs" & SQL_Concat(conn) & "' - '" & SQL_Concat(conn), "") & SQL_IF(conn, SQL_IsTrue(conn, SQL_IfIsNull(conn, "semplificata", "0")), "nomepage " & SQL_Concat(conn) & "' ( per email semplificate )'", "nomepage") & ") AS NAME " + _
		      IIF(AddOrdine, ", (1) AS ordine ", "") + _
			  " FROM tb_pages " + IIF(copia_pagine_tra_siti, " INNER JOIN tb_webs ON tb_pages.id_webs = tb_webs.id_webs ", "") + _
			  " WHERE "& SQL_IsTrue(conn, "template") & condition
	if not copia_pagine_tra_siti then
		if cIntero(Session("AZ_ID"))>0 then
			QryElencoTemplate = QryElencoTemplate + _
								" AND id_webs=" & Session("AZ_ID")
		end if
	end if
	QryElencoTemplate = QryElencoTemplate + _
			  " ORDER BY " & IIF(AddOrdine, "ordine, NAME", "nomepage")
end function


'esegue il check dei permessi di una pagina sito
Function ChkPrmPages(ID)
	dim PaginaSito
	PaginaSito = GetValueList(index.conn, NULL, QryPagineSito("id_pagineSito", ID))
	if PaginaSito > 0 then
		ChkPrmPages = index.content.ChkPrmF("tb_pagineSito", PaginaSito)
	else
		ChkPrmPages = true
	end if
End Function


'restituisce il nome della pagina completo di riferimento aggiuntivo
function PaginaSitoNome(rs, lingua) 
    lingua = IIF(cString(lingua)<>"", lingua, LINGUA_ITALIANO)
    PaginaSitoNome = rs("nome_ps_" + lingua)
    if cString(rs("nome_ps_interno"))<>"" then
        PaginaSitoNome = PaginaSitoNome + _
            "<span class=""note""> ( " + rs("nome_ps_interno") + " ) </span>"
    end if
end function


'restituisce il riepilogo dei numeri che compongono la pagina.
function GetPageNumbers(rs)
	dim lingua, i, Numbers
	Numbers = "pagina n&ordm; " & rs("id_pagineSito") & vbCrLf
	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
		lingua = Application("LINGUE")(i)
		if Session("LINGUA_" + lingua) then
			Numbers = Numbers & "lingua " &  GetNomeLingua(lingua) & ":" & vbCrLF & _
					  vbTab & "pagina di lavoro: n&ordm; " & vbTab & IIF(cInteger(rs("id_pagStage_" + lingua))>0, rs("id_pagStage_" + lingua), "n.d.") & vbCrLF & _
					  vbTab & "pagina pubblica: n&ordm; " & vbTab &  IIF(cInteger(rs("id_pagDyn_" + lingua))>0, rs("id_pagDyn_" + lingua), "n.d.") & vbCrLF
		end if
	next
	GetPageNumbers = Numbers
end function

%>
