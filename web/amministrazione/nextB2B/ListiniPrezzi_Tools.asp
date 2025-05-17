<%
'****************************************************************************************
'****************************************************************************************
'		ATTENZIONE:
'		l'aggiornamento dei listini va effettuato solo con la funzione salva riga
'		che integra la propagazione degli aggiornamenti ai listini derivati.
'		la cancellazione o l'aggiornamento possono essere eseguiti in qualsiasi modo
'		perche' gestiti da TRIGGER
'****************************************************************************************
'****************************************************************************************

const ARTICOLI_PER_PAGINA = 200

sub Listino_OpenRowRecordset(byref rsl, byref rs, listino, sql_where)
	sql = Listino_GetRowQuery(rsl, listino, sql_where)
	rs.open sql, conn, adOpenStatic, adLockReadOnly 
end sub


function Listino_GetRowQuery(byref rsl, listino, sql_where)
	dim sql
	sql = " SELECT rel_ordine, rel_id, rel_art_id, rel_cod_int, rel_prezzo, rel_scontoQ_id, " + _
		  " art_id, art_nome_it, art_cod_int, art_scontoQ_id, art_varianti, art_disabilitato, art_iva_id, " + _
		  " iva_id, iva_nome, iva_valore, " + _
		  " prz_id , prz_prezzo, prz_visibile, prz_promozione, prz_listino_id, prz_variante_id, prz_scontoQ_id, prz_iva_id, prz_var_euro, prz_var_sconto, " + _
		  " prz_offerta_dal, prz_offerta_al "
	if not rsl("listino_offerte") then
		'sql = sql + ", (SELECT COUNT(*) FROM gv_listino_offerte WHERE gv_listino_offerte.prz_variante_id=gv_listini.prz_variante_id) AS OFFERTA "
		dim id_listino_base, id_listino_offerte, conn
		set conn = Server.CreateObject("ADODB.Connection")
		conn.open Application("DATA_ConnectionString")
		id_listino_base = cIntero(GetValueList(conn, NULL, "SELECT listino_id FROM gtb_listini WHERE listino_base_attuale = 1"))
		id_listino_offerte = cIntero(GetValueList(conn, NULL, "SELECT listino_id FROM gtb_listini WHERE listino_offerte = 1"))
		sql = sql + ", (SELECT COUNT(*) FROM fn_listino_vendita_varianti("&id_listino_base&","&id_listino_offerte&","&cIntero(listino)&") f WHERE in_offerta = 1 AND f.rel_id = gv_listini.prz_variante_id) AS OFFERTA "		
	end if
	if not rsl("listino_base") then
		sql = sql + ", (SELECT prz_prezzo FROM gtb_prezzi WHERE gtb_prezzi.prz_variante_id=gv_listini.prz_variante_id " + _
			  		" AND gtb_prezzi.prz_listino_id=" & rsl("LB_ATTUALE") & ") AS PREZZO_BASE "
	end if
	sql = sql + " FROM gv_listini "
	if rsl("listino_base") then
		sql = sql & " WHERE prz_listino_id=" & cIntero(listino) & " " + sql_where + " ORDER BY rel_cod_int" 
	else
		sql = sql & " WHERE prz_listino_id=" & cIntero(listino) & " " + sql_where + _
					" UNION " + _
					sql & " WHERE (prz_listino_id =" & rsl("LB_ATTUALE") & " AND " + _
								 " prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_prezzi WHERE prz_listino_id=" & cIntero(listino) & ")) " + _
					sql_where + " ORDER BY rel_cod_int" 
	end if
	response.write "<!-- " & sql & "-->"
	Listino_GetRowQuery = sql
end function


sub Listino_Scheda(rsl, listino, spostamento)%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px; width:735px;">
		<caption>
			<% if spostamento then %>	
				<table align="right" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td align="right" style="font-size: 1px;">
							<a class="button" href="?ID=<%= listino %>&goto=PREVIOUS" title="listino precedente" <%= ACTIVE_STATUS %>>
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= listino %>&goto=NEXT" title="listino successivo" <%= ACTIVE_STATUS %>>
								SUCCESSIVO &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			<% end if %>
			Modifica dati del listino "<%= rsl("listino_codice") %>"
		</caption>
		<tr><th colspan="4">DATI DEL LISTINO</th></tr>	
		<tr>
			<td class="label">tipo</td>
			<td class="content">
				<% if rsl("listino_base") then 
					if rsl("listino_base_attuale") then%>
						<strong title="listino base attualmente in vigore">listino base</strong>
					<% else %>
						listino base
					<% end if
				elseif rsl("listino_offerte") then %>
					listino offerte speciali
				<% elseif rsl("listino_b2c") then %>
					listino al pubblico
				<% else
					if cInteger(rsl("listino_ancestor_id"))>0 then%>
						listino clienti derivato da <a href="ListiniMod.asp?ID=<%= rsl("listino_ancestor_id") %>" target="_blank"><%= GetValueList(rsl.ActiveConnection, NULL, "SELECT listino_codice FROM gtb_listini WHERE listino_id=" & rsl("listino_ancestor_id")) %></a>
					<% else %>
						listino clienti principale
					<% end if
				end if %>
			</td>
			<td class="label">stato / validit&agrave;</td>
			<td class="content">
				<% if rsl("listino_base_attuale") OR _
					  ((rsl("listino_offerte") OR rsl("listino_b2c")) AND _
					   (rsl("listino_DataCreazione") <= Date AND _
					    ((rsl("listino_DataScadenza") >= Date) OR isNull(rsl("listino_DataScadenza"))) )) then%>
					<strong>in vigore</strong>
					<% if rsl("listino_offerte") then %>
						<span class="Icona Offerte" title="listino offerte speciali in vigore">&nbsp;</span>
					<% end if %>
				<% else %>
					&nbsp;
				<% end if %>
			</td>
			<% if rsl("listino_offerte") then %>
				<tr>
					<td class="label">validit&agrave; offerte:</td>
					<td class="content" colspan="3">
						<% if isDate(rsl("listino_DataCreazione")) then%>
							dal <%= DateITA(rsl("listino_DataCreazione")) %>
						<% end if
						if isDate(rsl("listino_DataScadenza")) then %>
							fino al <%= DateITA(rsl("listino_DataScadenza")) %>
						<% end if%>
					</td>
				</tr>
			<% elseif rsl("listino_b2c") then %>
				<tr>
					<td class="label">pubblicazione:</td>
					<td class="content" colspan="3">visible al pubblico</td>
				</tr>
				<tr>
					<td class="label">validit&agrave; pubblicazione:</td>
					<td class="content" colspan="3">
						<% if isDate(rsl("listino_DataCreazione")) then%>
							dal <%= DateITA(rsl("listino_DataCreazione")) %>
						<% end if
						if isDate(rsl("listino_DataScadenza")) then %>
							fino al <%= DateITA(rsl("listino_DataScadenza")) %>
						<% end if%>
					</td>
				</tr>
			<% end if %>
		</tr>
	</table>
<% end sub

function Listino_StatoRiga(rs, rsl, listino, byref prezzo_base, _
											 byref prezzo_nuovo, _
											 byref prezzo_attuale, _
											 byref var_sconto, _
											 byref var_euro, _
											 byref personalizzato, _
											 byref iva_id, _
											 byref scontoQ_id, _
											 byref visibile, _
											 byref promozione, _
											 byref offerta_dal, _
											 byref offerta_al)
	if rs("prz_listino_id") = listino then
		'prezzo con variazione caricato dal listino corrente
		if rsl("listino_base") then		'listino base
			prezzo_base = rs("rel_prezzo")
		else							'altro listino
			prezzo_base = rs("PREZZO_BASE")
		end if
		prezzo_nuovo = rs("prz_prezzo")
		prezzo_attuale = rs("prz_prezzo")
		var_sconto = rs("prz_var_sconto")
		var_euro = rs("prz_var_euro")
		personalizzato = 1
	else
		'prezzo senza variazioni caricato dal lisitno base
		prezzo_base = rs("prz_prezzo")
		prezzo_nuovo = rs("prz_prezzo")
		prezzo_attuale = rs("prz_prezzo")
		var_sconto = 0
		var_euro = 0
		personalizzato = 0
	end if
	iva_id = rs("prz_iva_id")
	scontoQ_id = cInteger(rs("prz_scontoQ_id"))

	if rsl("listino_offerte") then
		if personalizzato>0 then
			visibile = rs("prz_visibile")
		else
			visibile = false
		end if
		promozione = false
		offerta_dal = rs("prz_offerta_dal")
		offerta_al = rs("prz_offerta_al")
	else
		visibile = rs("prz_visibile")
		promozione = rs("prz_promozione")
	end if
end function


function Listino_SalvaRiga(conn, rsl, rs, listino, variante_id, var_sconto, var_euro, iva_id, scontoQ_id, visibile, promozione, prezzo_nuovo, personalizzato, offerta_dal, offerta_al)
	dim prezzo_attuale, sql, child_sql
	dim ToBeSaved
	'se &egrave; un listino base deve essere salvata ogni riga
	ToBeSaved = rsl("listino_base")
	if not ToBeSaved then
		'verifica se e' presente una qualsiasi variazione
		ToBeSaved = (var_sconto<>0 OR var_euro<>0)
		if not ToBeSaved then	
			if scontoQ_id = 0 then scontoQ_id = NULL
		
			'verifica le altre impostazioni dal listino base
			sql = "SELECT * FROM gtb_prezzi WHERE prz_variante_id=" & cIntero(variante_id) & " AND prz_listino_id=" & rsl("LB_ATTUALE")
			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			'verifica se sono diversi gli sconti quantit&agrave; o la categoria IVA
			ToBeSaved = (iva_id <> cInteger(rs("prz_iva_id"))) OR (cInteger(scontoQ_id) <> cInteger(rs("prz_scontoQ_id")))
			if not ToBeSaved then
				'verifica se sono diversi gli stati di visibilit&agrave;
				if not rsl("listino_offerte") then
					'listino normale
					ToBeSaved = (visibile <> rs("prz_visibile") OR promozione <> rs("prz_promozione"))
				else
					'listino offerte speciali: salva la riga se impostata su "visibile"
					ToBeSaved = visibile
					'promozione non considerata su listino offerte"
				end if
			end if
			rs.close
		end if
	end if
	if ToBeSaved then
		sql = "SELECT * FROM gtb_prezzi WHERE prz_variante_id=" & cIntero(variante_id) & " AND prz_listino_id=" & cIntero(listino)
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		
		if rsl("listino_with_child") then
			if rs.eof then
				child_sql = ""
			else
				child_sql = " UPDATE gtb_prezzi SET " + _
							" prz_prezzo=" & ParseSQL(prezzo_nuovo, adNumeric) & ", " & _
							" prz_var_euro=" & ParseSQL(var_euro, adNumeric) & ", " & _
							" prz_var_sconto=" & ParseSQL(var_sconto, adNumeric) & ", " & _
							" prz_visibile=" & IIF(visibile, 1, 0) & ", " & _
							" prz_promozione=" & IIF(promozione, 1, 0) & ", " & _
							" prz_scontoQ_id=" & IIF(cInteger(scontoQ_id)=0, "NULL", scontoQ_id) & ", " & _
							" prz_iva_id=" & ParseSQL(iva_id, adNumeric) & _
							" WHERE prz_variante_id=" & variante_id & _
							" AND prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE listino_ancestor_id=" & cIntero(listino) & ") " & _
							" AND prz_prezzo=" & ParseSQL(rs("prz_prezzo"), adNumeric) & _
							" AND prz_var_euro=" & ParseSQL(rs("prz_var_euro"), adNumeric) & _
							" AND prz_var_sconto=" & ParseSQL(rs("prz_var_sconto"), adNumeric) & _
							" AND prz_visibile=" & IIF(rs("prz_visibile"), 1, 0) & _
							" AND prz_promozione=" & IIF(rs("prz_promozione"), 1, 0) & _
							" AND ISNULL(prz_scontoQ_id,0)=" & cInteger(rs("prz_scontoQ_id")) & _
							" AND prz_iva_id=" & ParseSQL(rs("prz_iva_id"), adNumeric)
			end if
		end if
		
		
		if rs.eof then
			rs.AddNew
		end if
		rs("prz_listino_id") = listino
		rs("prz_variante_id") = variante_id
		prezzo_attuale = rs("prz_prezzo")
		rs("prz_prezzo") = prezzo_nuovo
		rs("prz_var_euro") = var_euro
		rs("prz_var_sconto") = var_sconto
		rs("prz_visibile") = visibile
		rs("prz_promozione") = ((not rsl("listino_offerte")) AND promozione)
		rs("prz_scontoQ_id") = IIF(cInteger(scontoQ_id)=0, NULL, scontoQ_id)
		rs("prz_iva_id") = iva_id
		if rsl("listino_offerte") then
			if isDate(offerta_dal) then
				rs("prz_offerta_dal") = DateISO(offerta_dal)
			elseif offerta_dal = "azzera" then
				rs("prz_offerta_dal") = NULL
			end if
			if isDate(offerta_al) then
				rs("prz_offerta_al") = DateISO(offerta_al)
			elseif offerta_al = "azzera" then
				rs("prz_offerta_al") = NULL
			end if
		end if
		rs.update
		rs.close
		
		if rsl("listino_with_child") AND child_sql<>"" then
			CALL conn.execute(child_sql, , adCmdText)
		end if
		
		if rsl("listino_base_attuale") then
			if prezzo_attuale <> prezzo_nuovo then
				'se sono in modifica del listino base attuale aggiorno tutti i prezzi dei listini dipendenti dal listino base: listini offerte speciali e listini clienti
				CALL AggiornaPrezzoListiniDaListinoBase(conn, variante_id)
			end if
		end if
	else
		'verifica se la registrazione era gi&agrave; presente
		if personalizzato then
			'cancella la registrazione personalizzata perch&egrave; non pi&ugrave; necessaria
			sql = "DELETE FROM gtb_prezzi WHERE prz_variante_id=" & cIntero(variante_id) & " AND prz_listino_id=" & cIntero(listino)
			CALL conn.execute(sql, , adExecuteNoRecords)
		end if
	end if
end function


sub SearchForm_Init()
	session("B2B_PREZZI_LISTINI_SEARCHED") = true
	if request.form("tutti")<>"" then
		CALL SearchSession_Reset("list_")
		session("B2B_PREZZI_LISTINI_SEARCHED_ALL") = true
	elseif request.form("cerca")<>"" then
		CALL SearchSession_Reset("list_")
		CALL SearchSession_Set("list_")
		session("B2B_PREZZI_LISTINI_SEARCHED_ALL") = false
	end if
end sub

function SearchForm_Parse(categorie)
	dim sql_where
	sql_where = ""
	'ricerca per codici
	sql_where = sql_where + SearchForm_SEARCH_BY_codice(Session("list_codice_int"), Session("list_codice_pro"), cInteger(Session("list_codice_variante"))>0)
	
	'filtra per nome
	sql_where = sql_where + SearchForm_SEARCH_BY_nome(Session("list_nome"))
	
	'filtra per categoria
	sql_where = sql_where + SearchForm_SEARCH_BY_categoria(Session("list_categoria"))
	
	'filtra per marca
	if Session("list_marchio")<>"" then
		sql_where = sql_where & " AND art_marca_id=" & Session("list_marchio")
	end if
	
	'ricerca per visibilita a catalogo
	if Session("list_stato_catalogo")<>"" AND _
	   NOT (instr(1, Session("list_stato_catalogo"), "1", vbTextCompare) AND instr(1, Session("list_stato_catalogo"), "0", vbTextCompare))then
		sql_where = sql_where & " AND NOT (art_disabilitato=" & Session("list_stato_catalogo") & ") AND NOT (rel_disabilitato=" & Session("list_stato_catalogo") & ") "
	end if
	
	
	'ricerca per visibilit&agrave; impostata sul listino
	if Session("list_visibile")<>"" AND _
	   NOT (instr(1, Session("list_visibile"), "1", vbTextCompare) AND instr(1, Session("list_visibile"), "0", vbTextCompare)) then
		sql_where = sql_where & " AND ISNULL(prz_visibile, 0)=" & Session("list_visibile")
	end if
	
	
	'ricerca per variazione di prezzo
	sql_where = sql_where + SearchForm_SEARCH_BY_variazione(rsl, Session("list_Variazioni"), Session("list_variazione_valore"), Session("list_variazione_valore_tipo"))
	

	'ricerca per stato di promozione
	if Session("list_promozione")<>"" AND _
	   NOT (instr(1, Session("list_promozione"), "1", vbTextCompare) AND instr(1, Session("list_promozione"), "0", vbTextCompare)) then
		sql_where = sql_where & " AND ISNULL(prz_promozione, 0)=" & Session("list_promozione")
	end if
	
	'ricerca per stato di offerta speciale
	sql_where = sql_where + SearchForm_SEARCH_BY_Offerta(Session("list_offerte"))
	
	'filtra per aliquota iva
	if Session("list_iva")<>"" then
		sql_where = sql_where & " AND prz_iva_id=" & Session("list_iva")
	end if
	
	'ricerca per classe di sconto per quantit&agrave;
	sql_where = sql_where + SearchForm_SEARCH_BY_scontoQ(Session("list_scontoQ"), Session("list_scontoQ_id"))
	
	if Session("list_personalizzato")<>"" then
		if instr(1, Session("list_personalizzato"), "1", vbTextCompare)<1 OR instr(1, Session("list_personalizzato"), "0", vbTextCompare)<1 then
			if instr(1, Session("list_personalizzato"), "1", vbTextCompare)>0  then
				'ricerca solo prezzi personalizzati
				sql_where = sql_where + SearchForm_SEARCH_BY_personalizzazione_prezzo(rsL, true)
			elseif instr(1, Session("list_personalizzato"), "0", vbTextCompare)>0 then
				'ricerca solo prezzi non personalizzati
				sql_where = sql_where + SearchForm_SEARCH_BY_personalizzazione_prezzo(rsL, false)
			end if
		end if
	end if
	
	if sql_where = "" then
		Session("B2B_PREZZI_LISTINI_SEARCHED_ALL") = true
	end if
	
	SearchForm_Parse = sql_where
end function



function SearchForm_SEARCH_BY_codice(cod_int, cod_pro, includi_varianti)
	if cod_int<>"" OR cod_pro<>"" then
		dim cod_where
		if cod_int<>"" then
            cod_where = SQL_FullTextSearch(cod_int, "rel_cod_int")
		end if
		if cod_int<>"" AND cod_pro<>"" then
			cod_where = cod_where & " AND "
		end if
		if cod_pro<>"" then
			cod_where = cod_where & SQL_FullTextSearch(cod_pro, "rel_cod_pro")
		end if
		if includi_varianti then
			'ricerca anche varianti di questo articolo
			SearchForm_SEARCH_BY_codice = " AND rel_art_id IN (SELECT rel_Art_id FROM grel_art_valori WHERE " + cod_where + ")"
		else
			'ricerca il codice unico
			SearchForm_SEARCH_BY_codice = " AND " & cod_where
		end if
	end if
end function


function SearchForm_SEARCH_BY_nome(nome)
	if nome<>"" then
		SearchForm_SEARCH_BY_nome = " AND " & sql_FullTextSearch(nome, FieldLanguageList("art_nome_"))
	end if
end function


function SearchForm_SEARCH_BY_categoria(categoria)
	if cInteger(categoria)>0 then
		SearchForm_SEARCH_BY_categoria = " AND  art_tipologia_id IN (" & categorie.FoglieID(categoria) & " )"
	end if
end function


function SearchForm_SEARCH_BY_scontoQ(scontoQ_flag, scontoQ_value)
	if scontoQ_flag<>"" then
		SearchForm_SEARCH_BY_scontoQ =  " AND ("
		if instr(1, scontoQ_flag, "0", vbTextCompare) then
			SearchForm_SEARCH_BY_scontoQ = SearchForm_SEARCH_BY_scontoQ + " ISNULL(prz_scontoq_id, 0) = 0 "
			if instr(1, scontoQ_flag, "1", vbTextCompare) then
				SearchForm_SEARCH_BY_scontoQ = SearchForm_SEARCH_BY_scontoQ + " OR "
			end if 
		end if
		if instr(1, scontoQ_flag, "1", vbTextCompare) then
			if cInteger(scontoQ_value) > 0 then
				SearchForm_SEARCH_BY_scontoQ = SearchForm_SEARCH_BY_scontoQ + " ISNULL(prz_scontoq_id, 0) = " & scontoQ_value
			else
				SearchForm_SEARCH_BY_scontoQ = SearchForm_SEARCH_BY_scontoQ + " ISNULL(prz_scontoq_id, 0) <> 0 "
			end if
		end if 
		SearchForm_SEARCH_BY_scontoQ = SearchForm_SEARCH_BY_scontoQ + " )"
	end if
end function


function SearchForm_SEARCH_BY_personalizzazione_prezzo(rsL, solo_personalizzati)
	SearchForm_SEARCH_BY_personalizzazione_prezzo = SearchForm_SEARCH_BY_personalizzazione_prezzoEX(rsL, solo_personalizzati, 0, 0)
end function

function SearchForm_SEARCH_BY_personalizzazione_prezzoEX(rsL, solo_personalizzati, rel_id, art_id)
	'filtra per i prezzi personalizzati del listino (qualsiasi personalizzazione)
	rel_id = cInteger(rel_id)
	art_id = cInteger(art_id)
	if solo_personalizzati  then
		'ricerca solo prezzi personalizzati
		if cInteger(rsl("listino_ancestor_id"))>0 then
			'listino con prezzi derivati: deve escludere anche quelli
			SearchForm_SEARCH_BY_personalizzazione_prezzoEX = " AND " + _
				" (ISNULL(prz_id, 0) <> 0 AND " + _
				" ISNULL(prz_listino_id,0)=" & rsl("listino_id") & " AND " + _
				IIF(rel_id>0, "prz_variante_id = " & rel_id & " AND ", "") + _
				IIF(art_id>0 AND rel_id=0, "prz_variante_id IN( SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(art_id) & ") AND ", "") + _
				" prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_prezzi L_ancestor " + _
				" WHERE L_ancestor.prz_listino_id=" & rsl("listino_ancestor_id") & " AND " + _
				"	ISNULL(L_ancestor.prz_visibile,0) = ISNULL(gv_listini.prz_visibile,0) AND " + _
				"	ISNULL(L_ancestor.prz_promozione,0) = ISNULL(gv_listini.prz_promozione,0) AND " + _
				"	ISNULL(L_ancestor.prz_scontoQ_id,0) = ISNULL(gv_listini.prz_scontoQ_id,0) AND " + _
				"	ISNULL(L_ancestor.prz_iva_id,0) = ISNULL(gv_listini.prz_iva_id,0) AND " + _
				"	ISNULL(L_ancestor.prz_prezzo,0) = ISNULL(gv_listini.prz_prezzo,0) AND " + _
				"	ISNULL(L_ancestor.prz_var_euro,0) = ISNULL(gv_listini.prz_var_euro,0) AND " + _
				"	ISNULL(L_ancestor.prz_var_sconto,0) = ISNULL(gv_listini.prz_var_sconto,0) AND " + _
				"	L_ancestor.prz_variante_id = gv_listini.prz_variante_id  " + _
				IIF(rel_id>0, " AND L_ancestor.prz_variante_id = " & rel_id, "") + _
				IIF(art_id>0 AND rel_id=0, "AND L_ancestor.prz_variante_id IN( SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(art_id) & ")", "") + _
				" ) )"
		else
			if rsL("listino_base") then
				SearchForm_SEARCH_BY_personalizzazione_prezzoEX = " AND ( " + _
					" (prz_visibile = rel_disabilitato) OR " + _
					" (ISNULL(prz_promozione,0)=1) OR " + _
					" (ISNULL(prz_scontoQ_id,0) <> ISNULL(rel_scontoQ_id,0)) OR " + _
					" (ISNULL(prz_var_euro,0) <> 0) OR " + _
					" (ISNULL(prz_var_sconto,0) <> 0) )" + _
					IIF(rel_id>0, " AND prz_variante_id = " & rel_id, "") + _
					IIF(art_id>0 AND rel_id=0, "AND prz_variante_id IN( SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(art_id) & ")", "")
			else
				SearchForm_SEARCH_BY_personalizzazione_prezzoEX = " AND " + _
					" (ISNULL(prz_id, 0) <> 0 AND ISNULL(prz_listino_id,0)=" & rsl("listino_id") & ")" + _
					IIF(rel_id>0, " AND prz_variante_id = " & rel_id, "") + _
					IIF(art_id>0 AND rel_id=0, "AND prz_variante_id IN( SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(art_id) & ")", "")
			end if
		end if
	elseif not solo_personalizzati then
		'ricerca solo prezzi non personalizzati
		if cInteger(rsl("listino_ancestor_id"))>0 then
			'listino con prezzi derivati: deve includere anche quelli derivati dal listino "ancestor"
			SearchForm_SEARCH_BY_personalizzazione_prezzoEX = " AND " + _
				" (ISNULL(prz_id, 0) = 0 OR " + _
				" ISNULL(prz_listino_id,0)<>" & rsl("listino_id") & " OR " + _
				" prz_variante_id IN (SELECT prz_variante_id FROM gtb_prezzi L_ancestor " + _
				" WHERE L_ancestor.prz_listino_id=" & rsl("listino_ancestor_id") & " AND " + _
				"	ISNULL(L_ancestor.prz_visibile,0) = ISNULL(gv_listini.prz_visibile,0) AND " + _
				"	ISNULL(L_ancestor.prz_promozione,0) = ISNULL(gv_listini.prz_promozione,0) AND " + _
				"	ISNULL(L_ancestor.prz_scontoQ_id,0) = ISNULL(gv_listini.prz_scontoQ_id,0) AND " + _
				"	ISNULL(L_ancestor.prz_iva_id,0) = ISNULL(gv_listini.prz_iva_id,0) AND " + _
				"	ISNULL(L_ancestor.prz_prezzo,0) = ISNULL(gv_listini.prz_prezzo,0) AND " + _
				"	ISNULL(L_ancestor.prz_var_euro,0) = ISNULL(gv_listini.prz_var_euro,0) AND " + _
				"	ISNULL(L_ancestor.prz_var_sconto,0) = ISNULL(gv_listini.prz_var_sconto,0) AND " + _
				"	L_ancestor.prz_variante_id = gv_listini.prz_variante_id  " + _
				IIF(rel_id>0, " AND L_ancestor.prz_variante_id = " & rel_id, "") + _
				IIF(art_id>0 AND rel_id=0, "AND L_ancestor.prz_variante_id IN( SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & art_id & ")", "") + _
				" ) )" + _
				IIF(rel_id>0, "prz_variante_id = " & rel_id & " AND ", "") + _
				IIF(art_id>0 AND rel_id=0, "prz_variante_id IN( SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & art_id & ") AND ", "")
		else
			if rsL("listino_base") then
				SearchForm_SEARCH_BY_personalizzazione_prezzoEX = " AND ( " + _
					" (prz_visibile <> rel_disabilitato) AND " + _
					" (ISNULL(prz_promozione,0)=0) AND " + _
					" (ISNULL(prz_scontoQ_id,0) = ISNULL(rel_scontoQ_id,0)) AND " + _
					" (ISNULL(prz_var_euro,0) = 0) AND " + _
					" (ISNULL(prz_var_sconto,0) = 0) )" + _
					IIF(rel_id>0, " AND prz_variante_id = " & rel_id, "") + _
					IIF(art_id>0 AND rel_id=0, "AND prz_variante_id IN( SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & art_id & ")", "")
			else
				SearchForm_SEARCH_BY_personalizzazione_prezzoEX = " AND " + _
					" (ISNULL(prz_id, 0) = 0 OR ISNULL(prz_listino_id,0)<>" & rsl("listino_id") & ")" + _
					IIF(rel_id>0, " AND prz_variante_id = " & rel_id, "") + _
					IIF(art_id>0 AND rel_id=0, "AND prz_variante_id IN( SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & art_id & ")", "")
			end if
		end if
	end if
end function

'ricerca per variazione di prezzo
function SearchForm_SEARCH_BY_variazione(rsl, variazioni_flag, variazioni_valore, variazioni_tipo)
	
	if variazioni_flag<>"" then
		SearchForm_SEARCH_BY_variazione = " AND ("
		if instr(1, variazioni_flag, "0", vbTextCompare)>0 then
			'senza variazioni
			if rsl("listino_base") then
				SearchForm_SEARCH_BY_variazione = SearchForm_SEARCH_BY_variazione & " (ISNULL(prz_var_euro, 0)=0 AND ISNULL(prz_var_sconto, 0)=0) "
			else
				SearchForm_SEARCH_BY_variazione = SearchForm_SEARCH_BY_variazione & " (prz_listino_id<>" & rsL("Listino_id") & " OR (ISNULL(prz_var_euro, 0)=0 AND ISNULL(prz_var_sconto, 0)=0))"
			end if
		end if
		if instr(1, variazioni_flag, "0", vbTextCompare)>0 AND instr(1, variazioni_flag, "1", vbTextCompare)>0 then
			SearchForm_SEARCH_BY_variazione = SearchForm_SEARCH_BY_variazione & ") OR ("
		end if
		if instr(1, variazioni_flag, "1", vbTextCompare)>0 then
			'listino con variazioni
			if rsl("listino_base") then
				SearchForm_SEARCH_BY_variazione = SearchForm_SEARCH_BY_variazione & " (ISNULL(prz_var_euro, 0)<>0 OR ISNULL(prz_var_sconto, 0)<>0) "
			else
				SearchForm_SEARCH_BY_variazione = SearchForm_SEARCH_BY_variazione & " (prz_listino_id=" & rsL("Listino_id") & " AND (ISNULL(prz_var_euro, 0)<>0 OR ISNULL(prz_var_sconto, 0)<>0))"
			end if
			if variazioni_valore<>"" then
				if cReal(variazioni_valore)<>0 then
					SearchForm_SEARCH_BY_variazione = SearchForm_SEARCH_BY_variazione & " AND  ( prz_var_" & variazioni_tipo & "=" & ParseSQL(variazioni_valore, adNumeric) & " ) "
				end if
			end if 
		end if
		SearchForm_SEARCH_BY_variazione = SearchForm_SEARCH_BY_variazione & " )"
	end if
end function


function SearchForm_SEARCH_BY_Offerta(offerte_flag)
	if offerte_flag<>"" AND _
	   NOT (instr(1, offerte_flag, "1", vbTextCompare) AND instr(1, offerte_flag, "0", vbTextCompare)) then
	   	if instr(1, offerte_flag, "1", vbTextCompare) then	
			SearchForm_SEARCH_BY_Offerta = " AND prz_variante_id IN "
		else
			SearchForm_SEARCH_BY_Offerta = " AND prz_variante_id NOT IN "
		end if
		'SearchForm_SEARCH_BY_Offerta = SearchForm_SEARCH_BY_Offerta + " (SELECT prz_variante_ID FROM gv_listino_offerte) " --- Giacomo 2012/03/21 COMMENTATO perchè con questa condizione la query finale si impiantava
		SearchForm_SEARCH_BY_Offerta = SearchForm_SEARCH_BY_Offerta + "("&GetValueList(NULL, NULL, "SELECT prz_variante_ID FROM gv_listino_offerte")&")"
	end if
end function


sub SearchForm_Write(categorie, conn)%>
	<script language="JavaScript" type="text/javascript">
		//variabile utilizzata per il controllo del submit nel form di ricerca
		var ClickVediTutti = true;
		
		function verifica_intenzioni(){
			if (ClickVediTutti){
				return window.confirm('ATTENZIONE: se il numero di articoli e\' elevato la pagina potrebbe impiegare alcuni minuti per essere visualizzata. \n' + 
									  'Visualizzare comunque TUTTI gli articoli?')
			}
			else
				return true;
		}
		
		function RicercaCodice(){
			var codici_ricercati = "";
			codici_ricercati += ricerca.search_codice_int.value;
			codici_ricercati += ricerca.search_codice_pro.value;
			if(codici_ricercati!=""){
				DisableControl(document.all.search_codice_variante_0, false);
				DisableControl(document.all.search_codice_variante_1, false);
			}
			else{
				DisableControl(document.all.search_codice_variante_0, true);
				DisableControl(document.all.search_codice_variante_1, true);
			}
		}
	</script>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:7px; width:735px;">
		<form action="" method="post" id="ricerca" name="ricerca" onsubmit="return verifica_intenzioni();">
		<caption>
			<table cellspacing="0" cellpadding="0" align="right">
				<tr>
					<td align="right">
						<input type="submit" name="cerca" value="CERCA" class="button" onclick="ClickVediTutti=false;">
						<input type="submit" class="button" name="tutti" value="VEDI TUTTI" onclick="ClickVediTutti=true;">
					</td>
				</tr>
			</table>
			Opzioni di ricerca
		</caption>
		<tr>
			<th style="width:38%;" colspan="4" <%= Search_Bg("list_codice_int;list_codice_pro") %>>CODICE ARTICOLO</th>
			<th colspan="2" <%= Search_Bg("list_codice_variante") %>>INCLUSIONE VARIANTI</th>
			<th colspan="2" style="width:30%;" <%= Search_Bg("list_nome") %>>NOME ARTICOLO</th>
			<th class="vertical" rowspan="4" width="1.5%"> filtri sul catalogo</th>
		</tr>
		<tr>
			<td class="label" style="width:8%;">interno:</td>
			<td class="content" style="width:10%;">
				<input type="text" name="search_codice_int" value="<%= TextEncode(session("list_codice_int")) %>" style="width:100%;" onchange="RicercaCodice();" onblur="RicercaCodice();">
			</td>
			<td class="label" style="width:10%;">produttore:</td>
			<td class="content" style="width:10%;">
				<input type="text" name="search_codice_pro" value="<%= TextEncode(session("list_codice_pro")) %>" style="width:100%;" onchange="RicercaCodice();" onblur="RicercaCodice();">
			</td>
			<td class="content">
				<input type="radio" class="checkbox" name="search_codice_variante" id="search_codice_variante_0" value="0" <%= chk(cInteger(Session("list_codice_variante"))=0) %> <%= disable((session("list_codice_int") & session("list_codice_pro"))="") %>>
				codice indicato
			</td>
			<td class="content">
				<input type="radio" class="checkbox" name="search_codice_variante" id="search_codice_variante_1" value="1" <%= chk(cInteger(Session("list_codice_variante"))=1) %> <%= disable((session("list_codice_int") & session("list_codice_pro"))="") %>>
				includi varianti
			</td>
			<td class="content" colspan="2">
				<input type="text" name="search_nome" value="<%= TextEncode(session("list_nome")) %>" style="width:100%;">
			</td>
		</tr>
		<tr>
			<th colspan="4" <%= Search_Bg("list_categoria") %>>CATEGORIA</th>
			<th colspan="2" <%= Search_Bg("list_marchio") %>>MARCHIO</th>
			<th colspan="2" <%= Search_Bg("list_stato_catalogo") %>>STATO ARTICOLO A CATALOGO</th>
		</tr>
		<tr>
			<td class="content" colspan="4">
				<% CALL categorie.WritePicker("ricerca", "search_categoria", session("list_categoria"), false, false, 29) %>
			</td>
			<td class="content" colspan="2">
				<%	sql = "SELECT * FROM gtb_marche ORDER BY mar_nome_it"
				CALL dropDown(conn, sql, "mar_id", "mar_nome_it", "search_marchio", session("list_marchio"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="1" <%= chk(instr(1, session("list_stato_catalogo"), "1", vbTextCompare)>0) %>>
				visibile
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="0" <%= chk(instr(1, Session("list_stato_catalogo"), "0", vbTextCompare)>0) %>>
				non visibile
			</td>
		</tr>
		<tr>
			<th colspan="4" <%= Search_Bg("list_promozione") %>>PROMOZIONE</th>
			<th colspan="2" <%= Search_Bg("list_offerte") %>>OFFERTE SPECIALI</th>
			<th colspan="2" <%= Search_Bg("list_visibile") %>>VISIBILIT&Agrave; IMPOSTATA SUL LISTINO</th>
			<th class="vertical" rowspan="<%= IIF(rsl("listino_base"), "4", "6") %>" width="1.5%"> filtri sul listino</th>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="checkbox" <%= disable(rsl("listino_offerte")) %> name="search_promozione" value="1" <%= chk(instr(1, Session("list_promozione"), "1", vbTextCompare)>0) %>>
				in promozione
			</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="checkbox" <%= disable(rsl("listino_offerte")) %> name="search_promozione" value="0" <%= chk(instr(1, session("list_promozione"), "0", vbTextCompare)>0) %>>
				non in promozione
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_offerte" value="1" <%= chk(instr(1, Session("list_offerte"), "1", vbTextCompare)>0) %>>
				in offerta
				<span class="Icona Offerte" title="articolo in offerta">&nbsp;</span>
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_offerte" value="0" <%= chk(instr(1, session("list_offerte"), "0", vbTextCompare)>0) %>>
				non in offerta
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_visibile" value="1" <%= chk(instr(1, Session("list_visibile"), "1", vbTextCompare)>0) %>>
				visibile
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_visibile" value="0" <%= chk(instr(1, session("list_visibile"), "0", vbTextCompare)>0) %>>
				non visibile
			</td>
		</tr>
		<tr>
			<th colspan="5" <%= Search_Bg("list_variazioni;list_variazione_valore;list_variazione_valore_tipo") %>>VARIAZIONI PREZZO</th>
			<th colspan="3" <%= Search_Bg("list_scontoQ;list_scontoQ_id") %>>SCONTO PER QUANTITA'</th>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="checkbox" name="search_variazioni" value="0" <%= chk(instr(1, session("list_variazioni"), "0", vbTextCompare)>0) %> onclick="EnableIfChecked(document.all.search_variazioni_1, ricerca.search_variazione_valore);EnableIfChecked(document.all.search_variazioni_1, ricerca.search_variazione_valore_tipo);">
				senza variazione
			</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="checkbox" name="search_variazioni" id="search_variazioni_1" value="1" <%= chk(instr(1, Session("list_variazioni"), "1", vbTextCompare)>0) %> onclick="EnableIfChecked(this, ricerca.search_variazione_valore);EnableIfChecked(this, ricerca.search_variazione_valore_tipo);">
				con variazione
			</td>
			<td class="content">
				<input type="text" name="search_variazione_valore" value="<%= Session("list_variazione_valore") %>" size="2" <% if not(instr(1, session("list_variazioni"), "1", vbTextCompare)>0) then %> disabled class="disabled"<% end if %>>
				<select name="search_variazione_valore_tipo" <%= disable(not(instr(1, session("list_variazioni"), "1", vbTextCompare)>0)) %>>
					<option value="euro" <%= IIF(Session("list_variazione_valore_tipo")="euro", "selected", "") %>>&euro;</option>
					<option value="sconto" <%= IIF(Session("list_variazione_valore_tipo")="euro", "", "selected") %>>%</option>
				</select>
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_scontoQ" value="0" <%= chk(instr(1, Session("list_scontoQ"), "0", vbTextCompare)>0) %> onclick="EnableIfChecked(document.all.search_scontoQ_1, ricerca.search_scontoQ_id)">
				senza sconto
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_scontoQ" id="search_scontoQ_1" value="1" <%= chk(instr(1, session("list_scontoQ"), "1", vbTextCompare)>0) %> onclick="EnableIfChecked(this, ricerca.search_scontoQ_id)">
				con sconto
			</td>
			<td class="content">
				<%	sql = "SELECT * FROM gtb_scontiQ_classi ORDER BY scc_nome"
				CALL dropDown(conn, sql, "scc_id", "scc_nome", "search_scontoQ_id", session("list_scontoQ_id"), false, " style=""width:100%;"" " & disable(not(instr(1, session("list_scontoQ"), "1", vbTextCompare)>0)), LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<th colspan="9" <%= Search_Bg("list_personalizzato") %>>PERSONALIZZAZIONE LISTINO</th>
		</tr>
		<tr>
			<td class="content" colspan="3">
				<input type="checkbox" class="checkbox" name="search_personalizzato" value="0" <%= chk(instr(1, Session("list_personalizzato"), "0", vbTextCompare)>0) %>>
				senza alcuna personalizzazione
			</td>
			<td class="content" colspan="6">
				<input type="checkbox" class="checkbox" name="search_personalizzato" value="1" <%= chk(instr(1, Session("list_personalizzato"), "1", vbTextCompare)>0) %>>
				personalizzati
				<% if cInteger(rsl("listino_ancestor_id"))>0 then %>
					<span class="note">&nbsp;(esclusi anche i prezzi non variati dal listino principale)</span>
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="footer" style="padding-right:0px;" colspan="9">
				<input type="submit" name="cerca" value="CERCA" class="button" onclick="ClickVediTutti=false;">
				<input type="submit" class="button" name="tutti" value="VEDI TUTTI" onclick="ClickVediTutti=true;">
			</td>
		</tr>
		</form>
	</table>
<% end sub 



'funzione per il cambio del listino principale da cui deriva il listino corrente
sub Listino_ChangeAncestor(conn, ListinoId, AncestorNew, AncestorOld)
	dim sql
	'non tocca i dati del listino: cambia solo le righe

	'cancella le righe del listino corrente che derivano dal listino ancestor
	sql = "DELETE FROM gtb_prezzi WHERE " + _
		  " gtb_prezzi.prz_listino_id=" & ListinoId & " AND " + _
		  " gtb_prezzi.prz_variante_id IN (SELECT prz_variante_id FROM gtb_prezzi L_ancestor WHERE " + _
		  							   	  "ISNULL(L_ancestor.prz_visibile,0) = ISNULL(gtb_prezzi.prz_visibile,0) AND " + _
										  "ISNULL(L_ancestor.prz_promozione,0) = ISNULL(gtb_prezzi.prz_promozione,0) AND " + _
										  "ISNULL(L_ancestor.prz_scontoQ_id,0) = ISNULL(gtb_prezzi.prz_scontoQ_id,0) AND " + _
										  "ISNULL(L_ancestor.prz_iva_id,0) = ISNULL(gtb_prezzi.prz_iva_id,0) AND " + _
										  "ISNULL(L_ancestor.prz_prezzo,0) = ISNULL(gtb_prezzi.prz_prezzo,0) AND " + _
										  "ISNULL(L_ancestor.prz_var_euro,0) = ISNULL(gtb_prezzi.prz_var_euro,0) AND " + _
										  "ISNULL(L_ancestor.prz_var_sconto,0) = ISNULL(gtb_prezzi.prz_var_sconto,0) AND " + _
										  "L_ancestor.prz_variante_id = gtb_prezzi.prz_variante_id AND " + _
										  "L_ancestor.prz_listino_id=" & AncestorOld & ")"
	CALL conn.execute(sql, , adCmdText)

	'inserisce le righe del nuovo listino ancestor
	sql = " INSERT INTO gtb_prezzi (prz_listino_id, prz_prezzo, prz_visibile, prz_promozione, " + _
		  " prz_variante_id, prz_scontoQ_id, prz_iva_id, prz_var_euro, prz_var_sconto) " + _
		  " SELECT " & ListinoId & ", L_ancestor.prz_prezzo, L_ancestor.prz_visibile, L_ancestor.prz_promozione, " + _
		  " L_ancestor.prz_variante_id, L_ancestor.prz_scontoQ_id, L_ancestor.prz_iva_id, L_ancestor.prz_var_euro, L_ancestor.prz_var_Sconto " + _
		  " FROM gtb_prezzi L_ancestor WHERE prz_listino_id=" & AncestorNew & _
		  " AND L_ancestor.prz_variante_id NOT IN ( SELECT L_child.prz_variante_id FROM " + _
		  " gtb_prezzi L_child WHERE prz_listino_id=" & ListinoID & ")"
	CALL conn.execute(sql, , adCmdText)

end sub



function GetPrezzoMinimoVenditaListiniAncestor(conn, rsPrezzi, rel_id, prezzo_acquisto, intervallo_sconto_massimo)
	dim PrezzoMinimoAncestor, IntervalloEuro, IntervalloNonScontabile, IntervalloScontabile, sql, value, MinValue
	sql = " SELECT MIN(prz_prezzo) AS PREZZI_MINIMO, COUNT(*) AS PREZZI_NUMERO, " + _
		  " (SELECT COUNT(*) FROM gtb_listini WHERE listino_with_child=1) AS LISTINI_NUMERO " + _
		  " FROM gtb_prezzi INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + _
		  " WHERE gtb_prezzi.prz_variante_id =" & rel_id & " AND gtb_listini.listino_with_child=1 "
	rsPrezzi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if cInteger(rsPrezzi("PREZZI_NUMERO")) = cInteger(rsPrezzi("LISTINI_NUMERO")) then
		PrezzoMinimoAncestor = cReal(rsPrezzi("PREZZI_MINIMO"))
	else
		rsPrezzi.close
		sql = " SELECT MIN(prz_prezzo) AS PREZZI_MINIMO " + _
		  	  " FROM gtb_prezzi INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + _
		  	  " WHERE gtb_prezzi.prz_variante_id =" & rel_id & _
			  " AND (IsNull(gtb_listini.listino_with_child, 0)=1 OR IsNull(gtb_listini.listino_base_attuale, 0)=1)"
		rsPrezzi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		PrezzoMinimoAncestor = cReal(rsPrezzi("PREZZI_MINIMO"))
	end if
	rsPrezzi.close
	
	IntervalloEuro = PrezzoMinimoAncestor - prezzo_acquisto
	IntervalloScontabile = (IntervalloEuro/100) * intervallo_sconto_massimo
	IntervalloNonScontabile = IntervalloEuro - IntervalloScontabile
	GetPrezzoMinimoVenditaListiniAncestor = prezzo_acquisto + IntervalloNonScontabile
end function



function GetPrezzoOfferta(conn, rsPrezzo, rel_id)
	sql = "SELECT prz_prezzo FROM gv_listino_offerte WHERE prz_variante_id=" & rel_id
	rsPrezzo.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if rsPrezzo.eof then
		'articolo non in offerta
		GetPrezzoOfferta = 0
	else
		GetPrezzoOfferta = rs("prz_prezzo")
	end if
	rsPrezzo.close
end function


function IsRowAncestorDerived(conn, rsp, listino_id, listino_ancestor_id, prz_variante_id)
	dim sql, value
	sql = "SELECT COUNT(*) FROM gtb_prezzi WHERE " + _
		  " gtb_prezzi.prz_variante_id=" & prz_variante_id & " AND " + _
		  " gtb_prezzi.prz_listino_id=" & listino_id & " AND " + _
		  " gtb_prezzi.prz_variante_id IN (SELECT prz_variante_id FROM gtb_prezzi L_ancestor WHERE " + _
		  							   	  "ISNULL(L_ancestor.prz_visibile,0) = ISNULL(gtb_prezzi.prz_visibile,0) AND " + _
										  "ISNULL(L_ancestor.prz_promozione,0) = ISNULL(gtb_prezzi.prz_promozione,0) AND " + _
										  "ISNULL(L_ancestor.prz_scontoQ_id,0) = ISNULL(gtb_prezzi.prz_scontoQ_id,0) AND " + _
										  "ISNULL(L_ancestor.prz_iva_id,0) = ISNULL(gtb_prezzi.prz_iva_id,0) AND " + _
										  "ISNULL(L_ancestor.prz_prezzo,0) = ISNULL(gtb_prezzi.prz_prezzo,0) AND " + _
										  "ISNULL(L_ancestor.prz_var_euro,0) = ISNULL(gtb_prezzi.prz_var_euro,0) AND " + _
										  "ISNULL(L_ancestor.prz_var_sconto,0) = ISNULL(gtb_prezzi.prz_var_sconto,0) AND " + _
										  "L_ancestor.prz_variante_id = gtb_prezzi.prz_variante_id AND " + _
										  "L_ancestor.prz_listino_id=" & listino_ancestor_id & ")"
	if cInteger(GetValueList(Conn, rsp, sql))>0 then
		'riga uguale al padre: derivata
		IsRowAncestorDerived = false
	else
		'riga non derivata
		IsRowAncestorDerived = true
	end if
end function

%> 