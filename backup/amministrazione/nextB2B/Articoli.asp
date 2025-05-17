<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->

<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione articoli - elenco"
dicitura.puls_new = "nuovo:;ARTICOLO SINGOLO;ARTICOLO CON VARIANTI;BUNDLE;CONFEZIONE"
dicitura.link_new = ";ArticoliNew.asp?TYPE=AS;ArticoliNew.asp?TYPE=AV;ArticoliNew.asp?TYPE=B;ArticoliNew.asp?TYPE=C"
dicitura.scrivi_con_sottosez() 

dim conn, rs, rsc, rsv, sql, Pager, colore, txt, title, rs_spe
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("art_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("art_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
set rs_spe = Server.CreateObject("ADODB.RecordSet")

'filtra per codice interno
if Session("art_codice_int")<>"" then
	if sql <>"" then sql = sql & " AND "
    sql = sql & "( " & SQL_FullTextSearch(Session("art_codice_int"), "art_cod_int") & " OR " & _
                " art_id IN (SELECT rel_art_id FROM grel_art_valori WHERE " & SQL_FullTextSearch(Session("art_codice_int"), "rel_cod_int") & " ) )"
end if

'filtra per codice produttore
if Session("art_codice_pro")<>"" then
	if sql <>"" then sql = sql & " AND "
    sql = sql & "( " & SQL_FullTextSearch(Session("art_codice_pro"), "art_cod_pro") & " OR " & _
                " art_id IN (SELECT rel_art_id FROM grel_art_valori WHERE " & SQL_FullTextSearch(Session("art_codice_pro"), "rel_cod_pro") & " ) )"
end if

'filtra per nome
if Session("art_nome")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & SQL_FullTextSearch(Session("art_nome"), FieldLanguageList("art_nome_"))
end if

'filtra per categoria
if Session("art_categoria")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " art_tipologia_id IN ("&Session("art_categoria") & _
				IIF(categorie.DiscendentiID(Session("art_categoria"))<>"",","&categorie.DiscendentiID(Session("art_categoria")),"") & " )"
end if

'filtra per marca
if Session("art_marchio")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " art_marca_id=" & Session("art_marchio")
end if

'ricerca per stato a catalogo
if Session("art_stato_catalogo")<>"" then
	if not (instr(1, Session("art_stato_catalogo"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("art_stato_catalogo"), "0", vbTextCompare)>0 ) then
		if sql <>"" then sql = sql & " AND "
		if instr(1, Session("art_stato_catalogo"), "1", vbTextCompare)>0 then
			'articolo a catalogo
			sql = sql & " ISNULL(art_disabilitato, 0)=0 "
		elseif instr(1, Session("art_stato_catalogo"), "0", vbTextCompare)>0 then
			'articolo fuori catalogo
			sql = sql & " ISNULL(art_disabilitato, 0)=1 "
		end if
	end if
end if

'ricerca per varianti
if Session("art_varianti")<>"" then
	if not (instr(1, Session("art_varianti"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("art_varianti"), "0", vbTextCompare)>0 ) then
		if sql <>"" then sql = sql & " AND "
		if instr(1, Session("art_varianti"), "1", vbTextCompare)>0 then
			'articolo con varianti
			sql = sql & " ISNULL(art_varianti, 0)=1 "
		elseif instr(1, Session("art_varianti"), "0", vbTextCompare)>0 then
			'articolo senza varianti
			sql = sql & " ISNULL(art_varianti, 0)=0 "
		end if
	end if
end if

'ricerca per tipo di articolo
if Session("art_tipo")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " ( "
	'articolo singolo
	if instr(1, Session("art_tipo"), "AS", vbTextCompare)>0  then
		sql = sql & "(ISNULL(art_varianti, 0)=0 AND ISNULL(art_se_bundle, 0)=0 AND ISNULL(art_se_confezione,0)=0) OR "
	end if
	'articolo con varianti
	if instr(1, Session("art_tipo"), "AV", vbTextCompare)>0  then
		sql = sql & "((" & SQL_isTrue(conn, "art_varianti") & ") AND ISNULL(art_se_bundle, 0)=0 AND ISNULL(art_se_confezione,0)=0) OR "
	end if
	'bundle
	if instr(1, Session("art_tipo"), "B", vbTextCompare)>0  then
		sql = sql & "(" & SQL_isTrue(conn, "art_se_bundle") & ") OR "
	end if
	'confezione
	if instr(1, Session("art_tipo"), "C", vbTextCompare)>0  then
		sql = sql & "(" & SQL_isTrue(conn, "art_se_confezione") & ") OR "
	end if
	sql = left(sql, len(sql)-3) & " )"
end if

'ricerca per tipo di aggregazione
if Session("art_aggregazione")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " ( "
	'articolo singolo
	if instr(1, Session("art_aggregazione"), "A", vbTextCompare)>0  then
		sql = sql & "(ISNULL(art_in_bundle, 0)=0 AND ISNULL(art_in_confezione, 0)=0) OR "
	end if
	'bundle
	if instr(1, Session("art_aggregazione"), "B", vbTextCompare)>0  then
		sql = sql & "(" & SQL_isTrue(conn, "art_in_bundle") & ") OR "
	end if
	'confezione
	if instr(1, Session("art_aggregazione"), "C", vbTextCompare)>0  then
		sql = sql & "(" & SQL_isTrue(conn, "art_in_confezione") & ") OR "
	end if
	sql = left(sql, len(sql)-3) & " )"
end if

'ricerca per accessori
if Session("art_accessori")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " ( "
	'articolo senza accessori
	if instr(1, Session("art_accessori"), "S", vbTextCompare)>0  then
		sql = sql & " (ISNULL(art_ha_accessori, 0)=0) OR "
	end if
	'articolo con accessori
	if instr(1, Session("art_accessori"), "C", vbTextCompare)>0  then
		sql = sql & " (" & SQL_isTrue(conn, "art_ha_accessori") & ") OR "
	end if
	'accessorio di un articolo
	if instr(1, Session("art_accessori"), "A", vbTextCompare)>0  then
		sql = sql & " (" & SQL_isTrue(conn, "art_se_accessorio") & ") OR "
	end if
	sql = left(sql, len(sql)-3) & " )"
end if

'ricerca per tipo spedizione	  
if Session("art_spedizione")<>"" then
	if sql <>"" then sql = sql & " AND "
		sql = sql & " art_spedizione_id=" & Session("art_spedizione")
end if

'filtra per variante
if Session("art_variante")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " art_id IN (SELECT rel_art_id FROM grel_art_valori WHERE rel_id IN " + _
				" (SELECT rvv_art_var_id FROM grel_art_vv WHERE rvv_val_id IN (SELECT val_id FROM gtb_valori WHERE val_var_id = " & Session("art_variante") & ")))" 
end if

'filtra per valore variante
if Session("art_valore_variante")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " art_id IN (SELECT rel_art_id FROM grel_art_valori WHERE rel_id IN " + _
				" (SELECT rvv_art_var_id FROM grel_art_vv WHERE rvv_val_id = " & Session("art_valore_variante") & "))" 
end if

'ricerca full-text
if Session("art_full_text")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & SQL_FullTextSearch(Session("art_full_text"), FieldLanguageList("art_nome_;art_descr_"))
end if


if cIntero(Session("ID_CATEGORIA_BASE_CATALOGO")) > 0 then
	if Session("art_catalogo_vendita")<>"" then
		if not (instr(1, Session("art_catalogo_vendita"), "1", vbTextCompare)>0 AND _
				instr(1, Session("art_catalogo_vendita"), "0", vbTextCompare)>0 ) then
			if sql <>"" then sql = sql & " AND "
			if instr(1, Session("art_catalogo_vendita"), "1", vbTextCompare)>0 then
				'catalogo articoli in vendita
				sql = sql & " (art_tipologia_id IN " & _
							" (SELECT tip_id FROM gtb_tipologie where ',' + tip_tipologie_padre_lista + ',' LIKE '%,"&cIntero(Session("ID_CATEGORIA_BASE_CATALOGO"))&",%')) "
			elseif instr(1, Session("art_catalogo_vendita"), "0", vbTextCompare)>0 then
				'catalogo articoli non in vendita
				sql = sql & " (art_tipologia_id NOT IN " & _
							" (SELECT tip_id FROM gtb_tipologie where ',' + tip_tipologie_padre_lista + ',' LIKE '%,"&cIntero(Session("ID_CATEGORIA_BASE_CATALOGO"))&",%')) "
			end if
		end if
	end if
end if

if sql <> "" then sql = " WHERE " & sql
sql = " SELECT [art_id],[art_nome_it],[art_nome_en],[art_nome_fr],[art_nome_es],[art_nome_de],[art_cod_int],[art_cod_pro],[art_cod_alt],[art_prezzo_base]," & _
	  " [art_scontoQ_id],[art_giacenza_min],[art_lotto_riordino],[art_qta_min_ord],[art_NovenSingola],[art_se_accessorio],[art_ha_accessori],[art_se_bundle]," & _
	  " [art_se_confezione],[art_in_bundle],[art_in_confezione],[art_varianti],[art_disabilitato],[art_tipologia_id],[art_marca_id],[art_iva_id]," & _
	  " [art_external_id],[art_raggruppamento_id],[art_insData],[art_insAdmin_id],[art_modData],[art_modAdmin_id]," & _
	  " [art_non_vendibile],[art_applicativo_id],[art_unico],[art_spedizione_id], [art_ordine]," & _
	  " [art_dettagli_ord_tipo_id], gtb_marche.mar_nome_it " + _
	  "	FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id=gtb_marche.mar_id " + _
	  sql & " ORDER BY art_nome_it"
	  
Session("B2B_ARTICOLI_SQL") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0">
		<tr>
	  		<td style="width:27%;" valign="top">
<!-- BLOCCO DI RICERCA -->
				<form action="" method="post" id="ricerca" name="ricerca">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer" colspan="2">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<% if cIntero(Session("ID_CATEGORIA_BASE_CATALOGO")) > 0 then %>
									<tr><th colspan="2" <%= Search_Bg("art_catalogo_vendita") %>>SCEGLI IL CATALOGO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="checkbox" class="checkbox" name="search_catalogo_vendita" value="1" <%= chk(instr(1, session("art_catalogo_vendita"), "1", vbTextCompare)>0) %>>
											catalogo articoli in vendita
										</td>
									</tr>
									<tr>
										<td class="content" colspan="2">
											<input type="checkbox" class="checkbox" name="search_catalogo_vendita" value="0" <%= chk(instr(1, Session("art_catalogo_vendita"), "0", vbTextCompare)>0) %>>
											catalogo articoli pubblici
										</td>
									</tr>
								<% end if %>
								<tr><th colspan="2" <%= Search_Bg("art_codice_int;art_codice_pro") %>>CODICI</th></tr>
								<tr>
									<td class="label">interno:</td>
									<td class="content">
										<input type="text" name="search_codice_int" value="<%= TextEncode(session("art_codice_int")) %>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<td class="label">produttore:</td>
									<td class="content">
										<input type="text" name="search_codice_pro" value="<%= TextEncode(session("art_codice_pro")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("art_nome") %>>NOME</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="text" name="search_nome" value="<%= TextEncode(session("art_nome")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("art_stato_catalogo") %>>STATO ARTICOLO A CATALOGO</th></tr>
								<tr>
									<td class="content" style="width:45%;">
										<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="1" <%= chk(instr(1, session("art_stato_catalogo"), "1", vbTextCompare)>0) %>>
										visibile
									</td>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="0" <%= chk(instr(1, Session("art_stato_catalogo"), "0", vbTextCompare)>0) %>>
										non visibile
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("art_tipo") %>>TIPO</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_tipo" value="AS" <%= chk(instr(1, Session("art_tipo"), "AS", vbTextCompare)>0) %>>
										articolo singolo
									</td>
								</tr>
								<tr>
									<td class="content varianti" colspan="2">
										<input type="checkbox" class="checkbox" name="search_tipo" value="AV" <%= chk(instr(1, Session("art_tipo"), "AV", vbTextCompare)>0) %>>
										articolo con varianti
									</td>
								</tr>
								<tr>
									<td class="content bundle" colspan="2">
										<input type="checkbox" class="checkbox" name="search_tipo" value="B" <%= chk(instr(1, Session("art_tipo"), "B", vbTextCompare)>0) %>>
										bundle
									</td>
								</tr>
								<tr>
									<td class="content confezione" colspan="2">
										<input type="checkbox" class="checkbox" name="search_tipo" value="C" <%= chk(instr(1, Session("art_tipo"), "C", vbTextCompare)>0) %>>
										confezione
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("art_aggregazione") %>>AGGREGAZIONE</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_aggregazione" value="A" <%= chk(instr(1, session("art_aggregazione"), "A", vbTextCompare)>0) %>>
										articolo singolo
									</td>
								</tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_aggregazione" value="B" <%= chk(instr(1, session("art_aggregazione"), "B", vbTextCompare)>0) %>>
										componente di un bundle
									</td>
								</tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_aggregazione" value="C" <%= chk(instr(1, session("art_aggregazione"), "C", vbTextCompare)>0) %>>
										componente di una confezione
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("art_accessori") %>>ARTICOLI COLLEGATI</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_accessori" value="S" <%= chk(instr(1, session("art_accessori"), "S", vbTextCompare)>0) %>>
										senza articoli collegati
									</td>
								</tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_accessori" value="C" <%= chk(instr(1, session("art_accessori"), "C", vbTextCompare)>0) %>>
										con articoli collegati
									</td>
								</tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_accessori" value="A" <%= chk(instr(1, session("art_accessori"), "A", vbTextCompare)>0) %>>
										collegato ad un articolo
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("art_categoria") %>>CATEGORIA</th></tr>
								<tr>
									<td class="content" colspan="2">
										<% CALL categorie.WritePicker("ricerca", "search_categoria", session("art_categoria"), false, true, 32) %>
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("art_marchio") %>>MARCHIO / PRODUTTORE</th></tr>
								<tr>
									<td class="content" colspan="2">
										<%	sql = "SELECT * FROM gtb_marche ORDER BY mar_nome_it"
										CALL dropDown(conn, sql, "mar_id", "mar_nome_it", "search_marchio", session("art_marchio"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
									</td>
								</tr>
								<%	sql = "SELECT spa_id, spa_nome_it  FROM gtb_spese_spedizione_articolo ORDER BY spa_id"
									rs_spe.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
									if rs_spe.RecordCount > 1 then 
								%>
										<tr><th colspan="2" <%= Search_Bg("art_spedizione") %>>TIPO SPEDIZIONE</th></tr>
										<tr>
											<td class="content" colspan="2">
												<%
												CALL dropDown(conn, sql, "spa_id", "spa_nome_it", "search_spedizione", session("art_spedizione"), false, " style=""width:100%;""", LINGUA_ITALIANO) 
												%>
											</td>
										</tr>
								<%end if 
								%>
								<tr><th colspan="2" <%= Search_Bg("art_variante") %>>VARIANTE</th></tr>
								<tr>
									<td class="content" colspan="2">
										<%	sql = "SELECT var_id, var_nome_it FROM gtb_varianti ORDER BY var_nome_it"
										CALL dropDown(conn, sql, "var_id", "var_nome_it", "search_variante", session("art_variante"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
									</td>
								</tr>
								
								<%	sql = " SELECT TOP 1 val_id, gtb_varianti.var_nome_it + ' - ' + gtb_valori.val_nome_it AS valore FROM gtb_valori INNER JOIN gtb_varianti " + _
										  " ON gtb_valori.val_var_id = gtb_varianti.var_id ORDER BY gtb_varianti.var_nome_it, gtb_valori.val_nome_it" 
								if GetValueList(conn, NULL, sql) then %>
									<tr><th colspan="2" <%= Search_Bg("art_valore_variante") %>>VALORE VARIANTE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<%	sql = Replace(sql, "TOP 1", "")
											CALL dropDown(conn, sql, "val_id", "valore", "search_valore_variante", session("art_valore_variante"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
										</td>
									</tr>
								<% end if %>
								<tr><th colspan="2" <%= Search_Bg("art_full_text") %>>FULL-TEXT (tutti i campi)</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="text" name="search_full_text" value="<%= TextEncode(session("art_full_text")) %>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<td class="footer" colspan="2">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
				</form>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco articoli
					</caption>
					<% if not rs.eof then 
						'costruisce query per calcolare il numero di varianti
						sql = right(Session("B2B_ARTICOLI_SQL"), len(Session("B2B_ARTICOLI_SQL")) - instr(1, Session("B2B_ARTICOLI_SQL"), "FROM ",vbTextCompare)+1)
						sql = left(sql, instr(1, sql, "ORDER BY", vbTextCompare)-1)
						sql = "SELECT COUNT(*) FROM grel_art_valori WHERE rel_Art_id IN ( SELECT art_id " + sql + " ) " %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> records ( <%= cIntero(GetValueList(conn, rsv, sql)) %> articoli ) in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="<%= IIF(rs("art_disabilitato"), "header_disabled", "header") %>" colspan="7">
												<%= rs("art_nome_it") %>
											</td>
										</tr>
										<tr>
											<td class="header" colspan="7">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr colspan="2">
														<td style="font-size: 1px; text-align:right;">
															<% CALL index.WriteButton("gtb_Articoli", rs("art_id"), POS_ELENCO) %> &nbsp;
														
															<a class="button" href="ArticoliMod.asp?ID=<%= rs("art_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<% 	sql = " SELECT TOP 1 rel_id FROM grel_art_valori INNER JOIN gtb_dettagli_ord " & _
																	  " ON grel_art_valori.rel_id = gtb_dettagli_ord.det_art_var_id " & _
																	  " WHERE rel_art_id = " & rs("art_id")
																if GetValueList(conn, rsc, sql) > 0 then %>
																<a class="button_disabled" title="Impossibile cancellare l'articolo perch&egrave; &egrave; presente in almeno un ordine">
																	CANCELLA
																</a>
															<%elseif rs("art_in_bundle") then %>
																<a class="button_disabled" title="Impossibile cancellare l'articolo perch&egrave; fa parte almeno di un bundle.">
																	CANCELLA
																</a>
															<% elseif rs("art_in_confezione") then%>
																<a class="button_disabled" title="Impossibile cancellare l'articolo perch&egrave; fa parte almeno di una confezione.">
																	CANCELLA
																</a>
															<% 	else
																if cInteger(rs("art_external_ID"))>0 then
																	sql = cInteger(rs("art_external_ID"))
																else
																	sql = "SELECT COUNT(*) FROM grel_art_valori WHERE rel_art_id=" & rs("art_id") & " AND IsNull(rel_external_id, 0)>0"
																	sql = cInteger(GetValueList(conn, rsc, sql))
																end if
															if sql > 0 then%>
																<a class="button_disabled" title="Impossbile cancellare l'articolo perch&egrave; &egrave; collegato ad almeno un articolo esterno. Per cancellarlo rimuovere prima il collegamento esterno.">
																	CANCELLA
																</a>
															<% else %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ARTICOLI','<%= rs("art_id") %>');">
																	CANCELLA
																</a>
															<% end if
															end if %>
														</td>
													</tr>
													<tr>
														<td style="font-size: 1px; text-align:right;">	
															<a class="button_l2" href="ArticoliGiacenze.asp?ID=<%= rs("art_id") %>">
																GIACENZE
															</a>
															&nbsp;
															<a class="button_l2" href="ArticoliPrezzi.asp?ID=<%= rs("art_id") %>">
																PREZZI
															</a>
															&nbsp;
															<% if Session("ATTIVA_FAQ_ARTICOLI") then %>
																<a class="button_l2" href="ArticoliFaq.asp?ID=<%= rs("art_id") %>">
																	FAQ
																</a>
																&nbsp;
															<% end if %>
															<% if Session("ATTIVA_COMMENTI") then %>
																<a class="button_l2" href="ArticoliCommenti.asp?ID=<%= rs("art_id") %>">
																	COMMENTI
																</a>
																&nbsp;
															<% end if %>
														</td>
													</tr>
												</table>
												
											</td>
										</tr>
										<tr>
											<td class="label">codice:</td>
											<td class="label" style="width:9%;">interno:</td>
											<td class="content_b" style="width:16%;"><%= rs("art_cod_int") %></td>
											<td class="label" style="width:13%;">alternativo:</td>
											<td class="content" style="width:16%;"><%= rs("art_cod_alt") %></td>
											<td class="label" style="width:12%;">produttore:</td>
											<td class="content" style="width:19%;"><%= rs("art_cod_pro") %></td>
										</tr>
										<tr>
											<td class="label">categoria:</td>
											<td class="content" colspan="6">
												<%= categorie.NomeCompleto(rs("art_tipologia_id")) %>
											</td>
										</tr>
										<% if cInteger(rs("art_raggruppamento_id"))>0 then 
											sql = "SELECT rag_nome_it FROM gtb_tipologie_raggruppamenti WHERE rag_id=" & rs("art_raggruppamento_id")%>
											<tr>
												<td class="label">raggruppamento:</td>
												<td class="content" colspan="6">
													<%= GetValueList(conn, rsc, sql) %>
												</td>
											</tr>
										<% end if %>
										<tr>
											<td class="label">marca:</td>
											<td class="content" colspan="4">
												<%= rs("mar_nome_it") %>
											</td>
											<td class="label">stato:</td>
											<td class="content">
												<% if rs("art_disabilitato") then %>
													non a catalogo
												<% else %>
													a catalogo
												<% end if %>
											</td>
										</tr>
										<tr>
											<td class="label">tipo:</td>
											<% if rs("art_se_bundle") then %>
												<td class="content bundle" colspan="2">bundle di articoli</td>
											<% elseif rs("art_se_confezione") then %>
												<td class="content confezione" colspan="2">confezione di articoli</td>
											<% elseif rs("art_varianti") then %>
												<td class="content varianti" colspan="2">articolo con varianti</td>
											<% else %>
												<td class="content" colspan="2">articolo singolo</td>
											<% end if %>
											<% txt = ""
											title = ""
											if rs("art_in_bundle") then
												txt = txt &"in bundle - "
											end if
											if rs("art_in_confezione") then
												txt = txt &"in confezione - "
											end if
											if txt <> "" then txt = left(txt, len(txt)-3)%>
											<td class="content" colspan="2">
												<%= txt %>
											</td>
											<% txt = ""
											if rs("art_se_accessorio") then
												sql = " SELECT art_nome_it + ' (cod:' + art_cod_int + ')' FROM gtb_articoli where art_id IN ( " & _
													  " select aa_art_id from grel_art_acc " & _
													  " where aa_acc_id = " & rs("art_id") & _
													  " )"
												title = GetValueList(conn, NULL, sql)
												txt = txt &"collegato ad altro <strong title="""&title&""">articolo</strong> - "
											end if
											if rs("art_noVenSingola") then
												txt = txt &"non vendibile singolarmente - "
											end if
											if txt <> "" then txt = left(txt, len(txt)-3) %>
											<td class="content" colspan="2">
												<%= txt %>
											</td>
										</tr>
										<% 	if rs("art_se_bundle") OR rs("art_se_confezione") then 
											sql = " SELECT art_cod_int, art_nome_it, bun_quantita, art_varianti, rel_id, art_id, rel_cod_int FROM gtb_articoli " + _
												  " INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + _
												  " INNER JOIN gtb_bundle ON grel_art_valori.rel_id = gtb_bundle.bun_articolo_id " + _
												  " WHERE gtb_bundle.bun_bundle_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & rs("art_id") & ") " + _
												  " ORDER BY art_nome_it"
										
											rsc.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
											if not rsc.eof then%>
												<tr>
													<td class="label">componenti:</td>
													<td colspan="6">
														<% if rsc.recordcount>2 then %> 
															<span class="overflow">
														<% end if  %>
														<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
															<tr>
																<th class="l2_center" width="14%">codice</th>
																<th class="L2">nome</th>
																<th class="l2_center" width="8%">quantit&agrave;</th>
															</tr>
															<% while not rsc.eof %>
																<tr>
																	<% if rsc("art_varianti") then %>
																		<td class="content"><%= rsc("rel_cod_int")%></td>
																		<td class="content">
																			<%CALL ArticoloLink(rsc("art_id"), rsc("art_nome_it"), rsc("rel_cod_int"))
																			CALL ListValoriVarianti(conn, rsv, rsc("rel_id"))%>
																		</td>
																	<% else %>
																		<td class="content"><%= rsc("art_cod_int")%></td>
																		<td class="content"><%CALL ArticoloLink(rsc("art_id"), rsc("art_nome_it"), rsc("art_cod_int")) %></td>
																	<% end if %>
																	<td class="content_center"><%= rsc("bun_quantita")%></td>
																</tr>
																<%rsc.movenext
															wend %>
														</table>
														<% if rsc.recordcount>2 then %>
															</span>
														<% end if %>
													</td>
												</tr>
											<%end if
											rsc.close
										end if 
										
										CALL ListaCollegamentiArticolo(conn, rsc, rs("art_id"), rs("art_ha_accessori"), false) %>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend %>
						<tr>
							<td class="footer" style="border-top:0px; text-align:left;">
								<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
							</td>
						</tr>
					<%else%>
						<tr><td class="noRecords">Nessun record trovato</th></tr>
					<% end if %>
				</table>
			</td> 
		</tr>
		<tr><td>&nbsp;</td></tr>
	</table>		
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsc = nothing
set rsv = nothing
set conn = nothing
%>
