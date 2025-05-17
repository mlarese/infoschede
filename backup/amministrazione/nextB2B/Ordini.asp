<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" --> 
<!--#INCLUDE VIRTUAL="amministrazione/library/ExportTools.asp" -->
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione ordini - elenco"
dicitura.puls_new = "NUOVO ORDINE"
dicitura.link_new = "OrdiniNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, rsc, sql, pager, i, query_per_export, rsp

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator


'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("ord_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("ord_")
	end if
end if

'filtra per agente
if session("ord_agente") <> "" then
	sql = sql &" AND riv_agente_id="& session("ord_agente")
end if

'filtra per nome cliente
if Session("ord_denominazione")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("ord_denominazione"))
end if

'filtra per indirizzo di destinazione
if Session("ord_indirizzo")<>"" then
	sql = sql & " AND ord_id IN (SELECT det_ord_id FROM gtb_dettagli_ord INNER JOIN tb_indirizzario " + _
								" ON gtb_dettagli_ord.det_ind_id = tb_indirizzario.IDElencoIndirizzi WHERE " + _
                                SQL_FullTextSearch_Contatto_Indirizzo(conn, Session("ord_indirizzo")) + _
								" )"
end if

'filtra per citta
if Session("ord_citta")<>"" then
	sql = sql & " AND ord_id IN (SELECT det_ord_id FROM gtb_dettagli_ord INNER JOIN tb_indirizzario " + _
								" ON gtb_dettagli_ord.det_ind_id = tb_indirizzario.IDElencoIndirizzi WHERE " + _
								sql_FullTextSearch(Session("ord_citta"), "CittaElencoIndirizzi") + _
								" )"
end if

'filtra per articolo
if Session("ord_articolo")<>"" then
	sql = sql & " AND ord_id IN ( SELECT det_ord_id FROM gtb_dettagli_ord INNER JOIN grel_art_valori ON gtb_dettagli_ord.det_art_var_id = grel_art_valori.rel_id " + _
							    " WHERE gtb_dettagli_ord.det_ord_id = gtb_ordini.ord_id AND " & sql_FullTextSearch(Session("ord_articolo"),"rel_cod_int;rel_cod_alt;rel_cod_pro") & " ) "
end if

'filtra per codice
if session("ord_codice") <> "" then
	sql = sql & " AND (" & sql_FullTextSearch(Session("ord_codice"),"ord_cod") & " OR " & sql_FullTextSearch(Session("ord_codice"),"ord_id") & ") "
end if

'filtra per stato
if session("ord_stato") <> "" then
	sql = sql & " AND ord_stato_id="& session("ord_stato")
end if

'metodo di pagamento
if session("ord_metodopagamento") <> "" then
	sql = sql & " AND ord_modopagamento_id="& session("ord_metodopagamento")
end if

'filtra per data ordine
if isDate(Session("ord_data_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "ord_data", adCompareGreaterThan, Session("ord_data_from"))
end if
if isDate(Session("ord_data_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "ord_data", adCompareLessThan, Session("ord_data_to"))
end if

'ricerca per tipo
if Session("ord_tipo")<>"" then
	sql = sql + " AND ( "
	'non confermato
	if instr(1, Session("ord_tipo"), ORDINE_NON_CONFERMATO, vbTextCompare)>0  then
		sql = sql + "(NOT " + SQL_isTrue(conn, "ord_impegna") + _
					" AND NOT " + SQL_isTrue(conn, "ord_movimenta") + _
					" AND NOT " + SQL_isTrue(conn, "ord_archiviato") + _
		     		") OR "
	end if
	'confermato
	if instr(1, Session("ord_tipo"), ORDINE_CONFERMATO, vbTextCompare)>0  then
		sql = sql + "(" + SQL_isTrue(conn, "ord_impegna") + _
					" AND NOT " + SQL_isTrue(conn, "ord_movimenta") + _
					" AND NOT " + SQL_isTrue(conn, "ord_archiviato") + _
		     		") OR "
	end if
	'evaso
	if instr(1, Session("ord_tipo"), ORDINE_EVASO, vbTextCompare)>0  then
		sql = sql + "(" + SQL_isTrue(conn, "ord_movimenta") + _
					" AND NOT " + SQL_isTrue(conn, "ord_impegna") + _
					" AND NOT " + SQL_isTrue(conn, "ord_archiviato") + _
		     		") OR "
	end if
	'concluso
	if instr(1, Session("ord_tipo"), ORDINE_ARCHIVIATO, vbTextCompare)>0  then
		sql = sql & "(" & SQL_isTrue(conn, "ord_archiviato") & ") OR "
	end if
	sql = left(sql, len(sql)-3) & " )"
end if




'......................................................................................................
'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
CALL ADDON__ORDINI__ricerca_form_parse(conn, sql)
'......................................................................................................




'ricerca full text
if Session("ord_full_text")<>"" then
	sql = sql & " AND ((" + sql_FullTextSearch(Session("ord_full_text"),"ord_Note;mag_nome") + _
                       " OR " + SQL_FullTextSearch_Contatto_Nominativo(conn, Session("ord_full_text")) + _
				") OR ord_id IN (SELECT det_ord_id FROM gtb_dettagli_ord INNER JOIN tb_indirizzario " + _
								" ON gtb_dettagli_ord.det_ind_id = tb_indirizzario.IDElencoIndirizzi WHERE " + _
                                SQL_FullTextSearch_Contatto_Nominativo(conn, Session("ord_full_text")) + _
                                " OR " + SQL_FullTextSearch_Contatto_indirizzo(conn, Session("ord_full_text")) + _
								" ))"
end if


sql = " SELECT ord_id, ord_archiviato, ord_movimenta, ord_impegna, ord_data, ord_cod, IDElencoIndirizzi, so_stato_ordini, so_nome_it, " & _
	  " ord_totale, ord_totale_iva, ord_totale_spese, ord_totale_spese_iva, ord_modopagamento_id, isSocieta, NomeOrganizzazioneElencoIndirizzi, " & _
	  " CognomeElencoIndirizzi, NomeElencoIndirizzi, ord_exported " & _
	  " FROM (gtb_ordini INNER JOIN gv_rivenditori ON gtb_ordini.ord_riv_id = gv_rivenditori.riv_id) " + _
	  " INNER JOIN gtb_stati_ordine ON gtb_ordini.ord_stato_id=gtb_stati_ordine.so_id " + _
	  " INNER JOIN gtb_magazzini ON gtb_ordini.ord_magazzino_id = gtb_magazzini.mag_id " + _
	  " WHERE (1=1) " & sql
	  
query_per_export = "SELECT ord_id " & Right(sql, Len(sql) - InStr(1, sql, "FROM", 0) + 2)
	  
sql = sql & " ORDER BY ord_id DESC, ord_data DESC"
session("B2B_ORDINI_SQL") = sql


CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
				<form action="" method="post" id="ricerca" name="ricerca">
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
									<caption>Opzioni di ricerca</caption>
									<tr>
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("ord_codice") %>>RIFERIMENTO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_codice" value="<%= TextEncode(session("ord_codice")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("ord_tipo") %>>STATO ORDINE</td></tr>
									<% for i= lbound(STATI_ORDINE) to ubound(STATI_ORDINE) %>
										<tr>
											<td class="content<%= STILI_STATI_ORDINE(i) %>" colspan="2">
												<input type="checkbox" class="checkbox" name="search_tipo" value="<%= i %>" <%= chk(instr(1, session("ord_tipo"), cString(i), vbTextCompare)>0) %>>
												<%= STATI_ORDINE(i) %>
											</td>
										</tr>
									<% next %>
									<tr><th colspan="2" <%= Search_Bg("ord_stato") %>>STATO DI LAVORAZIONE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% sql = "SELECT so_id, so_nome_it FROM gtb_stati_ordine ORDER BY so_stato_ordini, so_ordine"
											CALL dropDown(conn, sql, "so_id", "so_nome_it", "search_stato", session("ord_stato"), false, " style=""width:100%;""", LINGUA_ITALIANO) %> 
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("ord_metodopagamento") %>>METODO DI PAGAMENTO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% sql = "SELECT mosp_id, mosp_nome_it FROM gtb_modipagamento ORDER BY mosp_nome_it"
											CALL dropDown(conn, sql, "mosp_id", "mosp_nome_it", "search_metodopagamento", session("ord_metodopagamento"), false, " style=""width:100%;""", LINGUA_ITALIANO) %> 
										</td>
									</tr>
									<% 
									
									
									'......................................................................................................
									'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
									CALL ADDON__ORDINI__ricerca_form(conn, rs)
									'......................................................................................................
									
									
									 %>
									<tr><th colspan="2" <%= Search_Bg("ord_denominazione") %>>CLIENTE</th></tr>
									<tr>
										<td class="content" colspan="2">
										<input type="text" name="search_denominazione" value="<%= TextEncode(session("ord_denominazione")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("ord_indirizzo;ord_citta") %>>DESTINAZIONE MERCE</th></tr>
									<tr>
										<td class="label">
											indirizzo:
										</td>
										<td class="content">
											<input type="text" name="search_indirizzo" value="<%= TextEncode(session("ord_indirizzo")) %>" style="width:100%;">
										</td>
									</tr>
									<tr>
										<td class="label">
											citt&agrave;:
										</td>
										<td class="content">
											<input type="text" name="search_citta" value="<%= TextEncode(session("ord_citta")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("ord_agente") %>>AGENTE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% 	sql = " SELECT IDElencoIndirizzi, "& _
													  "(CognomeElencoIndirizzi "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) &" NomeElencoIndirizzi "& SQL_concat(conn) &" ' - ' " & SQL_concat(conn) &" NomeOrganizzazioneElencoIndirizzi) AS NOMINATIVO " &_
													  " FROM gv_agenti " & _
													  " ORDER BY ModoRegistra"
												CALL dropDown(conn, sql, "IDElencoIndirizzi", "NOMINATIVO", "search_agente", session("ord_agente"), false, " style=""width:100%;""", LINGUA_ITALIANO)
											%> 
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("ord_articolo") %>>ARTICOLO</th></tr>
									<tr>
										<td class="content" colspan="2">
										<input type="text" name="search_articolo" value="<%= TextEncode(session("ord_articolo")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("ord_data_from;ord_data_to") %>>DATA</td></tr>
									<tr><td class="label" colspan="2">a partire dal:</td></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_from", Session("ord_data_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><td class="label" colspan="2">fino al:</td></tr>
									<tr>
										<td class="content" colspan="2">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_to", Session("ord_data_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("ord_full_text") %>>FULL-TEXT</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_full_text" value="<%= TextEncode(session("ord_full_text")) %>" style="width:100%;">
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
					
						<%					
						sql = "SELECT     gtb_ordini.ord_id AS NUMERO_ORDINE, gtb_ordini.ord_cod as RIFERIMENTO_ORDINE, gtb_ordini.ord_data AS DATA_ORDINE, " & vbcrlf & _
							  "           ISNULL(gv_rivenditori.NomeOrganizzazioneElencoIndirizzi + ' - ', '') + gv_rivenditori.NomeElencoIndirizzi + ' - ' + gv_rivenditori.CognomeElencoIndirizzi AS NOME_CLIENTE, " & vbcrlf & _
							  "                (SELECT     mosp_nome_it " & _
							  "                 FROM          gtb_modipagamento " & _
							  "                 WHERE      (mosp_id = gtb_ordini.ord_modopagamento_id)) AS TIPO_PAGAMENTO, " & vbcrlf & _
							  "               (SELECT     so_nome_it " & _
							  "                 FROM          gtb_stati_ordine " & _
							  "                 WHERE      (gtb_ordini.ord_stato_id = so_id)) AS STATO_ORDINE, " & vbcrlf & _
							  "               (SELECT     SUM(det_prezzo_unitario * det_qta) " & _
							  "                 FROM          gv_dettagli_ord " & _
							  "                 WHERE      (det_ord_id = gtb_ordini.ord_id)) AS SUBTOTALI, " & vbcrlf & _
							  "           CONVERT(money, (SELECT     SUM(det_prezzo_unitario * det_qta * det_iva) / 100 " & _
							  "                 FROM          gv_dettagli_ord AS gv_dettagli_ord_1 " & _
							  "                 WHERE      (det_ord_id = gtb_ordini.ord_id))) AS IVA, " & vbcrlf & _
							  "			  CONVERT(money, gtb_ordini.ord_spesespedizione) AS SPESE_SPEDIZIONE, " & vbcrlf & _
							  "           CONVERT(money, gtb_ordini.ord_speseincasso) AS SPESE_INCASSO, " & vbcrlf & _
							  "           CONVERT(money, gtb_ordini.ord_spesefisse) AS SPESE_FISSE, " & vbcrlf & _
							  "           CONVERT(money, gtb_ordini.ord_spesealtre) AS ALTRE_SPESE, " & vbcrlf & _
							  "           CONVERT (money," & _
							  "               (SELECT     SUM(det_prezzo_unitario * det_qta + det_prezzo_unitario * det_qta * det_iva / 100) AS Expr1 " & _
							  "                   FROM         gv_dettagli_ord AS gv_dettagli_ord_2 " & _
							  "                   WHERE     (det_ord_id = gtb_ordini.ord_id)) + ISNULL(gtb_ordini.ord_spesealtre, 0) + ISNULL(gtb_ordini.ord_spesefisse, 0) + ISNULL(gtb_ordini.ord_speseincasso, 0) " & _
							  "           + ISNULL(gtb_ordini.ord_spesespedizione, 0)) AS TOTALE " & vbcrlf & _
							  " FROM         gv_rivenditori INNER JOIN " & vbcrlf & _
							  "           gtb_ordini ON gv_rivenditori.riv_id = gtb_ordini.ord_riv_id " & vbcrlf & _
							  " WHERE gtb_ordini.ord_id IN (" & query_per_export & ") " & _
							  " ORDER BY NUMERO_ORDINE"
							  
						Session("STAT_SQL") = sql
						%>
						<tr><td style="font-size:4px;">&nbsp;</td></tr>
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
									<caption>Strumenti</caption>
									<tr>
										<th>Totali ordini</th>
									</tr>
									<% session("B2B_EXPORT_TOT_ORDINI_SQL") = query_per_export %>
									<tr>
										<td class="content_center">
											<a <%=IIF(not rs.eof, "", "disabled")%> style="display:block; width:94%; text-align:center; line-height:12px;" class="button<%=IIF(not rs.eof, "", "_disabled")%>"
												title="Visualizza i totali degli ordini" href="javascript:void(0);" 
												onClick="OpenAutoPositionedWindow('OrdiniViewTotali.asp?qry=B2B_EXPORT_TOT_ORDINI_SQL', 'export', 500, 310)" <%= ACTIVE_STATUS %>>
												VISUALIZZA TOTALI ORDINI
											</a>
										</td>
									</tr>
									<tr>
										<td class="content_center" style="padding-top:4px; padding-bottom:4px;">
											<%CALL WRITE_EXPORT_LINK("ESPORTA TOTALI PER ORDINE", "DATA_ConnectionString", "STAT_SQL", FORMAT_EXCEL_FILE, false) %>
										</td>
									</tr>
									<tr>
										<th>Articoli ordinati</th>
									</tr>
									<%
									sql = "SELECT (SELECT art_cod_int FROM gv_articoli WHERE rel_id = det_art_var_id) AS CODICE, " + vbCrlf + _
										  " (SELECT art_nome_it FROM gv_articoli WHERE rel_id = det_art_var_id) AS NOME, " + vbCrlf + _
										  " SUM(det_qta) AS QTA," + vbCrlf + _
										  " SUM(det_totale) AS IMPORTO" + vbCrlf + _
										  "FROM gtb_dettagli_ord" + vbCrlf + _
										  "WHERE det_ord_id IN (" & query_per_export & ") " + vbCrlf + _
										  "GROUP BY det_art_var_id" + vbCrlf + _
										  "ORDER BY CODICE"
									Session("QTA_ART_ORD_SQL") = sql
									%>
									<tr>
										<td class="content_center" style="padding-top:4px; padding-bottom:4px;">
											<%CALL WRITE_EXPORT_LINK("ESPORTA ARTICOLI ORDINATI", "DATA_ConnectionString", "QTA_ART_ORD_SQL", FORMAT_EXCEL_FILE, false) %>
										</td>
									</tr>
									<tr>
										<th>Clienti che hanno ordinato</th>
									</tr>
									<tr>
										<td class="content_center" style="padding-top:4px; padding-bottom:4px;">
											<% 
											sql = session("B2B_ORDINI_SQL")
											sql = "" & right(sql, len(sql) + 1 - instr(1, sql, "FROM gv_rivenditori", vbTextCompare))
											Session("RIVENDITORI_ELENCO_SQL") = sql
											%>
											<% CALL ExportContattiInRubrica(sql, "IDElencoIndirizzi", "", "") %>
										</td>
									</tr>
									<tr>
										<td class="content_center" style="padding-top:4px; padding-bottom:4px;">
											<% CALL WRITE_CONTATTI_EXPORT_LINK("ESPORTA ELENCO CLIENTI", "RIVENDITORI_ELENCO_SQL", FORMAT_EXCEL_FILE, false) %>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco ordini
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
										<%if rs("ord_archiviato") then%>
											<td class="header_disabled" colspan="4">
										<% elseif rs("ord_movimenta") then %>
											<td class="header OrdEvaso" colspan="4">
										<% elseif rs("ord_impegna") then %>
											<td class="header OrdConfermato" colspan="4">
										<% else %>
											<td class="header" colspan="4">
										<% end if %>	
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="OrdiniMod.asp?ID=<%= rs("ord_id") %>">
																MODIFICA
															</a>
															&nbsp;
														<% 	if rs("ord_movimenta") OR rs("ord_impegna") then %>
															<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare l'ordine: merce gi&agrave; movimentati">
																CANCELLA
															</a>
														<% 	elseif cBoolean(rs("ord_exported"), false) then %>
															<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare l'ordine: ordine gi&agrave; esportato">
																CANCELLA
															</a>
														<% 	else %>
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ORDINI','<%= rs("ord_id") %>');">
																CANCELLA
															</a>
														<% 	end if %>
														</td>
													</tr>
												</table>
												<%= rs("ord_id") %>
												 - 
												<%= DateIta(rs("ord_data")) %>
												<%= IIF(CString(rs("ord_cod")) <> "", " - " & rs("ord_cod"), "")%>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:21%;">cliente:</td>
											<td class="content" colspan="3">
												<% CALL ClienteLink(rs("IDElencoIndirizzi") , ContactFullName(rs)) %>
											</td>
										</tr>
										<tr>
											<td class="label">stato lavorazione:</td>
											<td class="content<%= STILI_STATI_ORDINE(rs("so_stato_ordini")) %>"><%= rs("so_nome_it") %></td>
											<% sql = "SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_ord_id=" & rs("ord_id")
											if cInteger(GetValueList(conn, rsc, sql))>0 then%>
												<td class="label">totale:</td>
												<td class="content">													
													<%= FormatPrice(cReal(rs("ord_totale"))+cReal(rs("ord_totale_iva"))+cReal(rs("ord_totale_spese"))+cReal(rs("ord_totale_spese_iva")) , 2, true) %>&euro;
												</td>
											</tr>
											<tr>
												<td class="label" style="width:21%;">metodo pagamento:</td>
												<td class="content" colspan="3">
													<%= GetValueList(conn, NULL, "SELECT mosp_nome_it FROM gtb_modipagamento WHERE mosp_id = " & rs("ord_modopagamento_id")) %>
												</td>
											</tr>
											<tr>
												<td class="label">destinazioni merce:</td>
												<td colspan="3">
													<% sql = " SELECT IDElencoIndirizzi, indirizzoElencoIndirizzi, capElencoIndirizzi, cittaElencoIndirizzi, statoProvElencoIndirizzi" & _
															 " FROM tb_indirizzario WHERE IDElencoIndirizzi IN " + _
												  	 		 " (SELECT DISTINCT det_ind_id FROM gtb_dettagli_ord WHERE det_ord_id=" & rs("ord_id") & ") " + _
															 " ORDER BY CntRel, IndirizzoElencoIndirizzi"
													rsc.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
													if rsc.recordcount>2 then %> 
														<span class="overflow">
													<% end if%>
													<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
														<tr>
															<th class="L2">indirizzo</th>
															<th class="L2">cap</th>
															<th class="L2">citt&agrave;</th>
															<th class="L2" width="28%">provincia / nazione</th>
														</tr>
														<% while not rsc.eof %>
															<tr>
																<td class="content">
																	<%= rsc("indirizzoElencoIndirizzi") %>
																	<% if rsc("IDElencoIndirizzi") = rs("IDElencoIndirizzi") then %>
																		(sede principale)
																	<% end if %>
																</td>
																<td class="content"><%= rsc("capElencoIndirizzi") %></td>
																<td class="content"><%= rsc("cittaElencoIndirizzi") %></td>
																<td class="content">
																	<%= rsc("statoProvElencoIndirizzi") %>
																	<%= IIF(cString(rsc("statoProvElencoIndirizzi"))<>"" AND cString(rsc("statoProvElencoIndirizzi"))<>"", " - ", "") %>
																	<%= rsc("statoProvElencoIndirizzi") %>
																</td>
															</tr>
															<%rsc.movenext
														wend %>
													</table>
													<% if rsc.recordcount>2 then %>
														</span>
													<% end if
													rsc.close %>
												</td>
											</tr>
											<%
											dim n_pagamenti
											n_pagamenti = cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_pagamenti WHERE pag_ordine_id = " & rs("ord_id")))
											%>
											<%if n_pagamenti>0 then%>
											<tr>
												<td class="label">pagamenti:</td>
												<td colspan="3">
													<% sql = " SELECT pag_id, pag_data, pag_importo, pag_ordine_id, pag_stato_ordine_id, pag_mosp_id, pag_RAW, mosp_nome_it, so_nome_it " & _
															 " FROM gtb_pagamenti " & _
														     " INNER JOIN gtb_modipagamento ON gtb_pagamenti.pag_mosp_id = gtb_modipagamento.mosp_id " & _
															 " INNER JOIN gtb_stati_ordine ON gtb_pagamenti.pag_stato_ordine_id = gtb_stati_ordine.so_id " & _
															 " WHERE pag_ordine_id = " & rs("ord_id")
													rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
													if rsp.recordcount>2 then %> 
														<span class="overflow">
													<% end if%>
													<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
														<tr>
															<th class="L2">metodo pagamento</th>
															<th class="L2">data</th>
															<th class="L2">importo</th>
															<th class="L2" width="28%">stato ordine</th>
														</tr>
														<% while not rsp.eof %>
															<tr>
																<td class="content"><%= rsp("mosp_nome_it") %></td>
																<td class="content"><%= rsp("pag_data") %></td>
																<td class="content">
																	<%= FormatPrice(cReal(rsp("pag_importo")), 2, true) %>&euro;
																</td>
																<td class="content"><%= rsp("so_nome_it") %></td>
															</tr>
															<%rsp.movenext
														wend %>
													</table>
													<% if rsp.recordcount>2 then %>
														</span>
													<% end if
													rsp.close %>
												</td>
											</tr>
											<%end if%>
											<% else %>
												<td class="content" colspan="2">&nbsp;</td>
											</tr>
											<tr>
												<td colspan="4" class="content_b">
													Ordine non completo: Nessun dettaglio trovato
												</td>
											</tr>
											<% end if 
											'......................................................................................................
											'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
											CALL ADDON__ORDINI__record_elenco(conn, rs)
											'......................................................................................................
											%>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
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
set conn = nothing%>
