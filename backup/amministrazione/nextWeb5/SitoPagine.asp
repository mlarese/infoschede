<%@ Language=VBScript %>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
Imposta_Proprieta_Sito("ID")
%>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<%
dim conn, sql, pager, rs, i, lingua, rsA, daPubblicare, templateSemplificato, riepilogo, indexAlberoPag, CanDelete

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsA = Server.CreateObject("ADODB.RecordSet")

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - indice delle pagine - elenco"
if index.ChkPrm(prm_pagine_altera, 0) then
	dicitura.puls_new = "INDIETRO A SITI;NUOVA PAGINA"
	dicitura.link_new = "Siti.asp;SitoPagineNew.asp"
else
	dicitura.puls_new = "INDIETRO A SITI"
	dicitura.link_new = "Siti.asp"
end if
dicitura.scrivi_con_sottosez()

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("pa_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("pa_")
	end if
elseif Session("WEB_PAGINE_SQL") = "" then
	session("pa_archiviata") = "0"
end if

'filtra per titolo
if Session("pa_titolo")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("pa_titolo"), FieldLanguageList("nome_ps_") + ";" + SQL_concatFields(conn, "nome_ps_it;nome_ps_interno")) & ")"
end if

'filtra per data modifica
if isDate(Session("pa_data_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "ps_modData", adCompareGreaterThan, Session("pa_data_from"))
end if
if isDate(Session("pa_data_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "ps_modData", adCompareLessThan, Session("pa_data_to"))
end if

'filtro per url
dim search_url
if Session("pa_url_exact")<>""then
	'ricerca esatta della chiave
	search_url = replace(trim(Session("pa_url_exact")), "\", "/")
	search_url = replace(search_url, "/", 1, 1, vbTextCompare)
	search_url = replace(search_url, "/", len(search_url)-1, 1, vbTextCompare)
	riepilogo = split(search_url, "/")
	search_url = "/" & riepilogo(ubound(riepilogo)) & "/"
	
	sql = sql & " AND ( id_pagineSito IN ( SELECT idx_link_pagina_id FROM v_indice " + _
										 " WHERE '/'"+ SQL_concat(conn) + "idx_link_url_rw_IT LIKE '%" + ParseSQL(search_url, adChar) + "' " + _
										 	" OR '/'"+ SQL_concat(conn) + "idx_link_url_rw_EN LIKE '%" + ParseSQL(search_url, adChar) + "' " + _
											" OR '/'"+ SQL_concat(conn) + "idx_link_url_rw_FR LIKE '%" + ParseSQL(search_url, adChar) + "' " + _
											" OR '/'"+ SQL_concat(conn) + "idx_link_url_rw_DE LIKE '%" + ParseSQL(search_url, adChar) + "' " + _
											" OR '/'"+ SQL_concat(conn) + "idx_link_url_rw_ES LIKE '%" + ParseSQL(search_url, adChar) + "' " + _
											" OR '/'"+ SQL_concat(conn) + "idx_link_url_rw_RU LIKE '%" + ParseSQL(search_url, adChar) + "' " + _
											" OR '/'"+ SQL_concat(conn) + "idx_link_url_rw_CN LIKE '%" + ParseSQL(search_url, adChar) + "' " + _
											" OR '/'"+ SQL_concat(conn) + "idx_link_url_rw_PT LIKE '%" + ParseSQL(search_url, adChar) + "' " + _
										 " )) "
end if

if Session("pa_url_fulltext")<>""then
	'ricerca fulltext
	search_url = replace(trim(Session("pa_url_fulltext")), "\", "/")
	search_url = replace(search_url, "/", 1, 1, vbTextCompare)
	search_url = replace(search_url, "/", len(search_url)-1, 1, vbTextCompare)
	sql = sql & " AND ( id_pagineSito IN (SELECT idx_link_pagina_id FROM tb_contents_index WHERE " & SQL_FullTextSearch(search_url, FieldLanguageList("idx_link_url_rw_")) & ") )"
end if


'filtra per numero della paginasito
if cIntero(Session("pa_numero_ps"))>0 then
	sql = sql & " AND id_pagineSito=" & Session("pa_numero_ps")
end if
'filtra per numero della pagina di navigazione / lavoro
if cIntero(Session("pa_numero_pagina"))>0 then
	sql = sql & " AND ( "
	for each lingua in application("LINGUE")
		if session("LINGUA_" + lingua) then
			sql = sql + " id_pagStage_" + lingua + "=" & Session("pa_numero_pagina") & " OR " + _
						" id_pagDyn_" + lingua + "=" & Session("pa_numero_pagina") & " OR "
		end if
	next
	sql = left(sql, len(sql)-3) & " )"
end if


'filtra per indice
if CIntero(session("pa_idx")) > 0 then
	sql = sql &" AND EXISTS (SELECT 1 FROM tb_contents c"& _
						   " INNER JOIN tb_contents_index i ON c.co_id = i.idx_content_id"& _
						   " WHERE co_F_key_id = id_pagineSito AND co_F_table_id = "& index.GetTable("tb_pagineSito") & _
						   " AND idx_padre_id = "& session("pa_idx") &")"
end if

'filtra per archiviata
if session("pa_archiviata") = "1" then
	sql = sql &" AND "& SQL_IsTrue(conn, "archiviata")
elseif session("pa_archiviata") = "0" then
	sql = sql &" AND NOT "& SQL_IsTrue(conn, "archiviata")
end if

'filtra per riservata
if session("pa_riservata") = "1" then
	sql = sql &" AND "& SQL_IsTrue(conn, "riservata")
elseif session("pa_riservata") = "0" then
	sql = sql &" AND NOT "& SQL_IsTrue(conn, "riservata")
end if


'filtra per indicizzabile
if session("pa_indicizzabile") = "1" then
	sql = sql &" AND "& SQL_IsTrue(conn, "indicizzabile")
elseif session("pa_indicizzabile") = "2" then
	sql = sql &" AND NOT "& SQL_IsTrue(conn, "indicizzabile")
end if


'filtra per template
if CIntero(session("pa_template")) > 0 then
	sql = sql &" AND EXISTS (SELECT 1 FROM tb_pages WHERE id_paginaSito = id_pagineSito AND id_template = "& session("pa_template") &")"
end if

'filtra per contenuto
if session("pa_testo") <> "" OR session("pa_img") <> "" OR session("pa_plugin") <> "" then
	sql = sql &" AND EXISTS (SELECT 1 FROM tb_pages p"& _
			   "			 INNER JOIN tb_layers l ON p.id_page = l.id_pag"& _
			   "			 WHERE id_paginaSito = id_pagineSito"
			   
	'filtra per testo
	if session("pa_testo") <> "" then
		sql = sql &" AND (id_tipo = " & LAYER_TEXT & " OR id_tipo = " & LAYER_S_TEXT & ") AND (" & SQL_FullTextSearch(Session("pa_testo"), "testo") & ")"
	end if
	'filtra per file
	if session("pa_img") <> "" then
		if Left(session("pa_img"), 1) = "/" then
			session("pa_img") = Right(session("pa_img"), Len(session("pa_img")) - 1)
		end if
		sql = sql &" AND (id_tipo = " & LAYER_IMAGE & " OR id_tipo = " & LAYER_FLASH & ") AND nome = '"& ParseSQL(_
			Replace(Replace(Session("pa_img"), "flash/", ""), "images/", ""), adChar) &"'"
	end if
	'filtra per plugin
	if session("pa_plugin") <> "" then
		sql = sql &" AND id_tipo = " & LAYER_OBJECT & " AND " + _
				   " ( id_objects = "& ParseSQL(Session("pa_plugin"), adNumeric) & " OR " & _
				   "   testo LIKE '%=%" + ParseSQL(GetValueList(conn, null, "SELECT name_objects FROM tb_objects WHERE id_objects =" & Session("pa_plugin")), adChar) + "%;%'" & _
				   "  ) "
	end if
	
	sql = sql &")"
end if

'filtra per stato
if CIntero(session("pa_stato")) > 0 then
	sql = sql &" AND"
	if session("pa_stato") = "1" then
		sql = sql &" NOT"
	end if
	
	dim condition, conditionLingue
	i = 0
	for each lingua in application("LINGUE")
		if session("LINGUA_" + lingua) then
			condition = condition + " OR id_page = s.id_pagStage_" + lingua
			conditionLingue = conditionLingue + " OR lingua = '"& lingua &"'"
			i = i + 1
		end if
	next
	condition = Right(condition, Len(condition) - 4)
	conditionLingue = Right(conditionLingue, Len(conditionLingue) - 4)
	sql = sql &" EXISTS (SELECT 1 FROM (tb_pagineSito s"& _
			   "		 INNER JOIN tb_pages p ON s.id_pagineSito = p.id_paginaSito)"& _
		       "	 	 LEFT JOIN tb_layers l ON p.id_page = l.id_pag"& _
			   "		 WHERE s.id_pagineSito = tb_pagineSito.id_pagineSito"& _
			   "		 AND ("& conditionLingue &")"& _
			   "		 GROUP BY lingua"& _
			   "		 HAVING SUM("& SQL_If(conn, condition, "-id_lay", "id_lay") &") < 0"& _
			   "		 OR SUM("& SQL_If(conn, condition, "-1", "1") &") <> 0"& _
			   "		 OR (MAX(id_template) > 0 AND SUM(id_template) / MAX(id_template) / COUNT(*) <> 1))"
end if

sql = " SELECT tb_paginesito.*, " & _
	  SQL_If(conn, "id_paginesito=id_home_page", "1", "0") &" AS HOME, " & _
	  SQL_If(conn, "id_paginesito=sito_in_aggiornamento_pagina", "1", "0") &" AS isAgg, " & _
	  SQL_If(conn, "id_paginesito=sito_in_costruzione_pagina", "1", "0") &" AS isCostr, " & _
	  SQL_If(conn, "id_paginesito=errore_pagina", "1", "0") &" AS isErr, " & _
	  SQL_If(conn, "id_paginesito=id_home_page_riservata", "1", "0") &" AS isHomeRis, " & _
	  SQL_If(conn, "id_paginesito=id_login_page_riservata", "1", "0") &" AS isLoginRis " & _
	  " FROM tb_PagineSito INNER JOIN tb_webs ON tb_pagineSito.id_web=tb_webs.id_webs" & _
	  " WHERE tb_paginesito.id_web=" & Session("AZ_ID") & sql & _
	  " ORDER BY "& SQL_If(conn, "id_paginesito=id_home_page", "0", "1") &", tb_paginesito.nome_ps_IT, tb_paginesito.nome_ps_interno "
session("WEB_PAGINE_SQL") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>
<div id="content">
	<table cellspacing="0" cellpadding="0" border="0" style="100%">
		<tr>
	  		<td width="27%" valign="top">
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
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("pa_titolo") %>>TITOLO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_titolo" value="<%= TextEncode(session("pa_titolo")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("pa_archiviata") %>>ARCHIVIO</th></tr>
									<tr>
										<td class="content" style="width:50%;">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td><input type="checkbox" class="checkbox" name="search_archiviata" value="1" <%= chk(instr(1, session("pa_archiviata"), "1", vbTextCompare)>0) %>></td>
													<td style="padding-right:4px;"><img src="../grafica/archiviata.gif" border="0" alt="Pagina archiviata"></td>
													<td>archiviata</td>
												</tr>
											</table>
										</td>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_archiviata" value="0" <%= chk(instr(1, Session("pa_archiviata"), "0", vbTextCompare)>0) %>>
											attiva
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("pa_stato") %>>STATO</th></tr>
									<tr>
										<td class="content" style="width:50%;">
											<input type="checkbox" class="checkbox" name="search_stato" value="1" <%= chk(instr(1, session("pa_stato"), "1", vbTextCompare)>0) %>>
											pubblicata
										</td>
										<td class="content dapubblicare">
											<input type="checkbox" class="checkbox" name="search_stato" value="2" <%= chk(instr(1, Session("pa_stato"), "2", vbTextCompare)>0) %>>
											da pubblicare
										</td>
									</tr>
									<% if IsAreaRiservataActive(conn) then %>
										<tr><th colspan="2" <%= Search_Bg("pa_riservata") %>>PROTEZIONE</th></tr>
										<tr>
											<td class="content">
												<table cellpadding="0" cellspacing="0">
													<tr>
														<td><input type="checkbox" class="checkbox" name="search_riservata" value="1" <%= chk(instr(1, session("pa_riservata"), "1", vbTextCompare)>0) %>></td>
														<td style="padding-right:4px;"><img src="../grafica/padlock.gif" border="0" alt="Pagina appartenente all'area protetta"></td>
														<td>protetta</td>
													</tr>
												</table>
											</td>
											<td class="content">
												<input type="checkbox" class="checkbox" name="search_riservata" value="0" <%= chk(instr(1, Session("pa_riservata"), "0", vbTextCompare)>0) %>>
												pubblica
											</td>
										</tr>
									<% end if %>
									<tr><th colspan="2" <%= Search_Bg("pa_indicizzabile") %>>INDICIZZABILE</th></tr>
									<tr>
										<td class="content">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td><input type="checkbox" class="checkbox" name="search_indicizzabile" value="1" <%= chk(instr(1, session("pa_indicizzabile"), "1", vbTextCompare)>0) %>></td>
													<td style="padding-right:4px;"><img src="../grafica/indicizzazione.gif" border="0" alt="Pagina indicizzabile"></td>
													<td>si</td>
												</tr>
											</table>
										</td>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_indicizzabile" value="2" <%= chk(instr(1, Session("pa_indicizzabile"), "2", vbTextCompare)>0) %>>
											no
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("pa_idx") %>>PAGINE COLLEGATE A</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL index.WritePicker("", "", "ricerca", "search_idx", session("pa_idx"), Session("AZ_ID"), false, true, 32, false, false) %>
										</td>
									</tr>
									<% sql = QryElencoTemplate("", false)
									if cString(GetValueList(conn, NULL, sql))<>"" then %>
										<tr><th colspan="2" <%= Search_Bg("pa_template") %>>TEMPLATE</th></tr>
										<tr>
											<td class="content" colspan="2">
												<% 	
												CALL dropDown(conn, sql, "id_page", "NAME", "search_template", _
															  session("pa_template"), false, "style=""width: 100%;""", LINGUA_ITALIANO) %>
											</td>
										</tr>
									<% end if %>
									<% sql = "SELECT URL_rewriting_attivo FROM tb_webs WHERE id_webs=" & Session("AZ_ID")
									rsA.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext 
									if cBoolean(rsA("URL_rewriting_attivo"), false) then%>
										<tr><th colspan="2" <%= Search_Bg("pa_url_exact;pa_url_fulltext") %>>URL</th></tr>
										<tr>
											<td class="label_no_width">
												Chiave esatta
											</td>
											<td class="content">
												<input type="text" name="search_url_exact" value="<%= TextEncode(session("pa_url_exact")) %>" style="width:100%;">
											</td>
										</tr>
										<tr>
											<td class="label_no_width">
												Full text
											</td>
											<td class="content">
												<input type="text" name="search_url_fulltext" value="<%= TextEncode(session("pa_url_fulltext")) %>" style="width:100%;">
											</td>
										</tr>
									<% end if
									rsa.close %>
									<tr><th colspan="2" <%= Search_Bg("pa_numero_ps;pa_numero_pagina") %>>NUMERO</th></tr>
									<tr>
										<td colspan="2">
											<table cellspacing="1" cellpadding="0" width="100%">
												<tr>
													<td class="label_no_width">
														della pagina principale:
													</td>
													<td class="content" style="width:25%;">
														<input type="text" name="search_numero_ps" value="<%= TextEncode(session("pa_numero_ps")) %>" style="width:100%;">
													</td>
												</tr>
												<tr>
													<td class="label_no_width">
														della pagina di navigazione:
													</td>
													<td class="content">
														<input type="text" name="search_numero_pagina" value="<%= TextEncode(session("pa_numero_pagina")) %>" style="width:100%;">
													</td>
												</tr>
											</table>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("pa_testo;pa_img;pa_plugin") %>>CONTENUTO</th></tr>
									<tr><th colspan="2" class="L2" <%= Search_Bg("pa_testo") %>>TESTO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_testo" value="<%= TextEncode(session("pa_testo")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" class="L2" <%= Search_Bg("pa_img") %>>FILE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteFilePicker_Input(Application("AZ_ID"), "", "ricerca", "search_img", session("pa_img"), "width:88px", false) %>
										</td>
									</tr>
									<tr><th colspan="2" class="L2" <%= Search_Bg("pa_plugin") %>>PLUGIN</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% 	sql = "SELECT * FROM tb_objects ORDER BY name_objects"
												CALL dropDown(conn, sql, "id_objects", "name_objects", "search_plugin", _
															  session("pa_plugin"), false, "style=""width: 100%;""", LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("pa_data_from;pa_data_to") %>>DATA DI ULTIMA MODIFICA</th></tr>
									<tr><td class="label" colspan="2">a partire dal:</td></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_from", Session("pa_data_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><td class="label" colspan="2">fino al:</td></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_to", Session("pa_data_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr>
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<% if Session("WEB_ADMIN")<>"" OR session("WEB_POWER") <> "" then %>
							<tr><td style="font-size:10px;">&nbsp;</td></tr>
							<tr>
								<td>
									<table cellspacing="1" cellpadding="0" class="tabella_madre">
										<caption class="border">Strumenti delle pagine</caption>
										<tr>
											<td class="content_right">
												<a class="button_block" title="Esegue una analisi dello stato delle pagine del sito segnalando se le pagine sono state create correttamente, se sono correttamente pubblicate o se necessitano di manutenzione."
												   href="SitoAnalisiPubblicazione.asp?FROM=PAGINE">
													STATO DI PUBBLICAZIONE
												</a>
											</td>
										</tr>
                                        <!-- 
										<tr>
											<td class="content_right">
												<a class="button_block" title="Esegue una analisi dell'utilizzo dei template e ne permette la modifica"
												   href="SitoAnalisiTemplates.asp?FROM=PAGINE">
													UTILIZZO DEI TEMPLATE
												</a>
											</td>
										</tr>
                                         -->
									</table>
								</td>
							</tr>
						<% end if %>
						<tr><td>&nbsp;</td></tr>
					</table>
				</form>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
			
<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
					<caption class="border">
						Indice delle pagine - albero
					</caption>
					<tr>
						<td class="content">
							Visualizza l'indice delle pagine come albero
						</td>
						<td class="content_right">
							<a class="button" href="SitoPagineAlbero.asp" title="Apre la visualizzazione ad albero.">
								VISUALIZZA COME ALBERO
							</a>
						</td>
					</tr>
				</table>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Indice delle pagine - elenco
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovate n&ordm; <%= Pager.recordcount %> pagine del sito in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo
							'imposta stato di base di pubblicazione
							daPubblicare = false
							
							'impostazione sql di base per verifica uso di template semplificati per email
							sql = ""
							
							for each lingua in application("LINGUE")
								if Session("LINGUA_" & lingua) then
									'verifica stato della pubblicazione
									if not daPubblicare then
										if must_be_published(conn, rsA, rs("id_pagStage_"& lingua), rs("id_pagDyn_"& lingua)) then
											daPubblicare = true
										end if
									end if
									'aggiunge porzione di query per verifica uso di template semplificati per email
									sql = sql + IIF(sql="", "", " OR ") & _
												" tb_pages.id_page = " & cIntero(rs("id_pagDyn_" & lingua)) & " OR " & _ 
												" tb_pages.id_page = " & cIntero(rs("id_pagStage_" & lingua))
								end if
							next 
							riepilogo = " title=""" & GetPageNumbers(rs) & """ "
							
							'completa query per verifica uso template semplificati per email
							if sql <> "" then
								sql = "SELECT COUNT(*) FROM tb_pages LEFT JOIN tb_pages tb_templates " + _
					 	  			  " ON tb_pages.id_template=tb_templates.id_page " + _
						  			  " WHERE " + SQL_IsTrue(conn, "tb_templates.semplificata") + _
									  " AND (" + sql + ")"
								if cIntero(GetValueList(conn, rsA, sql))>0 then
									templateSemplificato = true
								else
									templateSemplificato = false
								end if
							else
								templateSemplificato = false
							end if
							%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td title="stato: <%= IIF(daPubblicare, "da pubblicare", "pubblicata")%>" class=" header" colspan="4">
												<% if rs("archiviata") OR rs("riservata") OR templateSemplificato then %>
													<table border="0" cellspacing="0" cellpadding="0" align="left">
														<tr>
															<td style="padding-top:1px; padding-right:4px; font-size:1px;">
																<% if rs("archiviata") then %>
																	<img src="../grafica/archiviata.gif" border="0" alt="Pagina archiviata">
																	&nbsp;
																<% end if
																if rs("riservata") then %>
																	<img src="../grafica/padlock.gif" border="0" alt="Pagina appartenente all'area protetta">
																	&nbsp;
																<% end if 
																if templateSemplificato then%>
																	<img src="../grafica/notReadKnow.gif" border="0" alt="Pagina che utilizza almeno un template per email con visualizzazione semplificata.">
																	&nbsp;
																<% end if %>
															</td>
														</tr>
													</table>
												<% end if %>
												<%= PaginaSitoNome(rs, "") %>
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px; width:50%; text-align:right;">
															<% if index.content.ChkPrmF("tb_pagineSito", rs("id_pagineSito")) then %>
															    <% 	CALL index.WriteButton("tb_pagineSito", rs("id_pagineSito"), POS_ELENCO) %>
															    <a class="button" href="SitoPagineMod.asp?ID=<%= rs("id_pagineSito") %>">
																    MODIFICA
															    </a>
                                                                &nbsp;
                                                                <% 'verifica se la pagina e' usata da contenuti come supporto di pubblicazione.
																CanDelete = ""
																if rs("home") then
																	CanDelete = CanDelete + " &egrave; utilizzata come home page del sito." + vbCrLf
																end if
																if rs("isHomeRis") then
																	CanDelete = CanDelete + " &egrave; utilizzata come home dell'area riservata." + vbCrLf
																end if
																if rs("isLoginRis") then
																	CanDelete = CanDelete + " &egrave; utilizzata come pagina di login dell'area riservata." + vbCrLf
																end if
																if rs("isErr") then
																	CanDelete = CanDelete + " &egrave; utilizzata come pagina di errore." + vbCrLf
																end if
																if rs("isCostr") then
																	CanDelete = CanDelete + " &egrave; utilizzata come home page del sito in fase di costruzione." + vbCrLf
																end if
																if rs("isAgg") then
																	CanDelete = CanDelete + " &egrave; utilizzata come home page del sito in fase di aggiornamento." + vbCrLf
																end if
                                                                sql = " SELECT COUNT(*) FROM v_indice_it WHERE NOT (tab_name LIKE 'tb_paginesito') AND " + _
                                                                      " ( co_link_pagina_id=" & rs("id_pagineSito") & " OR idx_link_pagina_id = " & rs("id_pagineSito") & ")"
                                                                if cIntero(GetValueList(conn, rsA, sql))>0 then
																	CanDelete = CanDelete + "sono presenti contenuti dell'indice che vengono pubblicati per mezzo di essa." + vbCrLf
																end if
																if CanDelete = "" then%>
                                                                    <a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('PAGINE','<%= rs("id_pagineSito") %>');">
                                                                        CANCELLA
                                                                    </a>
                                                                <% else %>
                                                                    <a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la pagina:<%= vbCrLF %><%= CanDelete %>" <%= ACTIVE_STATUS %>>
                                                                        CANCELLA
                                                                    </a>
															    <% end if
                                                            else %>
															    <a class="button" title="Visualizza il layout della pagina" href="javascript:void(0)" onclick="OpenAutoPositionedWindow('SitoPagineView.asp?ID=<%= rs("id_pagineSito") %>', 'vedi', 400, 300)" <%= ACTIVE_STATUS %>>
																    VEDI
															    </a>
															<% end if %>
														</td>
													</tr>
												</table>
                                                
											</td>
										</tr>
										<tr>
											<td class="label" style="width:28%;">data ultima modifica</td>
											<td class="content"<%= riepilogo %>><%= rs("ps_modData") %></td>
											<td class="label" style="width:8%;">numero</td>
											<td class="content_right" style="width:7%;"<%= riepilogo %>><%= rs("id_pagineSito") %></td>
										</tr>
										<tr>
											<td class="label">stato della pagina</td>
											<td colspan="3">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<% 	if rs("home") then %>
													<tr>
														<td class="content homepage">Home page del sito</td>
													</tr>
													<% 	end if %>
													<% 	if rs("isHomeRis") then %>
													<tr>
														<td class="content homeareariservata">Home page dell'area riservata</td>
													</tr>
													<% 	end if %>
													<% 	if rs("isLoginRis") then %>
													<tr>
														<td class="content loginareariservata">Pagina di login dell'area riservata</td>
													</tr>
													<% 	end if %>
													<% 	if rs("isErr") then %>
													<tr>
														<td class="content paginaerrore">Pagina di errore</td>
													</tr>
													<% 	end if %>
													<% 	if rs("isCostr") then %>
													<tr>
														<td class="content incostruzione">Home page del sito in fase di costruzione</td>
													</tr>
													<% 	end if %>
													<% 	if rs("isAgg") then %>
													<tr>
														<td class="content inaggiornamento">Home page del sito in fase di aggiornamento</td>
													</tr>
													<% 	end if %>
													<tr>
														<td class="content<%= IIF(DaPubblicare, " dapubblicare", "") %>" <%= riepilogo %>>
															<% if daPubblicare then %>
																da pubblicare
															<% else %>
																pubblicata
															<% end if %>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<% 	'ELENCO VOCI COLLEGATE
											sql = " SELECT * FROM v_indice_it "& _
												  " WHERE co_F_key_id = "& rs("id_pagineSito") & _
												  " AND co_F_table_id = "& index.GetTable("tb_pagineSito") & _
												  " ORDER BY idx_ordine_assoluto, idx_ordine, idx_tipologie_padre_lista "
											rsA.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
											if not rsA.eof then 
                                                while not rsA.eof %>
                                                    <tr>
                                                        <% if rsA.absoluteposition = 1 then%>
                                                            <td class="label" rowspan="<%= rsA.recordcount %>">pubblicazioni sull'indice:</td>
                                                        <% end if %>
                                                        <td class="content" colspan="3">
															<% CALL index.WriteNodeLink(rsa, "", LINGUA_ITALIANO) %>
                                                        </td>
                                                    </tr>
                                                    <% rsa.movenext
                                                wend
                                            end if
											rsA.close %>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td class="footer" style="border-top:0px; text-align:left;" colspan="2">
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
set rsA = nothing
set conn = nothing
%>
