<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->

<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione modelli - elenco"
dicitura.puls_new = "NUOVO MODELLO;"
dicitura.link_new = "ArticoliNew.asp?TYPE=AS;"
dicitura.scrivi_con_sottosez() 

dim conn, rs, rsc, rsv, sql, Pager, colore, txt, rs_spe, sql_filtri
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
	sql = sql & " art_tipologia_id IN (0, " & categorie.DiscendentiID(Session("art_categoria")) & " )"
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


'ricerca per tipo spedizione	  
if Session("art_spedizione")<>"" then
	if sql <>"" then sql = sql & " AND "
		sql = sql & " art_spedizione_id=" & Session("art_spedizione")
end if


'ricerca full-text
if Session("art_full_text")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & SQL_FullTextSearch(Session("art_full_text"), FieldLanguageList("art_nome_;art_descr_"))
end if

sql_filtri = sql

sql = " WHERE art_tipologia_id IN ("&cat_modelli.DiscendentiID(0)&") "

if sql_filtri <> "" then sql = sql & " AND " & sql_filtri
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
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<form action="" method="post" id="ricerca" name="ricerca">
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
								<tr><th colspan="2" <%= Search_Bg("art_stato_catalogo") %>>STATO MODELLO A CATALOGO</th></tr>
								<tr>
									<td class="content" style="width:45%;">
										<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="1" <%= chk(instr(1, session("art_stato_catalogo"), "1", vbTextCompare)>0) %>>
										<b>visibile</b>
									</td>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="0" <%= chk(instr(1, Session("art_stato_catalogo"), "0", vbTextCompare)>0) %>>
										non visibile
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("art_categoria") %>>CATEGORIA</th></tr>
								<tr>
									<td class="content" colspan="2">
										<% CALL cat_modelli.WritePicker("ricerca", "search_categoria", session("art_categoria"), false, true, 32) %>
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
								if rs_spe.RecordCount > 1 then %>
										<tr><th colspan="2" <%= Search_Bg("art_spedizione") %>>TIPO SPEDIZIONE</th></tr>
										<tr>
											<td class="content" colspan="2">
												<%
												CALL dropDown(conn, sql, "spa_id", "spa_nome_it", "search_spedizione", session("art_spedizione"), false, " style=""width:100%;""", LINGUA_ITALIANO) 
												%>
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

					</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco modelli
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> modelli in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="<%= IIF(rs("art_disabilitato"), "header_disabled", "header") %>" colspan="7">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px; text-align:right;">
															<% 'CALL index.WriteButton("gtb_Articoli", rs("art_id"), POS_ELENCO) 
															%> &nbsp;
														
															<a class="button" href="ArticoliMod.asp?ID=<%= rs("art_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<% 	sql = " SELECT TOP 1 sc_id FROM sgtb_schede WHERE sc_modello_id IN (" & _
																	  " SELECT  rel_id FROM grel_art_valori WHERE rel_art_id = " & rs("art_id") & ")"
																if GetValueList(conn, rsc, sql) > 0 then %>
																<a class="button_disabled" title="Impossibile cancellare l'articolo perch&egrave; &egrave; presente almeno una scheda collegata">
																	CANCELLA
																</a>
															<% else %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ARTICOLI','<%= rs("art_id") %>');">
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												<%= rs("art_nome_it") %>
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

										<% CALL ListaCollegamentiArticolo(conn, rsc, rs("art_id"), rs("art_ha_accessori"), false) %>
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
