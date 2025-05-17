<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
'Gestione filtro per listino sui clienti
if cIntero(request.querystring("FILTRO_CLIENTI"))>0 then
	Session("")
end if

dim conn, rs, rsr, sql, pager
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione listini - elenco"

sql = "SELECT COUNT(*) FROM gtb_listini WHERE listino_base_attuale=1"
dicitura.puls_new = "nuovo:;"
dicitura.link_new = ";"
if cInteger(GetValueList(conn, NULL, sql))>0 then
	dicitura.puls_new = dicitura.puls_new + "LISTINO;LISTINO OFFERTE SPECIALI;"
	dicitura.link_new = dicitura.link_new + "ListiniNew.asp;ListiniNew.asp?TIPO=OFFERTE SPECIALI;"
end if
dicitura.puls_new = dicitura.puls_new + "LISTINO BASE"
dicitura.link_new = dicitura.link_new + "ListiniNew.asp?TIPO=BASE"

dicitura.scrivi_con_sottosez()  

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("ls_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("ls_")
	end if
end if

sql = ""
'filtra per codice listino e nome
if Session("ls_codice")<>"" then
	if sql<>"" then sql = Sql & " AND "
	sql = sql & "(" & sql_FullTextSearch(Session("ls_codice"), "listino_codice") & " OR "& sql_FullTextSearch(Session("ls_codice"), "listino_nome_it") &") " 
end if

if session("ls_stato")<>"" then
	if sql <> "" then sql = sql & " AND "
	sql = sql + " ("
	if instr(1, session("ls_stato"), "B", vbTextCompare) then
		sql = sql + " listino_Base_attuale=1 "
	end if
	if instr(1, session("ls_stato"), "O", vbTextCompare) then
		if right(sql, 1) <> "(" then sql = sql + " OR "
		sql = sql + " (listino_offerte=1 AND (" & SQL_Now(conn) & " BETWEEN listino_DataCreazione AND ISNULL(listino_DataScadenza, GetDate()) + 1))"
	end if
	if instr(1, session("ls_stato"), "P", vbTextCompare) then
		if right(sql, 1) <> "(" then sql = sql + " OR "
		sql = sql + " (listino_b2c=1 AND (" & SQL_Now(conn) & " BETWEEN listino_DataCreazione AND ISNULL(listino_DataScadenza, GetDate()) + 1))"
	end if
	sql = sql + ")"
end if

if Session("ls_tipo")<>"" then
	if sql <> "" then sql = sql & " AND "
	sql = sql + " ("
	if instr(1, session("ls_tipo"), "B", vbTextCompare) then
		sql = sql + " listino_Base=1 "
	end if
	if instr(1, session("ls_tipo"), "O", vbTextCompare) then
		if right(sql, 1) <> "(" then sql = sql + " OR "
		sql = sql + " listino_offerte=1 "
	end if
	if instr(1, session("ls_tipo"), "CP", vbTextCompare) then
		if right(sql, 1) <> "(" then sql = sql + " OR "
		sql = sql + " (IsNull(listino_offerte, 0)=0 AND IsNull(listino_Base, 0)=0 AND IsNull(listino_ancestor_id, 0)=0) "
	end if
	if instr(1, session("ls_tipo"), "CD", vbTextCompare) then
		if right(sql, 1) <> "(" then sql = sql + " OR "
		sql = sql + " (IsNull(listino_offerte, 0)=0 AND IsNull(listino_Base, 0)=0 AND IsNull(listino_ancestor_id, 0)>0) "
	end if
	if instr(1, session("ls_tipo"), "I", vbTextCompare) then
		if right(sql, 1) <> "(" then sql = sql + " OR "
		sql = sql + " (IsNull(listino_importato, 0)=1) "
	end if
	sql = sql + ")"
end if

if cInteger(Session("ls_ancestor"))>0 then
	if sql <> "" then sql = sql & " AND "
	sql = sql + " listino_ancestor_id=" & Session("ls_ancestor")
end if

if Session("ls_cliente")<>"" then
	if sql <> "" then sql = sql & " AND "
	sql = sql + " listino_id IN (SELECT riv_listino_id FROM gv_rivenditori WHERE " + _
                                 SQL_FullTextSearch_Contatto_Nominativo(conn, Session("ls_cliente")) + _
				")"
end if

sql = " SELECT * " + _
	  " FROM gtb_listini " + IIF(sql<>"", " WHERE " + sql, "") + " ORDER BY listino_base DESC, listino_offerte DESC, listino_codice"
session("B2B_LISTINI_SQL") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)%>
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
										<td class="footer">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("ls_tipo") %>>TIPO</th></tr>
									<tr>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_tipo" value="B" <%= chk(instr(1, Session("ls_tipo"), "B", vbTextCompare)>0) %>>
											listino base
										</td>
									</tr>
									<tr>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_tipo" value="O" <%= chk(instr(1, Session("ls_tipo"), "O", vbTextCompare)>0) %>>
											listino offerte speciali
										</td>
									</tr>
									<tr>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_tipo" value="CP" <%= chk(instr(1, Session("ls_tipo"), "CP", vbTextCompare)>0) %>>
											listino clienti principale
										</td>
									</tr>
									<tr>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_tipo" value="CD" <%= chk(instr(1, Session("ls_tipo"), "CD", vbTextCompare)>0) %>>
											listino clienti derivato
										</td>
									</tr>
									<tr>
										<td class="content bundle">
											<input type="checkbox" class="checkbox" name="search_tipo" value="I" <%= chk(instr(1, Session("ls_tipo"), "I", vbTextCompare)>0) %>>
											listino importato
										</td>
									</tr>
									<tr><th <%= Search_Bg("ls_codice") %>>NOME o CODICE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_codice" value="<%= TextEncode(session("ls_codice")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("ls_stato") %>>STATO</th></tr>
									<tr>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_stato" value="B" <%= chk(instr(1, Session("ls_stato"), "B", vbTextCompare)>0) %>>
											listino base in vigore
										</td>
									</tr>
									<tr>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_stato" value="O" <%= chk(instr(1, Session("ls_stato"), "O", vbTextCompare)>0) %>>
											listino offerte speciali in vigore
											<span class="Icona Offerte" title="listino offerte speciali in vigore">&nbsp;</span>
										</td>
									</tr>
									<tr>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_stato" value="P" <%= chk(instr(1, Session("ls_stato"), "P", vbTextCompare)>0) %>>
											listino visibile al pubblico
										</td>
									</tr>
									<tr><th <%= Search_Bg("ls_ancestor") %>>DERIVAZIONE DEL LISTINO</th></tr>
									<tr>
										<td class="content">
											<% sql = " SELECT listino_id, listino_codice FROM gtb_listini " + _
													 " WHERE listino_offerte=0 AND listino_base=0 AND IsNull(listino_ancestor_id,0)=0 ORDER BY listino_codice"
											CALL dropDown(conn, sql, "listino_id", "listino_codice", "search_ancestor", session("ls_ancestor"), false, " style=""width:100%;""", LINGUA_ITALIANO)%>
										</td>
									</tr>
									<tr><th <%= Search_Bg("ls_cliente") %>>NOMINATIVO CLIENTE ASSOCIATO</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_cliente" value="<%= TextEncode(session("ls_cliente")) %>" style="width:100%;">
										</td>
									</tr>
									<tr>
										<td class="footer">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTI" style="width: 49%;">
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
						Elenco listini
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> listini in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header <%=IIF(cBoolean(rs("listino_importato"), false),"bundle","")%>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px; text-align:right;">
															<% if cBoolean(rs("listino_importato"), false) AND not IsAdminCurrent(conn) then %>
																<a class="button" href="ListiniMod.asp?ID=<%= rs("listino_id") %>" title="Visualizza i dati del listino importato" <%= ACTIVE_STATUS %>>
																	VISUALIZZA
																</a>
															<% else %>
																<a class="button" href="ListiniMod.asp?ID=<%= rs("listino_id") %>" title="Modifica i dati del listino" <%= ACTIVE_STATUS %>>
																	MODIFICA
																</a>
																&nbsp;
																<% if rs("listino_base_attuale") then %>
																	<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il listino base attualmente attivo." <%= ACTIVE_STATUS %>>
																		CANCELLA
																	</a>
																<% else
																	if rs("listino_with_child") then %>
																		<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il listino: &egrave; associato ad almeno un cliente." <%= ACTIVE_STATUS %>>
																			CANCELLA
																		</a>
																	<% else 
																		sql = "SELECT COUNT(*) FROM gtb_listini WHERE listino_ancestor_id=" & rs("listino_id")
																		if cInteger(GetValueList(conn, rsr, sql)) > 0 then %>
																			<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il listino principale: da esso sono derivati almenno un listino." <%= ACTIVE_STATUS %>>
																				CANCELLA
																			</a>
																		<% else 
																			sql = "SELECT COUNT(*) FROM gtb_rivenditori WHERE riv_listino_id=" & rs("listino_id")
																			if cInteger(GetValueList(conn, rsr, sql)) > 0 then%>
																				<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il listino: associato ad almeno un cliente." <%= ACTIVE_STATUS %>>
																					CANCELLA
																				</a>
																			<% else %>
																				<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('LISTINO','<%= rs("listino_id") %>');" >
																					CANCELLA
																				</a>
																			<% end if
																		end if
																	end if 
																end if %>
															<% end if %>
														</td>
													</tr>
												</table>
												<%= rs("listino_nome_it")%> (<%= rs("listino_codice") %>)
											</td>
										</tr>
										<% if not rs("listino_offerte") then %>
											<tr>
												<td class="label" style="width:23%;">clienti associati:</td>
												<td class="content_right" colspan="3" style="padding-right:0px; font-size:1px;">
													<% CALL WriteCampoCerca("Clienti.asp", "listino", rs("listino_id"), "CLIENTI ASSOCIATI", "button_L2") %>
												</td>
											</tr>
										<% end if %>
										<tr>
											<td class="label" style="width:23%;">gestione dei prezzi:</td>
											<td class="content_right" colspan="3" style="padding-right:0px; font-size:1px;">
												<a class="button_L2" href="ListiniPrezzi_Gruppi.asp?ID=<%= rs("listino_id") %>" title="Permette la ricerca di gruppi di articoli e quindi la modifica diretta dei loro prezzi." <%= ACTIVE_STATUS %>>
													PER GRUPPI
												</a>
												<% if NOT (cBoolean(rs("listino_importato"), false) AND not IsAdminCurrent(conn)) then %>
													&nbsp;
													<a class="button_L2" href="ListiniPrezzi_RigaPerRiga.asp?ID=<%= rs("listino_id") %>" title="Permette la ricerca di gruppi di articoli e la modifica rapida riga per riga dei loro prezzi" <%= ACTIVE_STATUS %>>
														RIGA PER RIGA
													</a>
													&nbsp;
													<a class="button_L2" href="ListiniPrezzi_Avanzata.asp?ID=<%= rs("listino_id") %>" title="Permette la gestione di un massimo di 600 articoli permettendone la modifica e a gruppi o riga per riga." <%= ACTIVE_STATUS %>>
														AVANZATA
													</a>
												<% end if %>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:20%;">tipo:</td>
											<td class="content" style="width:39%;">
												<% if rs("listino_base") then 
													if rs("listino_base_attuale") then%>
														<strong title="listino base attualmente in vigore">listino base</strong>
													<% else %>
														listino base
													<% end if
												elseif rs("listino_offerte") then %>
													listino offerte speciali
												<% elseif rs("listino_b2c") then %>
													listino al pubblico
												<% else
													if cInteger(rs("listino_ancestor_id"))>0 then%>
														listino clienti derivato
													<% else %>
														listino clienti principale
													<% end if
												end if %>
											</td>
											<td class="label">stato:</td>
											<td class="content">
												<% if rs("listino_base_attuale") OR _
													  ((rs("listino_offerte") OR rs("listino_b2c")) AND _
													   (rs("listino_DataCreazione") <= Date AND _
													    ((rs("listino_DataScadenza") >= Date) OR isNull(rs("listino_DataScadenza"))) )) OR _
														rs("listino_offerte") AND isNull(rs("listino_DataCreazione")) AND isNull(rs("listino_DataScadenza")) then%>
													<strong>in vigore</strong>
													<% if rs("listino_offerte") then %>
														<span class="Icona Offerte" title="listino offerte speciali in vigore">&nbsp;</span>
													<% end if %>
												 <% else %>
													&nbsp;
												<% end if %>
											</td>
										</tr>
										<% if rs("listino_offerte") then %>
											<tr>
												<td class="label">validit&agrave; offerte:</td>
												<td class="content">
													<% if IsNull(rs("listino_DataCreazione")) AND IsNull(rs("listino_DataScadenza")) then %>
														<strong>gestito su ogni articolo</strong>
													<% elseif isDate(rs("listino_DataCreazione")) then%>
														dal <%= DateITA(rs("listino_DataCreazione")) %>
													<% end if
													if isDate(rs("listino_DataScadenza")) then %>
														fino al <%= DateITA(rs("listino_DataScadenza")) %>
													<% end if %>
												</td>
												<td class="label">articoli in offerta:</td>
												<td class="content">
													<% sql = " SELECT COUNT(*) FROM gtb_prezzi " + _
															 " WHERE prz_listino_id=" & rs("listino_id") & _
															 " AND IsNull(prz_visibile,0) = 1 AND (IsNull(prz_var_euro, 0)<>0 OR IsNull(prz_var_sconto, 0)<>0) " 
													dim n_art_off 
													n_art_off = GetValueList(conn, rsr, sql) %>
													<span style="float:left;"><%= n_art_off %></span>
													<% if n_art_off > 0 then %>
														<a class="button_L2" style="display:block; width:70px; text-align:center; float:right;" href="ListiniPrezzi_RigaPerRiga.asp?ID=<%= rs("listino_id") %>&list_offerte=1" title="Visualizza gli articoli in offerta" <%= ACTIVE_STATUS %>>
															VISUALIZZA ARTICOLI
														</a>
													<% end if %>
												</td>
											</tr>
										<% elseif rs("listino_b2c") then %>
											<tr>
												<td class="label">pubblicazione:</td>
												<td class="content" colspan="3">visible al pubblico</td>
											</tr>
											<tr>
												<td class="label">validit&agrave; pubblicazione:</td>
												<td class="content" colspan="3">
													<% if isDate(rs("listino_DataCreazione")) then%>
														dal <%= DateITA(rs("listino_DataCreazione")) %>
													<% end if
													if isDate(rs("listino_DataScadenza")) then %>
														fino al <%= DateITA(rs("listino_DataScadenza")) %>
													<% end if%>
												</td>
											</tr>
										<% end if 
										if cInteger(rs("listino_ancestor_id"))>0 then
											sql = "SELECT listino_codice FROM gtb_listini WHERE listino_id=" & rs("listino_ancestor_id")%>
											<tr>
												<td class="label">derivato dal listino:</td>
												<td class="content" colspan="3"><%= GetValueList(conn, rsr, sql) %></td>
											</tr>
										<% end if %>
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
set rsr = nothing
set conn = nothing%>