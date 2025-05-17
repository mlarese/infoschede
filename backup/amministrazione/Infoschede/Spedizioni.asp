<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" --> 
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione spedizioni - elenco"
dicitura.puls_new = "NUOVO DDT;NUOVA LETTERA D'ACCOMPAGNAMENTO"
dicitura.link_new = "SpedizioniNew.asp?CAT_ID="&DDT_CAT_ID&";SpedizioniNew.asp?CAT_ID="&LETTERE_CAT_ID&";"
dicitura.scrivi_con_sottosez()  

dim conn, rs, rsc, sql, pager, i, max_ddt, max_lettere

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("spe_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("spe_")
	end if
end if


'filtra per tipo
if Session("spe_tipo")<>"" then
	sql = sql & " AND ddt_categoria_id IN (" & Session("spe_tipo") & ") "
end if

'filtra numero ritiro
if session("spe_numero") <> "" then
	sql = sql & " AND ddt_numero IN (" & Session("spe_numero") & ") "
end if

'filtra numero scheda
if session("spe_numero_scheda") <> "" then
	sql = sql & " AND ddt_id IN (SELECT sc_rif_DDT_di_resa_id FROM sgtb_schede WHERE sc_numero = " & Session("spe_numero_scheda") & ") "
end if

'filtra data scheda
if isDate(Session("spe_data_scheda_from")) then
	sql = sql & " AND ddt_id IN (SELECT sc_rif_DDT_di_resa_id FROM sgtb_schede WHERE " & _
							SQL_CompareDateTime(conn, "sc_data_ricevimento", adCompareGreaterThan, Session("spe_data_scheda_from")) & ") "
end if
if isDate(Session("spe_data_scheda_to")) then
	sql = sql & " AND ddt_id IN (SELECT sc_rif_DDT_di_resa_id FROM sgtb_schede WHERE " & _
							SQL_CompareDateTime(conn, "sc_data_ricevimento", adCompareLessThan, Session("spe_data_scheda_to")) & ") "
end if

'filtra per causale
if session("spe_causale") <> "" then
	sql = sql & " AND ddt_causale_id = " & session("spe_causale")
end if

'filtra per nome cliente
if Session("spe_denominazione")<>"" then
	sql = sql & " AND ddt_cliente_id IN (SELECT riv_id FROM gv_rivenditori WHERE riv_profilo_id NOT IN ("&TRASPORTATORI&","&COSTRUTTORI&") AND " & _
							SQL_FullTextSearch_Contatto_Nominativo(conn, Session("spe_denominazione")) & ") "
end if

'filtra per nome trasportatore
if Session("spe_trasportatore")<>"" then
	sql = sql & " AND ddt_trasportatore_id IN (SELECT riv_id FROM gv_rivenditori WHERE riv_profilo_id IN ("&TRASPORTATORI&") AND " & _
							SQL_FullTextSearch_Contatto_Nominativo(conn, Session("spe_trasportatore")) & ") "
end if


'filtra per indirizzo di destinazione
if Session("spe_indirizzo")<>"" then
	sql = sql & " AND ddt_destinazione_id IN (SELECT IDElencoIndirizzi FROM tb_indirizzario WHERE " + _
                                SQL_FullTextSearch_Contatto_Indirizzo(conn, Session("spe_indirizzo")) & ") "
end if

'filtra per citta
if Session("spe_citta")<>"" then
	sql = sql & " AND ddt_destinazione_id IN (SELECT IDElencoIndirizzi FROM tb_indirizzario WHERE " + _
								sql_FullTextSearch(Session("spe_citta"), "CittaElencoIndirizzi") & ") "
end if





sql = " SELECT * FROM (sgtb_ddt INNER JOIN sgtb_ddt_categorie ON sgtb_ddt.ddt_categoria_id = sgtb_ddt_categorie.cat_id) " + _
	  " INNER JOIN sgtb_ddt_causali ON sgtb_ddt.ddt_causale_id=sgtb_ddt_causali.cau_id " + _
	  " WHERE (1=1) " & sql
sql = sql & " ORDER BY ddt_data DESC, ddt_id DESC"
session("INFOSCHEDE_SPEDIZIONI_SQL") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 6)
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
									<tr><th colspan="2" <%= Search_Bg("spe_tipo") %>>TIPO</td></tr>
									<tr>
										<td class="content OrdEvaso" colspan="2">
											<input type="checkbox" class="checkbox" name="search_tipo" value="<%=DDT_CAT_ID%>" <%= chk(cIntero(Session("spe_tipo"))=DDT_CAT_ID) %>>
											DDT
										</td>
									</tr>
									<tr>
										<td class="content OrdConfermato" colspan="2">
											<input type="checkbox" class="checkbox" name="search_tipo" value="<%=LETTERE_CAT_ID%>" <%= chk(cIntero(Session("spe_tipo"))=LETTERE_CAT_ID) %>>
											Lettere d'accompagnamento
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("spe_numero") %>>NUMERO DDT/LETTERA</th></tr>
									<tr>
										<td class="content" colspan="2">
										<input type="text" name="search_numero" value="<%= TextEncode(session("spe_numero")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("spe_numero_scheda") %>>NUMERO SCHEDA</th></tr>
									<tr>
										<td class="content" colspan="2">
										<input type="text" name="search_numero_scheda" value="<%= TextEncode(session("spe_numero_scheda")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("spe_data_scheda_from;spe_data_scheda_to")%>>DATA SCHEDA</td></tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("spe_data_scheda_from")%>>a partire dal:</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_scheda_from", Session("spe_data_scheda_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("spe_data_scheda_to")%>>fino al:</th></tr>
									<tr>
										<td class="content" colspan="2">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_scheda_to", Session("spe_data_scheda_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("spe_causale") %>>CAUSALE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% sql = "SELECT cau_id, cau_titolo_it FROM sgtb_ddt_causali ORDER BY cau_titolo_it"
											CALL dropDown(conn, sql, "cau_id", "cau_titolo_it", "search_causale", session("spe_causale"), false, " style=""width:100%;""", LINGUA_ITALIANO) %> 
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("spe_denominazione") %>>CLIENTE</th></tr>
									<tr>
										<td class="content" colspan="2">
										<input type="text" name="search_denominazione" value="<%= TextEncode(session("spe_denominazione")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("spe_trasportatore") %>>TRASPORTATORE</th></tr>
									<tr>
										<td class="content" colspan="2">
										<input type="text" name="search_trasportatore" value="<%= TextEncode(session("spe_trasportatore")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("spe_indirizzo;spe_citta") %>>DESTINAZIONE</th></tr>
									<tr>
										<td class="label">
											indirizzo:
										</td>
										<td class="content">
											<input type="text" name="search_indirizzo" value="<%= TextEncode(session("spe_indirizzo")) %>" style="width:100%;">
										</td>
									</tr>
									<tr>
										<td class="label">
											citt&agrave;:
										</td>
										<td class="content">
											<input type="text" name="search_citta" value="<%= TextEncode(session("spe_citta")) %>" style="width:100%;">
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
						Elenco spedizioni
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
										<% if cIntero(rs("ddt_categoria_id"))=DDT_CAT_ID then %>
											<td class="header OrdEvaso" colspan="4">
										<% elseif cIntero(rs("ddt_categoria_id"))=LETTERE_CAT_ID then %>
											<td class="header OrdConfermato" colspan="4">
										<% end if %>	
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
													<%	sql = " SELECT MAX(ddt_numero) FROM sgtb_ddt WHERE ddt_categoria_id = " & DDT_CAT_ID & _
															  " AND " & SQL_BetweenDate(conn, "ddt_data", "01/01/"&Year(Now()), "31/12/"&Year(Now()))
														max_ddt = cIntero(GetValueList(conn, NULL, sql))
														sql = " SELECT MAX(ddt_numero) FROM sgtb_ddt WHERE ddt_categoria_id = " & LETTERE_CAT_ID & _
															  " AND " & SQL_BetweenDate(conn, "ddt_data", "01/01/"&Year(Now()), "31/12/"&Year(Now()))
														max_lettere = cIntero(GetValueList(conn, NULL, sql))
													%>
														<td style="font-size: 1px;">
															<a class="button" href="SpedizioniMod.asp?ID=<%= rs("ddt_id") %>">
																MODIFICA
															</a>
															&nbsp;
														<%	sql = "SELECT COUNT(*) FROM sgtb_schede WHERE sc_rif_DDT_di_resa_id="& rs("ddt_id")%>
														<% 	if cIntero(GetValueList(conn,NULL,sql))=0 OR _
																(rs("ddt_numero")=max_ddt AND rs("ddt_categoria_id")=DDT_CAT_ID) OR _
																(rs("ddt_numero")=max_lettere AND rs("ddt_categoria_id")=LETTERE_CAT_ID) then %>
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('SPEDIZIONI','<%= rs("ddt_id") %>');">
																CANCELLA
															</a>
														<% 	else %>
															<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il record">
																CANCELLA
															</a>
														<% 	end if %>
														</td>
													</tr>
												</table>
												<%= rs("cat_nome_it") %>
												 n.
												<%= rs("ddt_numero") %>
												 del 
												<%= DateIta(rs("ddt_data")) %>
											</td>
										</tr>
										
										<% sql = " SELECT * FROM (sgtb_schede INNER JOIN gv_articoli ON sgtb_schede.sc_modello_id = gv_articoli.rel_id) " & _
												 " INNER JOIN sgtb_ddt ON sgtb_schede.sc_rif_DDT_di_resa_id = sgtb_ddt.ddt_id " & _
												 " WHERE ddt_id=" & rs("ddt_id") 
										rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
										<% if not rsc.eof then %>
											<tr>
												<td class="label">schede associate:</td>
												<td colspan="3">
													<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
														<tr>
															<th class="l2_center" style="width:28%;">num. scheda e data</th>
															<th class="l2_center" style="width:13%;">costo</th>
															<th class="L2">modello</th>
														</tr>
														<% while not rsc.eof%>
															<tr>
																<td class="content_center"><% CALL SchedaLink(rsc("sc_id"), rsc("sc_numero") & " del " & rsc("sc_data_ricevimento"))%></td>
																<td class="content_center">
																	<%= FormatPrice(cReal(rsc("sc_costo_riconsegna")), 2, false) %> &euro;
																</td>																<td class="content">
																	<% CALL ArticoloLink(rsc("art_id"), rsc("art_nome_it"), rsc("art_cod_int")) %>
																	<% if rsc("art_varianti") then %>
																		<%= ListValoriVarianti(conn, rsi, rsc("rel_id")) %>
																	<% else %>
																		&nbsp;
																	<% end if %>
																</td>
															</tr>
															<% rsc.moveNext %>
														<% wend %>
													</table>
												</td>
											</tr>
										<% end if %>
										<% rsc.close %>
										<tr>
											<td class="label">causale:</td>
											<td class="content" colspan="3"><%= rs("cau_titolo_it") %></td>
										</tr>
										<tr>
											<td class="label" style="width:21%;">cliente:</td>
											<td class="content" colspan="3">
												<% sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & rs("ddt_cliente_id")
												rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
												<% 'CALL ClienteLink(rsc("IDElencoIndirizzi") , ContactFullName(rsc)) 
												%>
												<%= ContactFullName(rsc)%>
												<% rsc.close %>
											</td>
										</tr>
										
										<tr>
											<td class="label" style="width:21%;">trasportatore:</td>
											<% if cIntero(rs("ddt_trasportatore_id"))>0 then %>
												<td class="content" colspan="3">
													<% sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & rs("ddt_trasportatore_id")
													rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
														<%= ContactFullName(rsc)%>
													<% rsc.close %>
												</td>
											<% else %>
												<td class="note" colspan="3">Trasportatore non inserito</td>
											<% end if %>
										</tr>
										
										<tr>
											<td class="label">destinazione:</td>
											<% if cIntero(rs("ddt_destinazione_id")) > 0 then
												sql = "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi = " & rs("ddt_destinazione_id")
												rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
												<td class="content" colspan="3">
													<%= ContactAddress(rsc) %>
												</td>
												<% rsc.close %>
											<% else %>
												<td class="note" colspan="3">Destinazione non inserita</td>
											<% end if %>
										</tr>
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
