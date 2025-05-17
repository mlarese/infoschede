<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "PROFILI CLIENTI"
dicitura.links(1) = "ClientiProfili.asp"
dicitura.sezione = "Gestione clienti - elenco"
dicitura.puls_new = "NUOVO CLIENTE"
dicitura.link_new = "ClientiGestione.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, rsv, rsr, sql, pager, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("riv_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("riv_")
	end if
end if

'filtra per iniziali
if Session("riv_iniziali")<>"" then
	sql = sql & " AND " & SQL_Ucase(conn) & "(LEFT(ModoRegistra, 1)) IN (" & Session("riv_iniziali") & ")"
end if

'filtra per nome
if Session("riv_denominazione")<>"" then
    sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("riv_denominazione"))
end if

'filtra per indirizzo
if Session("riv_indirizzo")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Indirizzo(conn, Session("riv_indirizzo"))
end if

'filtra per citta
if Session("riv_citta")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("riv_citta"), "CittaElencoIndirizzi")
end if

'filtra per profilo
if Session("riv_profilo")<>"" then
	sql = sql & " AND riv_profilo_id IN (" & Session("riv_profilo") & ")"
end if

'filtra per listino
if Session("riv_listino")<>"" then
	sql = sql & " AND riv_listino_id=" & Session("riv_listino")
end if

'filtra per agente
if Session("riv_agente")<>"" then
	sql = sql & " AND riv_agente_id=" & Session("riv_agente")
end if

'codice cliente
if Session("riv_codice")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("riv_codice"), "riv_codice")
end if


'filtra per e-mail
if Session("riv_email")<>"" then
	sql = sql & " AND IDElencoIndirizzi IN (SELECT id_indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero = 6 " & _
				" 							 AND " & SQL_FullTextSearch(Session("riv_email"), "ValoreNumero") & " ) "
end if

'filtra per login
if Session("riv_login")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("riv_login"), "ut_login")
end if

'filtra per abilitazione dell'account utente
if Session("riv_abilitato")<>"" then
	if not (instr(1, Session("riv_abilitato"), "A", vbTextCompare)>0 AND instr(1, Session("riv_abilitato"), "N", vbTextCompare)>0 ) then
		if instr(1, Session("riv_abilitato"), "A", vbTextCompare)>0 then
			sql = sql &" AND "& SQL_IsTrue(conn, "UT_Abilitato")
		elseif instr(1, Session("riv_abilitato"), "N", vbTextCompare)>0 then
			sql = sql & " AND NOT (" & SQL_IsTrue(conn, "UT_Abilitato") & ") "
		end if
	end if
end if

'filtra per attivazione dell'utente
if Session("riv_attivo")<>"" then
	if not (instr(1, Session("riv_attivo"), "A", vbTextCompare)>0 AND instr(1, Session("riv_attivo"), "N", vbTextCompare)>0 ) then
		if instr(1, Session("riv_attivo"), "A", vbTextCompare)>0 then
			sql = sql &" AND "& SQL_IsTrue(conn, "riv_attivo")
		elseif instr(1, Session("riv_attivo"), "N", vbTextCompare)>0 then
			sql = sql & " AND NOT (" & SQL_IsTrue(conn, "riv_attivo") & ") "
		end if
	end if
end if

'......................................................................................................
'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
CALL ADDON__CLIENTI__ricerca_form_parse(conn, sql)
'......................................................................................................

sql = "SELECT * FROM ((gv_rivenditori r INNER JOIN gtb_valute v ON r.riv_valuta_id=v.valu_id) "& _
	  "INNER JOIN gtb_listini l ON r.riv_listino_id=l.listino_id) "& _
	  "INNER JOIN tb_cnt_lingue cl ON r.lingua=cl.lingua_codice "& _
	  "WHERE (1=1) "& sql & _
	  " ORDER BY ModoRegistra"
session("B2B_RIVENDITORI_SQL") = sql


CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
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
									<!--
									<tr><th colspan="2" <%= Search_Bg("riv_iniziali") %>>INIZIALI</th></tr>
									<tr>
										<td colspan="2">
											<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
												<tr>
													<%for i=asc("A") to asc("Z")%>
					    								<TD class="content">
															<INPUT class="checkbox" type="checkbox" name="search_iniziali" value="'<%=chr(i)%>'" <%if instr(1, Session("riv_iniziali"), chr(i), vbTextCompare)>0 then %> checked <% end if %>>
															<%=chr(i)%>
														</TD>
					    								<%if i mod 4 = 0 then%>
															</tr>
															<tr>
														<%end if
													next %>
													<td class="content" colspan="2">&nbsp;</td>
												</tr>
											</table>
										</td>
									</tr>
									-->
									<tr><th colspan="2" <%= Search_Bg("riv_attivo") %>>STATO ATTIVAZIONE</th></tr>
									<tr>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_attivo" style="float:left;" value="A" <%= IIF(instr(1, Session("riv_attivo"), "A", vbTextCompare)>0, " checked", "") %>>
											<% WriteColor("#67c567") %>attivo
										</td>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_attivo" style="float:left;"  value="N" <%= IIF(instr(1, Session("riv_attivo"), "N", vbTextCompare)>0, " checked", "") %>>
											<% WriteColor("#f94d4d") %>non attivo
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("riv_abilitato") %>>STATO ACCESSO</th></tr>
									<tr>
										<td class="content_b">
											<input type="checkbox" class="checkbox" name="search_abilitato" value="A" <%= IIF(instr(1, Session("riv_abilitato"), "A", vbTextCompare)>0, " checked", "") %>>
											abilitato
										</td>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_abilitato" value="N" <%= IIF(instr(1, Session("riv_abilitato"), "N", vbTextCompare)>0, " checked", "") %>>
											non abilitato
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("riv_denominazione") %>>NOME / DENOMINAZIONE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_denominazione" value="<%= TextEncode(session("riv_denominazione")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("riv_codice") %>>CODICE CLIENTE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_codice" value="<%= TextEncode(session("riv_codice")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("riv_indirizzo") %>>INDIRIZZO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_indirizzo" value="<%= TextEncode(session("riv_indirizzo")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("riv_citta") %>>CITT&Agrave;</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_citta" value="<%= TextEncode(session("riv_citta")) %>" style="width:100%;">
										</td>
									</tr>									
									<% sql = "SELECT * FROM gtb_profili ORDER BY pro_nome_it" %>
									<% if cString(GetValueList(conn,NULL,sql))<>"" then %>
										<tr><th colspan="2" <%= Search_Bg("riv_profilo") %>>PROFILO</th></tr>
										<tr>
											<td class="content" colspan="2">
												<% CALL	dropDown(conn, sql, "pro_id", "pro_nome_it", "search_profilo", Session("riv_profilo"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
											</td>
										</tr>
									<% end if %>
									<tr><th colspan="2" <%= Search_Bg("riv_listino") %>>LISTINO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% sql = "SELECT listino_id, listino_codice FROM gtb_listini WHERE ISNULL(listino_offerte, 0)=0 ORDER BY listino_codice"
											CALL dropDown(conn, sql, "listino_id", "listino_codice", "search_listino", session("riv_listino"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("riv_agente") %>>AGENTE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% 	sql = " SELECT ag_id, "& _
													  "(CognomeElencoIndirizzi "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) &" NomeElencoIndirizzi "& SQL_concat(conn) &" ' - ' " & SQL_concat(conn) &" NomeOrganizzazioneElencoIndirizzi) AS NOMINATIVO " &_
													  " FROM gv_agenti " & _
													  " ORDER BY ModoRegistra"
												CALL dropDown(conn, sql, "ag_id", "NOMINATIVO", "search_agente", session("riv_agente"), false, " style=""width:100%;""", LINGUA_ITALIANO)
											%> 
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("riv_email") %>>E-MAIL</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_email" value="<%= TextEncode(session("riv_email")) %>" style="width:100%;">
										</td>
									</tr>	
									<tr><th colspan="2" <%= Search_Bg("riv_login") %>>LOGIN</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_login" value="<%= TextEncode(session("riv_login")) %>" style="width:100%;">
										</td>
									</tr>
									
									<%
									'......................................................................................................
									'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
									CALL ADDON__CLIENTI__ricerca_form(conn)
									'......................................................................................................
									%>
									<tr>
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr><td style="font-size:4px;">&nbsp;</td></tr>
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
									<caption>Strumenti</caption>
									<tr>
										<td class="content_center">
											<% CALL ExportContattiInRubrica(session("B2B_RIVENDITORI_SQL"), "IDElencoIndirizzi", "", "") %>
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
						Elenco clienti
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> clienti in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0">
										<tr>
											<td class="<%= IIF(rs("ut_abilitato"), "header", "header_disabled") %>" colspan="4" title="riv_id=<%=rs("riv_id")%> - cliente: <%= rs("IDElencoIndirizzi") %>">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<% if rs("lingua")<>"" then %>
															<td style="padding-right:4px; vertical-align:bottom;">
																<img src="../grafica/flag_mini_<%= rs("lingua") %>.jpg" alt="Lingua: <%= rs("lingua_nome_IT") %>">
															</td>
														<% end if %>
														<td style="font-size: 1px;">
															<% CALL WriteCampoCerca("Ordini.asp", "denominazione", rs("ModoRegistra"), "ORDINI", "button") %>
															&nbsp;
															<a class="button" href="ClientiGestione.asp?ID=<%= rs("IDElencoIndirizzi") %>">
																MODIFICA
															</a>
															&nbsp;
														<% 	if GetValueList(conn, rsv, "SELECT COUNT(*) FROM gtb_ordini WHERE ord_riv_id="& rs("riv_id")) > 0 then
																sql = "SELECT sito_nome FROM tb_siti WHERE id_sito IN (" & rs("ApplicationsLocker") & "0 )"%>
															<a class="button_disabled" title="Impossibile cancellare il cliente: sono presenti ordini associati">
																CANCELLA
															</a>
														<% 	else %>
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CLIENTI','<%= rs("riv_id") %>');">
																CANCELLA
															</a>
														<%	end if %>
														</td>
													</tr>
												</table>
												<% if cBoolean(rs("riv_attivo"), false) then %>
													<% WriteColor("#67c567") %>
												<% else %>
													<% WriteColor("#f94d4d") %>
												<% end if %>
												<%=ContactName(rs)%>
											</td>
										</tr>
										<% if rs("isSocieta") then 
											if rs("CognomeElencoIndirizzi") & rs("NomeElencoIndirizzi")<>"" then  %>
												<tr>
													<td class="label">contatto:</td>
													<td class="content" colspan="3"><%= rs("CognomeElencoIndirizzi") %>&nbsp;<%= rs("NomeElencoIndirizzi") %></td>
											<% end if
										else 
											if rs("NomeOrganizzazioneElencoIndirizzi")<>"" then  %>
												<tr>
													<td class="label">ente:</td>
													<td class="content" colspan="3"><%= rs("NomeOrganizzazioneElencoIndirizzi") %></td>
											<% end if
										end if %>
										</tr>
										<tr>
											<td class="label">codice aziendale:</td>
											<td class="content" width="38%"><%= rs("riv_codice") %></td>
											<td class="label">login:</td>
											<td class="content"><%= rs("ut_login") %></td>
										</tr>
										<tr>
											<td class="label">listino:</td>
											<td class="content" colspan="2">
												<a href="listiniMod.asp?ID=<%= rs("listino_id") %>" title="apri la scheda del listino in una nuova finestra" target="_blank" <%= ACTIVE_STATUS %>>
													<% if rs("listino_Base_attuale") then %>
														listino base in vigore (mantenuto automaticamente)
													<% else %>
														<%= rs("listino_codice")%>
													<% end if %>
												</a>
											</td>
											<td class="content_right" nowrap>
												<% CALL WriteCampoCerca("ListiniPrezzi_RigaPerRiga.asp?ID=" & rs("listino_id"), "personalizzato", 1, "prezzi personalizzati", "button_L2") %>
											</td>
										</tr>
										<%if CInteger(rs("riv_agente_id")) > 0 then 
											sql = "SELECT * FROM gv_agenti WHERE ag_id="& rs("riv_agente_id")
											rsv.Open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
											if not rsv.eof then %>
												<tr>
													<td class="label">agente:</td>
													<td class="content" colspan="3"><%= ContactFullName(rsv) %></td>
												</tr>
											<%end if
											rsv.Close
										end if %>
										<tr>
											<td class="label" style="width:20%;">indirizzo principale:</td>
											<td class="content" colspan="3">
												<%= rs("IndirizzoElencoIndirizzi") %>
												&nbsp;<%= rs("LocalitaElencoIndirizzi") %>
												&nbsp;<%= rs("CittaElencoIndirizzi") %>
											</td>
										</tr>
										<%sql = "SELECT * FROM tb_TipNumeri WHERE id_tipoNumero " &_
												 " IN (SELECT id_TipoNumero FROM tb_ValoriNumeri WHERE id_indirizzario=" & rs("IDElencoIndirizzi") & ") "
										rsv.Open sql, conn, AdOpenForwardOnly, adLockReadOnly, adCmdText
										while not rsv.eof
											sql = "SELECT id_TipoNumero, ValoreNumero FROM tb_ValoriNumeri " &_
												  " WHERE id_TipoNumero=" & rsv("id_tipoNumero") & " AND  id_Indirizzario=" & rs("IDElencoIndirizzi")
											rsr.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
											if not rsr.eof then%>
												<tr>
													<td class="label" nowrap><%= Lcase(rsv("nome_TipoNumero")) %>:</td>
													<td class="content" colspan="3">
														<%while not rsr.eof
															select case rsr("id_TipoNumero")
																case 6	'email 
																%>
																	<a href="mailto:<%= rsr("ValoreNumero") %>"><%= rsr("ValoreNumero") %></a>
																<% case 7	'web 
																%>
																	<a href="http://<%= rsr("ValoreNumero") %>" target="_blank"><%= rsr("ValoreNumero") %></a>
																<% case else %>
																	<%= rsr("ValoreNumero") %>
															<%end select
															rsr.movenext
															if not rsr.eof then%>
																,&nbsp;
															<%end if
														wend %>
													</td>
												</tr>
											<%end if
											rsr.close
											rsv.MoveNext
										wend
										rsv.Close
										
										
										'......................................................................................................
										'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
										CALL ADDON__CLIENTI__record_elenco(conn, rs)
										'......................................................................................................
										
										
										sql = " SELECT *, (SELECT ut_id FROM tb_utenti WHERE ut_nextCom_id = tb_indirizzario.IDElencoIndirizzi) AS UTENTE " + _
											  " FROM tb_Indirizzario WHERE CntRel=" & rs("IDElencoIndirizzi") & " ORDER BY ModoRegistra "
										rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
										if not rsr.eof then%>
											<tr>
												<td class="label">contatti interni / sedi alternative:</td>
												<td colspan="3">
													<% if rsr.recordcount>2 then %>
														<span class="overflow">
													<% end if %>
														<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
															<tr>
																<th class="L2" width="45%">contatto / sede</th>
																<th class="l2_center" width="8%">sede</th>
																<th class="l2_center" width="9%">accesso</th>
																<th class="l2_center"  width="11%">operazioni</th>
															</tr>
															<% while not rsr.eof %>
																<tr>
																	<td class="content" title="ruolo / qualifica: <%= rsr("QualificaElencoIndirizzi") %>"><%= ContactFullName(rsr) %></td>
																	<td class="content_center">
																		<input type="checkbox" class="checkbox" disabled <%= chk(rsr("isSocieta")) %> title="<%= IIF(rsr("isSocieta"), "sede alternativa o periferica", "contatto interno") %>">
																	</td>
																	<td class="content_center">
																		<input type="checkbox" class="checkbox" disabled <%= chk(cInteger(rsr("UTENTE"))>0) %> title="<%= IIF(cInteger(rsr("UTENTE"))>0, "con accesso subordinato", "senza accesso") %>">
																	</td>
																	<td class="content_center">
																		<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ClientiIntGestione.asp?ID=<%= rsr("IdElencoIndirizzi") %>', 'cntInt', 500, 450, true)">
																			MODIFICA
																		</a>
																	</td>
																</tr>
																<% rsr.movenext
															wend %>
														</table>
													<% if rsr.recordcount>2 then %>
														</span>
													<% end if %>
												</td>
											</tr>
										<% 	end if
										rsr.close%>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td class="footer" style="border-top:0px; text-align:left;">
								<% CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled") %>
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
<% if Session("B2B_HTTP_RESULT_SPEDIZ_CREDENZ_ACCESSO") <> "" then %>
	<script type="text/javascript">
		alert('Le credenziali di accesso sono state spedite.');
	</script>
	<% Session("B2B_HTTP_RESULT_SPEDIZ_CREDENZ_ACCESSO") = "" %>
<% end if %>
<% 
rs.close
conn.close 
set rs = nothing
set rsv = nothing
set rsr = nothing
set conn = nothing%>
