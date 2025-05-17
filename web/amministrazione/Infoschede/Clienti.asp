<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" --> 

<%
dim conn, dicitura, profilo, profiloS, id_profilo, sql, sql_filtri, sezione_anagrafiche
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")

set dicitura = New testata
dicitura.iniz_sottosez(0)

if cString(request("PROFILO"))<>"" then
	Session("PROFILO") = request("PROFILO")
end if

sezione_anagrafiche = false

sql = "SELECT pro_nome_it FROM gtb_profili WHERE pro_id = "
select case Session("PROFILO")
	case "trasportatori"
		id_profilo = TRASPORTATORI
	case "costruttori"
		id_profilo = COSTRUTTORI
	case else
		id_profilo = 0
		sezione_anagrafiche = true
end select
if cIntero(id_profilo) > 0 then
	sql = sql & id_profilo
	profilo = GetValueList(conn, NULL, sql)
	sql = Replace(sql, "pro_nome_it", "pro_codice")
	profiloS = GetValueList(conn, NULL, sql)
	
	dicitura.sezione = "Gestione "&lCase(profilo)&" - elenco"
	dicitura.puls_new = "NUOVO " & uCase(profiloS)
	dicitura.link_new = "ClientiGestione.asp?PROFILO="&Session("PROFILO")&"&IDPROFILO=" & id_profilo
else
	if cString(Session("riv_profilo"))="" then 
		Session("riv_profilo") = CLIENTI_PRIVATI&","&CLIENTI_PROFESSIONALI&","&RIVENDITORI&","&SUPERVISORI_NEGOZI
	end if
	profilo = "Anagrafiche clienti"
	profiloS = "Anagrafica cliente"
	dicitura.sezione = "Gestione anagrafiche clienti - elenco"
	'dicitura.puls_new = "nuovo:;CLIENTE PRIVATO;CLIENTE PROFESSIONALE;RIVENDITORE;SUPERVISORE NEGOZI;"
	'dicitura.link_new = ";ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO="&CLIENTI_PRIVATI&";ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO="&CLIENTI_PROFESSIONALI&";" & _
	'					"ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO="&RIVENDITORI&";ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO="&SUPERVISORI_NEGOZI
	dicitura.puls_new = "nuovo:;"
	dicitura.link_new = ";"
	dicitura.puls_2a_riga.Add "CLIENTE PRIVATO","ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO="&CLIENTI_PRIVATI
	dicitura.puls_2a_riga.Add "CLIENTE PROFESSIONALE","ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO="&CLIENTI_PROFESSIONALI
	dicitura.puls_2a_riga.Add "RIVENDITORE","ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO="&RIVENDITORI
	dicitura.puls_2a_riga.Add "SUPERVISORE NEGOZI","ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO="&SUPERVISORI_NEGOZI
end if

dicitura.scrivi_con_sottosez()  

dim rs, rsv, rsr, pager, i, colore

set rs = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator
sql = ""



'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("riv_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("riv_")
	else
		Session("riv_profilo") = CLIENTI_PRIVATI&","&CLIENTI_PROFESSIONALI&","&RIVENDITORI&","&SUPERVISORI_NEGOZI
	end if
end if



dim id_centro_assistenza
id_centro_assistenza = GetIdCentroAssistenzaLoggato()
if id_centro_assistenza > 0 then
	sql = " AND sc_centro_assistenza_id = " & id_centro_assistenza
end if


if id_centro_assistenza > 0 then
	'se l'utente è centro assistenza (oppure officina) mostro solo i clienti inseriti dell'utente stesso o i clienti associati alle schede collegate all'utente.
	sql = " AND ((cnt_insAdmin_id = "&CIntero(Session("ID_ADMIN"))&" OR cnt_modAdmin_id = "&CIntero(Session("ID_ADMIN"))&") " & _
		  " OR (EXISTS(SELECT ag_admin_id FROM gtb_agenti WHERE ag_admin_id = cnt_insAdmin_id AND ag_id = "&id_centro_assistenza&") " & _
		  " OR EXISTS(SELECT ut_NextCom_id FROM tb_utenti WHERE ut_NextCom_id = IDElencoIndirizzi AND " & _
		  " 				EXISTS(SELECT sc_cliente_id FROM sgtb_schede WHERE sc_cliente_id = ut_id AND sc_centro_assistenza_id = "&id_centro_assistenza&")))) "

end if


if sezione_anagrafiche then
	'ricerca per profilo
	if Session("riv_profilo")<>"" then
		sql = sql & " AND riv_profilo_id IN ( "
		'clienti privati
		if instr(1, Session("riv_profilo"), CLIENTI_PRIVATI, vbTextCompare)>0  then
			sql = sql & CLIENTI_PRIVATI & ","
		end if
		'clienti professionali
		if instr(1, Session("riv_profilo"), CLIENTI_PROFESSIONALI, vbTextCompare)>0  then
			sql = sql & CLIENTI_PROFESSIONALI & ","
		end if
		'rivenditori
		if instr(1, Session("riv_profilo"), RIVENDITORI, vbTextCompare)>0  then
			sql = sql & RIVENDITORI & ","
		end if
		'supervisori negozi
		if instr(1, Session("riv_profilo"), SUPERVISORI_NEGOZI, vbTextCompare)>0  then
			sql = sql & SUPERVISORI_NEGOZI & ","
		end if
		sql = left(sql, len(sql)-1) & " )"
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

'......................................................................................................
'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
CALL ADDON__CLIENTI__ricerca_form_parse(conn, sql)
'......................................................................................................

sql_filtri = sql

sql = " SELECT riv_profilo_id, ut_abilitato, lingua, lingua_nome_IT, IDElencoIndirizzi, riv_profilo_id, riv_id, " & _
	  "  NomeElencoIndirizzi, CognomeElencoIndirizzi, isSocieta, NomeOrganizzazioneElencoIndirizzi, IndirizzoElencoIndirizzi, " & _
	  "  LocalitaElencoIndirizzi, CittaElencoIndirizzi, ut_login, cnt_insAdmin_id, cnt_modAdmin_id" & _
	  " FROM gv_rivenditori r INNER JOIN tb_cnt_lingue cl ON r.lingua=cl.lingua_codice "
if cIntero(id_profilo) > 0 then
	sql = sql & "WHERE (riv_profilo_id IN ("&id_profilo&")) " & sql_filtri
else
	if cString(Session("riv_profilo"))="" AND session("PROFILO")="anagrafiche_clienti" then
		sql = sql & "WHERE (1=0) " & sql_filtri
	else
		sql = sql & "WHERE (1=1) " & sql_filtri
	end if
end if
sql = sql & " ORDER BY ModoRegistra"

session("B2B_"&UCase(Session("PROFILO"))&"_SQL") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>

<% if Session("PROFILO") = "trasportatori" OR Session("PROFILO") = "costruttori" then %>
	<div id="content">
<% else %>
	<div id="content_abbassato">
<% end if%>
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
									<% if sezione_anagrafiche then %>
										<tr><th colspan="2" <%= Search_Bg("riv_profilo") %>>PROFILO</th></tr>
										<tr>
											<td style="background:<%=COLOR_CLIENTI_PRIVATI%>;" class="content" colspan="2">
												<input type="checkbox"  class="checkbox" name="search_profilo" value="<%=CLIENTI_PRIVATI%>" <%= chk(instr(1, Session("riv_profilo"), CLIENTI_PRIVATI, vbTextCompare)>0) %>>
												<%=LABEL_CLIENTI_PRIVATI%>
											</td>
										</tr>
										<tr>
											<td style="background:<%=COLOR_CLIENTI_PROFESSIONALI%>;" class="content" colspan="2">
												<input type="checkbox"  class="checkbox" name="search_profilo" value="<%=CLIENTI_PROFESSIONALI%>" <%= chk(instr(1, Session("riv_profilo"), CLIENTI_PROFESSIONALI, vbTextCompare)>0) %>>
												<%=LABEL_CLIENTI_PROFESSIONALI%>
											</td>
										</tr>
										<tr>
											<td style="background:<%=COLOR_RIVENDITORI%>;" class="content" colspan="2">
												<input type="checkbox" class="checkbox" name="search_profilo" value="<%=RIVENDITORI%>" <%= chk(instr(1, Session("riv_profilo"), RIVENDITORI, vbTextCompare)>0) %>>
												<%=LABEL_RIVENDITORI%>
											</td>
										</tr>
										<tr>
											<td style="background:<%=COLOR_SUPERVISORI_NEGOZI%>;" class="content" colspan="2">
												<input type="checkbox" class="checkbox" name="search_profilo" value="<%=SUPERVISORI_NEGOZI%>" <%= chk(instr(1, Session("riv_profilo"), SUPERVISORI_NEGOZI, vbTextCompare)>0) %>>
												<%=LABEL_SUPERVISORI_NEGOZI%>
											</td>
										</tr>
									<% end if %>
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
									<tr><th colspan="2" <%= Search_Bg("riv_denominazione") %>>NOME / DENOMINAZIONE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_denominazione" value="<%= TextEncode(session("riv_denominazione")) %>" style="width:100%;">
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
									<% if not Session("PROFILO") = "trasportatori" then %>
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
									<% end if %>
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
					</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco <%=lCase(profilo)%>
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %>&nbsp;<%=lCase(profilo)%> in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0">
										<tr>
											<%
											if sezione_anagrafiche then
												select case rs("riv_profilo_id")
													case CLIENTI_PRIVATI
														colore = "style=""background:" & COLOR_CLIENTI_PRIVATI & """"
													case CLIENTI_PROFESSIONALI
														colore = "style=""background:" & COLOR_CLIENTI_PROFESSIONALI & """"
													case RIVENDITORI
														colore = "style=""background:" & COLOR_RIVENDITORI & """"
													case SUPERVISORI_NEGOZI
														colore = "style=""background:" & COLOR_SUPERVISORI_NEGOZI & """"
													case else
														colore = ""
												end select
											end if
											%>
											<td <%=colore%> class="<%= IIF(rs("ut_abilitato") OR Session("PROFILO")="trasportatori", "header", "header_disabled") %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<% if rs("lingua")<>"" then %>
															<td style="padding-right:4px; vertical-align:bottom;">
																<img src="../grafica/flag_mini_<%= rs("lingua") %>.jpg" alt="Lingua: <%= rs("lingua_nome_IT") %>">
															</td>
														<% end if %>
								
														<td style="font-size: 1px;">
															<% 'aggiunta: se non sei admin, deve permettere la modifica solo se l'anagrafica l'ha aggiunta l'amministratore loggato
															   if Session("INFOSCHEDE_ADMIN")<>"" OR rs("cnt_insAdmin_id")=cIntero(Session("ID_ADMIN")) OR rs("cnt_modAdmin_id")=cIntero(Session("ID_ADMIN")) then %>
																<a class="button" href="ClientiGestione.asp?ID=<%= rs("IDElencoIndirizzi") %>&PROFILO=<%=Session("PROFILO")%>&IDPROFILO=<%=rs("riv_profilo_id")%>">
																	MODIFICA
																</a>
															<% else %>
																<a class="button_disabled" title="Impossibile modificare l'anagrafica: &egrave; stata inserita da un'altro utente.">
																	MODIFICA
																</a>
															<% end if %>
															&nbsp;
														<% if Session("INFOSCHEDE_ADMIN")<>"" OR rs("cnt_insAdmin_id")=cIntero(Session("ID_ADMIN")) OR rs("cnt_modAdmin_id")=cIntero(Session("ID_ADMIN")) then %>
															<% 	if GetValueList(conn, rsv, "SELECT COUNT(*) FROM sgtb_ddt WHERE ddt_trasportatore_id="&rs("riv_id")&" OR ddt_cliente_id="&rs("riv_id")) > 0 then %>
																<a class="button_disabled" title="Impossibile cancellare l'anagrafica: documento di ritiro o riconsegna associati.">
																	CANCELLA
																</a>
															<% 	elseif GetValueList(conn, rsv, "SELECT COUNT(*) FROM sgtb_schede WHERE sc_cliente_id="& rs("riv_id")) > 0 then %>
																<a class="button_disabled" title="Impossibile cancellare l'anagrafica: presenti schede associate.">
																	CANCELLA
																</a>
															<% 	else %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CLIENTI','<%= rs("riv_id") %>');">
																	CANCELLA
																</a>
															<%	end if %>
														<% else %>
															<a class="button_disabled" title="Impossibile cancellare l'anagrafica: &egrave; stata inserita da un'altro utente.">
																CANCELLA
															</a>
														<% end if %>
														</td>
													</tr>
												</table>
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
											<td class="label" style="width:25%;">indirizzo principale:</td>
											<td class="content" colspan="2">
												<%= rs("IndirizzoElencoIndirizzi") %>
												&nbsp;<%= rs("LocalitaElencoIndirizzi") %>
												&nbsp;<%= rs("CittaElencoIndirizzi") %>
											</td>
											<td class="content_right" nowrap>
												<%
												sql = " SELECT TOP 1 sc_id FROM sgtb_schede WHERE sc_cliente_id IN ("&rs("riv_id")&") "
												if cIntero(GetValueList(conn, NULL, sql))>0 then %>
													<a href="javascript:void(0)" class="button_L2" href="../../Schede.asp"
														onclick="OpenAutoPositionedScrollWindow('Schede.asp?sch_riv_id=<%=rs("riv_id")%>', 'SCHEDE_RIV_<%=rs("riv_id")%>', 800, 800, true)" title="Click per visualizzare la schede create da questo contatto" <%= ACTIVE_STATUS %>>
														SCHEDE ASSOCIATE
													</a>
												<% else %>
													&nbsp;
												<% end if %>
											</td>
										</tr>
										<% if not Session("PROFILO") = "trasportatori" then %>
											<tr>
												<td class="label" style="width:10%;">login:</td>
												<td class="label_no_width" colspan="3"><%= rs("ut_login") %></td>
											</tr>
										<% end if %>
										
										<%sql = "SELECT id_tipoNumero, nome_TipoNumero FROM tb_TipNumeri WHERE id_tipoNumero " &_
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
set rsv = nothing
set rsr = nothing
set conn = nothing%>
