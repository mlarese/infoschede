<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione centri assistenza - elenco"
dicitura.puls_new = "NUOVO CENTRO ASSISTENZA"
dicitura.link_new = "AgentiGestione.asp"
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
	CALL SearchSession_Reset("ag_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("ag_")
	end if
end if

'filtra per iniziali
if Session("ag_iniziali")<>"" then
	sql = sql & " AND " & SQL_Ucase(conn) & "(LEFT(ModoRegistra, 1)) IN (" & Session("ag_iniziali") & ")"
end if

'filtra per nome
if Session("ag_denominazione")<>"" then
    sql = sql & " AND (" & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("ag_denominazione")) & _
				" OR IDElencoIndirizzi IN (SELECT cntRel FROM tb_Indirizzario WHERE "&SQL_FullTextSearch_Contatto_Nominativo(conn, Session("ag_denominazione"))&")) "
end if

'filtra per indirizzo
if Session("ag_indirizzo")<>"" then
    sql = sql & " AND " & SQL_FullTextSearch_Contatto_Indirizzo(conn, Session("ag_indirizzo"))
end if

'filtra per citta
if Session("ag_citta")<>"" then
     sql = sql & " AND " & SQL_FullTextSearch(Session("ag_citta"), "CittaElencoIndirizzi")
end if

'filtra per login
if Session("ag_login")<>"" then
    sql = sql & " AND " & SQL_FullTextSearch(Session("ag_login"), "admin_login")
end if

'filtra per email
if Session("ag_email")<>"" then
    sql = sql & " AND ((IDElencoIndirizzi IN (SELECT Id_Indirizzario FROM tb_valoriNumeri WHERE id_tipoNumero = 6 " & _
				" AND ValoreNumero LIKE '%" & Session("ag_email") & "%')) OR " & _
				" (IDElencoIndirizzi IN (SELECT cntRel FROM tb_indirizzario WHERE IDElencoIndirizzi IN " & _
				" (SELECT Id_Indirizzario FROM tb_valoriNumeri WHERE id_tipoNumero = 6 " & _
				" AND ValoreNumero LIKE '%" & Session("ag_email") & "%'))))"
end if



sql = "SELECT *, (SELECT COUNT(*) FROM gtb_rivenditori WHERE riv_agente_id = gv_agenti.ag_id) AS N_CLIENTI " + _
	  " FROM gv_agenti INNER JOIN tb_cnt_lingue ON gv_agenti.lingua=tb_cnt_lingue.lingua_codice " + _
	  " WHERE (1=1) " + sql + " ORDER BY ModoRegistra"
session("B2B_CENTRI_ASSISTENZA_SQL") = sql


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
										<td class="footer">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("ag_iniziali") %>>INIZIALI</th></tr>
									<tr>
										<td>
											<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
												<tr>
													<%for i=asc("A") to asc("Z")%>
					    								<TD class="content">
															<INPUT class="checkbox" type="checkbox" name="search_iniziali" value="'<%=chr(i)%>'" <%if instr(1, Session("ag_iniziali"), chr(i), vbTextCompare)>0 then %> checked <% end if %>>
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
									<tr><th <%= Search_Bg("ag_denominazione") %>>NOME / DENOMINAZIONE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_denominazione" value="<%= TextEncode(session("ag_denominazione")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("ag_indirizzo") %>>INDIRIZZO</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_indirizzo" value="<%= TextEncode(session("ag_indirizzo")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("ag_citta") %>>CITT&Agrave;</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_citta" value="<%= TextEncode(session("ag_citta")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("ag_login") %>>LOGIN</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_login" value="<%= TextEncode(session("ag_login")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("ag_email") %>>E-MAIL</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_email" value="<%= TextEncode(session("ag_email")) %>" style="width:100%;">
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
						Elenco centri assistenza
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> centri assistenza in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header<%= IIF(rs("ag_supervisore"), " supervisore", "") %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<% if rs("lingua")<>"" then %>
															<td style="padding-right:4px; vertical-align:bottom;">
																<img src="../grafica/flag_mini_<%= rs("lingua") %>.jpg" alt="Lingua: <%= rs("lingua_nome_IT") %>">
															</td>
														<% end if %>
														<td style="font-size: 1px;">
															<a class="button" href="AgentiGestione.asp?ID=<%= rs("IDElencoIndirizzi") %>">
																MODIFICA
															</a>
															&nbsp;
															<% if cInteger(rs("N_CLIENTI"))=0 then %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('AGENTI','<%= rs("ag_id") %>');">
																	CANCELLA
																</a>
															<% else %>
																<a class="button_disabled" title="impossibile cancellare il centro assistenza: &egrave; associato ad almeno un cliente" <%= ACTIVE_STATUS %>>
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
												</tr>
											<% end if
										else 
											if rs("NomeOrganizzazioneElencoIndirizzi")<>"" then  %>
												<tr>
													<td class="label">ente:</td>
													<td class="content" colspan="3"><%= rs("NomeOrganizzazioneElencoIndirizzi") %></td>
												</tr>
											<% end if
										end if %>
										<tr>
											<td class="label">login:</td>
											<td class="content"><%= rs("ut_login") %></td>
											<td class="label">codice</td>
											<td class="content" width="10%"><%= rs("ag_Codice") %></td>
										</tr>
										<tr>
											<td class="label">indirizzo:</td>
											<td class="content" colspan="3">
												<%= rs("IndirizzoElencoIndirizzi") %>
												&nbsp;<%= rs("LocalitaElencoIndirizzi") %>
												&nbsp;<%= rs("CittaElencoIndirizzi") %>
											</td>
										</tr>
										<% sql = "SELECT * FROM tb_TipNumeri WHERE id_tipoNumero " &_
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
										
										
										
										sql = " SELECT *, (SELECT ut_id FROM tb_utenti WHERE ut_nextCom_id = tb_indirizzario.IDElencoIndirizzi) AS UTENTE " + _
											  " FROM tb_Indirizzario WHERE CntRel=" & rs("IDElencoIndirizzi") & " ORDER BY ModoRegistra "
										rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
										if not rsr.eof then%>
											<tr>
												<td class="label">operatori:</td>
												<td colspan="3">
													<% if rsr.recordcount>2 then %>
														<span class="overflow">
													<% end if %>
														<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
															<tr>
																<th class="L2">operatore</th>
																<th class="L2" width="36%">e-mail</th>
																<th class="L2" width="16%">login</th>
																<th class="l2_center"  width="11%">operazioni</th>
															</tr>
															<% while not rsr.eof %>
																<tr>
																	<td class="content"><%= ContactFullName(rsr) %></td>
																	<td class="content">
																		<% sql = " SELECT ValoreNumero FROM tb_ValoriNumeri WHERE email_default = 1 AND id_TipoNumero = 6 " & _
																				 " AND id_Indirizzario = " & rsr("IDElencoIndirizzi") %>
																		<%= GetValueList(conn, NULL, sql) %>																	</td>
																	<td class="content">
																		<% sql = " SELECT ut_login FROM tb_utenti WHERE ut_nextCom_id = " & rsr("IDElencoIndirizzi") %>
																		<%= GetValueList(conn, NULL, sql) %>																	</td>
																	<td class="content_center">
																		<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('AgentiIntGestione.asp?ID=<%= rsr("IdElencoIndirizzi") %>', 'ageCntInt', 550, 450, true)">
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
