<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione agente - elenco"
dicitura.puls_new = "NUOVO AGENTE"
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
    sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("ag_denominazione"))
end if

'filtra per codice
if Session("ag_codice")<>"" then
    sql = sql & " AND " &SQL_FullTextSearch(Session("ag_codice"), "ag_codice")
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

'filtra per supervisore
if Session("ag_supervisione") = "0" then
	sql = sql & " AND NOT "& SQL_IsTrue(conn, "ag_supervisore")
elseif Session("ag_supervisione") = "1" then
	sql = sql & " AND "& SQL_IsTrue(conn, "ag_supervisore")
end if

sql = "SELECT *, (SELECT COUNT(*) FROM gtb_rivenditori WHERE riv_agente_id = gv_agenti.ag_id) AS N_CLIENTI " + _
	  " FROM gv_agenti INNER JOIN tb_cnt_lingue ON gv_agenti.lingua=tb_cnt_lingue.lingua_codice " + _
	  " WHERE (1=1) " + sql + " ORDER BY ModoRegistra"
session("B2B_AGENTI_SQL") = sql
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
									<tr><th <%= Search_Bg("ag_supervisione") %>>PORTAFOGLIO CLIENTI</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="checkbox" class="checkbox" name="search_supervisione" value="0" <%= chk(instr(1, session("ag_supervisione"), "0", vbTextCompare)>0) %>>
											solo clienti associati
										</td>
									</tr>
									<tr>
										<td class="content supervisore" colspan="2">
											<input type="checkbox" class="checkbox" name="search_supervisione" value="1" <%= chk(instr(1, session("ag_supervisione"), "1", vbTextCompare)>0) %>>
											tutti i clienti
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
									<tr><th <%= Search_Bg("ag_codice") %>>CODICE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_codice" value="<%= TextEncode(session("ag_codice")) %>" style="width:100%;">
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
									<tr>
										<td class="footer">
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
											<% 
											sql = session("B2B_AGENTI_SQL")
											sql = "SELECT * " & right(sql, len(sql) + 1 - instr(1, sql, "FROM gv_agenti", vbTextCompare))
											%>
											<% CALL ExportContattiInRubrica(sql, "IDElencoIndirizzi", "", "") %>
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
						Elenco agenti
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> agenti in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
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
															<% CALL WriteCampoCerca("Ordini.asp", "agente", rs("IDElencoIndirizzi"), "ORDINI", "button") %>
															&nbsp;
															<% CALL WriteCampoCerca("Clienti.asp", "agente", rs("ag_id"), "CLIENTI", "button") %>
															&nbsp;
															<a class="button" href="AgentiGestione.asp?ID=<%= rs("IDElencoIndirizzi") %>">
																MODIFICA
															</a>
															&nbsp;
															<% if cInteger(rs("N_CLIENTI"))=0 then %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('AGENTI','<%= rs("ag_id") %>');">
																	CANCELLA
																</a>
															<% else %>
																<a class="button_disabled" title="impossibile cancellare l'agente: &egrave; associato ad almeno un cliente" <%= ACTIVE_STATUS %>>
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
																case 6	'email %>
																	<a href="mailto:<%= rsr("ValoreNumero") %>"><%= rsr("ValoreNumero") %></a>
																<% case 7	'web %>
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
set rsv = nothing
set rsr = nothing
set conn = nothing%>
