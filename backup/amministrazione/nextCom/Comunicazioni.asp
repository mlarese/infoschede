<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->

<%'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Titolo_sezione, Sezione, HREF, Action
Titolo_sezione = "Comunicazioni in uscita - elenco"
HREF = ";ComunicazioniNew_Wizard_1.asp?new=1&type=" & MSG_EMAIL
Action = "invia nuova:;EMAIL"
if Session("FAX_ABILITATI") then
	HREF = HREF + ";ComunicazioniNew_Wizard_1.asp?new=1&type=" & MSG_FAX
	Action = Action + ";FAX"
end if
if Session("SMS_ABILITATI") then
	HREF = HREF + ";ComunicazioniNew_Wizard_1.asp?new=1&type=" & MSG_SMS
	Action = Action + ";SMS"
end if

SSezioniText = "TIPOLOGIE NEWSLETTER;"
SSezioniLink = "NewsletterTip.asp;"

%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine


'************************************************************************************************************
%>
<%
dim conn, rs, rsd, rse, rsr, sql, rubriche_visibili, Pager, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")
set rse = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("comunicazioni_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("comunicazioni_")
	end if
end if

if Session("COM_ADMIN")<>"" then
	sql = "SELECT * FROM tb_email WHERE (1=1) "
else
	sql = "SELECT * FROM tb_email WHERE email_id IN " + _
		  "		(SELECT log_email_id FROM log_cnt_email INNER JOIN rel_rub_ind " + _
		  "		 ON log_cnt_email.log_cnt_id = rel_rub_ind.id_indirizzo " + _
		  "		 WHERE rel_rub_ind.id_rubrica IN (" + IIF(Session("comunicazioni_rubriche")<>"", Session("comunicazioni_rubriche"), rubriche_visibili) + ") " + _
		  "		) "
end if

'filtra per tipo di messaggio
if Session("comunicazioni_tipo")<>"" then
	sql = sql & " AND (1=0 "
	if instr(1, Session("comunicazioni_tipo"), MSG_EMAIL, vbTextCompare)>0  then
		sql = sql & " OR ( email_tipi_messaggi_id = " & MSG_EMAIL & ") "
	end if
	if instr(1, Session("comunicazioni_tipo"), MSG_SMS, vbTextCompare)>0  then
		sql = sql & " OR ( email_tipi_messaggi_id = " & MSG_SMS & ") "
	end if
	if instr(1, Session("comunicazioni_tipo"), MSG_FAX, vbTextCompare)>0  then
		sql = sql & " OR ( email_tipi_messaggi_id = " & MSG_FAX & ") "
	end if
	sql = sql &")"
end if

'filtra per oggetto e testo
if Session("comunicazioni_oggetto") <> "" then
	sql = sql & " AND "& SQL_FullTextSearch(Session("comunicazioni_oggetto"), "email_object")
end if

'filtra per data di invio
if isDate(Session("comunicazioni_data_from")) then
	sql = sql & " AND email_data >=" & SQL_date(conn, Session("comunicazioni_data_from"))
end if
if isDate(Session("comunicazioni_data_to")) then
	sql = sql & " AND email_data <=" & SQL_date(conn, Session("comunicazioni_data_to"))
end if

'filtra per nome destinatario
if Session("comunicazioni_dest_nome") <> "" then
	sql = sql &" AND email_id IN (SELECT log_email_id FROM log_cnt_email l "& _
						    	 "INNER JOIN tb_indirizzario i ON l.log_cnt_id = i.IDElencoIndirizzi "& _
						    	 "WHERE "& SQL_FullTextSearch_Contatto_Nominativo(conn, Session("comunicazioni_dest_nome")) &")"
end if

'filtra per recapito destinatario
if Session("comunicazioni_dest_recapito") <> "" then
	sql = sql &" AND email_id IN (SELECT log_email_id FROM log_cnt_email "& _
						    	 "WHERE " & SQL_FullTextSearch(Session("comunicazioni_dest_recapito"), "log_email") & ")"
end if

'filtra per rubriche
if session("comunicazioni_rubriche") <> "" then
	if Session("comunicazioni_rubriche_tipo") = "C" then
		'filtra email tramite l'associazione contatti/rubriche
		sql = sql & " AND email_id IN (SELECT log_email_id FROM log_cnt_email em"& _
									" INNER JOIN rel_rub_ind rub ON em.log_cnt_id = rub.id_indirizzo"& _
									" WHERE id_rubrica IN ("& session("comunicazioni_rubriche") &"))"
	else
		'filtra email tramite il log di spedizione alle rubriche
		sql = Sql & " AND email_id IN (SELECT log_email_id FROM log_rubriche_email WHERE log_rubrica_id IN ("& session("comunicazioni_rubriche") &"))"
	end if
end if

'filtra per esito di invio
if Session("comunicazioni_esito")<>"" then
	sql = sql & " AND (1=0 "
	if instr(1, Session("comunicazioni_esito"), "I", vbTextCompare)>0  then				'inviata correttamente
		sql = sql & " OR ( NOT " & SQL_IsTrue(conn, "email_isBozza") & " AND email_id NOT IN (SELECT log_email_id FROM log_cnt_email WHERE NOT " & SQL_IsTrue(conn, "log_inviato_ok") & ")) "
	end if
	if instr(1, Session("comunicazioni_esito"), "E", vbTextCompare)>0  then				'inviata con errori
		sql = sql & " OR ( NOT " & SQL_IsTrue(conn, "email_isBozza") & " AND email_id IN (SELECT log_email_id FROM log_cnt_email WHERE NOT " & SQL_IsTrue(conn, "log_inviato_ok") & ")) "
	end if
	if instr(1, Session("comunicazioni_esito"), "S", vbTextCompare)>0  then				'bozza
		sql = sql &" OR " & SQL_IsTrue(conn, "email_isBozza")
	end if
	sql = sql &")"
end if
sql = sql & " ORDER BY email_data DESC, email_id DESC "
Session("SQL_COMUNICAZIONI") = sql			'serve per la visualizzazione dei record prec/succ
CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>
<div id="content">
	<form action="" method="post" id="ricerca" name="ricerca">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTE" style="width: 49%;">
									</td>
								</tr>
								<% if Session("FAX_ABILITATI") OR Session("SMS_ABILITATI") then %>
									<tr><th<%= Search_Bg("comunicazioni_tipo") %>>TIPO COMUNICAZIONE</th></tr>
									<tr>
										<td class="content">
											<input type="Checkbox" name="search_tipo" class="checkbox" value="<%= MSG_EMAIL %>" <%= chk(instr(1, Session("comunicazioni_tipo"), MSG_EMAIL, vbTextCompare)>0) %>>
											<img src="../grafica/icona_email.gif">
											Email
										</td>
									</tr>
									<% if Session("FAX_ABILITATI") then %>
										<tr>
											<td class="content">
												<input type="Checkbox" name="search_tipo" class="checkbox" value="<%= MSG_FAX %>" <%= chk(instr(1, Session("comunicazioni_tipo"), MSG_FAX, vbTextCompare)>0) %>>
												<img src="../grafica/icona_fax.gif">
												Fax
											</td>
										</tr>
									<% end if %>
									<% if Session("SMS_ABILITATI") then %>
										<tr>
											<td class="content">
												<input type="Checkbox" name="search_tipo" class="checkbox" value="<%= MSG_SMS %>" <%= chk(instr(1, Session("comunicazioni_tipo"), MSG_SMS, vbTextCompare)>0) %>>
												<img src="../grafica/icona_sms.gif">
												Sms
											</td>
										</tr>
									<% end if 
								end if%>
								<tr><th<%= Search_Bg("comunicazioni_esito") %>>ESITO DI INVIO</th></tr>
								<tr>
									<td class="content ok">
										<input type="Checkbox" name="search_esito" class="checkbox" value="I" <%= chk(instr(1, Session("comunicazioni_esito"), "I", vbTextCompare)>0) %>>
										inviate correttamente
									</td>
								</tr>
								<tr>
									<td class="content alert">
										<input type="Checkbox" name="search_esito" class="checkbox" value="E" <%= chk(instr(1, Session("comunicazioni_esito"), "E", vbTextCompare)>0) %>>
										inviate con errori di spedizione
									</td>
								</tr>
								<tr>
									<td class="content warning">
										<input type="Checkbox" name="search_esito" class="checkbox" value="S" <%= chk(instr(1, Session("comunicazioni_esito"), "S", vbTextCompare)>0) %>>
										salvate e non ancora inviate
									</td>
								</tr>
								<tr><th <%= Search_Bg("comunicazioni_oggetto") %>>OGGETTO</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_oggetto" value="<%= server.HTMLencode(session("comunicazioni_oggetto")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("comunicazioni_dest_nome;comunicazioni_dest_recapito") %>>DESTINATARIO</th></tr>
								<tr><th class="L2" <%= Search_Bg("comunicazioni_dest_nome") %>>per nome</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_dest_nome" value="<%= server.HTMLencode(session("comunicazioni_dest_nome")) %>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<th class="L2" <%= Search_Bg("comunicazioni_dest_recapito") %>>
										<% if Session("FAX_ABILITATI") OR Session("SMS_ABILITATI") then %>
											per email<%= IIF(Session("FAX_ABILITATI"), ", numero fax", "") %><%= IIF(Session("SMS_ABILITATI"), ", cellulare", "") %>
										<% else %>
											per indirizzo email
										<% end if %>
									</th>
								</tr>
								<tr>
									<td class="content">
										<input type="text" name="search_dest_recapito" value="<%= server.HTMLencode(session("comunicazioni_dest_recapito")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("comunicazioni_rubriche") %>>RUBRICHE DESTINATARI</th></tr>
								<tr>
									<td class="content">
										<script language="JavaScript" type="text/javascript">
											function ShowName(obj){
												var value = obj.options(obj.selectedIndex).text;
												if (value.length>35)
													alert(obj.options(obj.selectedIndex).text);
											}
										</script>
										<%sql = " SELECT " & _ 
													 IIF(DB_Type(conn) = DB_SQL, "(' ' + CAST(id_rubrica AS nvarchar(8)) + ' ') ", "(' ' & id_rubrica & ' ')") & " AS ID, " &_
													 " nome_rubrica FROM tb_rubriche " &_
													 " WHERE id_rubrica IN (" & rubriche_visibili & ") " &_
													 " ORDER BY rubrica_esterna, nome_rubrica"
										CALL dropDown(conn, sql, "ID", "nome_rubrica", "search_rubriche", Session("comunicazioni_rubriche"), true, _
													  "multiple size=""15"" style=""width:100%;"" onDblClick=""ShowName(this);""", LINGUA_ITALIANO)%>
									</td>
								</tr>
								<tr>
									<th class="L2">applica filtro:</th>
								</tr>
								<tr>
									<td class="content">
										<input type="radio" name="search_rubriche_tipo" class="checkbox" value="" <%= chk(Session("comunicazioni_rubriche_tipo")<>"C") %>>
										sulle rubriche "destinatarie"
									</td>
								</tr>
								<tr>
									<td class="content">
										<input type="radio" name="search_rubriche_tipo" class="checkbox" value="C" <%= chk(Session("comunicazioni_rubriche_tipo")="C") %>>
										tramite i contatti destinatari
									</td>
								</tr>
								<tr>
									<td class="note">
											Ctrl + Click per selezioni multiple.<br>
											Doppio click per visualizzare il nome.
									</td>
								</tr>
								<tr><th <%= Search_Bg("comunicazioni_data_from;comunicazioni_data_to") %>>DATA DI INVIO</td></tr>
								<tr><td class="label">a partire dal:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_from", Session("comunicazioni_data_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td class="label">fino al:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_to", Session("comunicazioni_data_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTE" style="width: 49%;">
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
											sql = session("SQL_COMUNICAZIONI")
											sql = "SELECT * FROM log_cnt_email INNER JOIN tb_email ON log_cnt_email.log_email_id = tb_email.email_id " & _
													right(sql, len(sql) + 1 - instr(1, sql, IIF(Session("COM_ADMIN")<>"","WHERE (1=1) ","WHERE email_id "), vbTextCompare))													 
											%>
											<% CALL ExportContattiInRubrica(sql, "log_cnt_id", "", "") %>
										</td>
									</tr>
								</table>
							</td>
						</tr>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco comunicazioni
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovate n&ordm; <%= Pager.recordcount %> comunicazioni in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo 
							'recupera errori di spedizione
							sql = GetQuery_LogContatti(conn, rs("email_id"), true)
							rse.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
							'recupera rubriche
							sql =  GetQuery_LogRubriche(conn, rs("email_id"))
							rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
							'recupera destinatari
							sql = GetQuery_LogContatti(conn, rs("email_id"), false)
							'response.write sql
							'response.end
							rsd.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
							%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td rowspan="<%= 3 + IIF(not rse.eof, 2, 0) + IIF(not rsd.eof, 1, 0) + IIF(not rsr.eof, 1, 0) + IIF(CString(rs("email_docs")) <> "", 1, 0) %>" align="center" valign="top" width="18">
												<% CALL Comunicazioni_Icona(rs("email_tipi_messaggi_id"))
												if rs("email_archiviata") then %>
													<img src="../grafica/archiviata.gif" border="0" alt="Comunicazione archiviata">
												<% end if %>
												<%if cIntero(rs("email_newsletter_tipo_id")) > 0 then %>
													<img src="../grafica/i.p.new.gif" border="0" alt="Newsletter">
												<% end if %>
											</td>
											<% if cBoolean(rs("email_isBozza"), false) then %>
												<td colspan="2" class="header sticker">
											<% elseif rse.recordcount>0 then %>
												<td colspan="2" class="header alert">
											<% else %>
												<td colspan="2" class="header ok">
											<% end if %>
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<% 	if cBoolean(rs("email_isBozza"), false) then %>
																<a class="button" href="ComunicazioniNew_Wizard_2.asp?ID=<%= rs("email_id") %>" title="riprende la procedura di spedizione.">
																	ESEGUI INVIO
																</a>
															<% else %>
																<a class="button" href="ComunicazioniView.asp?ID=<%= rs("email_id") %>" title="visualizza il messaggio inviato">
																	VISUALIZZA
																</a>
																<% 	if not cBoolean(rs("email_isBozza"), false) then %>
																	&nbsp;
																	<a class="button" href="ComunicazioniNew_Wizard_2.asp?Inoltra_id=<%= rs("email_id") %>" title="noltra il messaggio">
																		INOLTRA
																	</a>
																<% end if
															end if %>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('EMAIL','<%= rs("email_id") %>');">
																CANCELLA
															</a>
														</td>
													</tr>
												</table>
												<% if  rs("email_tipi_messaggi_id") <> MSG_SMS then%>
													<%= rs("email_object") %>
												<% else %>
													<%= Sintesi(rs("email_text"), 35, "...") %>
												<% end if %>
											</td>
										</tr>
										<tr>
											<td class="label">inviat<%= IIF(rs("email_tipi_messaggi_id") = MSG_EMAIL, "a", "o") %>:</td>
											<% 	if cBoolean(rs("email_isBozza"), false) then %>
												<td class="content_disabled">non inviata</td>
											<% else %>
												<td class="content">
													<%= DateTimeIta(rs("email_data")) %>
												</td>
											<% end if %>
										</tr>
										<tr>
											<th class="l2" colspan="2">destinatari <%= IIF(cBoolean(rs("email_isBozza"), false), "selezionati", "") %>:</td>
										</tr>
										<% 
										if rsr.recordcount > 0 then%>
											<tr>
												<td class="label">rubriche</td>
												<td>
													<% CALL Write_Log_Rubriche(rsr, 4) %>
												</td>
											</tr>
										<% end if
										
										if rse.recordcount > 0 AND not CBoolean(rs("email_isBozza"), false) then %>
											<tr>
												<td class="label" rowspan="2">errori di invio</td>
												<td class="content alert">
													<span style="float:right;">
														<a class="button_L2" title="Ritenta l'inivio dell'email ai destinatari ai quali non &egrave; stata ancora recapitata."
														   href="ComunicazioniNew_Wizard_2.asp?RitentaErrati_id=<%= rs("email_id") %>">
														   RITENTA INVIO <%= IIF(rse.recordcount=1, "AL CONTATTO", "AI " & rse.recordcount & " CONTATTI") %>
														</a>
													</span>
													<span class="smaller"> n&ordm; <%= rse.recordcount %> errori nella spedizione</span>
												</td>
											</tr>
											<tr>
												<td>
													<% CALL Write_Log_Contatti(rse, false, 4) %>
												</td>
											</tr>	
										<% end if
										
										if rsd.recordcount > 0 then %>
											<tr>
												<td class="label">contatti</td>
												<td>
													<% CALL Write_Log_Contatti(rsd, true, 4) %>
												</td>
											</tr>
										<% end if
										if CString(rs("email_docs")) <> "" then %>
										<tr>
												<td class="label">allegati:</td>
												<td class="content">
													<% CALL Write_Allegati(rs("email_docs"), rs("email_id")) %>
												</td>
											</tr>
										<% 	end if %>
									</table>
								</td>
							</tr>
											
							<%rse.close
							rsr.close
							rsd.close
							rs.moveNext
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
	</form>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsd = nothing
set rse = nothing
set rsr = nothing
set conn = nothing

%>
