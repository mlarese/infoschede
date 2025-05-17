<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
dim conn, rs, rsc, sql, Pager
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_ELENCO"), "IDElencoIndirizzi", "Pratiche.asp")
end if

'imposta filtro sul contatto
if request.QueryString("all")<>"" then		'se richiesta arriva da menu: visualizza tutte le pratiche
	Session("PRA_CONTATTO_ID") = ""
	Session("PRA_CONTATTO_NOME") = ""
	Session("PRA_PREFIX") = ""
	response.redirect "Pratiche.asp"
elseif request.Querystring("ID")<>"" then	'se richiesta arriva da contatti: visualizza tutte le pratiche del contatto
	Session("PRA_CONTATTO_ID") = cIntero(request.Querystring("ID"))
	Session("PRA_CONTATTO_NOME") = ""
	Session("PRA_PREFIX") = "CNT_"
	response.redirect "Pratiche.asp"
end if

%>
<%'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
if Session("PRA_CONTATTO_ID")<>"" then
	Titolo_sezione = "Anagrafica contatti - pratiche"
	HREF = "Contatti.asp;ContattiMod.asp?ID=" & Session("PRA_CONTATTO_ID") & ";ContattiRecapiti.asp?ID=" & Session("PRA_CONTATTO_ID") & ";"
	Action = "INDIETRO;SCHEDA CONTATTO;RECAPITI CONTATTO;"
else
	Titolo_sezione = "Pratiche - elenco"
	HREF = ""
	Action = ""
end if
HREF = HREF & "PraticaNew.asp"
Action = Action & "NUOVA PRATICA"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<%
dim Prefix, docC, attC
Prefix = Session("PRA_PREFIX")

Session("PageNavigator_VarPrefix") = Prefix
set Pager = new PageNavigator
'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	if request("tutti")<>"" then
		Session(Prefix & "pra_stato") = ""
		Session(Prefix & "pra_contatto") = ""
		Session(Prefix & "pra_codice") = ""
		Session(Prefix & "pra_nome") = ""
		Session(Prefix & "pra_data_creazione_from") = ""
		Session(Prefix & "pra_data_creazione_to") = ""
		Session(Prefix & "pra_data_UM_from") = ""
		Session(Prefix & "pra_data_UM_to") = ""
		Session(Prefix & "pra_full_text") = ""
	elseif request("cerca")<>"" then
		Session(Prefix & "pra_stato") = request("search_stato")
		Session(Prefix & "pra_contatto") = request("search_contatto")
		Session(Prefix & "pra_codice") = request("search_codice")
		Session(Prefix & "pra_nome") = request("search_nome")
		Session(Prefix & "pra_data_creazione_from") = request("search_data_creazione_from")
		Session(Prefix & "pra_data_creazione_to") = request("search_data_creazione_to")
		Session(Prefix & "pra_data_UM_from") = request("search_data_um_from")
		Session(Prefix & "pra_data_UM_to") = request("search_data_um_to")
		Session(Prefix & "pra_full_text") = request("search_full_text")
	end if
end if

'imposta criteri per ricerca semplice
sql = " SELECT tb_pratiche.*, tb_admin.admin_nome, tb_admin.admin_cognome, " &_ 
	  " tb_indirizzario.isSocieta, tb_indirizzario.NomeElencoIndirizzi, tb_indirizzario.CognomeElencoIndirizzi, " & _
	  " tb_indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_indirizzario.IDElencoIndirizzi " & _
	  " FROM (tb_pratiche INNER JOIN tb_admin ON tb_pratiche.pra_creatore_id=tb_admin.id_admin) " & _
	  " INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id=tb_indirizzario.IDElencoIndirizzi " & _
	  " WHERE (pra_creatore_id = "& Session("ID_ADMIN") &" OR "& AL_query(conn, AL_PRATICHE) &") "
	  
'filtra su id contatto (mostra solo pratiche del contatto)
if Session("PRA_CONTATTO_ID")<>"" then
	sql = sql & " AND pra_cliente_id=" & Session("PRA_CONTATTO_ID")
	if Session("PRA_CONTATTO_NOME") = "" then
		rs.open "SELECT isSocieta, NomeElencoIndirizzi, CognomeElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi " & _
			    "FROM tb_indirizzario WHERE IdElencoIndirizzi=" & Session("PRA_CONTATTO_ID"), _
				conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		Session("PRA_CONTATTO_NOME") = ContactFullName(rs)
		rs.close
	end if
end if
	
'filtra per stato della pratica
if Session(Prefix & "pra_stato")<>"" then
	sql = sql & " AND ( "
	if instr(1, Session(Prefix & "pra_stato"), "A", vbTextCompare)>0  then		'aperte
		sql = sql & " (NOT " & SQL_IsTrue(conn, "pra_archiviata") & ") OR "
	end if
	if instr(1, Session(Prefix & "pra_stato"), "C", vbTextCompare)>0  then		'chiuse
		sql = sql & " " & SQL_IsTrue(conn, "pra_archiviata") & " OR "
	end if
	sql = left(sql, len(sql)-3) & " )"
end if
	
'filtra per nominativo del contatto
if Session(Prefix & "pra_contatto")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch_Contatto_Nominativo(conn, Session(Prefix & "pra_contatto"))
end if

'filtra per codice pratica
if Session(Prefix & "pra_codice")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch(Session(Prefix & "pra_codice"), "pra_codice")
end if

'filtra per nome pratica
if Session(Prefix & "pra_nome")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch(Session(Prefix & "pra_nome"), "pra_nome")
end if

'filtra per data di creazione
if isDate(Session(Prefix & "pra_data_creazione_from")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "pra_dataI", adCompareGreaterThan, Session(Prefix & "pra_data_creazione_from")) & " "
end if
if isDate(Session(Prefix & "pra_data_creazione_to")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "pra_dataI", adCompareLessThan, Session(Prefix & "pra_data_creazione_to")) & " "
end if

'filtra per data dell'ultima attivita
if isDate(Session(Prefix & "pra_data_UM_from")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "pra_dataUM", adCompareGreaterThan, Session(Prefix & "pra_data_UM_from")) & " "
end if
if isDate(Session(Prefix & "pra_data_UM_to")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "pra_dataUM", adCompareLessThan, Session(Prefix & "pra_data_UM_to")) & " "
end if

'filtro full text
if Session(Prefix & "pra_full_text")<>"" then
	sql = sql &" AND " + SQL_FullTextSearch(Session(Prefix & "pra_full_text"), "pra_codice;pra_nome;pra_note")
end if
	
	
	if DB_Type(conn) = DB_SQL then
		sql = sql & " ORDER BY pra_archiviata, pra_dataUM DESC"
	else
		sql = sql & " ORDER BY pra_archiviata DESC, pra_dataUM DESC"
	end if

Session("SQL_PRATICHE") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<form action="Pratiche.asp" method="post" id="ricerca" name="ricerca">
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTE" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "pra_stato") %>>STATO PRATICA</th></tr>
								<tr>
									<td class="content pratiche">
										<input type="Checkbox" name="search_stato" class="checkbox" value="A" <%= IIF(instr(1, Session(Prefix & "pra_stato"), "A", vbTextCompare)>0, " checked", "") %>>
										<strong>aperte</strong> 
									</td>
								</tr>
								<tr>
									<td class="content">
										<input type="Checkbox" name="search_stato" class="checkbox" value="C" <%= IIF(instr(1, Session(Prefix & "pra_stato"), "C", vbTextCompare)>0, " checked", "") %>>
										archiviate 
									</td>
								</tr>
								<% if Session("PRA_CONTATTO_ID")="" then %>
									<tr><th <%= Search_Bg(Prefix & "pra_contatto") %>>NOMINATIVO CONTATTO</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_contatto" value="<%= replace(session(Prefix & "pra_contatto"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
								<% end if %>
								<tr><th <%= Search_Bg(Prefix & "pra_codice") %>>CODICE</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_codice" value="<%= replace(session(Prefix & "pra_codice"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "pra_nome") %>>NOME PRATICA</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= replace(session(Prefix & "pra_nome"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "pra_data_creazione_from;" & Prefix & "pra_data_creazione_to") %>>DATA CREAZIONE</td></tr>
								<tr><td class="label">a partire dal:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_creazione_from", Session(Prefix & "pra_data_creazione_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td class="label">fino al:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_creazione_to", Session(Prefix & "pra_data_creazione_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "pra_data_UM_from;" & Prefix & "pra_data_UM_to") %>>DATA ULTIMA ATTIVITA'</td></tr>
								<tr><td class="label">a partire dal:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_UM_from", Session(Prefix & "pra_data_UM_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td class="label">fino al:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_UM_to", Session(Prefix & "pra_data_UM_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "pra_full_text") %>>FULL-TEXT</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_full_text" value="<%= replace(session(Prefix & "pra_full_text"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTE" style="width: 49%;">
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
						<% if Session("PRA_CONTATTO_ID")<>"" then 
							'scorre contatti%>
							<table border="0" cellspacing="0" cellpadding="0" align="right">
								<tr>
									<td style="font-size: 1px; padding-right:1px;" nowrap>
										<a class="button" href="?ID=<%= Session("PRA_CONTATTO_ID") %>&goto=PREVIOUS" title="pratiche del contatto precedente">
											&lt;&lt; PRECEDENTE
										</a>
										&nbsp;
										<a class="button" href="?ID=<%= Session("PRA_CONTATTO_ID") %>&goto=NEXT" title="pratiche del contatto successivo">
											SUCCESSIVO &gt;&gt;
										</a>
									</td>
								</tr>
							</table>
							Elenco pratiche del contatto &quot;<%= Session("PRA_CONTATTO_NOME") %>&quot;
						<% else %>
							Elenco pratiche
						<% end if %>
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovate n&ordm; <%= Pager.recordcount %> pratiche in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header<%= IIF(rs("pra_archiviata"), "", " pratiche") %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<% If Session("COM_ADMIN") <> "" OR Session("COM_POWER") <> "" OR _
																  Session("ID_ADMIN") = rs("pra_creatore_id") then 
																sql = "SELECT COUNT(*) FROM tb_documenti WHERE doc_pratica_id=" & rs("pra_id")
																docC = cInteger(GetValueList(conn, rsc, sql))
																sql = "SELECT COUNT(*) FROM tb_attivita WHERE NOT "& SQL_IsTrue(conn, "att_conclusa") & " AND att_pratica_id=" & rs("pra_id")
																attC = cInteger(GetValueList(conn, rsc, sql))%>
																<a class="button" href="PraticaMod.asp?ID=<%= rs("pra_id") %>" title="modifica dati della pratica">
																	MODIFICA
																</a>
																&nbsp;
																<a class="button" href="Attivita.asp?PRA_ID=<%= rs("pra_id") %>" title="gestione attivit&agrave; della pratica">
																	ATTIVIT&Agrave;
																</a>
																&nbsp;
																<a class="button" href="Documenti.asp?PRA_ID=<%= rs("pra_id") %>" title="gestione documenti della pratica">
																	DOCUMENTI
																</a>
																&nbsp;
																<% if docC = 0 AND attC = 0 OR Session("COM_ADMIN") <> "" then %>
																	<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('PRATICHE','<%= rs("pra_id") %>');">
																		CANCELLA
																	</a>
																<% Else  %>
																	<a class="button_disabled" title="impossibile cancellare la pratica: documenti o attivit&agrave; associate!">
																		CANCELLA
																	</a>
																<%End If
															Else %>
																<a class="button" href="PraticaMod.asp?ID=<%= rs("pra_id") %>" title="modifica dati della pratica">
																	VISUALIZZA
																</a>
																&nbsp;
																<a class="button" href="Attivita.asp?PRA_ID=<%= rs("pra_id") %>" title="gestione attivit&agrave; della pratica">
																	ATTIVIT&Agrave;
																</a>
																&nbsp;
																<a class="button" href="Documenti.asp?PRA_ID=<%= rs("pra_id") %>" title="gestione documenti della pratica">
																	DOCUMENTI
																</a>
																&nbsp;
																<a class="button_disabled" title="impossibile cancellare la pratica!">
																	CANCELLA
																</a>
															<% 	End If %>
														</td>
													</tr>
												</table>
												<% If rs("pra_archiviata") then %>
													<span style="font-weight:normal;">
														<%= rs("pra_nome") %>
													</span>
												<% Else  %>
													<%= rs("pra_nome") %>
												<% End If %>
											</td>
										</tr>
										<% if Session("PRA_CONTATTO_ID")="" then %>
											<tr>
												<td class="label">contatto:</td>
												<td class="content" colspan="3">
													<% ContactLinkedName(rs) %>
												</td>
											</tr>
										<% end if %>
										<tr>
											<td class="label" style="width:20%;">codice:</td>
											<td class="content_b"><%= rs("pra_codice") %></td>
											<td class="label" style="width:20%;">Progressivo:</td>
											<td class="content"><%= rs("pra_ID") %></td>
										</tr>
										<tr>
											<td class="label">creatore:</td>
											<td class="content"><%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %></td>
											<td class="label" style="width:18%;">data creazione:</td>
											<td class="content" style="width:20%;"><%= DateTimeITA(rs("pra_dataI")) %></td>
										</tr>
										<tr>
											<td class="label">data ultima attivit&agrave;:</td>
											<td class="content" colspan="3"><%= DateTimeITA(rs("pra_dataUM")) %></td>
										</tr>
										
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
						<tr><td class="noRecords" colspan="2">Nessun record trovato</th></tr>
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
