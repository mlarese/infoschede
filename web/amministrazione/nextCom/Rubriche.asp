<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->

<%
'controllo accesso
if Session("COM_ADMIN")="" AND Session("COM_POWER")="" then
	response.redirect "Contatti.asp"
end if

dim conn, rs, rsg, sql, Pager

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsg = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("rub_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("rub_")
	end if
end if

'recupera rubriche visibili all'utente
dim rubriche_visibili
rubriche_visibili = GetList_Rubriche(conn, rs)

sql = ""
'filtra per nome della rubrica
if session("rub_nome")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(session("rub_nome"), FieldLanguageList("nome_pubblico_rubrica_") + ";" + SQL_concatFields(conn, "nome_rubrica")) & ")"
end if

'filtra per categoria
if session("rub_gruppo")<>"" then
	sql = sql & " AND id_rubrica IN (SELECT id_DellaRubrica FROM tb_rel_gruppirubriche WHERE id_gruppo_assegnato=" & session("rub_gruppo") & ")"
end if

sql = "SELECT * FROM tb_rubriche " &_
		" WHERE id_rubrica IN (" & rubriche_visibili & ") " &_
		sql & _
		" ORDER BY nome_rubrica"
Session("SQL_RUBRICHE_ELENCO") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>
<%'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Rubriche - elenco"
'Indirizzo pagina per link su sezione 
		HREF = "RubricheNew.asp"
'Azione sul link: {BACK | NEW}
	Action = "NUOVA RUBRICA"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
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
								<tr><th <%= Search_Bg("rub_nome") %>>NOME</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= Server.HTMLEncode(session("rub_nome")) %>" style="width:100%;">
									</td>
								</tr>
								<% if Cinteger(Application("NextCom_DefaultWorkGroup"))=0 then %>
									<tr><th <%= Search_Bg("rub_gruppo") %>>GRUPPI DI LAVORO ABILITATI</td></tr>
									<tr>
										<td class="content">
											<%sql = "SELECT * FROM tb_gruppi ORDER BY nome_gruppo"
											CALL dropDown(conn, sql, "id_gruppo", "nome_gruppo", "search_gruppo", session("rub_gruppo"), false, " style=""width:100%;""", LINGUA_ITALIANO)%>
										</td>
									</tr>
								<% end if %>
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
					<caption>Elenco rubriche</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="RubricheMod.asp?ID=<%= rs("id_rubrica") %>">
																MODIFICA
															</a>
															&nbsp;
															<% if not rs("locked_rubrica") AND not rs("rubrica_esterna") then %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('RUBRICHE','<%= rs("id_rubrica") %>');" >
																	CANCELLA
																</a>
															<% else %>
																<a class="button_disabled" href="javascript:void(0);" title="rubrica non cancellabile perch&egrave; utilizzata da alcuni automatismi dell'applicazione">
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												<%=rs("nome_rubrica")%>
											</td>
										</tr>
										<tr>	
											<td class="label">n&ordm; contatti</td>
											<% sql = "SELECT COUNT(*) FROM rel_rub_ind WHERE id_rubrica=" & rs("id_rubrica") %>
											<td class="content" style="width:25%;"><%= cInteger(GetValueList(conn, rsg, sql)) %></td>
											<td class="label">numero</td>
											<td class="content"><%= rs("id_rubrica") %></td>
										</tr>
										<% if Cinteger(Application("NextCom_DefaultWorkGroup"))=0 then %>	
											<tr>
												<td class="label" style="width:25%;">gruppi di lavoro abilitati</td>
												<% sql = "SELECT nome_Gruppo FROM tb_gruppi INNER JOIN tb_rel_gruppirubriche " &_
														 " ON tb_gruppi.id_Gruppo = tb_rel_gruppirubriche.id_Gruppo_assegnato " &_
														 " WHERE id_dellaRubrica=" & rs("id_rubrica") & " ORDER BY tb_gruppi.nome_Gruppo "%>
												<td class="content" colspan="3"><%= GetValueList(conn, rsg, sql) %></td>
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
set rsg = nothing
set conn = nothing%>
