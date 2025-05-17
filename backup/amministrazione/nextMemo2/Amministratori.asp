<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata  
dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "LOG ACCESSI"
dicitura.links(1) = "AmministratoriAccessi.asp"
dicitura.sezione = "Gestione utenti area amministrativa NextMemo 2.0 - elenco"
dicitura.puls_new = "NUOVO UTENTE"
dicitura.link_new = "AmministratoriNew.asp"
dicitura.scrivi_con_sottosez() 


dim conn, rs, sql, Pager, rsa, headerCss
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("a_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("a_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")

'filtra per nome e/o cognome dell'utente
if Session("a_nome")<>"" then
    sql = sql & SQL_FullTextSearch(Session("a_nome"), "admin_nome;admin_cognome;" + SQL_concatFields(conn, "admin_cognome;admin_nome"))
end if

'filtra per login
if Session("a_login")<>"" then
	if sql <>"" then sql = sql & " AND "
    sql = sql & SQL_FullTextSearch(Session("a_login"), "admin_login")
end if

'filtra per email
if Session("a_email")<>"" then
	if sql <>"" then sql = sql & " AND "
    sql = sql & SQL_FullTextSearch(Session("a_email"), "admin_email")
end if


'filtra per stato abilitazioni
if Session("a_abilitato")<> "" OR Session("a_accessoscaduto")<> "" then
	if not ( Session("a_abilitato")<> ""  AND Session("a_accessoscaduto")<> "") then
		if sql <>"" then sql = sql & " AND "
		sql = sql + " ( "
	
		if Session("a_abilitato")<> "" then
			sql = sql + " ( id_admin IN (SELECT admin_id FROM rel_admin_sito) AND ( " & SQL_IfIsNull(conn, "admin_scadenza", Sql_Now(conn) & " + 1") & " >= " & SQL_date(conn, Date()) & ") ) OR "
		end if
		if Session("a_accessoscaduto")<>"" then
			sql = sql + " ( " &  SQL_IfIsNull(conn, "admin_scadenza", Sql_Now(conn) & " + 1") & " < " & SQL_Now(conn) & ") OR "
		end if
		
		sql = left(sql, (len(sql)-3)) + " ) "
	end if
end if

'ricerca profilo
if Session("a_profilo")<>"" then
	sql = sql & "id_admin IN (SELECT rpa_admin_id FROM mrel_profili_admin WHERE rpa_profilo_id = " & Session("a_profilo") & ")"
end if


if sql<>"" then sql = " AND " & sql
sql = " SELECT (SELECT COUNT(*) FROM rel_admin_sito WHERE admin_id = tb_admin.id_admin) AS N_Abilitazioni, * " + _
	  " FROM tb_admin WHERE ID_admin IN (SELECT DISTINCT admin_id FROM rel_admin_sito WHERE rel_as_permesso > 1 AND " + _
	  " 														sito_id = " & NEXTMEMO2 & ") " & _
	  " 				AND ID_admin NOT IN (SELECT DISTINCT admin_id FROM rel_admin_sito WHERE rel_as_permesso = 1 AND " + _
	  " 														sito_id = " & NEXTMEMO2 & ")" & sql & " ORDER BY admin_cognome, admin_nome"

Session("SQL_AMMINISTRATORI") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<form action="Amministratori.asp" method="post" id="ricerca" name="ricerca">
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer">
										<input type="submit" class="button" name="cerca" value="CERCA" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("a_nome") %>>COGNOME E/O NOME</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= session("a_nome")%>" style="width:100%;">
									</td>
								</tr>
								
								<tr><th <%= Search_Bg("a_login") %>>Login</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_login" value="<%= session("a_login")%>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("a_abilitato;a_accessoscaduto") %>>STATO ABILITAZIONE</td></tr>
								<tr>
									<td class="content_b">
										<input type="checkbox" class="checkbox" name="search_abilitato" value="1" <%= chk(Session("a_abilitato")<>"") %>>
										abilitato
									</td>
								</tr>
								<tr>
									<td class="content warning">
										<input type="checkbox" class="checkbox" name="search_accessoscaduto" value="1" <%= chk(Session("a_accessoscaduto")<>"") %>>
										accesso scaduto
									</td>
								</tr>
								<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) then %>
									<% 
									sql = "SELECT * FROM mtb_profili ORDER BY pro_nome_it"
									if GetValueList(conn, NULL, sql) <>"" then %>
										<tr><th colspan="2" <%= Search_Bg("a_profilo") %>>PROFILO</th></tr>
										<tr>
											<td class="content" colspan="2">
												<% CALL dropDown(conn, sql, "pro_id", "pro_nome_it", "search_profilo", session("a_profilo"), false, "style=""width: 100%;""", Session("LINGUA")) %>
											</td>
										</tr>
									<% end if %>
								<% end if %>
								<tr><th <%= Search_Bg("a_email") %>>Email</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_email" value="<%= session("a_email")%>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<td class="footer">
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
					<caption>Elenco utenti area amministrativa</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> utenti  in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo
							headerCss = ""
							if not IsNull(rs("admin_scadenza")) then
								if rs("admin_Scadenza") < Date() then
									headerCss = " warning"
								end if
							end if
							if headerCss = "" and cIntero(rs("N_Abilitazioni"))=0 then
								headerCss = " alert"
							end if %>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header<%= headerCss %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="AmministratoriMod.asp?ID=<%= rs("id_admin") %>">
																MODIFICA
															</a>
															&nbsp;
															<% dim altre_applicazioni
															   sql = " SELECT COUNT(*) FROM rel_admin_sito WHERE admin_id = " & rs("id_admin") & " AND sito_id <> " & NEXTMEMO2
															   altre_applicazioni = GetValueList(conn, NULL, sql)
															%>
															<% if cIntero(altre_applicazioni) > 0 then %>  
																<a class="button_disabled" href="javascript:void(0);" title="Utente non cancellabile perchè collegato ad altri applicativi.">
																	CANCELLA
																</a>
															<% else %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('AMMINISTRATORI','<%= rs("id_admin") %>');" >
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												<%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:22%;">login:</td>
											<td class="content"><%= rs("admin_login") %></td>
											<td class="label_no_width" style="width:20%;">scadenza accesso:</td>
											<td class="content" style="width:14%;"><%= DateIta(rs("admin_Scadenza")) %></td>
										</tr>
										<tr>
											<td class="label">ultimo accesso:</td>
											<td class="content">
												<% sql = "SELECT TOP 1 log_data FROM log_admin WHERE log_admin_id=" & rs("id_admin") & " ORDER BY log_data DESC " 
												rsa.open sql, conn, adOpenStatic, adLockOptimistic, adcmdText
												if rsa.eof then%>
													accesso non ancora effettuato
												<% else %>
													<%= DateTimeIta(rsa("log_data")) %>
												<% end if
												rsa.close %>
											</td>
											<td class="label_no_width" style="text-align:right;">
												<%
												sql = "SELECT COUNT(*) FROM log_admin WHERE log_admin_id = " & rs("id_admin") 
												if cIntero(GetValueList(conn, NULL, sql)) > 0 then %>
													<a class="button_L2" href="AmministratoriAccessi.asp?ID=<%= rs("id_admin") %>" title="Accessi dell'utente">
														ACCESSI
													</a>
												<% else %>
													<a class="button_L2 button_disabled" style="padding-bottom:0px;" href="javascript:void(0);" title="Nessun accesso effettuato dall'utente.">
														ACCESSI
													</a>
												<% end if %>
											</td>
											<td class="label_no_width">
												<% CALL WriteCampoCerca("Documenti.asp", "admin_id", rs("id_admin"), "DOCUMENTI", "button_L2") %>
											</td>
										</tr>
										<tr>
											<td class="label">email in uscita:</td>
											<td class="content" colspan="3"><a href="mailto:<%= rs("admin_email") %>"><%= rs("admin_email") %></a></td>
										</tr>
										<tr>
											<td class="label">tipo di abilitazione:</td>
											<td class="content" colspan="3">
                                            <% dim permessi, permesso
											   sql = "SELECT rel_as_permesso FROM rel_admin_sito WHERE admin_id = " & rs("id_admin") & " AND sito_id = " & NEXTMEMO2
											   permessi = GetValueList(conn, NULL, sql)
											   permessi = replace(permessi, " ", "")
											   permessi = split(permessi, ",")
											   for each permesso in permessi
													sql = " SELECT DISTINCT sito_p" & permesso & " FROM tb_siti INNER JOIN rel_admin_sito " + _
														  " ON tb_siti.id_sito = rel_admin_sito.sito_id WHERE admin_id = " & rs("id_admin") & " AND sito_id = " & NEXTMEMO2
													response.write cString(GetValueList(conn, NULL, sql)) & "; "
											   next
                                            %>
                                            </td>
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
set rsa = nothing
set conn = nothing%>