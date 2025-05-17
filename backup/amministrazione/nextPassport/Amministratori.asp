<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata  
dicitura.iniz_sottosez(2)
dicitura.sottosezioni(1) = "ELENCO UTENTI"
dicitura.links(1) = "Amministratori.asp"
dicitura.sottosezioni(2) = "LOG ACCESSI"
dicitura.links(2) = "AmministratoriAccessi.asp"
dicitura.sezione = "Gestione utenti area amministrativa - elenco"
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

'filtra per applicazione di accesso
if Session("a_applicazione")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " id_admin IN (SELECT admin_id FROM rel_admin_sito WHERE sito_id=" & Session("a_applicazione") & " ) "
end if

'filtra per stato abilitazioni
if Session("a_abilitato")<> "" OR Session("a_senzaabilitazioni")<> "" OR Session("a_accessoscaduto")<> "" then
	if not ( Session("a_abilitato")<> "" AND Session("a_senzaabilitazioni")<> "" AND Session("a_accessoscaduto")<> "") then
		if sql <>"" then sql = sql & " AND "
		sql = sql + " ( "
	
		if Session("a_abilitato")<> "" then
			sql = sql + " ( id_admin IN (SELECT admin_id FROM rel_admin_sito) AND ( " & SQL_IfIsNull(conn, "admin_scadenza", Sql_Now(conn) & " + 1") & " >= " & SQL_date(conn, Date()) & ") ) OR "
		end if
		if Session("a_senzaabilitazioni")<> "" then
			sql = sql + " ( id_admin NOT IN (SELECT admin_id FROM rel_admin_sito) AND ( " & SQL_IfIsNull(conn, "admin_scadenza", Sql_Now(conn) & " + 1") & " >= " & SQL_date(conn, Date()) & ") ) OR "
		end if
		if Session("a_accessoscaduto")<>"" then
			sql = sql + " ( " &  SQL_IfIsNull(conn, "admin_scadenza", Sql_Now(conn) & " + 1") & " < " & SQL_Now(conn) & ") OR "
		end if
		
		sql = left(sql, (len(sql)-3)) + " ) "
	end if
end if

if sql<>"" then sql = " WHERE " & sql
sql = " SELECT (SELECT COUNT(*) FROM rel_admin_sito WHERE admin_id = tb_admin.id_admin) AS N_Abilitazioni, * " + _
	  " FROM tb_admin " & sql & " ORDER BY admin_cognome, admin_nome"

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
								<tr><th <%= Search_Bg("a_abilitato;a_senzaabilitazioni;a_accessoscaduto") %>>STATO ABILITAZIONE</td></tr>
								<tr>
									<td class="content_b">
										<input type="checkbox" class="checkbox" name="search_abilitato" value="1" <%= chk(Session("a_abilitato")<>"") %>>
										abilitato
									</td>
								</tr>
								<tr>
									<td class="content alert">
										<input type="checkbox" class="checkbox" name="search_senzaabilitazioni" value="1" <%= chk(Session("a_senzaabilitazioni")<>"") %>>
										senza abilitazioni
									</td>
								</tr>
								<tr>
									<td class="content warning">
										<input type="checkbox" class="checkbox" name="search_accessoscaduto" value="1" <%= chk(Session("a_accessoscaduto")<>"") %>>
										accesso scaduto
									</td>
								</tr>
								<tr><th <%= Search_Bg("a_email") %>>Email</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_email" value="<%= session("a_email")%>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("a_applicazione") %>>Applicazione</td></tr>
								<tr>
									<td class="content">
										<%sql = "SELECT * FROM tb_siti WHERE " & SQL_IsTrue(conn, "sito_Amministrazione") & " ORDER BY sito_nome"
										CALL dropDown(conn, sql, "id_sito", "sito_nome", "search_applicazione", Session("a_applicazione"), false, " style=""width:100%;""", LINGUA_ITALIANO)%>
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
															<%
															if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & COMMESSE)) <> "" then
																%>
																<a class="button" href="AmministratoriOrarioMod.asp?ID_ADMIN=<%= rs("id_admin") %>" title="Modifica orario di lavoro">
																	ORARIO
																</a>
																<%
															end if
															%>
                                                            <a class="button" href="AmministratoriAccessi.asp?ID=<%= rs("id_admin") %>" title="Accessi dell'utente">
																ACCESSI
															</a>
															&nbsp;
															<a class="button" href="AmministratoriMod.asp?ID=<%= rs("id_admin") %>">
																MODIFICA
															</a>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('AMMINISTRATORI','<%= rs("id_admin") %>');" >
																CANCELLA
															</a>
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
											<td class="content" style="width:30%;"><%= DateIta(rs("admin_Scadenza")) %></td>
										</tr>
										<tr>
											<td class="label" rowspan="2">Indirizzi email:</td>
											<td class="label_no_width">Email utente:</td>
											<td class="content" colspan="2"><a href="mailto:<%= rs("admin_email") %>"><%= rs("admin_email") %></a></td>
										</tr>
										<tr>
											<td class="label_no_width">Email per newsletter:</td>
											<td class="content" colspan="2"><a href="mailto:<%= rs("admin_email_newsletter") %>"><%= rs("admin_email_newsletter") %></a></td>
										</tr>
										<tr>
											<td class="label">ultimo accesso:</td>
											<td class="content" colspan="3">
												<% sql = "SELECT TOP 1 log_data FROM log_admin WHERE log_admin_id=" & rs("id_admin") & " ORDER BY log_data DESC " 
												rsa.open sql, conn, adOpenStatic, adLockOptimistic, adcmdText
												if rsa.eof then%>
													accesso non ancora effettuato
												<% else %>
													<%= DateTimeIta(rsa("log_data")) %>
												<% end if
												rsa.close %>
											</td>
										</tr>
										<tr>
											<td class="label">applicazioni abilitate:</td>
                                            <% sql = "SELECT * FROM tb_siti WHERE id_sito IN " &_
											         " (SELECT sito_id FROM rel_admin_sito WHERE admin_id=" & rs("id_admin") & ")"
                                            rsa.open sql, conn, adOpenStatic, adLockOptimistic, adcmdText
                                            if rsa.eof then %>
                                                <td class="content alert" colspan="3">Utente senza alcuna abilitazione</td>
                                            <% else %>
											    <td colspan="3">
                                                    <% if rsa.recordcount>2 then %>
														<span class="overflow">
													<% end if %>
														<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
                                                            <% while not rsa.eof %>
                                                                <tr>
                                                                    <% if cString(rsa("sito_prmEsterni_admin"))<>"" then %>
                                                                        <td class="content visibile" title="applicativo con permessi aggiuntivi">
                                                                    <% else %>
                                                                        <td class="content">
                                                                    <% end if %>
                                                                        <%= rsa("sito_nome") %>
                                                                    </td>
                                                                </tr>
                                                                <% rsa.movenext
                                                            wend %>
                                                        </table>
                                                    <% if rsa.recordcount>2 then %>
														</span>
													<% end if %>
											    </td>
                                            <% end if
                                            rsa.close%>
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