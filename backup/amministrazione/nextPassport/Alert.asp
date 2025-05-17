<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione alert - elenco"
if session("PASS_ADMIN") <> "" then
	dicitura.puls_new = "NUOVO ALERT"
	dicitura.link_new = "AlertNew.asp"
end if
dicitura.scrivi_con_sottosez()


dim conn, rs, sql, Pager, rsa, headerCss
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("ale_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("ale_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")

'filtra per nome
if Session("ale_nome")<>"" then
    sql = sql &" AND "& SQL_MultiLanguage(SQL_FullTextSearch(Session("ale_nome"), "sev_nome_<LINGUA>"), "OR")
end if

'filtra per applicazione di accesso
if CIntero(Session("ale_applicazione")) > 0 then
	sql = sql &" AND sev_sito_id = " & Session("ale_applicazione")
end if

'filtra per abilitazione
if Session("ale_abilitato") <> "" AND Session("ale_disabilitato") = "" then
	sql = sql &" AND "& SQL_IsTrue(conn, "sev_abilitato")
elseif Session("ale_disabilitato") <> "" AND Session("ale_abilitato") = "" then
	sql = sql &" AND NOT "& SQL_IsTrue(conn, "sev_abilitato")
end if

sql = " SELECT * FROM tb_siti_eventi e"& _
	  " LEFT JOIN tb_siti s ON e.sev_sito_id = s.id_sito"& _
	  " WHERE (1=1)"& sql & _
	  " ORDER BY sev_nome_it "
Session("SQL_ALERT") = sql
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
										<input type="submit" class="button" name="cerca" value="CERCA" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("ale_nome") %>>NOME</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= session("ale_nome")%>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("ale_abilitato;ale_disabilitato") %>>STATO ABILITAZIONE</td></tr>
								<tr>
									<td class="content_b">
										<input type="checkbox" class="checkbox" name="search_abilitato" value="1" <%= chk(Session("ale_abilitato")<>"") %>>
										abilitato
									</td>
								</tr>
								<tr>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_disabilitato" value="1" <%= chk(Session("ale_disabilitato")<>"") %>>
										non abilitato
									</td>
								</tr>
								<tr><th <%= Search_Bg("ale_applicazione") %>>Applicazione</td></tr>
								<tr>
									<td class="content">
									<%	sql = "SELECT * FROM tb_siti WHERE " & SQL_IsTrue(conn, "sito_Amministrazione") & " ORDER BY sito_nome"
										CALL dropDown(conn, sql, "id_sito", "sito_nome", "search_applicazione", Session("ale_applicazione"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
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
					<caption>Elenco alert</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> alert in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo
							headerCss = ""
							if NOT rs("sev_abilitato") then
								headerCss = "_disabled"
							end if %>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header<%= headerCss %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
                                                            <a class="button" href="AlertMod.asp?ID=<%= rs("sev_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ALERT','<%= rs("sev_id") %>');" >
																CANCELLA
															</a>
														</td>
													</tr>
												</table>
												<%= rs("sev_nome_it") %>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:22%;">codice:</td>
											<td class="content" colspan="3"><%= rs("sev_codice") %></td>
										</tr>
										<tr>
											<td class="label">applicazione:</td>
											<td class="content" colspan="3"><%= rs("sito_nome") %></td>
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