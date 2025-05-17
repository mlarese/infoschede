<%@ Language=VBScript CODEPAGE=65001
%>
<% Option Explicit 
%>
<% response.charset = "UTF-8" 
%>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim i, conn, rs, rsA, sql, Pager

set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsA = Server.CreateObject("ADODB.RecordSet")

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Log di sistema"
dicitura.puls_new = "INDIETRO A STRUMENTI"
dicitura.link_new = "Strumenti.asp"
dicitura.scrivi_con_sottosez()

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("log_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("log_")
	end if
end if

sql = ""

'filtra per codice operazione
if session("log_codice")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch(session("log_codice"), "log_codice")
end if

'filtra per nome tabella
if session("log_table_nome")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch(session("log_table_nome"), "log_table_nome")
end if

'filtra per categoria
if session("log_user_login")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch(session("ut_login"), "ut_login")
end if

'filtra per data di invio
if isDate(Session("log_data_from")) then
	sql = sql & " AND log_data >=" & SQL_date(conn, Session("log_data_from"))
end if
if isDate(Session("log_data_to")) then
	sql = sql & " AND log_data <=" & SQL_date(conn, Session("log_data_to"))
end if

if sql<>"" then	
	sql=" where 1=1 " & sql
end if

sql = "SELECT * FROM (log_framework LEFT JOIN tb_Utenti ON log_framework.log_user_id = tb_Utenti.ut_ID ) " +_
	  " LEFT JOIN tb_admin ON log_framework.log_admin_id= tb_admin.ID_admin " & sql & " ORDER BY log_id DESC "
CALL Pager.OpenSmartRecordset(conn, rs, sql, 25)
%>
<div id="content">
	<!-- BLOCCO DI RICERCA -->
	<form action="" method="post" id="ricerca" name="ricerca">
			<table width="15%" border="0" cellspacing="0" cellpadding="0" align="left">			
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
							<tr><th <%= Search_Bg("log_codice") %>>CODICE OPERAZIONE</td></tr>
							<tr>
								<td class="content">
									<input type="text" name="search_codice" value="<%= Server.HTMLEncode(session("log_codice")) %>" style="width:100%;">
								</td>
							</tr>						
							<tr><th <%= Search_Bg("log_table_nome") %>>NOME TABELLA</td></tr>
							<tr>
								<td class="content">
									<input type="text" name="search_table_nome" value="<%= Server.HTMLEncode(session("log_table_nome")) %>" style="width:100%;">
								</td>
							</tr>						
							<tr><th <%= Search_Bg("log_user_login") %>>LOGIN UTENTE</td></tr>
							<tr>
								<td class="content">
									<input type="text" name="search_user_login" value="<%= Server.HTMLEncode(session("log_user_login")) %>" style="width:100%;">
								</td>
							</tr>
							<tr><th <%= Search_Bg("log_data_from;log_data_to") %>>DATA DEI LOG</td></tr>
								<tr><td class="label">a partire dal:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_from", Session("log_data_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td class="label">fino al:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_to", Session("log_data_to"), "", "/", true, true, LINGUA_ITALIANO) %>
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
<!-- BLOCCO RISULTATI -->
		<table cellspacing="1" width="85% cellpadding="0" class="tabella_madre">		
			<%if not rs.eof then
			%>
			<caption>
					Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine
			</caption>
			<tr>
			<tr>
				<th class="center" width="15%">DATA</th>
				<th class="center" width="15%">NOME TABELLA</th>
				<th class="center" width="15%">CODICE OPERAZIONE</th>
				<th class="center" width="35%">DESCRIZIONE</th>
				<th class="center" width="10%">UTENTE</th>
				<th class="center" width="10%">AMMINISTRATORE</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo 
			%>
				<tr>
					<td class="content_center"><%= DateTimeIta(rs("log_data")) %></td>
					<td class="content"><%= rs("log_table_nome") %></td>
					<td class="content"><%= rs("log_codice") %></td>
					<td class="content"><%= rs("log_descrizione") %></td>					
					<td class="content"><a href="UtentiMod.asp?ID=<%=rs("ut_nextcom_id")%>" target="_blank"><%= rs("ut_login") %></a></td>	
					<td class="content"><a href="AmministratoriMod.asp?ID=<%=rs("log_admin_id")%>" target="_blank"><%= rs("admin_nome") %></a></td>					
				</tr>
				<!--
				<%= rs("log_http_request")%>
				-->
				<%rs.movenext
			wend%>
			
			<tr>
				<td colspan="6" class="footer" style="text-align:left;">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")
					%>
				</td>
			</tr>
		<% else 
		%>
			<tr><td class="noRecords">Nessun log rilevato</th></tr>
		<% end if 
		%>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing
%>