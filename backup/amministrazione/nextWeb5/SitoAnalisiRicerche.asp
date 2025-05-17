<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
Imposta_Proprieta_Sito("ID")
%>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_menu_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Analisi ricerche - Log completo"
dicitura.puls_new = "INDIETRO A SITI;LOG PAROLE CHIAVE"
dicitura.link_new = "Siti.asp;SitoAnalisiRicercheParoleChiave.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

dim Pager
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("logric_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("logric_")
	end if
end if

'filtra per data ricerche
if isDate(Session("logric_data_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "lor_data", adCompareGreaterThan, Session("logric_data_from"))
end if
if isDate(Session("logric_data_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "lor_data", adCompareLessThan, Session("logric_data_to"))
end if

'filtra per nome
if Session("logric_nome")<>"" then
    sql = sql & " AND " & SQL_FullTextSearch(Session("logric_nome"), "lor_ricerca")
end if

'filtra per risultati ottenuti
if session("logric_zero") = "1" then
	sql = sql &" AND lor_risultati_numero <> 0"
elseif session("logric_zero") = "2" then
	sql = sql &" AND lor_risultati_numero = 0"
end if


sql = "SELECT * FROM log_ricerche WHERE lor_web_id = "& session("AZ_ID") & sql & " ORDER BY lor_data DESC"
session("WEB_MENU_SQL") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 40)
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
										<td class="footer" nowrap>
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("logric_data_from;logric_data_to") %>>DATA RICERCA</td></tr>
									<tr><td class="label" colspan="2">a partire dal:</td></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_from", Session("logric_data_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><td class="label" colspan="2">fino al:</td></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_to", Session("logric_data_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th <%= Search_Bg("logric_nome") %>>CHIAVE DI RICERCA</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome" value="<%= TextEncode(session("logric_nome")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("logric_zero") %>>PER RISULTATO</th></tr>
									<tr>
										<td class="content" style="width:50%;">
											<input type="checkbox" class="checkbox" name="search_zero" value="1" <%= chk(instr(1, session("logric_zero"), "1", vbTextCompare)>0) %>>
											con risultato
										</td>
										<td class="content">
											<input type="checkbox" class="checkbox" name="search_zero" value="2" <%= chk(instr(1, Session("logric_zero"), "2", vbTextCompare)>0) %>>
											senza risultato
										</td>
									</tr>
									<tr>
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
		<caption>Log completo - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<% if not rs.eof then %>
			<tr>
				<th class="center" width="20%">DATA</th>
				<th>CHIAVE DI RICERCA</th>
				<th style="width:30%;">NUMERO DI RISULTATI</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
				while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
				<tr>
					<td class="content_center"><%= DateTimeIta(rs("lor_data")) %></td>
					<td class="content"><%= rs("lor_ricerca") %></td>
					<td class="content"><%= rs("lor_risultati_numero") %></td>
					
				</tr>
				<%rs.movenext
			wend%>
			<tr>
				<td class="footer" style="border-top:0px; text-align:left;" colspan="3">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
				</td>
			</tr>
		<%else%>
			<tr><td class="noRecords">Nessun record trovato</th></tr>
		<% end if %>
	</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing%>
