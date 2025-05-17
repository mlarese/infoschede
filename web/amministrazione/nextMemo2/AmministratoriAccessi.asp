<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
dim i, conn, rs, rsA, sql, Pager

set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsA = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	Pager.Reset
	CALL GotoRecord(conn, rs, Session("SQL_AMMINISTRATORI"), "id_admin", "AmministratoriAccessi.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 

dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione utenti area amministrativa NextMemo 2.0 - log accessi"
if request("ID")="" then
	dicitura.puls_new = "INDIETRO;"
	dicitura.link_new = "Amministratori.asp;"
else
	dicitura.puls_new = "INDIETRO;DATI"
	dicitura.link_new = "Amministratori.asp;AmministratoriMod.asp?ID=" & request("ID")
end if
dicitura.scrivi_con_sottosez() 


sql = "SELECT * FROM (log_admin INNER JOIN tb_admin ON log_admin.log_admin_id = tb_admin.id_admin) " &_
  	  " INNER JOIN tb_siti ON log_admin.log_sito_id = tb_siti.id_sito " & _
	  " WHERE log_sito_id = 36 "
if request("ID")<>"" then
	sql = sql & " AND id_admin=" & cIntero(request("ID"))
end if
sql = sql & " ORDER BY log_data DESC "
CALL Pager.OpenSmartRecordset(conn, rs, sql, 25)
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% if request("ID")<>"" then
			sql = "SELECT * FROM tb_admin WHERE id_admin=" & cIntero(request("ID"))
			rsA.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
			<caption>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<% if not rs.eof then %>
						<td class="caption">
							Accessi effettuati dall'utente &quot;<%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %>&quot;
							 - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine
						</td>
						<% Else  %>
						<!--
						<td class="caption">
							Nessun accesso effettuato.
						</td>
						-->
						<% End If %>
						<td align="right" style="font-size: 1px;" nowrap>
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="utente precedente">
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="utente successiva">
								SUCCESSIVO &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			</caption>
			<%'rs.close
			%>
		<% else %>
			<caption>Elenco accessi effettuati da tutti gli utenti - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<%end if
		
		if not rs.eof then%>
			<tr>
				<th class="center" width="20%">DATA</th>
				<th>UTENTE</th>
				<th class="center" width="40%">APPLICAZIONE</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
				<tr>
					<td class="content_center"><%= DateTimeIta(rs("log_data")) %></td>
					<td class="content"><%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %></td>
					<td class="content"><%= rs("sito_nome") %></td>
				</tr>
				<%rs.movenext
			wend%>
			
			<tr>
				<td colspan="3" class="footer" style="text-align:left;">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
				</td>
			</tr>
		<% else %>
			<tr><td class="noRecords">Nessun accesso effettuato</th></tr>
		<% end if %>
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