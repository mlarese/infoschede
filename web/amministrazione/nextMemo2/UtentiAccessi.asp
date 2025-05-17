<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
dim i, conn, rs, rsu, sql, Pager

set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsu = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	Pager.Reset
	CALL GotoRecord(conn, rs, Session("SQL_UTENTI"), "ut_id", "UtentiAccessi.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 

dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione utenti area riservata NextMemo 2.0 - log accessi"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Utenti.asp"
dicitura.scrivi_con_sottosez() 


sql = "SELECT * FROM ((tb_utenti INNER JOIN tb_Indirizzario " &_
	  " ON tb_utenti.ut_NextCom_ID=tb_Indirizzario.IDElencoIndirizzi) " & _
	  " INNER JOIN log_utenti ON tb_utenti.ut_id=log_utenti.log_ut_id) " &_
	  " INNER JOIN tb_siti ON log_utenti.log_sito_id=tb_siti.id_sito"
if request("ID")<>"" then
	sql = sql & " WHERE ut_id=" & cIntero(request("ID"))
end if

sql = sql & " ORDER BY log_data DESC "
CALL Pager.OpenSmartRecordset(conn, rs, sql, 25)
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% if request("ID")<>"" then
			sql = "SELECT * FROM tb_utenti INNER JOIN tb_Indirizzario " &_
				  " ON tb_utenti.ut_NextCom_ID=tb_Indirizzario.IDElencoIndirizzi WHERE ut_ID=" & request("ID")
			rsu.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
			<caption>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td class="caption">
							Accessi effettuati dall'utente &quot;
							<%= ContactFullName(rsu) %>
							&quot;
							 - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine
						</td>
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
			<%rsu.close%>
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
					<td class="content"><%= ContactFullName(rs) %></td>
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
set rsu = nothing
set conn = nothing
%>