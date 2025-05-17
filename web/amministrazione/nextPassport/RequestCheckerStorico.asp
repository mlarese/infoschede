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

%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Log del Request Checker"
dicitura.puls_new = "INDIETRO A STRUMENTI"
dicitura.link_new = "Strumenti.asp"
dicitura.scrivi_con_sottosez()


sql = "SELECT * FROM log_request_checker ORDER BY log_date DESC "
CALL Pager.OpenSmartRecordset(conn, rs, sql, 25)
%>
<div id="content_liquid">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">		
		<%if not rs.eof then%>
			<caption>
					Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine
			</caption>
			<tr>
			<tr>
				<th class="center" width="15%">DATA</th>
				<th class="center" width="30%">NOME PARAMETRO</th>
				<th class="center" width="20%">VALORE PARAMETRO</th>
				<th class="center" width="35%">URL</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
				<tr>
					<td class="content_center"><%= DateTimeIta(rs("log_date")) %></td>
					<td class="content"><%= Server.HtmlEncode(rs("log_parameter_name")) %></td>
					<td class="content"><%= Server.HtmlEncode(rs("log_parameter_value")) %></td>
					<td class="content"><%= rs("log_url") %></td>
				</tr>
				<%rs.movenext
			wend%>
			
			<tr>
				<td colspan="4" class="footer" style="text-align:left;">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
				</td>
			</tr>
		<% else %>
			<tr><td class="noRecords">Nessun errore rilevato</th></tr>
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