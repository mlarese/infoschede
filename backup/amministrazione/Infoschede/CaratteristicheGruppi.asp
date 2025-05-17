<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "CARATTERISTICHE"
dicitura.links(1) = "Caratteristiche.asp"
dicitura.sezione = "Gruppi di caratteristiche tecniche - elenco"
dicitura.puls_new = "NUOVO GRUPPO"
dicitura.link_new = "CaratteristicheGruppiNew.asp"
dicitura.scrivi_con_sottosez()

dim conn, rs, sql, Pager
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM gtb_carattech_raggruppamenti ORDER BY ctr_titolo_it"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco gruppi di caratteristiche tecniche - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<% if not rs.eof then %>
			<tr>
				<th class="center" style="width:3%;">ID</th>
				<th>NOME</th>
				<th class="center" width="8%">ORDINE</th>
				<th class="center" width="9%">DI SISTEMA</th>
				<th class="center" colspan="2" style="width:21%;">OPERAZIONI</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
				<tr>
					<td class="content_center"><%= rs("ctr_id") %></td>
					<td class="Content">
						<%= rs("ctr_titolo_IT") %>
					</td>
					<td class="Content_center"><%= rs("ctr_ordine") %></td>
					<td class="Content_center">
						<input type="checkbox" class="checkbox" <%= chk(rs("ctr_di_sistema")) %> disabled>
					</td>
					<td class="Content_center">
						<a class="button" href="CaratteristicheGruppiMod.asp?ID=<%= rs("ctr_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="Content_center">
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CTECH_GRUPPI','<%= rs("ctr_id") %>');" >
							CANCELLA
						</a>
					</td>
				</tr>
				<% rs.moveNext
			wend%>
			<tr>
				<td colspan="6" class="footer" style="text-align:left;">
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

	
