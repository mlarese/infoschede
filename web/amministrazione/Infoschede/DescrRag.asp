<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(2)
dicitura.sottosezioni(1) = "CONTROLLI RIPARAZIONI"
dicitura.links(1) = "Descrittori.asp"
dicitura.sottosezioni(2) = "GRUPPI"
dicitura.links(2) = "DescrRag.asp"
dicitura.sezione = "Gruppi controlli per riparazioni - elenco"
dicitura.puls_new = "NUOVO GRUPPO"
dicitura.link_new = "DescrRagNew.asp"
dicitura.scrivi_con_sottosez() 


dim conn, rs, sql, Pager
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM sgtb_descrittori_raggruppamenti"& _
	  " ORDER BY rag_titolo_it"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco gruppi caratteristiche anagrafiche - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<% if not rs.eof then %>
			<tr>
				<th class="center" style="width:3%;">ID</th>
				<th>NOME</th>
				<th class="center" width="8%">ORDINE</th>
				<th class="center" colspan="2" style="width:21%;">OPERAZIONI</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
				<tr>
					<td class="content_center"><%= rs("rag_id") %></td>
					<td class="Content">
						<%= rs("rag_titolo_IT") %>
						<% if cString(rs("rag_note"))<>"" then %>
							<br><span class="note"><%= rs("rag_note") %></span>
						<% end if %>
					</td>
					<td class="Content_center"><%= rs("rag_ordine") %></td>
					<td class="Content_center">
						<a class="button" href="DescrRagMod.asp?ID=<%= rs("rag_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="Content_center">
						<% if CInt(GetValueList(conn, NULL, "SELECT COUNT(*) FROM sgtb_descrittori WHERE des_raggruppamento_id="& rs("rag_id"))) = 0 then %>
							<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('DESCRAG','<%= rs("rag_id") %>');" >
								CANCELLA
							</a>
						<% else %>
							<a class="button_disabled" title="Impossibile cancellare il gruppo: sono presenti caratteristiche associate">
								CANCELLA
							</a>
						<% end if %>
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

	
