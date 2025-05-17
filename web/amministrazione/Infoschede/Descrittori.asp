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
dicitura.sezione = "Gestione controlli per riparazioni - elenco"
dicitura.puls_new = "NUOVO CONTROLLO"
dicitura.link_new = "DescrittoriNew.asp"
dicitura.scrivi_con_sottosez() 


dim conn, rs, sql, Pager
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM sgtb_descrittori LEFT OUTER JOIN sgtb_descrittori_raggruppamenti " & _
	  " ON sgtb_descrittori.des_raggruppamento_id = sgtb_descrittori_raggruppamenti.rag_id" & _
	  " ORDER BY des_nome_IT "
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco controlli - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<% if not rs.eof then %>
			<tr>
				<th class="center" style="width:3%;">ID</th>
				<th>NOME</th>
				<th width="20%">TIPO</th>
				<th class="center" width="8%">PRINCIPALE</th>
				<th class="center" width="20%">GRUPPO</th>
				<th class="center" colspan="2" style="width:21%;">OPERAZIONI</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
				<tr>
					<td class="content_center"><%= rs("des_id") %></td>
					<td class="Content"><%= rs("des_nome_IT") %></td>
					<td class="Content"><%= DesVisTipo(rs("des_tipo")) %></td>
					<td class="Content_center"><input type="checkbox" disabled class="Checkbox" <%= chk(rs("des_principale")) %>></td>
					<td class="Content"><%= rs("rag_titolo_it") %></td>
					<td class="Content_center">
						<a class="button" href="DescrittoriMod.asp?ID=<%= rs("des_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="Content_center">
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('DESCRITTORI','<%= rs("des_id") %>');" >
							CANCELLA
						</a>
					</td>
				</tr>
				<% rs.moveNext
			wend%>
			<tr>
				<td colspan="7" class="footer" style="text-align:left;">
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

	
