<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<% 
CALL CheckAutentication(Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE"))

dim Titolo_sezione, action, HREF
Titolo_sezione = "Gruppi di caratteristiche - elenco"
HREF = "ContattiDescrittoriGruppiNew.asp"
Action = "NUOVO GRUPPO"
SSezioniText = "CATEGORIE;CARATTERISTICHE"
SSezioniLink = "ContattiCategorie.asp;ContattiDescrittori.asp"
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%

dim conn, rs, sql, Pager
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM tb_indirizzario_carattech_raggruppamenti ORDER BY icr_titolo_it"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco gruppi di caratteristiche - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
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
					<td class="content_center"><%= rs("icr_id") %></td>
					<td class="Content">
						<%= rs("icr_titolo_IT") %>
					</td>
					<td class="Content_center"><%= rs("icr_ordine") %></td>
					<td class="Content_center">
						<input type="checkbox" class="checkbox" <%= chk(rs("icr_di_sistema")) %> disabled>
					</td>
					<td class="Content_center">
						<a class="button" href="ContattiDescrittoriGruppiMod.asp?ID=<%= rs("icr_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="Content_center">
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ContattiCTECH_GRUPPI','<%= rs("icr_id") %>');" >
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

	
