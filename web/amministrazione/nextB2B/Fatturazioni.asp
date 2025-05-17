<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tipologie di fatturazioni - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA TIPOLOGIA DI FATTURAZIONE"
dicitura.link_new = "Tabelle.asp;FatturazioniNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT *,(SELECT COUNT(*) FROM gtb_ordini WHERE ord_fatturazione_id=fatt_id) as N_ORD FROM gtb_fatturazioni ORDER BY fatt_codice"
session("B2B_FATTURAZIONI_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco tipologie di fatturazioni - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" width="5%">ID</th>
			<th class="center" width="25%">CODICE FATTURAZIONE</th>
			<th class="center" width="15%">NUMERO CORRENTE</th>
			<th class="center" width="20%">DATA CORRENTE</th>
			<th class="center" width="10%">SERIE</th>
			<th class="center" colspan="2" width="20%">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("fatt_id") %></td>
				<td class="content_center"><%= rs("fatt_codice") %></td>
				<td class="content_center"><%= rs("fatt_numero_corrente") %></td>
				<td class="content_center"><%= DateIta(rs("fatt_data_corrente")) %></td>
				<td class="content_center"><%= rs("fatt_serie") %></td>
				<td class="Content_center">
					<a class="button" href="FatturazioniMod.asp?ID=<%= rs("fatt_id") %>">
						MODIFICA
					</a>
				</td>
				<td style="vertical-align:middle;" class="Content_center">
					<% if rs("N_ORD") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la tipologia di fatturazione: sono presenti ordini associati"<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('FATTURAZIONE','<%= rs("fatt_id") %>');" >
							CANCELLA
						</a>
					<% end if %>
				</td>
			</tr>
			<%rs.movenext
		wend%>
	</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing%>