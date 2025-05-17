<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione valute - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA VALUTA"
dicitura.link_new = "Tabelle.asp;ValuteNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT *, (SELECT COUNT(*) FROM gtb_rivenditori WHERE riv_valuta_id=valu_id) AS N_RIV FROM gtb_valute ORDER BY valu_nome"
session("B2B_VALUTE_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco valute - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" width="5%">ID</th>
			<th class="center" width="10%">CODICE ISO</th>
			<th class="center" width="8%">SIMBOLO</th>
			<th>NOME</th>
			<th width="25%">TASSO DI CAMBIO VALUTA/&euro;URO</th>
			<th class="center" colspan="2" width="20%">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("valu_id") %></td>
				<td class="content_center"><%= rs("valu_codice") %></td>
				<td class="content_center"><%= rs("valu_simbolo") %></td>
				<td class="content"><%= rs("valu_nome") %></td>
				<td class="content"><%= rs("valu_cambio") %><%= rs("valu_simbolo") %> = 1&euro;</td>
				<td class="Content_center">
					<a class="button" href="ValuteMod.asp?ID=<%= rs("valu_id") %>">
						MODIFICA
					</a>
				</td>
				<td style="vertical-align:middle;" class="Content_center">
					<% if rs("N_RIV") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la valuta: sono presenti clienti associati"<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('VALUTE','<%= rs("valu_id") %>');" >
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