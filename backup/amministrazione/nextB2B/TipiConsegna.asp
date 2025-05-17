<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tipi consegna - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVO TIPI CONSEGNA"
dicitura.link_new = "Tabelle.asp;TipiConsegnaNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT *, " & _
	  " (SELECT COUNT(*) from gtb_ordini WHERE ord_tipo_consegna_id = tco_id) AS N_ORDINI, " & _
	  " (SELECT COUNT(*) from gtb_shopping_cart WHERE sc_tipo_consegna_id = tco_id) AS N_SHOP_CART " & _
	  " FROM gtb_tipo_consegna ORDER BY tco_ordine, tco_nome_it"
session("B2B_TIPI_CONSEGNA_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco tipi consegna - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" width="5%">ID</th>
			<th>NOME</th>
			<th class="center" width="10%">ORDINE</th>
			<th class="center" colspan="2" width="20%">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("tco_id") %></td>
				<td class="content"><%= rs("tco_nome_it") %></td>
				<td class="content_center"><%= rs("tco_ordine") %></td>
				<td class="content_center">
					<a class="button" href="TipiConsegnaMod.asp?ID=<%= rs("tco_id") %>">
						MODIFICA
					</a>
				</td>
				<td style="vertical-align:middle;" class="Content_center">
					<% if rs("N_ORDINI") > 0 OR rs("N_SHOP_CART") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il tipo consegna: sono presenti ordini o shopping cart associati"<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('TIPI_CONSEGNA','<%= rs("tco_id") %>');" >
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