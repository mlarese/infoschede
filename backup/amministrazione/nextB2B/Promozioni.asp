<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione promozioni - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA PROMOZIONE"
dicitura.link_new = "Tabelle.asp;PromozioniNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM gtb_promozioni"
session("B2B_VALUTE_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco promozioni - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" width="5%">ID</th>
			<th class="center" width="15%">NOME</th>
			<th class="center" width="25%">DESCRIZIONE</th>
			<th class="center" width="10%">VALORE</th>
			<th class="center" width="15%">INIZIO VALIDITA'</th>
			<th class="center" width="15%">FINE VALIDITA'</th>
			<th class="center" width="15%">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("promo_id") %></td>
				<td class="content_center"><%= rs("promo_nome_it") %></td>
				<td class="content_center"><%= rs("promo_descrizione_it") %></td>
				<td class="content_center"><%= rs("promo_valore") %>%</td>
				<td class="content_center"><%= rs("promo_inizio_validita") %></td>
				<td class="content_center"><%= rs("promo_fine_validita") %></td>
				<td class="content_center">
					<a class="button" href="PromozioniMod.asp?ID=<%= rs("promo_id") %>">
						MODIFICA
					</a>
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