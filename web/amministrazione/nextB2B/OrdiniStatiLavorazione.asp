<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione stati di lavorazione dell'ordine - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVO STATO DI LAVORAZIONE"
dicitura.link_new = "Tabelle.asp;OrdiniStatiLavorazioneNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT *, (SELECT COUNT(*) FROM gtb_ordini WHERE ord_stato_id = so_id) AS N_ORD FROM gtb_stati_ordine " + _
	  " ORDER BY so_stato_ordini, so_ordine"
session("B2B_STATI_ORDINE_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco stati degli ordini - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" width="5%">ID</th>
			<th>NOME</th>
			<th class="center" width="12%">ORDINAMENTO</th>
			<th class="center" width="22%">STATO ORDINI COLLEGATI</th>
			<th class="center" width="13%">ORDINI INTERNET</th>
			<th class="center" width="20%" colspan="2">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("so_id") %></td>
				<td class="content"><%= rs("so_nome_it") %></td>
				<td class="content_center"><%= rs("so_ordine") %></td>
				<td class="content<%= STILI_STATI_ORDINE(rs("so_stato_ordini")) %>" style="text-align:center;"><%= STATI_ORDINE(rs("so_stato_ordini")) %></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("so_internet")) %>></td>
				<td class="Content_center">
					<a class="button" href="OrdiniStatiLavorazioneMod.asp?ID=<%= rs("so_id") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<% if rs("N_ORD") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare lo stato: sono presenti ordini associati"<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% elseif rs("so_internet") then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare lo stato: stato di ingresso degli ordini internet.<%= vbCrLF %>Per cancellare lo stato corrente impostare come stato di ingresso un altro stato di lavorazione."<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('STATIO','<%= rs("so_id") %>');">
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