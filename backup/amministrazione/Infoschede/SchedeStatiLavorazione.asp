<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione stati di lavorazione delle schede - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVO STATO DI LAVORAZIONE"
dicitura.link_new = "Tabelle.asp;SchedeStatiLavorazioneNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT *, (SELECT COUNT(*) FROM sgtb_schede WHERE sc_stato_id = sts_id) AS N_SCHEDE FROM sgtb_stati_schede " + _
	  " ORDER BY sts_ordine, sts_nome_it"
session("INFOSCHEDE_STATI_ORDINE_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco stati delle schede - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th rowspan="2">NOME</th>
			<th class="center" rowspan="2" width="6%">ORDINE</th>
			<th class="center" colspan="3" width="19%" style="border-bottom:0px;">VISUALIZZIONE</th>
			<th class="center" colspan="3" width="19%" style="border-bottom:0px;">MODIFICA</th>
			<th class="center" colspan="2" width="14%" style="border-bottom:0px;">DOCUMENTI</th>
			<th class="center" rowspan="2" width="20%" colspan="2">OPERAZIONI</th>
		</tr>
		<tr>
			<th class="l2_center" style="border-bottom:1px solid #999999;">ADMIN</th>
			<th class="l2_center" style="border-bottom:1px solid #999999;">OFFICINA</th>
			<th class="l2_center" style="border-bottom:1px solid #999999;">C. ASSIST.</th>
			<th class="l2_center" style="border-bottom:1px solid #999999;">ADMIN</th>
			<th class="l2_center" style="border-bottom:1px solid #999999;">OFFICINA</th>
			<th class="l2_center" style="border-bottom:1px solid #999999;">C. ASSIST.</th>
			<th class="l2_center" style="border-bottom:1px solid #999999;">DDT</th>
			<th class="l2_center" style="border-bottom:1px solid #999999;">RITIRI</th>
			
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content"><%= rs("sts_nome_it") %></td>
				<td class="content_center"><%= rs("sts_ordine") %></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("sts_visibile_admin")) %>></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("sts_visibile_officina")) %>></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("sts_visibile_centr_assist")) %>></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("sts_modifica_admin")) %>></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("sts_modifica_officina")) %>></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("sts_modifica_centr_assist")) %>></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("sts_elenco_ddt_da_consegnare")) %>></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("sts_elenco_ddt_da_ritirare")) %>></td>
				<td class="Content_center">
					<a class="button" href="SchedeStatiLavorazioneMod.asp?ID=<%= rs("sts_id") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<% if rs("N_SCHEDE") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare lo stato: sono presenti schede associate"<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('STATIS','<%= rs("sts_id") %>');">
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