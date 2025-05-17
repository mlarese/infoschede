<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione esiti interventi - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVO ESITO"
dicitura.link_new = "Tabelle.asp;EsitiNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT *, (SELECT COUNT(*) FROM sgtb_schede WHERE sc_esito_intervento_id = esi_id) AS N_SCHEDE FROM sgtb_esiti " + _
	  " ORDER BY esi_nome_it"
session("INFOSCHEDE_ESITI_SQL") = sql

rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco esiti interventi - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" width="5%">ID</th>
			<th>NOME</th>
			<th class="center" width="20%" colspan="2">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("esi_id") %></td>
				<td class="content"><%= rs("esi_nome_it") %></td>
				<td class="Content_center">
					<a class="button" href="EsitiMod.asp?ID=<%= rs("esi_id") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<% if rs("N_SCHEDE") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare l'esito: sono presenti schede associati"<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ESITO','<%= rs("esi_id") %>');">
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