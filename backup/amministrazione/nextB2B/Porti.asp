<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione porti - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVO PORTO"
dicitura.link_new = "Tabelle.asp;PortiNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT *, (SELECT COUNT(*) FROM gtb_rivenditori WHERE riv_porto_default_id=prt_id) AS N_RIV FROM gtb_porti ORDER BY prt_nome_it"
session("B2B_PORTI_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco porti - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" width="5%">ID</th>
			<th>NOME</th>
			<th class="center" width="7%">CON SPESE</th>
			<th class="center" width="7%">CON VETTORE</th>
			<th class="center" width="7%">SCELTA SEDE</th>
			<th class="center" width="12%">SCELTA MOD. CONSEGNA</th>
			<th width="6%" class="center">ATTIVO</th>
			<th class="center" colspan="2" width="20%">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("prt_id") %></td>
				<td class="content"><%= rs("prt_nome_it") %></td>
				<td class="content_center">
					<input type="checkbox" disabled class="checkbox" <%= chk(rs("prt_con_spese")) %>>
				</td>
				<td class="content_center">
					<input type="checkbox" disabled class="checkbox" <%= chk(rs("prt_con_vettore")) %>>
				</td>
				<td class="content_center">
					<input type="checkbox" disabled class="checkbox" <%= chk(rs("prt_scelta_sede")) %>>
				</td>
				<td class="content_center">
					<input type="checkbox" disabled class="checkbox" <%= chk(rs("prt_scelta_modalita_consegna")) %>>
				</td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("prt_attivo")) %>></td>
				<td class="content_center">
					<a class="button" href="PortiMod.asp?ID=<%= rs("prt_id") %>">
						MODIFICA
					</a>
				</td>
				<td style="vertical-align:middle;" class="content_center">
					<% if rs("N_RIV") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il porto: sono presenti clienti associati"<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('PORTI','<%= rs("prt_id") %>');" >
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