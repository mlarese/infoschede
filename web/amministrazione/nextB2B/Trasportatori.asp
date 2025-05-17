<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione trasportatori - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVO TRAPORTATORE"
dicitura.link_new = "Tabelle.asp;TrasportatoriNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT *, (SELECT COUNT(*) FROM gtb_rivenditori WHERE riv_trasportatore_default_id=tra_id) AS N_RIV FROM gtb_trasportatori ORDER BY tra_nome_it"
session("B2B_TRASPORTATORI_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco trasportatori - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" width="5%">ID</th>
			<th>NOME</th>
			<th width="15%">CODICE</th>
			<th width="8%" class="center">ATTIVO</th>
			<th class="center" colspan="2" width="20%">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("tra_id") %></td>
				<td class="content"><%= rs("tra_nome_it") %></td>
				<td class="content"><%= rs("tra_codice") %></td>
				<td class="content_center"><input type="checkbox" disabled class="checkbox" <%= chk(rs("tra_attivo")) %>></td>
				<td class="content_center">
					<a class="button" href="TrasportatoriMod.asp?ID=<%= rs("tra_id") %>">
						MODIFICA
					</a>
				</td>
				<td style="vertical-align:middle;" class="Content_center">
					<% if rs("N_RIV") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il trasportatore: sono presenti clienti associati"<%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('TRASPORTATORI','<%= rs("tra_id") %>');" >
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