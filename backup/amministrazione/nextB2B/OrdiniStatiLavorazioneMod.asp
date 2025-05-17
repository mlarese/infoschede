<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("OrdiniStatiLavorazioneSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_STATI_ORDINE_SQL"), "so_id", "OrdiniStatiLavorazioneMod.asp")
end if


dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione stati di lavorazione dell'ordine - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "OrdiniStatiLavorazione.asp"
dicitura.scrivi_con_sottosez()  


sql = "SELECT * FROM gtb_stati_ordine WHERE so_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati dello stato di lavorazione</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="stato precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="stato successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="6">DATI DELLO STATO DI LAVORAZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content" colspan="2">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_so_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("so_nome_"& Application("LINGUE")(i)) %>" maxlength="200" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<% sql = "SELECT COUNT(*) FROM gtb_ordini WHERE ord_stato_id=" & rs("so_id") 
		if cInteger(GetValueList(conn, NULL, sql)) > 0 then%>
			<input type="hidden" name="tfn_so_stato_ordini" value="<%= rs("so_stato_ordini") %>">
			<tr>
				<td class="label" rowspan="2">stato degli ordini collegabili:</td>
				<td class="content<%= STILI_STATI_ORDINE(rs("so_stato_ordini")) %>" colspan="2">
					<%= STATI_ORDINE(rs("so_stato_ordini")) %>
				</td>
			</tr>
			<tr>
				<td class="note" colspan="2">
					La modifica dello stato degli ordini collegabili non &egrave; permessa perch&egrave; c'&egrave; almeno un ordine associato a tale stato.
				</td>
			</tr>
		<% else %>
			<tr>
				<td class="label" rowspan="4">stato degli ordini collegabili:</td>
				<td class="content<%= STILI_STATI_ORDINE(ORDINE_NON_CONFERMATO) %>" colspan="2">
					<input type="radio" class="checkbox" name="tfn_so_stato_ordini" value="<%= ORDINE_NON_CONFERMATO %>" <%= chk(rs("so_stato_ordini")=ORDINE_NON_CONFERMATO) %>>
					ordini non confermati
				</td>
			</tr>
			<tr>
				<td class="content<%= STILI_STATI_ORDINE(ORDINE_CONFERMATO) %>" colspan="2">
					<input type="radio" class="checkbox" name="tfn_so_stato_ordini" value="<%= ORDINE_CONFERMATO %>" <%= chk(rs("so_stato_ordini")=ORDINE_CONFERMATO) %>>
					ordini confermati
				</td>
			</tr>
			<tr>
				<td class="content<%= STILI_STATI_ORDINE(ORDINE_EVASO) %>" colspan="2">
					<input type="radio" class="checkbox" name="tfn_so_stato_ordini" value="<%= ORDINE_EVASO %>" <%= chk(rs("so_stato_ordini")=ORDINE_EVASO) %>>
					ordini evasi
				</td>
			</tr>
			<tr>
				<td class="content<%= STILI_STATI_ORDINE(ORDINE_ARCHIVIATO) %>" colspan="2">
					<input type="radio" class="checkbox" name="tfn_so_stato_ordini" value="<%= ORDINE_ARCHIVIATO %>" <%= chk(rs("so_stato_ordini")=ORDINE_ARCHIVIATO) %>>
					ordini archiviati
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tfn_so_ordine" value="<%= rs("so_ordine") %>" maxlength="10" size="3">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">stato di ingresso ordini internet:</td>
			<% if rs("so_internet") then %>
				<td class="content" width="15%">
					<input type="hidden" name="chk_so_internet" value="1">
					<input type="checkbox" class="noBorder" name="chk_so_internet" checked disabled>
				</td>
				<td class="note">Per impostare un altro stato come "ingresso di ordini internet" entrare in modifica dello stato stesso.</td>
			<% else%>
				<td class="content" width="15%">
					<input type="checkbox" name="chk_so_internet" class="noBorder" name="chk_so_internet">
				</td>
				<td class="note">Impostando lo stato come "stato di ingresso" verr&agrave; automaticamente tolta l'assegnazione allo stato attuale.</td>
			<% end if %>
		</tr>
		<tr><th colspan="3">NOTE</th></tr>
		<tr>
			<td class="content" colspan="3">
				<textarea style="width:100%;" rows="3" name="tft_so_descrizione"><%= rs("so_descrizione") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
set rs = nothing
conn.Close
set conn = nothing
%>