<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SchedeStatiLavorazioneSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("INFOSCHEDE_STATI_ORDINE_SQL"), "sts_id", "SchedeStatiLavorazioneMod.asp")
end if


dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione stati di lavorazione delle schede - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SchedeStatiLavorazione.asp"
dicitura.scrivi_con_sottosez()  


sql = "SELECT * FROM sgtb_stati_schede WHERE sts_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica stato di lavorazione delle schede</td>
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
					<input type="text" class="text" name="tft_sts_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("sts_nome_"& Application("LINGUE")(i)) %>" maxlength="200" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tfn_sts_ordine" value="<%= rs("sts_ordine") %>" maxlength="10" size="3">
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="3">visibilit&agrave;:</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_visibile_admin" <%= chk(rs("sts_visibile_admin")) %>>
				admin
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_visibile_officina" <%= chk(rs("sts_visibile_officina")) %>>
				officina
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_visibile_centr_assist" <%= chk(rs("sts_visibile_centr_assist")) %>>
				centro assistenza
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="3">modifica:</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_modifica_admin" <%= chk(rs("sts_modifica_admin")) %>>
				admin
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_modifica_officina" <%= chk(rs("sts_modifica_officina")) %>>
				officina
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_modifica_centr_assist" <%= chk(rs("sts_modifica_centr_assist")) %>>
				centro assistenza
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2">visualizza per composizione documenti:</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_elenco_ddt_da_ritirare" <%= chk(rs("sts_elenco_ddt_da_ritirare")) %>>
				richieste di ritiro
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_elenco_ddt_da_consegnare" <%= chk(rs("sts_elenco_ddt_da_consegnare")) %>>
				ddt di spedizione
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