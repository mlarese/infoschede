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
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione stati di lavorazione delle schede - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SchedeStatiLavorazione.asp"
dicitura.scrivi_con_sottosez()  

dim conn, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo stato di lavorazione delle schede</caption>
		<tr><th colspan="3">DATI DELLO STATO DI LAVORAZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content" colspan="2">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_sts_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_sts_nome_"& Application("LINGUE")(i)) %>" maxlength="200" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tfn_sts_ordine" value="<%= request("tfn_sts_ordine") %>" maxlength="10" size="3">
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="3">visibilit&agrave;:</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_visibile_admin" <%= chk(request("chk_sts_visibile_admin")<>"") %>>
				admin
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_visibile_officina" <%= chk(request("chk_sts_visibile_officina")<>"") %>>
				officina
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_visibile_centr_assist" <%= chk(request("chk_sts_visibile_centr_assist")<>"") %>>
				centro assistenza
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="3">modifica:</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_modifica_admin" <%= chk(request("chk_sts_modifica_admin")<>"") %>>
				admin
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_modifica_officina" <%= chk(request("chk_sts_modifica_officina")<>"") %>>
				officina
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_modifica_centr_assist" <%= chk(request("chk_sts_modifica_centr_assist")<>"") %>>
				centro assistenza
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2">visualizza per composizione documenti:</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_elenco_ddt_da_ritirare" <%= chk(request("chk_sts_elenco_ddt_da_ritirare")<>"") %>>
				richieste di ritiro
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" class="noBorder" name="chk_sts_elenco_ddt_da_consegnare" <%= chk(request("chk_sts_elenco_ddt_da_consegnare")<>"") %>>
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
conn.close
set conn = nothing
%>