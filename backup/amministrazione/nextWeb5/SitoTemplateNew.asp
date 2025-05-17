<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_template_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoTemplateSalva.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - templates - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoTemplate.asp"
dicitura.scrivi_con_sottosez() 

dim conn, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_template" value="1">
	<input type="hidden" name="tfn_id_webs" value="<%= Session("AZ_ID") %>">
	<input type="hidden" name="tfd_contRes" value="NOW">
	<input type="hidden" name="tfn_contatore" value="0">
	<input type="hidden" name="tfn_contUtenti" value="0">
	<input type="hidden" name="tfn_contCrawler" value="0">
	<input type="hidden" name="tfn_contAltro" value="0">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
		<caption>Inserimento nuovo template</caption>
		<tr><th colspan="2">DATI DEL TEMPLATE</th></tr>
		<tr>
			<td class="label">titolo:</td>
			<td class="content">
				<input type="text" class="text" name="tft_nomepage" value="<%= request("tft_nomepage") %>" maxlength="250" size="100">
				(*)
			</td>
		</tr>
		<% 
		if not Session("SITO_MOBILE") then
		%>
		<tr><th class="L2" colspan="4">PROPRIET&Agrave;</th></tr>
		<tr>
			<td class="label" rowspan="2">tipo:</td>
			<td class="content">
				<input type="radio" class="checkbox" name="tfn_semplificata" value="0" <%= Chk(cIntero(request("tfn_semplificata")) = 0) %>>
				per pagina normale
			</td>
		</tr>
		<tr>
			<td class="content">
				<table cellpadding="0" cellspacing="0">
					<tr>
						<td><input type="radio" class="checkbox" name="tfn_semplificata" value="1" <%= Chk(cIntero(request("tfn_semplificata")) = 1) %>></td>
						<td style="padding-right:4px;"><img src="../grafica/notReadKnow.gif" border="0" alt="Template per email con visualizzazione semplificata."></td>
						<td>per email semplificate</td>
					</tr>
				</table>
			</td>
		</tr>
		<% 
		end if
		%>
		<%
		sql = QryElencoTemplate("", false)
		if cString(GetValueList(conn, NULL, sql))<>"" then %>
			<tr><th class="L2" colspan="4">OPZIONI DI GENERAZIONE</th></tr>
			<tr>
				<td class="label">copia da template:</td>
				<td class="content">
					<% 
					CALL dropDown(conn, sql, "id_page", "name", "template_padre", "", FALSE, "", LINGUA_ITALIANO)
					%>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA &gt;&gt;">
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
set conn = nothing%>