<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoDominioSalva.asp")
end if

dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<%'--------------------------------------------------------
sezione_testata = "inserimento nuovo dominio aggiuntivo" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_dom_web_id" value="<%= request("WEB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Inserimento nuovo dominio aggiuntivo</caption>
			<tr><th colspan="3">DATI DEL DOMINIO</th></tr>
			<tr>
				<td class="label" colspan="2">nome:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_dom_name" value="<%= request("tft_dom_name") %>" maxlength="50" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">url completo:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_dom_url" value="<%= request("tft_dom_url") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">lingua:</td>
				<td class="content" nowrap>
					<% CALL DropLingue(conn, NULL, "tft_dom_lingua", request("tft_dom_lingua"), true, false, "width:100px;") %>
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">ordine:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tfn_dom_ordine" value="<%= request("tfn_dom_ordine") %>" size="5">
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">hreflang:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_dom_href_lang" value="<%= request("tft_dom_href_lang") %>" maxlength="50" size="10">
				</td>
			</tr>
			<script language="JavaScript" type="text/javascript">
				function CorrectUrl()
				{
					var url = form1.tft_dom_url;
					if (url.value.indexOf("http://") < 0)
						form1.tft_dom_url.value = "http://" + form1.tft_dom_url.value;
				}
			</script>
			<tr>
				<td class="footer" colspan="4">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA" onclick="CorrectUrl();">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
	</form>
		</table>
</div>
</body>
</html>

<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>