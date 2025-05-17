<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoUrlSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "inserimento nuovo url" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_dir_web_id" value="<%= request("WEB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Inserimento nuovo url</caption>
			<tr><th colspan="3">DATI DELL'URL</th></tr>
			<tr>
				<td class="label" colspan="2">url completo:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_dir_url" value="<%= request("tft_dir_url") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<th style="width:11%;"><a href="http://www.google.it" target="_blank"><img src="../grafica/Google/Logo_25wht.gif" width="75" height="32" alt="Google" border="1"></a></th>
				<th style="vertical-align:middle;" colspan="2">INTEGRATION</th>
			</tr>
			<tr>
				<td class="label_no_width" rowspan="2"><a href="http://mappe.google.it" target="_blank" title="apri mappe.google.it in una nuova finestra">Google Maps</a></td>
				<td class="label_no_width" rowspan="2">chiave:</td>
				<td class="content">
					<input type="text" class="text" name="tft_dir_google_maps_key" value="<%= request("tft_dir_google_maps_key") %>" maxlength="255" size="100">
				</td>
			</tr>
			<tr>
				<td class="content note">
					Per ottenere una chiave valida &egrave; necessario riempire il form disponibile presso <a href="http://code.google.com/apis/maps/signup.html" target="_blank" title="apre in una nuova finestra la pagina di registrazione di un nuovo indirizzo per l'utilizzo di Google Maps">questo link</a>.
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="4">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
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