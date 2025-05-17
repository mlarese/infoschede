<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoMetaSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "inserimento metatag aggiuntivo" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_meta_web_id" value="<%= request("WEB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Inserimento Metatag aggiuntivo</caption>
			<tr><th colspan="3">DATI DEL METATAG</th></tr>
			<tr>
				<td class="label" colspan="1">name:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_meta_name" value="<%= request("tft_meta_name") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="1">content:</td>
				<td class="content">
					<input type="text" class="text" name="tft_meta_content" value="<%= request("tft_meta_content") %>" maxlength="255" size="100">
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