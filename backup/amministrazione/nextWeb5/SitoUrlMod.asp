<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoUrlSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica dati url" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tb_webs_directories WHERE dir_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>


<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_dir_web_id" value="<%= request("WEB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica dati url</caption>
			<tr><th colspan="3">DATI DELL'URL</th></tr>
			<tr>
				<td class="label" colspan="2">url completo:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_dir_url" value="<%= CBR(rs, "dir_url", "tft_") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<th style="width:11%;"><a href="http://www.google.it" target="_blank"><img src="../grafica/Google/Logo_25wht.gif" width="75" height="32" alt="Google" border="1"></a></th>
				<th style="vertical-align:middle;" colspan="2">INTEGRATION</th>
			</tr>
			<tr>
				<td class="label_no_width" rowspan="2"><a href="http://www.google.it/analytics" target="_blank" title="apri www.google.it/analytics in una nuova finestra">Google Maps</a></td>
				<td class="label_no_width" rowspan="2">chiave:</td>
				<td class="content">
					<input type="text" class="text" name="tft_dir_google_maps_key" value="<%= CBR(rs, "dir_google_maps_key", "tft_") %>" maxlength="255" size="100">
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
<% rs.close
conn.close
set rs = nothing
set conn = nothing %>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>