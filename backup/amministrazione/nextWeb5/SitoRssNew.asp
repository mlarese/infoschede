<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoRssSalva.asp")
end if


dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
%>
<%'--------------------------------------------------------
sezione_testata = "inserimento nuovo url" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_rss_web_id" value="<%= request("WEB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica dati rss</caption>
			<tr><th colspan="3">DATI DELL'RSS</th></tr>
			<tr>
				<td class="label" colspan="2">lingua:</td>				
				<td class="content" nowrap>
					<% CALL DropLingue(conn, NULL, "tft_rss_lingua", request("tft_rss_lingua"), true, false, "width:100px;") %>
					(*)
				</td>
			</tr>	
			<tr>
				<td class="label" colspan="2">titolo:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_rss_titolo" value="<%= request("tft_rss_titolo") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">file:</td>
				<td class="content" colspan="3">
					<input type="text" class="text" name="tft_rss_file" value="<%= request("tft_rss_file") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">abilitato:</td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_rss_abilitato" <%= chk(cInteger(request("tfn_rss_abilitato"))>0) %>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_rss_abilitato" <%= chk(cInteger(request("tfn_rss_abilitato"))=0) %>>
					no
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">metatag:</td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_rss_metatag" <%= chk(cInteger(request("tfn_rss_metatag"))>0) %>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_rss_metatag" <%= chk(cInteger(request("tfn_rss_metatag"))=0) %>>
					no
				</td>
			</tr>
			<tr>				
				<td class="label" colspan="2">descrizione:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_rss_descrizione" value="<%= request("tft_rss_descrizione") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">immagine:</td>
				<td class="content" colspan="3">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_rss_image", request("tft_rss_image") , "width:425px;", false) %>
				</td>
			</tr>
			<tr>
				<td class="label" valign="top" colspan="2">query:</td>
				<td class="content"><textarea style="width:528px;" rows="6" name="tft_rss_query" ><%= request("tft_rss_query")%></textarea>
				(*)
				</td>
				
			</tr>
			<tr>				
				<td class="label" colspan="2">frequenza di generazione:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tfn_rss_freq_generazione" value="<%= request("tfn_rss_freq_generazione") %>" size="10">
					(*)
				</td>
			</tr>
			<tr>				
				<td class="label" colspan="2">data generazione:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tfd_rss_data_generazione" value="<%=Now() %>" size="50">
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