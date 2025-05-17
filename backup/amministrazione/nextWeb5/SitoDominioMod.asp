<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoDominioSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica dati dominio aggiuntivo" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tb_webs_domini WHERE dom_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_dom_web_id" value="<%= request("WEB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica dati url</caption>
			<tr><th colspan="3">DATI DEL DOMINIO AGGIUNTIVO</th></tr>
			<tr>
				<td class="label" colspan="2">nome:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_dom_name" value="<%= CBR(rs, "dom_name", "tft_") %>" maxlength="50" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">url completo:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_dom_url" value="<%= CBR(rs, "dom_url", "tft_") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">lingua:</td>
				<td class="content" nowrap>
					<% CALL DropLingue(conn, NULL, "tft_dom_lingua", rs("dom_lingua"), true, false, "width:100px;") %>
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">ordine:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tfn_dom_ordine" value="<%= rs("dom_ordine") %>" size="5">
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">href_lang:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_dom_href_lang" value="<%= rs("dom_href_lang") %>" maxlength="50" size="10">
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