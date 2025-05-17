<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniTabelleQuerySalva.asp")
end if

%>

<%'--------------------------------------------------------
sezione_testata = "modifica query per creazione tag" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tb_siti_tabelle_tag_query WHERE tq_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>


<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_tq_tab_id" value="<%= request("TAB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica dati query</caption>
			<tr><th colspan="3">DATI DELLA QUERY</th></tr>
			<tr>
				<td class="label" colspan="1">nome:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_tq_nome" value="<%= CBR(rs, "tq_nome", "tft_") %>" maxlength="100" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="1">separatore:</td>
				<td class="content">
					<input type="text" class="text" name="tft_tq_separatore" value="<%= CBR(rs, "tq_separatore", "tft_") %>" maxlength="10" size="10">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="1">query:</td>
				<td class="content">
					<textarea name="tft_tq_query" style="width:97%;" rows="10"><%= server.HtmlEncode(rs("tq_query")) %></textarea>					
					(*)
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="4">
					<%=Server.HtmlEncode("Per i campi multilingua inserirne il valore come campo_<lingua> mentre per l'id del contenuto inserire <id>")%>
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