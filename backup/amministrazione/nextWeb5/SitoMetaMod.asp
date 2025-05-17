<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoMetaSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica metatag" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tb_webs_metatag WHERE meta_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>


<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_meta_web_id" value="<%= request("WEB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica dati metatag</caption>
			<tr><th colspan="3">DATI DEL METATAG</th></tr>
			<tr>
				<td class="label" colspan="1">name:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_meta_name" value="<%= CBR(rs, "meta_name", "tft_") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="1">content:</td>
				<td class="content">
					<input type="text" class="text" name="tft_meta_content" value="<%= CBR(rs, "meta_content", "tft_") %>" maxlength="255" size="100">
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