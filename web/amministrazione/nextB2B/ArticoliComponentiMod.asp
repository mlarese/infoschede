<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ArticoliComponentiSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica del componente"
testata_show_back = true %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<% 
dim conn, rs, rsc, sql, typ, bun_id
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")

sql = " SELECT * FROM gv_articoli INNER JOIN gtb_bundle ON gv_articoli.rel_id = gtb_bundle.bun_articolo_id WHERE bun_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

sql = "SELECT art_id, art_se_bundle, art_se_confezione FROM gv_articoli WHERE rel_id=" & rs("bun_bundle_id")
rsc.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
typ = IIF(rsc("art_se_bundle"), "B", "C")
bun_id = rsc("art_id")
rsc.close
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_bun_bundle_id" value="<%= rs("bun_bundle_id") %>">
		<input type="hidden" name="tfn_bun_articolo_id" value="<%= rs("bun_articolo_id") %>">
		<input type="hidden" name="BUN_TYPE" value="<%= typ %>">
		<input type="hidden" name="BUN_ID" value="<%= bun_id %>">
		<input type="hidden" name="COM_ID" value="<%= rs("art_id") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Selezione nuovo componente per <%= IIF(typ="B", " il bundle ", " la confezione ") %></caption>
			<tr><th colspan="7">DATI ARTICOLO</th></tr>
			<% CALL ArticoloScheda (conn, rs, rsc) %>
			<tr><th colspan="7">DATI AGGREGAZIONE</th></tr>
			<tr>
				<td class="label"">quantit&agrave;</td>
				<td class="content" colspan="6">
					<input type="text" class="text" name="tfn_bun_quantita" value="<%= rs("bun_quantita") %>" maxlength="3" size="3"> (*)
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="7">
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
set rsc = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>