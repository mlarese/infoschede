<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("MagazziniCarichi_dettagliSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "inserimento nuovo dettaglio del carico a magazzino"
testata_show_back = true %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>
<SCRIPT LANGUAGE="javascript" src="Tools_B2B.js" type="text/javascript"></SCRIPT>

<% 
dim conn, rs, rsc, sql, cli_id, listino_id
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")



sql = "SELECT * FROM gv_articoli a "& _
	  "LEFT JOIN grel_giacenze g ON a.rel_id=g.gia_art_var_id "& _
	  "WHERE rel_id=" & cIntero(request("ART_ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_rcv_car_id" value="<%= request("CAR_ID") %>">
		<input type="hidden" name="tfn_rcv_art_var_id" value="<%= request("ART_ID") %>">
		<input type="hidden" name="MAG_ID" value="<%= request("MAG_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Selezione nuovo dettaglio</caption>
			<tr><th colspan="7">DATI ARTICOLO</th></tr>
			<% 	CALL ArticoloScheda (conn, rs, rsc) %>
			<tr>
				<td class="label">&nbsp;</td>
				<td class="content" colspan="2">&nbsp;</td>
				<td class="label">giacenza:</td>
				<td class="content" colspan="3"><%= rs("gia_qta") %></td>
			</tr>
			<tr>
				<td class="label">ordine min.:</td>
				<td class="content" colspan="2"><%= rs("rel_qta_min_ord") %></td>
				<td class="label">lotto riordino:</td>
				<td class="content" colspan="3"><%= rs("rel_lotto_riordino") %></td>
			</tr>
			<tr><th colspan="7">DATI DETTAGLIO</th></tr>
			<tr>
				<td class="label">quantit&agrave;</td>
				<td class="content" colspan="6">
					<input type="text" class="text" tabindex="1" name="tfn_rcv_qta" value="<%= rs("rel_qta_min_ord") %>" maxlength="10" size="3">
					(*)
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="7">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
		</table>
	</form>
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