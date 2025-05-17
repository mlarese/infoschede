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
sezione_testata = "modifica dettagli del carico"
testata_show_back = false %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>
<SCRIPT LANGUAGE="javascript" src="Tools_B2B.js" type="text/javascript"></SCRIPT>

<% 
dim conn, rs, rsc, sql, cli_id, listino_id
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")

' Recupera i dati del  dettaglio
sql = "SELECT * FROM gv_carichi LEFT JOIN grel_giacenze ON gv_carichi.rel_id = grel_giacenze.gia_art_var_id WHERE rcv_id =" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

'recupera i dati del prodotto, compresa la marca ...

dim caricoID,magazzinoID
sql = "SELECT * FROM gtb_carichi INNER JOIN gtb_magazzini ON gtb_carichi.car_magazzino_id = gtb_magazzini.mag_id WHERE gtb_carichi.car_id=" & rs("rcv_car_id")
rsc.open sql, conn
	caricoID = rsc("car_id")
	magazzinoID = rsc("mag_id")
rsc.close
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_rcv_car_id" value="<%= rs("rcv_car_id") %>">
		<input type="hidden" name="tfn_rcv_art_var_id" value="<%= rs("rcv_art_var_id") %>">
		<input type="hidden" name="old_qta" value="<%= rs("rcv_qta") %>">
		<input type="hidden" name="MAG_ID" value="<%= magazzinoID %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica dettaglio</caption>
			<tr><th colspan="7">DATI ARTICOLO</th></tr>
			<% CALL ArticoloScheda (conn, rs, rsc) %>
			<tr>
				<td class="label">giacenza:</td>
				<td class="content" colspan="3"><%= rs("gia_qta") %></td>
			</tr>
			
			<tr><th colspan="7">DATI DETTAGLIO</th></tr>
			<tr>
				<td class="label">quantit&agrave;</td>
				<td class="content" colspan="6">
					<input type="text" class="text" tabindex="1" name="tfn_rcv_qta" value="<%= rs("rcv_qta") %>" maxlength="10" size="3">
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

