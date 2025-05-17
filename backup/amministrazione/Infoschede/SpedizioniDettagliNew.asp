<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="../nextB2B/Tools_B2B.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SpedizioniDettagliSalva.asp")
end if

'--------------------------------------------------------
sezione_testata = "inserimento nuovo articolo sulla spedizione"
testata_show_back = true %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, rsc, sql, value
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")

sql = " SELECT * FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " & _
	  " INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id WHERE rel_id=" & cIntero(request("RELID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText


%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_dtd_ddt_id" value="<%= request("ID_EXT") %>">
		<input type="hidden" name="tfn_dtd_articolo_id" value="<%= request("RELID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Dettaglio articolo per la spedizione</caption>
			<tr><th colspan="7">DATI ARTICOLO DA COLLEGARE</th></tr>
			<% CALL ArticoloScheda(conn, rs, rsc) %>
			<tr><th colspan="7">RIEPILOGO</th></tr>
			<tr>
				<td class="label" style="width:20%;">codice:</td>
				<td class="content" colspan="6">
					<input type="text" class="text" name="tft_dtd_articolo_codice" value="<%= rs("rel_cod_int") %>" maxlength="100" size="20">
				</td>
			</tr>
			<tr>
				<td class="label">ricambio:</td>
				<td class="content" colspan="6">
					<input type="text" class="text" name="tft_dtd_articolo_nome" value="<%= rs("art_nome_it") %>" maxlength="255" size="50">
				</td>
			</tr>
			<tr>
				<td class="label">quantit&agrave;:</td>
				<td class="content" colspan="6">
					<input type="text" class="number" name="tfn_dtd_articolo_qta" value="1" size="4">
				</td>
			</tr>
			<tr>
				<td class="label">prezzo di listino:</td>
				<td class="content" colspan="6">
					<%= FormatPrice(cReal(rs("art_prezzo_base")), 2, false) %> &euro;
				</td>
			</tr>
			<tr>
				<td class="label">prezzo unitario:</td>
				<td class="content" colspan="6">
					<input type="text" class="number" name="tfn_dtd_articolo_prezzo_unitario" value="<%= FormatPrice(cReal(rs("art_prezzo_base")), 2, false) %>" size="7"> &euro;
				</td>
			</tr>
			<tr>
				<td class="label">sconto:</td>
				<td class="content" colspan="6">
					<input type="text" class="number" name="tfn_dtd_articolo_sconto" value="<%= request("tfn_dtd_articolo_sconto") %>" size="7"> %
				</td>
			</tr>
			<tr>
				<td class="label">rif. vs ddt:</td>
				<td class="content" colspan="6">
					<input type="text" class="text" name="tft_dtd_rif_vs_ddt" value="<%= request("tft_dtd_rif_vs_ddt") %>" maxlength="100" size="22">
				</td>
			</tr>
			<tr>
				<td class="label">in garanzia:</td>
				<td class="content" colspan="6">
					<input type="checkbox" class="noBorder" name="chk_dtd_in_garanzia" <%= chk(request("chk_dtd_in_garanzia")<>"") %>>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="7">
					(*) Campi obbligatori.
					<input style="width:22%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:22%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
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