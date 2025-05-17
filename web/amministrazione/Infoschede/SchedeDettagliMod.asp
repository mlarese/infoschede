<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="../nextB2B/Tools_B2B.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SchedeDettagliSalva.asp")
end if

'--------------------------------------------------------
sezione_testata = "modifica collegamento con altro articolo"%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------

dim conn, rs, rsc, sql, in_garanzia, value
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")

sql = " SELECT * FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " & _
	  " INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " & _
	  " RIGHT JOIN sgtb_dettagli_schede ON grel_art_valori.rel_id = sgtb_dettagli_schede.dts_ricambio_id " & _
	  " WHERE dts_id = " & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

sql = "SELECT sc_in_garanzia FROM sgtb_schede WHERE sc_id = " & rs("dts_scheda_id")
in_garanzia = cBoolean(GetValueList(conn, NULL, sql), false)

%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_dts_scheda_id" value="<%= rs("dts_scheda_id") %>">
		<input type="hidden" name="tfn_dts_ricambio_id" value="<%= rs("dts_ricambio_id") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Dettaglio ricambio per la scheda</caption>
			<% if cIntero(rs("rel_id"))>0 then %>
				<tr><th colspan="7">DATI RICAMBIO DA COLLEGARE</th></tr>
				<% CALL ArticoloScheda(conn, rs, rsc) %>
			<% end if %>
			<tr><th colspan="7">RIEPILOGO</th></tr>
			
			<script language="JavaScript" type="text/javascript">
				function Ricalcola(){
					var prezzo_unit = toNumber(form1.tfn_dts_ricambio_prezzo.value);
					var sconto = toNumber(form1.tfn_dts_ricambio_sconto.value);
					var qta = toNumber(form1.tfn_dts_ricambio_qta.value);
					qta = FormatNumber(qta, 0);
					qta = toNumber(qta);
					var totale = toNumber(prezzo_unit*qta);
					totale = totale - ((totale*sconto)/100);

					
					//form1.ricambio_prezzo.value = FormatNumber(prezzo_unit, 2);
					form1.tfn_dts_ricambio_prezzo.value = FormatNumber(prezzo_unit, 2);
					
					form1.tfn_dts_ricambio_qta.value = FormatNumber(qta, 0);
					
					//form1.ricambio_sconto.value = FormatNumber(sconto, 2);
					form1.tfn_dts_ricambio_sconto.value = FormatNumber(sconto, 2);

					form1.prezzo_totale.value = FormatNumber(totale, 2);
					form1.tfn_dts_prezzo_totale.value = FormatNumber(totale, 2);
				}		
			</script>
			
			<tr>
				<td class="label">prezzo di listino:</td>
				<td class="content" colspan="6">
					<%= FormatPrice(cReal(rs("art_prezzo_base")), 2, false) %> &euro;
				</td>
			</tr>
			
			<tr>
				<td class="label">codice:</td>
				<td class="content" colspan="6">
					<input type="text" class="text" name="tft_dts_ricambio_codice" value="<%= rs("dts_ricambio_codice") %>" maxlength="100" size="22">
				</td>
			</tr>
			<tr>
				<td class="label">ricambio:</td>
				<td class="content" colspan="6">
					<input type="text" class="text" name="tft_dts_ricambio_nome" value="<%= rs("dts_ricambio_nome") %>" maxlength="255" size="60">
				</td>
			</tr>
			<tr>
				<td class="label">prezzo unitario:</td>
				<td class="content" colspan="6">
					<input type="text" <%=IIF(in_garanzia, "disabled", "")%> class="number" name="tfn_dts_ricambio_prezzo" value="<%= FormatPrice(cReal(rs("dts_ricambio_prezzo")), 2, false) %>" size="7" onchange="Ricalcola()"> &euro;
				</td>
			</tr>
			<tr>
				<td class="label">quantit&agrave;:</td>
				<td class="content" colspan="6">
					<input type="text" class="number" name="tfn_dts_ricambio_qta" value="<%= rs("dts_ricambio_qta") %>" size="5" onchange="Ricalcola()">
				</td>
			</tr>
			<tr>
				<td class="label">sconto:</td>
				<td class="content" colspan="6">
					<input type="text" <%=IIF(in_garanzia, "disabled", "")%> class="number" name="tfn_dts_ricambio_sconto" value="<%= rs("dts_ricambio_sconto") %>" size="5" onchange="Ricalcola()"> %
				</td>
			</tr>
			
			<tr>
				<td class="label">prezzo totale:</td>
				<td class="content" colspan="6">
					<input type="hidden" name="tfn_dts_prezzo_totale" value="<%= FormatPrice(cReal(rs("dts_prezzo_totale")), 2, false) %>"> 
					<input disabled  type="text" class="number" name="prezzo_totale" value="<%= FormatPrice(cReal(rs("dts_prezzo_totale")), 2, false) %>" size="7"> &euro;
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
	Ricalcola();
</script>