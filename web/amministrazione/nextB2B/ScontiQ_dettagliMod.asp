<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ScontiQ_dettagliSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica intervallo di sconto" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM gtb_scontiQ WHERE sco_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica intervallo di sconto</caption>
			<tr><th colspan="3">DATI DELL'INTERVALLO</th></tr>
			<tr>
				<td class="label">a partire da</td>
				<td class="content" width="20%">
					<input type="text" class="text" name="tfn_sco_qta_da" value="<%= rs("sco_qta_da") %>" maxlength="10" size="5">
				</td>
				<td class="note">n&ordm; unit&agrave;</td>
			</tr>
			<script type="text/javascript">
				function TipoVariazione(isPerc)
				{
					if (isPerc)
					{
						document.form1.tfn_sco_prezzo.value = '0,00';
						document.form1.tfn_sco_prezzo.className = "text_disabled";
						document.form1.tfn_sco_sconto.className = "text";
					}
					else
					{
						document.form1.tfn_sco_sconto.value = '0';
						document.form1.tfn_sco_sconto.className = "text_disabled";
						document.form1.tfn_sco_prezzo.className = "text";
					}
				}
			</script>
			<tr>
				<td class="label" rowspan="2">variazione in</td>
				<td class="content">
					<input class="checkbox" type="radio" name="sconto" id="sconto_percentuale" <%= chk(cReal(rs("sco_sconto"))<>0) %> onclick="TipoVariazione(true);">
					percentuale
				</td>
				<td class="content">
					<input type="text" class="text" name="tfn_sco_sconto" value="<%= rs("sco_sconto") %>" maxlength="10" size="3"> %
				</td>
			</tr>
			<tr>
				<td class="content">
					<input class="checkbox" type="radio" name="sconto" <%= chk(cReal(rs("sco_prezzo"))>0) %> onclick="TipoVariazione(false);">
					euro
				</td>
				<td class="content">
					<input type="text" class="text" name="tfn_sco_prezzo" value="<%= FormatPrice(rs("sco_prezzo"), 2, true) %>" maxlength="10" size="6"> €
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
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
<script type="text/javascript">
	<% if cReal(rs("sco_prezzo"))>0 then %>
		TipoVariazione(false);
	<% else %>
		TipoVariazione(true);
	<% end if %>
	FitWindowSize(this);
</script>
<% rs.close
conn.close
set rs = nothing
set conn = nothing %>
