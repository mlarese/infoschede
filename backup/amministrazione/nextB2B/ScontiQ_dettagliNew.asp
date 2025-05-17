<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ScontiQ_dettagliSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "inserimento nuovo intervallo di sconto" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_sco_classe_id" value="<%= request("EXTID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Inserimento nuovo intervallo di sconto</caption>
			<tr><th colspan="3">DATI DELL'INTERVALLO</th></tr>
			<tr>
				<td class="label">a partire da</td>
				<td class="content" width="20%">
					<input type="text" class="text" name="tfn_sco_qta_da" value="<%= request("tfn_sco_qta_da") %>" maxlength="10" size="5">
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
					<input class="checkbox" type="radio" name="sconto" id="sconto_percentuale" <%= chk(request("sconto")="") %> onclick="TipoVariazione(true);">
					percentuale
				</td>
				<td class="content">
					<input type="text" class="text" name="tfn_sco_sconto" value="<%= request("tfn_sco_sconto") %>" maxlength="10" size="3"> %
				</td>
			</tr>
			<tr>
				<td class="content">
					<input class="checkbox" type="radio" name="sconto" onclick="TipoVariazione(false);">
					euro
				</td>
				<td class="content">
					<input type="text" class="text" name="tfn_sco_prezzo" value="<%= FormatPrice(request("tfn_sco_prezzo"), 2, true) %>" maxlength="10" size="6"> €
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
	TipoVariazione(true);
	FitWindowSize(this);
</script>

