<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ValuteSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione valute - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Valute.asp"
dicitura.scrivi_con_sottosez()  

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova valuta</caption>
		<tr><th colspan="3">DATI DELLA VALUTA</th></tr>
		<tr>
			<td class="label">codice ISO:</td>
			<td class="content" style="width:20%;">
				<input type="text" class="text" name="tft_valu_codice" value="<%= request("tft_valu_codice") %>" maxlength="3" size="3">
			</td>
			<td class="note">
				Codice internazionale della valuta secondo le specifiche ISO 4217.
			</td>
		</tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_valu_nome" value="<%= request("tft_valu_nome") %>" maxlength="50" size="75">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">tasso di cambio:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_valu_cambio" value="<%= request("tfn_valu_cambio") %>" maxlength="20" size="10">
				= 1 &euro;
				&nbsp;(*)
			</td>
			<td class="note">
				Tasso di cambio della valuta in euro. 
			</td>
		</tr>
		<tr><th colspan="3">FORMATTAZIONE DEI PREZZI</th></tr>
		<tr>
			<td class="label">simbolo:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_valu_simbolo" value="<%= request("tft_valu_simbolo") %>" maxlength="5" size="4">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">n&ordm; cifre decimali:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tfn_valu_num_decimali" value="<%= request("tfn_valu_num_decimali") %>" maxlength="1" size="1">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2"> caratteri separatori:</td>
			<td class="label_no_width">delle migliaia:</td>
			<td class="content">
				<input type="text" class="text" name="tft_valu_sep_migliaia" value="<%= request("tft_valu_sep_migliaia") %>" maxlength="1" size="1">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label_no_width">delle cifre decimali:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_valu_sep_decimali" value="<%= request("tft_valu_sep_decimali") %>" maxlength="1" size="1">
				(*)
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>