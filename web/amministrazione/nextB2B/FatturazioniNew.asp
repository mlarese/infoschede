<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("FatturazioniSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione valute - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Fatturazioni.asp"
dicitura.scrivi_con_sottosez()  

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova tipologia di fatturazione</caption>
		<tr><th colspan="3">DATI DELLA TIPOLOGIA DI FATTURAZIONE</th></tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content" style="width:45%;">
				<input type="text" class="text" name="tft_fatt_codice" value="<%= request("tft_fatt_codice") %>" maxlength="255" size="50">
			</td>
			<td class="note">
				Codice della tipologia di fatturazione.
			</td>
		</tr>
		<tr>
			<td class="label">numero corrente:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_fatt_numero_corrente" value="0" maxlength="5" size="5">
				(*)
			</td>
			<td class="note">
				Numero di partenza delle fatturazioni di questa tipologia.
			</td>
		</tr>
		<tr>
			<td class="label">data corrente:</td>
			<td class="content" >
				<table>
				<tr>
					<td class="content">
						<% CALL WriteDataPicker_Input("form1", "tfd_fatt_data_corrente", IIF(isDate(request("tfd_fatt_data_corrente")), request("tfd_fatt_data_corrente"), Date()), "", "/", true, true, LINGUA_ITALIANO) %>				
					</td>
					<td>&nbsp;(*)</td>
				</tr>
				</table>
			<td class="note">
				Data di creazione della tipologia di fatturazione. 
			</td>
		</tr>
		<tr>
			<td class="label">serie:</td>
			<td class="content">
				<input type="text" class="text" name="tft_fatt_serie" value="<%= request("tft_fatt_serie") %>" maxlength="10" size="10">
				(*)
			</td>
			<td class="note">
				Suffisso da apporre al numero di fatturazione.
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