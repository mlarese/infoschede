<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_immaginiFormati_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoImmaginiFormatiSalva.asp")
end if

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(1)
dicitura.sezione = "Gestione siti - formati immagini - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoImmaginiFormati.asp"
dicitura.sottosezioni(1) = "ELENCO FILES"
dicitura.links(1) = "SitoFileManager.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_imf_webId" value="<%= Session("AZ_ID") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
		<caption>Inserimento nuovo formato immagine</caption>
		<tr><th colspan="4">DATI DEL FORMATO</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_imf_nome" value="<%= request("tft_imf_nome") %>" maxlength="255" size="80">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">directory:</td>
			<td class="content" colspan="3">
				<% CALL WriteFileSystemPicker_Input(Session("AZ_ID"), FILE_SYSTEM_DIRECTORY, "", "", "form1", "tft_imf_dir", request("tft_imf_dir"), "", false, false) %>
			</td>
		</tr>
		<tr>
			<td class="label">suffisso:</td>
			<td class="content">
				<input type="text" class="text" name="tft_imf_suffisso" value="<%= request("tft_imf_suffisso") %>" maxlength="50" size="10">
			</td>
			<td class="label">suffisso formato:</td>
			<td class="content">
				<input type="checkbox" name="chk_imf_suffissoFormato" value="1" class="checkbox" <%= Chk(request("chk_imf_suffissoFormato") <> "") %>>
			</td>
		</tr>
		<tr>
			<td class="label">salva file originale:</td>
			<td class="content">
				<input type="checkbox" name="chk_imf_salvaOriginale" value="1" class="checkbox" <%= Chk(request("chk_imf_salvaOriginale") <> "") %>>
			</td>
			<td class="note" colspan="2">
				Salva una copia del file caricato dall'utente con lo stesso nome del file generato aggiunto di "_originale".
			</td>
		</tr>
		<tr><th colspan="4">DIMENSIONI</th></tr>
		<tr>
			<td class="label">larghezza:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_imf_width" value="<%= request("tfn_imf_width") %>" maxlength="10" size="4">
			</td>
			<td class="label">altezza:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_imf_height" value="<%= request("tfn_imf_height") %>" maxlength="10" size="4">
			</td>
		</tr>
		<tr>
			<td class="label" style="white-space: nowrap;">dimensioni massime:</td>
			<td class="content">
				<input type="checkbox" name="chk_imf_dimensioniMax" value="1" class="checkbox" <%= Chk(request("chk_imf_dimensioniMax") <> "") %>>
			</td>
			<td class="note" colspan="2">
				Deselezionando il flag le immagini avranno sempre le dimensioni impostate,
				selezionandolo le immagini verrano ridimensionate solamente se superano le dimensioni impostate.
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
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
<% 
conn.close
set rs = nothing
set conn = nothing%>