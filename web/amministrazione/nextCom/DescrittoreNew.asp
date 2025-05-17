<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<%
'controllo accesso
if Session("COM_ADMIN") = "" then
	response.redirect "Documenti.asp"
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("DescrittoreSalva.asp")
end if

dim conn, rs, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Descrittori documenti - nuovo"
'Indirizzo pagina per link su sezione 
		HREF = "Descrittori.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo descrittore</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_descr_nome" value="<%= request("tft_descr_nome") %>" maxlength="50" size="50">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">tipo:</td>
			<td class="content" colspan="3">
				<% CALL DesDropTipi("tfn_descr_tipo", "", request.form("tfn_descr_tipo")) %>
				<span id="nome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">ordine:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_descr_ordine" value="<%= getValueList(conn, rs, "SELECT MAX(descr_ordine)+1 FROM tb_descrittori") %>" maxlength="10" size="5">
			</td>
			<td class="label">principale:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_descr_principale" value="1" <%= chk(request.form("chk_descr_principale")) %>>
			</td>
		</tr>
		<tr><th colspan="4">TIPOLOGIE ASSOCIATE</th></tr>
		<tr>
			<td colspan="4">
				<% sql = "SELECT *, (NULL) AS rtd_descrittore_id FROM tb_tipologie " & _
						 "ORDER BY tipo_nome"
				CALL Write_Relations_Checker(conn, rs, sql, 2, "tipo_id", "tipo_nome", "rtd_descrittore_id", "tipi") %>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva_avanti" value="SALVA">
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