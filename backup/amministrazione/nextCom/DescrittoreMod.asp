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

dim conn, rs, rsg, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsg = Server.CreateObject("ADODB.RecordSet")

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Descrittori documenti - modifica"
'Indirizzo pagina per link su sezione 
		HREF = "Descrittori.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

sql = "SELECT * FROM tb_descrittori WHERE descr_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Modifica dati del descrittore</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_descr_nome" value="<%= rs("descr_nome") %>" maxlength="50" size="50">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">tipo:</td>
			<td class="content" colspan="3">
				<% CALL DesDropTipi("tfn_descr_tipo", "", rs("descr_tipo")) %>
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">ordine:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_descr_ordine" value="<%= rs("descr_ordine") %>" maxlength="10" size="5">
			</td>
			<td class="label">principale:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_descr_principale" value="1" <%= chk(rs("descr_principale")) %>>
			</td>
		</tr>
		<tr><th colspan="4">TIPOLOGIE ASSOCIATE</th></tr>
		<tr>
			<td colspan="4">
				<% sql = "SELECT tipo_id, tipo_nome, rtd_tipologia_id " &_
		  				 "FROM tb_tipologie t LEFT JOIN rel_tipologie_descrittori r ON (" &_
						 "t.tipo_id=r.rtd_tipologia_id AND r.rtd_descrittore_id=" & cIntero(request("ID")) & _
		  				 ") ORDER BY tipo_nome"
				   CALL Write_Relations_Checker(conn, rsg, sql, 2, "tipo_id", "tipo_nome", "rtd_tipologia_id", "tipi") 
				%>
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
rs.close
conn.close 
set rsg = nothing
set rs = nothing
set conn = nothing%>