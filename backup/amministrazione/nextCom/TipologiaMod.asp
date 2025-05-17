<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
'controllo accesso
if Session("COM_ADMIN") = "" then
	response.redirect "Documenti.asp"
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("TipologiaSalva.asp")
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
	Titolo_sezione = "Tipologie di documenti - modifica"
'Indirizzo pagina per link su sezione 
		HREF = "Tipologie.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

sql = "SELECT * FROM tb_tipologie WHERE tipo_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Modifica dati della tipologia</caption>
		<tr><th colspan="2">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label" style="width:22%;">nome tipologia:</td>
			<td class="content">
				<input type="text" class="text" name="tft_tipo_nome" value="<%= rs("tipo_nome") %>" maxlength="50" size="50">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<tr><th colspan="2">DESCRITTORI ASSOCIATI</th></tr>
		<tr>
			<td colspan="2">
				<% sql = "SELECT descr_id, descr_nome, rtd_descrittore_id " &_
		  				 "FROM tb_descrittori d LEFT JOIN rel_tipologie_descrittori r ON (" &_
						 "d.descr_id=r.rtd_descrittore_id AND r.rtd_tipologia_id=" & cIntero(request("ID")) & _
		  				 ") ORDER BY descr_nome"
				   CALL Write_Relations_Checker(conn, rsg, sql, 2, "descr_id", "descr_nome", "rtd_descrittore_id", "descr") 
				%>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
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