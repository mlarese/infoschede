<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("GruppiSalva.asp")
end if

dim dicitura
set dicitura = New testata  
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gruppi di lavoro - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Gruppi.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo gruppo</caption>
		<tr><th colspan="2">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">nome gruppo:</td>
			<td class="content">
				<input type="text" class="text" name="tft_nome_gruppo" value="<%= request("tft_nome_gruppo") %>" maxlength="250" size="75">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr><th colspan="2">UTENTI CHE COMPONGONO IL GRUPPO (*)</th></tr>
		<tr>
			<td colspan="2">
				<% sql = "SELECT ID_admin, (admin_Cognome " & SQL_concat(conn) & " ' ' " & SQL_concat(conn) & " admin_Nome) AS NOME, (NULL) AS id_rel_dipgruppi FROM tb_admin " & _
						 " ORDER BY admin_Cognome"
				CALL Write_Relations_Checker(conn, rs, sql, 2, "id_admin", "NOME", "id_rel_dipgruppi", "utenti") %>
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
conn.close 
set rs = nothing
set conn = nothing%>