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
dicitura.sezione = "Gruppi di lavoro - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Gruppi.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, rsu, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsu = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_gruppi WHERE id_gruppo=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Modifica dati del gruppo</caption>
		<tr><th colspan="2">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">nome gruppo:</td>
			<td class="content">
				<input type="text" class="text" name="tft_nome_gruppo" value="<%= rs("nome_gruppo") %>" maxlength="250" size="75">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr><th colspan="2">UTENTI CHE COMPONGONO IL GRUPPO (*)</th></tr>
		<tr>
			<td colspan="2">
				<% sql = "SELECT id_admin, (admin_Cognome " & SQL_concat(conn) & " ' ' " & SQL_concat(conn) & " admin_Nome) AS NOME, id_rel_dipgruppi " &_
						 " FROM tb_admin LEFT JOIN tb_rel_dipgruppi ON (tb_admin.ID_admin = tb_rel_dipgruppi.id_impiegato " &_
		  				 " AND tb_rel_dipgruppi.id_Gruppo=" & cIntero(request("ID")) & ")" & _
						 " ORDER BY admin_Cognome"
				CALL Write_Relations_Checker(conn, rsu, sql, 2, "id_admin", "NOME", "id_rel_dipgruppi", "utenti") %>
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
set rsu = nothing
set rs = nothing
set conn = nothing%>