<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>
<% Titolo_sezione = "Import dati dei contatti da file in formato NEXT-com"
Action = "INDIETRO"
href = "default.asp"%>
<!--#include file="Intestazione.asp"-->

<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Import tutti i dati</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		<% if request("importa")="" OR request("conn_import")="" then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" rowspan="2" style="width:18%;">connessione da cui importare:</td>
				<td class="content" colspan="2">
					<input type="text" name="conn_import" style="width:90%;" value="" /><br>
                    <span class="note">(*) scrivi la connessione al database di orgine.</span>
				</td>
			</tr>
			<tr>
				<td class="note">Provider=Microsoft.Jet.OLEDB.4.0;Data Source=&lt;database path&gt;;</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA CONTATTI">
				</td>
			</tr>
			</form>
		<% else 
				
			dim sconn, dconn, rss, rsd, sql
			set dconn = Server.CreateObject("ADODB.Connection")
			set sconn = Server.CreateObject("ADODB.Connection")
			
			dconn.open Application("DATA_ConnectionString")
			sconn.open request("conn_import")
			
			set rss = Server.CreateObject("ADODB.RecordSet")
			set rsd = Server.CreateObject("ADODB.RecordSet")
			
			CALL CopyTable(sconn, dconn, "SELECT * FROM tb_indirizzario", "SELECT * FROM tb_indirizzario", "IdeLEncoindirizzi")
			%>	
			
		<% end if %>
	</table>
</div>

<%
conn.close
set rs = nothing
set rsr = nothing
set rsv = nothing
set conn = nothing
%>