<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Titolo_sezione = "Import dati dei contatti in formato vCard"
Action = "INDIETRO"
href = "default.asp"%>
<!--#include file="Intestazione.asp"-->
<% 
dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Import dati dei contatti in formato vCard</caption>
		<tr><th colspan="2">PARAMETRI DI IMPORT</th></tr>
		<% if request("importa")="" OR _
			  request("rubrica_import")="" OR _
			  request("cartella_import")="" then %>
			<form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:18%;">cartella con i file vCard:</td>
				<td class="content">
					<% CALL WriteFilePicker_Input(0, "", "form1", "cartella_import", request("cartella_import") , "width:400px;", true) %>
					<span class="note">Selezionare la cartella che contiene i file vCard.</span>
				</td>
			</tr>
			<tr>
				<td class="label">rubrica di destinazione:</td>
				<td class="content">
					<% sql = " SELECT id_rubrica, nome_rubrica FROM tb_rubriche " &_
							 " ORDER BY nome_rubrica"
					CALL dropDown(conn, sql, "id_rubrica", "nome_rubrica", "rubrica_import", request("rubrica_import"), true, "", LINGUA_ITALIANO)%>(*)<br>
					<span class="note">Selezionare la rubrica nella quale verranno inseriti i contatti.</span>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="2">
					(*) Campi obbligatori.
					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA CONTATTI">
				</td>
			</tr>
			</form>
		<% else %>
			<tr>
				<td class="label" style="width:18%;">cartella con i file vCard:</td>
				<td class="content">
					<%= request("cartella_import") %>
				</td>
			</tr>
			<tr>
				<td class="label">rubrica di destinazione:</td>
				<td class="content">
					<% sql = "SELECT nome_rubrica FROM tb_rubriche WHERE id_rubrica=" & cIntero(request("rubrica_import")) %>
					<%= GetValueList(conn, rs, sql) %>
				</td>
			</tr>
		<% end if %>
	</table>
</div>

<%
conn.close
set rs = nothing
set conn = nothing
%>