<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
Imposta_Proprieta_Sito("ID")
%>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_immaginiFormati_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(1)
dicitura.sezione = "Gestione files - formati immagini - elenco"
dicitura.puls_new = "INDIETRO A SITI;NUOVO FORMATO"
dicitura.link_new = "Siti.asp;SitoImmaginiFormatiNew.asp"
dicitura.sottosezioni(1) = "ELENCO FILES"
dicitura.links(1) = "SitoFileManager.asp"
dicitura.scrivi_con_sottosez()

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_immaginiFormati WHERE imf_webId = "& session("AZ_ID") &" ORDER BY imf_nome"
session("WEB_IMMAGINIFORMATI_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco formati delle immagini - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<% if not rs.eof then %>
			<tr>
				<th class="center" width="3%" rowspan="2">ID</th>
				<th rowspan="2">NOME</th>
				<th class="center" style="width: 12%;" rowspan="2">SALVA FILE ORIGINALE</th>
				<th class="center" style="width: 25%; border-bottom: 0px;" colspan="3">DIMENSIONI</th>
				<th class="center" colspan="2" style="width:19%;" rowspan="2">OPERAZIONI</th>
			</tr>
			<tr>
				<th style="text-align: right;">larghezza</th>
				<th>X</th>
				<th>altezza</th>
			</tr>
			<% while not rs.eof %>
				<tr>
					<td class="content_center"><%= rs("imf_id") %></td>
					<td class="content"><%= rs("imf_nome") %></td>
					<td class="content_center"><input type="checkbox" <%= chk(rs("imf_salvaOriginale")) %> disabled></td>
					<td class="content_right"><%= IIF(rs("imf_width") = 0, IIF(rs("imf_height") = 0, "invariata", "proporzionata"), rs("imf_width")) %></td>
					<td class="content_center">X</td>
					<td class="content"><%= IIF(rs("imf_height") = 0, IIF(rs("imf_width") = 0, "invariata", "proporzionata"), rs("imf_height")) %></td>
					<td style="vertical-align:middle;" class="Content_center">
						<a class="button" href="SitoImmaginiFormatiMod.asp?ID=<%= rs("imf_id") %>">
							MODIFICA
						</a>
					</td>
					<td style="vertical-align:middle;" class="Content_center">
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('IMMAGINIFORMATI','<%= rs("imf_id") %>');" >
							CANCELLA
						</a>
					</td>
				</tr>
				<%rs.movenext
			wend
		else%>
			<tr><td class="noRecords">Nessun record trovato</th></tr>
		<% end if %>
	</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing%>
