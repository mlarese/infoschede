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
CALL CheckAutentication(index.ChkPrm(prm_menu_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - menu - elenco"
dicitura.puls_new = "INDIETRO A SITI;NUOVO MENU"
dicitura.link_new = "Siti.asp;SitoMenuNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_menu WHERE m_id_webs = "& session("AZ_ID") &" ORDER BY m_nome_it"
session("WEB_MENU_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco menu - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<% if not rs.eof then %>
			<tr>
				<th class="center" width="3%">ID</th>
				<th>NOME</th>
				<th class="center" colspan="2" style="width:19%;">OPERAZIONI</th>
			</tr>
			<% while not rs.eof %>
				<tr>
					<td class="content_center"><%= rs("m_id") %></td>
					<td class="content"><%= rs("m_nome_it") %></td>
					<td style="vertical-align:middle;" class="Content_center">
						<a class="button" href="SitoMenuMod.asp?ID=<%= rs("m_id") %>">
							MODIFICA
						</a>
					</td>
					<td style="vertical-align:middle;" class="Content_center">
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('MENU','<%= rs("m_id") %>');" >
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
