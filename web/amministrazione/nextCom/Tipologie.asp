<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->

<%
'controllo accesso
if Session("COM_ADMIN") = "" then
	response.redirect "Pratiche.asp"
end if

dim conn, rs, rsg, sql, Pager

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsg = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

sql = "SELECT * FROM tb_tipologie ORDER BY tipo_nome"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>
<%'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Tipologie di documenti - elenco"
'Indirizzo pagina per link su sezione 
		HREF = "TipologiaNew.asp"
'Azione sul link: {BACK | NEW}
	Action = "NUOVA TIPOLOGIA"
	if Session("COM_ADMIN") <> "" then
		SSezioniText = "DOCUMENTI;TIPOLOGIE;DESCRITTORI"
		SSezioniLink = "documenti.asp;tipologie.asp;descrittori.asp"
	end if
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco tipologie - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<% if not rs.eof then %>
			<tr>
				<th style="text-align:center; width:5%;">ID</th>
				<th>TIPOLOGIA</th>
				<th>DESCRITTORI</th>
				<th colspan="2" style="text-align:center; width:20%;">OPERAZIONI</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
				<tr>
					<td class="content_center"><%= rs("tipo_id") %></td>
					<td class="content"><%= rs("tipo_nome") %></td>
					<% sql = "SELECT descr_nome FROM tb_descrittori d INNER JOIN rel_tipologie_descrittori r "& _
							 "ON d.descr_id=r.rtd_descrittore_id WHERE rtd_tipologia_id="& rs("tipo_id") %>
					<td class="content"><%= GetValueList(conn, rsg, sql) %>
					</td>
					<td class="Content_center">
						<a class="button" href="TipologiaMod.asp?ID=<%= rs("tipo_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="Content_center">
						<% if CInt(GetValueList(conn, rsg, "SELECT COUNT(*) FROM tb_documenti WHERE doc_tipologia_id="& rs("tipo_id"))) = 0 then %>
							<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('TIPI','<%= rs("tipo_id") %>');" >
								CANCELLA
							</a>
						<% else %>
							<a class="button_disabled" href="javascript:void(0);" title="tipologia non cancellabile perch&egrave; esistono documenti associati">
								CANCELLA
							</a>
						<% end if %>
					</td>
				</tr>
				<% rs.moveNext
			wend%>
			<tr>
				<td colspan="5" class="footer" style="text-align:left;">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
				</td>
			</tr>
		<%else%>
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
set rsg = nothing
set conn = nothing%>
