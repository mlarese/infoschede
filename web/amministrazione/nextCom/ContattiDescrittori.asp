<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%

dim conn, rs, rsg, sql, Pager

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsg = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

sql = "SELECT * FROM tb_indirizzario_carattech ORDER BY ict_nome_it "
Session("NEXTCOM_CONTATTI_CTECH_SQL") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>
<%'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Descrittori contatti - elenco"
'Indirizzo pagina per link su sezione 
		HREF = "ContattiDescrittoreNew.asp"
'Azione sul link: {BACK | NEW}
	Action = "NUOVO DESCRITTORE"
	if Session("COM_ADMIN") <> "" then
		SSezioniText = "CATEGORIE;RAGGRUPPAMENTI"
		SSezioniLink = "ContattiCategorie.asp;ContattiDescrittoriGruppi.asp"
	end if
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco descrittori - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<% if not rs.eof then %>
			<tr>
				<th style="text-align:center; width:5%;">ID</th>
				<th>DESCRITTORE</th>
				<th>TIPOLOGIE</th>
				<th class="center" style="width:10%">RICERCABILE</th>
				<th class="center" style="width:10%">CONFRONTABILE</th>
				<th colspan="2" style="text-align:center; width:20%;">OPERAZIONI</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
				<tr>
					<td class="content_center"><%= rs("ict_id") %></td>
					<td class="content"><%= rs("ict_nome_it") %></td>
					<% sql = "SELECT icat_nome_it FROM tb_indirizzario_categorie t INNER JOIN rel_categ_ctech r "& _
							 "ON t.icat_id=r.rcc_categoria_id WHERE rcc_ctech_id="& rs("ict_id") %>
					<td class="content"><%= GetValueList(conn, rsg, sql) %>
					<td class="content_center"><input disabled class="checkbox" type="checkbox" name="chk_<%= rs("ict_id") %>" value="1" <%= chk(rs("ict_per_ricerca")) %>></td>
					<td class="content_center"><input disabled class="checkbox" type="checkbox" name="chk_<%= rs("ict_id") %>" value="1" <%= chk(rs("ict_per_confronto")) %>></td>
					</td>
					<td class="Content_center">
						<a class="button" href="ContattiDescrittoreMod.asp?ID=<%= rs("ict_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="Content_center">
						<% if CInt(GetValueList(conn, rsg, "SELECT COUNT(*) FROM rel_cnt_ctech WHERE ric_ctech_id="& rs("ict_id"))) = 0 then %>
							<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ContattiCTECH','<%= rs("ict_id") %>');" >
								CANCELLA
							</a>
						<% else %>
							<a class="button_disabled" href="javascript:void(0);" title="descrittore non cancellabile perch&egrave; esistono contatti associati">
								CANCELLA
							</a>
						<% end if %>
					</td>
				</tr>
				<% rs.moveNext
			wend%>
			<tr>
				<td colspan="7" class="footer" style="text-align:left;">
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
