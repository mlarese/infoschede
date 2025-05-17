<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/ExportTools.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
if request("ID")="" then 
	dicitura.iniz_sottosez(1)
	dicitura.sottosezioni(1) = "ELENCO DOCUMENTI"
	dicitura.links(1) = "Documenti.asp"
else
	dicitura.iniz_sottosez(0)
end if
dicitura.sezione = "Gestione documenti - log download"
if request("ID")="" then
	dicitura.puls_new = ""
	dicitura.link_new = ""
else
	dicitura.puls_new = "INDIETRO;DATI DOCUMENTO"
	dicitura.link_new = "Documenti.asp;DocumentiMod.asp?ID=" & request("ID")
end if
dicitura.scrivi_con_sottosez() 



dim i, conn, rsd, rs, sql, Pager

set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	Pager.Reset
	CALL GotoRecord(conn, rs, Session("SQL_DOCUMENTI"), "doc_id", "DocumentiLog.asp")
end if


sql = " FROM (((log_documenti LEFT JOIN tb_admin ON log_documenti.log_dip_id = tb_admin.id_admin) " &_
	  " LEFT JOIN tb_Utenti ON log_documenti.log_ut_id=tb_utenti.ut_id) " &_
	  " LEFT JOIN tb_Indirizzario ON tb_utenti.ut_nextCom_id=tb_indirizzario.IDElencoIndirizzi) " &_
  	  " INNER JOIN  mtb_documenti ON log_documenti.log_doc_id = mtb_documenti.doc_id " & _
	  " LEFT JOIN mtb_documenti_categorie ON mtb_documenti.doc_categoria_id = mtb_documenti_categorie.catc_id " & _
	  " LEFT JOIN mtb_documenti_categorie catp ON mtb_documenti_categorie.catC_tipologia_padre_base = catp.catc_id "
if request("ID")<>"" then
	sql = sql & " WHERE doc_id=" & cIntero(request("ID"))
end if
sql = sql & " ORDER BY log_data DESC "

Session("EXPORT_MEMO_DOWNLOAD" & request("ID")) = " SELECT log_data AS [Data download], " & _
												  " (CASE WHEN isnull(log_dip_id,0)<>0 THEN admin_login ELSE Isnull(NomeOrganizzazioneElencoIndirizzi,'') + ' - ' + isnull(CognomeElencoIndirizzi,'') + ' ' + IsNull(NomeElencoIndirizzi,'') END) as Utente," & _
												  " (doc_titolo_it) As [Documento]," & _
												  " (mtb_documenti_categorie.catc_nome_it) AS [Categoria], " & _
												  " (catp.catc_nome_it) AS [Categoria base]" & _
												  sql
												  
CALL Pager.OpenSmartRecordset(conn, rs, "SELECT * " & sql, 25)
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% if request("ID")<>"" then
			sql = "SELECT * FROM mtb_documenti WHERE doc_id=" & cIntero(request("ID"))
			rsd.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
			<caption>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td class="caption">Log dei download del documento &quot;<%= rsd("doc_titolo_it") %>&quot;</td>
						<td align="right" style="font-size: 1px;">
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="documento precedente">
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="documento successivo">
								SUCCESSIVO &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			</caption>
			<%rsd.close%>
		<% else %>
			<caption>
				<table width="100%">
				<tr>
					<td>
					Log di download di tutti i documenti
					<% if not rs.eof then %>
						- eseguiti n&ordm; <%= Pager.recordcount %> download
					<% end if %>
					</td>
					<td align="right">
						<%CALL WRITE_EXPORT_LINK_ADV("ESPORTA IN EXCEL", "DATA_ConnectionString", "EXPORT_MEMO_DOWNLOAD" & request("ID"), FORMAT_EXCEL_XML, true, "../library/") %>
					</td>
				</tr>
				</table>
			</caption>
		<%end if
		if not rs.eof then%>
			<tr>
				<th class="center" width="20%">DATA</th>
				<th>UTENTE</th>
				<th class="center" width="10%">AREA</th>
				<th width="40%">DOCUMENTO</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
				<tr>
					<td class="content_center"><%= DateTimeIta(rs("log_data")) %></td>
					<% if cIntero(rs("log_dip_id"))>0 then %>
						<td class="content"><%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %></td>
						<td class="content_center">AMMINISTRAZIONE</td>
					<% elseif cIntero(rs("log_ut_id"))>0 then%>
						<td class="content"><%= ContactFullName(rs) %></td>
						<td class="content_center">RISERVATA</td>
					<% else %>
						<td class="content">anonimo</td>
						<td class="content_center">PUBBLICA</td>
					<% end if %>
					<td class="content"><%= rs("doc_titolo_it") %></td>
				</tr>
				<%rs.movenext
			wend%>
			
			<tr>
				<td colspan="4" class="footer" style="text-align:left;">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
				</td>
			</tr>
		<% else %>
			<tr><td class="noRecords">Nessun download effettuato</th></tr>
		<% end if %>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsd = nothing
set conn = nothing
%>