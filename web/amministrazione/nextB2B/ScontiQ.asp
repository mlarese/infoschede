<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione classi di sconto per quantit&agrave; - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA CLASSE"
dicitura.link_new = "Tabelle.asp;ScontiQNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, rsv, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * ,(SELECT COUNT(*) FROM gtb_prezzi WHERE prz_scontoQ_id=gtb_scontiQ_classi.scc_id) AS N_PREZZI, " + _
	  " (SELECT COUNT(*) FROM grel_art_valori WHERE rel_scontoQ_id=gtb_scontiQ_classi.scc_id) AS N_REL, " + _
	  " (SELECT COUNT(*) FROM gtb_articoli WHERE art_scontoQ_id = gtb_scontiQ_classi.scc_id) AS N_ART " + _
	  " FROM gtb_scontiQ_classi ORDER BY scc_nome"
session("B2B_SCONTIQ_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco classi di sconto per quantit&agrave; - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" style="width:5%;">ID</th>
			<th>NOME</th>
			<th class="center" colspan="2" style="width:20%;">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("scc_id") %></td>
				<td class="content"><%= rs("scc_nome") %></td>
				<td class="Content_center">
					<a class="button" href="ScontiQMod.asp?ID=<%= rs("scc_id") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<% if rs("N_PREZZI")>0 OR rs("N_ART")>0 OR rs("N_REL")>0 then%>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la classe perch&egrave; associata ad almeno un articolo." <%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('SCONTIQ','<%= rs("scc_id") %>');" >
							CANCELLA
						</a>
					<% end if %>
				</td>
			</tr>
			<% rs.movenext
		wend%>
	</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsv = nothing
set conn = nothing%>