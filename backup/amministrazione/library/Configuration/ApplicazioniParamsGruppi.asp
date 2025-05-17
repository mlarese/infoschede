<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../../nextPassport/ToolsApplicazioni.asp" -->
<%

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(3)
dicitura.sottosezioni(1) = "APPLICAZIONI"
dicitura.links(1) = "Applicazioni.asp"
dicitura.sottosezioni(2) = "PARAMETRI"
dicitura.links(2) = "ApplicazioniParams.asp"
dicitura.sottosezioni(3) = "GRUPPI DI PARAMETRI"
dicitura.links(3) = "ApplicazioniParamsGruppi.asp"
dicitura.puls_new = "NUOVO GRUPPO"
dicitura.link_new = "ApplicazioniParamsGruppiNew.asp"
dicitura.sezione = "Gestione gruppi di parametri - elenco"
dicitura.scrivi_con_sottosez()


dim conn, rs, sql, Pager
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open GetConfigurationConnectionstring()
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM tb_siti_descrittori_raggruppamenti"& _
	  " ORDER BY sdr_titolo_it"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco gruppi parametri applicazioni - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<% if not rs.eof then %>
			<tr>
				<th class="center" style="width:3%;">ID</th>
				<th>NOME</th>
				<th class="center" width="8%">ORDINE</th>
				<th class="center" colspan="2" style="width:21%;">OPERAZIONI</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
				<tr>
					<td class="content_center"><%= rs("sdr_id") %></td>
					<td class="Content">
						<%= rs("sdr_titolo_IT") %>
					</td>
					<td class="Content_center"><%= rs("sdr_ordine") %></td>
					<td class="Content_center">
						<a class="button" href="ApplicazioniParamsGruppiMod.asp?ID=<%= rs("sdr_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="Content_center">
					<% 		if CInt(GetValueList(conn, NULL, "SELECT COUNT(*) FROM tb_siti_descrittori WHERE sid_raggruppamento_id="& rs("sdr_id"))) = 0 then %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('APPLICAZIONI_PARAMS_RAG','<%= rs("sdr_id") %>');" >
							CANCELLA
						</a>
					<% 		else %>
						<a class="button_disabled" title="Impossibile cancellare il gruppo: sono presenti parametri associati">
							CANCELLA
						</a>
					<% 		end if %>
					</td>
				</tr>
				<% rs.moveNext
			wend%>
			<tr>
				<td colspan="6" class="footer" style="text-align:left;">
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
set conn = nothing%>

	
