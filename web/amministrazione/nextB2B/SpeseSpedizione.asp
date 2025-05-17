<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione spese di spedizione / modalit&agrave; spedizione ordine"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA MODALIT&Agrave;"
dicitura.link_new = "Tabelle.asp;SpeseSpedizioneNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql, pager, i, disabled

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("sp_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("sp_")
	end if
end if

sql = ""

'filtra per codice
if Session("sp_area_nome")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("sp_area_nome"), FieldLanguageList("sp_area_nome_")) & ") "
end if

'filtra per codice
if Session("sp_condizioni")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("sp_condizioni"), FieldLanguageList("sp_condizioni_")) & ") "
end if

sql = "SELECT * FROM gtb_spese_spedizione LEFT JOIN gtb_iva ON gtb_spese_spedizione.sp_iva_id = gtb_iva.iva_id " + _
	  " WHERE (1=1) " + sql + " ORDER BY sp_area_nome_it"
session("B2B_SP_SQL") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
				<form action="" method="post" id="ricerca" name="ricerca">
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
									<caption>Opzioni di ricerca</caption>
									<tr>
										<td class="footer">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("sp_area_nome") %>>NOME</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome" value="<%= TextEncode(session("sp_area_nome")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("sp_condizioni") %>>DESCRIZONE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_descrizione" value="<%= TextEncode(session("sp_condizioni")) %>" style="width:100%;">
										</td>
									</tr>
									<tr>
										<td class="footer">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
				<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>Elenco Modalit&agrave; Spedizione - Trovati n&ordm; <%= Pager.recordcount %> metodi in n&ordm; <%= Pager.PageCount %> pagine</caption>
					<% if not rs.eof then %>
						<tr>
							<th>MODALIT&Agrave; SPEDIZIONE ORDINE</th>
							<th class="center" style="width:14%;">COSTO</th>
							<th class="center" style="width:10%;">PERC.</th>
							<th class="center" style="width:15%;">COSTO MIN.</th>
							<th class="center" style="width:8%;">I.V.A.</th>
							<th class="center" colspan="2" style="width: 20%;">OPERAZIONI</th>
						</tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="content"><%= rs("sp_area_nome_it") %></td>
								<td class="content_center">
									<% if cReal(rs("sp_percentuale"))=0 then %>
										<%= FormatPrice(rs("sp_importo_euro"), 2, true) %> &euro;
									<% else %>
										&nbsp;
									<% end if %>
								</td>
								<td class="content_center">
									<% if cReal(rs("sp_percentuale"))>0 then %>
										<%= rs("sp_percentuale") %> %
									<% else %>
										&nbsp;
									<% end if %>
								</td>
								<td class="content_center">
									<% if cReal(rs("sp_percentuale"))>0 then %>
										<%= FormatPrice(rs("sp_importo_euro"), 2, true) %> &euro;
									<% else %>
										&nbsp;
									<% end if %>
								</td>
								<td class="content_center">
									<% if cReal(rs("iva_valore"))>0 then %>
										<%= FormatPrice(rs("iva_valore"), 0, true) %>%
									<% else %>
										&nbsp;
									<% end if %>
								</td>
								<td style="vertical-align:middle;" class="Content_center">
									<a class="button" href="SpeseSpedizioneMod.asp?ID=<%= rs("sp_id") %>">
										MODIFICA
									</a>
								</td>
								<td style="vertical-align:middle;" class="Content_center">
									<% disabled = false ' To Do Correggere appena inserita relazione con ordini
									if disabled then %>
										<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare l'area ordini già consegnati">
											CANCELLA
										</a>
									<% else %>
										<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('SPESESPEDIZIONE','<%= rs("sp_id") %>');" >
											CANCELLA
										</a>
									<% end if %>
								</td>
							</tr>
							<%rs.movenext
						wend%>
						<tr>
							<td class="footer" style="text-align:left;" colspan="7">
								<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
							</td>
						</tr>
					<%else%>
						<tr><td class="noRecords">Nessun record trovato</th></tr>
					<% end if %>
				</table>
			</td> 
		</tr>
		<tr><td>&nbsp;</td></tr>
	</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing%>