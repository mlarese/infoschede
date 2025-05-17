<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione modalit&agrave; di spedizione dell'articolo"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA MODALIT&Agrave;"
dicitura.link_new = "Tabelle.asp;SpeseSpedizioneArticoloNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql, pager, i, disabled

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("spa_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("spa_")
	end if
end if

sql = ""

'filtra per codice
if Session("spa_nome")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("spa_nome"), FieldLanguageList("spa_nome_")) & ") "
end if


sql = "SELECT (SELECT COUNT(*) FROM gtb_articoli WHERE art_spedizione_id = spa_id) AS NUM_ART, " & _
	  " * FROM gtb_spese_spedizione_articolo " + _
	  " WHERE (1=1) " + sql + " ORDER BY spa_nome_it"
session("B2B_SPA_SQL") = sql
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
									<tr><th <%= Search_Bg("spa_nome") %>>NOME</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome" value="<%= TextEncode(session("spa_nome")) %>" style="width:100%;">
										</td>
									</tr>
								
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
					<caption>Modalit&agrave; Spedizione - Trovati n&ordm; <%= Pager.recordcount %> metodi in n&ordm; <%= Pager.PageCount %> pagine</caption>
					<% if not rs.eof then %>
						<tr>
							<th>MODALITA' SPEDIZIONE ARTICOLO</th>
							<th class="center" style="width: 25%;">SPESE SPEDIZIONE</th>
							<th class="center" colspan="2" style="width: 20%;">OPERAZIONI</th>
						</tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="content"><%= rs("spa_nome_it") %></td>
								<td class="content"><%= FormatPrice(rs("spa_importo_spese"), 2, true) %> &euro;</td>
								<td style="vertical-align:middle;" class="Content_center">
									<a class="button" href="SpeseSpedizioneArticoloMod.asp?ID=<%= rs("spa_id") %>">
										MODIFICA
									</a>
								</td>
								<td style="vertical-align:middle;" class="Content_center">
									<% if cIntero(rs("NUM_ART")) > 0 then %>
										<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare l'area ordini già consegnati">
											CANCELLA
										</a>
									<% else %>
										<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('SPESESPEDIZIONEARTICOLO','<%= rs("spa_id") %>');" >
											CANCELLA
										</a>
									<% end if %>
								</td>
							</tr>
							<%rs.movenext
						wend%>
						<tr>
							<td class="footer" style="text-align:left;" colspan="6">
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