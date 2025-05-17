<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
'controllo accesso
if Session("COM_ADMIN")="" AND Session("COM_POWER")="" then
	response.redirect "Contatti.asp"
end if

dim conn, rs, rsg, sql, Pager, sql_export

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsg = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("cmp_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("cmp_")
	end if
end if

'recupera rubriche visibili all'utente
dim rubriche_visibili
rubriche_visibili = GetList_Rubriche(conn, rs)

sql = ""
'filtra per nome della rubrica
if session("cmp_nome")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(session("cmp_nome"), SQL_concatFields(conn, "inc_nome")) & ")"
end if

sql_export = sql
sql = " SELECT (SELECT COUNT(*) FROM rel_cnt_campagne WHERE rcc_campagna_id = inc_id) AS N_CONTATTI_COLLEGATI, " & _
	  " * FROM tb_indirizzario_campagne " &_
	  " WHERE (1 = 1) " &_
	  sql & _
	  " ORDER BY inc_nome"
Session("SQL_CAMPAGNE_ELENCO") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>
<%'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Campagne marketing - elenco"
'Indirizzo pagina per link su sezione 
	HREF = "CampagneNew.asp"
'Azione sul link: {BACK | NEW}
	Action = "NUOVA CAMPAGNA MARKETING"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<!--#INCLUDE FILE ="../library/ExportTools.asp" -->
<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
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
								<tr><th <%= Search_Bg("cmp_nome") %>>NOME</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= Server.HTMLEncode(session("cmp_nome")) %>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><td style="font-size:4px;">&nbsp;</td></tr>
								<tr>
									<td>
										<table cellspacing="1" cellpadding="0" class="tabella_madre">
											<caption class="border">Strumenti</caption>
											<%
											sql_export = "SELECT inc_id FROM tb_indirizzario_campagne WHERE (1 = 1) " & sql_export
											
											sql_export = " SELECT IDElencoIndirizzi AS [ID], NomeOrganizzazioneElencoIndirizzi AS [Societa], NomeElencoIndirizzi AS [NOME], SecondoNomeElencoIndirizzi [SECONDO NOME], " & _
														 " CognomeElencoIndirizzi AS [COGNOME], IndirizzoElencoIndirizzi AS [Indirizzo], CittaElencoIndirizzi AS [Citta], " & _
														 " StatoProvElencoIndirizzi AS [STATO / PROV.], CAPElencoIndirizzi AS [CAP], CountryElencoIndirizzi AS [Nazione], " & _
														 " ZonaElencoIndirizzi AS [Zona], partita_iva AS [P.IVA], inc_nome AS [NOME CAMPAGNA], inc_note AS [DESCRIZIONE CAMPAGNA], " & _
														 " (CASE ISDATE(rcc_data_conclusione) WHEN 0 THEN 'no' ELSE 'si' END) AS [CONCLUSA], rcc_data_conclusione AS [DATA CONCLUSIONE], " & _
														 "  ina_note AS [NOTE CONCLUSIONE], (CASE WHEN ina_da_richiamare = 1 THEN 'Da richiamare il' WHEN ina_preso_appuntamento = 1 THEN 'Preso appuntamento il' " & _
														 "  WHEN ina_non_raggiungibili = 1 THEN 'Non raggiungibili' WHEN ina_non_interessati = 1 THEN 'Non interessati' END ) AS CONCLUSIONE, " & _
														 " (CASE WHEN ISDATE(ina_da_richiamare)=1 THEN ina_da_richiamare WHEN ISDATE(ina_data_appuntamento)=1 THEN ina_data_appuntamento END) AS [DATA] " & _
														 " FROM (tb_Indirizzario INNER JOIN rel_cnt_campagne ON tb_Indirizzario.IDElencoIndirizzi = rel_cnt_campagne.rcc_cnt_id " & _
														 " INNER JOIN tb_indirizzario_campagne ON rel_cnt_campagne.rcc_campagna_id = tb_indirizzario_campagne.inc_id) " & _
														 " LEFT OUTER JOIN tb_indirizzario_attivita ON tb_indirizzario_campagne.inc_id = tb_indirizzario_attivita.ina_campagna_conclusa_id " & _
														 " AND tb_Indirizzario.IDElencoIndirizzi = tb_indirizzario_attivita.ina_anagrafica_id " & _
														 " WHERE inc_id IN ("& sql_export &") " & _
														 " ORDER BY NomeOrganizzazioneElencoIndirizzi "
											Session("CONTATTI_CAMPAGNE_EXPORT_SQL") = sql_export
											%>
											<tr>
												<td class="content_center">
													<%
													CALL WRITE_EXPORT_LINK("ESPORTA CAMPAGNE MARKETING", "DATA_ConnectionString", "CONTATTI_CAMPAGNE_EXPORT_SQL", FORMAT_EXCEL_FILE, false)
													%>
												</td>
											</tr>
										</table>
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
					<caption>Elenco campagne marketing</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="4" style="background-color:#b9ddb9;">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="CampagneMod.asp?ID=<%= rs("inc_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<% if cIntero(rs("N_CONTATTI_COLLEGATI")) = 0 then %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CAMPAGNE','<%= rs("inc_id") %>');" >
																	CANCELLA
																</a>
															<% else %>
																<a class="button_disabled" href="javascript:void(0);" title="campagna non cancellabile perch&egrave; sono presenti dei contatti collegati">
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												<%=rs("inc_nome")%>
											</td>
										</tr>
										<tr>	
											<td class="label">n&ordm; contatti</td>
											<td class="content" style="width:25%;"><%= rs("N_CONTATTI_COLLEGATI") %></td>
											<td class="label">numero</td>
											<td class="content"><%= rs("inc_id") %></td>
										</tr>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td class="footer" style="border-top:0px; text-align:left;">
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
set rsg = nothing
set conn = nothing%>
