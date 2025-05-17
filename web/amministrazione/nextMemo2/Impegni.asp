<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->

<% 	
dim dicitura
set dicitura = New testata

dicitura.iniz_sottosez(3)
dicitura.sottosezioni(1) = "TIPOLOGIE"
dicitura.links(1) = "ImpegniTipologie.asp"
dicitura.sottosezioni(2) = "CONFIGURAZIONE"
dicitura.links(2) = "AgendaConfigura.asp"
dicitura.sottosezioni(3) = "LOG AVVISI"
dicitura.links(3) = "ImpegniLogAvvisi.asp"

dicitura.sezione = "Gestione impegni/appuntamenti - elenco"
dicitura.puls_new = "NUOVO IMPEGNO"
dicitura.link_new = "ImpegniNew.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, Pager

set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


%>
<!--#INCLUDE FILE="ImpegniFiltriRicerca.asp" -->
<%

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)

dim profili_attivi
sql = "SELECT pro_id FROM mtb_profili"
if cString(GetValueList(conn, NULL, sql)) <> "" then
	profili_attivi = true
else
	profili_attivi = false
end if

%>

<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
				<% CALL WriteBloccoRicerca(conn,"vertical") %>
			</td>
			
			<!-- BLOCCO DEI RISULTATI -->
			<td width="1%">&nbsp;</td>
			<td valign="top">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
					<caption class="border">
						Lista impegni/appuntamenti - calendario
					</caption>
					<tr>
						<td class="content">
							Visualizza gli impegni e gli appuntamenti nel calendario
						</td>
						<td class="content_right">
							<a class="button" href="ImpegniCalendarioView.asp?FIRSTDATE=<%=IIF(Session("imp_data_inizio")<>"",Session("imp_data_inizio"),Date())%>" title="Apre la visualizzazione del calendario.">
								VISUALIZZA CALENDARIO
							</a>
						</td>
					</tr>
				</table>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco impegni/appuntamenti
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>										
											<%
											dim header 
											if DateISO(rs("imp_data_ora_fine")) < DateIso(Now()) then
												header = "header_disabled"
											else
												header = "header"
											end if
											%>
											<td class="<%=header%>" colspan="7">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<% CALL index.WriteButton("mtb_impegni", rs("imp_id"), POS_ELENCO) %>
															<a class="button" href="ImpegniMod.asp?ID=<%= rs("imp_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('IMPEGNI','<%= rs("imp_id") %>');" >
																CANCELLA
															</a>
														</td>
													</tr>
												</table>
												<%= rs("imp_titolo_it") %>
											</td>
										</tr>
										<tr>
											<td class="label">tipologia:</td>
											<td class="content" colspan="3">
												<% sql = "SELECT tim_colore FROM mtb_tipi_impegni WHERE tim_id = " & rs("imp_tipo_id") %>
												<% if GetValueList(conn,NULL,sql) <> "" then %>
													<% WriteColor(GetValueList(conn,NULL,sql))%>
													<% sql = "SELECT tim_nome_it FROM mtb_tipi_impegni WHERE tim_id = " & rs("imp_tipo_id") %>
													<%= GetValueList(conn,NULL,sql) %>
												<% else %>
													<span class="note">tipologia non impostata</span>
												<% end if %>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:20%;">orario inizio:</td>
											<td class="content" style="width:28%;"><%=TimeIta(rs("imp_data_ora_inizio"))%></td>
											<td class="label_right">orario fine:</td>
											<td class="content"><%=TimeIta(rs("imp_data_ora_fine"))%></td>
										</tr>
										<tr>
											<td class="label" style="width:20%;">attivo dal:</td>
											<td class="content" style="width:28%;"><%=DateIta(rs("imp_data_ora_inizio"))%></td>
											<td class="label_right">scadenza:</td>
											<% if Trim(DateIta(rs("imp_data_ora_fine"))) = Trim(DateIta(DATA_SENZA_FINE)) then %>
												<td class="content">&nbsp;-</td>
											<% else %>
												<td class="content"><%=DateIta(rs("imp_data_ora_fine"))%></td>
											<% end if %>
										</tr>
										<tr>
											<td class="label" style="width:20%;">visibile:</td>
											<td class="content <%= IIF(rs("imp_protetto"), " OrdConfermato", " OrdEvaso")%>" colspan="3">
												<% if rs("imp_protetto") then %>
													<img src="../grafica/padlock.gif" border="0" alt="Pagina appartenente all'area protetta">
													solo da chi &egrave; associato all'impegno/appuntamento
												<% else %>
													da tutti
												<% end if %>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend %>
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
set conn = nothing
%>