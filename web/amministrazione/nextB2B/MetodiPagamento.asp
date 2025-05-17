<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione modalit&agrave; di pagamento - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA MODALIT&Agrave;"
dicitura.link_new = "Tabelle.asp;MetodiPagamentoNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql, pager, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("modpaga_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("modpaga_")
	end if
end if

sql = ""
'filtra per nome della modalità
if Session("modpaga_nome")<>"" then
	sql = sql & " AND "
	sql = sql & SQL_FullTextSearch(Session("modpaga_nome"), FieldLanguageList("mosp_nome_"))
end if

'filtra per codice della modalità
if Session("modpaga_codice")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("modpaga_codice"), "mosp_codice")
end if

sql = "SELECT *, (SELECT COUNT(*) FROM gtb_ordini WHERE ord_modopagamento_id=mosp_id) AS N_ORDINI FROM gtb_modipagamento " + _
	  " WHERE (1=1) " + sql + " ORDER BY mosp_nome_it"
session("B2B_MODPAGA_SQL") = sql

'response.write sql
'response.end
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
									<tr><th <%= Search_Bg("modpaga_nome") %>>NOME</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome" value="<%= TextEncode(session("modpaga_nome")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("modpaga_codice") %>>CODICE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_codice" value="<%= TextEncode(session("modpaga_codice")) %>" style="width:100%;">
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
				<!--
				<%= sql %>
				-->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>Elenco modalit&agrave; - Trovati n&ordm; <%= Pager.recordcount %> modalit&agrave; in n&ordm; <%= Pager.PageCount %> pagine</caption>
					<% if not rs.eof then %>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="<%= IIF(rs("mosp_se_abilitato"), "header", "header_disabled") %>" colspan="4">
												<%= rs("mosp_nome_it") %>
											</td>
											<td class="header" colspan="2">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr colspan="2">
														<td style="font-size: 1px; text-align:right;">
															<a class="button" href="MetodiPagamentoMod.asp?ID=<%= rs("mosp_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<% if rs("N_ORDINI") > 0 then %>
																<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la modalit&agrave;: sono presenti ordini associati">
																	CANCELLA
																</a>
															<% else %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('MODPAGA','<%= rs("mosp_id") %>');" >
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												
											</td>
										</tr>
										<tr>
											<td class="label">codice:</td>
											<td class="content" colspan="5">
												<%= rs("mosp_codice") %>
											</td>
										</tr>
										<tr>
											<td class="label">abilitato:</td>
											<td class="content">
												<input type="checkbox" disabled class="checkbox" <%= chk(rs("mosp_se_abilitato")) %>>
											</td>
											<td class="label">con spese:</td>
											<td class="content">
												<input type="checkbox" disabled class="checkbox" <%= chk(rs("mosp_se_spesespedizione")) %>>
											</td>
											<td class="label">default:</td>
											<td class="content">
												<input type="checkbox" disabled class="checkbox" <%= chk(rs("mosp_default")) %>>
											</td>
										</tr>
									
									</table>
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