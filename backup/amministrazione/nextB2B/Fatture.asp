<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione fatture - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA FATTURA"
dicitura.link_new = "Tabelle.asp;FattureNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql, pager, i, sql_filtri

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("fatture_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("fatture_")
	end if
end if

sql = ""
'filtra per codice
if Session("fatture_codice")<>"" then
    sql = sql & " AND (" & SQL_FullTextSearch(Session("fatture_codice"), "mar_codice") & ") "
end if

'filtra per nome
if Session("fatture_nome")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("fatture_nome"), FieldLanguageList("mar_nome_")) & ") "
end if

'filtra per nome cliente
if Session("fatture_denominazione")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("fatture_denominazione"))
end if

'filtra per descrizione
if Session("fatture_descrizione")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("fatture_descrizione"), FieldLanguageList("mar_descr_")) & ") "
end if

sql_filtri = sql

sql = "SELECT * FROM gtb_fatture " 
if Session("fatture_denominazione") <> "" then
	sql = sql + " INNER JOIN gv_agenti ON gtb_fatture.fa_emittente_id = gv_agenti.ag_id "
end if
sql = sql + "  WHERE (1=1) " + sql_filtri + " ORDER BY fa_id DESC "



session("B2B_fatture_SQL") = sql
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
									<tr><th <%= Search_Bg("fatture_codice") %>>CODICE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_codice" value="<%= TextEncode(session("fatture_codice")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("fatture_nome") %>>NOME</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome" value="<%= TextEncode(session("fatture_nome")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("fatture_denominazione") %>>CLIENTE</th></tr>
									<tr>
										<td class="content" colspan="2">
										<input type="text" name="search_denominazione" value="<%= TextEncode(session("fatture_denominazione")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("fatture_descrizione") %>>DESCRIZONE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_descrizione" value="<%= TextEncode(session("fatture_descrizione")) %>" style="width:100%;">
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
			<%
			dim is_bozza, type_doc
			%>
			<td valign="top">
				<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>Elenco fatture</caption>
					<% if not rs.eof then %>
						<tr><th>Trovate n&ordm; <%= Pager.recordcount %> fatture in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<%
							is_bozza = cBooloean(rs("fa_is_bozza"), false)
							if is_bozza then
								type_doc = "bozza"
							else
								type_doc = "fattura"
							end if
							%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px; text-align:right;">
															<% if is_bozza then %>
																<a class="button" href="FattureMod.asp?ID=<%= rs("fa_id") %>">
																	MODIFICA
																</a>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('FATTURE','<%= rs("fa_id") %>');" >
																	CANCELLA
																</a>
															<% else %>
																<a class="button" href="FattureMod.asp?ID=<%= rs("fa_id") %>">
																	VISUALIZZA
																</a>
																<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la fattura">
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												<%
												response.write type_doc & " " & rs("fa_numero") & IIF(cIntero(rs("fa_serie")>0,"/"&rs("fa_serie"),"")) & " del " & rs("fa_anno")
												%>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:23%;">test:</td>
											<td class="content_right" colspan="3" style="padding-right:0px; font-size:1px;">
												ciao
											</td>
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
set conn = nothing%>