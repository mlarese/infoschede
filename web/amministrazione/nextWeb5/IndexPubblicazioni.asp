<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<% 
Reset_Proprieta_Sito()

'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_Pubblicazioni_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Indice generale - Pubblicazioni automatiche dei dati"
dicitura.puls_new = "NUOVA PUBBLICAZIONE AUTOMATICA"
dicitura.link_new = "IndexPubblicazioniNew.asp"
dicitura.scrivi_con_sottosez()


dim conn, rs, rsv, sql, Pager
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("pub_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("pub_")
	end if
end if

sql = ""
if Session("pub_titolo")<>"" then
	sql = sql & IIF(sql <> "", " AND ", " WHERE ")
	sql = sql & SQL_FullTextSearch(Session("pub_titolo"), "pub_titolo")
end if
if Session("pub_tabella")<>"" then
	sql = sql & IIF(sql <> "", " AND ", " WHERE ")
	sql = sql & SQL_FullTextSearch(Session("pub_tabella"), "tab_titolo")
end if
if cInteger(Session("pub_sito"))>0 then
    sql = sql & IIF(sql <> "", " AND ", " WHERE ")
	sql = sql & " id_sito = " & Session("pub_sito")
end if


sql = " SELECT * FROM (((tb_siti_tabelle_pubblicazioni"& _
	  " INNER JOIN tb_siti_Tabelle ON tb_siti_tabelle_pubblicazioni.pub_tabella_id = tb_siti_Tabelle.tab_id) " + _
	  " INNER JOIN tb_siti ON tb_siti_tabelle.tab_sito_id = tb_siti.id_sito) " + _ 
	  " LEFT JOIN tb_pagineSito ON tb_siti_Tabelle_pubblicazioni.pub_pagina_id = tb_paginesito.id_pagineSito) " + _
	  " LEFT JOIN tb_webs ON tb_paginesito.id_web = tb_webs.id_webs " + _
	  sql + _
	  " ORDER BY pub_titolo "
session("PUBBLICAZIONI_SQL") = sql
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
										<input type="submit" class="button" name="cerca" value="CERCA" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("pub_titolo") %>>TITOLO PUBBLICAZIONE</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_titolo" value="<%= session("pub_titolo")%>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("pub_tabella") %>>ORIGINE DATI</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_tabella" value="<%= session("pub_tabella")%>" style="width:100%;">
									</td>
								</tr>
                                <tr><th colspan="2" <%= Search_Bg("pub_sito") %>>APPLICAZIONE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% 	sql = "SELECT * FROM tb_siti WHERE id_sito IN (SELECT tab_sito_id FROM tb_siti_tabelle) ORDER BY sito_nome"
												CALL dropDown(conn, sql, "id_sito", "sito_nome", "search_sito", _
															  session("pub_sito"), false, "style=""width: 100%;""", LINGUA_ITALIANO) %>
										</td>
									</tr>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
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
					<caption>Elenco pubblicazione automatiche</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> pubblicazioni in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size:1px;">
															<a class="button" href="IndexPubblicazioniMod.asp?ID=<%= rs("pub_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('PUBBLICAZIONI','<%= rs("pub_id") %>');" >
																CANCELLA
															</a>
														</td>
													</tr>
												</table>
												<%= rs("pub_titolo") %>
											</td>
										</tr>
                                        <tr>
                                            <td class="label_no_width" style="width:25%;">voci pubblicate:</td>
                                            <% sql = "SELECT COUNT(*) FROM rel_index_pubblicazioni WHERE rip_pub_id=" & rs("pub_id") %>
                                            <td class="content"><%= cIntero(GetValueList(conn, rsv, sql)) %></td>
                                        </tr>
										<tr>
											<td class="label_no_width">sorgente dati:</td>
											<td class="content"><%= rs("sito_nome") %> - <span style="color:<%= rs("tab_colore") %>;"><%= rs("tab_titolo") %></span></td>
										</tr>
										<tr>
											<td class="label_no_width">pagina di pubblicazione:</td>
											<td class="content"><%= rs("nome_webs") %> - <%= PaginaSitoNome(rs, "") %></td>
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
	&nbsp;
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsv = nothing
set conn = nothing
%>