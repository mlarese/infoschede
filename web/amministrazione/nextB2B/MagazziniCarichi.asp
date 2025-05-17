<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione magazzini carico merci - elenco"
dicitura.puls_new = "CAMBIA MAGAZZINO;NUOVO CARICO"
dicitura.link_new = "Magazzini.asp;MagazziniCarichiNew.asp?IDMAG=" & request("IDMAG")
dicitura.scrivi_con_sottosez()  

dim conn, rs, aux, sql, pager, header

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("carico_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("carico_")
	end if
end if

'filtra per codice fornitore
if session("carico_codice") <> "" then
    sql = sql &" AND "& SQL_FullTextSearch(Session("carico_codice"), "car_fornitore_cod")
end if

'filtra per nome fornitore
if session("carico_nome_fornitore") <> "" then
    sql = sql &" AND "& SQL_FullTextSearch(Session("carico_nome_fornitore"), "car_note") 
end if

'filtra per codice articolo
if session("carico_codiceA") <> "" then
	sql = sql & " AND car_id IN (SELECT rcv_car_id FROM gv_carichi WHERE" & _
                SQL_FullTextSearch(Session("carico_codiceA"), "rel_cod_int") & ") "
end if

'filtra per data carico
if isDate(Session("carico_data_ins_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "car_data", adCompareGreaterThan, Session("carico_data_ins_from"))
end if
if isDate(Session("carico_data_ins_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "car_data", adCompareLessThan, Session("carico_data_ins_to"))
end if

'ricerca per stato del carico
if Session("carico_stato_movimentato")<>"" then
	if not (instr(1, Session("carico_stato_movimentato"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("carico_stato_movimentato"), "0", vbTextCompare)>0 ) then
		sql = sql & " AND "
		if instr(1, Session("carico_stato_movimentato"), "1", vbTextCompare)>0 then
			'carico movimentato
			sql = sql & SQL_IsTrue(conn, "car_movimentato") 
		elseif instr(1, Session("carico_stato_movimentato"), "0", vbTextCompare)>0 then
			'carico in attesa
			sql = sql & " NOT (" &  SQL_IsTrue(conn, "car_movimentato") & ") "
		end if
	end if
end if

'filtra per campo note
if session("carico_note") <> "" then
    sql = sql &" AND "& SQL_FullTextSearch(Session("carico_note"), "car_note")
end if

sql = "SELECT * FROM gtb_carichi INNER JOIN "& _
	  "gtb_magazzini ON gtb_carichi.car_magazzino_id = gtb_magazzini.mag_id"& _
	  " WHERE (mag_ID="&cIntero(request("IDMAG"))&") "& sql & _
	  " ORDER BY car_data DESC"	  
session("B2B_CARICHI_SQL") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
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
									<tr><th <%= Search_Bg("carico_codice") %>>CODICE FORNITORE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_codice" value="<%= TextEncode(session("carico_codice")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("carico_nome_fornitore") %>>NOME / DENOMINAZIONE FORNITORE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome_fornitore" value="<%= TextEncode(session("carico_nome_fornitore")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("carico_codiceA") %>>CODICE ARTICOLO</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_codiceA" value="<%= TextEncode(session("carico_codiceA")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("carico_data_ins_from;carico_data_ins_to") %>>DATA DI CARICO</td></tr>
									<tr><td class="label">a partire dal:</td></tr>
									<tr>
										<td class="content">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_from", Session("carico_data_ins_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><td class="label">fino al:</td></tr>
									<tr>
										<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_to", Session("carico_data_ins_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("carico_stato_movimentato") %>>STATO DEL CARICO</td></tr>
								<tr>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_stato_movimentato" value="1" <%= chk(instr(1, session("carico_stato_movimentato"), "1", vbTextCompare)>0) %>>
										movimentato
									</td>
									</tr>
									<tr>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_stato_movimentato" value="0" <%= chk(instr(1, Session("carico_stato_movimentato"), "0", vbTextCompare)>0) %>>
										non non movimentato
									</td>
								</tr>
									<tr><th <%= Search_Bg("carico_note") %>>NOTE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_note" value="<%= TextEncode(session("carico_note")) %>" style="width:100%;">
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
				<% if not rs.eof then %>
					<caption>
						Elenco carichi effettuati nel magazzino: <%= rs("mag_nome") %>
					</caption>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> carichi in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
										<%	
											header = "header"
										%>
											<td class="<%= header %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="MagazziniCarichiMod.asp?ID=<%= rs("car_id") %>">
																MODIFICA
															</a>
															&nbsp;
														<% if rs("car_movimentato") > 0 then %>
															<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare un carico: per ora ...">
																CANCELLA
															</a>
														<% else %>
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CARICHI','<%= rs("car_id") %>');" >
																CANCELLA
															</a>
														<% end if %>
														
														</td>
													</tr>
												</table>
											<%=rs("car_data")%>
											</td>
										</tr>
										<tr>
											<td class="label">Fornitore:</td>
											<td class="content" width="50%"><%= rs("car_fornitore") %></td>
											<td class="label">Codice:</td>
											<td class="content"><%= rs("car_fornitore_cod") %></td>
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
						<caption>
							Elenco carichi effettuati.
						</caption>
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
set aux = nothing
set conn = nothing%>
