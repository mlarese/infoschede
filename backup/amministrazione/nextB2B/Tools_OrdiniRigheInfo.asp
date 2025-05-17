
<%
dim conn, rs, rsc, sql, Pager, TipCount
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("DetOrdInfo_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("DetOrdInfo_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")


sql = ""
'filtra per nome
if Session("DetOrdInfo_nome")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("DetOrdInfo_nome"), FieldLanguageList("dod_nome_"))
end if

'filtra per tipo
if cInteger(session("DetOrdInfo_tipo"))<>0 then
	sql = sql & " AND dod_tipo=" & cInteger(session("DetOrdInfo_tipo"))
end if

'filtra per categoria
if session("DetOrdInfo_categoria")<>"" then
	sql = sql & " AND dod_id IN (SELECT rtd_descrittore_id FROM grel_dettagli_ord_tipo_des WHERE rtd_tipo_id=" & session("DetOrdInfo_categoria") & ") "
end if

sql = " SELECT *, (SELECT COUNT(*) FROM grel_dettagli_ord_tipo_des WHERE rtd_descrittore_id=dod_id) AS N_TIPI " + _
	  " FROM gtb_dettagli_ord_des " + _
	  " WHERE (1=1) " + sql + " ORDER BY dod_nome_it"
Session(name_session_sql) = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0">
		<tr>
	  		<td style="width:27%;" valign="top">
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
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("DetOrdInfo_nome") %>>NOME</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= TextEncode(session("DetOrdInfo_nome")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("DetOrdInfo_tipo") %>>TIPO DI DATO</th></tr>
								<tr>
									<td class="content">
										<% CALL  DesAdvancedDropTipi("search_tipo", " style=""width:100%"" ", cInteger(session("DetOrdInfo_tipo")), false) %>
									</td>
								</tr>
								<tr><th <%= Search_Bg("DetOrdInfo_categoria") %>>TIPOLOGIA A CUI &Egrave; ASSOCIATA</th></tr>
								<tr>
									<td class="content">
                                        <%	sql = "SELECT * FROM gtb_dettagli_ord_tipo ORDER BY dot_nome_it"
										CALL dropDown(conn, sql, "dot_id", "dot_nome_it", "search_categoria", session("DetOrdInfo_categoria"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
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
					<caption>
						Elenco informazioni
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="OrdiniRigheInfoMod.asp?ID=<%= rs("dod_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<% if rs("N_TIPI") > 0 then %>
																<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare l'informazione: sono presenti tipologie associate">
																	CANCELLA
																</a>
															<% else %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ORDINI_INFO_RIGHE','<%= rs("dod_id") %>');" >
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												<%= rs("dod_nome_it") %>
											</td>
										</tr>
                                        <tr>
											<td class="label" style="width:22%;">tipo di dato:</td>
											<td class="content" style="width:50%;"><%= DesVisTipo(rs("dod_tipo")) %></td>
											<td class="label" style="width:8%;">codice:</td>
											<td class="content"><%= rs("dod_codice") %></td>
										</tr>
										<tr>
											<td class="label">categorie associate:</td>
											<td class="content" colspan=3>
												<% sql = " SELECT * FROM gtb_dettagli_ord_tipo INNER JOIN grel_dettagli_ord_tipo_des ON gtb_dettagli_ord_tipo.dot_id = grel_dettagli_ord_tipo_des.rtd_tipo_id " + _
                                                         " WHERE rtd_descrittore_id = " & rs("dod_id")
												rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext 
												TipCount = 0
												if rsc.eof then%>
													Non associata a nessuna categoria.
												<% else%>
													<span id="categorie_<%= rs("dod_id")%>">
													<%while not rsc.eof
														TipCount = TipCount + 1 %>
														<%= rsc("dot_nome_it") %>
														<% rsc.movenext
														if not rsc.eof then %>
															<br>
														<%end if
													wend %>
													</span>
													<%if TipCount > 4 then %>
														<style type="text/css">
															#categorie_<%= rs("dod_id")%>{
																height:50px; 
																overflow:auto;
																width:100%;
															}
														</style>
													<%end if
												end if
												rsc.close %>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<%rs.moveNext
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
set rsc = nothing
set conn = nothing%>