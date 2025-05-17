
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->

<%

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if cString(request("FIRSTDATE")) = "" then
		'se sono sull'elenco e non sul calendario
		Pager.Reset()
	end if
	CALL SearchSession_Reset("imp_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("imp_")
	end if
end if

sql = ""

'ricerca full text sul contenuto
if Session("imp_titolo")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("imp_titolo"), FieldLanguageList("imp_titolo_"))
end if


'ricerca per tipologia impegno
if Session("imp_tipologia")<>"" then
	sql = sql & " AND imp_tipo_id IN (" & Session("imp_tipologia") & ")"
end if


'ricerca se scaduti
if Session("imp_scadenza")<>"" then
	if not (instr(1, Session("imp_scadenza"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("imp_scadenza"), "0", vbTextCompare)>0 ) then
		if instr(1, Session("imp_scadenza"), "1", vbTextCompare)>0 then
			'documento scaduto
			sql = sql & " AND " & SQL_CompareDateTime(conn, "imp_data_ora_fine", adCompareLessThan, DateIso(Now()))
		elseif instr(1, Session("imp_scadenza"), "0", vbTextCompare)>0 then
			'documento non scaduto
			sql = sql & " AND " & SQL_CompareDateTime(conn, "imp_data_ora_fine", adCompareGreaterThan, DateIso(Now()))
		end if
	end if
end if


'ricerca se protetto
if Session("imp_protetto")<>"" then
	if not (instr(1, Session("imp_protetto"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("imp_protetto"), "0", vbTextCompare)>0 ) then
		if instr(1, Session("imp_protetto"), "1", vbTextCompare)>0 then
			'documento protetto
			sql = sql & " AND " & SQL_IsTrue(conn, "imp_protetto")
		elseif instr(1, Session("imp_protetto"), "0", vbTextCompare)>0 then
			'documento non protetto
			sql = sql & " AND NOT(" & SQL_IsTrue(conn, "imp_protetto") & ") "
		end if
	end if
end if


'filtra per nome utente area riservata
if Session("imp_nome_utente")<>"" then
	dim ut_ids
	ut_ids = " SELECT ut_ID FROM tb_utenti WHERE ut_NextCom_id IN " & _
			 "				(SELECT IDElencoIndirizzi FROM tb_indirizzario WHERE ( " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("imp_nome_utente")) & " ))"
	
	sql = sql & " AND (((imp_id IN (SELECT riu_impegno_id FROM mrel_impegni_utenti WHERE riu_utente_id IN (" & ut_ids & "))) OR " & _
				" 		(imp_id IN (SELECT rip_impegno_id FROM mrel_impegni_profili WHERE rip_profilo_id IN " & _
				"											(SELECT rpu_profilo_id FROM mrel_profili_utenti WHERE rpu_utenti_id IN (" & ut_ids & "))))))"
	
	'mostro anche i documenti visibili a tutti, se il filtro per nome utente area riservata dà risultati
	'sql = sql & " OR (NOT " & SQL_IsTrue(conn, "imp_protetto") & " AND EXISTS (" & ut_ids & ")))"
end if


'ricerca per profilo collegato agli impegni
if Session("imp_profilo")<>"" then
	sql = sql & " AND imp_id IN (SELECT rip_impegno_id FROM mrel_impegni_profili WHERE rip_profilo_id = " & Session("imp_profilo") & ")"
end if


if isDate(Session("imp_data_inizio")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "imp_data_ora_fine", adCompareGreaterThan, Session("imp_data_inizio"))
end if
if isDate(Session("imp_data_fine")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "imp_data_ora_inizio", adCompareLessThan, Session("imp_data_fine"))
end if


sql = " SELECT * FROM mtb_impegni " + _
	  " WHERE (1=1) " + sql + _
	  " ORDER BY imp_data_ora_fine, imp_titolo_it "
Session("SQL_IMPEGNI") = sql





Function WriteBloccoRicerca(conn, orientation)
	dim rst
	set rst = Server.CreateObject("ADODB.RecordSet")
	%>
	<!-- BLOCCO DI RICERCA -->
	<% if orientation = "vertical" then %>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
			<form action="" method="post" id="ricerca" name="ricerca">
			<tr>
				<td>
					<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption>Opzioni di ricerca</caption>
						<tr>
							<td class="footer" colspan="2">
								<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
								<input type="submit" class="button" name="tutti" value="VEDI TUTTE" style="width: 49%;">
							</td>
						</tr>
						<tr><th colspan="2" <%= Search_Bg("imp_titolo") %>>TITOLO</td></tr>
						<tr>
							<td class="content" colspan="2">
								<input type="text" name="search_titolo" value="<%= TextEncode(session("imp_titolo")) %>" style="width:100%;">
							</td>
						</tr>
						
						<tr><th colspan="2" <%= Search_Bg("imp_tipologia") %>>TIPOLOGIA</th></tr>
						<% sql = "SELECT * FROM mtb_tipi_impegni ORDER BY tim_nome_it" 
						rst.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
						while not rst.eof 
							%>
							<tr>
								<td class="content" colspan="2">
									<input type="checkbox" class="checkbox" name="search_tipologia" value="<%=rst("tim_id")%>" <%= chk(inStr(session("imp_tipologia"),cString(rst("tim_id")))>0) %>>
									<% WriteColor(rst("tim_colore"))%>
									<%=rst("tim_nome_it")%>
								</td>
							</tr>
							<%
							rst.moveNext
						wend
						%>

						<tr><th colspan="2" <%= Search_Bg("imp_scadenza") %>>SCADENZA</th></tr>
						<tr>
							<td class="content" style="width:45%;" colspan="2">
								<input type="checkbox" class="checkbox" name="search_scadenza" value="1" <%= chk(instr(1, session("imp_scadenza"), "1", vbTextCompare)>0) %>>
								impegni scaduti
							</td>
						</tr>
						<tr>
							<td class="content_b" colspan="2">
								<input type="checkbox" class="checkbox" name="search_scadenza" value="0" <%= chk(instr(1, Session("imp_scadenza"), "0", vbTextCompare)>0) %>>
								impegni non scaduti
							</td>
						</tr>
						
						<tr><th colspan="2" <%= Search_Bg("imp_protetto") %>>VISIBILITA'</th></tr>
						<tr>
							<td class="content OrdEvaso">
								<input type="checkbox" class="checkbox OrdEvaso" name="search_protetto" value="0" <%= chk(instr(1, Session("imp_protetto"), "0", vbTextCompare)>0) %>>
								pubblico
							</td>
							<td class="content OrdConfermato" style="width:45%;">
								<input type="checkbox" class="checkbox OrdConfermato" name="search_protetto" value="1" <%= chk(instr(1, session("imp_protetto"), "1", vbTextCompare)>0) %>>
								<img src="../grafica/padlock.gif" border="0" alt="Pagina appartenente all'area protetta">		
								privato
							</td>
						</tr>
						
						<% if cBoolean(Session("CONDIVISIONE_PUBBLICA"), false) then %>
							<tr><th colspan="2" <%= Search_Bg("imp_nome_utente;imp_profilo") %>>UTENTI IMPEGNATI</th></tr>
							<tr><th colspan="2" class="L2" <%= Search_Bg("imp_nome_utente") %>>UTENTE</th></tr>
							<tr>
								<td class="content" colspan="2">
									<input type="text" name="search_nome_utente" value="<%= TextEncode(session("imp_nome_utente")) %>" style="width:100%;">
								</td>
							</tr>
							<% if profili_attivi then 
								sql = "SELECT * FROM mtb_profili ORDER BY pro_nome_it"
								if GetValueList(conn, NULL, sql) <>"" then %>
									<tr><th colspan="2" class="L2" <%= Search_Bg("imp_profilo") %>>PROFILO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL dropDown(conn, sql, "pro_id", "pro_nome_it", "search_profilo", session("imp_profilo"), false, "style=""width: 100%;""", Session("LINGUA")) %>
										</td>
									</tr>
								<% end if %>
							<% end if %>
						<% end if %>
					
						<tr><th colspan="2" <%= Search_Bg("imp_data_inizio;imp_data_fine")%>>IMPEGNI/APPUNTAMENTi</td></tr>
						<tr><td colspan="2" class="label">compresi tra il:</td></tr>
						<tr>
							<td colspan="2" class="content">
								<% CALL WriteDataPicker_Input("ricerca", "search_data_inizio", Session("imp_data_inizio"), "", "/", true, true, LINGUA_ITALIANO) %>
							</td>
						</tr>
						<tr><td colspan="2" class="label">e il:</td></tr>
						<tr>
							<td colspan="2" class="content">
								<% CALL WriteDataPicker_Input("ricerca", "search_data_fine", Session("imp_data_fine"), "", "/", true, true, LINGUA_ITALIANO) %>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="footer">
								<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
								<input type="submit" class="button" name="tutti" value="VEDI TUTTE" style="width: 49%;">
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr><td style="font-size:4px;">&nbsp;</td></tr>
			<tr>
				<td>
					<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption>Strumenti</caption>
						<tr>
							<td class="content_center">
								<% dim sql_export
								sql_export = session("SQL_IMPEGNI")
								sql_export = right(sql_export, len(sql_export) + 1 - instr(1, sql_export, "WHERE (1=1)", vbTextCompare))
								if instrRev(sql_export, "ORDER BY", vbTrue,vbTextCompare) > 0 then
									sql_export = left(sql_export, instrRev(sql_export, "ORDER BY", vbTrue,vbTextCompare) - 1)
								end if
								sql_export = "SELECT imp_id FROM mtb_impegni " & sql_export
								
								sql_export = " SELECT * FROM tb_utenti WHERE ut_ID IN (SELECT riu_utente_id FROM mrel_impegni_utenti WHERE riu_impegno_id IN ("&sql_export&")) OR " & _
												" ut_ID IN (SELECT rpu_utenti_id FROM mrel_profili_utenti WHERE rpu_profilo_id IN (SELECT rip_profilo_id FROM mrel_impegni_profili " & _
																																		" WHERE rip_impegno_id IN ("&sql_export&")))"
								%>
								<% CALL ExportContattiInRubrica(sql_export, "ut_NextCom_ID", "", "") %>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			</form>
		</table>
	<% elseif orientation = "horizontal" then %>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
			<form action="" method="post" id="ricerca" name="ricerca">
			<tr>
				<td>
					<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption>Opzioni di ricerca</caption>
						<tr>
							<td bgcolor="#f4f4f2" style="padding-right:1px; border-right:1px solid #999999; vertical-align:top; width:34%;">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
									<tr><th colspan="2" <%= Search_Bg("imp_titolo") %>>TITOLO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_titolo" value="<%= TextEncode(session("imp_titolo")) %>" style="width:100%;">
										</td>
									</tr>
									
									<tr><th colspan="2" <%= Search_Bg("imp_tipologia") %>>TIPOLOGIA</th></tr>
									<% sql = "SELECT * FROM mtb_tipi_impegni ORDER BY tim_nome_it" 
									rst.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
									while not rst.eof 
										%>
										<tr>
											<td class="content" colspan="2">
												<input type="checkbox" class="checkbox" name="search_tipologia" value="<%=rst("tim_id")%>" <%= chk(inStr(session("imp_tipologia"),cString(rst("tim_id")))>0) %>>
												<% WriteColor(rst("tim_colore"))%>
												<%=rst("tim_nome_it")%>
											</td>
										</tr>
										<%
										rst.moveNext
									wend
									%>

								</table>
							</td>
						
							<td bgcolor="#f4f4f2" style="padding-right:1px; border-right:1px solid #999999; vertical-align:top; width:33%;">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
								
									<tr><th colspan="2" <%= Search_Bg("imp_protetto") %>>VISIBILITA'</th></tr>
									<tr>
										<td class="content OrdEvaso">
											<input type="checkbox" class="checkbox OrdEvaso" name="search_protetto" value="0" <%= chk(instr(1, Session("imp_protetto"), "0", vbTextCompare)>0) %>>
											pubblico
										</td>
										<td class="content OrdConfermato" style="width:45%;">
											<input type="checkbox" class="checkbox OrdConfermato" name="search_protetto" value="1" <%= chk(instr(1, session("imp_protetto"), "1", vbTextCompare)>0) %>>
											<img src="../grafica/padlock.gif" border="0" alt="Pagina appartenente all'area protetta">		
											privato
										</td>
									</tr>
									
									<% if cBoolean(Session("CONDIVISIONE_PUBBLICA"), false) then %>
										<tr><th colspan="2" <%= Search_Bg("imp_nome_utente;imp_profilo") %>>UTENTI IMPEGNATI</th></tr>
										<tr><th colspan="2" class="L2" <%= Search_Bg("imp_nome_utente") %>>UTENTE</th></tr>
										<tr>
											<td class="content" colspan="2">
												<input type="text" name="search_nome_utente" value="<%= TextEncode(session("imp_nome_utente")) %>" style="width:100%;">
											</td>
										</tr>
										<% if profili_attivi then 
											sql = "SELECT * FROM mtb_profili ORDER BY pro_nome_it"
											if GetValueList(conn, NULL, sql) <>"" then %>
												<tr><th colspan="2" class="L2" <%= Search_Bg("imp_profilo") %>>PROFILO</th></tr>
												<tr>
													<td class="content" colspan="2">
														<% CALL dropDown(conn, sql, "pro_id", "pro_nome_it", "search_profilo", session("imp_profilo"), false, "style=""width: 100%;""", Session("LINGUA")) %>
													</td>
												</tr>
											<% end if %>
										<% end if %>
									<% end if %>
								
								</table>
							</td>						
							<td bgcolor="#f4f4f2" style="padding-left:1px; vertical-align:top;">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
									<tr><th colspan="2" <%= Search_Bg("imp_scadenza") %>>SCADENZA</th></tr>
									<tr>
										<td class="content" style="width:45%;" colspan="2">
											<input type="checkbox" class="checkbox" name="search_scadenza" value="1" <%= chk(instr(1, session("imp_scadenza"), "1", vbTextCompare)>0) %>>
											impegni scaduti
										</td>
									</tr>
									<tr>
										<td class="content_b" colspan="2">
											<input type="checkbox" class="checkbox" name="search_scadenza" value="0" <%= chk(instr(1, Session("imp_scadenza"), "0", vbTextCompare)>0) %>>
											impegni non scaduti
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("imp_data_inizio;imp_data_fine")%>>IMPEGNI/APPUNTAMENTi</th></tr>
									<tr><td colspan="2" class="label">compresi tra il:</td></tr>
									<tr>
										<td colspan="2" class="content">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_inizio", Session("imp_data_inizio"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><td colspan="2" class="label">e il:</td></tr>
									<tr>
										<td colspan="2" class="content">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_fine", Session("imp_data_fine"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td colspan="3" class="footer" align="right">
								<input type="submit" name="cerca" value="CERCA" class="button" style="width:80px;">
								<input type="submit" class="button" name="tutti" value="VEDI TUTTE" style="width:80px;">
							</td>
						</tr>
					</table>
				</td>
			</tr>
			</form>
		</table>
	
	
	<% end if 
end function


%>