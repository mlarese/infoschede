<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata


dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "LOG DOWNLOAD COMPLETO"
dicitura.links(1) = "DocumentiLog.asp"


dicitura.sezione = "Gestione documenti - elenco"
dicitura.puls_new = "NUOVO DOCUMENTO / CIRCOLARE"
dicitura.link_new = "DocumentiNew.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, Pager, rs_t, profili_attivi

set Pager = new PageNavigator


'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("docs_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("docs_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rs_t = Server.CreateObject("ADODB.RecordSet")


sql = ""
'ricerca full text sul contenuto
if Session("docs_testo")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("docs_testo"), FieldLanguageList("doc_titolo_"))
end if

'filtra sul numero / protocollo
if Session("docs_numero")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("docs_numero"), "doc_numero")
end if

'filtra per data di pubblicazione
if isDate(Session("docs_data_from")) then
	sql = sql & " AND doc_pubblicazione >=" & SQL_date(conn, Session("docs_data_from"))
end if
if isDate(Session("docs_data_to")) then
	sql = sql & " AND doc_pubblicazione <=" & SQL_date(conn, Session("docs_data_to"))
end if

'ricerca visibilita' / pubblicazione
if Session("docs_visibile")<>"" then
	if not (instr(1, Session("docs_visibile"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("docs_visibile"), "0", vbTextCompare)>0 ) then
		if instr(1, Session("docs_visibile"), "1", vbTextCompare)>0 then
			'documento visibile
			sql = sql & " AND " & SQL_IsTrue(conn, "doc_visibile")
		elseif instr(1, Session("docs_visibile"), "0", vbTextCompare)>0 then
			'documento non visibile
			sql = sql & " AND NOT(" & SQL_IsTrue(conn, "doc_visibile") & ") "
		end if
	end if
end if


'ricerca protezione
if Session("docs_protetto")<>"" then
	if not (instr(1, Session("docs_protetto"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("docs_protetto"), "0", vbTextCompare)>0 ) then
		if instr(1, Session("docs_protetto"), "1", vbTextCompare)>0 then
			'documento protetto
			sql = sql & " AND " & SQL_IsTrue(conn, "doc_protetto")
		elseif instr(1, Session("docs_protetto"), "0", vbTextCompare)>0 then
			'documento non protetto
			sql = sql & " AND NOT(" & SQL_IsTrue(conn, "doc_protetto") & ") "
		end if
	end if
end if


'ricerca profilo
if Session("docs_profilo")<>"" then
	sql = sql & " AND doc_id IN (SELECT rdp_doc_id FROM mrel_doc_profili WHERE rdp_profilo_id = " & Session("docs_profilo") & ")"
end if


'filtra per nome utente area riservata
if Session("docs_nome_utente")<>"" then
	dim ut_ids
	ut_ids = " SELECT ut_ID FROM tb_utenti WHERE ut_NextCom_id IN " & _
			 "				(SELECT IDElencoIndirizzi FROM tb_indirizzario WHERE ( " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("docs_nome_utente")) & " ))"
	
	sql = sql & " AND (((doc_id IN (SELECT rdu_doc_id FROM mrel_doc_utenti WHERE rdu_utenti_id IN (" & ut_ids & "))) OR " & _
				" 		(doc_id IN (SELECT rdp_doc_id FROM mrel_doc_profili WHERE rdp_profilo_id IN " & _
				"											(SELECT rpu_profilo_id FROM mrel_profili_utenti WHERE rpu_utenti_id IN (" & ut_ids & ")))))"
	
	'mostro anche i documenti visibili a tutti, se il filtro per nome utente area riservata dà risultati
	sql = sql & " OR (NOT " & SQL_IsTrue(conn, "doc_protetto") & " AND EXISTS (" & ut_ids & ")))"
end if

'filtro per id utente area riservata
if Session("docs_utente_id")<>"" then
	dim ut_id
	ut_id = cIntero(Session("docs_utente_id"))
	sql = sql & " AND (((doc_id IN (SELECT rdu_doc_id FROM mrel_doc_utenti WHERE rdu_utenti_id = " & ut_id & ")) OR " & _
				" 		(doc_id IN (SELECT rdp_doc_id FROM mrel_doc_profili WHERE rdp_profilo_id IN " & _
				"											(SELECT rpu_profilo_id FROM mrel_profili_utenti WHERE rpu_utenti_id = " & ut_id & "))))"
	
	'mostro anche i documenti visibili a tutti, se il filtro per nome utente area riservata dà risultati
	sql = sql & " OR (NOT " & SQL_IsTrue(conn, "doc_protetto") & "))"
end if


'filtra per nome utente area amministrativa
if Session("docs_nome_admin")<>"" then
	dim adm_ids
	adm_ids = " SELECT ID_admin FROM tb_admin WHERE admin_nome LIKE '%" & Session("docs_nome_admin") & "%' OR admin_cognome LIKE '%" & Session("docs_nome_admin") & "%'"
	
	sql = sql & " AND (((doc_id IN (SELECT rda_doc_id FROM mrel_doc_admin WHERE rda_admin_id IN (" & adm_ids & "))) " & _
				"		OR (doc_id IN (SELECT rdp_doc_id FROM mrel_doc_profili WHERE rdp_profilo_id IN (SELECT rpa_profilo_id FROM mrel_profili_admin WHERE rpa_admin_id IN (" & adm_ids & ")))))"
	
	'mostro anche i documenti visibili a tutti, se il filtro per nome utente area amministrativa dà risultati
	sql = sql & " OR (NOT " & SQL_IsTrue(conn, "doc_protetto") & " AND EXISTS (" & adm_ids & ")))"
end if

'filtro per id utente area amministrativa
if Session("docs_admin_id")<>"" then
	dim adm_id
	adm_id = cIntero(Session("docs_admin_id"))
	sql = sql & " AND (((doc_id IN (SELECT rda_doc_id FROM mrel_doc_admin WHERE rda_admin_id = " & adm_id & ")) " & _
				"		OR (doc_id IN (SELECT rdp_doc_id FROM mrel_doc_profili WHERE rdp_profilo_id IN (SELECT rpa_profilo_id FROM mrel_profili_admin WHERE rpa_admin_id = " & adm_id & "))))"
	
	'mostro anche i documenti visibili a tutti, se il filtro per nome utente area amministrativa dà risultati
	sql = sql & " OR (NOT " & SQL_IsTrue(conn, "doc_protetto") & "))"
end if


'filtra per categoria
if cIntero(request("catc_id")) > 0 then
	Session("docs_categoria") = cIntero(request("catc_id"))
end if
if Session("docs_categoria")<>"" then
	sql = sql & " AND doc_categoria_id IN (" & categorie.FoglieID(Session("docs_categoria")) & " )"
end if


dim testo, filtroDescrittori
CALL DesRicercaQuery(filtroDescrittori, testo, "mtb_carattech", "ct_id", "ct_nome_it", "ct_unita_it", "doc_id", "mrel_doc_ctech", "rdc_ctech_id", "rdc_doc_id", "rdc_valore_", "docs")
if filtroDescrittori<>"" AND sql = "" then
	sql = filtroDescrittori
else
	sql = sql & filtroDescrittori
end if


sql = " SELECT * FROM mtb_documenti " + _
	  " WHERE (1=1) " + sql + _
	  " ORDER BY doc_pubblicazione DESC, doc_titolo_it"
Session("SQL_DOCUMENTI") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)

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
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<form action="" method="post" id="ricerca" name="ricerca">
					<% if cBoolean(Session("CATEGORIE_NEXTMEMO2_ABILITATE"), false) then %>
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px;">
									<caption>Categorie base</caption>
									<% dim selected
									sql = "SELECT catC_id, catC_nome_it FROM mtb_documenti_categorie WHERE catC_livello = 0 "
									rs_t.open sql, conn, adOpenstatic, adLockOptimistic, adCmdText
									while not rs_t.eof
										selected = (rs_t("catC_id") = Session("docs_categoria"))
										%>
										<tr>
											<td class="content" colspan="2" <%=IIF(selected, Search_Bg("docs_categoria"), "")%>>
												<a href="?catc_id=<%=rs_t("catC_id")%>" style="padding-top:4px;padding-bottom:4px;display:block;<%=IIF(selected,"color:white !important;","")%>"><%=rs_t("catC_nome_it")%></a>
											</td>
										</tr>
										<%
										rs_t.moveNext
									wend
									rs_t.close %>
								</table>
							</td>
						</tr>
					<% end if %>
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
								<tr><th colspan="2" <%= Search_Bg("docs_testo") %>>TITOLO</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="text" name="search_testo" value="<%= TextEncode(session("docs_testo")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("docs_numero") %>>NUMERO / PROTOCOLLO</th></tr>
								<tr>
									<td colspan="2" class="content">
										<input type="text" name="search_numero" value="<%= TextEncode(session("docs_numero")) %>" style="width:100%;">
									</td>
								</tr>
								<% if cBoolean(Session("CATEGORIE_NEXTMEMO2_ABILITATE"), false) then %>
									<tr><th colspan="2" <%= Search_Bg("docs_categoria") %>>CATEGORIA</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL categorie.WritePicker("ricerca", "search_categoria", session("docs_categoria"), false, true, 32) %>
										</td>
									</tr>
								<% end if %>
								<tr><th colspan="2" <%= Search_Bg("docs_visibile") %>>VISIBILE / PUBBLICATO</th></tr>
								<tr>
									<td class="content_b" style="width:45%;">
										<input type="checkbox" class="checkbox" name="search_visibile" value="1" <%= chk(instr(1, session("docs_visibile"), "1", vbTextCompare)>0) %>>
										visibile
									</td>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_visibile" value="0" <%= chk(instr(1, Session("docs_visibile"), "0", vbTextCompare)>0) %>>
										non visibile
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("docs_protetto") %>>PROTEZIONE</th></tr>
								<tr>
									<td class="content OrdConfermato" style="width:45%;">
										<input type="checkbox" class="checkbox OrdConfermato" name="search_protetto" value="1" <%= chk(instr(1, session("docs_protetto"), "1", vbTextCompare)>0) %>>
										protetto
									</td>
									<td class="content OrdEvaso">
										<input type="checkbox" class="checkbox OrdEvaso" name="search_protetto" value="0" <%= chk(instr(1, Session("docs_protetto"), "0", vbTextCompare)>0) %>>
										non protetto
									</td>
								</tr>
								<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) _ 
											AND profili_attivi then 
									sql = "SELECT * FROM mtb_profili ORDER BY pro_nome_it"
									if GetValueList(conn, NULL, sql) <>"" then %>
										<tr><th colspan="2" <%= Search_Bg("docs_profilo") %>>PROFILO</th></tr>
										<tr>
											<td class="content" colspan="2">
												<% CALL dropDown(conn, sql, "pro_id", "pro_nome_it", "search_profilo", session("docs_profilo"), false, "style=""width: 100%;""", Session("LINGUA")) %>
											</td>
										</tr>
									<% end if %>
								<% end if %>
								<% if cBoolean(Session("CONDIVISIONE_PUBBLICA"), false) then %>
									<% if cIntero(session("docs_utente_id"))=0 then %>
										<tr><th colspan="2" <%= Search_Bg("docs_nome_utente") %>>VISIBILE ALL'UTENTE AREA RISERVATA</th></tr>
										<tr>
											<td class="content" colspan="2">
												<input type="text" name="search_nome_utente" value="<%= TextEncode(session("docs_nome_utente")) %>" style="width:100%;">
											</td>
										</tr>
									<% else %>
										<tr><th colspan="2" <%= Search_Bg("docs_utente_id") %>>VISIBILE ALL'UTENTE AREA RISERVATA</th></tr>
										<tr>
											<td class="content_b" colspan="2">
												<% sql = "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi = (SELECT ut_NextCom_id FROM tb_utenti WHERE ut_id = " & cIntero(Session("docs_utente_id")) & ")"
												rs_t.open sql, conn, adOpenstatic, adLockOptimistic, adCmdText %>
												<%= ContactFullName(rs_t) %>
												<input type="hidden" name="search_utente_id" value="<%= Session("docs_utente_id") %>" style="width:100%;">
												<% rs_t.close %>
											</td>
										</tr>
									<% end if %>
								<% end if %>
								<% if cBoolean(Session("CONDIVISIONE_INTERNA"), false) then %>
									<% if cIntero(session("docs_admin_id"))=0 then %>
										<tr><th colspan="2" <%= Search_Bg("docs_nome_admin") %>>VISIBILE ALL'UTENTE AREA AMMINISTRATIVA</th></tr>
										<tr>
											<td class="content" colspan="2">
												<input type="text" name="search_nome_admin" value="<%= TextEncode(session("docs_nome_admin")) %>" style="width:100%;">
											</td>
										</tr>
									<% else %>
										<tr><th colspan="2" <%= Search_Bg("docs_admin_id") %>>VISIBILE ALL'UTENTE AREA AMMINISTRATIVA</th></tr>
										<tr>
											<td class="content_b" colspan="2">
												<% sql = "SELECT * FROM tb_admin WHERE ID_admin = " & cIntero(Session("docs_admin_id"))
												rs_t.open sql, conn, adOpenstatic, adLockOptimistic, adCmdText %>
												<%= rs_t("admin_nome")%>&nbsp<%= rs_t("admin_cognome")%>
												<input type="hidden" name="search_admin_id" value="<%= Session("docs_admin_id") %>" style="width:100%;">
												<% rs_t.close %>
											</td>
										</tr>
									<% end if %>
								<% end if %>
								<tr><th colspan="2" <%= Search_Bg("docs_data_from;docs_data_to") %>>DATA DI PUBBLICAZIONE</th></tr>
								<tr><td colspan="2" class="label">a partire dal:</td></tr>
								<tr>
									<td colspan="2" class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_from", Session("docs_data_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td colspan="2" class="label">fino al:</td></tr>
								<tr>
									<td colspan="2" class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_to", Session("docs_data_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><th colspan="2">CARATTERISTICHE</th></tr>
								<tr>
									<td colspan="2">
										<table cellspacing="1" cellpadding="0" style="width:100%;">
											<% sql = "SELECT * FROM mtb_carattech WHERE ISNULL(ct_per_ricerca, 0) = 1 ORDER BY ct_nome_it"
												CALL DesRicercaEX(conn, sql, "mtb_carattech", "ct_id", "ct_nome_it", "ct_tipo", "ct_unita_it", "search", "docs") %>
										</table>
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
					</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco documenti
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="<%= IIF(rs("doc_visibile"), "header", "header_disabled") %>" colspan="7">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="DocumentiLog.asp?ID=<%= rs("doc_id") %>" title="registrazione dei download del documento">
																LOG
															</a>
															&nbsp;
															<% CALL index.WriteButton("mtb_documenti", rs("doc_id"), POS_ELENCO) %>
															<a class="button" href="DocumentiMod.asp?ID=<%= rs("doc_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('DOCUMENTI','<%= rs("doc_id") %>');" >
																CANCELLA
															</a>
														</td>
													</tr>
												</table>
												<% if cString(rs("doc_numero")) <> "" then %>
													<%= rs("doc_numero") %> - <%= rs("doc_titolo_it") %>
												<% else %>
													<%= rs("doc_titolo_it") %>
												<% end if %>
											</td>
										</tr>
										<% if cBoolean(Session("CATEGORIE_NEXTMEMO2_ABILITATE"), false) then %>
											<tr>
												<td class="label">categoria:</td>
												<td class="content" colspan="4">
													<% if cInteger(rs("doc_categoria_id"))>0 then %>
														<%= categorie.NomeCompleto(rs("doc_categoria_id")) %>
													<% else %>
														<span class="note">categoria non impostata</span>
													<% end if %>
												</td>
											</tr>
										<% end if %>
										<tr>
											<td class="label" style="width:20%;">pubblicazione:</td>
											<td class="label_right" style="width:8%;">dal:</td>
											<td class="content" style="width:28%;"><%=DateIta(rs("doc_pubblicazione"))%></td>
											<% if cString(rs("doc_scadenza")) <> "" then %>
												<td class="label_right">al:</td>
												<td class="content"><%=DateIta(rs("doc_scadenza"))%></td>
											<% else %>
												<td class="content" colspan="2">&nbsp;</td>
											<% end if %>
										</tr>
										<tr>
											<td class="label">stato:</td>
											<td class="label_right">visibilit&agrave;:</td>
											<td class="content"><%= IIF(rs("doc_visibile"), "visibile", "non visibile") %></td>
											<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) then %>
												<td class="label_right">protezione:</td>
												<td class="content <%= IIF(rs("doc_protetto"), " OrdConfermato", " OrdEvaso") %>"><%= IIF(rs("doc_protetto"), "protetto", "non protetto") %></td>
											<% else %>
												<td class="content" colspan="2">&nbsp;</td>
											<% end if %>
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