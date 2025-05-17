<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
Imposta_Proprieta_Sito("ID")
%>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->
<% 
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_template_accesso, 0))

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - templates - elenco"
dicitura.puls_new = "INDIETRO A SITI;NUOVO TEMPLATE"
dicitura.link_new = "Siti.asp;SitoTemplateNew.asp"
dicitura.scrivi_con_sottosez()

dim conn, sql, pager, rs, rst

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rst = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("te_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("te_")
	end if
end if

'filtra per titolo
if Session("te_titolo")<>"" then
    sql = sql & " AND " & SQL_FullTextSearch(Session("te_titolo"), "nomepage")
end if

'filtra per tipo
if session("te_semplificato") = "2" then
	sql = sql &" AND "& SQL_IsTrue(conn, "semplificata")
elseif session("te_semplificata") = "1" then
	sql = sql &" AND NOT "& SQL_IsTrue(conn, "semplificata")
end if

'filtra per contenuto
if session("te_testo") <> "" OR session("te_img") <> "" OR session("te_plugin") <> "" then
	sql = sql &" AND EXISTS (SELECT 1 FROM tb_layers WHERE tb_pages.id_page = tb_layers.id_pag "
			   
	'filtra per testo
	if session("te_testo") <> "" then
		sql = sql &" AND id_tipo = 1 AND (" & SQL_FullTextSearch(Session("te_testo"), "testo") & ")"
	end if
	
	'filtra per file
	if session("te_img") <> "" then
		if Left(session("te_img"), 1) = "/" then
			session("te_img") = Right(session("te_img"), Len(session("te_img")) - 1)
		end if
		sql = sql &" AND (id_tipo = 2 OR id_tipo = 3) AND nome = '"& ParseSQL(_
			Replace(Replace(Session("te_img"), "flash/", ""), "images/", ""), adChar) &"'"
	end if
	
	'filtra per plugin
	if session("te_plugin") <> "" then
		sql = sql &" AND id_tipo = 4 AND " + _
				   " ( id_objects = "& ParseSQL(Session("te_plugin"), adNumeric) & " OR " & _
				   "   testo LIKE '%=%" + ParseSQL(GetValueList(conn, null, "SELECT name_objects FROM tb_objects WHERE id_objects =" & Session("te_plugin")), adChar) + "%;%'" & _
				   "  ) "
	end if
	
	sql = sql &")"
end if

sql = "SELECT * FROM tb_pages WHERE "& SQL_IsTrue(conn, "template") &" AND id_webs=" & Session("AZ_ID") & sql &" ORDER BY nomepage"
session("WEB_TEMPLATE_SQL") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
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
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("te_titolo") %>>TITOLO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_titolo" value="<%= TextEncode(session("te_titolo")) %>" style="width:100%;">
										</td>
									</tr>
									<% 
									if not Session("SITO_MOBILE") then
									%>
									<tr><th colspan="2" <%= Search_Bg("te_semplificato") %>>TIPO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="checkbox" class="checkbox" name="search_semplificato" value="1" <%= chk(instr(1, session("te_semplificato"), "1", vbTextCompare)>0) %>>
											per pagine normali
										</td>
									</tr>
									<tr>
										<td class="content">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td><input type="checkbox" class="checkbox" name="search_semplificato" value="2" <%= chk(instr(1, Session("te_semplificato"), "2", vbTextCompare)>0) %>></td>
													<td style="padding-right:4px;"><img src="../grafica/notReadKnow.gif" border="0" alt="Template per email con visualizzazione semplificata."></td>
													<td>per email semplificate</td>
												</tr>
											</table>
										</td>
									</tr>
									<% 
									end if
									%>
									<tr><th colspan="2" <%= Search_Bg("te_testo;te_img;te_plugin") %>>CONTENUTO</th></tr>
									<tr><th colspan="2" class="L2" <%= Search_Bg("te_testo") %>>TESTO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_testo" value="<%= TextEncode(session("te_testo")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" class="L2" <%= Search_Bg("pa_img") %>>FILE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteFilePicker_Input(Application("AZ_ID"), "", "ricerca", "search_img", session("pa_img"), "width:88px", false) %>
										</td>
									</tr>
									<tr><th colspan="2" class="L2" <%= Search_Bg("te_plugin") %>>PLUGIN</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% 	sql = "SELECT * FROM tb_objects ORDER BY name_objects"
												CALL dropDown(conn, sql, "id_objects", "name_objects", "search_plugin", _
															  session("te_plugin"), false, "style=""width: 100%;""", LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr>
										<td class="footer" colspan="2">
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
					<caption>
						Elenco templates
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> templates in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="4">
												<% if rs("semplificata") AND (not session("SITO_MOBILE")) then %>
													<table border="0" cellspacing="0" cellpadding="0" align="left">
														<tr>
															<td style="padding-top:1px;font-size:1px;">
																<img src="../grafica/notReadKnow.gif" border="0" alt="Template per email con visualizzazione semplificata.">
																&nbsp;
															</td>
														</tr>
													</table>
												<% end if %>
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="dynalay.asp?PAGINA=<%= rs("id_page") %>&lingua=it" target="_blank" title="apre l'anteprima del template in una nuova finestra." <%= ACTIVE_STATUS %>>
																VEDI
															</a>
															&nbsp;
															<a class="button" href="SitoTemplateMod.asp?ID=<%= rs("id_page") %>">
																MODIFICA
															</a>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('TEMPLATE','<%= rs("id_page") %>');" >
																CANCELLA
															</a>
														</td>
													</tr>
												</table>
												<%= rs("nomepage") %>
											</td>
										</tr>
										<%sql = "SELECT COUNT(*) FROM tb_pages WHERE id_template=" & rs("id_page") & " AND (" + _
											    IIF(Session("LINGUA_EN"), " id_page IN (SELECT <tipo_pagina>_EN FROM tb_pagineSito) OR ", "") + _
												IIF(Session("LINGUA_FR"), " id_page IN (SELECT <tipo_pagina>_FR FROM tb_pagineSito) OR ", "") + _
												IIF(Session("LINGUA_DE"), " id_page IN (SELECT <tipo_pagina>_DE FROM tb_pagineSito) OR ", "") + _
												IIF(Session("LINGUA_ES"), " id_page IN (SELECT <tipo_pagina>_ES FROM tb_pagineSito) OR ", "") + _
												IIF(Session("LINGUA_RU"), " id_page IN (SELECT <tipo_pagina>_RU FROM tb_pagineSito) OR ", "") + _
												IIF(Session("LINGUA_CN"), " id_page IN (SELECT <tipo_pagina>_CN FROM tb_pagineSito) OR ", "") + _
												IIF(Session("LINGUA_PT"), " id_page IN (SELECT <tipo_pagina>_PT FROM tb_pagineSito) OR ", "") + _
												" id_page IN (SELECT <tipo_pagina>_IT FROM tb_pagineSito) ) "
										%>
										<tr>
											<td class="label" rowspan="2">usato in:</td>
											<td class="content" colspan="2">
												n&ordm; <%= cInteger(GetValueList(conn, rst, replace(sql, "<tipo_pagina>", "id_pagDyn"))) %>
												pagine pubblicate
											</td>
											<td class="content_right" rowspan="2">
												<% 	CALL WriteCampoCerca("SitoPagine.asp", "template", rs("id_page"), "ELENCO PAGINE COLLEGATE", "button_L2") %>
											</td>
										</tr>
										<tr>
											<td class="content" colspan="2">
												n&ordm; <%= cInteger(GetValueList(conn, rst, replace(sql, "<tipo_pagina>", "id_pagStage"))) %>
												pagine di lavoro
											</td>
										</tr>
										<tr>
											<td class="label">numero</td>
											<td class="content" style="width:29%;"><%= rs("id_page") %></td>
											<td class="label" style="width:10%;">tipo</td>
											<% if rs("semplificata") then %>
												<td class="content_b">Template per email semplificate</td>
											<% else %>
												<td class="content">Template per pagine normali</td>
											<% end if %>
										</tr>
										<tr>
											<td class="label" style="width:20%;">ultima modifica</td>
											<td class="content" colspan="3"><%= rs("page_modData") %></td>
										</tr>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td class="footer" style="border-top:0px; text-align:left;" colspan="2">
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
set rst = nothing
set conn = nothing%>
