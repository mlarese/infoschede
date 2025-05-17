<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<style type="text/css"> 
	.ascx {
	  background-color: #d0e4f5 !important;
	}

	.class {
	  background-color: #d7efd7 !important;	 		  
	}
	
	.html {
	  background-color: #fce0c7 !important;
	  color: #8e4302 !important;
	}
</style>

<%
Imposta_Proprieta_Sito("ID")
%>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<% 
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_plugin_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - plugin - elenco"
dicitura.puls_new = "INDIETRO A SITI;NUOVO PLUGIN"
dicitura.link_new = "Siti.asp;SitoPluginNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, rsp, sql, pager, isUsed
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("og_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("og_")
	end if
elseif cIntero(request("PAGINA"))<>0 then
	Pager.Reset()
	CALL SearchSession_Reset("og_")
	Session("og_pagina") = request("PAGINA")
end if

'filtra per nome
if Session("og_nome")<>"" then
    sql = sql & " AND " & SQL_FullTextSearch(Session("og_nome"), "name_objects")
end if

'filtra per tipo
if(Session("og_type")<>"") then
	sql = sql & "AND (1=0"
	if inStr(Session("og_type"),"0")>0 then
		sql = sql & " OR obj_type='ascx' "
	end if
	if inStr(Session("og_type"),"1")>0 then
		sql = sql & " OR obj_type='class' "
	end if
	if inStr(Session("og_type"),"2")>0 then
		sql = sql & " OR obj_type='html' "
	end if
	sql=sql & ")"
end if

'filtra per classe
if Session("og_classe")<>"" then
    sql = sql & " AND " & SQL_FullTextSearch(Session("og_classe"), "identif_objects")
end if

'filtra per parametro
if Session("og_parametri")<>"" then
    sql = sql & " AND (" & SQL_FullTextSearch(Session("og_parametri"), "param_list") & _
				" OR " & SQL_FullTextSearch(Session("og_parametri"), FieldLanguageList("obj_html_")) & ")"
end if

'filtra per pagina in cui è usato
if cIntero(Session("og_pagina"))>0 then
	
end if


'filtra per data inserimeto
if isDate(Session("og_data_ins_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "obj_insData", adCompareGreaterThan, Session("og_data_ins_from"))
end if
if isDate(Session("og_data_ins_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "obj_insData", adCompareLessThan, Session("og_data_ins_to"))
end if


'filtra per data modifica
if isDate(Session("og_data_mod_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "obj_modData", adCompareGreaterThan, Session("og_data_mod_from"))
end if
if isDate(Session("og_data_mod_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "obj_modData", adCompareLessThan, Session("og_data_mod_to"))
end if



sql = " SELECT id_objects, name_objects, obj_type, identif_objects FROM tb_Objects WHERE id_webs=" & Session("AZ_ID") & sql & " ORDER BY name_objects"
session("WEB_PLUGIN_SQL") = sql

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
										<td class="footer" nowrap>
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("og_nome") %>>NOME</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome" value="<%= TextEncode(session("og_nome")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("og_type") %>>TIPO</th></tr>
									<tr>
										<td class="content ascx" colspan="2">
											<input type="checkbox" class="checkbox" name="search_type" value="0" <%= chk(instr(1, session("og_type"), "0", vbTextCompare)>0) %>>
											plugin
										</td>
									</tr>
									<tr>
										<td class="content class" colspan="2">
											<input type="checkbox" class="checkbox" name="search_type" value="1" <%= chk(instr(1, Session("og_type"), "1", vbTextCompare)>0) %>>
											classe
										</td>
									</tr>
									<tr>
										<td class="content html" colspan="2">
											<input type="checkbox" class="checkbox" name="search_type" value="2" <%= chk(instr(1, Session("og_type"), "2", vbTextCompare)>0) %>>
											html
										</td>
									</tr>
									<tr><th <%= Search_Bg("og_classe") %>>CLASSE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_classe" value="<%= TextEncode(session("og_classe")) %>" style="width:100%;">
										</td>
									</tr>									
									<tr><th <%= Search_Bg("og_parametri") %>>PARAMETRI O HTML</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_parametri" value="<%= TextEncode(session("og_parametri")) %>" style="width:100%;">
										</td>
									</tr>
									<!--
									<tr><th <%= Search_Bg("og_pagina") %>>ID PAGINA</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_pagina" value="<%= TextEncode(session("og_pagina")) %>" style="width:100%;">
										</td>
									</tr>
									-->
									<tr><th colspan="2" <%= Search_Bg("og_data_ins_from;og_data_ins_to") %>>DATA INSERIMENTO</td></tr>
									<tr><td class="label" colspan="2">a partire dal:</td></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_ins_from", Session("og_data_ins_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><td class="label" colspan="2">fino al:</td></tr>
									<tr>
										<td class="content" colspan="2">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_ins_to", Session("og_data_ins_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("og_data_mod_from;og_data_mod_to") %>>DATA MODIFICA</td></tr>
									<tr><td class="label" colspan="2">a partire dal:</td></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_mod_from", Session("og_data_mod_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><td class="label" colspan="2">fino al:</td></tr>
									<tr>
										<td class="content" colspan="2">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_mod_to", Session("og_data_mod_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr>
										<td class="footer" nowrap>
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
						Elenco plugin
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovate n&ordm; <%= Pager.recordcount %> plugin del sito in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						
						dim N_PAGES, N_TEMPLATE, N_OBJECTS, values
						while not rs.eof and rs.AbsolutePage = Pager.PageNo						
							
							sql = " SELECT COUNT(*) FROM tb_layers INNER JOIN tb_pages ON ( tb_layers.id_pag = tb_pages.id_page AND " & SQL_IsTrue(conn, "tb_pages.template") & ") " & _
								  " WHERE id_webs=" & Session("AZ_ID") & _
										" AND ((tb_layers.id_objects=" & rs("id_objects") & ") OR " + _
								             " (tb_layers.testo LIKE '%=%" & ParseSQL(rs("name_objects"), adChar) & "%;%' AND tb_layers.id_tipo = " & LAYER_OBJECT & "))"				
							N_TEMPLATE = cIntero(GetValueList(conn, rsp, sql))
							
							
							values = ""
							sql = " SELECT DISTINCT "&SQL_IfIsNull(conn, "id_paginasito", "0")&" FROM tb_layers INNER JOIN tb_pages ON ( tb_layers.id_pag = tb_pages.id_page AND NOT " & SQL_IsTrue(conn, "tb_pages.template") & ") " & _
								  " WHERE id_webs=" & Session("AZ_ID") & _
										" AND ((tb_layers.id_objects=" & rs("id_objects") & ") OR " + _
											" (tb_layers.id_tipo = " & LAYER_OBJECT & " AND tb_layers.testo LIKE '%=%" & ParseSQL(rs("name_objects"), adChar) & "%;%')) "			
							values = GetValueList(conn, rsp, sql)
							if values = "" then
								N_PAGES = 0
							else
								sql = " SELECT COUNT(id_paginesito) FROM tb_paginesito " + _
									  " WHERE id_web=" & Session("AZ_ID") & " AND id_paginesito IN ("&values&")"							
								N_PAGES = cIntero(GetValueList(conn, rsp, sql))
							end if							
						
							sql = " SELECT COUNT(*) FROM tb_objects sub WHERE id_webs=" & Session("AZ_ID") & " AND sub.param_list LIKE '%=%" & ParseSQL(rs("name_objects"), adChar) & "%;%'"
							N_OBJECTS = cIntero(GetValueList(conn, rsp, sql))
							
							if (N_PAGES + N_TEMPLATE + N_OBJECTS) > 0 then
								isUsed = true								
							else
								isUsed = false
							end if 
							
							%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="<%= IIF(not isUsed, "header_disabled", "header") & " " & rs("obj_type")%>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="SitoPluginMod.asp?ID=<%= rs("id_objects") %>">
																MODIFICA
															</a>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('PLUGIN','<%= rs("id_objects") %>');" >
																CANCELLA
															</a>
														</td>
													</tr>
												</table>
												<%= rs("name_objects") %>
											</td>
										</tr>
										<% if not isUsed then %>
											<tr>
												<td class="content_center_disabled" colspan="4">
													<b>Plugin non utilizzato.</b>
												</td>
											</tr>
										<% else %>
											<tr>
												<td class="label_no_width" style="width:20%;">numero istanze:</td>
												<td colspan="3">
													<table width="100%" cellspacing="1" cellpadding="0">
														<% if N_PAGES>0 then %>
															<tr>
																<td class="label_no_width" style="width:18%;">nelle pagine:</td>
																<td class="content">
																	n&ordm; <b><%= N_PAGES%></b> istanze
																</td>
																<td class="content_right">
																	<% CALL WriteCampoCerca("SitoPagine.asp", "plugin", rs("id_objects"), "PAGINE COLLEGATE", "button_L2") %>
																</td>
															</tr>
														<% end if
														if N_TEMPLATE>0 then %>
															<tr>
																<td class="label_no_width" style="width:18%;">nei template:</td>
																<td class="content">
																	n&ordm; <b><%= N_TEMPLATE%></b> istanze
																</td>
																<td class="content_right">
																	<% CALL WriteCampoCerca("SitoTemplate.asp", "plugin", rs("id_objects"), "TEMPLATE COLLEGATI", "button_L2") %>
																</td>
															</tr>
														<% end if 
														if N_OBJECTS>0 then%>
															<tr>
																<td class="label_no_width" rowspan="2">nei plugin:</td>
																<td class="content">
																	n&ordm; <b><%=N_OBJECTS%></b> come parametro
																</td>
																<td class="content_right">
																	<% CALL WriteCampoCerca("SitoPlugin.asp", "parametri", rs("name_objects"), "PLUGIN CONTENITORI", "button_L2") %>
																</td>
															</tr>
														<% end if %>
													</table>
												</td>
											</tr>
										<% end if%>
										<tr>
											<td class="label_no_width" style="width:20%;">Classe sorgente:</td>
											<td class="content"><%= rs("identif_objects") %></td>
											<td class="label" style="width:17%;">id:</td>
											<td class="content"><%= rs("id_objects") %></td>
										</tr>
									</table>
								</td>
							</tr>
							<%
							rs.moveNext
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
set conn = nothing%>
