<%@ Language=VBScript CODEPAGE=65001 %>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->

<% 	
dim dicitura
set dicitura = New testata 
dicitura.sezione = "Gestione guasti - elenco"
dicitura.iniz_sottosez(0)
dicitura.puls_new = "INSERISCI GUASTO"
dicitura.link_new = "ProblemiNew.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, last, Pager, sql, rsa, rows, color, dataStart
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("prb_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("prb_")
	end if
end if

'filtra per natura guasto
if(Session("prb_natura")<>"") then
	sql = sql & "AND (1=0"
	if inStr(Session("prb_natura"),"0")>0 then
		sql = sql & " OR " & SQL_IsTrue(conn, "prb_riscontrato")
	end if
	if inStr(Session("prb_natura"),"1")>0 then
		sql = sql & " OR NOT " & SQL_IsTrue(conn, "prb_riscontrato")
	end if
	sql=sql & ")"
end if

'ricerca per nome/titolo
if Session("prb_titolo")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("prb_titolo"), FieldLanguageList("prb_nome_"))
end if

'ricerca per descrizione
if Session("prb_descrizione")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("prb_descrizione"), FieldLanguageList("prb_descrizione_"))
end if

'ricerca per avviso per conferma
if Session("prb_avviso")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("prb_avviso"), FieldLanguageList("prb_avviso_per_conferma_"))
end if

'ricerca per visibilità
if Session("prb_visibile")<>"" then
	if not (instr(1, Session("prb_visibile"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("prb_visibile"), "0", vbTextCompare)>0 ) then
		sql = sql & " AND "
		if instr(1, Session("prb_visibile"), "1", vbTextCompare)>0 then
			'articolo a catalogo
			sql = sql & " ISNULL(prb_visibile, 0)=0 "
		elseif instr(1, Session("prb_visibile"), "0", vbTextCompare)>0 then
			'articolo fuori catalogo
			sql = sql & " ISNULL(prb_visibile, 0)=1 "
		end if
	end if
end if

'ricerca per avviso per profilo
if Session("prb_profilo")<>"" then
	sql = sql & " AND prb_id IN (SELECT rpp_problema_id FROM srel_problemi_profili WHERE rpp_profilo_id = " & Session("prb_profilo") & ")"
end if

'ricerca per avviso per marchio
if Session("prb_marca")<>"" then
	sql = sql & " AND prb_id IN (SELECT rpm_problema_id FROM srel_problemi_mar_tip WHERE rpm_marchio_id = " & Session("prb_marca") & ")"
end if

'ricerca per avviso per categoria
if Session("prb_categoria")<>"" then
	sql = sql & " AND prb_id IN (SELECT rpm_problema_id FROM srel_problemi_mar_tip WHERE rpm_tipologia_id = " & Session("prb_categoria") & ")"
end if

'ricerca per avviso per categoria
if Session("prb_articolo")<>"" then
	sql = sql & " AND prb_id IN (SELECT rpa_problema_id FROM srel_problemi_articoli WHERE rpa_articolo_rel_id = " & Session("prb_articolo") & ")"
end if


'filtra per tipo autorizzazione
if(Session("prb_modalita")<>"") then
	sql = sql & "AND (1=0"
	if inStr(Session("prb_modalita"),"0")>0 then
		sql = sql & " OR " & SQL_IsTrue(conn, "prb_modalita_easy")
	end if
	if inStr(Session("prb_modalita"),"1")>0 then
		sql = sql & " OR NOT " & SQL_IsTrue(conn, "prb_modalita_easy")
	end if
	sql=sql & ")"
end if


sql = " SELECT * FROM sgtb_problemi WHERE (1=1) " + sql + " ORDER BY prb_nome_it"


Session("ELENCO_PROBLEMI_SQL") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 8)

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
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("prb_natura") %>>GUASTO</th></tr>
								<tr>
									<td class="content riscontrato" colspan="2">
										<input type="checkbox" class="checkbox" name="search_natura" value="0" <%= chk(instr(1, session("prb_natura"), "0", vbTextCompare)>0) %>>
										riscontrato
									</td>
								</tr>
								<tr>
									<td class="content segnalato" colspan="2">
										<input type="checkbox" class="checkbox" name="search_natura" value="1" <%= chk(instr(1, Session("prb_natura"), "1", vbTextCompare)>0) %>>
										segnalato
									</td>
								</tr>
								
								<tr><th colspan="2" <%= Search_Bg("prb_titolo") %>>TITOLO GUASTO</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="text" name="search_titolo" value="<%= TextEncode(session("prb_titolo")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("prb_descrizione") %>>DESCRIZIONE</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="text" name="search_descrizione" value="<%= TextEncode(session("prb_descrizione")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("prb_avviso") %>>AVVISO PER CONFERMA</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="text" name="search_avviso" value="<%= TextEncode(session("prb_avviso")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("prb_modalita") %>>MODALITA'</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_modalita" value="0" <%= chk(instr(1, session("prb_modalita"), "0", vbTextCompare)>0) %>>
										<% WriteColor(MODALITA_EASY_COLOR)%><%=MODALITA_EASY%>
									</td>
								</tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_modalita" value="1" <%= chk(instr(1, Session("prb_modalita"), "1", vbTextCompare)>0) %>>
										<% WriteColor(MODALITA_NON_EASY_COLOR)%><%=MODALITA_NON_EASY%>
									</td>
								</tr>
								
								<tr><th colspan="2" <%= Search_Bg("prb_visibile") %>>VISIBILITA'</th></tr>
								<tr>
									<td class="content" style="width:45%;">
										<input type="checkbox" class="checkbox" name="search_visibile" value="1" <%= chk(instr(1, session("prb_visibile"), "1", vbTextCompare)>0) %>>
										<b>visibile</b>
									</td>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_visibile" value="0" <%= chk(instr(1, Session("prb_visibile"), "0", vbTextCompare)>0) %>>
										non visibile
									</td>
								</tr>
								
								<% sql = "SELECT * FROM gtb_profili ORDER BY pro_nome_it" %>
								<% if GetValueList(conn, NULL, sql)<>"" then %>
									<tr><th colspan="2" <%= Search_Bg("prb_profilo") %>>PROFILO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL dropDown(conn, sql, "pro_id", "pro_nome_it", "search_profilo", session("prb_profilo"), false, "style=""width: 100%;""", Session("LINGUA")) %>
										</td>
									</tr>
								<% end if %>
								
								<% sql = "SELECT * FROM gtb_marche ORDER BY mar_nome_it" %>
								<% if GetValueList(conn, NULL, sql)<>"" then %>
									<tr><th colspan="2" <%= Search_Bg("prb_marca") %>>MARCHIO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL dropDown(conn, sql, "mar_id", "mar_nome_it", "search_marca", session("prb_marca"), false, "style=""width: 100%;""", Session("LINGUA")) %>
										</td>
									</tr>
								<% end if %>
								
								<% sql = "SELECT * FROM gtb_tipologie ORDER BY tip_nome_it" %>
								<% if GetValueList(conn, NULL, sql)<>"" then %>
									<tr><th colspan="2" <%= Search_Bg("prb_categoria") %>>CATEGORIA</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL cat_modelli.WritePicker("ricerca", "search_categoria", session("prb_categoria"), false, true, 32) %>
										</td>
									</tr>
								<% end if %>
								
								<tr><th colspan="2" <%= Search_Bg("prb_articolo") %>>MODELLO</th></tr>
								<tr>
									<td class="content" colspan="2">
										<% CALL WritePicker_ArticoloVariante(conn, rsa, "ricerca", "search_articolo", session("prb_articolo"), 12, false, "Infoschede/ArticoliSeleziona.asp?TYPE=M&") %>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="footer">
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
						Elenco guasti
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> guasti in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
                            <tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0">
										<tr>
											<td class="<%=IIF(rs("prb_visibile"), "header ", "header_disabled ")%> <%=IIF(rs("prb_riscontrato"), "riscontrato", "segnalato")%>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
                                                    <tr>
                                                        <td style="font-size: 1px;">
                                                            <a class="button" href="ProblemiMod.asp?ID=<%= rs("prb_id") %>">MODIFICA</a>
                                                            &nbsp;
                                                            <% if false then %>
                        										<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il marchio: sono presenti anagrafiche che utilizzano questo profilo">
                        											CANCELLA
                        										</a>
                        									<% else %>
                        										<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('PROBLEMI','<%= rs("prb_id") %>');" >
                        											CANCELLA
                        										</a>
                        									<% end if %>
														</td>
													</tr>
												</table>
												<%= rs("prb_nome_it") %>
											</td>
										</tr>
										<% if not rs("prb_riscontrato") then %>
											<tr>
												<td class="label_no_width" style="width:18%;">modalit&agrave;:</td>
												<td class="content" colspan="3">
													<% if rs("prb_modalita_easy") then %>
														<% WriteColor(MODALITA_EASY_COLOR)%><%=MODALITA_EASY%>
													<% else%>
														<% WriteColor(MODALITA_NON_EASY_COLOR)%><%=MODALITA_NON_EASY%>
													<% end if %>
												</td>
											</tr>
										<% end if %>
                                        <tr>
                                            <td class="label_no_width" style="width:18%;">visibilit&agrave;:</td>
                                            <td class="content"  style="width:45%;"><%= IIF(rs("prb_visibile"),"","non ") %>visibile</td>
											<td class="label_no_width" style="width:15%;">ordine:</td>
                                            <td class="content"><%= rs("prb_ordine")%></td>
										</tr>
										<% dim writed_assoc 
										writed_assoc = false
										
										sql = " SELECT pro_nome_it FROM gtb_profili WHERE pro_id IN " & _
												 " (SELECT rpp_profilo_id FROM srel_problemi_profili WHERE rpp_problema_id = "&rs("prb_id")&")"
										if GetValueList(conn, NULL, sql)<>"" then %>
											<% if not writed_assoc then %>
												<% writed_assoc = true %>
												<tr><th class="L2" colspan="4">ASSOCIAZIONI</th></tr>
												<tr>
												<td class="label_no_width">profili:</td>
												<td class="content" colspan="3"><%=GetValueList(conn, NULL, sql)%></td>
											<% end if %>
										<% end if %>
										
										<%
										sql = " SELECT *, ISNULL((SELECT mar_nome_it FROM gtb_marche WHERE mar_id=rpm_marchio_id),'Tutte le marche') AS MARCA, " & _
												 " ISNULL((SELECT tip_nome_it FROM gtb_tipologie WHERE tip_id=rpm_tipologia_id),'Tutti le categorie') AS TIPOLOGIA " & _
												 " FROM srel_problemi_mar_tip " & _
												 " WHERE rpm_problema_id = " & rs("prb_id") & _
												 " ORDER BY MARCA "
										rsa.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
										<% if not rsa.eof then %>
											<% if not writed_assoc then %>
												<% writed_assoc = true %>
												<tr><th class="L2" colspan="4">ASSOCIAZIONI</th></tr>
											<% end if %>
											<tr>
												<td class="label_no_width">marca / tipologia:</td>
												<td class="content" colspan="3">
													<%=IIF(cIntero(rsa("rpm_marchio_id"))=0,"<b>","")%><%=rsa("MARCA")%><%=IIF(cIntero(rsa("rpm_marchio_id"))=0,"</b>","")%>
													&nbsp;/&nbsp;
													<%=IIF(cIntero(rsa("rpm_tipologia_id"))=0,"<b>","")%><%=rsa("TIPOLOGIA")%><%=IIF(cIntero(rsa("rpm_tipologia_id"))=0,"</b>","")%>
												</td>
											</tr>
											<% rsa.moveNext %>
										<% end if %>
										<% while not rsa.eof %>
											<tr>
												<td class="label_no_width">&nbsp;</td>
												<td class="content" colspan="3">
													<%=IIF(cIntero(rsa("rpm_marchio_id"))=0,"<b>","")%><%=rsa("MARCA")%><%=IIF(cIntero(rsa("rpm_marchio_id"))=0,"</b>","")%>
													&nbsp;/&nbsp;
													<%=IIF(cIntero(rsa("rpm_tipologia_id"))=0,"<b>","")%><%=rsa("TIPOLOGIA")%><%=IIF(cIntero(rsa("rpm_tipologia_id"))=0,"</b>","")%>
												</td>
											</tr>
											<% rsa.moveNext %>
										<% wend %>
										<% rsa.close %>
										
										<% 
										sql = " SELECT * FROM gv_articoli INNER JOIN srel_problemi_articoli " & _
											  " ON gv_articoli.rel_id = srel_problemi_articoli.rpa_articolo_rel_id " & _
											  " WHERE rpa_problema_id = " & rs("prb_id") & " ORDER BY art_nome_it "
										rsa.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
										<% if not rsa.eof then %>
											<% if not writed_assoc then %>
												<% writed_assoc = true %>
												<tr><th class="L2" colspan="4">ASSOCIAZIONI</th></tr>
											<% end if %>
											<tr>
												<td class="label_no_width">modelli:</td>
												<td class="content" colspan="3">
													<% CALL ArticoloLink(rsa("art_id"), rsa("art_nome_it"), rsa("art_cod_int")) %>
													<% if rsa("art_varianti") then %>
														<%= ListValoriVarianti(conn, NULL, rsa("rel_id")) %>
													<% else %>
														&nbsp;
													<% end if %>
													<span>&nbsp;(<%=rsa("tip_nome_it")%>)</span>
												</td>
											</tr>
											<% rsa.moveNext %>
										<% end if %>
										<% while not rsa.eof %>
											<tr>
												<td class="label_no_width">&nbsp;</td>
												<td class="content" colspan="3">
													<% CALL ArticoloLink(rsa("art_id"), rsa("art_nome_it"), rsa("art_cod_int")) %>
													<% if rsa("art_varianti") then %>
														<%= ListValoriVarianti(conn, NULL, rsa("rel_id")) %>
													<% else %>
														&nbsp;
													<% end if %>
													&nbsp;(<%=rsa("tip_nome_it")%>)
												</td>
											</tr>
											<% rsa.moveNext %>
										<% wend %>
										<% rsa.close %>
										
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td class="footer" style="text-align:left;" colspan="6">
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
set rsa = nothing
set conn = nothing
%>
