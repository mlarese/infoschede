<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "GRUPPI"
dicitura.links(1) = "CaratteristicheGruppi.asp"
dicitura.sezione = "Gestione caratteristhce - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA CARATTERISTICA"
dicitura.link_new = "Tabelle.asp;CaratteristicheNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, rsc, sql, Pager, TipCount
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("ctech_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("ctech_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

sql = ""
'filtra per nome
if Session("ctech_nome")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("ctech_nome"), FieldLanguageList("ct_nome_"))
end if

'filtra per tipo
if cInteger(session("ctech_tipo"))<>0 then
	sql = sql & " AND ct_tipo=" & cInteger(session("ctech_tipo"))
end if

'filtra per categoria
if session("ctech_categoria")<>"" then
	sql = sql & " AND ct_id IN (SELECT rcc_ctech_id FROM mrel_categ_ctech WHERE rcc_categoria_id=" & session("ctech_categoria") & ") "
end if

'filtra per raggruppamento
if CIntero(Session("ctech_raggruppamento_id")) > 0 then
	sql = sql & " AND ct_raggruppamento_id = "& Session("ctech_raggruppamento_id")
end if

'filtra per ricercabilita'
if Session("ctech_ricercabile")<>"" then
	if not (instr(1, Session("ctech_ricercabile"), "S", vbTextCompare)>0 AND instr(1, Session("ctech_ricercabile"), "N", vbTextCompare)>0 ) then
		if instr(1, Session("ctech_ricercabile"), "S", vbTextCompare)>0 then
			sql = sql &" AND "& SQL_IsTrue(conn, "ct_per_ricerca")
		elseif instr(1, Session("ctech_ricercabile"), "N", vbTextCompare)>0 then
			sql = sql & " AND NOT (" & SQL_IsTrue(conn, "ct_per_ricerca") & ") "
		end if
	end if
end if

'filtra per confrontabilita'
if Session("ctech_confrontabile")<>"" then
	if not (instr(1, Session("ctech_confrontabile"), "S", vbTextCompare)>0 AND instr(1, Session("ctech_confrontabile"), "N", vbTextCompare)>0 ) then
		if instr(1, Session("ctech_confrontabile"), "S", vbTextCompare)>0 then
			sql = sql &" AND "& SQL_IsTrue(conn, "ct_per_confronto")
		elseif instr(1, Session("ctech_confrontabile"), "N", vbTextCompare)>0 then
			sql = sql & " AND NOT (" & SQL_IsTrue(conn, "ct_per_confronto") & ") "
		end if
	end if
end if

sql = " SELECT *, (SELECT COUNT(*) FROM mrel_doc_ctech WHERE rdc_ctech_id=ct_id) AS N_DOC " + _
	  " FROM mtb_carattech " + _
	  " WHERE (1=1) " + sql + " ORDER BY ct_nome_it"
Session("MEMO2_CTECH_SQL") = sql

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
									<td class="footer" colspan="2">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("ctech_nome") %> colspan="2">NOME</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="text" name="search_nome" value="<%= TextEncode(session("ctech_nome")) %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("ctech_tipo") %> colspan="2">TIPO</th></tr>
								<tr>
									<td class="content" colspan="2">
										<% CALL  DesAdvancedDropTipi("search_tipo", " style=""width:100%"" ", cInteger(session("ctech_tipo")), false) %>
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("ctech_ricercabile") %>>RICERCABILIT&Agrave;</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_ricercabile" value="S" <%=chk(instr(1, Session("ctech_ricercabile"), "S", vbTextCompare)>0) %>>
										usato come filtro di ricerca
									</td>
								</tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_ricercabile" value="N" <%=chk(instr(1, Session("ctech_ricercabile"), "N", vbTextCompare)>0) %>>
										non utilizzato
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("ctech_confrontabile") %>>CONFRONTABILIT&Agrave;</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_confrontabile" value="S" <%=chk(instr(1, Session("ctech_confrontabile"), "S", vbTextCompare)>0) %>>
										usato nel confronto documenti
									</td>
								</tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_confrontabile" value="N" <%=chk(instr(1, Session("ctech_confrontabile"), "N", vbTextCompare)>0) %>>
										non utilizzato
									</td>
								</tr>
								<% if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM mtb_carattech_raggruppamenti"))>0 then %>
									<tr><th <%= Search_Bg("ctech_raggruppamento_id") %> colspan="2">GRUPPO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% sql = "SELECT * FROM mtb_carattech_raggruppamenti ORDER BY ctr_titolo_it"
							                CALL dropDown(conn, sql, "ctr_id", "ctr_titolo_it", "search_raggruppamento_id", Session("ctech_raggruppamento_id"), false, " style=""width:100%"" ", LINGUA_ITALIANO) %>
										</td>
									</tr>
								<% end if %>
								<tr><th <%= Search_Bg("ctech_categoria") %>>CATEGORIA A CUI &Egrave; ASSOCIATA</th></tr>
								<tr>
									<td class="content" colspan="2">
										<% CALL categorie.WritePicker("ricerca", "search_categoria", session("ctech_categoria"), true, true, 32) %>
									</td>
								</tr>
								<tr>
									<td class="footer" colspan="2">
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
						Elenco caratteristiche
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
															<a class="button" href="CaratteristicheMod.asp?ID=<%= rs("ct_id") %>">
																MODIFICA
															</a>
															&nbsp;
															<% if rs("N_DOC") > 0 then %>
																<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la caratteristica: sono presenti documenti associati">
																	CANCELLA
																</a>
															<% else %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CTECH','<%= rs("ct_id") %>');" >
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												<%= rs("ct_nome_it") %>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:22%;">tipo di dato:</td>
											<td class="content"><%= DesVisTipo(rs("ct_tipo")) %></td>
										</tr>
										<tr>
											<td class="label">categorie associate:</td>
											<td class="content" style="max-height:200px; overflow-y:scroll; float:left; width:98%;">
												<% sql = "tip_L0.catC_id IN (SELECT rcc_categoria_id FROM mrel_categ_ctech WHERE rcc_ctech_id=" & rs("ct_id") & ")"
												sql = categorie.QueryElenco(false, sql)
												rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext 
												TipCount = 0
												if rsc.eof then%>
													Non associata a nessuna categoria.
												<% else%>
													<span id="categorie_<%= rs("ct_id")%>">
													<%while not rsc.eof
														TipCount = TipCount + 1 %>
														<%= rsc("NAME") %>
														<% rsc.movenext
														if not rsc.eof then %>
															<br>
														<%end if
													wend %>
													</span>
													<%if TipCount > 4 then %>
														<style type="text/css">
															#categorie_<%= rs("ct_id")%>{
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
						<tr><th class="noRecords">Nessun record trovato</th></tr>
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