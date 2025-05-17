<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Documenti - elenco"
dicitura.puls_new = ""
dicitura.link_new = ""
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, Pager, Found, CI_Current

set Pager = new PageNavigator

if cString(request("CAT_ID"))<>"" then
	if cString(request("CAT_ID")) = "0" then
		Session("DOWNLOAD_CAT_ID") = ""
	else
		Session("DOWNLOAD_CAT_ID") = cString(request("CAT_ID"))
	end if
end if

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("Ddoc_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("Ddoc_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


sql = " SELECT * FROM mtb_documenti WHERE " & SQL_isTrue(conn, "doc_visibile") & " AND " & _
												"(" & SQL_now(conn) & " BETWEEN doc_pubblicazione AND ISNULL(doc_scadenza," & SQL_now(conn) & ")) "



'filtra per protezione documento solo per user download
if Session("MEMO2_DOWNLOAD") <> "" then
	sql = sql + " AND (((doc_id IN (SELECT rda_doc_id FROM mrel_doc_admin WHERE rda_admin_id = " & Session("ID_ADMIN") & ")) " + _
				" 		OR (doc_id IN (SELECT rdp_doc_id FROM mrel_doc_profili WHERE rdp_profilo_id IN (SELECT rpa_profilo_id FROM mrel_profili_admin WHERE rpa_admin_id = " & Session("ID_ADMIN") & "))))" + _ 
				" OR NOT (" & SQL_isTrue(conn, "doc_protetto") & "))"
end if

'ricerca full text sul contenuto
if Session("Ddoc_testo")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("Ddoc_testo"), FieldLanguageList("doc_titolo_"))
	sql = sql & " OR " & SQL_FullTextSearch(Session("Ddoc_testo"), FieldLanguageList("doc_estratto_")) & ")"
end if

'filtra sul numero / protocollo
if Session("Ddoc_numero")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("Ddoc_numero"), "doc_numero")
end if


'filtra per categoria
if cString(Session("DOWNLOAD_CAT_ID"))<>"" then
	sql = sql & " AND doc_categoria_id IN (" & cString(Session("DOWNLOAD_CAT_ID")) & " )"
end if

'filtra per categoria
'if Session("Ddoc_categoria")<>"" then
'	sql = sql & " AND doc_categoria_id IN (" & categorie.FoglieID(Session("Ddoc_categoria")) & " )"
'end if


'filtra per data di pubblicazione
if isDate(Session("Ddoc_data_from")) then
	sql = sql & " AND doc_pubblicazione >=" & SQL_date(conn, Session("Ddoc_data_from"))
end if
if isDate(Session("Ddoc_data_to")) then
	sql = sql & " AND doc_pubblicazione <=" & SQL_date(conn, Session("Ddoc_data_to"))
end if

sql = sql & " ORDER BY doc_pubblicazione DESC, doc_titolo_it"

Session("SQL_DOWNLOAD") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>

<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<form action="Download.asp" method="post" id="ricerca" name="ricerca">
					<tr>
						<td>
							<% if cBoolean(Session("CATEGORIE_NEXTMEMO2_ABILITATE"), false) then %>
								<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px;">
									<caption class="border">Categorie</caption>
									<%  dim rs_cat
										set rs_cat = Server.CreateObject("ADODB.RecordSet")
										rs_cat.open categorie.QueryElenco(true, ""), conn, adOpenStatic, adLockReadOnly, adCmdText
										
										while not rs_cat.eof
											%>
											<tr>
												<td class="content" style="padding-top:2px; padding-bottom:2px; font-weight:bold">
													<a href="Download.asp?CAT_ID=<%=rs_cat("catC_id")%>" <%=IIF(rs_cat("catC_id")=cIntero(Session("DOWNLOAD_CAT_ID")), "style='color:#ff9a00; text-transform:uppercase;'","")%>>
														<%= rs_cat("NAME") %>
													</a>
												</td>
											</tr>
											<%
											rs_cat.moveNext
										wend 
										rs_cat.close
										set rs_cat = nothing
										%>
										<tr>
											<td class="label" style="text-align:center;">
												<a class="button_L2_block" style="width:100%;text-align:center;" href="Download.asp?CAT_ID=0">VEDI TUTTE</a>
											</td>
										</tr>
								</table>
							<% end if %>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTE" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("Ddoc_testo") %>>TITOLO ED ESTRATTO</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_testo" value="<%= replace(session("Ddoc_testo"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("Ddoc_numero") %>>NUMERO / PROTOCOLLO</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_numero" value="<%= replace(session("Ddoc_numero"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>								   
								<tr><th <%= Search_Bg("Ddoc_data_from;Ddoc_data_to") %>>DATA DI PUBBLICAZIONE</th></tr>
								<tr><td class="label">a partire dal:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_from", Session("Ddoc_data_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td class="label">fino al:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_to", Session("Ddoc_data_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr>
									<td class="footer">
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
<!-- BLOCCO RISULTATI -->
				<% if not rs.eof then 
					if cInteger(request("ID"))>0 then
						'cerca circolare selezionata
						rs.Find "doc_id=" & request("ID")
						Found = TRUE
						if rs.eof then
							'circolare non trovata
							Found = FALSE
							rs.MoveFirst
						end if
					else
						Found = FALSE
						rs.MoveFirst
					end if
					'salva ID circolare visualizzata
					CI_Current = rs("doc_id")
					%>
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
<!-- CIRCOLARE SELEZIONATA -->
						<tr>
							<td style="padding-bottom:10px;">
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
									<caption>
										<% if found then %>
											Documento / circolare selezionata.
										<% else %>
											Documento / circolare pi&ugrave; recente.
										<% end if %>
									</caption>
									<tr><th colspan="2" class="center"><%= rs("doc_titolo_it") %></th></tr>
									<tr>
										<td class="content">
											<table cellspacing="0" cellpadding="0">
												<tr>
													<% if rs("doc_numero")<>"" then %>
														<td class="label" style="width:45px;">n&ordm; / prot.</td>
														<td class="content_b" nowrap ><%= rs("doc_numero")%></td>
													<% end if %>
													<td class="label" style="width:18px;">&nbsp;del</td>
													<td class="content_b"><%= DateIta(rs("doc_pubblicazione")) %> </td>
													<% if cBoolean(Session("CATEGORIE_NEXTMEMO2_ABILITATE"), false) then %>
														<td class="label" style="width:28px;">&nbsp;&nbsp;categoria</td>
														<% sql = "SELECT catC_nome_it FROM mtb_documenti_categorie WHERE catC_id = " & rs("doc_categoria_id") %>
														<td class="content_b"><%= GetValueList(conn, NULL, sql) %></td>
													<% end if %>
												</tr>
											</table>
										</td>
										<td class="content_right" style="width:21%;">
											<% CALL write_download(rs("doc_id"), rs("doc_file_it")) %>
												Download del file.
											</a>
										</td>
									</tr>
									<tr>
										<td class="content_center" colspan="2" style="padding:3px; padding-left:10px;padding-right:10px;">
											<iframe name="estratto" width="100%" height="100" align="left" frameborder="0">
											</iframe>
										</td>
									</tr>
									<script language="JavaScript" type="text/javascript">
										estratto.document.open();
										estratto.document.write('<html></head><link rel="stylesheet" type="text/css" href="../library/stili.css"><body leftmargin=0 topmargin=0 class="IFRAME">');
										estratto.document.write('<div style="text-align: justify; padding-right:5px;"><%=JsEncode(rs("doc_estratto_it")&"", "'")%></div>');
										estratto.document.write('</body></html>');
										estratto.document.close();
										</script>
								</table>
							</td>
						</tr>
<!-- ELENCO CIRCOLARI -->
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
									<caption>Elenco documenti - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
									<tr>
										<th>DOCUMENTO / CIRCOLARE</th>
										<th style="width:15%;">PROT.</th>
										<th style="width:12%;" class="center">DATA</th>
										<th style="width:13%;" class="center">DOWNLOAD</th>
									</tr>
									<% if not rs.eof then 
										rs.AbsolutePage = Pager.PageNo
										while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
											<tr>
												<td class="content">
													<%if rs("doc_id") = CI_Current then		'circolare gia' visualizzata
													%>
														<a class="content_selected" title="Documento / circolare visualizzati">
													<%else									'altra circolare
													%>
														<a href="?ID=<%=rs("doc_id")%>" class="content" title="Apri scheda completa">
													<%end if%>
														<%=Server.HTMLEncode(rs("doc_titolo_it")&"")%>
													</a>
												</td>
												<td class="content">
													<% if rs("doc_numero")<>"" then %>
														n&ordm; <%= rs("doc_numero") %>
													<% end if %>
													&nbsp;
												</td>
												<td class="content_center"><%= DateIta(rs("doc_pubblicazione")) %></td>
												<td class="Content_center">
													<%if rs("doc_file_it")<>"" then
														if rs("doc_id") = CI_Current then		'circolare gia' visualizzata
														%>
															<a class="content_selected" title="Documento / circolare visualizzati">
														<%else									'altra circolare
															CALL write_download(rs("doc_id"), rs("doc_file_it"))
														end if%>
															Download
														</a>
													<%else%>
														&nbsp;
													<% end if %>
												</td>
											</tr>
											<% rs.moveNext
										wend%>
										<tr>
											<td colspan="6" class="footer" style="text-align:left;">
												<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
											</td>
										</tr>
									<%else%>
										<tr><td class="noRecords">Nessun record trovato</th></tr>
									<% end if %>		
								</table
							</td>
						</tr>
					</table>
				<% else %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption>
							Elenco documenti / circolari
						</caption>
						<tr><td class="noRecords">Nessun record trovato</th></tr>
					</table>
				<% end if %>
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

sub write_download(ID, file)%>
	<a class="content" title="Click per fare il download del file." 
		<% if file<>"" then %>
			href="DownloadFile.asp?ID=<%= ID %>&DIP=<%= Session("LOGIN_4_LOG") %>">
		<% else %>
			>
		<% end if %>
<% end sub %>
