<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="IndexMetaTag_TOOLS.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_indice_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.sezione = "Indice generale"
dicitura.puls_new = "nuovo:;RAGGRUPPAMENTO;VOCE"
dicitura.link_new = ";IndexRaggruppamentoGestione.asp?FROM="& FROM_ELENCO &";IndexGestione.asp?FROM=" & FROM_ELENCO

if index.ChkPrm(prm_Pubblicazioni_accesso, 0) then
	dicitura.iniz_sottosez(2)
	dicitura.sottosezioni(2) = "PUBBLICAZIONI AUTOMATICHE"
	dicitura.links(2) = "IndexPubblicazioni.asp"
else
	dicitura.iniz_sottosez(1)
end if
dicitura.sottosezioni(1) = "META TAG"
dicitura.links(1) = "IndexMetaTag.asp?FROM=Indice"

dicitura.scrivi_con_sottosez()

dim sql, rs, Pager
set Pager = new PageNavigator
set rs = Server.CreateObject("ADODB.RecordSet")
    	

sql = IndexSearchEngineSetFilter(Index.conn, false)

sql = " SELECT idx_id, idx_autopubblicato, co_chiave_it, idx_livello, visibile_assoluto, co_titolo_it, tab_colore, tab_titolo FROM v_indice " + _
	  IIF(sql<>"", " WHERE " & sql, "") + _
	  " ORDER BY co_titolo_it"
Session("IDX_SQL") = sql

CALL Pager.OpenSmartRecordset(Index.conn, rs, sql, 10) %>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0">
		<tr>
	  		<td style="width:27%;" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
					<form action="" method="post" id="ricerca" name="ricerca">
					<caption>Opzioni di ricerca</caption>
					<tr>
						<td class="footer" colspan="2">
							<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
							<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("idx_chiave") %>>CODICE UNIVOCO</th></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="text" name="search_chiave" value="<%= TextEncode(session("idx_chiave")) %>" style="width:100%;">
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("idx_titolo") %>>TITOLO</th></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="text" name="search_titolo" value="<%= TextEncode(session("idx_titolo")) %>" style="width:100%;">
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("idx_id") %>>ID</th></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="text" name="search_id" value="<%= TextEncode(session("idx_id")) %>" style="width:100%;">
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("idx_visibile") %>>VISIBILE</th></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="checkbox" class="checkbox" name="search_visibile" value="1" <%= chk(instr(1, Session("idx_visibile"), "1", vbTextCompare)>0) %>>
							visibile
						</td>
					</tr>
					<tr>
						<td class="content" colspan="2">
							<input type="checkbox" class="checkbox" name="search_visibile" value="0" <%= chk(instr(1, session("idx_visibile"), "0", vbTextCompare)>0) %>>
							non visibile
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("idx_livello") %>>LIVELLO DELLA VOCE</th></tr>
					<tr>
						<td class="content" colspan="2">
						<% 	sql = "SELECT MAX(idx_livello) FROM tb_contents_index"
							dim levels, i
							set levels = Server.CreateObject("Scripting.Dictionary")
							CALL levels.Add("0", "Voci base")
							for i=1 to cInteger(GetValueList(Index.conn, NULL, sql))
								CALL levels.Add(cString(i), "Voci livello " & i)
							next
							CALL DropDownDictionary(levels, "search_livello", Session("idx_livello"), false, "style=""width:100%;""", LINGUA_ITALIANO)%>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("idx_tipoContenuto") %>>TIPO DELLA VOCE</th></tr>
					<tr>
						<td class="content" colspan="2">
						<% CALL index.content.DropDownTipi("search_tipoContenuto", "", session("idx_tipoContenuto")) %>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("idx_categoria") %>>VOCI COLLEGATE A</th></tr>
					<tr>
						<td class="content" colspan="2">
						<% CALL index.WritePicker("", "", "ricerca", "search_categoria", session("idx_categoria"), 0, false, true, 32, false, false) %>
						</td>
					</tr>
					<tr>
						<td class="footer" colspan="2">
							<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
							<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
						</td>
					</tr>
					</form>
				</table>
				<% if Session("WEB_ADMIN")<>"" then %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption class="border">Strumenti</caption>
						<tr>
							<td class="content_center">
								<a class="button_block" target="_blank"
								   title="Apre la palette di export dei dati"
								   href="IndexTabellaContenuti.asp">
								   	ESPORTA COME TABELLA
								</a>
							</td>
						</tr>
					</table>
				<% end if %>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
				
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
					<caption class="border">
						Indice generale - albero
					</caption>
					<tr>
						<td class="content">
							Visualizza l'indice generale come albero:
						</td>
						<td class="content_right">
							<a class="button" href="IndexAlbero.asp" title="Apre la visualizzazione ad albero.">
								VISUALIZZA COME ALBERO
							</a>
						</td>
					</tr>
				</table>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Indice generale - elenco
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
															<a class="button" href="IndexSottosezioni.asp?ID=<%= rs("idx_id") %>" title="Apre l'elenco delle sotto-voci." <%= ACTIVE_STATUS %>>
																VOCI COLLEGATE
															</a>
															&nbsp;
															<a class="button" href="IndexGestione.asp?ID=<%= rs("idx_id") %>&FROM=<%= FROM_ELENCO %>">
																MODIFICA
															</a>
															&nbsp;
															<% if rs("idx_autopubblicato") then %>
																<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la voce perch&egrave; fa parte delle seguenti pubblicazioni automatiche:<%= vbCrLF & index.GetPubblicazioniLockers(rs("idx_id")) %>."<%= ACTIVE_STATUS %> >
																	CANCELLA
																</a>
															<% else 
                                                                   CALL index.WriteDeleteButton("", rs("idx_id"))
                                                               end if %>
														</td>
													</tr>
												</table>
												<% CALL index.content.WriteNomeETipo(rs) %>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:23%;">posizione nell'indice:</td>
											<td class="content" colspan="3"><%= index.NomeCompleto(rs("idx_id")) %></td>
										</tr>
										<tr>
											<td class="label">codice univoco:</td>
											<td class="content"><%= rs("co_chiave_it") %></td>
											<td class="label">livello:</td>
											<td class="content" style="width:30%;">
												<% if rs("idx_livello")=0 then %>
													voce base
												<% else %>
													voce livello <%= rs("idx_livello") %>
												<% end if %>
											</td>
										</tr>
										<tr>
											<td class="label">visibile:</td>
											<td class="content"><input type="checkbox" class="checkbox" disabled <%= chk(rs("visibile_assoluto")) %>></td>
											<td class="label">id:</td>
											<td class="content"><%= rs("idx_id") %></td>
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
<% rs.close


set rs = nothing
set index = nothing
%>