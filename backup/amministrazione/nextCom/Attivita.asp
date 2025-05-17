<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
dim conn, rs, rsp, sql, sqlp, Pager, N_DOCS
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	if Session("ATT_PRA_ID")<>"" then
		CALL GotoRecord(conn, rs, Session(Session("PRA_PREFIX") & "SQL_PRATICHE"), "pra_id", "Attivita.asp")
	elseif Session("ATT_DOC_ID")<>"" then
		CALL GotoRecord(conn, rs, Session(Session("DOC_PREFIX") & "SQL_DOCUMENTI"), "doc_id", "Attivita.asp")
	end if
end if

if request.QueryString("all")<>"" then		'se richiesta arriva da menu: visualizza tutte le attivita
	'da pratiche
	Session("ATT_PRA_ID") = ""
	Session("ATT_PRA_NOME") = ""
	Session("ATT_PRA_CNT_NOME") = ""
	'da documenti
	Session("ATT_DOC_ID") = ""
	Session("ATT_DOC_NOME") = ""
	
	Session("ATT_PREFIX") = ""
	response.redirect "Attivita.asp"
elseif request.Querystring("PRA_ID")<>"" OR _
	(request.querystring("ID")<>"" AND Session("ATT_PRA_ID")<>"") then	
					'se richiesta arriva da pratiche: visualizza tutte le attivita' della pratica
	'da pratiche
	Session("ATT_PRA_ID") = IIF(request.Querystring("PRA_ID")<>"", request.Querystring("PRA_ID"), request.querystring("ID"))
	Session("ATT_PRA_NOME") = ""
	Session("ATT_PRA_CNT_NOME") = ""
	'da documenti
	Session("ATT_DOC_ID") = ""
	Session("ATT_DOC_NOME") = ""
	
	Session("ATT_PREFIX") = "ATT_"
	response.redirect "Attivita.asp"
elseif request.Querystring("DOC_ID")<>"" OR _
	(request.querystring("ID")<>"" AND Session("ATT_DOC_ID")<>"") then	
					'se richiesta arriva da documenti: visualizza tutte le attivita' collegate al documento
	'da pratiche
	Session("ATT_PRA_ID") = ""
	Session("ATT_PRA_NOME") = ""
	Session("ATT_PRA_CNT_NOME") = ""
	'da documenti
	Session("ATT_DOC_ID") = IIF(request.Querystring("DOC_ID")<>"", request.Querystring("DOC_ID"), request.querystring("ID"))
	Session("ATT_DOC_NOME") = ""
	
	Session("ATT_PREFIX") = "DOC_"
	response.redirect "Attivita.asp"
end if


'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
if Session("ATT_PRA_ID")<>"" then
	'attivita' della pratica
	Titolo_sezione = "Pratiche - attivit&agrave; della pratica"
	HREF = "Pratiche.asp;PraticaMod.asp?ID=" & Session("ATT_PRA_ID") & ";Documenti.asp?pra_id=" & Session("ATT_PRA_ID") & ";"
	Action = "INDIETRO;SCHEDA PRATICA;DOCUMENTI PRATICA;"
elseif Session("ATT_DOC_ID")<>"" then
	'attivita' collegate al documento
	Titolo_sezione = "Documenti - attivit&agrave; collegate al documento"
	HREF = "Documenti.asp;DocumentoMod.asp?ID=" & Session("ATT_DOC_ID") & ";"
	Action = "INDIETRO;SCHEDA DOCUMENTO;"
else
	Titolo_sezione = "Attivit&agrave; - elenco"
	HREF = ""
	Action = ""
end if
HREF = HREF & "AttivitaNew.asp"
Action = Action & "NUOVA ATTIVITA'"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************


dim Prefix, ClassHeader
Prefix = Session("ATT_PREFIX")

'imposta criteri iniziali
if Session(Prefix & "SQL_ATTIVITA")="" then
	Session(Prefix & "att_stato") = "A"
end if

Session("PageNavigator_VarPrefix") = Prefix
set Pager = new PageNavigator
'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	if request("tutti")<>"" then
		Session(Prefix & "att_stato") = ""
		Session(Prefix & "att_tipo") = ""
		Session(Prefix & "att_pratica") = ""
		Session(Prefix & "att_contatto") = ""
		Session(Prefix & "att_oggetto") = ""
		Session(Prefix & "att_data_creazione_from") = ""
		Session(Prefix & "att_data_creazione_to") = ""
		Session(Prefix & "att_data_scadenza_from") = ""
		Session(Prefix & "att_data_scadenza_to") = ""
		Session(Prefix & "att_verso") = ""
		Session(Prefix & "att_full_text") = ""
	elseif request("cerca")<>"" then
		Session(Prefix & "att_stato") = request("search_stato")
		Session(Prefix & "att_tipo") = request("search_tipo")
		Session(Prefix & "att_pratica") = request("search_pratica")
		Session(Prefix & "att_contatto") = request("search_contatto")
		Session(Prefix & "att_oggetto") = request("search_oggetto")
		Session(Prefix & "att_data_creazione_from") = request("search_data_creazione_from")
		Session(Prefix & "att_data_creazione_to") = request("search_data_creazione_to")
		Session(Prefix & "att_data_scadenza_from") = request("search_data_scadenza_from")
		Session(Prefix & "att_data_scadenza_to") = request("search_data_scadenza_to")
		Session(Prefix & "att_verso") = request("search_verso")
		Session(Prefix & "att_full_text") = request("search_full_text")
	end if
end if

'imposta criteri base per ricerca
sql = " SELECT * " & IIf(DB_Type(conn) = DB_SQL, ", ISNULL(att_dataS, " & Sql_DATE(conn, DateAdd("y", 10, Date())) & ") AS OrdScadenza ", "") & _
	  " FROM tb_attivita INNER JOIN tb_admin ON tb_attivita.att_mittente_id=tb_admin.id_admin WHERE " & _
	  " ( " & AL_query(conn, AL_ATTIVITA) & " OR ( att_mittente_id = " & Session("ID_ADMIN") & " )) " & _
	  " AND NOT " & SQL_IsTrue(conn, "att_sistema") & _
	  " AND (NOT " & SQL_IsTrue(conn, "att_inSospeso") & " OR (" & SQL_IsTrue(conn, "att_inSospeso") & " AND att_mittente_id = "& Session("ID_ADMIN") &")) "
	  
if Session("ATT_PRA_ID")<>"" then
	'filtra su id pratica (mostra solo attivita della pratica)
	sql = sql & " AND att_pratica_id=" & Session("ATT_PRA_ID")

	'recupera nome pratica e nome contatto
	if Session("ATT_PRA_NOME") = "" OR Session("ATT_PRA_CNT_NOME") = "" then
		sqlp = "SELECT pra_nome, isSocieta, NomeElencoIndirizzi, CognomeElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi " & _
			   " FROM tb_pratiche INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id = tb_Indirizzario.IDElencoIndirizzi " & _
			   " WHERE pra_id=" & Session("ATT_PRA_ID")
		rs.open sqlp, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		Session("ATT_PRA_NOME") = rs("pra_nome")
		Session("ATT_PRA_CNT_NOME") = ContactName(rs)
		rs.close
	end if
	
elseif Session("ATT_DOC_ID")<>"" then
	'filtra su id documento: mostra solo attivita' con allegato il documento
	sql = sql & " AND att_id IN (SELECT all_attivita_id FROM tb_allegati WHERE all_documento_id=" & Session("ATT_DOC_ID") & ") "
	
	'recupera nome documento
	sqlp = "SELECT doc_nome FROM tb_documenti WHERE doc_id=" & Session("ATT_DOC_ID")
	rs.open sqlp, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	Session("ATT_DOC_NOME") = rs("doc_nome")
	rs.close
end if
	
'filtra per stato delle attivita'
if Session(Prefix & "att_stato")<>"" then
	sql = sql & " AND ( "
	if instr(1, Session(Prefix & "att_stato"), "A", vbTextCompare)>0  then		'aperte
		sql = sql & " (NOT " & SQL_IsTrue(conn, "att_conclusa") & ") OR "
	end if
	if instr(1, Session(Prefix & "att_stato"), "C", vbTextCompare)>0  then		'chiuse
		sql = sql & " " & SQL_IsTrue(conn, "att_conclusa") & " OR "
	end if
	sql = left(sql, len(sql)-3) & " )"
end if

'filtra per tipo attivita'
if Session(Prefix & "att_tipo")<>"" then
	sql = sql & " AND ( "
	if instr(1, Session(Prefix & "att_tipo"), "A", vbTextCompare)>0  then		'attivita'
		sql = sql & " (att_pratica_id <> 0 AND NOT("& SQL_IsNULL(conn, "att_pratica_id") &")) OR "
	end if
	if instr(1, Session(Prefix & "att_tipo"), "S", vbTextCompare)>0  then		'sticker
		sql = sql & " (att_pratica_id = 0 OR "& SQL_IsNULL(conn, "att_pratica_id") &") OR "
	end if
	if instr(1, Session(Prefix & "att_tipo"), "B", vbTextCompare)>0  then		'bozze
		sql = sql & " ("& SQL_IsTrue(conn, "att_inSospeso") & " AND att_mittente_id=" & Session("ID_ADMIN") & ") OR "
	end if
	sql = left(sql, len(sql)-3) & " )"
end if

'filtra per pratica e per contatto 
if Session(Prefix & "att_pratica")<>"" OR Session(Prefix & "att_contatto")<>"" then
	sql = sql & " AND att_pratica_id IN (SELECT pra_id FROM tb_pratiche "
	if Session(Prefix & "att_contatto")<>"" then
		sql = sql & " INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id=tb_indirizzario.idElencoIndirizzi " & _
					" WHERE ( " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session(Prefix & "att_contatto")) & ") "
		if Session(Prefix & "att_pratica")<>"" then
			sql = sql & " AND "
		end if
	else
		sql = sql & " WHERE "
	end if
	if Session(Prefix & "att_pratica")<>"" then
	    sql = sql & SQL_FullTextSearch(Session(Prefix & "att_pratica"), "pra_nome")
	end if
	sql = sql & " )"
end if

'filtra per oggetto
if Session(Prefix & "att_oggetto")<>"" then
	    sql = sql & " AND " & SQL_FullTextSearch(Session(Prefix & "att_oggetto"), "att_oggetto")
end if

'filtra per data di creazione
if isDate(Session(Prefix & "att_data_creazione_from")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "att_dataCrea", adCompareGreaterThan, Session(Prefix & "att_data_creazione_from")) & " "
end if
if isDate(Session(Prefix & "att_data_creazione_to")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "att_dataCrea", adCompareLessThan, Session(Prefix & "att_data_creazione_to")) & " "
end if

'filtra per data di creazione
if isDate(Session(Prefix & "att_data_scadenza_from")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "att_dataS", adCompareGreaterThan, Session(Prefix & "att_data_scadenza_from")) & " "
end if
if isDate(Session(Prefix & "att_data_scadenza_to")) then
	sql = sql & " AND " & SQL_CompareDateTime(conn, "att_dataS", adCompareLessThan, Session(Prefix & "att_data_scadenza_to")) & " "
end if

'filtra per tipo mittenza
if Session(Prefix & "att_verso")<>"" then
	sql = sql & " AND ( "
	if instr(1, Session(Prefix & "att_verso"), "I", vbTextCompare)>0  then		'inviate
		sql = sql & " ( att_mittente_id=" & Session("ID_ADMIN") & ") OR "
	end if
	if instr(1, Session(Prefix & "att_verso"), "R", vbTextCompare)>0  then		'ricevute
		sql = sql & " ( att_mittente_id<>" & Session("ID_ADMIN") & ") OR "
	end if
	sql = left(sql, len(sql)-3) & " )"
end if

'ricerca full-text
if Session(Prefix & "att_full_text")<>"" then
	    sql = sql & " AND " & SQL_FullTextSearch(Session(Prefix & "att_full_text"), "att_oggetto;att_testo;att_note")
end if

'imposta ordinamento
if DB_Type(conn) = DB_SQL then
	sql = sql & " ORDER BY att_conclusa, OrdScadenza, att_priorita DESC, att_dataCrea DESC, att_id DESC"
else
	sql = sql & " ORDER BY att_conclusa DESC, IIF(att_dataS IS NULL, #2079-01-01#, att_dataS), att_priorita, att_dataCrea DESC, att_id DESC "
end if

Session(Prefix & "SQL_ATTIVITA") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
				<form action="Attivita.asp" method="post" id="ricerca" name="ricerca">
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTE" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "att_stato") %>>STATO ATTIVIT&Aacute;</th></tr>
								<tr>
									<td class="content_b">
										<input type="Checkbox" name="search_stato" class="checkbox" value="A" <%= IIF(instr(1, Session(Prefix & "att_stato"), "A", vbTextCompare)>0, " checked", "") %>>
										in corso
									</td>
								</tr>
								<tr>
									<td class="content">
										<input type="Checkbox" name="search_stato" class="checkbox" value="C" <%= IIF(instr(1, Session(Prefix & "att_stato"), "C", vbTextCompare)>0, " checked", "") %>>
										concluse
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "att_tipo") %>>TIPO ATTIVIT&Aacute;</th></tr>
								<tr>
									<td class="content attivita">
										<input type="Checkbox" name="search_tipo" class="checkbox" value="A" <%= IIF(instr(1, Session(Prefix & "att_tipo"), "A", vbTextCompare)>0, " checked", "") %>>
										attivit&agrave;
									</td>
								</tr>
								<tr>
									<td class="content sticker">
										<input type="Checkbox" name="search_tipo" class="checkbox" value="S" <%= IIF(instr(1, Session(Prefix & "att_tipo"), "S", vbTextCompare)>0, " checked", "") %>>
										sticker
									</td>
								</tr>
								<tr>
									<td class="content bozza">
										<input type="Checkbox" name="search_tipo" class="checkbox" value="B" <%= IIF(instr(1, Session(Prefix & "att_tipo"), "B", vbTextCompare)>0, " checked", "") %>>
										bozza
									</td>
								</tr>
								<tr><th<%= Search_Bg(Prefix & "att_oggetto") %>>OGGETTO</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_oggetto" value="<%= replace(Session(Prefix & "att_oggetto"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<% If Session("ATT_PRA_ID") = "" then %>
									<tr><th<%= Search_Bg(Prefix & "att_pratica") %>>PRATICA</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_pratica" value="<%= replace(Session(Prefix & "att_pratica"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
									<tr><th<%= Search_Bg(Prefix & "att_contatto") %>>CONTATTO</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_contatto" value="<%= replace(Session(Prefix & "att_contatto"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
								<% End If %>
								<tr><th<%= Search_Bg(Prefix & "att_data_creazione_from;" & Prefix & "att_data_creazione_to") %>>DATA CREAZIONE</td></tr>
								<tr><td class="label">a partire dal:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_creazione_from", Session(Prefix & "att_data_creazione_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td class="label">fino al:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_creazione_to", Session(Prefix & "att_data_creazione_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><th<%= Search_Bg(Prefix & "att_data_scadenza_from;" & Prefix & "att_data_scadenza_to") %>>DATA SCADENZA</th></tr>
								<tr><td class="label">a partire dal:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_scadenza_from", Session(Prefix & "att_data_scadenza_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td class="label">fino al:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_scadenza_to", Session(Prefix & "att_data_scadenza_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><th<%= Search_Bg(Prefix & "att_verso") %>>TIPO MITTENTE</th></tr>
								<tr>
									<td class="content">
										<input type="Checkbox" name="search_verso" class="checkbox" value="I" <%= IIF(instr(1, Session(Prefix & "att_verso"), "I", vbTextCompare)>0, " checked", "") %>>
										inviate
									</td>
								</tr>
								<tr>
									<td class="content">
										<input type="Checkbox" name="search_verso" class="checkbox" value="R" <%= IIF(instr(1, Session(Prefix & "att_verso"), "R", vbTextCompare)>0, " checked", "") %>>
										ricevute
									</td>
								</tr>
								<tr><th<%= Search_Bg(Prefix & "att_full_text") %>>FULL-TEXT</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_full_text" value="<%= replace(session(Prefix & "att_full_text"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTE" style="width: 49%;">
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
						<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption class="border">Strumenti</caption>
						<tr>
							<td class="content_right">
								<a style="width:100%; text-align:center; line-height:12px;" class="button"
									title="Apre la palette di export dei dati" 
									onclick="OpenAutoPositionedScrollWindow('AttivitaExport.asp', 'export', 640, 480, true);" href="javascript:void(0);">
									EXPORT DATI
								</a>
							</td>
						</tr>
						</table>
						<td>
					</tr>
				</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
				
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						<% if Session("ATT_PRA_ID")<>"" then 
							'scorre per pratica
							if Session(Session("PRA_PREFIX") & "SQL_PRATICHE")<>"" then 
								'verifica se e' disponibile la query per l'elenco delle pratiche%>
								<table border="0" cellspacing="0" cellpadding="0" align="right">
									<tr>
										<td style="font-size: 1px; padding-right:1px;" nowrap>
											<a class="button" href="?ID=<%= Session("ATT_PRA_ID") %>&goto=PREVIOUS" title="pratica precedente">
												&lt;&lt; PRECEDENTE
											</a>
											&nbsp;
											<a class="button" href="?ID=<%= Session("ATT_PRA_ID") %>&goto=NEXT" title="pratica successiva">
												SUCCESSIVA &gt;&gt;
											</a>
										</td>
									</tr>
								</table>
							<% end if %>
							Elenco attivit&agrave; della pratica &quot;<%= Session("ATT_PRA_NOME") %>&quot;
							<span style="white-space:nowrap;">per il cliente &quot;<%= Session("ATT_PRA_CNT_NOME") %>&quot;</span>
						<% elseif Session("ATT_DOC_ID")<>"" then 
							'scorre per documento
							if Session(Session("DOC_PREFIX") & "SQL_DOCUMENTI")<>"" then 
								'verifica se e' disponibile la query per l'elenco dei documenti%>
								<table border="0" cellspacing="0" cellpadding="0" align="right">
									<tr>
										<td style="font-size: 1px; padding-right:1px;" nowrap>
											<a class="button" href="?ID=<%= Session("ATT_DOC_ID") %>&goto=PREVIOUS" title="documento precedente">
												&lt;&lt; PRECEDENTE
											</a>
											&nbsp;
											<a class="button" href="?ID=<%= Session("ATT_DOC_ID") %>&goto=NEXT" title="documento successivo">
												SUCCESSIVO &gt;&gt;
											</a>
										</td>
									</tr>
								</table>
							<% end if %>
							Elenco attivit&agrave; collegato al documento &quot;<%= Session("ATT_DOC_NOME") %>&quot;
						<% else %>	
							Elenco attivit&agrave;
						<% end if %>
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovate n&ordm; <%= Pager.recordcount %> attivit&agrave; in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo
							'recupera numero documenti eventualmente allegati
							sql = " SELECT COUNT(*) FROM tb_documenti INNER JOIN tb_allegati ON tb_documenti.doc_id = tb_allegati.all_documento_id " & _
	  							  " WHERE "& AL_query(conn, AL_DOCUMENTI) & " AND all_attivita_id=" & rs("att_ID")
							N_DOCS = cInteger(GetValueList(conn, rsp, sql)) %>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<% if rs("att_pratica_id") <> 0 then %>
												<td rowspan="6" align="center" valign="top" width="18">
													<%	if rs("att_conclusa") then %>
														<img src="../grafica/AttConclusa.gif" border="0" alt="Attivita' conclusa" title="Attivita' conclusa">
													<%  else %>
													<img src="../grafica/AttAperta.gif" border="0" alt="Attivita' in corso" title="Attivita' in corso">
												<% 	end if %>
											<% Elseif rs("att_inSospeso") then %>
												<td rowspan="4" align="center" valign="top" width="18">
													<img src="../grafica/notes.gif" border="0" alt="Bozza" title="Bozza">
											<% else %>
												<td rowspan="4" align="center" valign="top" width="18">
													<img src="../grafica/notes.gif" border="0" alt="Nota interna" title="Nota interna">
											<% End If
											if rs("att_priorita") then%>
												<img src="../grafica/AttPriorita.gif" border="0" alt="Priorit&agrave; alta" title="Priorit&agrave; alta">
											<%end if
											if N_DOCS > 0 then %>
												<a href="Documenti.asp?ATT_ID=<%= rs("att_id") %>" title="Elenco documenti allegati">
													<img src="..\grafica\AttAttach.gif" border="0" alt="Attivita' con allegati" title="Documenti allegati">
												</a>
											<% End If%>
											</td>
											<% 	If rs("att_conclusa") then						'conclusa
													ClassHeader = ""
											   	elseif rs("att_inSospeso") then					'bozza
													ClassHeader = " bozza"
												elseif cInteger(rs("att_pratica_id")) = 0 then	'stickers
													ClassHeader = " sticker"
												else
													ClassHeader = " attivita"
											   	end if 
											%>
											<td class="header<%= ClassHeader %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<% If N_DOCS > 0 then %>
																<a class="button" href="Documenti.asp?ATT_ID=<%= rs("att_id") %>" title="Elenco documenti allegati">
																	ALLEGATI
																</a>
																&nbsp;
															<% End If
															' Mostra i pulsanti di MODIFICA/VISUALIZZA e di CANCELLA in base al tipo di attivita
															if Session("COM_ADMIN") <> "" OR rs("att_inSospeso") then %>
																<a class="button" href="AttivitaMod.asp?ID=<%= rs("att_id") %>" title="modifica attivit&agrave;">
																	MODIFICA
																</a>
																&nbsp;
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ATTIVITA','<%= rs("att_id") %>');">
																	CANCELLA
																</a>
															<% 	Else %>
																<a class="button" href="AttivitaMod.asp?ID=<%= rs("att_id") %>" title="visualizza attivit&agrave;">
																	VISUALIZZA
																</a>
																&nbsp;
																<a class="button_disabled" title="Impossibile cancellare: non si hanno i permessi sull'attivit&agrave;.!">
																	CANCELLA
																</a>
															<% 	End If 'fine attivita di sistema %>
														</td>
													</tr>
												</table>
												<% If rs("att_conclusa") then %>
													<span style="font-weight:normal;">
														<%= rs("att_oggetto") %>
													</span>
												<% Else  %>
													<%= rs("att_oggetto") %>
												<% End If %>
											</td>
										</tr>
										<% if Session("ATT_PRA_ID")="" AND rs("att_pratica_id")>0 then
											sql = "SELECT pra_id, pra_nome, isSocieta, NomeElencoIndirizzi, CognomeElencoIndirizzi, " & _
												  " NomeOrganizzazioneElencoIndirizzi, IDElencoIndirizzi " & _
												  " FROM tb_pratiche INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id = tb_Indirizzario.IDElencoIndirizzi " & _
												  " WHERE pra_id=" & rs("att_pratica_id")
											rsp.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
											if not rsp.eof then%> 
												<tr>
													<td class="label">contatto:</td>
													<td class="content" colspan="3">
														<% ContactLinkedName(rsp) %>
													</td>
												</tr>
												<tr>
													<td class="label">pratica:</td>
													<td class="content" colspan="3">
														<% PraticaLinkedName(rsp) %>
													</td>
												</tr>
											<% else %>
												<tr>
													<td class="content_b" colspan="4">
														Errore dati dell'attivit&agrave;. Pratica collegata non trovata.
													</td>
												</tr>
											<% end if
											rsp.close
										end if %>
										<tr>
											<td class="label" style="width:20%;">data creazione:</td>
											<td class="content" style="width:25%;"><%= DateTimeITA(rs("att_dataCrea")) %></td>
											<% If rs("att_conclusa") then %>
												<td class="label">data chiusura:</td>
												<td class="content"><%= DateTimeITA(rs("att_dataChiusa")) %></td>
											<% Elseif isDate(rs("att_dataS")) AND NOT IsNull(rs("att_dataS")) then%>
												<td class="label" style="width:16%;">data scadenza:</td>
												<td class="content">
													<% if DateDiff("d", Date, rs("att_dataS")) < 0 then 
														'scaduta%>
														<strong>
															Scaduta il <%= DateIta(rs("att_dataS")) %>
														</strong>
													<% elseif DateDiff("d", Date, rs("att_dataS"))=0 then 
														'scade oggi%>
															Scade oggi
													<% elseif DateDiff("d", Date, rs("att_dataS"))=1 then
														'scade domani %>
															Scade domani
													<% elseif DateDiff("d", Date, rs("att_dataS"))>1 then 
														'scade fra x giorni%>
															Scade fra n&deg; <%= DateDiff("d", Date, rs("att_dataS")) %> giorni
													<% end if %>
												</td>
											<% Else %>
												<td colspan="2" class="content">&nbsp;</td>
											<% End If %>
										</tr>
										<tr>
											<td class="label">mittente:</td>
											<td class="content" colspan="3"><%= rs("admin_nome") &" "& rs("admin_cognome") %></td>
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
						<tr><td class="noRecords" colspan="2">Nessun record trovato</th></tr>
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
set rsp = nothing
set rs = nothing
set conn = nothing
%>


