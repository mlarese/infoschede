<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<%
dim conn, rs, rsp, sql, sqlp, Pager, var
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	if Session("DOC_PRA_ID")<>"" then
		CALL GotoRecord(conn, rs, Session(Session("PRA_PREFIX") & "SQL_PRATICHE"), "pra_id", "Documenti.asp")
	elseif Session("DOC_ATT_ID")<>"" then
		CALL GotoRecord(conn, rs, Session(Session("ATT_PREFIX") & "SQL_ATTIVITA"), "att_id", "Documenti.asp")
	end if
end if

if request.QueryString("all")<>"" then		'se richiesta arriva da menu: visualizza tutti i documenti
	'da pratiche
	Session("DOC_PRA_ID") = ""
	Session("DOC_PRA_NOME") = ""
	Session("DOC_PRA_CNT_NOME") = ""
	'da attivita
	Session("DOC_ATT_ID") = ""
	Session("DOC_ATT_OGGETTO") = ""
	
	Session("DOC_PREFIX") = ""
	response.redirect "Documenti.asp"
elseif request.Querystring("PRA_ID")<>"" OR _
	(request.querystring("ID")<>"" AND Session("DOC_PRA_ID")<>"") then	
					'se richiesta arriva da pratiche: visualizza i documenti della pratica
	'da pratiche
	Session("DOC_PRA_ID") = IIF(request.Querystring("PRA_ID")<>"", request.Querystring("PRA_ID"), request.querystring("ID"))
	Session("DOC_PRA_NOME") = ""
	Session("DOC_PRA_CNT_NOME") = ""
	'da attivita'
	Session("DOC_ATT_ID") = ""
	Session("DOC_ATT_OGGETTO") = ""
	
	Session("DOC_PREFIX") = "ATT_"
	response.redirect "Documenti.asp"
elseif request.Querystring("ATT_ID")<>"" OR _
	(request.querystring("ID")<>"" AND Session("DOC_ATT_ID")<>"") then	
					'se richiesta arriva da attvita': visualizza tutti i documenti allegati all'attivita'
	'da pratiche
	Session("DOC_PRA_ID") = ""
	Session("DOC_PRA_NOME") = ""
	Session("DOC_PRA_CNT_NOME") = ""
	'da documenti
	Session("DOC_ATT_ID") = IIF(request.Querystring("ATT_ID")<>"", request.Querystring("ATT_ID"), request.querystring("ID"))
	Session("DOC_ATT_OGGETTO") = ""
	
	Session("DOC_PREFIX") = "ATT_"
	response.redirect "Documenti.asp"
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
if Session("DOC_PRA_ID")<>"" then
	'documenti della pratica
	Titolo_sezione = "Pratiche - documenti della pratica"
	HREF = "Pratiche.asp;PraticaMod.asp?ID=" & Session("DOC_PRA_ID") & ";Attivita.asp?pra_id=" & Session("DOC_PRA_ID") & ";"
	Action = "INDIETRO;SCHEDA PRATICA;ATTIVITA' PRATICA;"
elseif Session("DOC_ATT_ID")<>"" then
	'attivita' collegate al documento
	Titolo_sezione = "Attivit&agrave; - documenti allegati"
	HREF = "Attivita.asp;AttivitaMod.asp?ID=" & Session("DOC_ATT_ID")
	Action = "INDIETRO;SCHEDA ATTIVITA'"
else
	Titolo_sezione = "Documenti - elenco"
	HREF = ""
	Action = ""
end if
if Session("DOC_ATT_ID")= "" then
	HREF = HREF & "DocumentoNew.asp"
	Action = Action & "NUOVO DOCUMENTO"
end if
if Session("COM_ADMIN") <> "" AND Session("DOC_PREFIX")="" then
	SSezioniText = "DOCUMENTI;TIPOLOGIE;DESCRITTORI"
	SSezioniLink = "documenti.asp;tipologie.asp;descrittori.asp"
end if
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************


dim Prefix, i, FileList, N_ATTIVITA, N_FILES
Prefix = Session("DOC_PREFIX")

if request.querystring("semplice") <> "" then			'ho scelto la ricerca semplice
	'azzera le variabili per le richieste: "vedi tutti" o per richiesta di ricerca semplice
	Session(Prefix & "ADV_doc_TXT") = ""
	Session(Prefix & "ADV_doc_SQL") = ""
	for each var in Session.Contents
		if left(var, len(Prefix &"adv_doc_")) = Prefix &"adv_doc_" then
			Session(var) = ""
		end if
	next
end if

Session("PageNavigator_VarPrefix") = Prefix
set Pager = new PageNavigator
'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	if request("tutti")<>"" then
		Session(Prefix & "doc_nome") = ""
		Session(Prefix & "doc_file") = ""
		Session(Prefix & "doc_tipologia") = ""
		Session(Prefix & "doc_pratica") = ""
		Session(Prefix & "doc_contatto") = ""
		Session(Prefix & "doc_data_creazione_from") = ""
		Session(Prefix & "doc_data_creazione_to") = ""
		Session(Prefix & "doc_full_text") = ""
	elseif request("cerca")<>"" then
		Session(Prefix & "doc_nome") = request("search_nome")
		Session(Prefix & "doc_file") = request("search_file")
		Session(Prefix & "doc_tipologia") = request("search_tipologia")
		Session(Prefix & "doc_pratica") = request("search_pratica")
		Session(Prefix & "doc_contatto") = request("search_contatto")
		Session(Prefix & "doc_data_creazione_from") = request("search_data_creazione_from")
		Session(Prefix & "doc_data_creazione_to") = request("search_data_creazione_to")
		Session(Prefix & "doc_full_text") = request("search_full_text")
	end if
	
	'azzera le variabili per le richieste: "vedi tutti" o per richiesta di ricerca semplice
	Session(Prefix & "ADV_doc_TXT") = ""
	Session(Prefix & "ADV_doc_SQL") = ""
	for each var in Session.Contents
		if left(var, len(Prefix &"adv_doc_")) = Prefix &"adv_doc_" then
			Session(var) = ""
		end if
	next
end if


if Session(Prefix & "ADV_doc_SQL")<>"" AND Session(Prefix & "ADV_doc_TXT")<>"" then
	'imposta query di ricerca avanzata
	sql = Session(Prefix & "ADV_doc_SQL")
else

 
	'imposta criteri per ricerca semplice
	sql = " SELECT * FROM (tb_documenti INNER JOIN tb_tipologie ON tb_documenti.doc_tipologia_id=tb_tipologie.tipo_id) "& _
		  " INNER JOIN tb_admin ON tb_documenti.doc_creatore_id=tb_admin.id_admin "& _
		  " WHERE (doc_creatore_id="& Session("ID_ADMIN") &" OR "& AL_query(conn, AL_DOCUMENTI) &")"
		  
	if Session("DOC_PRA_ID")<>"" then
		'filtra su id pratica (mostra solo documenti della pratica)
		sql = sql & " AND doc_pratica_id=" & Session("DOC_PRA_ID")
	
		'recupera nome pratica e nome contatto
		if Session("DOC_PRA_NOME") = "" OR Session("DOC_PRA_CNT_NOME") = "" then
			sqlp = "SELECT pra_nome, isSocieta, NomeElencoIndirizzi, CognomeElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi " & _
				   " FROM tb_pratiche INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id = tb_Indirizzario.IDElencoIndirizzi " & _
				   " WHERE pra_id=" & Session("DOC_PRA_ID")
			rs.open sqlp, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			Session("DOC_PRA_NOME") = rs("pra_nome")
			Session("DOC_PRA_CNT_NOME") = ContactName(rs)
			rs.close
		end if
		
	elseif Session("DOC_ATT_ID")<>"" then
		'filtra su id documento: mostra solo documenti allegati alla pratica
		sql = sql & " AND doc_id IN (SELECT all_documento_id FROM tb_allegati WHERE all_attivita_id=" & Session("DOC_ATT_ID") & ") "
		
		'recupera oggetto attivita
		if Session("DOC_ATT_OGGETTO")="" then
			sqlp = "SELECT att_oggetto FROM tb_attivita WHERE att_id=" & Session("DOC_ATT_ID")
			rs.open sqlp, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			Session("DOC_ATT_OGGETTO") = rs("att_oggetto")
			rs.close
		end if
	end if
	
	'filtra per nome documento
	if Session(Prefix & "doc_nome")<>"" then
		sql = sql & " AND " + SQL_FullTextSearch(Session(Prefix & "doc_nome"), "doc_nome")
	end if
	
	'filtra per full-text su nome e note
	if Session(Prefix & "doc_full_text") <> "" then
		sql = sql & " AND " + SQL_FullTextSearch(Session(Prefix & "doc_full_text"), "doc_nome;doc_note")
	end if
	
	'filtra per nome file contenuto
	if Session(Prefix & "doc_file")<>"" then
		sql = sql & " AND doc_id IN (SELECT rel_documento_id FROM rel_documenti_files " & _
					" INNER JOIN tb_files ON rel_documenti_files.rel_files_id = tb_files.f_id " & _
					" WHERE " + SQL_FullTextSearch(Session(Prefix & "doc_file"), "f_original_name") + ")"
	end if
	
	'filtra per pratica e per contatto 
	if Session(Prefix & "doc_pratica")<>"" OR Session(Prefix & "doc_contatto")<>"" then
		sql = sql & " AND doc_pratica_id IN (SELECT pra_id FROM tb_pratiche "
		if Session(Prefix & "doc_contatto")<>"" then
			sql = sql & " INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id=tb_indirizzario.idElencoIndirizzi " & _
						" WHERE ( " + SQL_FullTextSearch_Contatto_Nominativo(conn, Session(Prefix & "doc_contatto")) + " ) "
			if Session(Prefix & "doc_pratica")<>"" then
				sql = sql & " AND "
			end if
		else
			sql = sql & " WHERE "
		end if
		if Session(Prefix & "doc_pratica")<>"" then
			sql = sql & SQL_FullTextSearch(Session(Prefix & "doc_pratica"), "pra_nome")
		end if
		sql = sql & " )"
	end if
	
	'filtra per data di creazione
	if isDate(Session(Prefix & "doc_data_creazione_from")) then
		sql = sql & " AND " & SQL_CompareDateTime(conn, "doc_dataC", adCompareGreaterThan, Session(Prefix & "doc_data_creazione_from")) & " "
	end if
	if isDate(Session(Prefix & "doc_data_creazione_to")) then
		sql = sql & " AND " & SQL_CompareDateTime(conn, "doc_dataC", adCompareLessThan, Session(Prefix & "doc_data_creazione_to")) & " "
	end if
	
	'filtra per tipologia
	if Session(Prefix & "doc_tipologia")<>"" then
		sql = sql & " AND doc_tipologia_id = " & ParseSQL(Session(Prefix & "doc_tipologia"), adChar)
	end if
	
	sql = sql & " ORDER BY doc_nome"
end if

Session(Prefix & "SQL_DOCUMENTI") = sql

rs.Open sql, conn, AdOpenStatic, adLockReadOnly, adCmdText
rs.PageSize = 8
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
				<form action="Documenti.asp" method="post" id="ricerca" name="ricerca">
				<% if not (Session(Prefix & "ADV_doc_SQL")<>"" AND Session(Prefix & "ADV_doc_TXT")<>"") then %>
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg(prefix & "doc_nome") %>>NOME DOCUMENTO</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= replace(session(prefix & "doc_nome"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg(prefix & "doc_file") %>>NOME FILE</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_file" value="<%= replace(session(prefix & "doc_file"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "doc_tipologia") %>>TIPOLOGIA</th></tr>
								<tr>
									<td class="content">
										<% sql = "SELECT * FROM tb_tipologie ORDER BY tipo_nome"
										CALL dropDown(conn, sql, "tipo_id", "tipo_nome", "search_tipologia", Session(Prefix & "doc_tipologia"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
									</td>
								</tr>
								<% If Session("DOC_PRA_ID") = "" AND Session("DOC_ATT_ID")="" then %>
									<tr><th <%= Search_Bg(Prefix & "doc_pratica") %>>PRATICA</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_pratica" value="<%= replace(Session(Prefix & "doc_pratica"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg(Prefix & "doc_contatto") %>>CONTATTO</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_contatto" value="<%= replace(Session(Prefix & "doc_contatto"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
								<% End If %>
								<tr><th <%= Search_Bg(Prefix & "doc_data_creazione_from;" & Prefix & "doc_data_creazione_to") %>>DATA CREAZIONE</td></tr>
								<tr><td class="label">a partire dal:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_creazione_from", Session(Prefix & "doc_data_creazione_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><td class="label">fino al:</td></tr>
								<tr>
									<td class="content">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_creazione_to", Session(Prefix & "doc_data_creazione_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><th <%= Search_Bg(Prefix & "doc_full_text") %>>FULL-TEXT</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_full_text" value="<%= replace(session(Prefix & "doc_full_text"), """", "&quot;") %>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<% else %>
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca avanzata</caption>
								<tr>
									<td class="footer">
										<input type="button" name="cerca" value="CAMBIA RICERCA" class="button" style="width: 59%;" onclick="OpenRicercaAvanzata()">
										<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 39%;">
									</td>
								</tr>
								<tr><th>CRITERI IMPOSTATI</td></tr>
								<%= Session(Prefix & "ADV_doc_TXT") %>
								<tr>
									<td class="footer">
										<input type="button" name="cerca" value="CAMBIA RICERCA" class="button" style="width: 59%;" onclick="OpenRicercaAvanzata()">
										<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTI" style="width: 39%;">
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<% end if%>
					<script language="JavaScript" type="text/javascript">
						function OpenRicercaAvanzata(){
							if (!(this.name))
								this.name = "ElencoContatti_<%= Session.SessionID %>";
							OpenPositionedScrollWindow('DocumentiRicercaAvanzata.asp', 'ricercaavanzata', window.screenLeft - 40, window.screenTop, 410, 360, true)
						}
						
					</script>
					<tr><td style="font-size:4px;">&nbsp;</td></tr>
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption class="border">Strumenti</caption>
								<% if Session(Prefix & "ADV_doc_SQL")<>"" AND Session(Prefix & "ADV_doc_TXT")<>"" then %>
									<tr>
										<td class="content_right">
											<a style="width:100%; text-align:center; line-height:12px;" class="button"
												title="Annulla la ricerca avanzata in corso."
												href="?semplice=si">
												RICERCA SEMPLICE
											</a>
										</td>
									</tr>
								<% else %>
									<tr>
										<td class="content_right">
											<a style="width:100%; text-align:center; line-height:12px;" class="button"
												title="Apre la palette di ricerca avanzata" 
												onclick="OpenRicercaAvanzata()" href="javascript:void(0);">
												RICERCA AVANZATA
											</a>
										</td>
									</tr>
								<% end if %>								
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
						<% if Session("DOC_PRA_ID")<>"" then 
							'scorre per pratica
							if Session(Session("PRA_PREFIX") & "SQL_PRATICHE")<>"" then 
								'verifica se e' disponibile la query per l'elenco delle pratiche%>
								<table border="0" cellspacing="0" cellpadding="0" align="right">
									<tr>
										<td style="font-size: 1px; padding-right:1px;" nowrap>
											<a class="button" href="?ID=<%= Session("DOC_PRA_ID") %>&goto=PREVIOUS" title="pratica precedente">
												&lt;&lt; PRECEDENTE
											</a>
											&nbsp;
											<a class="button" href="?ID=<%= Session("DOC_PRA_ID") %>&goto=NEXT" title="pratica successiva">
												SUCCESSIVA &gt;&gt;
											</a>
										</td>
									</tr>
								</table>
							<% end if %>
							Elenco documenti della pratica &quot;<%= Session("DOC_PRA_NOME") %>&quot;
							<span style="white-space:nowrap;">per il cliente &quot;<%= Session("DOC_PRA_CNT_NOME") %>&quot;</span>
						<% elseif Session("DOC_ATT_ID")<>"" then 
							'scorre per documento
							if Session(Session("ATT_PREFIX") & "SQL_ATTIVITA")<>"" then 
								'verifica se e' disponibile la query per l'elenco dei documenti%>
								<table border="0" cellspacing="0" cellpadding="0" align="right">
									<tr>
										<td style="font-size: 1px; padding-right:1px;" nowrap>
											<a class="button" href="?ID=<%= Session("DOC_ATT_ID") %>&goto=PREVIOUS" title="attivit&agrave; precedente">
												&lt;&lt; PRECEDENTE
											</a>
											&nbsp;
											<a class="button" href="?ID=<%= Session("DOC_ATT_ID") %>&goto=NEXT" title="attivit&agrave; successiva">
												SUCCESSIVA &gt;&gt;
											</a>
										</td>
									</tr>
								</table>
							<% end if %>
							Elenco documenti allegati all'attivit&agrave; &quot;<%= Session("DOC_ATT_OGGETTO") %>&quot;
						<% else %>	
							Elenco documenti
						<% end if %>
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= rs.recordcount %> documenti in n&ordm; <%= rs.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo	
							sql = "SELECT COUNT(*) FROM tb_allegati WHERE all_documento_id=" & rs("doc_id")
							N_ATTIVITA = cInteger(GetValueList(conn, rsp, sql))%>
							<tr>
								<td class="body" colspan="2">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<% If N_ATTIVITA > 0 then %>
																<a class="button" href="Attivita.asp?DOC_ID=<%= rs("doc_id") %>" title="Attivita' che hanno il documento in allegato">
																	ATTIVIT&Aacute; COLLEGATE
																</a>
																&nbsp;
															<% 	End If
															If Session("COM_ADMIN") <> "" OR Session("COM_POWER") <> "" OR _
															   Session("ID_ADMIN") = rs("doc_creatore_id") then 
																sql = "SELECT COUNT(*) FROM rel_documenti_files WHERE rel_documento_id=" & rs("doc_id")
																N_FILES = cInteger(GetValueList(conn, rsp, sql))%>
																<a class="button" href="DocumentoMod.asp?ID=<%= rs("doc_id") %>" title="modifica dati del documento">
																	MODIFICA
																</a>
																&nbsp;
																<%if Session("COM_ADMIN") <> "" OR (N_ATTIVITA=0 AND N_FILES=0) then %>
																	<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('DOCUMENTI','<%= rs("doc_id") %>');">
																		CANCELLA
																	</a>
																<% Else%>
																	<a class="button_disabled" title="impossibile cancellare il documento: attivit&agrave; o file associati!">
																		CANCELLA
																	</a>
																<% End If
															Else  %>
																<a class="button" href="DocumentoMod.asp?ID=<%= rs("doc_id") %>" title="apri scheda del documento in visualizzazione">
																	VISUALIZZA
																</a>
																&nbsp;
																<a class="button_disabled" title="impossibile cancellare il documento per mancanza dei permessi su di esso.">
																	CANCELLA
																</a>
															<% 	End If %>
														</td>
													</tr>
												</table>
												<%= rs("tipo_nome") %> - <%= rs("doc_nome") %>
											</td>
										</tr>
										<% if Session("DOC_PRA_ID")="" AND rs("doc_pratica_id")>0 then
											sql = "SELECT pra_id, pra_nome, isSocieta, NomeElencoIndirizzi, CognomeElencoIndirizzi, " & _
												  " NomeOrganizzazioneElencoIndirizzi, IDElencoIndirizzi " & _
												  " FROM tb_pratiche INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id = tb_Indirizzario.IDElencoIndirizzi " & _
												  " WHERE pra_id=" & rs("doc_pratica_id")
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
											<td class="label">creato da:</td>
											<td class="content" style="width:40%"><%= rs("admin_nome") &" "& rs("admin_cognome") %></td>
											<td class="label" style="width:25%">data creazione:</td>
											<td class="content"><%= DateTimeITA(rs("doc_dataC")) %></td>
										</tr>
										<% if rs("doc_mod_utente") <> 0 then 
											sql = "SELECT admin_cognome, admin_nome FROM tb_admin WHERE id_admin=" & rs("doc_mod_utente")
											rsp.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
											<tr>
												<td class="label" nowrap>modificato da:</td>
												<td class="content"><%= rsp("admin_nome") &" "& rsp("admin_cognome") %></td>
												<td class="label">data ultima modifica:</td>
												<td class="content"><%= DateTimeITA(rs("doc_dataC")) %></td>
											</tr>
											<%rsp.close
										end if
										%>
										<tr>
											<td class="label">files:</td>
											<td colspan="3">
												<% CALL ElencoFileAssociati(conn, rsp, rs("doc_id")) %>
											</td>
										</tr>
									<%	'VISUALIZZO DESCRITTORI PRINCIPALI
										sql = " SELECT * FROM tb_descrittori d " & _
											  " INNER JOIN rel_documenti_descrittori r ON d.descr_id = r.rdd_descrittore_id " & _
											  " WHERE "& SQL_isTrue(conn, "descr_principale") &" AND rdd_documento_id="& rs("doc_id")
										rsp.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
										if not rsp.eof then
											while not rsp.eof %>
										<tr>
											<td class="label"><%= rsp("descr_nome") %>:</td>
											<td colspan="3" class="content">
												<%= DesFormat(rsp("descr_tipo"), rsp("rdd_valore"), "target=""_blank""", "", "&euro;") %>
											</td>
										</tr>
											<%	rsp.movenext
											wend
										end if
										rsp.close
									%>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td class="footer" style="border-top:0px; text-align:left;" colspan="2">
								<% 	CALL Pager.Render_GroupNavigator(10, rs.PageCount, "", "button", "button_disabled")%>
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
set rs = nothing
set conn = nothing
%>
