<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Anagrafica contatti - elenco"
'Indirizzo pagina per link su sezione 
		HREF = "ContattiNew.asp"
'Azione sul link: {BACK | NEW}
	Action = "NUOVO CONTATTO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<!--#INCLUDE FILE ="../library/ExportTools.asp" -->
<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>

<%
dim conn, rs, rsr, rsv, rst, sql, rubriche_visibili, lettera, Pager, nominativo, i, var, sql_export, checked

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
set rst = Server.CreateObject("ADODB.RecordSet")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" OR request("ALL")<>"" then
	Pager.Reset()
	
	CALL SearchSession_Reset("search_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("search_")
	end if
	
	'azzera le variabili per le richieste: "vedi tutti" o per richiesta di ricerca semplice
	Session("ADV_search_TXT") = ""
	Session("ADV_search_SQL") = ""
end if

if Session("ADV_search_SQL")<>"" AND Session("ADV_search_TXT")<>"" then
	'imposta query di ricerca avanzata
	sql = Session("ADV_search_SQL")
else
	'imposta criteri per ricerca semplice
	sql = " SELECT IDElencoIndirizzi, Lingua, lingua_nome_IT, LockedByApplication, ApplicationsLocker, SyncroApplication, " +_
		  " isSocieta, CognomeElencoIndirizzi, NomeElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, TitoloElencoIndirizzi, " + _
		  " indirizzoElencoIndirizzi, capElencoIndirizzi, cittaElencoIndirizzi, statoProvElencoIndirizzi, countryElencoIndirizzi, localitaelencoindirizzi, DataIscrizione " + _
		  " FROM tb_indirizzario INNER JOIN tb_cnt_lingue ON tb_indirizzario.lingua=tb_cnt_lingue.lingua_codice WHERE " & _
          " (CntRel = 0 OR "  & SQL_IsNull(conn, "CntRel") & ") AND " & _
		  " IDElencoIndirizzi IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica IN ("
	
	'filtra sulle rubriche
	if Session("search_rubriche")<>"" then
		sql = sql & Session("search_rubriche")
	elseif rubriche_visibili<>"" then
		sql = sql & rubriche_visibili
	else
		sql = sql & "0"
	end if
	sql = sql & ")) "
	
	if cIntero(Session("search_categoria"))>0 then
		sql = sql & " AND cnt_categoria_id=" & Session("search_categoria")
	end if
	
	if cIntero(Session("search_campagna"))>0 then
		sql = sql & " AND IDElencoIndirizzi IN (SELECT rcc_cnt_id FROM rel_cnt_campagne WHERE rcc_campagna_id = " & Session("search_campagna") & ")"
	end if
	if cIntero(Session("search_campagna_conclusa"))=1 then
		sql = sql & " AND IDElencoIndirizzi IN (SELECT rcc_cnt_id FROM rel_cnt_campagne WHERE rcc_data_conclusione IS NULL )"
	elseif cIntero(Session("search_campagna_conclusa"))=2 then
		sql = sql & " AND IDElencoIndirizzi IN (SELECT rcc_cnt_id FROM rel_cnt_campagne WHERE rcc_data_conclusione IS NOT NULL )"
	end if
	
	if cString(Session("search_trattativa"))<>"" then
		sql = sql & " AND IDElencoIndirizzi IN (SELECT ima_contatto_id FROM tb_indirizzario_macchine WHERE IsNull(ima_stato_trattativa, 0) = 1 and IsNull(ima_esito_trattativa,0)="&Session("search_trattativa")&")"
	end if
	
	if Session("search_iniziali")<>"" then
		sql = sql & " AND " & SQL_Ucase(conn) & "(LEFT(ModoRegistra, 1)) IN (" & Session("search_iniziali") & ")"
	end if
	
	if Session("search_denominazione")<>"" then
	    sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("search_denominazione"))
	end if
	
	if Session("search_indirizzo")<>"" then
	    sql = sql & " AND " & SQL_FullTextSearch_Contatto_Indirizzo(conn, Session("search_indirizzo"))
	end if
	
	if Session("search_citta")<>"" then
	    sql = sql & " AND " & SQL_FullTextSearch(Session("search_citta"), "CittaElencoIndirizzi")
	end if
	
	if Session("search_provincia")<>"" then
	    sql = sql & " AND " & SQL_FullTextSearch(Session("search_provincia"), "statoProvElencoIndirizzi")
	end if
	
	if Session("search_email")<>"" then
		sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri " &_
					" WHERE " + SQL_FullTextSearch(session("search_email"), "ValoreNumero") + " AND (id_TipoNumero<=6))"
	end if
						
	'sql = sql & " ORDER BY ModoRegistra"
end if

if inStr(sql, "ORDER BY") = 0 then
	if DB_Type(conn) = DB_SQL then
		sql = sql & " ORDER BY (CASE isSocieta WHEN 1 THEN NomeOrganizzazioneElencoIndirizzi ELSE CognomeElencoIndirizzi+NomeElencoIndirizzi END), ModoRegistra"
	else
		sql = sql & " ORDER BY (IIF(isSocieta, NomeOrganizzazioneElencoIndirizzi, CognomeElencoIndirizzi+NomeElencoIndirizzi)), ModoRegistra"
	end if
end if
Session("SQL_ELENCO") = sql

'calcola query per utilizzo come sottoquery.
sql_export =  Session("SQL_ELENCO")
sql_export = "SELECT IDElencoIndirizzi " & Right(sql_export,Len(sql_export)-InStr(sql_export,"FROM")+1)
sql_export = Left(sql_export,Len(sql_export)-(Len(sql_export)-InStr(sql_export,"ORDER BY") +1))

'response.write sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)

%>
<!-- <%= sql_export %>-->
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<form action="contatti.asp" method="post" id="ricerca" name="ricerca">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<% if not (Session("ADV_search_SQL")<>"" AND Session("ADV_search_TXT")<>"") then %>
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
									<tr><th<%= Search_Bg("search_rubriche") %>>RUBRICHE</th></tr>
									<tr>
										<td class="content">
											<script language="JavaScript" type="text/javascript">
												function ShowName(obj){
													var value = obj.options(obj.selectedIndex).text;
													if (value.length>33)
														alert(obj.options(obj.selectedIndex).text);
												}
											</script>
											<% sql = " SELECT " & _ 
													 IIF(DB_Type(conn) = DB_SQL, "(' ' + CAST(id_rubrica AS nvarchar(8)) + ' ') ", "(' ' & id_rubrica & ' ')") & " AS ID, " &_
													 " nome_rubrica FROM tb_rubriche " &_
													 " WHERE id_rubrica IN (" & rubriche_visibili & ") " &_
													 " ORDER BY rubrica_esterna, nome_rubrica"
											CALL dropDown(conn, sql, "ID", "nome_rubrica", "search_rubriche", Session("search_rubriche"), true, _
														  " multiple size=""20"" style=""width:100%;"" onDblClick=""ricerca.submit();""", LINGUA_ITALIANO)%>
											<div class="note">
												Ctrl + Click per selezioni multiple.<br>
												<!--Doppio click per visualizzare il nome.-->
											</div>
										</td>
									</tr>
									<% if Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE") then %>
										<tr><th <%= Search_Bg("search_categoria") %>>CATEGORIA</th></tr>
										<tr>
											<td class="content">
												<%CALL dropDown(conn, CatContatti.QueryElenco(true, ""), "icat_id", "NAME", "search_categoria", Session("search_categoria"), false, "style=""width:100%""", LINGUA_ITALIANO)%>
											</td>
										</tr>
									<% end if %>
									<% if Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") then %>
										<tr><th <%= Search_Bg("search_campagna;search_campagna_conclusa;") %>>CAMPAGNA MARKETING</th></tr>
										<tr>
											<td class="content">
												<%CALL dropDown(conn, "SELECT * FROM tb_indirizzario_campagne ORDER BY inc_nome", "inc_id", "inc_nome", "search_campagna", Session("search_campagna"), false, "style=""width:100%""", LINGUA_ITALIANO)%>
											</td>
										</tr>
										<tr>
											<td class="content" style="padding-left:0px; padding-right:0px;">
												<input type="checkbox" class="noborder" name="search_campagna_conclusa" value="2" <%=chk(Session("search_campagna_conclusa")="2")%>>Solo contatti conclusi
											</td>
										</tr>
										<tr>
											<td class="content" style="padding-left:0px; padding-right:0px;">
												<input type="checkbox" class="noborder" name="search_campagna_conclusa" value="1" <%=chk(Session("search_campagna_conclusa")="1")%>>Solo contatti ancora da concludere
											</td>
										</tr>
									<% end if %>
									<% if Session("ATTIVA_PARCO_MACCHINE") then %>
										<tr><th <%= Search_Bg("search_trattativa") %>>TRATTATIVA</th></tr>
										<tr>
											<td class="content">
												<% checked = chk(cIntero(Session("search_trattativa")) = 0 AND cString(Session("search_trattativa"))<>"") %>
												<input type="radio" name="search_trattativa" value="0" class="noborder<%=IIF(checked<>"", " selected", "")%>" onclick="ClickTrattativa(this)" <%=checked%>>
												In corso
											</td>
										</tr>
										<tr>
											<td class="content">
												<% checked = chk(cIntero(Session("search_trattativa")) = 1) %>
												<input type="radio" name="search_trattativa" value="1" class="noborder<%=IIF(checked<>"", " selected", "")%>" onclick="ClickTrattativa(this)" <%=checked%>>
												Vinta
											</td>
										</tr>
										<tr>
											<td class="content">
												<% checked = chk(cIntero(Session("search_trattativa")) = 2) %>
												<input type="radio" name="search_trattativa" value="2" class="noborder<%=IIF(checked<>"", " selected", "")%>" onclick="ClickTrattativa(this)" <%=checked%>>
												Persa
											</td>
										</tr>
										<script language="JavaScript" type="text/javascript">
											function ClickTrattativa(clicked){
												if (clicked.className.indexOf(' selected') > 0){
													clicked.checked = false;
													clicked.className = clicked.className.replace(' selected', '');
												}
												else{
													clicked.checked = true; // abilito il radiobutton cliccato
													clicked.className = clicked.className + ' selected';
												}
											}
										</script>
									<% end if 
									%>
									<!--
									<tr><th<%= Search_Bg("search_iniziali") %>>INIZIALI</th></tr>
									<tr>
										<td>
											<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
												<tr>
													<%for i=asc("A") to asc("Z")%>
					    								<TD class="content">
															<INPUT class="checkbox" type="checkbox" name="search_iniziali" value="'<%=chr(i)%>'" <%if instr(1, Session("search_iniziali"), chr(i), vbTextCompare)>0 then %> checked <% end if %>>
															<%=chr(i)%>
														</TD>
					    								<%if i mod 4 = 0 then%>
															</tr>
															<tr>
														<%end if
													next %>
													<td class="content" colspan="2">&nbsp;</td>
												</tr>
											</table>
										</td>
									</tr>
									-->
									<tr><th <%= Search_Bg("search_denominazione") %>>NOME / DENOMINAZIONE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_denominazione" value="<%= replace(session("search_denominazione"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("search_indirizzo") %>>INDIRIZZO</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_indirizzo" value="<%= replace(session("search_indirizzo"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("search_citta") %>>CITT&Agrave;</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_citta" value="<%= replace(session("search_citta"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("search_provincia") %>>PROVINCIA</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_provincia" value="<%= replace(session("search_provincia"), """", "&quot;") %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("search_email") %>>E-MAIL</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_email" value="<%= replace(session("search_email"), """", "&quot;") %>" style="width:100%;">
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
									<tr><th>CRITERI IMPOSTATI</th></tr>
									<!-- <%= "sql ricerca: " & Session("ADV_search_SQL") %> -->
									<%= Session("ADV_search_TXT") %>
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
							OpenPositionedScrollWindow('ContattiRicercaAvanzata.asp', 'ricercaavanzata', window.screenLeft - 40, window.screenTop, 410, 750, true)
						}
						
					</script>
					<tr><td style="font-size:4px;">&nbsp;</td></tr>
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption class="border">Strumenti</caption>
								<% if Session("ADV_search_SQL")<>"" AND Session("ADV_search_TXT")<>"" then %>
									<tr>
										<td class="content_center">
											<a style="width:100%; text-align:center; line-height:12px;" class="button"
												title="Annulla la ricerca avanzata in corso."
												href="?ALL=1">
												RICERCA SEMPLICE
											</a>
										</td>
									</tr>
								<% else %>
									<tr>
										<td class="content_center">
											<a style="width:100%; text-align:center; line-height:12px;" class="button"
												title="Apre la palette di ricerca avanzata" 
												onclick="OpenRicercaAvanzata()" href="javascript:void(0);">
												RICERCA AVANZATA
											</a>
										</td>
									</tr>
								<% end if %>
								<tr>
									<td class="content_center">
										<% CALL ExportContattiInRubrica(Session("SQL_ELENCO"), "IDElencoIndirizzi", "", "") %>
									</td>
								</tr>
								<tr>
									<td class="content_center">
										<a style="width:100%; text-align:center; line-height:12px;" class="button"
											title="Apre la palette di export dei dati" 
											onclick="OpenAutoPositionedScrollWindow('ContattiExport.asp', 'export', 240, 142, true);" href="javascript:void(0);">
											EXPORT DATI
										</a>
									</td>
								</tr>
								<% if Session("ATTIVA_PARCO_MACCHINE") then %>
									<%
									Session("CONTATTI_MACCHINE_EXPORT_SQL") = _
												 " SELECT IDElencoIndirizzi AS [ID], NomeOrganizzazioneElencoIndirizzi AS [Societa], NomeElencoIndirizzi AS [NOME], SecondoNomeElencoIndirizzi [SECONDO NOME], " & _
												 " CognomeElencoIndirizzi AS [COGNOME], IndirizzoElencoIndirizzi AS [Indirizzo], CittaElencoIndirizzi AS [Citta], " & _
												 " StatoProvElencoIndirizzi AS [STATO / PROV.], CAPElencoIndirizzi AS [CAP], CountryElencoIndirizzi AS [Nazione], " & _
												 " ZonaElencoIndirizzi AS [Zona], partita_iva AS [P.IVA], ima_marchio AS [MARCHIO MACCHINA], ima_modello AS [MODELLO], " & _
												 " ima_numero AS [NUMERO MACCHINA], ima_tipocolore AS [TIPO COLORE], ima_contratto AS [CONTRATTO], " & _
												 " ima_installazione AS [INSTALLAZIONE], ima_scadenza_data AS [SCADENZA - DATA], ima_scadenza AS [SCADENZA - NOTE], " & _
												 " ima_matricola AS [MATRICOLA], ima_fornitore AS [FORNITORE], " & _
												 " CASE ima_stato_trattativa WHEN 1 THEN " & _
												 "								(CASE ISNULL(ima_esito_trattativa, 0) WHEN 0 THEN 'In corso' WHEN 1 THEN 'Vinta' WHEN 2 THEN 'Persa' END) " & _
												 " 							 ELSE '-' " & _
												 " END AS [TRATTATIVA], " & _
												 " ima_chiusura_trattativa_data AS [DATA CHIUSURA TRATTATIVA] " & _
												 " FROM tb_Indirizzario LEFT JOIN tb_indirizzario_macchine " & _
												 " ON tb_Indirizzario.IDElencoIndirizzi = tb_indirizzario_macchine.ima_contatto_id " & _
												 " WHERE (ISNULL(tb_Indirizzario.cntRel, 0) = 0) " & _
												 " AND IDElencoIndirizzi IN ("& sql_export &") " & _
												 " ORDER BY NomeOrganizzazioneElencoIndirizzi "
									%>
									<tr>
										<td class="content_center">
											<%
											CALL WRITE_EXPORT_LINK("ESPORTA MACCHINE", "DATA_ConnectionString", "CONTATTI_MACCHINE_EXPORT_SQL", FORMAT_EXCEL_FILE, false)
											%>
										</td>
									</tr>
								<% end if 
								if Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") then 
									Session("CONTATTI_RIASSUNTO_EXPORT_SQL") = _
												 " SELECT IDElencoIndirizzi AS [ID], NomeOrganizzazioneElencoIndirizzi AS [Societa], NomeElencoIndirizzi AS [NOME], SecondoNomeElencoIndirizzi [SECONDO NOME], " & _
												 " CognomeElencoIndirizzi AS [COGNOME], IndirizzoElencoIndirizzi AS [Indirizzo], CittaElencoIndirizzi AS [Citta], " & _
												 " StatoProvElencoIndirizzi AS [STATO / PROV.], CAPElencoIndirizzi AS [CAP], CountryElencoIndirizzi AS [Nazione], " & _
												 " ZonaElencoIndirizzi AS [Zona], partita_iva AS [P.IVA], " & _
												 " ( SELECT top 1 IsNull(ina_insData,ina_modData) " & _
												 "			FROM tb_indirizzario_attivita where ina_anagrafica_id = tb_Indirizzario.IDElencoIndirizzi " & _
												 "			ORDER BY IsNull(ina_insData,ina_modData) ) AS [Contatto cliente], " & _
												 " ( CASE WHEN " & _
												 " 		EXISTS( SELECT ina_id FROM tb_indirizzario_attivita where ina_anagrafica_id = tb_Indirizzario.IDElencoIndirizzi AND Isnull(ina_non_raggiungibili,0)=1 ) THEN 'Non raggiungibile'" & _
												 " 		WHEN EXISTS( SELECT ina_id FROM tb_indirizzario_attivita where ina_anagrafica_id = tb_Indirizzario.IDElencoIndirizzi AND Isnull(ina_non_interessati,0)=1 ) THEN 'No' " & _
												 "		WHEN EXISTS( SELECT ina_id FROM tb_indirizzario_attivita where ina_anagrafica_id = tb_Indirizzario.IDElencoIndirizzi AND Isnull(ina_non_raggiungibili,0)=0 AND Isnull(ina_non_interessati,0)=0) THEN 'Si' " & _
												 "		ELSE '' END) AS [Interessato], " & _
												 " ( CASE WHEN EXISTS( SELECT ina_id FROM tb_indirizzario_attivita where ina_anagrafica_id = tb_Indirizzario.IDElencoIndirizzi AND Isnull(ina_preso_appuntamento,0)=1) " & _
												 "   THEN 'Si' ELSE '' END) AS [Visita / Presentazione], " & _
												 " (CASE WHEN EXISTS( SELECT ima_id FROM tb_indirizzario_macchine WHERE ima_contatto_id = tb_Indirizzario.IDElencoIndirizzi AND IsNull(ima_stato_trattativa,0)=1 " & _
												 "	) THEN 'Si' ELSE '' END) AS [Presentazione offerta], " & _
												 " (CASE WHEN " & _
												 "		EXISTS( SELECT ima_id FROM tb_indirizzario_macchine WHERE ima_contatto_id = tb_Indirizzario.IDElencoIndirizzi AND IsNull(ima_stato_trattativa,0)=1 ) " & _
												 "	THEN " & _
												 " 		(CASE  (SELECT top 1 ima_esito_trattativa FROM tb_indirizzario_macchine WHERE ima_contatto_id = tb_Indirizzario.IDElencoIndirizzi AND IsNull(ima_stato_trattativa,0)=1 ORDER BY ima_esito_trattativa) " & _
												 "			WHEN 0 THEN 'In corso' " & _
												 "			WHEN 1 THEN 'Vinta' " & _
												 "			ELSE 'Persa' " & _
												 "  		END ) " & _
												 " 	ELSE '' END ) AS [Chiusura / vinta / persa], " & _
												 " (CASE WHEN " & _
												 "		EXISTS( SELECT ima_id FROM tb_indirizzario_macchine WHERE ima_contatto_id = tb_Indirizzario.IDElencoIndirizzi AND IsNull(ima_stato_trattativa,0)=1 ) " & _
												 "  THEN (SELECT top 1 ima_modello FROM tb_indirizzario_macchine WHERE ima_contatto_id = tb_Indirizzario.IDElencoIndirizzi AND IsNull(ima_stato_trattativa,0)=1 ORDER BY ima_esito_trattativa) " & _
												 "	ELSE '' END ) AS [Modello], " & _
												 " (SELECT top 1 ina_note FROM tb_indirizzario_attivita where ina_anagrafica_id = tb_Indirizzario.IDElencoIndirizzi ORDER BY ina_insData DESC) AS [Ultima attivita], " & _
												 " (SELECT top 1 CASE WHEN IsNull(ina_da_richiamare,0)=1 THEN 'Da richiamare il ' + CONVERT(nvarchar(10), ina_data_ricontatto, 102) " & _
												 "					  WHEN ISNULL(ina_preso_appuntamento,0)=1 THEN 'Preso appuntamento il ' + CONVERT(nvarchar(10), ina_data_appuntamento, 102) " & _
												 "				 END " & _
												 " 	FROM tb_indirizzario_attivita where ina_anagrafica_id = tb_Indirizzario.IDElencoIndirizzi ORDER BY ina_insData DESC) AS [Prossima attivita] " & _
												 " FROM tb_Indirizzario " & _
												 " WHERE (ISNULL(tb_Indirizzario.cntRel, 0) = 0) " & _
												 " AND IDElencoIndirizzi IN ("& sql_export &") " & _
												 " ORDER BY NomeOrganizzazioneElencoIndirizzi "
									%>
									<!-- <%= Session("CONTATTI_RIASSUNTO_EXPORT_SQL") %> -->
									<tr>
										<td class="content_center">
											<% CALL WRITE_EXPORT_LINK("ESPORTA RIASSUNTO", "DATA_ConnectionString", "CONTATTI_RIASSUNTO_EXPORT_SQL", FORMAT_EXCEL_FILE, false) %>
										</td>
									</tr>
									
								<% end if %>
							</table>
						</td>
					</tr>
				</table>
				</form>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>Elenco contatti</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> contatti in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="2">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<% if rs("lingua")<>"" then %>
															<td style="padding-right:4px; vertical-align:bottom;">
																<img src="../grafica/flag_mini_<%= rs("lingua") %>.jpg" alt="Lingua: <%= rs("lingua_nome_IT") %>">
															</td>
														<% end if %>
														<td style="font-size: 1px;">
															<a class="button" href="ContattiMod.asp?ID=<%= rs("IDElencoIndirizzi") %>" title="modifica dati anagrafici e rubriche">
																MODIFICA
															</a>
															<!--
															&nbsp;
															<a class="button" href="ContattiRecapiti.asp?ID=<%= rs("IDElencoIndirizzi") %>" title="gestione recapiti (telefono, fax, email, ecc..)">
																RECAPITI
															</a>-->
															&nbsp;
															<% If Application("NextCrm") then %>
																<a class="button" href="Pratiche.asp?ID=<%= rs("IDElencoIndirizzi") %>" title="apri le pratiche associate al contatto">
																	PRATICHE
																</a>
																&nbsp;
															<% End If %>
															<% if cInteger(rs("LockedByApplication"))>0 then
																sql = "SELECT sito_nome FROM tb_siti WHERE id_sito IN (" & rs("ApplicationsLocker") & "0 )"%>
																<a class="button_disabled" title="contatto non cancellabile perch&egrave; bloccato dalle applicazioni: <%= GetValueList(conn, rsr, sql) %>.">
																	CANCELLA
																</a>
															<% elseif cInteger(rs("SyncroApplication"))>0 then
																sql = "SELECT sito_nome FROM tb_siti WHERE id_sito=" & rs("SyncroApplication")%>
																<a class="button_disabled" title="contatto gestito completamente dall'applicazione: <%= GetValueList(conn, rsr, sql) %>.">
																	CANCELLA
																</a>
															<% else
																if Application("NextCrm") then 
																	sql = "SELECT (COUNT(*)) AS N_PRATICHE FROM tb_pratiche WHERE pra_cliente_id=" & rs("IDElencoIndirizzi")
																	rsv.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
																	var = rsv("N_PRATICHE")>0
																	rsv.close
																else
																	var = 0
																end if
																if var= 0 then %>
																	<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CONTATTI','<%= rs("IDElencoIndirizzi") %>');">
																		CANCELLA
																	</a>
																<% else %>
																	<a class="button_disabled" title="contatto non cancellabile perch&egrave; associato a delle pratiche.">
																		CANCELLA
																	</a>
																<% end if
															end if %>
														</td>
													</tr>
												</table>
												<%=ContactName(rs)%><%= IIF(cString(rs("TitoloElencoIndirizzi"))<>"" AND not rs("isSocieta"), ",&nbsp;" & rs("TitoloElencoIndirizzi"), "")%>
											</td>
										</tr>
										<% if rs("isSocieta") then 
											if rs("CognomeElencoIndirizzi") & rs("NomeElencoIndirizzi")<>"" then  %>
												<tr>
													<td class="label">contatto:</td>
													<td class="content"><%= rs("CognomeElencoIndirizzi") %>&nbsp;<%= rs("NomeElencoIndirizzi") %><%= IIF(cString(rs("TitoloElencoIndirizzi"))<>"", ",&nbsp;" & rs("TitoloElencoIndirizzi"), "") %></td>
												</tr>
											<% end if
										else 
											if rs("NomeOrganizzazioneElencoIndirizzi")<>"" then  %>
												<tr>
													<td class="label">ente:</td>
													<td class="content"><%= rs("NomeOrganizzazioneElencoIndirizzi") %></td>
												</tr>
											<% end if
										end if %>
										<tr>
											<td class="label" style="width:22%;">rubriche:</td>
											<% sql = " SELECT '<span style=""white-space:nowrap"">' + nome_rubrica + '</span>' FROM tb_rubriche " &_
													 " WHERE tb_rubriche.id_rubrica IN ( SELECT id_rubrica FROM rel_rub_ind WHERE rel_rub_ind.id_indirizzo=" & rs("IDElencoIndirizzi") & ") " & _
													 " ORDER BY nome_rubrica " 
											%>
											<td class="content"><%= GetValueList(conn, rsr, sql) %></td>
										</tr>
										<tr>
											<td class="label">indirizzo:</td>
											<td class="content">
												<%= ContactAddress(rs) %>
											</td>
										</tr>
										<% sql = "SELECT * FROM tb_TipNumeri WHERE id_tipoNumero " &_
												 " IN (SELECT id_TipoNumero FROM tb_ValoriNumeri WHERE id_indirizzario=" & rs("IDElencoIndirizzi") & ") "
										rsv.Open sql, conn, AdOpenForwardOnly, adLockReadOnly, adCmdText
										while not rsv.eof
											sql = "SELECT id_TipoNumero, ValoreNumero FROM tb_ValoriNumeri " &_
												  " WHERE id_TipoNumero=" & rsv("id_tipoNumero") & " AND  id_Indirizzario=" & rs("IDElencoIndirizzi")
											rsr.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
											if not rsr.eof then%>
												<tr>
													<td class="label" nowrap><%= Lcase(rsv("nome_TipoNumero")) %>:</td>
													<td class="content">
														<%while not rsr.eof
															select case rsr("id_TipoNumero")
																case 6	'email 
																%>
																	<a href="mailto:<%= rsr("ValoreNumero") %>"><%= rsr("ValoreNumero") %></a>
																<% case 7	'web 
																%>
																	<a href="http://<%= rsr("ValoreNumero") %>" target="_blank"><%= rsr("ValoreNumero") %></a>
																<% case else %>
																	<%= rsr("ValoreNumero") %>
															<%end select
															rsr.movenext
															if not rsr.eof then%>
																,&nbsp;
															<%end if
														wend %>
													</td>
												</tr>
											<%end if
											rsr.close
											rsv.MoveNext
										wend
										rsv.Close
										if isDate(rs("DataIscrizione")) then%>
											<tr>
												<td class="label">data iscrizione:</td>
												<td class="content">
													<%= DateTimeIta(rs("DataIscrizione")) %>
												</td>
											</tr>
										<% 	end if 
										sql = " SELECT IDElencoIndirizzi, CognomeElencoIndirizzi, NomeElencoIndirizzi, QualificaElencoIndirizzi, isSocieta, " + _
											  " NomeOrganizzazioneElencoIndirizzi " + _
											  " FROM tb_indirizzario WHERE CntRel=" & rs("IDElencoIndirizzi") & " ORDER BY ModoRegistra "
										rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
										if not rsr.eof then%>
											<tr>
												<td class="label">contatti interni / sedi alternative:</td>
												<td>
													<% if rsr.recordcount>2 then %>
														<span class="overflow">
													<% end if %>
														<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
															<tr>
																<th class="L2">contatto / sede</th>
																<th class="l2_center" width="10%" nowrap>sede</th>
																<th class="l2_center" width="10%" colspan="1">operazioni</th>
															</tr>
															<% while not rsr.eof %>
																<tr>
																	<td class="content" title="ruolo / qualifica:<%= rsr("qualificaElencoIndirizzi") %>"><%= ContactFullName(rsr) %></td>
																	<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(rsr("isSocieta")) %> title="<%= IIF(rsr("isSocieta"), "sede alternativa o periferica", "contatto interno") %>"></td>
																	<td class="content_center">
																		<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedWindow('ContattiInterniMod.asp?CNT=<%= rs("IdElencoIndirizzi") %>&ID=<%= rsr("IdElencoIndirizzi") %>', 'cntInt', 950, 310)">
																			MODIFICA
																		</a>
																	</td>
																</tr>
																<% rsr.movenext
															wend %>
														</table>
													<% if rsr.recordcount>2 then %>
														</span>
													<% end if %>
												</td>
											</tr>
										<% 	end if
											rsr.close
										%>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
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
<% rs.close %>

<% if Session("ATTIVA_PARCO_MACCHINE") or Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") then %>
	<%
	dim queryString

	'query telefonate o visite
	sql = " SELECT (CASE ISNULL(ina_da_richiamare, 0) WHEN 1 THEN ina_data_ricontatto ELSE ina_data_appuntamento END) AS DATA_IMPEGNO, " & _
		  " ina_note AS TESTO_IMPEGNO, " & _
		  " (CASE ISNULL(ina_da_richiamare, 0) WHEN 1 THEN 'Da ricontattare il ' ELSE 'Appuntamento il ' END) AS TIPO_IMPEGNO, " & _
		  " ina_anagrafica_id AS ID_ANAGRAFICA, " & _
		  " ina_insAdmin_id AS ID_ADMIN, " & _
		  "	0 as IS_MACCHINE, ina_da_richiamare, " & _
		  " IDElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, NomeElencoIndirizzi, CognomeElencoIndirizzi, IsSocieta " & _
		  " FROM tb_indirizzario_attivita INNER JOIN tb_indirizzario ON tb_indirizzario_attivita.ina_anagrafica_id = tb_indirizzario.IdelencoIndirizzi " & _
		  " WHERE ((ina_da_richiamare = 1 AND ISNULL(ina_richiamare_fatto, 0) = 0) OR " & _
		  " (ina_preso_appuntamento = 1 AND ISNULL(ina_appuntamento_fatto, 0) = 0)) " & _
		  " AND ina_anagrafica_id IN (" & sql_export & ") " & _
		  " ORDER BY (CASE ISNULL(ina_da_richiamare, 0) WHEN 1 THEN ina_data_ricontatto ELSE ina_data_appuntamento END), ina_insData"
		  '" AND (" & SQL_CompareDateTime(conn, "(CASE ISNULL(ina_da_richiamare, 0) WHEN 1 THEN ina_data_ricontatto ELSE ina_data_appuntamento END)", adCompareGreaterThan, Now()) & ") " & _
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText

	'query scadenza contratti macchine
	sql = " SELECT DISTINCT ima_scadenza_data, IDElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, NomeElencoIndirizzi, CognomeElencoIndirizzi, IsSocieta, ima_stato_trattativa, ima_esito_trattativa " & _
		  " FROM tb_indirizzario_macchine INNER JOIN tb_indirizzario ON tb_indirizzario_macchine.ima_contatto_id = tb_indirizzario.IdelencoIndirizzi " & _
		  " WHERE ima_scadenza_data IS NOT NULL " & _
		  " AND ima_contatto_id IN (" & sql_export & ") " & _
		  " AND " & SQL_CompareDateTime(conn, "ima_scadenza_data", adCompareGreaterThan, Now()) & _
		  " AND (IsNull(ima_stato_trattativa, 0) = 0 OR IsNull(ima_esito_trattativa,0)=1 OR IsNull(ima_esito_trattativa,0)=2 ) " & _
		  " ORDER BY ima_scadenza_data "
	rsv.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText

	'query trattative in corso
	sql = " SELECT DISTINCT ima_scadenza_data, IDElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, NomeElencoIndirizzi, CognomeElencoIndirizzi, ModoRegistra, IsSocieta " & _
		  " FROM tb_indirizzario_macchine INNER JOIN tb_indirizzario ON tb_indirizzario_macchine.ima_contatto_id = tb_indirizzario.IdelencoIndirizzi " & _
		  " WHERE ima_contatto_id IN (" & sql_export & ") " & _
		  " AND IsNull(ima_stato_trattativa, 0) = 1 and IsNull(ima_esito_trattativa,0)=0 " & _
		  " ORDER BY NomeOrganizzazioneElencoIndirizzi, CognomeElencoIndirizzi"
	rst.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText

	if not rs.eof OR not rsv.eof OR not rst.eof then %>
		<div id="pulsanti" style="position:absolute; top:93px; left:790px; width:680px; text-align:center;">
			<table cellspacing="0" cellpadding="0" width="100%">
				<tr>
					<td style="vertical-align:top;">
						<% if not rs.eof then %>
							
							<table cellspacing="1" cellpadding="0" class="tabella_madre contattiattivita" style="width:360px; margin-right:15px;">
								<caption class="">Elenco prossimi impegni e appuntamenti</caption>
								<tr>
									<th style="width:52%;">anagrafica</th>
									<th>impegno</th>
									<th style="width:23%;">data</th>
								</tr>
								<tr>
									<td colspan="3">
										<div style="height:900px; overflow-y:scroll; width:100%;">
											<table cellspacing="1" cellpadding="0" style="width:100%;">
												<% while not rs.eof %>
													<%
													dim scaduto
													scaduto = false
													if (isDate(rs("DATA_IMPEGNO")) AND DateISO(rs("DATA_IMPEGNO")) < DateISO(Now())) then
														scaduto = true
													end if
													%>
													<tr>
														<td style="width:55%;" class="content<%=IIF(rs("ina_da_richiamare"), "", " evidenzia")%>">
															<% if scaduto then %>
																<img style="padding:3px; padding-left:0px; float:left;" src="<%= GetSiteUrl(null, 0, 0) & "/amministrazione/grafica/attivita-dimenticata.gif" %>" />
																<b><%= ContactLinkedNameExtra(rs, false, "#IFrameContattiAttivita") %></b>
															<% else %>
																<%= ContactLinkedNameExtra(rs, false, "#IFrameContattiAttivita") %>
															<% end if %>
														</td>
														<td class="content<%=IIF(rs("ina_da_richiamare"), "", " evidenzia")%>">
															<% if scaduto then %><b><% end if %>
															<%= rs("TIPO_IMPEGNO")%>
															<% if scaduto then %></b><% end if %>
														</td>
														<td  style="width:13%;" class="content_right<%=IIF(rs("ina_da_richiamare"), "", " evidenzia")%>">
															<% if scaduto then %><b><% end if %>
															<%= rs("DATA_IMPEGNO") %>
															<% if scaduto then %></b><% end if %>
														</td>
													</tr>
													<% rs.moveNext
												wend %>
											</table>
										</div>
									</td>
								</tr>
							</table>
						<% end if %>
					</td>
					<td style="vertical-align:top;">
						<% if not rst.eof then %>
							<table cellspacing="1" cellpadding="0" class="tabella_madre contattimacchine" style="width:360px; margin-bottom:15px;">
								<caption class="">Elenco trattative in corso</caption>
								<tr>
									<th>cliente</th>
								</tr>
								<% while not rst.eof
									%>
									<tr>
										<td class="content trattative">
											<%= ContactLinkedNameExtra(rst, false, "#IFrameMacchine") %>
										</td>
									</tr>
									<% rst.moveNext
								wend %>
							</table>
						<% end if %>
						<% if not rsv.eof then %>
							<table cellspacing="1" cellpadding="0" class="tabella_madre contattimacchine" style="width:360px;">
								<caption class="">Elenco scadenze contratti macchine</caption>
								<tr>
									<th style="width:70%;">anagrafica</th>
									<th class="right">data scadenza</th>
								</tr>
								<%
								while not rsv.eof
									%>
									<tr>
										<td class="content">
											<% if rsv("ima_stato_trattativa") then
												if cIntero(rsv("ima_esito_trattativa")) = 1 then
													CALL WriteColoreTipo(COLORE_CONTATTI_TRATTATIVE, "Trattativa vinta")
												else
													CALL WriteColoreTipo("#FF0606", "Trattativa persa")
												end if
											end if %>
											<%= ContactLinkedNameExtra(rsv, false, "#IFrameMacchine") %>
										</td>
										<td class="content_right">
											<%= rsv("ima_scadenza_data") %>
										</td>
									</tr>
									<%
									rsv.moveNext
								wend
								%>
							</table>			
						<% end if %>
					</td>
				<tr>
			</table>
		</div>
	<% end if %>
	<% rs.close
	rst.close
	rsv.close %>
<% end if %>
</body>
</html>
<% 
conn.close 
set rs = nothing
set rsv = nothing
set rsr = nothing
set conn = nothing
%>
 