<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.buffer = true %>
<% response.charset = "UTF-8" %>

<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ContattiSalva.asp")
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Anagrafica contatti - modifica"
'Indirizzo pagina per link su sezione 
	'HREF = "Contatti.asp;ContattiRecapiti.asp?ID=" & request("ID")
	HREF = "Contatti.asp"
'Azione sul link: {BACK | NEW}
	'Action = "INDIETRO;RECAPITI"
	Action = "INDIETRO"
If Application("NextCrm") then
	HREF = HREF & ";Pratiche.asp?ID=" & request("ID")
	Action = Action & ";PRATICHE"
end if
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

dim conn, rs, rsr, rsa, sql, rubriche_visibili, isLocked, textStyle, value, iframe_url, sql_campagne, iframe_ID_macchine, iframe_ID_attivita

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_ELENCO"), "IDElencoIndirizzi", "ContattiMod.asp")
end if


'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rsr)

sql = "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext

'controlla se il contato e' bloccato
if cInteger(rs("SyncroApplication"))>0 then
	isLocked = "disabled"
	textStyle = "text_disabled"
else
	isLocked = ""
	textStyle= "text"
end if
%>
<script language="JavaScript" type="text/javascript">
	function set_modo_registra(){
		var isSocieta = document.getElementById('chk_issocieta_true');
		if (isSocieta.checked)
			form1.tft_modoregistra.value = form1.tft_nomeorganizzazioneelencoindirizzi.value;
		else
			form1.tft_modoregistra.value = form1.tft_cognomeelencoindirizzi.value;
		return true;
	}
	
	function show_mandatory(){
		var isSocieta = document.getElementById('chk_issocieta_true');
		var span_nome = document.getElementById('nome')
		var span_cognome = document.getElementById('cognome')
		var span_ente = document.getElementById('ente')

		if (isSocieta.checked){
			span_ente.innerHTML='(*)'
			span_cognome.innerHTML=''
			span_nome.innerHTML=''
		}
		else{
			span_ente.innerHTML=''
			span_cognome.innerHTML='(*)'
			span_nome.innerHTML='(*)'
		}
		
	}
</script>
<div id="content_float" style="min-width: 1200px;">
	<script language="JavaScript" type="text/javascript">
		function ShowDatiAggiuntivi(state){
			if (document.getElementById("Agg1").style.visibility == "visible" || state == "hide"){
				document.getElementById("Agg1").style.visibility = 'hidden';
				document.getElementById("Agg1").style.display = 'none';
				document.getElementById("Agg2").style.visibility = 'hidden';
				document.getElementById("Agg2").style.display = 'none';
				document.getElementById("Agg3").style.visibility = 'hidden';
				document.getElementById("Agg3").style.display = 'none';
				document.getElementById("PulsanteAgg").innerHTML = 'Mostra dati aggiuntivi';
			}
			else
			{
				document.getElementById("Agg1").style.visibility = 'visible';
				document.getElementById("Agg1").style.display = '';
				document.getElementById("Agg2").style.visibility = 'visible';
				document.getElementById("Agg2").style.display = '';
				document.getElementById("Agg3").style.visibility = 'visible';
				document.getElementById("Agg3").style.display = '';
				document.getElementById("PulsanteAgg").innerHTML = 'Nascondi dati aggiuntivi';
			}
		}
	</script>
	<form action="" method="post" id="form1" name="form1" onsubmit="set_modo_registra();">
	<input type="hidden" name="tft_modoregistra" value="">
	<input type="hidden" name="isLocked" value="<%= isLocked %>">
	<input type="hidden" name="LockedByApplication" value="<%= rs("LockedByApplication") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="width:99%;">
		<caption colspan="2">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption" width="48%;">
						<table cellspacing="0" cellpadding="0" style="width:99%;">
							<tr>
								<td class="caption">Modifica dati del contatto</td>
								<td align="right" style="font-size:10px;">
									<a href="javascript:void(0)" class="button_L2" onclick="ShowDatiAggiuntivi('');" id="PulsanteAgg">Mostra dati aggiuntivi</a>
									&nbsp;&nbsp;&nbsp;&nbsp;
									<a class="button_L2" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="contatto precedente">
										&lt;&lt; PRECEDENTE
									</a>
									&nbsp;
									<a class="button_L2" href="?ID=<%= request("ID") %>&goto=NEXT" title="contatto successivo">
										SUCCESSIVO &gt;&gt;
									</a>
								</td>
							</tr>
						</table>
					</td>
					<td align="right" style="font-size:10px;">
						<input style="width:25%" type="button" class="button" id="salva_eltop" name="salva_eltop" value="SALVA & TORNA ALL'ELENCO" onclick="SubmitFrames('elenco');">
						<input style="width:10%;" type="button" class="button" id="salva_top" name="salva" value="SALVA" onclick="SubmitFrames('')">
					</td>
				</tr>
			</table>
		</caption>
		<tr>
			<td id="prima_colonna" style="width:48%; border-right:1px solid #999999; vertical-align:top;">
				<table cellspacing="1" cellpadding="0" style="" style="width:100%;">
					<tr>
						<% if isLocked<>"" then %>
							<th colspan="5">ANAGRAFICA&nbsp;&nbsp;( parzialmente modificabile )</th>
						<% else %>
							<th colspan="4">ANAGRAFICA</th>
						<% end if %>
					</tr>
					<tr>
						<td class="label_no_width" style="width:15%;">salva come:</td>
						<td class="content" style="width:45%;">
							<table border="0" cellspacing="0" cellpadding="0" align="left">
								<tr>
									<td><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_false" value="" <%= chk(not rs("isSocieta"))%> onClick="show_mandatory()"></td>
									<td width="28%">persona fisica</td>
									<td><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_true" value="1" <%= chk(rs("isSocieta"))%> onClick="show_mandatory()"></td>
									<td>ente / societ&agrave; / organizzazione</td>
								</tr>
							</table>
						</td>
						<td class="label_no_width">lingua comunicazioni:</td>
						<td class="content" style="width:15%;" <%= IIF(isLocked<>"", " colspan=""2"" ", "") %>>
							<% CALL DropLingue(conn, NULL, "tft_lingua", rs("lingua"), true, false, "width:100%;") %>
						</td>
					</tr>
					<tr>
						<td class="label_no_width">ente:</td>
						<td class="content" colspan="3">
							<input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_nomeorganizzazioneelencoindirizzi" value="<%= rs("NomeOrganizzazioneElencoIndirizzi") %>" maxlength="250" style="width:95%;">
							<span id="ente">(*)</span>
						</td>
						<% if isLocked<>"" then %>
							<th rowspan="11" class="vertical" style="width:15px;" title="Sorgente: <%= rs("SyncroTable") %>; Chiave: <%= rs("SyncroKey") %>">Dati sincronizzati automaticamente</th>
						<% end if %>
					</tr>
					<tr>
						<td class="label_no_width">titolo:</td>
						<td class="content" colspan="3"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_TitoloElencoIndirizzi" value="<%= rs("TitoloElencoIndirizzi") %>" maxlength="50" style="width:22%;"></td>
					</tr>
					<tr>
						<td class="label_no_width">nome:</td>
						<td class="content" colspan="3">
							<input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_nomeelencoindirizzi" value="<%= rs("NomeElencoIndirizzi") %>" maxlength="100" style="width:70%;">
							<span id="nome">(*)</span>
						</td>
					</tr>
					<tr id="Agg1">
						<td class="label_no_width" style="width:20%;">secondo nome:</td>
						<td class="content" colspan="3"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_secondonomeelencoindirizzi" value="<%= rs("SecondoNomeElencoIndirizzi") %>" maxlength="30" style="width:70%;"></td>
					</tr>
					<tr>
						<td class="label_no_width">cognome:</td>
						<td class="content" colspan="3">
							<input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_cognomeelencoindirizzi" value="<%= rs("CognomeElencoIndirizzi") %>" maxlength="100" style="width:70%;">
							<span id="cognome">(*)</span>
						</td>
					</tr>
					<tr id="Agg2">	
						<td class="label_no_width">ruolo / qualifica:</td>
						<td class="content" colspan="3"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_qualificaelencoindirizzi" value="<%= rs("qualificaelencoindirizzi") %>" maxlength="250" style="width:56%;"></td>
					</tr>
					<tr>	
						<td class="label_no_width">codice fiscale:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_CF" value="<%= rs("CF") %>" maxlength="16" style="width:60%;"></td>
						<td class="label_no_width">partita i.v.a.:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_partita_iva" value="<%= rs("partita_iva") %>" maxlength="11" style="width:100%;"></td>
					</tr>
					<tr id="Agg3">
						<td class="label_no_width">luogo di nascita:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_luogonascita" value="<%= rs("luogonascita") %>" maxlength="255" style="width:100%;"></td>
						<td class="label_no_width">data di nascita:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tfd_dtnascelencoindirizzi" value="<%= rs("DTNASCElencoIndirizzi") %>" maxlength="10" style="width:100%;"></td>
					</tr>
					<% if Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE") then %>
						<tr>
							<td class="label">categoria:</td>
							<td class="content" colspan="3">
								<%CALL dropDown(conn, CatContatti.QueryElenco(true, ""), "icat_id", "NAME", "tfn_cnt_categoria_id", IIF(cInteger(request("tfn_cnt_categoria_id"))>0, cInteger(request("tfn_cnt_categoria_id")), rs("cnt_categoria_id")), false, " onchange='form1.submit()'", LINGUA_ITALIANO)%>
							</td>
						</tr>
					<% end if %>
					<script language="JavaScript" type="text/javascript">
						show_mandatory();
					</script>
					<tr>
						<% if isLocked<>"" then %>
							<th colspan="4">INDIRIZZO&nbsp;&nbsp;( non modificabile )</th>
						<% else %>
							<th colspan="4">INDIRIZZO</th>
						<% end if %>
					</tr>
					<tr>
						<td class="label_no_width">indirizzo:</td>
						<td class="content" colspan="3"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_IndirizzoElencoIndirizzi" value="<%= rs("IndirizzoElencoIndirizzi") %>" maxlength="250" style="width:100%;"></td>
					</tr>
					<tr>
						<td class="label_no_width">localit&agrave;:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_LocalitaElencoIndirizzi" value="<%= rs("LocalitaElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
						<td class="label_no_width">cap:</td>
						<td class="content" ><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_CAPElencoIndirizzi" value="<%= rs("CAPElencoIndirizzi") %>" maxlength="20" style="width:100%;"></td>
					</tr>
					<tr>
						<td class="label_no_width">citt&agrave;:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_cittaElencoIndirizzi" value="<%= rs("cittaElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
						<td class="label_no_width">provincia / stato:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_StatoProvElencoIndirizzi" value="<%= rs("StatoProvElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
					</tr>
					<tr>
						<td class="label_no_width">zona:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_ZonaElencoIndirizzi" value="<%= rs("ZonaElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
						<td class="label_no_width">nazione:</td>
						<td class="content"><input <%= isLocked %> type="text" class="<%= TextStyle %>" name="tft_CountryElencoIndirizzi" value="<%= rs("CountryElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
					</tr>
					
					<% If (Session("COM_ADMIN") & Session("COM_POWER") <> "") AND _
							InStr(Application("NextCom_codice"), "<PREFISSOCLIENTE>") > 0 AND _
							Application("NextCrm") then
						dim PraticheCount
						sql = "SELECT COUNT(*) FROM tb_pratiche WHERE pra_cliente_id=" & rs("IDElencoIndirizzi")
						PraticheCount = GetValueList(conn, rsr, sql)%>
						<tr><th colspan="4">PARAMETRI GESTIONE PRATICHE</th></tr>
						<tr>
							<td class="label_no_width">prefisso:</td>
							<td class="content" colspan="3">
								<input type="text" class="text" name="tft_PraticaPrefisso" value="<%= rs("PraticaPrefisso") %>" maxlength="5" size="5" <%= IIF(PraticheCount, " disabled ", "") %>>
								<span class="note">
									<% if PraticheCount>0 then %>
										Prefisso non modificabile perch&egrave; sono presenti n&deg;<%= PraticheCount %> pratiche registrate.
									<% else %>
										Prefisso per la generazione del codice di ogni pratica.
									<% end if %>
								</span>
							</td>
						</tr>
						<% If InStr(Application("NextCom_codice"), "<COUNTCLIENTE>") > 0 then %>
							<tr>
								<td class="label_no_width">progressivo:</td>
								<td class="content" colspan="3">
									<input type="text" class="text" name="tfn_PraticaCount" value="<%= rs("PraticaCount") %>" maxlength="5" size="5" <%= IIF(PraticheCount, " disabled ", "") %>>
									<span class="note">
										<% if PraticheCount>0 then %>
											progressivo di pratica non modificabile perch&egrave; sono presenti n&deg;<%= PraticheCount %> pratiche registrate.
										<% else %>
											Numero progressivo con il quale viene calcolato il codice della pratica.
										<% end if %>
									</span>
								</td>
							</tr>
						<%end if
					end if%>
					<% if Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE") then %>
						<tr><th colspan="4">CARATTERISTICHE</th></tr>
					<%	sql = " SELECT * FROM (tb_indirizzario_carattech " + _
							  "	INNER JOIN rel_categ_ctech ON (tb_indirizzario_carattech.ict_id = rel_categ_ctech.rcc_ctech_id " & _
							  "						AND rel_categ_ctech.rcc_categoria_id=" & IIF(cInteger(request("tfn_cnt_categoria_id"))>0, cInteger(request("tfn_cnt_categoria_id")), cInteger(rs("cnt_categoria_id"))) & ") )" + _
							  " LEFT JOIN rel_cnt_ctech ON (tb_indirizzario_carattech.ict_id=rel_cnt_ctech.ric_ctech_id AND rel_cnt_ctech.ric_cnt_id="& CInteger(request("ID")) &") " + _
							  " LEFT JOIN tb_indirizzario_carattech_raggruppamenti ON tb_indirizzario_carattech.ict_raggruppamento_id = tb_indirizzario_carattech_raggruppamenti.icr_id " + _
							  " ORDER BY tb_indirizzario_carattech_raggruppamenti.icr_ordine, rel_categ_ctech.rcc_ordine "
						'CALL DesElenco_EXT(conn, sql, "tb_indirizzario_carattech", "ict_id", "ict_nome_it", "ict_tipo", "", "ric_valore_", "ric_valore_", false, 4)
						CALL DesForm  (conn, sql, "tb_indirizzario_carattech", "ict_id", "ict_nome_it", "ict_tipo", "ict_unita_it", "", "ric_valore_", "ric_valore_", "icr_titolo_it", cIntero(request("ID")) = 0, 4)
					end if%>
					<% if Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") then %>
							<tr>
								<td colspan="4">
									<table class="campagnemarketing" cellspacing="1" cellpadding="1" style="width:100%;">
										<tr>
											<th colspan="4">CAMPAGNE MARKETING</th>
										</tr>
										<% 
										'query campagne collegate 
										sql = " SELECT * FROM tb_indirizzario_campagne c INNER JOIN rel_cnt_campagne r ON  " & _
												 " c.inc_id = r.rcc_campagna_id AND r.rcc_cnt_id = " & request("ID") & _
												 " ORDER BY inc_nome "
										rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdtext
										dim rowspan 
										rowspan = cIntero(rsr.recordCount)
										
										'query campagne da collegare
										sql_campagne = " SELECT * FROM tb_indirizzario_campagne WHERE inc_id NOT IN " & _
													 " (SELECT rcc_campagna_id FROM rel_cnt_campagne WHERE rcc_cnt_id = "&request("ID")&")" & _
													 " ORDER BY inc_nome "
										rsa.open sql_campagne, conn, adOpenStatic, adLockReadOnly, adCmdtext

										if rsr.eof then
											%>
											<tr>
												<td class="label" style="width:23%;">
													campagne marketing collegate:
												</td>
												<td class="content" colspan="3">(nessuna)</td>
											</tr>
										<% end if %>
										<% while not rsr.eof %>
											<tr>
												<% if cIntero(rsr.AbsolutePosition) = 1 then %>
													<td class="label" style="width:23%;" rowspan="<%=rowspan%>">
														campagne marketing collegate:
													</td>
												<% end if %>
												<% if isDate(rsr("rcc_data_conclusione")) then  %>
													<td class="content" colspan="2">
														<%= rsr("inc_nome") %>
													</td>
													<td class="note" style="width:35%;">
														conclusa in data: <%= rsr("rcc_data_conclusione") %>
													</td>
												<% else %>
													<td class="content" colspan="2">
														<%= rsr("inc_nome") %>
													</td>
													<td class="content_right">
														<% if session("COM_ADMIN")<>"" then %>
															<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('CAMPAGNA_CONTATTO','<%= rsr("rcc_id") %>');">
																rimuovi
															</a>
														<% end if %>
														&nbsp;
													<td>
												<% end if %>
											</tr>
											<% rsr.moveNext %>
										<% wend %>
										<% if rsa.recordCount > 0 then %>
											<tr>
												<td class="content_right" colspan="4">
													<input type="checkbox" class="noborder" name="abilita_aggiunta_campagna" value="<%=request("abilita_aggiunta_campagna")%>" onclick="EnableIfChecked(this, document.getElementById('ext_new_campagna_id'));" />
													aggiungi campagna: 
													<% CALL dropDown(conn, sql_campagne, "inc_id", "inc_nome", "ext_new_campagna_id", cInteger(request("ext_new_campagna_id")), false, "", LINGUA_ITALIANO)%>
													<% if cInteger(request("ext_new_campagna_id")) = 0 then %>
														<script language="JavaScript" type="text/javascript">
															document.getElementById('ext_new_campagna_id').disabled = true;
															document.getElementById('ext_new_campagna_id').className = 'disabled';
														</script>
													<% end if %>
												</td>
											</tr>
											<%
										end if
										rsa.close %>
										</tr>
									</table>
								</td>
							</tr>
						<% rsr.close %>
					<% end if %>
					<tr>
						<th colspan="4">RUBRICHE (*)</th>
					</tr>
					<%  sql = " SELECT DISTINCT nome_rubrica FROM tb_rubriche " &_
							  " INNER JOIN rel_rub_ind ON tb_rubriche.id_rubrica=rel_rub_ind.id_rubrica " &_
							  " WHERE " & SQL_IsTrue(conn, "tb_rubriche.rubrica_esterna") & _
							  " AND rel_rub_ind.id_indirizzo=" & rs("IDElencoIndirizzi") & _
							  " ORDER BY nome_rubrica "
					value = GetValueList(conn, NULL, sql)
					if value<>"" then %>
						<tr>
							<td class="label_no_width">non modificabili:</td>
							<td class="content content_checked" colspan="3"><%= replace(value, ", ", "<br>") %></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="4">
							<% sql = "SELECT tb_rubriche.id_rubrica, tb_rubriche.nome_rubrica, rel_rub_ind.id_rub_ind " &_
									 " FROM tb_rubriche LEFT JOIN rel_rub_ind ON (tb_rubriche.id_Rubrica = rel_rub_ind.id_rubrica " &_
									 " AND rel_rub_ind.id_indirizzo=" & cIntero(request("ID")) & ")" & _
									 " WHERE tb_rubriche.id_rubrica IN (" & GetList_Rubriche(conn, rsr) & ")" &_
									 " AND NOT(" & SQL_IsTrue(conn, "tb_rubriche.rubrica_esterna") & ") " &_
									 " ORDER BY nome_rubrica"
							CALL Write_Relations_Checker(conn, rsr, sql, 3, "id_rubrica", "nome_rubrica", "id_rub_ind", "rubriche")%>
						</td>
						<% if  isLocked<>"" then %>
							<th rowspan="11" class="vertical">Dati modificabili</th>
						<% end if %>
					</tr>
					
				</table>
			</td>
			<td id="seconda_colonna" style="vertical-align:top;">
				<%
				iframe_url = GetUrl() & "/amministrazione/nextCom/ContattiRecapiti_iFrame.asp?MODE=iframe&ID=" & request("ID")
				iframe_ID_macchine = 0
				%>
				<iframe id="IFrameRecapiti" name="" style="border:0px; width:100%;" src="<%= iframe_url %>"  frameborder="0" scrolling="no">
				</iframe>
				<%
				'CONTATTI INTERNI
				iframe_url = GetUrl() & "/amministrazione/nextCom/ContattiInterni_iFrame.asp?ID=" & rs("IDElencoIndirizzi")
				iframe_ID_macchine = iframe_ID_macchine + 1
				%>
				<iframe id="IFrameContattiInterni" style="border:0px; width:100%;" src="<%= iframe_url %>"  frameborder="0" scrolling="no"></iframe>
				
				<% if Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") then %>
					<%
					'ATTIVITA' CON I CONTATTI
					iframe_url = GetUrl() & "/amministrazione/nextCom/ContattiAttivita_iFrame.asp?ID=" & rs("IDElencoIndirizzi")
					iframe_ID_macchine = iframe_ID_macchine + 1
					iframe_ID_attivita = iframe_ID_macchine
					%>
					<iframe id="IFrameContattiAttivita" style="border:0px; width:100%;" src="<%= iframe_url %>"  frameborder="0" scrolling="no"></iframe>
				<% end if %>
				
				<table id="text_area_container" border="0" cellspacing="0" cellpadding="0" align="left" style="width:100%;">
					<tr><th colspan="4">NOTE SUL CONTATTO</th></tr>
					<tr>
						<td class="content" colspan="4">
							<textarea id="text_area" style="width:100%;" rows="4" cols="" name="tft_NoteElencoIndirizzi"><%=rs("NoteElencoIndirizzi")%></textarea>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			function SetTextAreaHeight(){
				var Hpixels;
				var colonnaContainer = document.getElementById('seconda_colonna');
				var textArea = document.getElementById('text_area');
				var textAreaContainer = document.getElementById('text_area_container');
				
				Hpixels = (colonnaContainer.clientHeight - (textAreaContainer.offsetTop + textAreaContainer.clientHeight) + textArea.offsetHeight) - 2;
				//alert(Hpixels);
				if (Hpixels > textArea.offsetHeight){
					textArea.style.height = Hpixels;
				}
			}
		</script>
		
		<% if Session("ATTIVA_PARCO_MACCHINE") then %>
			<tr>
				<td colspan="2">
					<% 'GESTIONE PARCO MACCHINE
					iframe_url = GetUrl() & "/amministrazione/nextCom/ContattiMacchine.asp?ID=" & rs("IDElencoIndirizzi")
					iframe_ID_macchine = iframe_ID_macchine + 1
					%>
					<iframe id="IFrameMacchine" style="border:0px; width:100%;" src="<%= iframe_url %>"  frameborder="0" scrolling="no">
					</iframe>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="2">
				<% if isDate(rs("DataIscrizione")) then %>
					<span style="float:left;">data iscrizione: <%= DateTimeIta(rs("DataIscrizione")) %></span>
				<% end if %>
				(*) Campi obbligatori.
				<input type="hidden" name="salva_elenco" id="salva_elenco" value="">
				<input style="width:14%" type="button" class="button" name="salva_el" id="salva_el" value="SALVA & TORNA ALL'ELENCO" onclick="SubmitFrames('elenco');">
				<input style="width:5%;" type="button" class="button" name="salva" id="salva" value="SALVA" onclick="SubmitFrames('')">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
	<% if Trim(rs("SecondoNomeElencoIndirizzi")&rs("qualificaelencoindirizzi")&rs("luogonascita")&rs("DTNASCElencoIndirizzi"))="" then  %>
		<script language="JavaScript" type="text/javascript">
			ShowDatiAggiuntivi('hide');
		</script>
	<% else %>
		<script language="JavaScript" type="text/javascript">
			ShowDatiAggiuntivi('show');
		</script>
	<% end if %>
	<script language="JavaScript" type="text/javascript">
		function SubmitFrames(redirect){
			if (redirect == 'elenco'){
				document.getElementById("salva_elenco").value = "salva_elenco";
			}
		
			<% if Session("ATTIVA_PARCO_MACCHINE") then %>
				//salva parco macchine
				window.frames[<%=iframe_ID_macchine%>].document.forms[0].submit();
			<% end if %>
			//salva recapiti
			window.frames[0].document.forms[0].submit();
			<% if Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") then %>
				window.frames[<%=(iframe_ID_attivita)%>].document.forms[0].submit();
			<% end if %>
		}

		setTimeout(function()
		{
			SetTextAreaHeight();
		}, 2000);
	</script>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsr = nothing
set rsa = nothing
set conn = nothing%>
