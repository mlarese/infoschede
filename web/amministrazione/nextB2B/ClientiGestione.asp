<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<%
dim conn, rs, rsr, rsa, sql, OBJ_Contatto, id_ext
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")
set OBJ_contatto = new IndirizzarioLock

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_RIVENDITORI_SQL"), "IDElencoIndirizzi", "ClientiGestione.asp")
end if

if request("IDCNT") <> "" OR request("ID") <> "" then
	OBJ_contatto.LoadFromDB(cInteger(request("IDCNT") & request("ID")))
	if request("ID") <> "" then
		id_ext = GetValueList(conn, rs, "SELECT riv_id FROM gv_rivenditori WHERE IDElencoIndirizzi="& cIntero(request("ID")))
	end if
else
	OBJ_contatto.LoadFromForm("extC_riv_attivo;isSocieta;chk_abilitato")
end if

if request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim pagina_redirect
	pagina_redirect = ""

	%>
	<!--#INCLUDE FILE="ClientiGestioneSalva_Tools.asp" -->
	<%	
end if

'gestione dei campi esterni a IndirizzarioLock come se fossero interni
CaricaCampiEsterni conn, rs, OBJ_contatto, "SELECT * FROM gtb_rivenditori", "riv_id", id_ext
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione clienti"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Clienti.asp"
dicitura.scrivi_con_sottosez()  
%>

<script language="JavaScript" type="text/javascript">

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
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="Hidden" name="old_login" value="<%= OBJ_contatto("login") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
	<% 	if request("ID") = "" then %>
		<caption>Inserimento nuovo cliente</caption>
	<% 	else %>
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del cliente</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="cliente precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="cliente successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
	<% 	end if
        if request("ID") = "" then %>
		<tr><th colspan="4">SELEZIONA IL CLIENTE DAI CONTATTI</th></tr>
		<tr>
			<td class="label">contatto:</td>
			<td class="content" colspan="3">
				<%
                sql = " (( ut_id NOT IN (SELECT riv_id FROM gtb_rivenditori) AND ut_id NOT IN (SELECT ag_id FROM gtb_agenti)) OR " + SQL_IsNULL(OBJ_contatto.conn, "ut_id") + ") " + _
                          " AND (" + SQL_IsNULL(OBJ_contatto.conn, "cntRel") + " OR CntRel=0) "
                CALL WriteContactPicker_Input(OBJ_contatto.conn, rs, sql, "", "form1", "tfn_IDElencoIndirizzi", request("IDCNT"), "LOGIN EMAIL", false, false, false, "REDIRECT")
                %>
			</td>
		</tr>
		<% 	else %>
		<input type="hidden" name="tfn_IDElencoIndirizzi" value="<%= request("ID") %>">
		<% 	end if %>
		<tr><th colspan="4">CLIENTE</th></tr>
		<tr>
			<td class="label">stato:</td>
			<td class="content">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder" type="radio" name="extC_riv_attivo" value="1" <%=chk(OBJ_contatto("riv_attivo") OR cString(OBJ_contatto("riv_attivo")) = "")%>></td>
						<td width="30%"><% WriteColor("#67c567") %>attivo</td>
						<td><input class="noBorder" type="radio" name="extC_riv_attivo" value="" <%=chk(not OBJ_contatto("riv_attivo"))%>></td>
						<td><% WriteColor("#f94d4d") %>non attivo</td>
					</tr>
				</table>
			</td>
			<td class="label">codice aziendale:</td>
			<td class="content">
				<input type="text" class="text" name="extT_riv_codice" value="<%= OBJ_contatto("riv_codice") %>" maxlength="20" size="15">
			</td>
		</tr>
		<tr><th colspan="4">DATI ANAGRAFICI</th></tr>
		<tr>
			<td class="label">salva come:</td>
			<td class="content">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder" type="radio" name="isSocieta" id="chk_issocieta_false" value="" <%=chk(not OBJ_contatto("isSocieta"))%> onClick="show_mandatory()"></td>
						<td width="30%">persona fisica</td>
						<td><input class="noBorder" type="radio" name="isSocieta" id="chk_issocieta_true" value="1" <%=chk(OBJ_contatto("isSocieta"))%> onClick="show_mandatory()"></td>
						<td>ente / societ&agrave; / organizzazione</td>
					</tr>
				</table>
			</td>
			<td class="label">lingua:</td>
			<td class="content">
				<% CALL DropLingue(OBJ_contatto.conn, rs, "tft_lingua", OBJ_contatto("lingua"), true, false, "width:100px;") %>
			</td>
		</tr>
		<tr>
			<td class="label" style="width:19%;">ente / societ&agrave;:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_NomeOrganizzazioneElencoIndirizzi" value="<%= OBJ_contatto("NomeOrganizzazioneElencoIndirizzi") %>" maxlength="250" size="100">
				<span id="ente">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= OBJ_contatto("CognomeElencoIndirizzi") %>" maxlength="100" size="75">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr> 
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= OBJ_contatto("NomeElencoIndirizzi") %>" maxlength="100" size="75">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			show_mandatory();
		</script>
		<tr>
			<td class="label">luogo di nascita:</td>
			<td class="content">
				<input type="text" class="text" name="tft_LuogoNascita" value="<%= OBJ_contatto("LuogoNascita") %>" maxlength="255" size="50">
			</td>
			<td class="label">data di nascita:</td>
			<td class="content">
				<input type="text" class="text" name="tft_DTNASCElencoIndirizzi" value="<%= OBJ_contatto("DTNASCElencoIndirizzi") %>" maxlength="10" size="10">
			</td>
		</tr>
		<tr>
			<td class="label">codice fiscale:</td>
			<td class="content">
				<input type="text" class="text" name="tft_CF" value="<%= OBJ_contatto("CF") %>" maxlength="16" size="16">
			</td>
			<td class="label">partita i.v.a.:</td>
			<td class="content"><input type="text" class="text" name="tft_partita_iva" value="<%= OBJ_contatto("partita_iva") %>" maxlength="11" size="14"></td>
		</tr>
		<tr><th colspan="4">INDIRIZZO PRINCIPALE</th></tr>
		<tr>
			<td class="label">indirizzo:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= OBJ_contatto("IndirizzoElencoIndirizzi") %>" maxlength="250" size="100"></td>
		</tr>
		<tr>
			<td class="label">localit&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_LocalitaElencoIndirizzi" value="<%= OBJ_contatto("LocalitaElencoIndirizzi") %>" maxlength="50" size="50"></td>
			<td class="label">cap:</td>
			<td class="content" ><input type="text" class="text" name="tft_CAPElencoIndirizzi" value="<%= OBJ_contatto("CAPElencoIndirizzi") %>" maxlength="20" size="10"></td>
		</tr>
		<tr>
			<td class="label">citt&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_cittaElencoIndirizzi" value="<%= OBJ_contatto("cittaElencoIndirizzi") %>" maxlength="50" size="50"></td>
			<td class="label">provincia / stato:</td>
			<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= OBJ_contatto("StatoProvElencoIndirizzi") %>" maxlength="50" size="30"></td>
		</tr>
		<tr><th colspan="4">RECAPITI</th></tr>
		<tr>
			<td class="label">telefono:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_telefono" value="<%= OBJ_contatto("telefono") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">fax:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_fax" value="<%= OBJ_contatto("fax") %>" maxlength="20" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">cellulare:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cellulare" value="<%= OBJ_contatto("cellulare") %>" maxlength="20" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">email:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_email" value="<%= OBJ_contatto("email") %>" maxlength="250" size="75">
				(*)
			</td>
		</tr>
		<% if request("ID") <> "" then
			sql = " ut_id IN (SELECT riv_id FROM gtb_rivenditori) "
		%>
			<tr>
				<th colspan="4">CAPOGRUPPO</th>
			</tr>
			<tr>
				<td class="label" style="width:19%;">azienda capogruppo:</td>
				<td class="content" colspan="3">
					<%
					CALL WriteContactPicker_Input(OBJ_contatto.conn, rs, sql, "", "form1", "extN_riv_azienda_capogruppo_id", OBJ_contatto("riv_azienda_capogruppo_id"), "LOGIN LOGINID EMAIL", false, false, false, "")
					%>
				</td>
			</tr>
		<% end if %>
		<% if request("ID") <> "" then 'gestione contatti interni
			sql = " SELECT *, (SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_ind_id = tb_indirizzario.IDElencoIndirizzi) AS N_ORD, " + _
				  " (SELECT ut_id FROM tb_utenti WHERE ut_nextCom_id = tb_indirizzario.IDElencoIndirizzi) AS UTENTE, " + _
				  " (SELECT COUNT(*) FROM rel_utenti_sito INNER JOIN tb_utenti ON rel_utenti_sito.rel_ut_id = tb_utenti.ut_id " + _
				  " WHERE tb_utenti.ut_nextCom_id = tb_indirizzario.IDElencoIndirizzi) AS N_PERMESSI " + _
				  " FROM tb_Indirizzario INNER JOIN tb_cnt_lingue ON tb_Indirizzario.lingua = tb_cnt_lingue.lingua_codice " + _
				  " WHERE CntRel=" & cIntero(request("ID")) & " ORDER BY ModoRegistra "
			rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText%>
			<tr>
				<th colspan="4">CONTATTI INTERNI / SEDI DIVERSE</th>
			</tr>
			<tr>
				<td colspan="4">
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label" style="width:30%">
								<% if rsr.eof then %>
									Nessun contatto / sede inseriti.
								<% else %>
									Trovati n&ordm; <%= rsr.recordcount %> record
								<% end if %>
							</td>
							<td colspan="6" class="content_right" style="padding-right:0px;">
								<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ClientiIntGestione.asp?CNT=<%= request("ID") %>', 'cntInt', 500, 450, true)">
									NUOVO CONTATTO / SEDE
								</a>
							</td>
						</tr>
						<% if not rsr.eof then %>
							<tr>
								<th class="L2">contatto / sede</th>
								<th class="L2" width="25%">indirizzo</th>
								<th class="l2_center" width="6%">sede</th>
								<th class="l2_center" width="7%">accesso</th>
								<th class="L2" width="3%">&nbsp;</th>
								<th class="l2_center" width="14%" colspan="2">operazioni</th>
							</tr>
							<% while not rsr.eof %>
								<tr>
									<td class="content" title="ruolo / qualifica: <%= rsr("QualificaElencoIndirizzi") %>"><%= ContactFullName(rsr) %></td>
									<td class="content">
										<%= ContactAddress(rsr) %>
										<% if cIntero(rsr("CntSede"))>0 then 
											if ContactAddress(rsr)<>"" then %><br><%end if
											sql = "SELECT * FROM tb_Indirizzario WHERE idElencoIndirizzi = " & rsr("CntSede")
											rsa.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
											if not rsa.eof then%>
												<span class="note">sede: <strong><%= ContactFullName(rsa) %></strong> - <%= ContactAddress(rsa) %></span>
											<% end if
											rsa.close
										end if %> 
									</td>
									<td class="content_center">
										<input type="checkbox" class="checkbox" disabled <%= chk(rsr("isSocieta")) %> title="<%= IIF(rsr("isSocieta"), "sede alternativa o periferica", "contatto interno") %>">
									</td>
									<td class="content_center">
										<input type="checkbox" class="checkbox" disabled <%= chk(cInteger(rsr("UTENTE"))>0) %> title="<%= IIF(cInteger(rsr("UTENTE"))>0, "con accesso subordinato", "senza accesso") %>">
									</td>
									<td class="content_center">
										<% if rsr("lingua")<>"" then %>
											<img src="../grafica/flag_mini_<%= rsr("lingua") %>.jpg" alt="Lingua: <%= rsr("lingua_nome_it") %>">
										<% end if %>
									</td>
									<td class="content_center">
										<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ClientiIntGestione.asp?ID=<%= rsr("IdElencoIndirizzi") %>', 'cntInt', 500, 450, true)">
											MODIFICA
										</a>
									</td>
									<td class="content_center">
										<% if cInteger(rsr("N_ORD"))>0 then %>
											<a class="button_L2_disabled" title="impossibile cancellare l'indirizzo altenativo perch&egrave; &egrave; associato ad almeno una riga d'ordine">
												CANCELLA
											</a>
										<% elseif cInteger(rsr("N_PERMESSI"))>1 then %>
											<a class="button_L2_disabled" title="impossibile cancellare il contatto associato all'indirizzo in quanto ha accesso anche ad altre applicazioni dell'area riservata.">
												CANCELLA
											</a>
										<% else %>
											<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('CLIENTI_INDIRIZZI','<%= rsr("IDElencoIndirizzi") %>');">
												CANCELLA
											</a>
										<% end if %>
									</td>
								</tr>
								<%rsr.movenext
							wend
						end if%>
					</table>
				</td>
			</tr>
			<% rsr.close 
		end if %>

	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<% if request("ID")="" then 
			'in inserimento permette anche l'inserimento di un nuovo listino
			%>
			<tr><th colspan="4">LISTINO ASSOCIATO</th></tr>
			<script type="text/javascript">
				function ImpostaStatoControlli(){
					var obj_listino_associato_esistente = document.getElementById("listino_associato_esistente");
					var obj_listino_associato_nuovo = document.getElementById("listino_associato_nuovo");
					
					//abilista e disabilita controlli di stato
					EnableIfChecked(obj_listino_associato_esistente, form1.extn_riv_listino_id);
					
					EnableIfChecked(obj_listino_associato_nuovo, form1.listino_codice);
					EnableIfChecked(obj_listino_associato_nuovo, document.all.copia_da_base);
					EnableIfChecked(obj_listino_associato_nuovo, document.all.copia_da_altro);
					EnableIfChecked(obj_listino_associato_nuovo, document.all.copia_da_ancestor);
					if (obj_listino_associato_esistente.checked){
						EnableIfChecked(obj_listino_associato_nuovo, form1.copia_da_altro_id);
						EnableIfChecked(obj_listino_associato_nuovo, form1.copia_da_ancestor_id);
					}else
					{
						EnableIfChecked(document.all.copia_da_altro, form1.copia_da_altro_id);
						EnableIfChecked(document.all.copia_da_ancestor, form1.copia_da_ancestor_id);
					}
				}
			</script>
			<tr>
				<td class="label" rowspan="5">listino associato:</td>
				<td class="content" width="18%">
					<input class="checkbox" type="radio" name="listino_associato" id="listino_associato_esistente" value="" <%= chk(request("listino_associato")="") %> onclick="ImpostaStatoControlli()">
					listino <strong>esistente</strong>:
				</td>
				<td class="content" colspan="2">
					<% 	sql = " SELECT gtb_listini.listino_id, (CASE " + _
							  "	WHEN gtb_listini.listino_base_attuale=1 THEN 'listino base in vigore (mantenuto automaticamente)' " + _
							  " WHEN ISNULL(gtb_listini.listino_ancestor_id,0)>0 THEN gtb_listini.listino_codice + ' (derivato da ' + L_ancestor.listino_codice + ')'" + _
							  " ELSE gtb_listini.listino_codice END) AS denominazione " + _
						 	  " FROM gtb_listini LEFT JOIN gtb_listini L_ancestor ON gtb_listini.listino_ancestor_id = L_ancestor.listino_id " + _
							  " WHERE IsNull(gtb_listini.listino_offerte,0)=0 OR gtb_listini.listino_base_attuale=1 ORDER BY gtb_listini.listino_codice"
					dropDown conn, sql, "listino_id", "denominazione", "extn_riv_listino_id", OBJ_contatto("riv_listino_id"), true, " ID=""extn_riv_listino_id_select""", LINGUA_ITALIANO %>
				</td>
			</tr>
			<tr>
				<td class="content" rowspan="4">
					<input class="checkbox" type="radio" name="listino_associato" id="listino_associato_nuovo" value="nuovo" <%= chk(request("listino_associato")="nuovo") %> onclick="ImpostaStatoControlli()">
					listino <strong>nuovo</strong>:
				</td>
				<td class="label" style="width:9%;">nome</td>
				<td class="content">
					<input type="text" class="text" name="listino_codice" value="<%= request("listino_codice") %>" maxlength="50" size="60" onclick="ImpostaStatoControlli()"><br>
					<span class="note">
						Se lasciato vuoto viene generato automaticamente dal nome cliente.
					</span>
				</td>
			</tr>
			<tr>
				<td class="label" style="width:9%;" rowspan="3">modalit&agrave;</td>
				<td class="content">
					<input class="checkbox" type="radio" name="copia_da" id="copia_da_base" value="" <%= chk(request("copia_da")="") %> onclick="ImpostaStatoControlli()">
					<strong>copia</strong> prezzi da listino base in vigore
				</td>
			</tr>
			<tr>
				<td class="content">
					<input class="checkbox" type="radio" name="copia_da" id="copia_da_altro" value="altro" <%= chk(request("copia_da")="altro") %> onclick="ImpostaStatoControlli()">
					<strong>copia</strong> prezzi da altro listino:<br>
					<span style="padding-left:22px;">
						<% sql = " SELECT listino_id, listino_codice FROM gtb_listini " + _
								 " WHERE listino_offerte=0 AND listino_base=0 ORDER BY listino_codice"
						CALL dropDown(conn, sql, "listino_id", "listino_codice", "copia_da_altro_id", request("copia_da_altro_id"), true, "", LINGUA_ITALIANO)%>
					</span>
				</td>
			</tr>
			<tr>
				<td class="content">
					<input class="checkbox" type="radio" name="copia_da" id="copia_da_ancestor" value="ancestor" <%= chk(request("copia_da")="ancestor") %> onclick="ImpostaStatoControlli()">
					<strong>deriva</strong> prezzi da altro listino:<br>
					<span style="padding-left:22px;">
						<% sql = " SELECT listino_id, listino_codice FROM gtb_listini " + _
								 " WHERE listino_offerte=0 AND listino_base=0 AND IsNull(listino_ancestor_id,0)=0 ORDER BY listino_codice"
						CALL dropDown(conn, sql, "listino_id", "listino_codice", "copia_da_ancestor_id", request("copia_da_ancestor_id"), true, "", LINGUA_ITALIANO)%>
					</span>
				</td>
			</tr>
			<script type="text/javascript">ImpostaStatoControlli();</script>
			<tr><th colspan="4">ALTRE INFORMAZIONI COMMERCIALI</th></tr>
		<% else
			'in modifica permette solo di variare il listino associato al cliente
			%>
			<tr><th colspan="4">PROFILO COMMERCIALE</th></tr>
			<tr>
				<td class="label">listino associato:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT gtb_listini.listino_id, (CASE " + _
							  "	WHEN gtb_listini.listino_base_attuale=1 THEN 'listino base in vigore (mantenuto automaticamente)' " + _
							  " WHEN ISNULL(gtb_listini.listino_ancestor_id,0)>0 THEN gtb_listini.listino_codice + ' (derivato da ' + L_ancestor.listino_codice + ')'" + _
							  " ELSE gtb_listini.listino_codice END) AS denominazione " + _
						 	  " FROM gtb_listini LEFT JOIN gtb_listini L_ancestor ON gtb_listini.listino_ancestor_id = L_ancestor.listino_id " + _
							  " WHERE IsNull(gtb_listini.listino_offerte,0)=0 OR gtb_listini.listino_base_attuale=1 ORDER BY gtb_listini.listino_codice"
					dropDown conn, sql, "listino_id", "denominazione", "extN_riv_listino_id", OBJ_contatto("riv_listino_id"), true, "", LINGUA_ITALIANO %>
					(*)
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">sconto su ordine:</td>
			<td class="content" colspan="3">
				<input type="text" class="number" name="extN_riv_sconto_ordine" value="<%= OBJ_contatto("riv_sconto_ordine") %>" maxlength="3" size="5"> %
			</td>
		</tr>
		<tr>
			<td class="label" style="width:19%;">agente di zona:</td>
			<td class="content" colspan="3">
				<%
                sql = " ut_id IN (SELECT ag_id FROM gtb_agenti) "
                CALL WriteContactPicker_Input(OBJ_contatto.conn, rs, sql, "", "form1", "extN_riv_agente_id", OBJ_contatto("riv_agente_id"), "LOGIN LOGINID EMAIL", false, false, false, "")
                %>
			</td>
		</tr>
		<tr>
			<td class="label">valuta:</td>
			<td class="content" colspan="3">
				<% 	dropDown conn, "SELECT * FROM gtb_valute ORDER BY valu_nome", "valu_id", "valu_nome", "extN_riv_valuta_id", _
							 IIF(OBJ_contatto("riv_valuta_id")<>"", OBJ_contatto("riv_valuta_id"), GetValueList(conn, rsr, "SELECT TOP 1 valu_id FROM gtb_valute WHERE valu_nome LIKE '%eur%'")), _
							 true, "", LINGUA_ITALIANO %>
			</td>
		</tr>
		<tr>
			<td class="label">lista codici personalizzati:</td>
			<td class="content" colspan="3">
				<% 	dropDown conn, "SELECT * FROM gtb_lista_codici WHERE NOT "& SQL_isTrue(conn, "lstCod_sistema") &" ORDER BY lstCod_Nome", "lstcod_id", "lstcod_nome", "extN_riv_lstcod_id", OBJ_contatto("riv_lstcod_id"), false, "", LINGUA_ITALIANO %>
			</td>
		</tr>
		<tr>
			<td class="label">modalit&agrave; di pagamento:</td>
			<td class="content" colspan="3">
				<% 	dropDown conn, "SELECT * FROM gtb_modipagamento where "& SQL_isTrue(conn, "mosp_se_abilitato") &" ORDER BY mosp_nome_it", "mosp_id", "mosp_nome_it", "extN_riv_modopagamento_id", OBJ_contatto("riv_modopagamento_id"), false, "", LINGUA_ITALIANO %>
			</td>
		</tr>
		
		<tr><th colspan="4">DATI CONSEGNA MERCE</th></tr>
		<% if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_spese_spedizione"))>0 then %>
			<tr>
				<td class="label">tipologia spese di spedizione:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT * FROM gtb_spese_spedizione ORDER BY sp_area_nome_it"
					dropDown conn, sql, "sp_id", "sp_area_nome_it", "extN_riv_spese_spedizione_id", OBJ_contatto("riv_spese_spedizione_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
		<% if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_porti"))>0 then %>
			<tr>
				<td class="label">porto di default:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT prt_id, (prt_nome_it + (CASE WHEN IsNull(prt_attivo,0)=0 THEN ' (non attivo)' ELSE '' END)) AS prt_nome FROM gtb_porti ORDER BY prt_nome_it"
					dropDown conn, sql, "prt_id", "prt_nome", "extN_riv_porto_default_id", OBJ_contatto("riv_porto_default_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
		<% if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_trasportatori"))>0 then %>
			<tr>
				<td class="label">trasportatore di default:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT * FROM gtb_trasportatori ORDER BY tra_nome_it"
					dropDown conn, sql, "tra_id", "tra_nome_it", "extN_riv_trasportatore_default_id", OBJ_contatto("riv_trasportatore_default_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
	</table>
	
	<%
	'......................................................................................................
	'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
	CALL ADDON__CLIENTI__form_gestione(OBJ_contatto, rs)
	'......................................................................................................
	%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="4">PROFILO DI ACCESSO</th></tr>
		<tr>
			<td class="label" style="width:19%;">stato:</td>
			<td class="content" colspan="1">
				<input type="checkbox" class="checkbox" name="chk_abilitato" <%= Chk(IIF(request("ID")<>"" OR request.serverVariables("REQUEST_METHOD")="POST", OBJ_contatto("abilitato"), true)) %>>
				abilitato all'accesso
			</td>
			<% Session("B2B_OLD_VALUE_CHK_ABILITATO") = IIF(request("ID")<>"" OR request.serverVariables("REQUEST_METHOD")="POST", OBJ_contatto("abilitato"), true) %>
			<td class="label" style="text-align:right;"colspan="2">
				<% if cIntero(Session("B2B_ID_PAG_SPEDIZ_CREDENZ_ACCESSO"))>0 AND cBool(OBJ_contatto("abilitato")) then %>
					<a class="button_L2" style="text-align:center; width:190px;" href="<%= GetPageSiteUrl(conn, Session("B2B_ID_PAG_SPEDIZ_CREDENZ_ACCESSO"), OBJ_contatto("lingua"))&"&RIV_ID="&OBJ_contatto("riv_id")&"&ID_ADMIN="&Session("ID_ADMIN")&"&IDCNT="&OBJ_contatto("IDElencoIndirizzi")&"&KEY="&OBJ_contatto("codiceInserimento")&"&HTML_FOR_EMAIL=1"%>" target="ID_RIVENDITORE_<%=OBJ_contatto("riv_id")%>"
						onclick="OpenAutoPositionedScrollWindow('', 'ID_RIVENDITORE_<%=OBJ_contatto("riv_id")%>', 800, 800, true)" title="Click per visualizzare la lettera di vettura" <%= ACTIVE_STATUS %>>
						SPEDISCI CREDENZIALI DI ACCESSO
					</a>
				<% end if %>
			</td>
		</tr>
		<% sql = "SELECT * FROM gtb_profili ORDER BY pro_nome_it" %>
		<% if cString(GetValueList(conn,NULL,sql))<>"" then %>
			<tr>
				<td class="label">profilo:</td>
				<td class="content" colspan="3">
					<% CALL	dropDown(conn, sql, "pro_id", "pro_nome_it", "extn_riv_profilo_id", OBJ_contatto("riv_profilo_id"), false, "", LINGUA_ITALIANO) %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">Login:</td>
			<td class="content">
				<input type="text" class="text" name="tft_login" value="<%= OBJ_contatto("login") %>" maxlength="50" size="25">
				(*)
			</td>
			<td class="label_right" style="width:24%;">Scadenza:</td>
			<td class="content_right">
				<% CALL WriteDataPicker_Input("form1", "tfd_scadenza", OBJ_contatto("Scadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">Password:</td>
			<td class="content">
				<input type="password" class="text" name="tft_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="25">
				(*)
			</td>
			<td class="note" colspan="2" rowspan="2" style="width:55%;">
				Per i valori di login e password utilizzare solo caratteri alfanumerici o &quot;_&quot; 
				indifferentemente con lettere minuscole o maiuscole, ma senza spazi bianchi.
				<span style="letter-spacing:2px;">(<%= LOGIN_VALID_CHARSET %>)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Conferma password:</td>
			<td class="content">
				<input type="password" class="text" name="conferma_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="25">
				(*)
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<% if request("ID")="" then %>
					<input type="submit" class="button" name="salva" value="SALVA &gt;&gt">
				<% else %>
					<input type="submit" class="button" name="salva_modifica" value="SALVA">
					<input type="submit" class="button" name="salva" value="SALVA &amp; TORNA AD ELENCO">
				<% end if %>
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% if Session("B2B_HTTP_RESULT_SPEDIZ_CREDENZ_ACCESSO") <> "" then %>
	<script type="text/javascript">
		alert('Le credenziali di accesso sono state spedite.');
	</script>
	<% Session("B2B_HTTP_RESULT_SPEDIZ_CREDENZ_ACCESSO") = "" %>
<% end if %>
<% conn.close
set rs = nothing
set rsa = nothing
set rsr = nothing
set conn = nothing
%>