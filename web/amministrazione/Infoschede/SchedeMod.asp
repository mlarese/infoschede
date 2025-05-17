<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
dim post, is_officina, trasportatoreId

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	post = true
end if

if cString(Session("INFOSCHEDE_OFFICINA"))<>"" then
	is_officina = true
else
	is_officina = false
end if

trasportatoreId = 0

if (request("salva")<>"" OR request("salva_continua")<>"") AND post then
	Server.Execute("SchedeSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<SCRIPT LANGUAGE="javascript"  src="../library/utils4dynalay.js" type="text/javascript"></SCRIPT>

<% 	

dim conn, rs, rsa, rsd, sql, i, ordine, assegnata, cat_riconsegna, prof_id
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")


sql = " SELECT sgtb_schede.*, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.codiceInserimento, sts_elenco_ddt_da_consegnare, sts_elenco_ddt_da_ritirare " & _
	  " FROM sgtb_schede INNER JOIN gtb_rivenditori ON sgtb_schede.sc_cliente_id = gtb_rivenditori.riv_id " & _
	  " INNER JOIN tb_Utenti ON gtb_rivenditori.riv_id = tb_Utenti.ut_ID " & _
	  " INNER JOIN tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi " & _
	  " LEFT JOIN sgtb_stati_schede ON sgtb_schede.sc_stato_id = sgtb_stati_schede.sts_id " & _
	  " WHERE sc_id = " & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

if cIntero(rs("sc_centro_assistenza_id")) > 0 then
	assegnata = true
else
	assegnata = false
end if


dim dicitura, data
set dicitura = New testata 
dicitura.iniz_sottosez(0)
if assegnata then
	dicitura.sezione = "Gestione schede di assistenza - modifica"
else
	dicitura.sezione = "Gestione richieste di assistenza - modifica"
end if
dicitura.puls_new = "INDIETRO"
if assegnata then
	dicitura.link_new = "Schede.asp?ASSEGNATA=true"
else
	dicitura.link_new = "Schede.asp?ASSEGNATA=false"
end if
dicitura.scrivi_con_sottosez() 


if request("goto")<>"" then
	CALL GotoRecord(conn, rsa, session("INFOSCHEDE_SCHEDE_SQL"), "sc_id", "SchedeMod.asp")
end if

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati della scheda</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="scheda precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="scheda successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="6">DATI PRINCIPALI</th></tr>
		<% if cIntero(rs("sc_external_id")) > 0 then %>
			<tr>
				<td class="header OrdConfermato" colspan="2" style="font-weight:normal !important; font-size:10px;">
					scheda importata da db Access
				</td>
				<td class="header OrdConfermato" align="right" colspan="2" style="font-size:10px;">
					ext. id: &nbsp; <%=cIntero(rs("sc_external_id"))%>
				</td>
			</tr>
		<% end if %>
		
		<% sql = "SELECT * FROM sgtb_stati_schede ORDER BY sts_ordine, sts_nome_it" %>
		<tr>
			<td class="label">stato scheda:</td>
			<td class="content" colspan="3">
				<% if is_officina then %><input type="hidden" name="tfn_sc_stato_id" value="<%=rs("sc_stato_id")%>"><% end if %>
				<% CALL dropDown(conn, sql, "sts_id", "sts_nome_it", "tfn_sc_stato_id", IIF(post,request("tfn_sc_stato_id"),rs("sc_stato_id")), true, IIF(is_officina,"disabled",""), Session("LINGUA")) %>
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" style="width:24%;">numero:</td>
			<td class="content_b" colspan="3">
				<%=rs("sc_numero")%>
			</td>
		</tr>
		<tr>
			<td class="label">data ricevimento:</td>
			<td class="content_b" colspan="3">
				<%= DateIta(rs("sc_data_ricevimento")) %>
			</td>
		</tr>
		
		<% if assegnata then %>
			<tr>
				<td class="label">centro assistenza:</td>
				<td class="content_b" colspan="2" style="width:65%;">
					<% sql = " SELECT * FROM gtb_agenti INNER JOIN tb_Utenti ON gtb_agenti.ag_id = tb_Utenti.ut_ID INNER JOIN " & _
							 "               tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi" & _
							 " WHERE ag_id = " & cIntero(rs("sc_centro_assistenza_id"))
					rsa.open sql, conn
					response.write ContactFullName(rsa) 
					rsa.close %>
				</td>
				<td class="content_right">
					<% if Session("INFOSCHEDE_ADMIN")<>"" AND cString(request("ID_CENTRO_ASSISTENZA"))="" then %>
						<a href="javascript:void(0)" class="button_L2"
							onclick="OpenAutoPositionedScrollWindow('SchedeAssegnaCentroAssistenza.asp?ID_SCHEDA=<%= rs("sc_id") %>&ID_CENTRO=<%=cIntero(rs("sc_centro_assistenza_id"))%>', 'SelezioneCentroAssistenza', 450, 480, true)" 
							title="Click per aprire la finestra per la selezione del centro assistenza">
							CAMBIA CENTRO ASSISTENZA
						</a>
					<% else %>
						&nbsp;
					<% end if %>
				</td>
				</td>
			</tr>
		<% else %>
			<tr>
				<td class="label">centro assistenza:</td>
				<td class="content_right" colspan="3">
					<a href="javascript:void(0)" class="button_L2"
						onclick="OpenAutoPositionedScrollWindow('SchedeAssegnaCentroAssistenza.asp?ID_SCHEDA=<%= rs("sc_id") %>', 'SelezioneCentroAssistenza', 450, 480, true)" 
						title="Click per aprire la finestra per la selezione del centro assistenza">
						ASSEGNA A CENTRO ASSISTENZA
					</a>
				</td>
			</tr>					
		<% end if %>
		
		<% if is_officina then %>
			<tr>
				<td class="label" rowspan="2">cliente:</td>
				<td class="content_b" colspan="3">
					<input type="hidden" name="tfn_sc_cliente_id" value="<%=rs("sc_cliente_id")%>">
					<% sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & rs("sc_cliente_id")
					rsa.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
					response.write ContactFullName(rsa) 
					%>
				</td>
			</tr>
		<% else %>
			<tr>
				<td class="label" rowspan="2">cliente:</td>
				<td class="content" colspan="3">
					<table cellpadding="0" cellspacing="0" width="100%">
						<tr>
							<td>
								<input type="hidden" name="tfn_sc_cliente_id" value="<%=IIF(post, cIntero(request("tfn_sc_cliente_id")), rs("sc_cliente_id"))%>">
								<% sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & IIF(post, cIntero(request("tfn_sc_cliente_id")), rs("sc_cliente_id"))
								rsa.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
								%>
								<input READONLY type="text" name="cliente" style="padding-left:3px; width:100%" value="<%=IIF(post, request.form("cliente"), ContactFullName(rsa)) %>" 
									   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=cliente&field_id=tfn_sc_cliente_id&selected=' + tfn_sc_cliente_id.value + '&filtro_profilo=<%=TRASPORTATORI&","&COSTRUTTORI%>&filtro_exclude=true&BUTTONS_ADD=true&CENTRO_ASSISTENZA_ID=<%=request("ID_CENTRO_ASSISTENZA")%>&AFTER=submit', 'SelezioneCliente', 620, 480, true)" 
											title="Click per aprire la finestra per la selezione del cliente">
							</td>
							<td width="31%">
								<a class="button_input" href="javascript:void(0)" onclick="form1.cliente.onclick();" 
									 title="Apre la finestra per la selezione del cliente" <%= ACTIVE_STATUS %> style="display:inline;">
									SCEGLI
								</a>
								<a class="button_input" onclick="OpenAutoPositionedScrollWindow('ClientiGestione.asp?ID=<%=rsa("IDElencoIndirizzi")%>&PROFILO=anagrafiche_clienti&STANDALONE=true', 'DatiCliente', 500, 500, true)" 
									 href="javascript:void(0)"  title="Apre la finestra per la visualizzazione o la modifica dei dati del cliente" <%= ACTIVE_STATUS %>
									 style="display:inline;">
									VISUALIZZA DATI
								</a>
								&nbsp;(*)
							</td>
						</tr>
					</table>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td colspan="3">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr><th class="L2" colspan="2">DATI ANAGRAFICA</th></tr>
					<tr>
						<td class="label">indirizzo:</td>
						<td class="content"><%=ContactAddress(rsa)%></td>
					</tr>
					<tr>
						<td class="label">codice fiscale:</td>
						<td class="content"><%=rsa("CF")%></td>
					</tr>
					<tr>
						<td class="label">p. iva:</td>
						<td class="content"><%=rsa("partita_iva")%></td>
					</tr>
					<% dim Obj_Cnt 
					set Obj_Cnt = new IndirizzarioLock 
					Obj_Cnt.LoadFromDB(rsa("IDElencoIndirizzi"))
					%>
					<tr>
						<td class="label">telefono:</td>
						<td class="content"><%=Obj_Cnt("telefono")%></td>
					</tr>
					<tr>
						<td class="label">fax:</td>
						<td class="content"><%=Obj_Cnt("fax")%></td>
					</tr>
					<tr>
						<td class="label">cellulare:</td>
						<td class="content"><%=Obj_Cnt("cellulare")%></td>
					</tr>
					<tr>
						<td class="label">e-mail:</td>
						<td class="content"><a href="mailto:<%= Obj_Cnt("email") %>"><%= Obj_Cnt("email") %></a></td>
					</tr>
				</table>
			</td>
		</tr>
		<% rsa.close %>
		
		<% if not is_officina then %>
			<tr>
				<td class="label">note del cliente:</td>
				<td class="content" colspan="3">
					<textarea style="width:100%;" rows="5" name="tft_sc_note_cliente"><%=IIF(post,request("tft_sc_note_cliente"),rs("sc_note_cliente")) %></textarea>
				</td>
			</tr>
			<tr>
				<td class="label">riferimento cliente:</td>
				<td class="content" colspan="3">
					<input type="text" class="text" name="tft_sc_rif_cliente" value="<%=IIF(post,request("tft_sc_rif_cliente"),rs("sc_rif_cliente")) %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
		<% end if %>
		<tr><th colspan="6">DATI DEL MODELLO</th></tr>
		<tr>
			<td class="label" <%if cString(rs("sc_modello_altro"))<>"" OR cIntero(request("tfn_sc_modello_id")) = MODELLO_DEFAULT then%>rowspan="2"<%end if%>>modello:</td>
			<td class="content" colspan="3" nowrap>
				<% CALL WritePicker_ArticoloVariante(conn, rsa, "form1", "tfn_sc_modello_id", IIF(post,request("tfn_sc_modello_id"),rs("sc_modello_id")), 89, true, "Infoschede/ArticoliSeleziona.asp?SUBMIT_AFTER=true&TYPE=M&") %>
			</td>
		</tr>
		<% if cString(rs("sc_modello_altro"))<>"" OR cIntero(request("tfn_sc_modello_id")) = MODELLO_DEFAULT then %>
			<tr>
				<td class="content" colspan="5">
					nome modello:&nbsp;
					<input type="text" class="text" name="tft_sc_modello_altro" value="<%= rs("sc_modello_altro")  %>" maxlength="500" size="91">
				</td>
			</tr>
		<% end if %>
		<% sql = "SELECT rel_cod_int FROM grel_art_valori WHERE rel_id = " & cIntero(IIF(post,request("tfn_sc_modello_id"),rs("sc_modello_id"))) 
		dim codice
		codice = cString(GetValueList(conn, NULL, sql))
		%>
		<tr>
			<td class="label">codice:</td>
			<td class="content" colspan="3">
				<input type="text" READONLY class="text disabled" name="sc_codice" value="<%= codice %>" maxlength="255" size="52">
			</td>
		</tr>
		<tr>
			<td class="label">matricola:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_sc_matricola" value="<%=IIF(post,request("tft_sc_matricola"),rs("sc_matricola")) %>" maxlength="255" size="52">
			</td>
		</tr>
		
		<tr><th colspan="5">DATI DI ACQUISTO</th></tr>
		<% if not is_officina then %>
			<tr>
				<td class="label">negozio di acquisto:</td>
				<td class="content" colspan="3">
					<input type="text" class="text" name="tft_sc_negozio_acquisto" value="<%=IIF(post,request("tft_sc_negozio_acquisto"),rs("sc_negozio_acquisto")) %>" maxlength="255" size="52">
				</td>
			</tr>
			<tr>
				<td class="label">data acquisto:</td>
				<td class="content" colspan="3">
					<% CALL WriteDataPicker_Input_Manuale("form1", "tfd_sc_data_acquisto", IIF(post,request("tfd_sc_data_acquisto"),rs("sc_data_acquisto")), "", "/", IIF(is_officina,false,true), true, LINGUA_ITALIANO, "", true, "") 
					%>
					
					<% if is_officina then %>
						<script language="JavaScript" type="text/javascript">
							//document.getElementById("tfd_sc_data_acquisto").disabled = true;
							//document.getElementById("form1_link_scegli_tfd_sc_data_acquisto").onclick = '';
						</script>
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="label">numero scontrino:</td>
				<td class="content" colspan="3">
					<input type="text" class="text" name="tft_sc_numero_scontrino" value="<%=IIF(post,request("tft_sc_numero_scontrino"),rs("sc_numero_scontrino")) %>" maxlength="100" size="52">
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" style="line-height:20px;">richiesta garanzia:</td>
			<td class="content" colspan="3">
				<input type="hidden" name="chk_sc_richiesta_garanzia" value="<%= IIF(cBoolean(rs("sc_richiesta_garanzia"),false),"1","")%>">
				<input type="checkbox" class="noBorder" name="chk_sc_richiesta_garanzia" disabled value="1" <%= chk(cBoolean(rs("sc_richiesta_garanzia"),false))%>>
			</td>
		</tr>
		<% if cString(Session("INFOSCHEDE_ADMIN")) <> "" then %>
			<script language="JavaScript" type="text/javascript">
				function SetGaranzia(in_garanzia){
					if (in_garanzia){
						document.form1.chk_sc_in_garanzia.value = 1;
						document.getElementById("garanzia_si").className = 'button_disabled'
						document.getElementById("garanzia_no").className = 'button_selected'
					}
					else{
						document.form1.chk_sc_in_garanzia.value = '';
						document.getElementById("garanzia_si").className = 'button_selected'
						document.getElementById("garanzia_no").className = 'button_disabled'
					}
					document.form1.tfd_sc_data_garanzia_controllata.value = '<%=Now()%>';
				}
			</script>
			<tr>
				<td class="label" style="line-height:24px;">riparazione in garanzia:</td>
				<td class="content" colspan="3">
					<a class="button" id="garanzia_si" style="height:20px;" href="javascript:void(0)" onclick="SetGaranzia(true)">
						SI
					</a>
					&nbsp;
					<a class="button" id="garanzia_no" style="height:20px;" href="javascript:void(0)" onclick="SetGaranzia(false)">
						NO
					</a>
				</td>
				<input type="hidden" name="chk_sc_in_garanzia" value="">
				<input type="hidden" name="tfd_sc_data_garanzia_controllata" value="">
			</tr>
			<script language="JavaScript" type="text/javascript">
				<% if cBoolean(rs("sc_in_garanzia"), false) then %>
					SetGaranzia(true);
				<% else %>
					SetGaranzia(false);
				<% end if %>
			</script>
		<% else %>
			<tr>
				<td class="label" style="line-height:20px;">riparazione in garanzia:</td>
				<td class="note" colspan="3">
					<% if cString(rs("sc_data_garanzia_controllata"))="" AND cBoolean(rs("sc_richiesta_garanzia"),false) then %>
						in attesa di approvazione
					<% else %>
						<input type="hidden" name="chk_sc_in_garanzia" value="1" <%= chk(cBoolean(rs("sc_in_garanzia"),false))%>>
						<input type="checkbox" class="noBorder" name="chk_sc_richiesta_garanzia" disabled <%= chk(cBoolean(rs("sc_in_garanzia"),false))%>>
					<% end if %>
				</td>
			</tr>
		<% end if %>
		
		<% sql = "SELECT * FROM sgtb_accessori ORDER BY acc_nome_it" %>
		<tr>
			<td class="label" rowspan="2">accessori presenti:</td>
			<td class="content" colspan="3">
				<% if is_officina then %><input type="hidden" name="chk_sc_in_garanzia" value="<%=rs("sc_accessori_presenti_id")%>"><% end if %>
				<% CALL dropDown(conn, sql, "acc_id", "acc_nome_it", "tfn_sc_accessori_presenti_id", IIF(post,request("tfn_sc_accessori_presenti_id"),rs("sc_accessori_presenti_id")), false, "style=""width:50%;""", Session("LINGUA")) %>
			</td>
		</tr>
		<tr>
			<td class="content" colspan="3">
				altro&nbsp;
				<input type="text" class="text" name="tft_sc_accessori_presenti_altro" value="<%=IIF(post,request("tft_sc_accessori_presenti_altro"),rs("sc_accessori_presenti_altro")) %>" maxlength="500" style="width:100%;">
			</td>
		</tr>
		
		<tr><th colspan="6">DATI DELLA RIPARAZIONE</th></tr>
		<% 'sql = "SELECT * FROM sgtb_problemi WHERE ISNULL(prb_riscontrato, 0)=0 ORDER BY prb_nome_it" 
		dim marca_id, tipologia_id, rel_art_id
		rel_art_id = IIF(post,request("tfn_sc_modello_id"),rs("sc_modello_id"))
		marca_id = CIntero(GetValueList(conn, NULL, "SELECT art_marca_id FROM gv_articoli WHERE rel_id = "&rel_art_id))
		tipologia_id = CIntero(GetValueList(conn, NULL, "SELECT art_tipologia_id FROM gv_articoli WHERE rel_id = "&rel_art_id))
		sql = "SELECT DISTINCT prb_id, prb_nome_it FROM sgtb_problemi " & _
			 " LEFT JOIN srel_problemi_articoli ON prb_id = rpa_problema_id" & _
			 " LEFT JOIN srel_problemi_mar_tip ON prb_id = rpm_problema_id" & _
			 " LEFT JOIN grel_art_valori ON rel_id = rpa_articolo_rel_id" & _
				 " WHERE prb_id = "&cIntero(rs("sc_guasto_segnalato_id")) & " OR " & _
				   " (ISNULL(prb_riscontrato, 0)=0 " & _
				   " AND prb_visibile = 1" & _
				   " AND ((rpa_problema_id = prb_id AND rel_art_id =" & rel_art_id & ")" & _
					 " OR (rpm_problema_id = prb_id" & _
				   " AND (rpm_tipologia_id = " & tipologia_id &" OR rpm_tipologia_id = 0)" & _
				   " AND (rpm_marchio_id = " & marca_id & " OR rpm_marchio_id = 0))))"
		%>
		<tr>
			<td class="label" rowspan="2">guasto segnalato:</td>
			<td class="content" colspan="3">
				<% if is_officina then %><input type="hidden" name="tfn_sc_guasto_segnalato_id" value="<%=rs("sc_guasto_segnalato_id")%>"><% end if %>
				<% CALL dropDown(conn, sql, "prb_id", "prb_nome_it", "tfn_sc_guasto_segnalato_id", IIF(post,request("tfn_sc_guasto_segnalato_id"),rs("sc_guasto_segnalato_id")), false, "style=""width:50%;"""&IIF(is_officina," disabled",""), Session("LINGUA")) %>
			</td>
		</tr>
		<tr>
			<td class="content" colspan="3">
				altro&nbsp;
				<input type="text" class="text" name="tft_sc_guasto_segnalato_altro" value="<%=IIF(post,request("tft_sc_guasto_segnalato_altro"),rs("sc_guasto_segnalato_altro"))%>" maxlength="500" style="width:100%;" <%=IIF(is_officina,"disabled","")%>>
			</td>
		</tr>
		
		<% if assegnata then %>
			<% sql = "SELECT * FROM sgtb_problemi WHERE ISNULL(prb_riscontrato, 0)=1 ORDER BY prb_nome_it" 
			' sql = "SELECT DISTINCT prb_id, prb_nome_it FROM sgtb_problemi " & _
			 ' " LEFT JOIN srel_problemi_articoli ON prb_id = rpa_problema_id" & _
			 ' " LEFT JOIN srel_problemi_mar_tip ON prb_id = rpm_problema_id" & _
			 ' " LEFT JOIN grel_art_valori ON rel_id = rpa_articolo_rel_id" & _
				 ' " WHERE prb_id = "&cIntero(rs("sc_guasto_riscontrato_id")) & " OR " & _
				   ' " (ISNULL(prb_riscontrato, 0)=1 " & _
				   ' " AND prb_visibile = 1" & _
				   ' " AND ((rpa_problema_id = prb_id AND rel_art_id =" & rel_art_id & ")" & _
					 ' " OR (rpm_problema_id = prb_id" & _
				   ' " AND (rpm_tipologia_id = " & tipologia_id &" OR rpm_tipologia_id = 0)" & _
				   ' " AND (rpm_marchio_id = " & marca_id & " OR rpm_marchio_id = 0))))"
			
			%>
			<tr>
				<td class="label" rowspan="2">guasto riscontrato:</td>
				<td class="content" colspan="3">
					<% CALL dropDown(conn, sql, "prb_id", "prb_nome_it", "tfn_sc_guasto_riscontrato_id", IIF(post,request("tfn_sc_guasto_riscontrato_id"),rs("sc_guasto_riscontrato_id")), false, "style=""width:50%;""", Session("LINGUA")) %>
				</td>
			</tr>
			<tr>
				<td class="content" colspan="3">
					altro&nbsp;
					<input type="text" class="text" name="tft_sc_guasto_riscontrato_altro" value="<%= IIF(post,request("tft_sc_guasto_riscontrato_altro"),rs("sc_guasto_riscontrato_altro")) %>" maxlength="500" style="width:100%;">
				</td>
			</tr>
			
			<% sql = "SELECT * FROM sgtb_esiti ORDER BY esi_nome_it" %>
			<tr>
				<td class="label">esito dell'intervento:</td>
				<td class="content" colspan="3">
					<% CALL dropDown(conn, sql, "esi_id", "esi_nome_it", "tfn_sc_esito_intervento_id", IIF(post,request("tfn_sc_esito_intervento_id"),rs("sc_esito_intervento_id")), false, "style=""width:50%;""", Session("LINGUA")) %>
				</td>
			</tr>
			
			<tr>
				<td class="label">data fine lavoro:</td>
				<td class="content" colspan="3">
					<% CALL WriteDataPicker_Input_Manuale("form1", "tfd_sc_data_fine_lavoro", IIF(post,request("tfd_sc_data_fine_lavoro"),rs("sc_data_fine_lavoro")), "", "/", true, true, LINGUA_ITALIANO, "", true, "") 
					%>
				</td>
			</tr>
			<tr>
				<td class="label">ore manodopera:</td>
				<td class="content">
					<input type="text" class="number" name="tfn_sc_ora_manodopera_intervento" value="<%=IIF(post,request("tfn_sc_ora_manodopera_intervento"),rs("sc_ora_manodopera_intervento")) %>" size="7" onchange="FunzPrezzoManodopera()">
				</td>
				<td class="label_no_width" style="width:25%;">prezzo orario manodopera:</td>
				<td class="content">
					<input type="text" class="number" name="tfn_sc_prezzo_manodopera" value="<%= FormatPrice(cReal(IIF(post,request("tfn_sc_prezzo_manodopera"),rs("sc_prezzo_manodopera"))), 2, false) %>" size="7" onchange="FunzPrezzoManodopera()"> &euro;
				</td>
			</tr>
			
			<tr>
				<td class="label">ricambi utilizzati:</td>
				<td colspan="3">
					<% dim lista_id, tot_ricambi
					tot_ricambi = 0
					sql = " SELECT * FROM grel_art_valori RIGHT JOIN sgtb_dettagli_schede " & _
						  " ON grel_art_valori.rel_id = sgtb_dettagli_schede.dts_ricambio_id " & _
						  " WHERE dts_scheda_id = " & cIntero(request("ID")) & " ORDER BY dts_id "
					rsa.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch 
					
					sql = Replace(sql,"*","ISNULL(rel_id, 0)")
					lista_id = cString(GetValueList(conn, NULL, sql))
					lista_id = Replace(lista_id, " ", "")
					%>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label" colspan="3" style="width:60%">
								<% if rsa.eof then %>
									Nessun ricambio associato.
								<% else %>
									Trovati n&ordm; <%= rsa.recordcount %> record
								<% end if %>
							</td>
							<% sql = "SELECT mar_anagrafica_id FROM gv_articoli WHERE rel_id = " & rs("sc_modello_id") %>
							<td colspan="5" class="content_right" style="margin-right:0px;">
								<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la selezione di un nuovo ricambio associato" <%= ACTIVE_STATUS %>
								   onclick="OpenAutoPositionedScrollWindow('ArticoliSeleziona.asp?IDSCH=<%=request("ID")%>&TYPE=R&EXCLUDE_IDS=<%=lista_id%>&SelectedPage=SchedeDettagliNew&COSTR_ID=<%=GetValueList(conn, NULL, sql)%>', 'DettScheda', 600, 525, true)">
									AGGIUNGI RICAMBIO
								</a>
							</td>
						</tr>
						<% if not rsa.eof then %>
							<tr>
								<th class="L2" width="14%">codice</th>
								<th class="L2">ricambio</th>
								<th class="l2_center" width="11%">prezzo</th>
								<th class="l2_center" width="8%">quantit&agrave;</th>
								<th class="l2_center" width="10%">sconto</th>
								<th class="l2_center" width="13%">totale</th>
								<th colspan="2" class="l2_center" width="20%">operazioni</th>
							</tr>
							<% while not rsa.eof %>
								<tr>
									<td class="content"><%= rsa("dts_ricambio_codice")%></td>
									<td class="content"><%= rsa("dts_ricambio_nome")%></td>
									<td class="content_center"><%= FormatPrice(cReal(rsa("dts_ricambio_prezzo")), 2, false)%> &euro;</td>
									<td class="content_center"><%= rsa("dts_ricambio_qta")%></td>
									<td class="content_center"><%= rsa("dts_ricambio_sconto")%> %</td>
									<td class="content_center" nowrap><%= FormatPrice(cReal(rsa("dts_prezzo_totale")), 2, false)%> &euro;</td>
									<% tot_ricambi = tot_ricambi + cReal(rsa("dts_prezzo_totale")) %>
									<td class="content_center">
										<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la procedura di modifica del dettaglio della scheda" <%= ACTIVE_STATUS %>
										   onclick="OpenAutoPositionedWindow('SchedeDettagliMod.asp?ID=<%=rsa("dts_id") %>', 'DettScheda', 530, 450)">
											MODIFICA
										</a>
									</td>
									<td class="content_center">
										<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione del dettaglio della scheda" <%= ACTIVE_STATUS %>
										   onclick="OpenDeleteWindow('DETTAGLI_SCHEDE','<%= rsa("dts_id") %>');">
											CANCELLA
										</a>
									</td>
								</tr>
								<%rsa.movenext
							wend
						end if
						rsa.close %>
					</table>
				</td>
			</tr>
			<tr>
				<td class="label">note di chiusura:</td>
				<td class="content" colspan="3">
					<textarea style="width:100%;" rows="5" name="tft_sc_note_chiusura"><%=IIF(post,request("tft_sc_note_chiusura"),rs("sc_note_chiusura")) %></textarea>
				</td>
			</tr>
			
			<tr>
				<td colspan="4">
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<%	sql = " SELECT * FROM (sgtb_descrittori d" + _
							  " LEFT JOIN srel_descrittori_schede r ON (d.des_id = r.rds_descrittore_id AND r.rds_scheda_id="& CIntero(request("ID")) &")) " + _
							  " LEFT JOIN sgtb_descrittori_raggruppamenti g ON d.des_raggruppamento_id = g.rag_id " + _
							  " ORDER BY rds_ordine "
						CALL DesFullFormComplete(NULL, conn, sql, "sgtb_descrittori", "des_id", "des_nome_it", "des_tipo", "", "", "", "", "rds_valore_", "rds_memo_", "rag_titolo_it", (request("ID") = ""), 2, true, 24)	
						%>
					</table>
				</td>
			</tr>
		<% end if %>
		
		<tr>
			<td colspan="4">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px; <%=IIF(is_officina,"visibility:hidden; display:none; ","")%>">
					<tr><th colspan="4">DATI DEL TRASPORTO</th></tr>
					<tr><th class="L2" colspan="4">presa</th></tr>
					<tr>
						<td class="label" style="width:20%;">numero DDT di carico:</td>
						<td class="content" colspan="3">
							<input type="text" class="text" name="tft_sc_numero_DDT_di_carico" value="<%=IIF(post,request("tft_sc_numero_DDT_di_carico"),rs("sc_numero_DDT_di_carico")) %>" maxlength="255" size="52">
						</td>
					</tr>
					<tr>
						<td class="label">data DDT di carico:</td>
						<td class="content" colspan="3">
							<% CALL WriteDataPicker_Input_Manuale("form1", "tfd_sc_data_DDT_di_carico", IIF(post,request("tfd_sc_data_DDT_di_carico"),rs("sc_data_DDT_di_carico")), "", "/", false, true, LINGUA_ITALIANO, "", true, "") 
							%>				
							<% if is_officina then %>
								<script language="JavaScript" type="text/javascript">
									//document.getElementById("tfd_sc_data_DDT_di_carico").disabled = true;
									//document.getElementById("form1_link_scegli_tfd_sc_data_DDT_di_carico").onclick = '';
								</script>
							<% end if %>
						</td>
					</tr>
				</table>
				<% if assegnata then %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px; <%=IIF(is_officina,"visibility:hidden; display:none; ","")%>">
						<% if cIntero(rs("sc_documento_ritiro_id")) > 0 then 
							sql = " SELECT * FROM sgtb_ddt INNER JOIN sgtb_ddt_categorie ON sgtb_ddt.ddt_categoria_id=sgtb_ddt_categorie.cat_id " & _
								  " WHERE ddt_id = " & cIntero(rs("sc_documento_ritiro_id"))
							rsa.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
							%>
							<tr>
								<td class="label">numero e data:</td>
								<td class="content" colspan="3">
									<% CALL RitiroLink(rs("sc_documento_ritiro_id")&"&CAT_ID="&rsa("ddt_categoria_id")&"&STANDALONE=true", "richiesta di ritiro n. "&rsa("ddt_numero")&" del "&rsa("ddt_data"))%>
								</td>
							</tr>
							<% rsa.close() %>
						<% end if %>
						<tr>
							<td class="label" style="width:20%;">costo ritiro:</td>
							<td class="content" colspan="3">
								<input type="text" class="number" name="tfn_sc_costo_presa" value="<%= FormatPrice(cReal(IIF(post,request("tfn_sc_costo_presa"),rs("sc_costo_presa"))), 2, false) %>" size="7" onchange="FunzPrezzoPresaRiconsegna()"> &euro;
							</td>
						</tr>				
						
						<% if cIntero(rs("sc_rif_DDT_di_resa_id")) > 0 then 
							sql = " SELECT * FROM sgtb_ddt INNER JOIN sgtb_ddt_categorie ON sgtb_ddt.ddt_categoria_id=sgtb_ddt_categorie.cat_id " & _
								  " WHERE ddt_id = " & cIntero(rs("sc_rif_DDT_di_resa_id"))
							rsa.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
							%>			
							<tr><th class="L2" colspan="4">riconsegna - <%= uCase(rsa("cat_nome_it"))%></th></tr>
							<tr>
								<td class="label">costo riconsegna:</td>
								<td class="content" colspan="3">
									<input type="text" class="number" name="tfn_sc_costo_riconsegna" value="<%= FormatPrice(cReal(IIF(post,request("tfn_sc_costo_riconsegna"),rs("sc_costo_riconsegna"))), 2, false) %>" size="7" onchange="FunzPrezzoPresaRiconsegna()"> &euro;
								</td>
							</tr>
							<tr>
								<td class="label">numero e data:</td>
								<td class="content" colspan="3">
									<% CALL SpedizioneLink(rsa("ddt_id")&"&CAT_ID="&rsa("ddt_categoria_id")&"&STANDALONE=true", lCase(rsa("cat_nome_it"))&" n. "&rsa("ddt_numero")&" del "&rsa("ddt_data"))%>
								</td>
							</tr>
							<tr>
								<td class="label">trasportatore:</td>
								<%
								if cIntero(rsa("ddt_trasportatore_id"))>0 then
									trasportatoreId = cIntero(rsa("ddt_trasportatore_id"))
									sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & rsa("ddt_trasportatore_id")
									rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
									%>
									<td class="content" colspan="3"><%=ContactFullName(rsd)%></td>
									<%
									rsd.close 
								else
									%>
									<td class="note" colspan="3">non specificato</td>
									<%
								end if
								%>
							</tr>
							<% rsa.close() %>
						<% end if %>
					</table>	
						
					<script language="JavaScript" type="text/javascript">
						function FunzPrezzoPresaRiconsegna(){
							var p_presa = toNumber(document.form1.tfn_sc_costo_presa.value);
							<% if cIntero(rs("sc_rif_DDT_di_resa_id")) > 0 then %>
								var p_riconsegna = toNumber(document.form1.tfn_sc_costo_riconsegna.value);
							<% else %>
								var p_riconsegna = 0;
							<% end if %>
							document.getElementById("PrezzoPresaRiconsegna").innerHTML = FormatNumber(p_presa + p_riconsegna, 2);
							document.form1.tfn_sc_costo_presa.value = FormatNumber(p_presa, 2);
							<% if cIntero(rs("sc_rif_DDT_di_resa_id")) > 0 then %>
								document.form1.tfn_sc_costo_riconsegna.value = FormatNumber(p_riconsegna, 2);
							<% end if %>
							RicalcolaTotali();
						}
						
						function FunzPrezzoManodopera(){
							var p_manodopera = toNumber(document.form1.tfn_sc_prezzo_manodopera.value);
							var o_manodopera = toNumber(document.form1.tfn_sc_ora_manodopera_intervento.value);

							document.getElementById("PrezzoManodopera").innerHTML = FormatNumber(p_manodopera * o_manodopera, 2);
							document.form1.tfn_sc_prezzo_manodopera.value = FormatNumber(p_manodopera, 2);
							
							RicalcolaTotali();
						}
						
						function RicalcolaTotali(){
							var p_1 = toNumber(document.getElementById("PrezzoPresaRiconsegna").innerHTML);
							var p_2 = toNumber(document.getElementById("PrezzoManodopera").innerHTML);
							var p_3 = toNumber(document.getElementById("PrezzoRicambi").innerHTML);
							document.getElementById("PrezzoTotale").innerHTML = FormatNumber(p_1 + p_2 + p_3, 2);
						}
						
					</script>
					
					<tr><th colspan="4">RIEPILOGO COSTI</th></tr>
					<tr>
						<td class="label">costi di presa/riconsegna:</td>
						<td class="content" colspan="3">
							<span id="PrezzoPresaRiconsegna"></span>&nbsp;&euro;
						</td>
					</tr>
					<tr>
						<td class="label">costo di manodopera:</td>
						<td class="content" colspan="3">
							<span id="PrezzoManodopera"></span>&nbsp;&euro;
						</td>
					</tr>
					<tr>
						<td class="label">costo totale ricambi:</td>
						<td class="content" colspan="3">
							<span id="PrezzoRicambi"><%=FormatPrice(cReal(tot_ricambi), 2, false)%></span>&nbsp;&euro;
						</td>
					</tr>
					<tr>
						<td class="label">TOTALE SCHEDA:</td>
						<td class="content" colspan="3">
							<span id="PrezzoTotale"></span>&nbsp;&euro;
						</td>
					</tr>
				<% else %>
					<tr>
						<td class="label">costo ritiro:</td>
						<td class="content" colspan="3">
							<input type="text" class="number" name="tfn_sc_costo_presa" value="<%= FormatPrice(cReal(IIF(post,request("tfn_sc_costo_presa"),rs("sc_costo_presa"))), 2, false) %>" size="7" onchange="FunzPrezzoPresaRiconsegna()" <%=IIF(is_officina,"disabled","")%>> &euro;
						</td>
					</tr>
				<% end if %>
			</td>
		</tr>
		<tr>
			<% CALL Form_DatiModifica(conn, rs, "sc_") %>		
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
				<input type="submit" class="button" name="salva_continua" value="SALVA & TORNA ALL'ELENCO">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>


<div id="pulsanti" style="position:absolute; top:93px; left:810px; width:200px;">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Operazioni</caption>
		<tr><th>SCHEDA</th></tr>
		<tr>
			<td class="content_center" style="font-size:5px;">
				&nbsp;<br>
				<a class="button" style="width:120px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_STAMPA_SCHEDA"), "it") & "&SCHEDAID="&rs("sc_id")&"&CLIENTEID="&rs("sc_cliente_id")&"&IDCNT="&rs("IDElencoIndirizzi")&"&KEY="&rs("codiceInserimento")&"&SHOW_GARANZIA=true"%>" target="ID_SCHEDA_<%=rs("sc_id")%>_stampa"
				onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_stampa', 800, 800, true)" title="visualizza riepilogo scheda" <%= ACTIVE_STATUS %>>
					RIEPILOGO SCHEDA COMPLETA
				</a>
				<br>&nbsp;<br>
			</td>
		</tr>
		<tr>
			<td class="content_center" style="font-size:5px;">
				&nbsp;<br>
				<a class="button" style="width:120px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_STAMPA_SCHEDA"), "it") & "&SCHEDAID="&rs("sc_id")&"&CLIENTEID="&rs("sc_cliente_id")&"&IDCNT="&rs("IDElencoIndirizzi")&"&KEY="&rs("codiceInserimento")%>" target="ID_SCHEDA_<%=rs("sc_id")%>_stampa"
				onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_stampa', 800, 800, true)" title="visualizza riepilogo scheda" <%= ACTIVE_STATUS %>>
					RIEPILOGO SCHEDA PER CLIENTE
				</a>
				<br>&nbsp;<br>
			</td>
		</tr>
		<tr>
			<td class="content_center" style="font-size:1px;">
				<% if cIntero(Session("INFOSCHEDE_ID_PAG_PREVENTIVO"))>0 then %>
					<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_PREVENTIVO"), "it") & "&ID_SCHEDA="&rs("sc_id")&"&ID_ADMIN="&Session("ID_ADMIN")%>" target="ID_SCHEDA_<%=rs("sc_id")%>_preventivo"
					onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_preventivo', 520, 300, true)" title="invia il preventivo al cliente" <%= ACTIVE_STATUS %>>
						INVIA PREVENTIVO
					</a>
				<% end if %>
				<br>&nbsp;<br>
				<% if cIntero(Session("INFOSCHEDE_ID_PAG_CONSUNTIVO"))>0 then %>
					<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_CONSUNTIVO"), "it") & "&ID_SCHEDA="&rs("sc_id")&"&ID_ADMIN="&Session("ID_ADMIN")%>" target="ID_SCHEDA_<%=rs("sc_id")%>_consuntivo"
					onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_consuntivo', 520, 300, true)" title="invia il consuntivo al cliente" <%= ACTIVE_STATUS %>>
						INVIA CONSUNTIVO
					</a>
					<br>&nbsp;
				<% end if %>
			</td>
		</tr>
		<% if Session("INFOSCHEDE_ADMIN")<>"" then %>
			<% 'per mostrare la sezione "ritiri" la scheda deve essere in uno degli stati che permettono il ritiro oppure avere già collegata una richiesta di ritiro
			if cBoolean(rs("sts_elenco_ddt_da_ritirare"), false) OR cIntero(rs("sc_documento_ritiro_id")) > 0 then %>
				<tr><th>RITIRI</th></tr>
				<tr>
					<td class="content_center" style="font-size:1px;">
						&nbsp;<br>
						<% if cBoolean(rs("sts_elenco_ddt_da_ritirare"), false) AND cIntero(rs("sc_documento_ritiro_id")) = 0 then %>
							<a class="button_L2" style="width:160px;text-align:center;" href="javascript:void(0)" 
								onclick="OpenAutoPositionedScrollWindow('RitiriNew.asp?CAT_ID=<%=RITIRI_CAT_ID%>&ID_SCHEDA=<%=request("ID")%>&ID_CLIENTE=<%= IIF(post,request.form("tfn_sc_cliente_id"),rs("sc_cliente_id")) %>', 'Spedisci', 450, 480, true)" 
								title="Click per aprire la finestra per creare il documento di ritiro">
								RICHIEDI RITIRO
							</a>
						<% elseif cIntero(rs("sc_documento_ritiro_id")) > 0 then %>
							<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_RICH_RITIRO"), "it") & "&SCHEDAID="&rs("sc_id")&"&ID_ADMIN="&Session("ID_ADMIN")&"&DDTID="&cIntero(rs("sc_documento_ritiro_id"))&"&CLIENTEID="&rs("sc_cliente_id")&"&IDCNT="&rs("IDElencoIndirizzi")&"&KEY="&rs("codiceInserimento")%>" target="ID_SCHEDA_<%=rs("sc_id")%>_ritiro"
							onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_ritiro', 800, 800, true)" title="Click per aprire la finestra per visualizzare il documento di ritiro" <%= ACTIVE_STATUS %>>
								VISUALIZZA RICHIESTA DI RITIRO
							</a>
							<br>&nbsp;<br>
							<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_INVIO_RICH_RIT"), "it") & "&ID_ADMIN="&Session("ID_ADMIN")&"&ID_DDT="&cIntero(rs("sc_documento_ritiro_id"))%>" target="ID_SCHEDA_<%=rs("sc_id")%>_invia_ritiro"
							onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_invia_ritiro', 520, 300, true)" title="Click per inviare il documento di ritiro" <%= ACTIVE_STATUS %>>
								INVIA RICHIESTA DI RITIRO
							</a>
						<% end if %>
						<br>&nbsp;
					</td>
				</tr>
			<% end if %>
			<% if assegnata then 'assegnata ad un centro assistenza 
				'per mostrare la sezione "spedizioni" la scheda deve essere in uno degli stati che permettono la restituzione oppure avere già collegata un ddt
				if cBoolean(rs("sts_elenco_ddt_da_consegnare"), false) OR cIntero(rs("sc_rif_DDT_di_resa_id")) > 0 then %>
					<tr><th>SPEDIZIONI</th></tr>
					<tr>
						<td class="content_center" style="font-size:1px;">
							&nbsp;<br>
							<% dim pag_visualizz, pag_invio_ddt
								sql = "SELECT riv_profilo_id FROM gtb_rivenditori WHERE riv_id = " & rs("sc_cliente_id")
								prof_id = cIntero(GetValueList(conn, NULL, sql))
								if prof_id = CLIENTI_PRIVATI then
									cat_riconsegna = LETTERE_CAT_ID
									pag_visualizz = Session("INFOSCHEDE_ID_PAG_LETT_ACCOMP")
									pag_invio_ddt = Session("INFOSCHEDE_ID_PAG_INVIO_LETT_ACC")
								else
									cat_riconsegna = DDT_CAT_ID
									pag_visualizz = Session("INFOSCHEDE_ID_PAG_DDT")
									pag_invio_ddt = Session("INFOSCHEDE_ID_PAG_INVIO_DDT")
								end if
								%>
							<% if cIntero(rs("sc_rif_DDT_di_resa_id")) = 0 AND cBoolean(rs("sts_elenco_ddt_da_consegnare"), false) then %>
								<a class="button_L2" style="width:160px;text-align:center;" href="javascript:void(0)" 
									onclick="OpenAutoPositionedScrollWindow('SpedizioniNew.asp?CAT_ID=<%=cat_riconsegna%>&ID_SCHEDA=<%=request("ID")%>&ID_CLIENTE=<%= IIF(post,request.form("tfn_sc_cliente_id"),rs("sc_cliente_id")) %>', 'Spedisci', 450, 480, true)" 
									title="Click per aprire la finestra per creare il documento di spedizione">
									SPEDISCI MERCE
								</a>
							<% elseif cIntero(rs("sc_rif_DDT_di_resa_id")) > 0 then %>
								<% dim label 
								sql = "SELECT cat_nome_it FROM sgtb_ddt_categorie WHERE cat_id = " & cat_riconsegna
								label = GetValueList(conn, NULL, sql)
								%>
								<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, pag_visualizz, "it") & "&SCHEDAID="&rs("sc_id")&"&ID_ADMIN="&Session("ID_ADMIN")&"&DDTID="&cIntero(rs("sc_rif_DDT_di_resa_id"))&"&CLIENTEID="&rs("sc_cliente_id")&"&IDCNT="&rs("IDElencoIndirizzi")&"&KEY="&rs("codiceInserimento")%>" target="ID_SCHEDA_<%=rs("sc_id")%>_spedizione"
								onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_spedizione', 800, 800, true)" title="Click per aprire la finestra per visualizzare il documento di riconsegna" <%= ACTIVE_STATUS %>>
									VISUALIZZA <%=uCase(label)%>
								</a>
								<br>&nbsp;<br>
								<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, pag_invio_ddt, "it") & "&ID_ADMIN="&Session("ID_ADMIN")&"&ID_DDT="&cIntero(rs("sc_rif_DDT_di_resa_id"))%>" target="ID_SCHEDA_<%=rs("sc_id")%>_invia_ddt"
								onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_invia_ddt', 800, 800, true)" title="Click per inviare <%=uCase(label)%>" <%= ACTIVE_STATUS %>>
									INVIA <%=uCase(label)%>
								</a>								
								<br>&nbsp;<br>
								<br>&nbsp;<br>
								<% if trasportatoreId > 0 then %>
									<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_LETT_VETT"), "it") & "&SCHEDAID="&rs("sc_id")&"&ID_ADMIN="&Session("ID_ADMIN")&"&DDTID="&cIntero(rs("sc_rif_DDT_di_resa_id"))&"&CLIENTEID="&rs("sc_cliente_id")&"&IDCNT="&rs("IDElencoIndirizzi")&"&KEY="&rs("codiceInserimento")%>" target="ID_SCHEDA_<%=rs("sc_id")%>_lett_vettura"
									onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_lett_vettura', 800, 800, true)" title="Click per visualizzare la lettera di vettura" <%= ACTIVE_STATUS %>>
										VISUALIZZA LETTERA DI VETTURA
									</a>
									<br>&nbsp;<br>
									<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_INVIO_LETT_VETT"), "it") & "&ID_ADMIN="&Session("ID_ADMIN")&"&ID_DDT="&cIntero(rs("sc_rif_DDT_di_resa_id"))%>" target="ID_SCHEDA_<%=rs("sc_id")%>_invia_lett_vettura"
									onclick="OpenAutoPositionedScrollWindow('', 'ID_SCHEDA_<%=rs("sc_id")%>_invia_lett_vettura', 520, 300, true)" title="Click per inviare la lettera di vettura" <%= ACTIVE_STATUS %>>
										INVIA LETTERA DI VETTURA
									</a>
								<% else %>
									<a class="button_L2 disabled" style="width:160px;text-align:center;" title="Per visualizzare il documento di ritiro occorre selezionare il trasportatore" <%= ACTIVE_STATUS %>>
										VISUALIZZA LETTERA DI VETTURA
									</a>
									<br>&nbsp;<br>
									<a class="button_L2 disabled" style="width:160px;text-align:center;" title="Per inviare il documento di ritiro occorre selezionare il trasportatore" <%= ACTIVE_STATUS %>>
										INVIA LETTERA DI VETTURA
									</a>
								<% end if %>
							<% end if %>
							<br>&nbsp;
						</td>
					</tr>
				<% end if %>
			<% end if %>
		<% end if %>
	</table>
</div>
<script language="JavaScript" type="text/javascript">
	//floatMaking('pulsanti', getXCoord(document.getElementById('pulsanti')), getYCoord(document.getElementById('pulsanti')), 10);

	<% if assegnata then %>
		FunzPrezzoPresaRiconsegna();
		FunzPrezzoManodopera();
	<% end if %>
	
	FitWindowSize(this);
</script>
</body>
</html>
<%
set rs = nothing
set rsa = nothing
set rsd = nothing

%>