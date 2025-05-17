<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" AND (request("salva")<>"" OR request("salva_elenco")<>"") then
	Server.Execute("SpedizioniSalva.asp")
end if

dim conn, rs, rsd, sql, rsi, label, standalone
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.Recordset")
set rsi = Server.CreateObject("ADODB.Recordset")

if request("STANDALONE") = "true" then
	standalone = true
else
	standalone = false
end if

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("INFOSCHEDE_SPEDIZIONI_SQL"), "ddt_ID", "SpedizioniMod.asp")
end if

%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	

sql = " SELECT *, (SELECT COUNT(sc_id) FROM sgtb_schede WHERE sc_rif_DDT_di_resa_id = sgtb_ddt.ddt_id) AS N_IND_DIV " + _
	  " FROM sgtb_ddt INNER JOIN sgtb_ddt_causali ON sgtb_ddt.ddt_causale_id = sgtb_ddt_causali.cau_id " + _
	  " INNER JOIN gv_rivenditori ON sgtb_ddt.ddt_cliente_id = gv_rivenditori.riv_id " + _
	  " WHERE ddt_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 

label = GetValueList(conn, NULL, "SELECT cat_nome_it FROM sgtb_ddt_categorie WHERE cat_id = "& rs("ddt_categoria_id"))


dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione spedizioni - modifica " & label
if standalone then
	dicitura.puls_new = ""
	dicitura.link_new = ""
else
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = "Spedizioni.asp"
end if
dicitura.scrivi_con_sottosez() 

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_ddt_categoria_id" value="<%= rs("ddt_categoria_id") %>">
	<% if standalone then %><input type="hidden" name="reload" value="true"><% end if %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati <%=label%></td>
					<td align="right" style="font-size: 1px;">
						<% if standalone then %>
							<a class="button" href="javascript:window.close();" title="chiudi la finestra" <%= ACTIVE_STATUS %>>
								CHIUDI</a>
						<% else %>
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="<%=Lcase(label)%> precedente" <%= ACTIVE_STATUS %>>
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="<%=Lcase(label)%> successivo" <%= ACTIVE_STATUS %>>
								SUCCESSIVO &gt;&gt;
							</a>
						<% end if %>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI <%=Ucase(label)%></th></tr>
		<tr>
			<td class="label" style="width:18%;">numero:</td>
			<td class="content" colspan="3"><%=rs("ddt_numero")%></td>
			<input type="hidden" name="tfn_ddt_numero" value="<%=rs("ddt_numero")%>">
		</tr>
		<tr>
			<td class="label">data:</td>
			<td class="content" colspan="3">
				<%= DateIta(rs("ddt_data"))%>
			</td>
			<input type="hidden" name="tfd_ddt_data" value="<%=rs("ddt_data")%>">
		</tr>
		<tr>
			<td class="label">causale:</td>
			<td class="content" colspan="3">
				<% sql = "SELECT * FROM sgtb_ddt_causali ORDER BY cau_ordine, cau_titolo_it"
				CALL dropDown(conn, sql, "cau_id", "cau_titolo_it", "tfn_ddt_causale_id", rs("ddt_causale_id"), true, "", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">trasporto a cura:</td>
			<td class="content" colspan="3">
				<% sql = "SELECT * FROM sgtb_ddt_trasporto ORDER BY tra_titolo_it"
				CALL dropDown(conn, sql, "tra_id", "tra_titolo_it", "tfn_ddt_trasporto_id", rs("ddt_trasporto_id"), true, "", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">porto:</td>
			<td class="content" colspan="3">
				<% sql = "SELECT * FROM sgtb_ddt_porto ORDER BY por_titolo_it"
				CALL dropDown(conn, sql, "por_id", "por_titolo_it", "tfn_ddt_porto_id", rs("ddt_porto_id"), true, "", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">cliente:</td>
			<td class="content" colspan="3"><%= ContactFullName(rs)%></td>
			<input type="hidden" name="tfn_ddt_cliente_id" value="<%=rs("ddt_cliente_id")%>">
		</tr>
		<tr>
			<td class="label">destinazione:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<% dim nome_destinazione
					if cIntero(rs("ddt_destinazione_id"))>0 then
						sql = "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi = " & rs("ddt_destinazione_id")
						rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						nome_destinazione = ContactAddress(rsd)
						rsd.close 
					else
						nome_destinazione = ""
					end if					
					%>
					<td>
						<input type="hidden" name="tfn_ddt_destinazione_id" value="<%= rs("ddt_destinazione_id") %>">
						<input READONLY type="text" name="destinazione" style="padding-left:3px; width:100%" value="<%= nome_destinazione %>" 
							   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=destinazione&field_id=tfn_ddt_destinazione_id&selected=' + tfn_ddt_destinazione_id.value, 'SelezioneDestinazione', 620, 520, true)" title="Click per aprire la finestra per la selezione della destinazione">
					</td>
					<td width="30%" nowrap>
						<a class="button_input" href="javascript:void(0)" onclick="form1.destinazione.onclick();" 
							 title="Apre la filnestra per la selezione del destinazione" <%= ACTIVE_STATUS %>>
							SELEZIONA DESTINAZIONE
						</a>
						&nbsp;(*)
					</td>
				</tr>
				</table>
			</td>
		</tr>
		
		<script language="JavaScript" type="text/javascript">
			function abilita(){
				var trasp = document.getElementById('trasportatore');
				var peso = document.getElementById('tft_ddt_peso');
				var span_peso = document.getElementById('span_peso');
				var volume = document.getElementById('tft_ddt_volume');
				var span_volume = document.getElementById('span_volume');
				var colli = document.getElementById('tft_ddt_numero_colli');
				var span_colli = document.getElementById('span_colli');
				if (trasp.value != ''){
					peso.disabled = '';
					peso.className = 'text';
					span_peso.innerHTML='(*)'
					volume.disabled = '';
					volume.className = 'text';
					span_volume.innerHTML='(*)'
					colli.disabled = '';
					colli.className = 'text';
					span_colli.innerHTML='(*)'
				}
				else{
					peso.disabled = 'disabled';
					peso.className = 'text disabled';
					span_peso.innerHTML=''
					volume.disabled = 'disabled';
					volume.className = 'text disabled';
					span_volume.innerHTML=''
					colli.disabled = 'disabled';
					colli.className = 'text disabled';
					span_colli.innerHTML=''
				}
			}
			
			function resetTrasp(){
				document.form1.tfn_ddt_trasportatore_id.value = '';
				document.form1.trasportatore.value = '';
				abilita();
			}
			
		</script>
		<tr>
			<td class="label">trasportatore:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<% dim nome_trasportatore
					if cIntero(rs("ddt_trasportatore_id"))>0 then
						sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & rs("ddt_trasportatore_id")
						rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						nome_trasportatore = ContactFullName(rsd)
						rsd.close 
					else
						nome_trasportatore = ""
					end if					
					%>
					<td>
						<input type="hidden" name="tfn_ddt_trasportatore_id" value="<%= rs("ddt_trasportatore_id") %>"  onchange="abilita()">
						<input READONLY type="text" name="trasportatore" id="trasportatore" style="padding-left:3px; width:100%" value="<%= nome_trasportatore %>" 
							   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=trasportatore&field_id=tfn_ddt_trasportatore_id&selected=' + tfn_ddt_trasportatore_id.value + '&filtro_profilo=<%=TRASPORTATORI%>&AFTER=onchange', 'SelezioneTrasportatore', 620, 480, true)" title="Click per aprire la finestra per la selezione del trasportatore">
					</td>
					<td width="41%" nowrap>
						<a class="button_input" href="javascript:void(0)" onclick="form1.trasportatore.onclick();" 
							 title="Apre la filnestra per la selezione del trasportatore" <%= ACTIVE_STATUS %>>
							SELEZIONA TRASPORTATORE
						</a>
						<a class="button_input" href="javascript:void(0);" id="trasp_reset" onclick="resetTrasp();" title="Reset" <%= ACTIVE_STATUS %>>
							RESET
						</a>
						&nbsp;(*)
					</td>
				</tr>
				</table>
			</td>	
		</tr>
		<tr>
			<td class="label">peso:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_peso" id="tft_ddt_peso" value="<%=rs("ddt_peso")%>" maxlength="255" size="22">
				<span id="span_peso">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">volume:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_volume" id="tft_ddt_volume" value="<%=rs("ddt_volume")%>" maxlength="255" size="22">
				<span id="span_volume">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">numero colli:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_numero_colli" id="tft_ddt_numero_colli" value="<%=rs("ddt_numero_colli")%>" maxlength="255" size="22">
				<span id="span_colli">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">contrassegno:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_contrassegno" id="tft_ddt_contrassegno" value="<%=rs("ddt_contrassegno")%>" maxlength="100" size="22">
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			abilita();

			function rimuovi(idScheda){
				var campo_hidden = document.getElementById('scheda_' + idScheda)
				var button = document.getElementById('button_' + idScheda)
				campo_hidden.value = '';
				button.disabled = 'disabled';
				button.className = 'button_L2_disabled';
			}
		</script>
		<% if cIntero(rs("ddt_cliente_id"))>0 then %>
			<% sql = " SELECT * FROM (sgtb_schede INNER JOIN gv_articoli ON sgtb_schede.sc_modello_id = gv_articoli.rel_id) " & _
					 " INNER JOIN sgtb_stati_schede ON sgtb_schede.sc_stato_id = sgtb_stati_schede.sts_id " & _
					 " LEFT JOIN sgtb_ddt ON sgtb_schede.sc_rif_DDT_di_resa_id = sgtb_ddt.ddt_id " & _
					 " WHERE sc_rif_DDT_di_resa_id="&cIntero(request("ID")) & _
					 " ORDER BY sc_data_ricevimento DESC, sc_numero "
			rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
				<tr><th colspan="5">SCHEDE ASSOCIATE</th></tr>
				<% if rsd.eof then %>
					<tr><td colspan="4" class="note">Nessuna scheda trovata per il cliente selezionato</td></tr>
				<% else %>
					<tr>
						<th class="l2_center" style="width:8%">&nbsp;</th>
						<th class="l2_center" style="width:22%">numero scheda e data</th>
						<th class="l2_center" style="width:18%">stato</th>
						<th class="l2_center" style="width:14%">costo restituzione</th>
						<th class="L2">modello</th>
					</tr>
				<% end if %>
				<% while not rsd.eof %>
					<tr>
						<td class="content_center">
							<input type="hidden" id="scheda_<%=rsd("sc_id")%>" name="id_schede" value="<%=rsd("sc_id")%>">
							<% if cIntero(rsd.recordcount)>1 then %>
								<a class="button_L2" href="javascript:void(0)" id="button_<%=rsd("sc_id")%>" onclick="rimuovi(<%=rsd("sc_id")%>);" 
									title="Rimuovi associazione con questa scheda" <%= ACTIVE_STATUS %>>
									rimuovi
								</a>
							<% else %>
								<a class="button_L2_disabled" disabled href="javascript:void(0)" title="Impossibile rimuovere associazione con questa scheda">
									rimuovi
								</a>
							<% end if %>
						</td>
						<td class="content_center"><% CALL SchedaLink(rsd("sc_id"), rsd("sc_numero") & " del " & rsd("sc_data_ricevimento"))%></td>
						<td class="content"><%=rsd("sts_nome_it")%></td>
						<td class="content_center">
							<input type="text" class="number" name="costo_scheda_<%=rsd("sc_id")%>" value="<%= FormatPrice(cReal(rsd("sc_costo_riconsegna")), 2, false) %>" size="7"> &euro;
						</td>
						<td class="content">
							<% CALL ArticoloLink(rsd("art_id"), rsd("art_nome_it"), rsd("art_cod_int")) %>
							<% if rsd("art_varianti") then %>
								<%= ListValoriVarianti(conn, rsi, rsd("rel_id")) %>
							<% else %>
								&nbsp;
							<% end if %>
						</td>
					</tr>
				<% rsd.moveNext %>
			<% wend %>
			</table>
			<% rsd.close %>
		<% end if %>
		
		<% sql = " SELECT * FROM (sgtb_schede INNER JOIN gv_articoli ON sgtb_schede.sc_modello_id = gv_articoli.rel_id) " & _
			     " INNER JOIN sgtb_stati_schede ON sgtb_schede.sc_stato_id = sgtb_stati_schede.sts_id " & _
				 " LEFT JOIN sgtb_ddt ON sgtb_schede.sc_rif_DDT_di_resa_id = sgtb_ddt.ddt_id " & _
				 " WHERE ISNULL(sts_elenco_ddt_da_consegnare,0)=1 AND  ISNULL(sc_rif_DDT_di_resa_id, 0)=0 " & _
				 " AND sc_cliente_id = " & cIntero(rs("ddt_cliente_id")) & _
				 " ORDER BY sc_data_ricevimento DESC "
		rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr><th colspan="5">SCHEDE NON ASSOCIATE</th></tr>
			<% if rsd.eof then %>
				<tr><td colspan="4" class="note">Nessuna scheda trovata per il cliente selezionato</td></tr>
			<% else %>
				<tr>
					<th class="l2_center" style="width:8%">associa</th>
					<th class="l2_center" style="width:22%">numero scheda e data</th>
						<th class="l2_center" style="width:18%">stato</th>
					<th class="l2_center" style="width:14%">costo restituzione</th>
					<th class="L2">modello</th>
				</tr>
			<% end if %>
			<% while not rsd.eof %>
				<tr>
					<td class="content_center">
						<input type="checkbox" class="noBorder <%=IIF(cIntero(rsd("sc_rif_DDT_di_resa_id"))=cIntero(request("ID")),"checked","")%>" name="id_schede" value="<%=rsd("sc_id")%>" 
								<%= chk(cIntero(rsd("sc_rif_DDT_di_resa_id"))=cIntero(request("ID")))%>>
					</td>
					<td class="content_center"><% CALL SchedaLink(rsd("sc_id"), rsd("sc_numero") & " del " & rsd("sc_data_ricevimento"))%></td>
					<td class="content"><%=rsd("sts_nome_it")%></td>
					<td class="content_center">
						<input type="text" class="number" name="costo_scheda_<%=rsd("sc_id")%>" value="<%= FormatPrice(cReal(rsd("sc_costo_riconsegna")), 2, false) %>" size="7"> &euro;
					</td>
					<td class="content">
						<% CALL ArticoloLink(rsd("art_id"), rsd("art_nome_it"), rsd("art_cod_int")) %>
						<% if rsd("art_varianti") then %>
							<%= ListValoriVarianti(conn, rsi, rsd("rel_id")) %>
						<% else %>
							&nbsp;
						<% end if %>
					</td>
				</tr>
			<% rsd.moveNext %>
		<% wend %>
		<% rsd.close %>
		</table>
		
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr><th colspan="5">ARTICOLI ASSOCIATI</th></tr>
			<tr>
				<td colspan="4">
					<% dim lista_id, tot_ricambi
					tot_ricambi = 0
					sql = " SELECT * FROM grel_art_valori RIGHT JOIN sgtb_dettagli_ddt " & _
						  " ON grel_art_valori.rel_id = sgtb_dettagli_ddt.dtd_articolo_id " & _
						  " WHERE dtd_ddt_id = " & cIntero(request("ID")) & " ORDER BY dtd_id "
					rsd.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch 
					
					sql = Replace(sql,"*","ISNULL(rel_id, 0)")
					lista_id = cString(GetValueList(conn, NULL, sql))
					lista_id = Replace(lista_id, " ", "")
					%>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td colspan="10" class="content_right" style="margin-right:0px;">
								<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la selezione di un nuovo ricambio associato" <%= ACTIVE_STATUS %>
								   onclick="OpenAutoPositionedScrollWindow('ArticoliSeleziona.asp?ID_EXT=<%=request("ID")%>&TYPE=MR&EXCLUDE_IDS=<%=lista_id%>&SelectedPage=SpedizioniDettagliNew', 'DettScheda', 600, 525, true)">
									AGGIUNGI ARTICOLO
								</a>
							</td>
						</tr>
						<% if not rsd.eof then %>
							<tr>
								<th class="L2" width="13%">codice</th>
								<th class="L2">articolo</th>
								<th class="l2_center" width="6%">quantit&agrave;</th>
								<th class="l2_center" width="12%">prezzo unitario</th>
								<th class="l2_center" width="6%">sconto</th>
								<th class="l2_center" width="10%">rif. vs DDT</th>
								<th class="l2_center" width="6%">garanzia</th>
								<th colspan="2" class="l2_center" width="20%">operazioni</th>
							</tr>
							<% while not rsd.eof %>
								<tr>
									<td class="content"><%= rsd("dtd_articolo_codice")%></td>
									<td class="content"><%= rsd("dtd_articolo_nome")%></td>
									<td class="content_center"><%= rsd("dtd_articolo_qta")%></td>
									<td class="content_center"><%= FormatPrice(cReal(rsd("dtd_articolo_prezzo_unitario")), 2, false) %> &euro;</td>
									<td class="content_center"><%= rsd("dtd_articolo_sconto")%> %</td>
									<td class="content_center"><%= rsd("dtd_rif_vs_ddt")%></td>
									<td class="content_center" nowrap><input type="checkbox" disabled class="checkbox" <%= chk(rsd("dtd_in_garanzia")) %>></td>
									<td class="content_center">
										<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la procedura di modifica del dettaglio della scheda" <%= ACTIVE_STATUS %>
										   onclick="OpenAutoPositionedWindow('SpedizioniDettagliMod.asp?ID=<%=rsd("dtd_id") %>', 'DettScheda', 530, 450)">
											MODIFICA
										</a>
									</td>
									<td class="content_center">
										<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione del dettaglio della scheda" <%= ACTIVE_STATUS %>
										   onclick="OpenDeleteWindow('DETTAGLI_DDT','<%= rsd("dtd_id") %>');">
											CANCELLA
										</a>
									</td>
								</tr>
								<%rsd.movenext
							wend
						end if
						rsd.close %>
					</table>
				</td>
			</tr>
		</table>

		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<tr><th colspan="4">NOTE</th></tr>
			<tr>
				<td class="content" colspan="4">
					<textarea style="width:100%;" rows="3" name="tft_ddt_note"><%= rs("ddt_note") %></textarea>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="4">
					(*) Campi obbligatori.
					<input type="submit" class="button" name="salva" value="SALVA">
					<% if not standalone then %>
						<input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA A ELENCO">
					<% end if %>
				</td>
			</tr>
		</table>
		&nbsp;
	</form>
</div>



<div id="pulsanti" style="position:absolute; top:91px; left:760px; width:180px;">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Operazioni</caption>
			<% if Session("INFOSCHEDE_ADMIN")<>"" then %>
				<tr><th>SPEDIZIONI</th></tr>
				<tr>
					<td class="content_center" style="font-size:1px;">
						&nbsp;<br>
						<% dim pag_visualizz, IDCNT, KEY, prof_id, url
						sql = "SELECT riv_profilo_id FROM gtb_rivenditori WHERE riv_id = " & rs("ddt_cliente_id")
						prof_id = cIntero(GetValueList(conn, NULL, sql))
						if prof_id = CLIENTI_PRIVATI then
							' cat_riconsegna = LETTERE_CAT_ID
							pag_visualizz = Session("INFOSCHEDE_ID_PAG_LETT_ACCOMP")
						else
							' cat_riconsegna = DDT_CAT_ID
							pag_visualizz = Session("INFOSCHEDE_ID_PAG_DDT")
						end if
						' sql = "SELECT cat_nome_it FROM sgtb_ddt_categorie WHERE cat_id = " & cat_riconsegna
						' label = GetValueList(conn, NULL, sql)
						
						sql = "SELECT IDElencoIndirizzi FROM gv_rivenditori WHERE riv_id = " & rs("ddt_cliente_id")
						IDCNT = GetValueList(conn, NULL, sql)
						sql = "SELECT codiceInserimento FROM gv_rivenditori WHERE riv_id = " & rs("ddt_cliente_id")
						KEY = GetValueList(conn, NULL, sql)
						
						url = GetPageSiteUrl(conn, pag_visualizz, "it")&"&ID_ADMIN="&Session("ID_ADMIN")&"&DDTID="&cIntero(rs("ddt_id"))&"&CLIENTEID="&rs("ddt_cliente_id")&"&IDCNT="&IDCNT&"&KEY="&KEY 
						%>
						<a class="button_L2" style="width:160px;text-align:center;" href="<%= url %>" target="visualizza_<%=rs("ddt_id")%>"
						onclick="OpenAutoPositionedScrollWindow('', 'visualizza_<%=rs("ddt_id")%>', 800, 800, true)" title="Click per aprire la finestra per visualizzare il documento di riconsegna" <%= ACTIVE_STATUS %>>
							VISUALIZZA <%=uCase(label)%>
						</a>
						<br>&nbsp;<br>
						<% if rs("ddt_trasportatore_id") > 0 then 
							url = GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_LETT_VETT"), "it") & "&ID_ADMIN="&Session("ID_ADMIN")&"&DDTID="&cIntero(rs("ddt_id"))&"&CLIENTEID="&rs("ddt_cliente_id")&"&IDCNT="&IDCNT&"&KEY="&KEY
							%>
							<a class="button_L2" style="width:160px;text-align:center;" href="<%= url %>" target="visualizza_lett_vettura_<%=rs("ddt_id")%>"
							onclick="OpenAutoPositionedScrollWindow('', 'visualizza_lett_vettura_<%=rs("ddt_id")%>', 800, 800, true)" title="Click per visualizzare la lettera di vettura" <%= ACTIVE_STATUS %>>
								VISUALIZZA LETTERA DI VETTURA
							</a>
							<br>&nbsp;<br>
							<% url = GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_INVIO_LETT_VETT"), "it") & "&ID_ADMIN="&Session("ID_ADMIN")&"&ID_DDT="&cIntero(rs("ddt_id"))&"&IDCNT="&IDCNT&"&KEY="&KEY %>
							<a class="button_L2" style="width:160px;text-align:center;" href="<%= url %>" target="invia_lett_vettura_<%=rs("ddt_id")%>"
							onclick="OpenAutoPositionedScrollWindow('', 'invia_lett_vettura_<%=rs("ddt_id")%>', 520, 300, true)" title="Click per inviare la lettera di vettura" <%= ACTIVE_STATUS %>>
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
						<br>&nbsp;
					</td>
				</tr>	
			<% end if %>
	</table>
</div>

</body>
</html>

<% if standalone then %>
	<script language="JavaScript" type="text/javascript">
		FitWindowSize(this);
	</script>
<% end if %>

<%
rs.close
set rs = nothing
set rsd = nothing
set rsi = nothing
conn.Close
set conn = nothing
%>