<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("OrdiniSalva.asp")
end if

dim conn, rs, rsd, sql, rsi, rsp
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.Recordset")
set rsi = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_ORDINI_SQL"), "ord_ID", "OrdiniMod.asp")
end if

%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione ordini - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Ordini.asp"
dicitura.scrivi_con_sottosez() 
%>
<script language="JavaScript" type="text/javascript">
	function  VerificaStatoOrdine(){
		var t
		for (var i=0; i < form1.stato_ordine.length; i++){
			EnableIfChecked(document.getElementById('stato_ordine_' + form1.stato_ordine[i].value),
							document.getElementById('tfn_ord_stato_id_' + form1.stato_ordine[i].value));
		}
	}
</script>
<div id="content">
<%
sql = " SELECT *, (SELECT COUNT(det_ind_id) FROM gtb_dettagli_ord WHERE det_ord_id = gtb_ordini.ord_id) AS N_IND_DIV " + _
	  " FROM gtb_ordini INNER JOIN gtb_magazzini ON gtb_ordini.ord_magazzino_id=gtb_magazzini.mag_id " + _
	  " INNER JOIN gtb_stati_ordine ON gtb_ordini.ord_stato_id = gtb_stati_ordine.so_id " + _
	  " INNER JOIN gv_rivenditori ON gtb_ordini.ord_riv_id = gv_rivenditori.riv_id " + _
	  " INNER JOIN gtb_listini ON gv_rivenditori.riv_listino_id = gtb_listini.listino_id " + _
	  " LEFT JOIN gtb_lista_codici ON gv_rivenditori.riv_LstCod_id = gtb_lista_codici.lstCod_id " + _
	  " WHERE ord_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfd_ord_data_ultima_mod" value="NOW">
	<input type="hidden" name="tfn_ord_riv_id" value="<%= rs("ord_riv_id") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati dell'ordine</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="ordine precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="ordine successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI DEL'ORDINE</th></tr>
		<tr>
			<td class="label">numero:</td>
			<td class="content_b" colspan="3"><%= rs("ord_id") %></td>
		</tr>
		<tr>
			<td class="label">data:</td>
			<td class="content" style="width:25%;">
				<% if rs("ord_archiviato") OR rs("ord_movimenta") OR rs("ord_impegna") then %>
					<%= DateIta(rs("ord_data")) %>
				<% else %>
					<% CALL WriteDataPicker_Input("form1", "tfd_ord_data", DateITA(rs("ord_data")), "", "/", false, true, LINGUA_ITALIANO) %>
				<% end if %>
			</td>
			<td class="label">riferimento:</td>
			<td class="content">
				<% if rs("ord_archiviato") or rs("ord_movimenta") then %>
					<%= rs("ord_cod") %>
				<% else %>
					<input type="text" class="text" name="tft_ord_cod" value="<%= request("tft_ord_cod") %>" maxlength="50" size="20">
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="<%= IIF(cInteger(rs("riv_lstCod_id"))>0, "4", "3") %>">cliente:</td>
			<td class="label">nominativo:</td>
			<td class="content" colspan="2">
				<a href="ClientiGestione.asp?ID=<%= rs("IDElencoIndirizzi") %>" title="apri la scheda con i dati del cliente" <%= ACTIVE_STATUS %> target="cliente">
					<%= ContactFullName(rs) %>
				</a>
			</td>
		</tr>
		<tr>
			<td class="label">listino prezzi associato:</td>
			<td class="content" colspan="2"><%= rs("listino_codice") %></td>
		</tr>
		<% if cInteger(rs("riv_lstCod_id"))>0 then %>
			<tr>
				<td class="label">lista codici personalizzata:</td>
				<td class="content" colspan="2"><%= rs("LstCod_nome") %></td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">tipo di pagamento:</td>
			<td class="content" colspan="2">
				<% 	dropDown conn, "SELECT * FROM gtb_modipagamento where "& SQL_isTrue(conn, "mosp_se_abilitato") &" ORDER BY mosp_nome_it", "mosp_id", "mosp_nome_it", "tfn_ord_modopagamento_id", rs("ord_modopagamento_id"), false, " disabled ", LINGUA_ITALIANO %>
			</td>
		</tr>
		
		<% if cString(rs("ord_conferma_ordine_data"))<>"" then %>
			<tr>
				<td class="label">data di conferma:</td>
				<td class="content" colspan="3">
					<%= DateTimeIta(rs("ord_conferma_ordine_data")) %>
				</td>
			</tr>
		<% end if %>
		
		<tr><th colspan="4">GESTIONE DEL'ORDINE</th></tr>
		<tr>
			<td class="label">magazzino:</td>
			<td class="content" colspan="3">
				<%= rs("mag_nome") %>
			</td>
		</tr>
		<tr>
			<td class="label">stato dell'ordine:</td>
			<td colspan="3">
				<table cellpadding="0" cellspacing="1" width="100%">
					<input type="hidden" name="old_stato_ordine" value="<%= rs("so_stato_ordini") %>">
					<% 'stato ordine non confermato
					if NOT(rs("ord_movimenta")) then
						'disponibile per: ordini non confermati, ordini confermati ed ordni archiviati non evasi
						%>
						<tr>
							<td width="22%" class="content<%= STILI_STATI_ORDINE(ORDINE_NON_CONFERMATO) %>">
								<input type="radio" class="checkbox" name="stato_ordine" id="stato_ordine_<%= ORDINE_NON_CONFERMATO %>" value="<%= ORDINE_NON_CONFERMATO %>" <%= chk(NOT(rs("ord_impegna")) AND NOT(rs("ord_movimenta")) AND NOT(rs("ord_archiviato"))) %> onclick="VerificaStatoOrdine()">
								<%= STATI_ORDINE(ORDINE_NON_CONFERMATO) %>
							</td>
							<td class="content<%= STILI_STATI_ORDINE(ORDINE_NON_CONFERMATO) %>" colspan="2">
								<% sql = "SELECT * FROM gtb_stati_ordine WHERE so_stato_ordini=" & ORDINE_NON_CONFERMATO & " ORDER BY so_ordine"
								CALL dropDown(conn, sql, "so_id", "so_nome_it", "tfn_ord_stato_id", rs("ord_stato_id"), true, " id=""tfn_ord_stato_id_" & ORDINE_NON_CONFERMATO & """", LINGUA_ITALIANO) %>
							</td>
						</tr>
					<% end if
					'stato ordine confermato
					if cInteger(rs("N_IND_DIV"))>0 then
						if NOT(rs("ord_movimenta")) AND NOT(rs("ord_archiviato")) then
							'disponibile per: ordini non confermati, ordini confermati
							%>
							<tr>
								<td width="22%" class="content<%= STILI_STATI_ORDINE(ORDINE_CONFERMATO) %>">
									<input type="radio" class="checkbox" name="stato_ordine" id="stato_ordine_<%= ORDINE_CONFERMATO %>" value="<%= ORDINE_CONFERMATO %>" <%= chk(rs("ord_impegna")) %> onclick="VerificaStatoOrdine()">
									<%= STATI_ORDINE(ORDINE_CONFERMATO) %>
								</td>
								<td class="content<%= STILI_STATI_ORDINE(ORDINE_CONFERMATO) %>" colspan="2">
									<% sql = "SELECT * FROM gtb_stati_ordine WHERE so_stato_ordini=" & ORDINE_CONFERMATO & " ORDER BY so_ordine"
									CALL dropDown(conn, sql, "so_id", "so_nome_it", "tfn_ord_stato_id", rs("ord_stato_id"), true, " id=""tfn_ord_stato_id_" & ORDINE_CONFERMATO & """", LINGUA_ITALIANO) %>
								</td>
							</tr>
						<%end if
						'stato ordine evaso
						if rs("ord_impegna") OR rs("ord_movimenta") then
							'disponibili per: ordini confermati, ordini evasi
							%>
							<tr>
								<td width="22%" class="content<%= STILI_STATI_ORDINE(ORDINE_EVASO) %>">
									<input type="radio" class="checkbox" name="stato_ordine" id="stato_ordine_<%= ORDINE_EVASO %>" value="<%= ORDINE_EVASO %>" <%= chk(rs("ord_movimenta")) %> onclick="VerificaStatoOrdine()">
									<%= STATI_ORDINE(ORDINE_EVASO) %>
								</td>
								<td class="content<%= STILI_STATI_ORDINE(ORDINE_EVASO) %>" colspan="2">
									<% sql = "SELECT * FROM gtb_stati_ordine WHERE so_stato_ordini=" & ORDINE_EVASO & " ORDER BY so_ordine"
									CALL dropDown(conn, sql, "so_id", "so_nome_it", "tfn_ord_stato_id", rs("ord_stato_id"), true, " id=""tfn_ord_stato_id_" & ORDINE_EVASO & """", LINGUA_ITALIANO) %>
								</td>
							</tr>
						<%end if
						'stato ordine archiviato
						if NOT(rs("ord_impegna")) then 
							'disponibile per: ordini non confermati, ordini evasi
							%>
							<tr>
								<td width="22%" class="content<%= STILI_STATI_ORDINE(ORDINE_ARCHIVIATO) %>">
									<input type="radio" class="checkbox" name="stato_ordine" id="stato_ordine_<%= ORDINE_ARCHIVIATO %>" value="<%= ORDINE_ARCHIVIATO %>" <%= chk(rs("ord_archiviato")) %> onclick="VerificaStatoOrdine()">
									<%= STATI_ORDINE(ORDINE_ARCHIVIATO) %>
								</td>
								<td class="content<%= STILI_STATI_ORDINE(ORDINE_ARCHIVIATO) %>" colspan="2">
									<% sql = "SELECT * FROM gtb_stati_ordine WHERE so_stato_ordini=" & ORDINE_ARCHIVIATO & " ORDER BY so_ordine"
									CALL dropDown(conn, sql, "so_id", "so_nome_it", "tfn_ord_stato_id", rs("ord_stato_id"), true, " id=""tfn_ord_stato_id_" & ORDINE_ARCHIVIATO & """", LINGUA_ITALIANO) %>
								</td>
							</tr>
						<% end if 
					end if%>
				</table>
			</td>
		</tr>
		<% if cString(rs("ord_evasione_ordine_data"))<>"" then %>
			<tr>
				<td class="label"><%= IIF(rs("ord_movimenta"), "data di evasione", "evasione prevista")%>:</td>
				<td class="content" colspan="3">
					<%= DateTimeIta(rs("ord_evasione_ordine_data")) %>
				</td>
			</tr>
		<% end if %>
		<%
		dim n_porti, n_tipo_consegna, n_trasportatori
		n_porti = cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_porti"))
		n_tipo_consegna = cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_tipo_consegna"))
		n_trasportatori = cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_trasportatori"))
		%>
		<% if n_porti > 0 OR n_tipo_consegna > 0 OR n_trasportatori > 0 OR cString(rs("ord_evasione_ordine_data") & rs("ord_consegna_ordine_data"))<>"" then %>
			<tr><th colspan="4">TRASPORTO E CONSEGNA</th></tr>
		<% end if %>
		<% if n_porti > 0 then %>
			<tr>
				<td class="label">porto:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT * FROM gtb_porti ORDER BY prt_nome_it"
					dropDown conn, sql, "prt_id", "prt_nome_it", "tfn_ord_porto_id", rs("ord_porto_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
		<% if n_tipo_consegna > 0 then %>
			<tr>
				<td class="label">modalit&agrave; consegna:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT * FROM gtb_tipo_consegna ORDER BY tco_ordine, tco_nome_it"
					dropDown conn, sql, "tco_id", "tco_nome_it", "tfn_ord_tipo_consegna_id", rs("ord_tipo_consegna_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
		<% if n_trasportatori > 0 then %>
			<tr>
				<td class="label">trasportatore:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT * FROM gtb_trasportatori ORDER BY tra_nome_it"
					dropDown conn, sql, "tra_id", "tra_nome_it", "tfn_ord_trasportatore_id", rs("ord_trasportatore_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
		<% if cString(rs("ord_spedizione_ordine_data"))<>"" then %>
			<tr>
				<td class="label">data di spedizione:</td>
				<td class="content" colspan="3">
					<%= DateTimeIta(rs("ord_spedizione_ordine_data")) %>
				</td>
			</tr>
		<% end if %>
		<% if cString(rs("ord_consegna_ordine_data"))<>"" then %>
			<tr>
				<td class="label">data di consegna:</td>
				<td class="content" colspan="3">
					<%= DateTimeIta(rs("ord_consegna_ordine_data")) %>
				</td>
			</tr>
		<% end if %>
		<%
			dim n_pagamenti
			n_pagamenti = cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_pagamenti WHERE pag_ordine_id = " & rs("ord_id")))
		%>
		<%if n_pagamenti>0 then%>
		<tr><th colspan="4">PAGAMENTI</th></tr>
		<tr>
			<td colspan="4">
				<% sql = " SELECT pag_id, pag_data, pag_importo, pag_ordine_id, pag_stato_ordine_id, pag_mosp_id, pag_RAW, mosp_nome_it, so_nome_it " & _
						" FROM gtb_pagamenti " & _
						" INNER JOIN gtb_modipagamento ON gtb_pagamenti.pag_mosp_id = gtb_modipagamento.mosp_id " & _
						" INNER JOIN gtb_stati_ordine ON gtb_pagamenti.pag_stato_ordine_id = gtb_stati_ordine.so_id " & _
						" WHERE pag_ordine_id = " & rs("ord_id")
				rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
				if rsp.recordcount>2 then %> 
					<span class="overflow">
				<% end if%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<th class="L2">metodo pagamento</th>
						<th class="L2">data</th>
						<th class="L2">importo</th>
						<th class="L2">stato ordine</th>
					</tr>
					<% while not rsp.eof %>
					<tr>
						<td class="content"><%= rsp("mosp_nome_it") %></td>
						<td class="content"><%= rsp("pag_data") %></td>
						<td class="content"><%= FormatPrice(cReal(rsp("pag_importo")), 2, true) %>&euro;</td>
						<td class="content"><%= rsp("so_nome_it") %></td>
					</tr>
					<%rsp.movenext
					wend %>
				</table>
				<% if rsp.recordcount>2 then %>
					</span>
				<% end if
				rsp.close %>
			</td>
		</tr>
		<%end if%>
	</table>
	<% 
	'......................................................................................................
	'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
	CALL ADDON__ORDINI__form_update(conn, rs)
	'......................................................................................................
	%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="4">ARTICOLI ORDINATI</th></tr>
		<tr>
			<td colspan="4">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% 	sql = " SELECT * FROM gv_dettagli_ord LEFT JOIN tb_indirizzario ON gv_dettagli_ord.det_ind_id=tb_indirizzario.IDElencoIndirizzi "
					if cInteger(rs("riv_LstCod_id"))>0 then
						'aggiunge codifica dei codici
						sql = sql + " LEFT JOIN gtb_codici ON (gv_dettagli_ord.rel_id = gtb_codici.cod_variante_id AND cod_lista_id=" & rs("riv_LstCod_id") &") "
					end if
					sql = sql + " WHERE det_ord_id="& rs("ord_id") & _
							    " ORDER BY cntRel, IndirizzoElencoIndirizzi, (CASE WHEN ISNULL(det_art_var_id, 0) > 0 THEN 0 ELSE 1 END), art_nome_it, rel_cod_int"
					rsd.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
					%>
					<%
					dim modificabile
					modificabile = NOT rs("ord_archiviato") AND NOT rs("ord_movimenta")
					
					if cBoolean(rs("ord_imported"), false) then
						modificabile = true
					end if
					%>
					<tr>
						<td class="label" style="width:84%;" colspan="10">
							<% if rsd.eof then %>
								nessun dettaglio definito per l'ordine.
							<% else %>
								n&ordm; <%= rsd.recordcount %> dettagli per l'ordine.
							<% end if %>
						</td>
						<% if modificabile then %>
							<td colspan="3" class="content_right" style="padding-right:0px;">
								<% if NOT cBoolean(rs("ord_imported"), false) then %>
									<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ArticoliSelPrz.asp?ORD_ID=<%= request("ID") %>', 'OrdiniDettagli', 640, 580, true)"
									   title="apre la finestra per l'inserimento di un nuovo dettaglio dell'ordine" <%= ACTIVE_STATUS %>>
										NUOVO DETTAGLIO
									</a>
								<% end if %>
							</td>
						<% end if %>
					</tr>
					<% if not rsd.eof then %>
						<tr>
							<th class="l2_center" colspan="2" style="border-bottom:0px; width:8%">codici</th>
							<th class="L2" rowspan="2" style="width:60%">ARTICOLO</th>
							<th class="l2_center" style="border-bottom:0px; width:7%" colspan="2">QUANTIT&Agrave;</th>
							<th class="l2_center" colspan="3" style="border-bottom:0px; width:12%">PREZZI IN &euro;</th>
							<th class="l2_center" colspan="2" style="border-bottom:0px; width:8%">SPESE IN &euro;</th>
							<% if modificabile then %>
								<th class="l2_center" rowspan="2" colspan="2" style="width:5%">OPERAZIONI</th>
							<% end if %>
						</tr>
						<tr>
							<th class="L2">INTERNO</th>
							<% if cInteger(rs("riv_LstCod_id"))>0 then %>
								<th class="L2">CLIENTE</th>
							<% else %>
								<th class="L2">PRODUTTORE</th>
							<% end if %>
							<th class="l2_center">ORDINATA</th>
							<th class="l2_center">EVASA</th>
							<th class="l2_center">UNITARIO</th>
							<th class="l2_center">TOTALE</th>
							<th class="l2_center">I.V.A.</th>
							<th class="l2_center">TOTALE</th>
							<th class="l2_center">I.V.A.</th>
						</tr>
						<% dim row_tot, row_iva, ind_id, totale, iva, ind_totale, ind_iva, indirizzi_diversi, rowspan_note, piede_ultima_destinazione,_
						tot_colli, tot_peso_netto, tot_peso_lordo, tot_volume
						ind_id = 0
						totale = 0
						iva = 0
						ind_totale = 0
						ind_iva = 0
						indirizzi_diversi = false
						tot_colli = 0
						tot_peso_netto = 0
						tot_peso_lordo = 0
						tot_volume = 0

						while not rsd.eof 
							if ind_id <> rsd("det_ind_id") then
								if ind_id>0 then%>
									<tr>
										<td class="label_right" colspan="6">totale per destinazione</td>
										<td class="content_right"><%= FormatPrice(ind_totale, 2, true) %></td>
										<td class="content_right"><%= FormatPrice(ind_iva, 2, true) %></td>
										<% if modificabile then %><td class="content" colspan="2">&nbsp;</td><% end if %>
									</tr>
									<%ind_totale = 0
									ind_iva = 0
									indirizzi_diversi = true
								end if %>
								<tr>
									<td class="header_L2" colspan="12">
										<% if cInteger(rsd("cntRel"))=0 then %>
											sede principale
										<% else %>
											<%= rsd("NomeOrganizzazioneElencoIndirizzi") %>
										<% end if %>
										 - 
										<%= ContactAddress(rsd) %>
									</td>
								</tr>
								<% ind_id = rsd("det_ind_id")
							end if 
							row_tot = rsd("det_prezzo_unitario") * rsd("det_qta")
							row_iva = GetIva(row_tot, rsd("det_iva"))
							
							rowspan_note = IIF(cString(rsd("det_note"))<>"", 2, 0)
							if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then
								rowspan_note = rowspan_note + 1
							end if
							
							if rowspan_note > 0 then
								rowspan_note = "rowspan=""" & rowspan_note & """ style=""vertical-align:middle;"""
							end if
							
							if piede_ultima_destinazione <> ind_id AND cInteger(rsd("rel_id"))=0 then
								piede_ultima_destinazione = ind_id%>
								<tr>
									<td class="header_L3" colspan="12">sconti / costi aggiuntivi</td>
								</tr>
							<% end if %>
							<tr>
								<% IF cInteger(rsd("rel_id"))>0 then %>
									<td class="content">
										<%= rsd("rel_cod_int") %>
									</td>
									<% if cInteger(rs("riv_LstCod_id"))>0 then %>
										<td class="content">
											<% if NOT isNull(rsd("cod_codice")) then %>
												<%= rsd("cod_codice") %>
											<% else %>
												<%= rsd("rel_cod_pro") %>
											<% end if %>
										</td>
									<% else %>
										<td class="content"><%= rsd("rel_cod_pro") %></td>
									<% end if %>
									<td class="content">
										<% CALL ArticoloLink(rsd("art_id"), rsd("art_nome_it"), rsd("rel_cod_int")) %>
										<% if rsd("art_varianti") then %>
											<%= ListValoriVarianti(conn, rsi, rsd("rel_id")) %>
										<% end if %>
										<% if cString(rsd("det_cod_promozione"))<>"" then %>
											<span class="Icona Promozioni" title="dettaglio d'ordine a cui viene applicata la promozione &quot;<%= rsd("det_cod_promozione") %>&quot;">&nbsp;</span>
										<% end if %>
									</td>
									<td class="content_center">
										<% if cIntero(rsd("det_qta"))<>0 then %>
											<%= rsd("det_qta") %>
										<% else %>
											&nbsp;
										<% end if %>
									</td>
									<td class="content_center">
										<% if cIntero(rsd("det_qta_evasa"))<>0 then %>
											<%= rsd("det_qta_evasa") %>
										<% else %>
											&nbsp;
										<% end if %>
									</td>
									<td class="content_right <%= StiliCampoTestoAZero(rsd("det_prezzo_unitario"))%>"
										title="Prezzo di listino: <%= FormatPrice(rsd("det_prezzo_listino"), 2, true) & "&euro;" & vbCrLf%> variazione / sconto applicati: <%= FormatPrice(rsd("det_sconto"), 2, true) %>%">
										<%= FormatPrice(rsd("det_prezzo_unitario"), 2, true) %>
									</td>
									<td class="content_right <%=StiliCampoTestoAZero(rsd("det_totale"))%>"
										title="Prezzo di listino: <%= FormatPrice(cReal(rsd("det_totale")), 2, true) & "&euro;" & vbCrLf%> variazione / sconto applicati: <%= FormatPrice(rsd("det_sconto"), 2, true) %>%">
										<%= FormatPrice(cReal(rsd("det_totale")), 2, true) %>
									</td>
									<td class="content_right <%=StiliCampoTestoAZero(rsd("det_totale_iva"))%>"
										title="iva: <%= IIF(rsd("det_iva")>0, FormatPrice(rsd("det_iva"), 2, true) & "%", "esente iva") %>">
										<%= FormatPrice(cReal(rsd("det_totale_iva")),2, true) %>
									</td>
								<% else %>
									<td class="content" colspan="4">
										<%=TextEncode(rsd("det_descr_it"))%>
									</td>
									<td class="content_center">
										<% if cIntero(rsd("det_qta"))<>0 then %>
											<%= rsd("det_qta") %>
										<% else %>
											&nbsp;
										<% end if %>
									</td>
									<td class="content_center">
										<% if cIntero(rsd("det_qta_evasa"))<>0 then %>
											<%= rsd("det_qta_evasa") %>
										<% else %>
											&nbsp;
										<% end if %>
									</td>
									<td class="content_right <%=StiliCampoTestoAZero(rsd("det_totale"))%>">
										<%= FormatPrice(cReal("det_totale"), 2, true) %>
									</td>
									<td class="content_right <%=StiliCampoTestoAZero(rsd("det_totale_iva"))%>">
										<%= FormatPrice(cReal("det_totale_iva"),2, true) %>
									</td>
								<% end if %>
									<td class="content_right <%=StiliCampoTestoAZero(rsd("det_totale_spese"))%>"
											title="Spese di spedizione: <%= FormatPrice(cReal(rsd("det_spesespedizione")), 2, true) & "&euro;" & vbCrLf%>Spese di incasso: <%= FormatPrice(cReal(rsd("det_speseincasso")), 2, true) & "&euro;" & vbCrLf%>Spese fisse: <%= FormatPrice(cReal(rsd("det_spesefisse")), 2, true) & "&euro;" & vbCrLf%>Altre spese: <%= FormatPrice(cReal(rsd("det_spesealtre")), 2, true) & "&euro;" & vbCrLf%>">
											<%= FormatPrice(cReal(rsd("det_totale_spese")), 2, true) %>
									</td>
									<td class="content_right <%=StiliCampoTestoAZero(rsd("det_totale_spese_iva"))%>"
											title="Iva spese di spedizione: <%= cReal(rsd("det_spesespedizione")) & "%" & vbCrLf%>Iva spese di incasso: <%= cReal(rsd("det_speseincasso")) & "%" & vbCrLf%>Iva spese fisse: <%= cReal(rsd("det_spesefisse")) & "%" & vbCrLf%>Iva altre spese: <%= cReal(rsd("det_spesealtre")) & "%" & vbCrLf%>">
											<%= FormatPrice(cReal(rsd("det_totale_spese_iva")),2, true) %>
									</td>
								<% if modificabile then %>
									<td class="content_center" <%= rowspan_note %>>
										<% if cInteger(rsd("rel_id"))>0 then %>
											<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('Ordini_dettagli_art_Mod.asp?ID=<%= rsd("det_id") %>', '_blank', 640, 520, true)">
												MODIFICA
											</a>
										<% else %>
											<a class="button_L2_disabled" title="modifica non permessa">MODIFICA</a>
										<% end if %>
									</td>
									<td class="content_center" <%= rowspan_note %>>
										<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('DETORD','<%= rsd("det_id") %>');">
											CANCELLA
										</a>
									</td>
								<% end if %>
							</tr>
							<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
								<% 
								tot_colli = tot_colli + rsd("det_tot_colli")
								tot_peso_netto = tot_peso_netto + rsd("det_tot_peso_netto")
								tot_peso_lordo = tot_peso_lordo + rsd("det_tot_peso_lordo")
								tot_volume = tot_volume + rsd("det_tot_volume")
								%>
								<tr>
									<td class="note content_right">&nbsp;</td>
									<td class="note content_right" colspan="9" style="padding-top:0px;">
										colli: <%= rsd("det_tot_colli") %>, 
										peso netto: <%= rsd("det_tot_peso_netto") %>, 
										peso lordo: <%= rsd("det_tot_peso_lordo") %>, 
										volume: <%= rsd("det_tot_volume") %>
									</td>
								</tr>
							<% end if %>
							<% if cString(rsd("det_note"))<>"" then%>
								<tr>
									<td class="content">&nbsp;</td>
									<td class="content" colspan="9"><span class="label">note del dettaglio: </span><span class="note"><%= rsd("det_note") %></span></td>
								</tr>
							<%end if
							
							ind_totale = ind_totale + row_tot
							ind_iva = ind_iva + row_iva
							rsd.MoveNext
						wend
						
						if indirizzi_diversi then %>
							<tr>
								<td class="label_right" colspan="6">totale per destinazione</td>
								<td class="content_right"><%= FormatPrice(ind_totale, 2, true) %></td>
								<td class="content_right"><%= FormatPrice(ind_iva, 2, true) %></td>
								<% if modificabile then %><td class="content" colspan="2">&nbsp;</td><% end if %>
							</tr>
						<% end if %>
						<tr>
							<td class="header_L2" colspan="12">
								totali generali ordine
							</td>
						</tr>						
						<tr>
							<td class="label_right" colspan="6">subtotali</td>
							<td class="content_right <%= StiliCampoTestoAZero(rs("ord_totale"))%>"><%= FormatPrice(cReal(rs("ord_totale")), 2, true) %></td>
							<td class="content_right <%= StiliCampoTestoAZero(rs("ord_totale_iva"))%>"><%= FormatPrice(cReal(rs("ord_totale_iva")), 2, true) %></td>
							<td class="content_right <%= StiliCampoTestoAZero(rs("ord_det_totale_spese"))%>"><%= FormatPrice(cReal(rs("ord_det_totale_spese")), 2, true) %></td>
							<td class="content_right <%= StiliCampoTestoAZero(rs("ord_det_totale_spese_iva"))%>"><%= FormatPrice(cReal(rs("ord_det_totale_spese_iva")), 2, true) %></td>
							<% if modificabile then %><td class="content" colspan="4">&nbsp;</td><% end if %>
						</tr>
						<tr>
							<% dim spese_spedizione, spese_spedizione_iva
							   spese_spedizione = FormatPrice(cReal(rs("ord_spesespedizione")), 2, true)
							   spese_spedizione_iva = FormatPrice(GetIva(cReal(rs("ord_spesespedizione")), cReal(rs("ord_spesespedizione_iva"))), 2, true)
							%>
							<td class="label_right" colspan="6">spese spedizione</td>
							<td class="content" colspan="2">&nbsp;</td>
							<td class="content_right <%= StiliCampoTestoAZero(spese_spedizione)%>"><%= spese_spedizione%></td>
							<td class="content_right <%= StiliCampoTestoAZero(spese_spedizione_iva)%>" title="Iva: <%=cReal(rs("ord_spesespedizione_iva"))%>%"><%= spese_spedizione_iva%></td>
							<% if modificabile then %><td class="content" colspan="4">&nbsp;</td><% end if %>
						</tr>
						<tr>
							<% dim spese_fisse, spese_fisse_iva
							   spese_fisse = FormatPrice(cReal(rs("ord_spesefisse")), 2, true)
							   spese_fisse_iva = FormatPrice(GetIva(cReal(rs("ord_spesefisse")), cReal(rs("ord_spesefisse_iva"))), 2, true)
							%>
							<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
								<td colspan="3" rowspan="5">
									<table cellspacing="0" cellpadding="0" border="0" style="width:100%;">
										<tr>
											<th class="l2" colspan="2">TOTALE COLLI, PESO E VOLUME</th>
										</tr>
										<tr>
											<td style="width:40%" class="label">totale colli</td>
											<td class="content"><%=tot_colli%></td>
										</tr>
										<tr>
											<td class="label">totale peso netto</td>
											<td class="content"><%=tot_peso_netto%></td>
										</tr>
										<tr>
											<td class="label">totale peso lordo</td>
											<td class="content"><%=tot_peso_lordo%></td>
										</tr>
										<tr>
											<td class="label">totale volume</td>
											<td class="content"><%=tot_volume%></td>
										</tr>
									</table>
								</td>
								<td class="label_right" colspan="3">spese fisse</td>
							<% else %>
								<td class="label_right" colspan="6">spese fisse</td>
							<% end if %>
							<td class="content" colspan="2">&nbsp;</td>
							<td class="content_right <%= StiliCampoTestoAZero(spese_fisse)%>"><%= spese_fisse %></td>
							<td class="content_right <%= StiliCampoTestoAZero(spese_fisse_iva)%>" title="Iva: <%=cReal(rs("ord_spesefisse_iva"))%>%"><%= spese_fisse_iva %></td>
							<% if modificabile then %><td class="content" colspan="4">&nbsp;</td><% end if %>
						</tr>
						<tr>
							<% dim spese_incasso, spese_incasso_iva
							   spese_incasso = FormatPrice(cReal(rs("ord_speseincasso")), 2, true)
							   spese_incasso_iva = FormatPrice(GetIva(cReal(rs("ord_speseincasso")), cReal(rs("ord_speseincasso_iva"))), 2, true)
							%>
							<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
								<td class="label_right" colspan="3">spese incasso</td>
							<% else %>
								<td class="label_right" colspan="6">spese incasso</td>
							<% end if %>
							<td class="content" colspan="2">&nbsp;</td>
							<td class="content_right <%= StiliCampoTestoAZero(spese_incasso)%>"><%= spese_incasso %></td>
							<td class="content_right <%= StiliCampoTestoAZero(spese_incasso_iva)%>" title="Iva: <%=cReal(rs("ord_speseincasso_iva"))%>%"><%= spese_incasso_iva %></td>
							<% if modificabile then %><td class="content" colspan="4">&nbsp;</td><% end if %>
						</tr>
						<tr>
							<% dim altre_spese, altre_spese_iva
							   altre_spese = FormatPrice(cReal(rs("ord_spesealtre")), 2, true)
							   altre_spese_iva = FormatPrice(GetIva(cReal(rs("ord_spesealtre")), cReal(rs("ord_spesealtre_iva"))), 2, true)
							%>
							<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
								<td class="label_right" colspan="3">altre spese</td>
							<% else %>
								<td class="label_right" colspan="6">altre spese</td>
							<% end if %>
							<td class="content" colspan="2">&nbsp;</td>
							<td class="content_right <%= StiliCampoTestoAZero(altre_spese)%>"><%= altre_spese %></td>
							<td class="content_right <%= StiliCampoTestoAZero(altre_spese_iva)%>" title="Iva: <%=cReal(rs("ord_spesealtre_iva"))%>%"><%= altre_spese_iva %></td>
							<% if modificabile then %><td class="content" colspan="4">&nbsp;</td><% end if %>
						</tr>
						<tr>
							<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
								<td class="label_right" colspan="3">subtotali</td>
							<% else %>
								<td class="label_right" colspan="6">subtotali</td>
							<% end if %>
							<td class="content_right <%= StiliCampoTestoAZero(rs("ord_totale"))%>"><%= FormatPrice(cReal(rs("ord_totale")), 2, true) %></td>
							<td class="content_right <%= StiliCampoTestoAZero(rs("ord_totale_iva"))%>"><%= FormatPrice(cReal(rs("ord_totale_iva")), 2, true) %></td>
							<td class="content_right <%= StiliCampoTestoAZero(rs("ord_totale_spese"))%>"><%= FormatPrice(cReal(rs("ord_totale_spese")), 2, true) %></td>
							<td class="content_right <%= StiliCampoTestoAZero(rs("ord_totale_spese_iva"))%>"><%= FormatPrice(cReal(rs("ord_totale_spese_iva")), 2, true) %></td>
							<% if modificabile then %><td class="content" colspan="4">&nbsp;</td><% end if %>
						</tr>
						<tr>
							<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
								<td class="label_right" colspan="3">totale ordine</td>
							<% else %>
								<td class="label_right" colspan="6">totale ordine</td>
							<% end if %>
							<td class="content_right" colspan="4"><strong><%= FormatPrice(cReal(rs("ord_totale"))+cReal(rs("ord_totale_spese")) + cReal(rs("ord_totale_iva"))+cReal(rs("ord_totale_spese_iva")) , 2, true) %>&nbsp;&euro; </strong></td>
							<% if modificabile then %><td class="content" colspan="4">&nbsp;</td><% end if %>
						</tr>
						<% 
						'......................................................................................................
						'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
						CALL ADDON__ORDINI__form_update_piede(conn, rs, rsd)
						'......................................................................................................
						%>
					<%end if
					rsd.close%>
				</table>
			</td>
		</tr>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_ord_note"><%= rs("ord_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="label">data inserimento:</td>
			<td class="content"><%= DateTimeIta(rs("ord_data_ins")) %></td>
			<td class="label">ultima modifica:</td>
			<td class="content"><%= DateTimeIta(rs("ord_data_ultima_mod")) %></td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" style="width:23%;" class="button" name="mod" value="SALVA & TORNA ALL'ELENCO">
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
	 VerificaStatoOrdine();
</script>
<%
rs.close
set rs = nothing
set rsd = nothing
set rsi = nothing
conn.Close
set conn = nothing
%>