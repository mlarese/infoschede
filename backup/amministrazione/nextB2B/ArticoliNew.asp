<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ArticoliSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, tipo
if request("TYPE")="B" then
	tipo = "bundle"
elseif request("TYPE")="C" then
	tipo = "confezione"
elseif request("TYPE")="AV" then
	tipo ="articolo con varianti"
else
	tipo ="articolo singolo"
end if
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione articoli - nuovo " & tipo
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Articoli.asp"
dicitura.scrivi_con_sottosez()

dim conn, rs, rsv, sql, i, rs_spe
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")
set rs_spe = Server.CreateObject("ADODB.Recordset")
%>
<div id="content">
	<form action="ArticoliNew.asp?external=<%= request("external") %>&type=<%= request("TYPE") %>" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_art_applicativo_id" value="<%= NEXTB2B %>">
	<input type="hidden" name="tfn_art_se_accessorio" value="0">
	<input type="hidden" name="tfn_art_ha_accessori" value="0">
	<input type="hidden" name="tfn_art_in_bundle" value="0">
	<input type="hidden" name="tfn_art_in_confezione" value="0">
	<input type="hidden" name="tfn_art_se_bundle" value="<%= IIF(request("TYPE")="B", "1", "0") %>">
	<input type="hidden" name="tfn_art_se_confezione" value="<%= IIF(request("TYPE")="C", "1", "0") %>">
	<input type="hidden" name="tfn_art_varianti" value="<%= IIF(request("TYPE")="AV", "1", "0") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Inserimento nuovo <%= tipo %></caption>
		<tr><th colspan="7">DATI PRINCIPALI</th></tr>
		<% if request("TYPE")<>"AV" then 
			 sql = "SELECT * FROM gtb_lista_codici WHERE lstCod_sistema=1 ORDER BY lstCod_nome" 
			rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
			<tr>
				<td class="label" style="width:16%;"<% if rsv.recordcount>0 then %>rowspan="<%= 1+rsv.recordcount %>"<% end if %>>codici:</td>
				<td class="label" style="width:8%;">interno:</td>
				<td class="content" style="width:22%;">
					<input type="text" class="text" name="tft_art_cod_int" value="<%= request("tft_art_cod_int") %>" maxlength="50" size="15">
					(*)
				</td>
				<td class="label" style="width:8%;">alternativo:</td>
				<td class="content">
					<input type="text" class="text" name="tft_art_cod_alt" value="<%= request("tft_art_cod_alt") %>" maxlength="50" size="15">
				</td>
				<td class="label" style="width:8%;">produttore:</td>
				<td class="content">
					<input type="text" class="text" name="tft_art_cod_pro" value="<%= request("tft_art_cod_pro") %>" maxlength="50" size="15">
				</td>
			</tr>
			<% while not rsv.eof %>
				<tr>
					<td class="label_no_width" colspan="2">
						<%= rsv("lstCod_nome") %>
					</td>
					<td class="content" colspan="4">
						<input type="text" class="text" name="codice_articolo_<%= rsv("lstCod_id") %>" value="<%= request("codice_articolo_" & rsv("lstCod_id")) %>" maxlength="50" size="23">
					</td>
				</tr>
				<% rsv.movenext
			wend 
			rsv.close
		end if
		for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
				<td class="content" colspan="6">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_art_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_art_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">categoria:</td>
			<td class="content" colspan="6">
				<%CALL dropDown(conn, categorie.QueryElenco(false, ""), "tip_id", "NAME", "tfn_art_tipologia_id", request("tfn_art_tipologia_id"), false, " onchange='form1.submit()'", LINGUA_ITALIANO)%>
				(*)
			</td>
		</tr>
		<% sql = "SELECT COUNT(*) FROM gtb_tipologie_raggruppamenti"
		if cIntero(getValueList(conn, rsv, sql))>0 then %>
			<tr>
				<td class="label">&nbsp;</td>
				<td class="label" colspan="2">raggruppamento di pubblicazione:</td>
				<td class="content" colspan="5">
					<% if cInteger(request("tfn_art_tipologia_id"))>0 then
						sql = " SELECT * FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=" & request("tfn_art_tipologia_id")
						rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						if rsv.eof then %>
							<span class="note">Nessun raggruppamento disponibile per questa categoria di prodotti</span>
							<input type="hidden" name="nfn_art_raggruppamento_id" value="NULL">
						<% else
							CALL DropDownRecordset(rsv, "rag_id", "rag_nome_it", "nfn_art_raggruppamento_id", request("nfn_art_raggruppamento_id"), false, "", LINGUA_ITALIANO)
						end if
						rsv.close
					end if %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">marchio / produttore:</td>
			<td class="content" colspan="6">
				<%CALL dropDown(conn, "SELECT mar_id, mar_nome_it FROM gtb_marche ORDER BY mar_nome_it", _
							    "mar_id", "mar_nome_it", "tfn_art_marca_id", request("tfn_art_marca_id"), true, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<tr><th colspan="7">DATI PER LA GESTIONE</th></tr>
		<tr>
			<td class="label" colspan="2">non a catalogo:</td>
			<% if request("TYPE")="B" OR request("TYPE")="C" then %>	
				<input type="hidden" name="chk_art_disabilitato" value="1">
				<td class="content"><input type="checkbox" class="checkbox" disabled checked></td>
				<td class="note" colspan="1">
					Sar&agrave; possibile mettere a catalogo <%= IIF(request("TYPE")="B", "il", "la") %>&nbsp;<%= tipo %> dopo aver selezionato i relativi componenti.
				</td>
			<% else %>
				<td class="content" colspan="2"><input type="checkbox" class="checkbox" name="chk_art_disabilitato" <%= chk(request("chk_art_disabilitato")<>"") %>></td>
			<% end if %>
			<td class="label" colspan="2">ordine di pubblicazione:</td>
			<td class="content" colspan="1"><input type="text" class="text" name="tfn_art_ordine" value="<%= request("tfn_art_ordine") %>" size="7"></td>
		</tr>
		<tr>
			<td class="label" colspan="2">non vendibile singolarmente:</td>
			<td class="content" colspan="5"><input type="checkbox" class="checkbox" name="chk_art_NoVenSingola" <%= chk(request("chk_art_NoVenSingola")<>"") %>></td>
		</tr>
		<tr>
			<td class="label" colspan="2">pezzo unico:</td>
			<td class="content" colspan="5"><input type="checkbox" class="checkbox" name="chk_art_unico" <%= chk(request("chk_art_unico")<>"") %>></td>
		</tr>
		<% if request("TYPE")<>"AV" then %>
			<tr>
				<td class="label" colspan="2">prezzo base:</td>
				<td class="content" colspan="5">
					<input type="text" class="number" name="tfn_art_prezzo_base" value="<%= FormatPrice(cReal(request("tfn_art_prezzo_base")), 2, false) %>" size="7"> &euro;
					<span style="padding-left:5px;">(*)</span>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" colspan="2">categoria i.v.a.:</td>
			<td class="content" colspan="5">
				<% sql = "SELECT * FROM gtb_iva ORDER BY iva_ordine"
				CALL dropDown(conn, sql, "iva_id", "iva_nome", "tfn_art_iva_id", request("tfn_art_iva_id"), true, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<% sql = "SELECT spa_id, spa_nome_it  FROM gtb_spese_spedizione_articolo ORDER BY spa_id"
		   rs_spe.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		if rs_spe.RecordCount = 1 then %>
			<input type="hidden" name="tfn_art_spedizione_id" value="<%= GetValueList(conn, , "SELECT TOP 1 spa_id  FROM gtb_spese_spedizione_articolo ORDER BY spa_id") %>"">
		<% else %>
			<tr>
				<td class="label" colspan="2" nowrap>Metodo di spedizione:</td>
				<td class="content" colspan="5">
					<%CALL dropDown(conn, sql, "spa_id", "spa_nome_it", "tfn_art_spedizione_id", request("tfn_art_spedizione_id"), true, "", LINGUA_ITALIANO)%>
				</td>
			</tr>
		<% end if%>
		<% if request("TYPE")="AV" then %>
			<tr><th colspan="7">DATI PER LA GENERAZIONE DELLE VARIANTI</th></tr>
			<tr><td class="note" colspan="7">I dati immessi nei seguenti campi verranno utilizzati per generare le varianti dell'articolo.</td></tr>
			<tr>
				<td class="label" style="width:16%;">radici del codice:</td>
				<td class="label" style="width:8%;">interno:</td>
				<td class="content" style="width:17%;">
					<input type="text" class="text" name="tft_art_cod_int" value="<%= request("tft_art_cod_int") %>" maxlength="50" size="10">
					<span id="art_cod_int">(*)</span>
				</td>
				<td class="label" style="width:8%;">alternativo:</td>
				<td class="content">
					<input type="text" class="text" name="tft_art_cod_alt" value="<%= request("tft_art_cod_alt") %>" maxlength="50" size="10">
				</td>
				<td class="label" style="width:8%;">produttore:</td>
				<td class="content">
					<input type="text" class="text" name="tft_art_cod_pro" value="<%= request("tft_art_cod_pro") %>" maxlength="50" size="10">
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">prezzo base delle varianti:</td>
				<td class="content" colspan="5">
					<input type="text" class="number" name="tfn_art_prezzo_base" value="<%= FormatPrice(cReal(request("tfn_art_prezzo_base")), 2, false) %>" size="7"> &euro;
					<span id="art_prezzo_base" style="padding-left:5px;">(*)</span>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" colspan="2" nowrap>classe di sconto per quantit&agrave;:</td>
			<td class="content" colspan="5">
				<%CALL dropDown(conn, "SELECT scc_id, scc_nome FROM gtb_scontiQ_classi ORDER BY scc_nome", _
							    "scc_id", "scc_nome", "nfn_art_scontoQ_id", request("nfn_art_scontoQ_id"), false, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<tr>
			<td class="label" colspan="2">giacenza minima:</td>
			<td class="content"><input type="text" class="text" name="tfn_art_giacenza_min" value="<%= IIF(request("tfn_art_giacenza_min")<>"", request("tfn_art_giacenza_min"), "1") %>" size="7"></td>
			<td class="note" colspan="4">Limite minimo di quantit&agrave; di prodotto per la segnalazione dello stato "in esaurimento"</td>
		</tr>
		<tr>
			<td class="label" colspan="2">quantit&agrave; minima ordinabile:</td>
			<td class="content" colspan="5"><input type="text" class="text" name="tfn_art_qta_min_ord" value="<%= IIF(request("tfn_art_qta_min_ord")<>"", request("tfn_art_qta_min_ord"), "1") %>" size="7"></td>
		</tr>
		<tr>
			<td class="label" colspan="2">quantit&agrave; massima ordinabile:</td>
			<td class="content" colspan="5"><input type="text" class="text" name="tfn_art_qta_max_ord" value="<%=request("tfn_art_qta_max_ord")%>" size="7"></td>
		</tr>
		<tr>
			<td class="label" colspan="2">lotto di riordino:</td>
			<td class="content"><input type="text" class="text" name="tfn_art_lotto_riordino" value="<%= IIF(request("tfn_art_lotto_riordino")<>"", request("tfn_art_lotto_riordino"), "1") %>" size="7"></td>
			<td class="note" colspan="4">Indica il numero di articoli che compongono il lotto ordinato dal cliente.</td>
		</tr>
		<% if request("TYPE")<>"AV" OR TRUE then %>
			<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
				<tr><th colspan="7">COLLI, PESI E DIMENSIONI</th></tr>
				<tr>
					<td class="label" colspan="2">peso netto:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="extN_rel_peso_netto" value="<%= request("extN_rel_peso_netto") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">peso lordo:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="extN_rel_peso_lordo" value="<%= request("extN_rel_peso_lordo") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">numero colli:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="extN_rel_colli_num" value="<%= IIF(request("extN_rel_colli_num")<>"",request("extN_rel_colli_num"),"1") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">numero pezzi per collo:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="extN_rel_collo_pezzi_per" value="<%= IIF(request("extN_rel_collo_pezzi_per")<>"",request("extN_rel_collo_pezzi_per"),"1") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">larghezza collo:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="extN_rel_collo_width" value="<%= request("extN_rel_collo_width") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">altezza collo:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="extN_rel_collo_height" value="<%= request("extN_rel_collo_height") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">lunghezza collo:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="extN_rel_collo_lenght" value="<%= request("extN_rel_collo_lenght") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">volume collo:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="extN_rel_collo_volume" value="<%= request("extN_rel_collo_volume") %>" size="7"></td>
				</tr>
			<% else %>
				<input type="hidden" name="extN_rel_peso_netto" value="0">
				<input type="hidden" name="extN_rel_peso_lordo" value="0">
				<input type="hidden" name="extN_rel_colli_num" value="1">
				<input type="hidden" name="extN_rel_collo_pezzi_per" value="1">
				<input type="hidden" name="extN_rel_collo_width" value="0">
				<input type="hidden" name="extN_rel_collo_height" value="0">
				<input type="hidden" name="extN_rel_collo_lenght" value="0">
				<input type="hidden" name="extN_rel_collo_volume" value="0">
			<% end if %>
		<% end if %>
		
		<% if request("TYPE")="AV" then %>
			<tr><th colspan="7">VARIANTI DELL'ARTICOLO</th></tr>
			<% sql = " SELECT * FROM gtb_varianti INNER JOIN gtb_valori ON gtb_varianti.var_id = gtb_valori.val_var_id " + _
					 " ORDER BY var_nome_it, val_nome_it"
			rsv.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
			if rsv.eof then%>
				<tr><td class="content_b alert" colspan="7">Nessuna variante definita.<br>Prima di inserire l'articolo inserire le varianti ed i relativi valori.</th></tr>
				<tr>
					<td class="footer" colspan="7">
						<a class="button" href="Articoli.asp">INDIETRO</a>
					</td>
				</tr>
			</table>
				<%response.end
			else
				dim Current%>
				<script language="JavaScript" type="text/javascript">
					function SelectVariante(variante, var_id){
						var value = "valore_" + var_id
						//ciclo su tutti gli elementi del form per trovare i check dei valori della variante
						for (var i=0; i<form1.elements.length; i++){
							if (!form1.elements[i].id.indexOf(value)){
								//imposta il valore dei check a quello della variante
								form1.elements[i].checked = variante.checked;
							}
						}
						CampiObbligatoriGenerazioneVarianti()
					}
					
					function SelectValore(valore, var_id){
						var variante = document.getElementById('variante_' + var_id);
						if (!valore.checked){
							variante.checked = false;
						}
						else{
							//verifica anche tutti gli altri checkbox
							var value = "valore_" + var_id;
							var toCheck = true;
							for (var i=0; i<form1.elements.length; i++){
								if (!form1.elements[i].id.indexOf(value)){
									if (!form1.elements[i].checked){
										toCheck = false;
										i = form1.elements.length + 1;
									}
								}
							}
							variante.checked = toCheck;
						}
						CampiObbligatoriGenerazioneVarianti()
					}
					
					function CampiObbligatoriGenerazioneVarianti(){
						//verifica se esiste almeno uno dei valori delle varianti selezionati
						var value = "valore_";
						var obbligatori = false;
						
						for (var i=0; i<form1.elements.length; i++){
							if (!form1.elements[i].id.indexOf(value)){
								if (form1.elements[i].checked){
									obbligatori = true;
									i = form1.elements.length + 1;
								}
							}
						}
						
						var art_cod_int = document.getElementById('art_cod_int');
						var art_prezzo_base = document.getElementById('art_prezzo_base');
						
						if (obbligatori){
							art_cod_int.innerHTML = '(*)';
							art_prezzo_base.innerHTML = '(*)';
						}
						else {
							art_cod_int.innerHTML = '';
							art_prezzo_base.innerHTML = '';
						}
						
					}
				</script>
				<tr>
					<td colspan="5">
						<span class="overflow" style="height:150px;">
							<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
								<tr>
									<th class="L2">VARIANTE</th>
									<th class="L2" width="4%">&nbsp;</th>
									<th class="L2" width="8%">CODICE</th>
									<th class="L2">VALORE</th>
								</tr>
								<%Current = ""
								while not rsv.eof %>
									<tr>
										<% if Current <> rsv("var_id") then
											Current = rsv("var_id") %>
											<td class="content">
												<input type="checkbox" class="checkbox" id="variante_<%= rsv("var_id") %>" name="varianti" value=" <%= rsv("var_id") %> " <%= chk(instr(1, request("varianti"), " " & rsv("var_id") & " ", vbTextCompare)) %> onclick="SelectVariante(this, '<%= rsv("var_id") %>')">
												<%= rsv("var_nome_it") %>
											</td>
										<% else %>	
											<td class="content">&nbsp;</td>
										<% end if %>
										<td class="content"><input type="checkbox" class="checkbox" id="valore_<%= rsv("var_id") %>_<%= rsv("val_id") %>" name="valori" value=" <%= rsv("val_id") %> " <%= chk(instr(1, request("valori"), " " & rsv("val_id") & " ", vbTextCompare)) %>  onclick="SelectValore(this, '<%= rsv("var_id") %>')"></td>
										<td class="content"><%= rsv("val_cod_int") %></td>
										<td class="content"><%= rsv("val_nome_it") %></td>
									</tr>
									<%rsv.MoveNext
								wend%>
							</table>
						</span>
					</td>
					<td colspan="2" class="note" style="vertical-align:top;">
						<% if rsv.recordcount > 1 then %>
							Selezionare le varianti o una parte dei loro valori per i quali si vogliono generare le varianti dell'articolo. 
							L'eventuale selezione di altri valori, creazione di combinazioni di essi, la loro modifica o la loro 
							cancellazione &egrave; permessa anche dopo aver salvato.<br>
							Selezionando pi&ugrave; valori di pi&ugrave; varianti verr&ograve; creata una combinazione per ogni valore con 
							ogni valore delle altre varianti.<br>
							Es. selezionando 2 valori della variante "A" e 3 valori della variante "B" ne risulteranno 6 varianti per l'articolo.
						<% else %>
							Selezionare i valori della variante che descrivono le caratteristiche dell'articolo.
						<% end if %>
					</td>
				</tr>
				<script language="JavaScript" type="text/javascript">
				<!--
					CampiObbligatoriGenerazioneVarianti();
				//-->
				</script>
			<% end if
			rsv.close %>
		<% elseif request("TYPE")<>"AS" then %>
			<tr><th colspan="7">COMPOSIZIONE</th></tr>
			<tr><td class="note" colspan="7">Sar&agrave; possibile definire tutti i componenti e le loro quantit&agrave; dopo aver salvato.</td></tr>
		<% end if %>
		<tr><th colspan="7">DESCRIZIONE RIASSUNTIVA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="7">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="3" name="tft_art_descr_riassunto_<%= Application("LINGUE")(i) %>"><%= request("tft_art_descr_riassunto_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
		<tr><th colspan="7">DESCRIZIONE ESTESA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="7">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="4" name="tft_art_descr_<%= Application("LINGUE")(i) %>"><%= request("tft_art_descr_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
						<% 
						if Session("B2B_ABILITA_DESCRIZIONE_HTML") then 
							CALL activateCKEditor("tft_art_descr_"&Application("LINGUE")(i)) 
						end if 
						%>
					</table>
				</td>
			</tr>
		<%next %>
	</table>

	<% sql = " SELECT TOP 1 ct_id FROM gtb_carattech INNER JOIN gtb_tip_ctech ON gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id " & _
		  " WHERE rct_tipologia_id = " & cIntero(request("tfn_art_tipologia_id"))
	if cIntero(GetValueList(conn, NULL, sql)) > 0 then 
	%>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr><th colspan="7">CARATTERISTICHE TECNICHE</th></tr>
			<% if cInteger(request("tfn_art_tipologia_id"))>0 then 
				'sql = " SELECT *, ('') rel_ctech_it, ('') rel_ctech_en, ('') rel_ctech_de, ('') rel_ctech_fr, ('') rel_ctech_es " + _
				'	  " FROM gtb_carattech INNER JOIN gtb_tip_ctech ON gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id " + _
				'	  " WHERE rct_tipologia_id=" & request("tfn_art_tipologia_id")
				'CALL DesElenco(conn, sql, "gtb_carattech", "ct_id", "ct_nome_it", "ct_tipo", "ct_unita_it", "rel_ctech_", true, 7) 
				
				sql = " SELECT *" + _
					  " FROM gtb_carattech"& _
					  " INNER JOIN gtb_tip_ctech ON (gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id AND rct_tipologia_id=" & request("tfn_art_tipologia_id") & ")" + _
					  " LEFT JOIN grel_art_ctech ON (gtb_carattech.ct_id = grel_art_ctech.rel_ctech_id AND grel_art_ctech.rel_art_id=" & CInteger(request("ID")) & ")"& _
					  " LEFT JOIN gtb_carattech_raggruppamenti ON gtb_carattech.ct_raggruppamento_id = gtb_carattech_raggruppamenti.ctr_id " & _
					  " ORDER BY ctr_ordine, ctr_id, rct_ordine"
				CALL DesForm  (conn, sql, "gtb_carattech", "ct_id", "ct_nome_it", "ct_tipo", "ct_unita_it", "", "rel_ctech_", "rel_ctech_", "ctr_titolo_it", cIntero(request("ID")) = 0, 7)
				%>
			<% else %>
				<tr><td class="label" colspan="7">Per descrivere le caratteristiche tecniche dell'articolo selezionare prima la sua categoria.</td></tr>
			<% end if %>
		</table>
	<% end if %>
	
	
	<% 	CALL oArticoliFoto.Elenco(request("ID"), "FOTO") %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="7">ARTICOLI COLLEGATI</th></tr>
		<tr><td colspan="7" class="label">Sar&agrave; possibile definire gli articoli collegati dopo aver salvato.</td></tr>
		<tr><th colspan="7">NOTE INTERNE</th></tr>
		<tr>
			<td class="content" colspan="7">
				<textarea style="width:100%;" rows="3" name="tft_art_note"><%= request("tft_art_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="7">
				(*) Campi obbligatori.
				<input <%= Disable(cInteger(request("tfn_art_tipologia_id"))=0) %> type="submit" class="button" name="salva" value="SALVA &gt;&gt;">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<%
set rs = nothing
set rsv = nothing
conn.Close
set conn = nothing
%>