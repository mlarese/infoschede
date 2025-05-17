<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if (request("salva")<>"" OR request("salva_elenco")<>"") AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ArticoliSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<!--#INCLUDE FILE="Tools4Save_B2B.asp" -->

<% 	
dim conn, rs, rsv, rsp, sql, i, aux, txt, rs_spe
dim categoria

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")
set rs_spe = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_ARTICOLI_SQL"), "art_id", "ArticoliMod.asp")
end if

sql = " SELECT * FROM (gtb_articoli INNER JOIN gtb_iva ON gtb_Articoli.art_iva_id = gtb_iva.iva_id) " + _
	  " LEFT JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_Valori.rel_art_id " + _
	  " WHERE art_id=" & request("ID")
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

if not rs("art_varianti") then
	rs.close
	sql = " SELECT *, (SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_art_var_id = gv_articoli.rel_id) AS N_ORDINI " + _
		  " FROM gv_articoli WHERE art_id=" & request("ID")
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
end if


dim dicitura, tipo
if cBoolean(rs("art_se_bundle"),false) then
	tipo = "bundle"
elseif cBoolean(rs("art_se_confezione"),false) then
	tipo = "confezione"
elseif cBoolean(rs("art_varianti"),false) then
	tipo ="articolo con varianti"
else
	tipo ="articolo singolo"
end if


set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione articoli - modifica " & tipo
dicitura.puls_new = "INDIETRO;"
dicitura.link_new = "Articoli.asp;"
dicitura.puls_2a_riga.Add "PREZZI","ArticoliPrezzi.asp?ID=" & request("ID")
dicitura.puls_2a_riga.Add "GIACENZE","ArticoliGiacenze.asp?ID=" & request("ID")
if Session("ATTIVA_FAQ_ARTICOLI") then
	dicitura.puls_2a_riga.Add "FAQ","ArticoliFaq.asp?ID=" & request("ID")
end if
if Session("ATTIVA_COMMENTI") then
	dicitura.puls_2a_riga.Add "COMMENTI","ArticoliCommenti.asp?ID=" & request("ID")
end if

CALL dicitura.InitializeIndex(Index, "gtb_articoli", request("ID"))
dicitura.scrivi_con_sottosez()

%>

<div id="content_abbassato">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_art_varianti" value="<%= IIF(rs("art_varianti"), "1", "0") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati <%= tipo %> con codice &quot;<%= rs("art_cod_int") %>&quot;</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" name="<%= Server.HTMLEncode(rs("art_cod_int")) %>" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="articolo precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="articolo successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="7">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">tipo:</td>
			<% if rs("art_se_bundle") then %>
				<td class="content bundle" colspan="6">bundle</td>
			<% elseif rs("art_se_confezione") then %>
				<td class="content confezione" colspan="6">confezione</td>
			<% elseif rs("art_varianti") then %>
				<td class="content varianti" colspan="6">articolo con varianti</td>
			<% else %>
				<td class="content" colspan="6">
					articolo singolo
					<% if IsNextAim() then %>
						<a class="button_L2" href="javascript:void(0)" title="converti in articolo con varianti" <%= ACTIVE_STATUS %> style="margin-left:20px;"
						   onclick="OpenAutoPositionedScrollWindow('ArticoliModConverti.asp?ID=<%= rs("art_id") %>', 'CambiaInVarianti<%=rs("art_id")%>', 530, 400, true)">CONVERTI in Articolo con Varianti</a>
					<% end if %>
				</td>
			<% end if %>
		</tr>
		<% if not rs("art_varianti") then 
			sql = " SELECT * FROM gtb_lista_codici LEFT JOIN gtb_codici ON " + _
				  " ( gtb_lista_codici.lstCod_id = gtb_codici.Cod_lista_id AND gtb_codici.cod_variante_id=" & rs("rel_id") & " )" + _
				  " WHERE lstCod_sistema=1 ORDER BY lstCod_nome" 
			rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
			<tr>
				<td class="label" style="width:16%;"<% if rsv.recordcount>0 then %>rowspan="<%= 1+rsv.recordcount %>"<% end if %>>codici:</td>
				<td class="label" style="width:8%;">interno:</td>
				<td class="content" style="width:22%;">
					<% if cIntero(rs("art_external_id")) > 0 then %>
						<input type="text" readonly class="text_disabled" value="<%= rs("art_cod_int") %>" size="15">
						<input type="hidden" name="tft_art_cod_int" value="<%= rs("art_cod_int") %>">
						<% if IsNextAim() then %>
							<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica del codice articolo." <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedScrollWindow('ArticoliModCodice.asp?ID=<%= rs("art_id") %>', 'CambiaCodice<%=rs("art_id")%>', 530, 400, true)">CAMBIA</a>
						<% end if
					else %>
						<input type="text" class="text" name="tft_art_cod_int" value="<%= rs("art_cod_int") %>" maxlength="50" size="15">
						(*)
					<% end if %>
				</td>
				<td class="label" style="width:8%;">alternativo:</td>
				<td class="content">
					<input type="text" class="text" name="tft_art_cod_alt" value="<%= rs("art_cod_alt") %>" maxlength="50" size="15">
				</td>
				<td class="label" style="width:8%;">produttore:</td>
				<td class="content">
					<input type="text" class="text" name="tft_art_cod_pro" value="<%= rs("art_cod_pro") %>" maxlength="50" size="15">
				</td>
			</tr>
			<% while not rsv.eof %>
				<tr>
					<td class="label_no_width" colspan="2">
						<%= rsv("lstCod_nome") %>
					</td>
					<td class="content" colspan="4">
						<input type="text" class="text" name="codice_articolo_<%= rsv("lstCod_id") %>" value="<%= rsv("cod_codice") %>" maxlength="50" size="23">
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
					<input type="text" class="text" name="tft_art_nome_<%= Application("LINGUE")(i) %>" value="<%= textEncode(rs("art_nome_"& Application("LINGUE")(i))) %>" maxlength="250" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next 
		
		categoria = IIF(cInteger(request("tfn_art_tipologia_id"))>0, request("tfn_art_tipologia_id"), rs("art_tipologia_id"))%>
		<tr>
			<td class="label">categoria:</td>
			<td class="content" colspan="6">
				<%CALL dropDown(conn, categorie.QueryElenco(false, ""), "tip_id", "NAME", "tfn_art_tipologia_id", categoria, true, " onchange=""form1.submit()"" style=""width:97%;""", LINGUA_ITALIANO)%>
				(*)
			</td>
		</tr>
		<% sql = "SELECT COUNT(*) FROM gtb_tipologie_raggruppamenti"
		if cIntero(getValueList(conn, rsv, sql))>0 then %>
			<tr>
				<td class="label">&nbsp;</td>
				<td class="label" colspan="2">raggruppamento di pubblicazione:</td>
				<td class="content" colspan="5">
					<% if cInteger(categoria)>0 then
						sql = " SELECT * FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=" & categoria
						rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						if rsv.eof then %>
							<span class="note">Nessun raggruppamento disponibile per questa categoria di prodotti</span>
							<input type="hidden" name="nfn_art_raggruppamento_id" value="NULL">
						<% else
							CALL DropDownRecordset(rsv, "rag_id", "rag_nome_it", "nfn_art_raggruppamento_id", rs("art_raggruppamento_id"), false, "", LINGUA_ITALIANO)
						end if
						rsv.close
					end if %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" nowrap>marchio / produttore:</td>
			<td class="content" colspan="6">
				<%CALL dropDown(conn, "SELECT mar_id, mar_nome_it FROM gtb_marche ORDER BY mar_nome_it", _
							    "mar_id", "mar_nome_it", "tfn_art_marca_id", rs("art_marca_id"), true, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<tr><th colspan="7">DATI PER LA GESTIONE</th></tr>
		<tr>
			<td class="label" colspan="2">non a catalogo:</td>
			<% 'verifica che sia stato messo almeno un componente.
			if rs("art_se_bundle") OR rs("art_se_confezione") then 
				sql = "SELECT COUNT(*) FROM gtb_bundle WHERE bun_bundle_id=" & rs("rel_id")
				aux = cInteger(GetValueList(conn, rsv, sql))>0
			elseif rs("art_varianti") then
				sql = "SELECT COUNT(*) FROM grel_art_valori WHERE rel_art_id=" & rs("art_id")
				aux = cInteger(GetValueList(conn, rsv, sql))>0
			else
				aux = true
			end if
			if aux then
				'articolo singolo (con o senza varianti) o bundle/confezione con almeno un componente
			%>
				<td class="content" colspan="2"><input type="checkbox" class="checkbox" name="chk_art_disabilitato" <%= chk(rs("art_disabilitato")) %>></td>
			<% else 
				'bundle o confezione senza alcun componente definito.
				'articolo con varianti ma senza varianti definite
				%>
				<input type="hidden" name="chk_art_disabilitato" value="1">
				<td class="content"><input type="checkbox" class="checkbox" disabled checked></td>
				<td class="note" colspan="1">
					<% if rs("art_varianti") then  %>
						Sar&agrave; possibile mettere a catalogo l'articolo dopo aver definito almeno una variante.
					<% else %>
						Sar&agrave; possibile mettere a catalogo <%= IIF(rs("art_se_bundle"), "il", "la") %>&nbsp;<%= tipo %> dopo aver selezionato i relativi componenti.
					<% end if %>
				</td>
			<% end if %>
			<td class="label" colspan="2">ordine di pubblicazione:</td>
			<td class="content" colspan="1"><input type="text" class="text" name="tfn_art_ordine" value="<%= rs("art_ordine") %>" size="7"></td>
		</tr>
		<tr>
			<td class="label" colspan="2">non vendibile singolarmente:</td>
			<td class="content" colspan="5"><input type="checkbox" class="checkbox" name="chk_art_NoVenSingola" <%= chk(rs("art_NoVenSingola")) %>></td>
		</tr>
		<tr>
			<td class="label" colspan="2">pezzo unico:</td>
			<td class="content" colspan="5"><input type="checkbox" class="checkbox" name="chk_art_unico" <%= chk(rs("art_unico")) %>></td>
		</tr>
		<% sql = "SELECT spa_id, spa_nome_it  FROM gtb_spese_spedizione_articolo ORDER BY spa_id"
		   rs_spe.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		if rs_spe.RecordCount = 1 then %>
			<input type="hidden" name="tfn_art_spedizione_id" value="<%= rs_spe("spa_id") %>"">
		<% else %>
			<tr>
				<td class="label" colspan="2">metodo di spedizione:</td>
				<td class="content_b" colspan="5">
					<% CALL dropDown(conn, sql, "spa_id", "spa_nome_it", "tfn_art_spedizione_id", rs("art_spedizione_id"), true, "", LINGUA_ITALIANO) %>
				</td>
			</tr>
		<% end if %>
		<% if not rs("art_varianti") then%>
			<tr>
				<td class="label" colspan="2">giacenza minima:</td>
				<td class="content"><input type="text" class="text" name="tfn_art_giacenza_min" value="<%= rs("art_giacenza_min") %>" size="7"></td>
				<td class="note" colspan="4">Limite minimo di quantit&agrave; di prodotto per la segnalazione dello stato "in esaurimento"</td>
			</tr>
			<tr>
				<td class="label" colspan="2">quantit&agrave; minima ordinabile:</td>
				<td class="content" colspan="5"><input type="text" class="text" name="tfn_art_qta_min_ord" value="<%= rs("art_qta_min_ord") %>" size="7"></td>
			</tr>
			<tr>
				<td class="label" colspan="2">quantit&agrave; massima ordinabile:</td>
				<td class="content" colspan="5"><input type="text" class="text" name="tfn_art_qta_max_ord" value="<%= rs("art_qta_max_ord") %>" size="7"></td>
			</tr>
			<tr>
				<td class="label" colspan="2">lotto di riordino:</td>
				<td class="content"><input type="text" class="text" name="tfn_art_lotto_riordino" value="<%= rs("art_lotto_riordino") %>" size="7"></td>
				<td class="note" colspan="4">Indica il numero di articoli che compongono il lotto ordinato dal cliente.</td>
			</tr>
		<% else 
			sql = " SELECT *, (SELECT COUNT(*) FROM gtb_bundle WHERE gtb_bundle.bun_articolo_id=grel_art_valori.rel_id) AS COMPONENTE, " +_
				  " (SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_art_var_id= grel_art_valori.rel_id) AS ORDINI " + _
			 	  " FROM grel_art_valori LEFT JOIN gtb_scontiQ_classi ON grel_art_valori.rel_ScontoQ_id = gtb_scontiQ_classi.scc_id " + _
				  " WHERE rel_art_id=" & request("ID") & " ORDER BY rel_ordine "
			rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
			
			<tr><th colspan="7">IMPOSTAZIONI DI BASE DELL'ARTICOLO</th></tr>
			<% if not rsp.eof then%>
				<tr>
					<td class="label" rowspan="3">codici</td>
					<td class="label">interno:</td>
					<td class="content">
						<input type="hidden" name="tft_art_cod_int" value="<%= rs("art_cod_int") %>">
						<%= rs("art_cod_int") %>
					</td>
					<td class="note" rowspan="3" colspan="4">
						La modifica dei codici non &egrave; permessa perch&egrave; sono gi&agrave; state generate delle varianti. 
						Eseguire le modifiche direttamente nelle singole varianti.
					</td>
				</tr>
				<tr>
					<td class="label">alternativo:</td>
					<td class="content">&nbsp;<%= rs("art_cod_alt") %></td>
				</tr>
				<tr>
					<td class="label">produttore:</td>
					<td class="content">&nbsp;<%= rs("art_cod_pro") %></td>
				</tr>
			<% else %>
				<tr>
					<td class="label" style="width:16%;">codici:</td>
					<td class="label" style="width:8%;">interno:</td>
					<td class="content" style="width:17%;">
						<input type="text" class="text" name="tft_art_cod_int" value="<%= rs("art_cod_int") %>" maxlength="50" size="10">
					</td>
					<td class="label" style="width:8%;">alternativo:</td>
					<td class="content">
						<input type="text" class="text" name="tft_art_cod_alt" value="<%= rs("art_cod_alt") %>" maxlength="50" size="10">
					</td>
					<td class="label" style="width:8%;">produttore:</td>
					<td class="content">
						<input type="text" class="text" name="tft_art_cod_pro" value="<%= rs("art_cod_pro") %>" maxlength="50" size="10">
					</td>
				</tr>
				<tr>
					<td class="label" colspan="2">giacenza minima:</td>
					<td class="content"><input type="text" class="text" name="tfn_art_giacenza_min" value="<%= rs("art_giacenza_min") %>" size="7"></td>
					<td class="note" colspan="4">Limite minimo di quantit&agrave; di prodotto per la segnalazione dello stato "in esaurimento"</td>
				</tr>
				<tr>
					<td class="label" colspan="2">quantit&agrave; minima ordinabile:</td>
					<td class="content" colspan="5"><input type="text" class="text" name="tfn_art_qta_min_ord" value="<%= rs("art_qta_min_ord") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">lotto di riordino:</td>
					<td class="content"><input type="text" class="text" name="tfn_art_lotto_riordino" value="<%= rs("art_lotto_riordino") %>" size="7"></td>
					<td class="note" colspan="4">Indica il numero di articoli che compongono il lotto ordinato dal cliente.</td>
				</tr>
			<% end if
		end if %>
		<% if not cBoolean(rs("art_varianti"),false) then %>
			<% 
			sql = " SELECT * FROM grel_art_valori " + _
				  " WHERE rel_art_id=" & request("ID")
			rsv.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
			
			if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then 
				%>
				<tr><th colspan="7">COLLI, PESO E VOLUME</th></tr>
				<tr>
					<td class="label" rowspan="2" colspan="2">colli:</td>
					<td class="label">numero colli</td>
					<td class="content" colspan="4"><input type="text" class="number" name="extN_rel_colli_num" value="<%= rs("rel_colli_num") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label">numero pezzi per collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="extN_rel_collo_pezzi_per" value="<%= rs("rel_collo_pezzi_per") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" rowspan="2" colspan="2">peso:</td>
					<td class="label">peso netto</td>
					<td class="content" colspan="4"><input type="text" class="number" name="extN_rel_peso_netto" value="<%= rs("rel_peso_netto") %>" size="7"> Kg</td>
				</tr>
				<tr>
					<td class="label">peso lordo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="extN_rel_peso_lordo" value="<%= rs("rel_peso_lordo") %>" size="7"> Kg</td>
				</tr>
				<tr>
					<td class="label" rowspan="4" colspan="2">volume:</td>
					<td class="label">larghezza collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="extN_rel_collo_width" value="<%= rs("rel_collo_width") %>" size="7"> m</td>
				</tr>
				<tr>
					<td class="label">altezza collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="extN_rel_collo_height" value="<%= rs("rel_collo_height") %>" size="7"> m</td>
				</tr>
				<tr>
					<td class="label">lunghezza collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="extN_rel_collo_lenght" value="<%= rs("rel_collo_lenght") %>" size="7"> m</td>
				</tr>
				<tr>
					<td class="label">volume collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="extN_rel_collo_volume" value="<%= rs("rel_collo_volume") %>" size="7"> m</td>
				</tr>
			<% else %>
				<input type="hidden" name="extN_rel_peso_netto" value="<%= rsv("rel_peso_netto") %>">
				<input type="hidden" name="extN_rel_peso_lordo" value="<%= rsv("rel_peso_lordo") %>">
				<input type="hidden" name="extN_rel_colli_num" value="<%= rsv("rel_colli_num") %>">
				<input type="hidden" name="extN_rel_collo_pezzi_per" value="<%= rsv("rel_collo_pezzi_per") %>">
				<input type="hidden" name="extN_rel_collo_width" value="<%= rsv("rel_collo_width") %>">
				<input type="hidden" name="extN_rel_collo_height" value="<%= rsv("rel_collo_height") %>">
				<input type="hidden" name="extN_rel_collo_lenght" value="<%= rsv("rel_collo_lenght") %>">
				<input type="hidden" name="extN_rel_collo_volume" value="<%= rsv("rel_collo_volume") %>">
			<% end if %>
			<% rsv.close %>
		<% end if %>
		<tr><th colspan="7">DEFINIZIONE PREZZO</th></tr>
		<tr>
			<input type="hidden" name="tfn_art_prezzo_base" value="<%= rs("art_prezzo_base")%>">
			<td class="label" colspan="2">prezzo base:</td>
			<td class="content_b" style="width:17%;"><%= FormatPrice(rs("art_prezzo_base"), 2, true) %> &euro;</td>
			<td class="note" colspan="4" rowspan="3">
				La modifica del prezzo, della categoria i.v.a. e della classe di sconto per quantit&agrave; pu&ograve; essere fatta dalla sezione 
				<a class="button_L2" target="_blank" href="ArticoliPrezzi.asp?ID=<%= rs("art_id") %>" title="Apre la gestione dei prezzi dell'articolo in una nuova finestra" <%= ACTIVE_STATUS %>>PREZZI</a>
				dell'articolo.
			</td>
		</tr>
		<tr>
			<input type="hidden" name="tfn_art_iva_id" value="<%= rs("art_iva_id")%>">
			<td class="label" colspan="2">categoria i.v.a.:</td>
			<td class="content_b"><%= rs("iva_nome") %></td>
		</tr>
		<tr>
			<td class="label" colspan="2" nowrap>classe di sconto per quantit&agrave;:</td>
			<td class="content">
				<% if cInteger(rs("art_scontoQ_id"))>0 then 
					sql = "SELECT scc_nome FROM gtb_scontiQ_classi WHERE scc_id=" & rs("art_scontoQ_id")%>
					<%= GetValueList(conn, NULL, sql) %>
				<% end if %>
			</td>
		</tr>
		<tr>
			<th class="L2" colspan="7">specifiche aggiuntive sul prezzo:</th>
		</tr>
		<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="7">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="2" name="tft_art_descr_prezzo_<%= Application("LINGUE")(i) %>"><%= rs("art_descr_prezzo_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
		
		<% if rs("art_varianti") then %>
			<tr><th colspan="7">VARIANTI DELL'ARTICOLO</th></tr>
			<tr>
				<td colspan="7">
					<% if rsp.recordcount > 5 then %>	
						<span class="overflow" style="height:300px;">
					<% end if %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label" style="width:74%">
								<% if rsp.eof then %>
									Nessuna variante definita per il prodotto
								<% else %>
									Trovati n&ordm; <%= rsp.recordcount %> record
								<% end if %>
							</td>
							<% 'verifica se sono possibili altre combinazioni prima di permettere l'inserimento di nuove varianti
							aux = rsp.eof
							if not aux then
								sql = " SELECT var_id FROM gv_articoli_varianti " + _
									  " WHERE rvv_art_var_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & request("ID") & ") " + _
									  " GROUP BY var_id "
								sql = " SELECT COUNT(*) FROM gtb_valori WHERE val_var_id IN (" & GetValueList(conn, rsv, sql) & ") AND " + _
									  " val_id NOT IN (SELECT val_id FROM gv_articoli_Varianti WHERE rvv_art_var_id IN " + _
									  " (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & request("ID") & ")) "
								aux = (cInteger(GetValueList(Conn, rsv, sql))>0)
							end if %>
							<td class="content_right" style="padding-right:0px;">
								<% if aux then %>
									<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra l'inserimento di una variante per l'articolo" <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedScrollWindow('ArticoliVariantiNew.asp?ART_ID=<%= request("ID") %>', 'Varianti', 530, 440, true)">
										NUOVA VARIANTE
									</a>
								<% else %>
									<a class="button_L2_disabled" title="per le varianti scelte non ci sono pi&ugrave; combinazioni possibili." <%= ACTIVE_STATUS %>>
										NUOVA VARIANTE
									</a>
								<% end if %>
							</td>
						</tr>
						<% if not rsp.eof then %>
							<tr>
								<th class="L2" colspan="2">Elenco varianti del prodotto</th>
							</tr>
							<% while not rsp.eof %>
								<tr>
									<td class="body_L2" colspan="2">
											<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
												<tr>
													<td class="<%= IIF(rsp("rel_disabilitato"), "header_L2_disabled", "Header_L2") %>" colspan="9">
														<table border="0" cellspacing="0" cellpadding="0" align="right">
															<tr>
																<td style="font-size: 1px;">
																	<a name="<%= Server.HTMLEncode(rsp("rel_cod_int")) %>"
																	   class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica della variante" <%= ACTIVE_STATUS %>
											   						   onclick="OpenAutoPositionedScrollWindow('ArticoliVariantiMod.asp?ID=<%= rsp("rel_id") %>', 'Varianti', 530, 390, true)">
																		MODIFICA
																	</a>
																	&nbsp;
																	<% if cInteger(rsp("ORDINI"))>0 then %>
																		<a class="button_L2_disabled" title="cancellazione non permessa: &egrave; presente almeno un ordine per questa variante." <%= ACTIVE_STATUS %>>
																			CANCELLA
																		</a>
																	<% elseif cInteger(rsp("COMPONENTE"))>0 then %>
																		<a class="button_L2_disabled" title="cancellazione non permessa: la variante &egrave; un componente di almeno un bundle o di una confezione." <%= ACTIVE_STATUS %>>
																			CANCELLA
																		</a>
																	<% elseif cInteger(rsp("rel_external_id"))>0 then %>
																		<a class="button_L2_disabled" title="cancellazione non permessa: la variante &egrave; collegata ad un articolo esterno." <%= ACTIVE_STATUS %>>
																			CANCELLA
																		</a>
																	<% else %>
																		<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione della variante" <%= ACTIVE_STATUS %>
												   						   onclick="OpenDeleteWindow('ARTICOLI_VARIANTI','<%= rsp("rel_id") %>');">
																			CANCELLA
																		</a>
																	<% end if %>
																</td>
															</tr>
														</table>
														<% CALL TableValoriVarianti(conn, rsv, rsp("rel_id"), IIF(rsp("rel_disabilitato"), "header_L2_disabled", "Header_L2")) %>
													</td>
												</tr>
												<tr>
													<td style="width:7%;"  class="label" rowspan="3">codici:</td>
													<td style="width:9%;"  class="label">interno:</td>
													<td style="width:17%;" class="content"><%= rsp("rel_cod_int") %></td>
													<td style="width:7%;"  class="label" rowspan="3">magazzino:</td>
													<td style="width:15%;" class="label">giacenza minima:</td>
													<td style="width:5%;"  class="content_right"><%= rsp("rel_giacenza_min") %></td>
													<td style="width:7%;"  class="label" rowspan="2">prezzo:</td>
													<td style="width:16%;" class="label">prezzo variante:</td>
													<% if rsp("rel_prezzo_indipendente") then %>
														<td class="content" title="prezzo variante indipendente da prezzo base articolo">
													<% else %>
														<td class="content" title="prezzo variante dipendente da prezzo articolo e ricalcolato sulla base degli sconti applicati">
													<% end if %>
														<strong>
															<%= FormatPrice(rsp("rel_prezzo"), 2, true) %> &euro;
														</strong>
														<% if cReal(rsp("rel_var_sconto"))<>0 then %>
															( <%= FormatPrice(rsp("rel_var_sconto"), 2, true) %>% )
														<% elseif cReal(rsp("rel_var_euro"))<>0 then %>
															( <%= FormatPrice(rsp("rel_var_euro"), 2, true) %> &euro;)
														<% end if %>
													</td>
												</tr>
												<tr>
													<td style="width:9%;"  class="label">alternativo:</td>
													<td style="width:17%;" class="content"><%= rsp("rel_cod_alt") %></td>
													<td style="width:15%;" class="label">qta min. ordinabile:</td>
													<td style="width:5%;"  class="content_right"><%= rsp("rel_qta_min_ord") %></td>
													<td style="width:16%;" class="label">classe di sconto qta:</td>
													<td class="content"><%= rsp("scc_nome") %></td>
												</tr>
												<tr>
													<td style="width:9%;"  class="label">produttore:</td>
													<td style="width:17%;" class="content"><%= rsp("rel_cod_pro") %></td>
													<td style="width:15%;" class="label">lotto di riordino:</td>
													<td style="width:5%;"  class="content_right"><%= rsp("rel_lotto_riordino") %></td>
													<td style="width:9%;"  class="label">stato:</td>
													<td style="width:36%;" class="content" colspan="2">
														<table width="100%" cellspacing="0" cellpadding="0">
															<tr>
																<td class="content"><%= IIF(rsp("rel_disabilitato"), "non visibile a catalogo", "a catalogo") %></td>
																<% if cInteger(rsp("componente"))>0 then 
																	sql = " SELECT COUNT(*) FROM gv_articoli WHERE rel_id IN " + _
																		  " (SELECT bun_bundle_id FROM gtb_bundle WHERE bun_articolo_id=" & rsp("rel_id") & ")" + _
																		  " AND " + SQL_IsTrue(conn, "art_se_bundle")
																	if GetValueList(conn, rsv, sql)>0 then%>
																		<td class="content bundle">in bundle</td>
																	<% end if
																	sql = " SELECT COUNT(*) FROM gv_articoli WHERE rel_id IN " + _
																		  " (SELECT bun_bundle_id FROM gtb_bundle WHERE bun_articolo_id=" & rsp("rel_id") & ")" + _
																		  " AND " + SQL_IsTrue(conn, "art_se_confezione")
																	if GetValueList(conn, rsv, sql)>0 then %>
																		<td class="content confezione">in confezione</td>
																	<%end if
																end if %>
															</tr>
														</table>
													</td>
												</tr>
											</table>
									</td>
								</tr>
								<%rsp.movenext
							wend 
						end if%>
					</table>
					<% if rsp.recordcount > 5 then %>
						</span>
					<%end if
					rsp.close %>
				</td>
			</tr>
		<% end if 
		if rs("art_se_bundle") OR rs("art_se_confezione") then%>
			<tr><th colspan="7">COMPOSIZIONE</th></tr>
			<tr>
				<td colspan="7">
					<% sql = " SELECT * FROM gv_articoli INNER JOIN gtb_bundle ON gv_articoli.rel_id = gtb_bundle.bun_articolo_id " + _
							 " WHERE gtb_bundle.bun_bundle_id=" & rs("rel_id")
					rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label" colspan="2" style="width:74%">
								<% if rsp.eof then %>
									Nessun componente definito.
								<% else %>
									Trovati n&ordm; <%= rsp.recordcount %> record
								<% end if %>
							</td>
							<td colspan="4" class="content_right" style="padding-right:0px;">
								<% if cInteger(rs("N_ORDINI"))>0 then %>
									<a class="button_L2_disabled" title="la modifica della composizione del bundle non &egrave; permessa perch&egrave; sono gi&agrave; stati effettuati ordini di questo prodotto." <%= ACTIVE_STATUS %>>
										NUOVO COMPONENTE
									</a>
								<% else %>
									<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la selezione ed inserimento del nuovo componente" <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedScrollWindow('ArticoliSeleziona.asp?TYPE=C&EXCLUDE_ID=<%= rs("rel_id") %>&SelectedPage=ArticoliComponentiNew', 'COMArt', 530, 400, true)">
										NUOVO COMPONENTE
									</a>
								<% end if %>
							</td>
						</tr>
						<% if not rsp.eof then %>
							<tr>
								<th class="l2_center" width="14%">codice</th>
								<th class="L2">descrizione componente</th>
								<th class="l2_center" width="8%">quantit&agrave;</th>
								<th class="l2_center" width="17%" colspan="2">operazioni</th>
							</tr>
							<% while not rsp.eof %>
								<tr>
									<td class="content"><%= rsp("rel_cod_int")%></td>
									<td class="content">
										<% ArticoloLink rsp("art_id"), rsp("art_nome_it"), rsp("rel_cod_int") %>
										<% if rsp("art_varianti") then %>
											<%= ListValoriVarianti(conn, rsv, rsp("rel_id")) %>
										<% end if %>
									</td>
									<td class="content_center"><%= rsp("bun_quantita")%></td>
									<% if cInteger(rs("N_ORDINI"))>0 then %>
										<td class="content_center">
											<a class="button_L2_disabled" title="la modifica della composizione del bundle non &egrave; permessa perch&egrave; sono gi&agrave; stati effettuati ordini di questo prodotto." <%= ACTIVE_STATUS %>>
												MODIFICA
											</a>
										</td>
										<td class="content_center">
											<a class="button_L2_disabled" title="la modifica della composizione del bundle non &egrave; permessa perch&egrave; sono gi&agrave; stati effettuati ordini di questo prodotto." <%= ACTIVE_STATUS %>>
												CANCELLA
											</a>
										</td>
									<% else %>
										<td class="content_center">
											<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica del componente" <%= ACTIVE_STATUS %>
											   onclick="OpenAutoPositionedWindow('ArticoliComponentiMod.asp?ID=<%= rsp("bun_id") %>', 'COMArt', 510, 270)">
												MODIFICA
											</a>
										</td>
										<td class="content_center">
											<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione del componente" <%= ACTIVE_STATUS %>
											   onclick="OpenDeleteWindow('ARTICOLI_COMPONENTI','<%= rsp("bun_id") %>');">
												CANCELLA
											</a>
										</td>
									<% end if %>
								</tr>
								<%rsp.movenext
							wend 
						end if%>
						<tr>
							<th class="L2" colspan="5">note aggiuntive</th>
						</tr>
						<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
							<tr>
								<td class="content" colspan="7">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
										<tr>
											<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
											<td><textarea style="width:100%;" rows="2" name="tft_art_composizione_note_<%= Application("LINGUE")(i) %>"><%= rs("art_composizione_note_" & Application("LINGUE")(i)) %></textarea></td>
										</tr>
									</table>
								</td>
							</tr>
						<%next %>
					</table>
					<% rsp.close %>
				</td>
			</tr>
		<% end if %>
		<tr><th colspan="7">DESCRIZIONE RIASSUNTIVA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="7">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="3" name="tft_art_descr_riassunto_<%= Application("LINGUE")(i) %>"><%= rs("art_descr_riassunto_" & Application("LINGUE")(i)) %></textarea></td>
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
							<td><textarea style="width:100%;" rows="5" name="tft_art_descr_<%= Application("LINGUE")(i) %>"><%= rs("art_descr_" & Application("LINGUE")(i)) %></textarea></td>
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
	
	<% 'sql = " SELECT TOP 1 ct_id FROM gtb_carattech INNER JOIN gtb_tip_ctech ON gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id " & _
		'	  " WHERE rct_tipologia_id = " & IIF(cInteger(request("tfn_art_tipologia_id"))>0, request("tfn_art_tipologia_id"), rs("art_tipologia_id"))
  
		'if cIntero(GetValueList(conn, NULL, sql)) > 0 then 
	%>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr><th colspan="7">CARATTERISTICHE TECNICHE</th></tr>
			<% 	sql = " SELECT * " & _
					  " FROM gtb_carattech"& _
					  " left JOIN gtb_tip_ctech ON (gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id AND rct_tipologia_id=" & IIF(cInteger(request("tfn_art_tipologia_id"))>0, request("tfn_art_tipologia_id"), rs("art_tipologia_id")) &")" + _
					  " LEFT JOIN grel_art_ctech ON (gtb_carattech.ct_id = grel_art_ctech.rel_ctech_id AND grel_art_ctech.rel_art_id=" & rs("art_id") &")"& _
					  " LEFT JOIN gtb_carattech_raggruppamenti ON gtb_carattech.ct_raggruppamento_id = gtb_carattech_raggruppamenti.ctr_id " & _
					  " ORDER BY ctr_ordine, ctr_id, rct_ordine"		
			'CALL DesForm  (conn, sql, "gtb_carattech", "ct_id", "ct_nome_it", "ct_tipo", "ct_unita_it", "", "rel_ctech_", "rel_ctech_", "ctr_titolo_it", cIntero(request("ID")) = 0, 7)
			CALL DesFullFormComplete(NULL, conn, sql, "gtb_carattech", "ct_id", "ct_nome_it", "ct_tipo", "ct_unita_it",  "", "", "", "rel_ctech_", "rel_ctech_", "ctr_titolo_it", cIntero(request("ID")) = 0, 7, false, "")	
			%>
		</table>

	
	<% 	CALL oArticoliFoto.Elenco(request("ID"), "FOTO") %>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="7">ARTICOLI COLLEGATI</th></tr>
		<tr>
			<td colspan="7">
				<% sql = " SELECT * FROM gtb_articoli INNER JOIN grel_art_acc ON gtb_articoli.art_id = grel_art_acc.aa_acc_id " + _
						 " INNER JOIN gtb_accessori_tipo ON grel_art_acc.aa_tipo_id = gtb_accessori_tipo.at_id " + _
						 " WHERE grel_art_acc.aa_art_id=" & request("ID") & " ORDER BY gtb_accessori_tipo.at_ordine, grel_Art_acc.aa_ordine"
				rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label" colspan="2" style="width:74%">
							<% if rsp.eof then %>
								Nessun articolo collegato definito.
							<% else %>
								Trovati n&ordm; <%= rsp.recordcount %> record
							<% end if %>
						</td>
						<td colspan="5" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la selezione di un nuovo articolo collegato" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedScrollWindow('ArticoliSeleziona.asp?TYPE=A&EXCLUDE_ID=<%= request("ID") %>&SelectedPage=ArticoliAccessoriNew', 'COMAcc', 530, 525, true)">
								NUOVO ARTICOLO COLLEGATO
							</a>
						</td>
					</tr>
					<% if not rsp.eof then %>
						<tr>
							<th class="L2" width="15%">tipo</th>
							<th class="l2_center" width="10%">codice</th>
							<th class="L2">nome</th>
							<th class="l2_center" width="6%">ordine</th>
							<th width="14%" class="l2_center">non vendibile sing.</th>
							<th colspan="2" class="l2_center" width="16%">operazioni</th>
						</tr>
						<% while not rsp.eof %>
							<tr>
								<td class="content"><%= rsp("at_nome_it")%></td>
								<td class="content"><%= rsp("art_cod_int")%></td>
								<td class="content"><% ArticoloLink rsp("art_id"), rsp("art_nome_it"), rsp("art_cod_int") %></td>
								<td class="content_center"><%= rsp("aa_ordine")%></td>
								<td class="content_center">
									<% if rsp("at_vincolo_vendita") then %>
										<input type="checkbox" class="checkbox" disabled <%= chk(rsp("art_noVenSingola")) %>>
									<% else %>
										&nbsp;
									<% end if %>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica dell'articolo collegato" <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedScrollWindow('ArticoliAccessoriMod.asp?ID=<%= rsp("aa_id") %>', 'COMAcc', 510, 250, true)">
										MODIFICA
									</a>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione dell'articolo collegato" <%= ACTIVE_STATUS %>
									   onclick="OpenDeleteWindow('ARTICOLI_ACCESSORI','<%= rsp("aa_id") %>');">
										CANCELLA
									</a>
								</td>
							</tr>
							<%rsp.movenext
						wend 
					end if%>
					<tr>
						<th class="L2" colspan="6">note aggiuntive</th>
					</tr>
					<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
						<tr>
							<td class="content" colspan="7">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
									<tr>
										<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
										<td><textarea style="width:100%;" rows="2" name="tft_art_accessori_note_<%= Application("LINGUE")(i) %>"><%= rs("art_accessori_note_" & Application("LINGUE")(i)) %></textarea></td>
									</tr>
								</table>
							</td>
						</tr>
					<%next %>
				</table>
				<% rsp.close %>
			</td>
		</tr>
		<tr><th colspan="7">NOTE INTERNE</th></tr>
		<tr>
			<td class="content" colspan="7">
				<textarea style="width:100%;" rows="3" name="tft_art_note"><%= rs("art_note") %></textarea>
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% CALL Form_DatiModifica(conn, rs, "art_") %>
		<tr>
			<td class="footer" colspan="7">
				(*) Campi obbligatori.
				<input type="submit" style="width:23%;" class="button" name="salva_elenco" value="SALVA & TORNA ALL'ELENCO">
				<input type="submit" class="button" name="salva" value="SALVA">
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
set rsp = nothing
conn.Close
set conn = nothing
%>