<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" 

dim conn, rs, rsp, rsv, rsl, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")
set rsl = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_ARTICOLI_SQL"), "art_id", "ArticoliPrezzi.asp")
end if

response.buffer = false
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="ListiniPrezzi_tools.asp" -->
<% 	

sql = " SELECT * FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + _
	  " INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = gtb_iva.iva_id " + _
	  " LEFT JOIN gtb_scontiq_classi ON gtb_articoli.art_scontoQ_id = gtb_scontiq_classi.scc_id " + _
	  " WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

dim dicitura, tipo, listino
if rs("art_se_bundle") then
	tipo = "bundle"
elseif rs("art_se_confezione") then
	tipo = "confezione"
elseif rs("art_varianti") then
	tipo ="articolo con varianti"
else
	tipo ="articolo singolo"
end if
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione articoli - prezzi " & tipo
dicitura.puls_new = "INDIETRO;SCHEDA ARTICOLO;"
dicitura.link_new = "Articoli.asp;ArticoliMod.asp?ID=" & request("ID")
dicitura.puls_2a_riga.Add "GIACENZE","ArticoliGiacenze.asp?ID=" & request("ID")
if Session("ATTIVA_FAQ_ARTICOLI") then
	dicitura.puls_2a_riga.Add "FAQ","ArticoliFaq.asp?ID=" & request("ID")
end if
if Session("ATTIVA_COMMENTI") then
	dicitura.puls_2a_riga.Add "COMMENTI","ArticoliCommenti.asp?ID=" & request("ID")
end if

dicitura.scrivi_con_sottosez()
%>


<div id="content_abbassato">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati <%= tipo %> con codice &quot;<%= rs("art_cod_int") %>&quot;</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="articolo precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="articolo successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="7">DATI DELL'ARTICOLO</th></tr>
		<% CALL ArticoloScheda (conn, rs, rsp) %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<caption>Prezzo base <%= IIF(rs("art_varianti"), " e delle varianti dell'articolo", "") %></caption>
		<tr><th colspan="3">PREZZO BASE DELL'ARTICOLO</th></tr>
		<tr>
			<td class="label">prezzo base:</td>
			<td class="content" width="20%" title="<%= FormatPrice(cReal(rs("art_prezzo_base")) + GetIva(cReal(rs("art_prezzo_base")), rs("iva_valore")), 2, true) %>&euro; iva inclusa."><%= FormatPrice(cReal(rs("art_prezzo_base")), 2, true) %> &euro;</td>
			<td class="content">
				<a class="button_L2" href="javascript:void(0);" title="Apre la modifica del prezzo in una nuova finestra" <%= ACTIVE_STATUS %>
				   onclick="OpenAutoPositionedScrollWindow('ArticoliPrezzi_Mod_Art_Prz.asp?ID=<%= rs("art_id") %>', 'Prezzi', 510, 380, false)">
					MODIFICA PREZZO
				</a>
			</td>
		</tr>
		<tr>
			<td class="label">categoria i.v.a.:</td>
			<td class="content" title="<%= rs("iva_valore") %>%"><%= rs("iva_nome") %></td>
			<td class="content">
				<a class="button_L2" href="javascript:void(0);" title="Apre la modifica della categoria i.v.a. in una nuova finestra" <%= ACTIVE_STATUS %>
				   onclick="OpenAutoPositionedScrollWindow('ArticoliPrezzi_Mod_Art_IVA.asp?ID=<%= rs("art_id") %>', 'IVA', 510, 380, false)">
					MODIFICA CATEGORIA I.V.A.
				</a>
			</td>
		</tr>
		<tr>
			<td class="label"nowrap>classe di sconto per quantit&agrave;:</td>
			<td class="content">
				<%= rs("scc_nome") %>
			</td>
			<td class="content">
				<a class="button_L2" href="javascript:void(0)" title="Apre la modifica della classe di sconto per quantit&agrave; in una nuova finestra" <%= ACTIVE_STATUS %>
				   onclick="OpenAutoPositionedScrollWindow('ArticoliPrezzi_Mod_Art_ScQta.asp?ID=<%= rs("art_id") %>', 'Sconti', 510, 380, false)">
					MODIFICA CLASSE DI SCONTO PER QUANTIT&Agrave;
				</a>
			</td>
		</tr>
		<% if rs("art_varianti") then %>
			<tr><th colspan="3">PREZZI BASE DELLE VARIANTI</th></tr>
			<tr>
				<td colspan="3">
					<%sql = " SELECT * FROM grel_art_valori LEFT JOIN gtb_scontiQ_classi ON grel_art_valori.rel_ScontoQ_id = gtb_scontiQ_classi.scc_id " + _
				  			" WHERE rel_art_id=" & cIntero(request("ID")) & " ORDER BY rel_ordine "
					rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText%>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label" style="width:100%" colspan="10">
								<% if rsp.eof then %>
									Nessuna variante definita per il prodotto
								<% else %>
									Trovati n&ordm; <%= rsp.recordcount %> record
								<% end if %>
							</td>
						</tr>
						<tr>
							<th class="L2">codice</th>
							<th class="L2">variante</th>
							<th style="width:9%;" class="l2_center">a catalogo</th>
							<th class="L2_right">prezzo</th>
							<th style="width:10%;" class="l2_center" title="prezzo indipendente dal prezzo base dell'articolo">indipendente</th>
							<th class="L2_right" title="variazione da prezzo base articolo">var. &euro;</th>
							<th class="L2_right" title="variazione da prezzo base articolo">var. %</th>
							<th class="l2_center">classe sc. qta</th>
							<th class="l2_center" colspan="2">modifica</th>
						</tr>
						<% while not rsp.eof %>
							<tr>
								<td class="content"><%= rsp("rel_cod_int") %></td>
								<td class="content">
									<% CALL TableValoriVarianti(conn, rsv, rsp("rel_id"), IIF(rsp("rel_disabilitato"), "content_disabled", "content")) %>
								</td>
								<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(not rsp("rel_disabilitato")) %>></td>
								<td class="content_right"><%= FormatPrice(rsp("rel_prezzo"), 2, true) %>&nbsp;&euro;</td>
								<td class="content_center" title="<%= IIF(rsp("rel_prezzo_indipendente"), "prezzo indipendente dal prezzo base dell'articolo", "prezzo variante calcolato applicando le variazioni impostate al prezzo base articolo") %>"><input type="checkbox" class="checkbox" disabled <%= chk(rsp("rel_prezzo_indipendente")) %>></td>
								<td class="content_right">
									<% if cReal(rsp("rel_var_euro"))<>0 then %>
										<%= FormatPrice(cReal(rsp("rel_var_euro")), 2, true) %>&nbsp;&euro;
									<% end if %>
								</td>
								<td class="content_right">
									<% if cReal(rsp("rel_var_sconto"))<>0 then %>
										<%= FormatPrice(cReal(rsp("rel_var_sconto")), 2, true) %>&nbsp;%
									<% end if %>
								</td>
								<td class="content_center"><%= rsp("scc_nome") %></td>
								<td class="content_center" style="width:7%;">
									<a class="button_L2" href="javascript:void(0);" title="Apre la modifica del prezzo in una nuova finestra" <%= ACTIVE_STATUS %>
				   					   onclick="OpenAutoPositionedScrollWindow('ArticoliPrezzi_Mod_Var_Prz.asp?ID=<%= rsp("rel_id") %>', 'Prezzi', 510, 370, true)">
										PREZZO
									</a>
								</td>
								<td class="content_center" style="width:11%;">
									<a class="button_L2" href="javascript:void(0)" title="Apre la modifica della classe di sconto per quantit&agrave; in una nuova finestra" <%= ACTIVE_STATUS %>
				   					   onclick="OpenAutoPositionedScrollWindow('ArticoliPrezzi_Mod_Var_ScQta.asp?ID=<%= rsp("rel_id") %>', 'Sconti', 510, 340, false)">
										SCONTO QTA
									</a>
								</td>
							</tr>
							<% rsp.movenext
						wend %>
					</table>
					<% rsp.close %>
				</td>
			</tr>
		<% end if %>
	</table>
	
	<% 'recupera dati del listino base attuale
	dim listino_base
	sql = "SELECT listino_id FROM gtb_listini WHERE listino_base_attuale=1"
	listino_base = GetValueList(conn, rsl, sql) 
	
	'imposta variabili di sessione per visualizzazione esploso listini
	if request("listini_base_mostra")<>"" then
		Session("ARTICOLI_PREZZI_LISTINI_BASE") = true
	elseif request("listini_base_nascondi")<>"" then
		Session("ARTICOLI_PREZZI_LISTINI_BASE") = false
	end if
	
	if request("listini_offerte_mostra")<>"" then
		Session("ARTICOLI_PREZZI_LISTINI_OFFERTE") = true
	elseif request("listini_offerte_nascondi")<>"" then
		Session("ARTICOLI_PREZZI_LISTINI_OFFERTE") = false
	end if
	
	if request("listini_clienti_mostra")<>"" then
		Session("ARTICOLI_PREZZI_LISTINI_CLIENTI") = true
	elseif request("listini_clienti_nascondi")<>"" then
		Session("ARTICOLI_PREZZI_LISTINI_CLIENTI") = false
	end if
	
	if request("listini_personal_clienti_mostra")<>"" then
		Session("ARTICOLI_PREZZI_PERSONALIZZATI_LISTINI_CLIENTI") = true
	elseif request("listini_personal_clienti_nascondi")<>"" then
		Session("ARTICOLI_PREZZI_PERSONALIZZATI_LISTINI_CLIENTI") = false
	end if
	%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<form action="" method="post">
		<caption <% if Session("ARTICOLI_PREZZI_LISTINI_BASE") then %>class="border"<% end if %>>
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Prezzi nei listini base</td>
					<td align="right">
						<% if Session("ARTICOLI_PREZZI_LISTINI_BASE") then %>
							<input type="submit" name="listini_base_nascondi" value="NASCONDI ELENCO COMPLETO PREZZI" class="button" style="width:210px;">
						<% else%>
							<input type="submit" name="listini_base_mostra" value="VISUALIZZA ELENCO COMPLETO PREZZI" class="button" style="width:215px;">
						<% end if %>
					</td>
				</tr>
			</table>
		</caption>
	</form>
		<% if Session("ARTICOLI_PREZZI_LISTINI_BASE") then %>
			<% sql = "SELECT * FROM gtb_listini WHERE listino_base=1 ORDER BY listino_codice"
			CALL EsplosoListini(sql)
		end if %>
	</table>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<form action="" method="post">
		<caption <% if Session("ARTICOLI_PREZZI_LISTINI_OFFERTE") then %>class="border"<% end if %>>
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Prezzi nei listini offerte speciali</td>
					<td align="right">
						<% if Session("ARTICOLI_PREZZI_LISTINI_OFFERTE") then %>
							<input type="submit" name="listini_offerte_nascondi" value="NASCONDI ELENCO COMPLETO PREZZI" class="button" style="width:210px;">
						<% else%>
							<input type="submit" name="listini_offerte_mostra" value="VISUALIZZA ELENCO COMPLETO PREZZI" class="button" style="width:215px;">
						<% end if %>
					</td>
				</tr>
			</table>
		</caption>
	</form>
		<% if Session("ARTICOLI_PREZZI_LISTINI_OFFERTE") then %>
			<% sql = "SELECT * FROM gtb_listini WHERE listino_offerte=1 ORDER BY listino_codice"
			CALL EsplosoListini(sql)
		end if %>
	</table>
		
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<form action="" method="post">
		<caption <% if Session("ARTICOLI_PREZZI_LISTINI_CLIENTI") then %>class="border"<% end if %>>
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Prezzi nei listini clienti principali</td>
					<td align="right">
						<% if Session("ARTICOLI_PREZZI_LISTINI_CLIENTI") then %>
							<input type="submit" name="listini_clienti_nascondi" value="NASCONDI ELENCO COMPLETO PREZZI" class="button" style="width:210px;">
						<% else%>
							<input type="submit" name="listini_clienti_mostra" value="VISUALIZZA ELENCO COMPLETO PREZZI" class="button" style="width:215px;">
						<% end if %>
					</td>
				</tr>
			</table>
		</caption>
	</form>
		<% if Session("ARTICOLI_PREZZI_LISTINI_CLIENTI") then %>
			<% sql = "SELECT * FROM gtb_listini WHERE ISNULL(listino_base, 0)=0 AND ISNULL(listino_offerte, 0)=0 AND ISNULL(listino_ancestor_id, 0)=0 ORDER BY listino_codice"
			CALL EsplosoListini(sql)
		end if %>
	</table>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<form action="" method="post">
		<caption <% if Session("ARTICOLI_PREZZI_PERSONALIZZATI_LISTINI_CLIENTI") then %>class="border"<% end if %>>
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Prezzi personalizzati nei listini clienti</td>
					<td align="right">
						<% if Session("ARTICOLI_PREZZI_PERSONALIZZATI_LISTINI_CLIENTI") then %>
							<input type="submit" name="listini_personal_clienti_nascondi" value="NASCONDI ELENCO COMPLETO PREZZI" class="button" style="width:210px;">
						<% else%>
							<input type="submit" name="listini_personal_clienti_mostra" value="VISUALIZZA ELENCO COMPLETO PREZZI" class="button" style="width:215px;">
						<% end if %>
					</td>
				</tr>
			</table>
		</caption>
	</form>
		<% if Session("ARTICOLI_PREZZI_PERSONALIZZATI_LISTINI_CLIENTI") then %>
			<% sql = " SELECT gtb_listini.* FROM gtb_listini INNER JOIN gtb_listini listino_ancestor ON gtb_listini.listino_ancestor_id = listino_ancestor.listino_id " + _
					 " WHERE ISNULL(gtb_listini.listino_base, 0)=0 AND ISNULL(gtb_listini.listino_offerte, 0)=0 AND " + _
					 " EXISTS (SELECT 1 FROM gtb_prezzi INNER JOIN grel_art_valori ON gtb_prezzi.prz_variante_id = grel_art_valori.rel_id " + _
					 "		   INNER JOIN gtb_prezzi prezzi_ancestor ON gtb_prezzi.prz_variante_id=prezzi_ancestor.prz_variante_id " + _
					 "			  	 AND grel_art_valori.rel_art_id=" & cIntero(request("ID")) & " AND gtb_prezzi.prz_listino_id=gtb_listini.listino_id " + _
					 "				 AND prezzi_ancestor.prz_listino_id = gtb_listini.listino_ancestor_id " + _
					 "		   WHERE ISNULL(prezzi_ancestor.prz_visibile,0) <> ISNULL(gtb_prezzi.prz_visibile,0) OR " + _
					 "				 ISNULL(prezzi_ancestor.prz_promozione,0) <> ISNULL(gtb_prezzi.prz_promozione,0) OR " + _
					 "				 ISNULL(prezzi_ancestor.prz_scontoQ_id,0) <> ISNULL(gtb_prezzi.prz_scontoQ_id,0) OR " + _
					 "				 ISNULL(prezzi_ancestor.prz_iva_id,0) <> ISNULL(gtb_prezzi.prz_iva_id,0) OR " + _
					 "				 ISNULL(prezzi_ancestor.prz_prezzo,0) <> ISNULL(gtb_prezzi.prz_prezzo,0) OR " + _
					 "				 ISNULL(prezzi_ancestor.prz_var_euro,0) <> ISNULL(gtb_prezzi.prz_var_euro,0) OR " + _
					 "				 ISNULL(prezzi_ancestor.prz_var_sconto,0) <> ISNULL(gtb_prezzi.prz_var_sconto,0) ) " + _
					 " ORDER BY gtb_listini.listino_codice"
			CALL EsplosoListini(sql)
		end if %>
	</table>
	&nbsp;
</div>
</body>
</html>
<%
set rs = nothing
set rsv = nothing
set rsp = nothing
conn.Close
set conn = nothing







sub EsplosoListini(SqlListini)
	dim Sql
	rsl.open SqlListini, conn, adOpenStatic ,adLockOptimistic, adCmdText %>
	<tr>
		<td class="label" style="width:100%" colspan="10">
			<% if rsl.eof then %>
				Nessun listino individuato
			<% else %>
				Trovati n&ordm; <%= rsl.recordcount %> record
			<% end if %>
		</td>
	</tr>
	<% while not rsl.eof 
		sql = " SELECT * , " + _
			  " (SELECT COUNT(*) FROM gv_listino_offerte WHERE gv_listino_offerte.prz_variante_id=gv_listini.prz_variante_id) AS OFFERTA " + _
			  " FROM gv_listini LEFT JOIN gtb_scontiQ_classi ON gv_listini.prz_scontoQ_id= gtb_scontiQ_classi.scc_id  " + _
  			  " WHERE art_id=" & cIntero(request("ID")) & " AND "
		if rsl("listino_base") then
			sql = sql & " prz_listino_id=" & rsl("listino_id")
		else
			sql = sql & " (prz_listino_id=" & rsl("listino_id") & " OR " + _
						" (prz_listino_id =" & listino_base & " AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_prezzi WHERE prz_listino_id=" & rsl("listino_id") & "))) "
		end if
		if cInteger(rsL("listino_ancestor_id"))>0 then
			sql = sql + SearchForm_SEARCH_BY_personalizzazione_prezzoEX(rsL, true, 0, cIntero(request("ID")))
		end if
		sql = sql + " ORDER BY art_cod_int, rel_ordine"
		rsp.open sql, conn, adOpenStatic ,adLockOptimistic, adCmdText %>
		<tr>
			<th colspan="10">
				<%= rsl("listino_codice") %>
				<% if rsl("listino_offerte") AND (rsl("listino_datacreazione")<= Date AND (IsNull(rsl("listino_datascadenza")) OR rsl("listino_datascadenza")>= Date)) then %>
					<span class="Icona Offerte" title="listino offerte speciali attualmente in vigore">&nbsp;</span>
				<% end if %>
			</th>
		</tr>
		<form action="ListiniPrezzi_Avanzata.asp?ID=<%= rsl("listino_id") %>" method="post" target="listino" id="form_<%= rsl("listino_id") %>" name="form_<%= rsl("listino_id") %>">
		<tr>
			<td class="content_right" colspan="10">
				<% if not rs("art_varianti") then%>
					<input type="hidden" name="search_codice_int" value="<%= rs("art_cod_int") %>">
				<% end if %>
				<input type="hidden" name="search_nome" value="<%= TextEncode(rs("art_nome_it")) %>">
				<input type="hidden" name="search_categoria" value="<%= rs("art_tipologia_id") %>">
				<input type="hidden" name="search_marchio" value="<%= rs("art_marca_id") %>">
				<input type="submit" name="cerca" value="APRI LISTINO" class="button_L2" title="Apre il listino in modifica" onclick="OpenAutoPositionedScrollWindow('', 'listino', 760, 450, true);">
			</td>
		</tr>
		</form>
		<tr>
			<% if rs("art_varianti") then %>
				<th class="L2">codice</th>
				<th class="L2">variante</th>
			<% else %>
				<th class="L2" colspan="2">codice</th>
			<% end if %>
			<% if rsl("listino_offerte") then %>
				<th class="l2_center" colspan="2">in offerta</th>
			<% else %>
				<th class="l2_center" style="width:4%;">vis.</th>
				<th class="l2_center" style="width:5%;">promo.</th>
			<% end if %>
			<th class="L2_right">prezzo</th>
			<th class="L2_right">var. &euro;</th>
			<th class="L2_right">var. %</th>
			<th class="L2_right">i.v.a.</th>
			<th class="l2_center" style="width:9%;">classe sc.</th>
		</tr>
		<% 
		
		while not rsp.eof%>
			<tr>
				<% if rs("art_varianti") then %>
					<td class="content"><%= rsp("rel_cod_int") %></td>
					<td class="<%= IIF(rsp("rel_disabilitato"), "content_disabled", "content") %>">
						<% CALL TableValoriVarianti(conn, rsv, rsp("rel_id"), IIF(rsp("rel_disabilitato"), "content_disabled", "content")) %>
					</td>
				<% else %>
					<td class="content" colspan="2"><%= rsp("rel_cod_int") %></td>
				<% end if %>
				<% if rsl("listino_offerte") then %>
					<td class="content_center" colspan="2">
						<input type="checkbox" class="checkbox" disabled <%= chk(rsp("prz_listino_id")=rsl("listino_id") AND rsp("prz_visibile")) %>>
						<% if cInteger(rsp("OFFERTA"))>0 then %>
							<span class="Icona Offerte" title="articolo attualmente in offerta speciale">&nbsp;</span>
						<% end if %>
					</td>
				<% else %>
					<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(rsp("prz_visibile")) %>></td>
					<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(rsp("prz_promozione")) %>></td>
				<% end if %>
				<td class="content_right" nowrap><%= FormatPrice(rsp("prz_prezzo"), 2, true) %>&nbsp;&euro;</td>
				<td class="content_right" nowrap>
					<%= IIF(rsp("prz_var_euro")<>0, "<b>", "") %>
					<%= FormatPrice(rsp("prz_var_euro"), 2, true) %> &euro;
					<%= IIF(rsp("prz_var_euro")<>0, "</b>", "") %>
				</td>
				<td class="content_right" nowrap>
					<%= IIF(rsp("prz_var_sconto")<>0, "<b>", "") %>
					<%= FormatPrice(rsp("prz_var_sconto"), 2, true) %> %
					<%= IIF(rsp("prz_var_sconto")<>0, "</b>", "") %>
				</td>
				<td class="content_right" title="<%=rsp("iva_valore")%>"><%= rsp("iva_nome") %></td>
				<td class="content_center"><%= rsp("scc_nome") %></td>
			</tr>
			<% rsp.movenext
		wend 
		rsp.close
		rsl.movenext
	wend
	rsl.close

end sub
%>