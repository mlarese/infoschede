<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeOut = 100000 %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE file="ListiniPrezzi_Tools.asp" -->
<% 	

dim conn, rs, rsl, rsv, rs_scc, sql, listino, sql_where
dim prezzo_base, prezzo_attuale, prezzo_nuovo, iva_id, scontoQ_id, var_sconto, var_euro, visibile, promozione, personalizzato, offerta_dal, offerta_al
dim style, show_scc
set conn = Server.CreateObject("ADODB.Connection")
conn.CommandTimeout = 120
conn.ConnectionTimeout = 120
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsl = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
set rs_scc = Server.CreateObject("ADODB.RecordSet")

listino = cInteger(request.querystring("ID"))

if request.querystring("goto")<>"" then
	if session("B2B_PREZZI_LISTINI_SEARCHED_ALL") then
		session("B2B_PREZZI_LISTINI_SEARCHED") = false
	end if
	CALL GotoRecord(conn, rs, Session("B2B_LISTINI_SQL"), "listino_id", "ListiniPrezzi_RigaPerRiga.asp")
end if

sql = " SELECT *, (SELECT listino_id FROM gtb_listini WHERE listino_base_attuale=1) AS LB_ATTUALE " + _
	  " FROM gtb_listini WHERE listino_id=" & listino
rsl.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

if request.querystring("list_offerte")="1" then
	Session("list_offerte") = 1 
end if

dim Pager
set Pager = new PageNavigator

if (request.form("cerca")<>"" or request.form("tutti")<>"") then
	'richiesta ricerca o "vedi tutti"
	Pager.Reset()
	CALL SearchForm_Init()
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione listini - gestione riga per riga"
dicitura.puls_new = "INDIETRO;SCHEDA LISTINO;gestione prezzi:;PER GRUPPI;AVANZATA"
dicitura.link_new = "Listini.asp;ListiniMod.asp?ID=" & listino & ";;ListiniPrezzi_Gruppi.asp?ID=" & request("ID") & ";ListiniPrezzi_Avanzata.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez()  

sql_where = SearchForm_Parse(categorie)

'apre recordset sconti per quantità
sql = "SELECT scc_id, scc_nome FROM gtb_scontiQ_classi ORDER BY scc_nome"
rs_scc.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
show_scc = not rs_scc.eof%>

<script language="JavaScript" type="text/javascript">
<!--
	//funzioni per il box di ricerca
		
	//funzioni per l'area di gestione dei prezzi
	function InvertiVariazioni(row){
		var value_euro, value_perc;
		var prezzo_var_euro = eval('document.form_prezzi_' + row + '.prz_var_euro_' + row);
		var prezzo_var_perc = eval('document.form_prezzi_' + row + '.prz_var_sconto_' + row);
		value_euro = toNumber(prezzo_var_euro.value);
		value_perc = toNumber(prezzo_var_perc.value);
		
		if (value_euro){
			//variazione corrente: in euro
			prezzo_var_euro.value = '0,00'
			
			//cambia la variazione in percentuale
			CalcolaVariazione(eval('document.form_prezzi_' + row + '.prz_prezzo_base_' + row), 
							  eval('document.form_prezzi_' + row + '.prz_prezzo_' + row), 
							  prezzo_var_perc);
		}
		else {
			//variazione corrente: percentuale
			prezzo_var_perc.value = '0,00'
			
			//cambia la variazione in euro
			CalcolaDifferenza(eval('document.form_prezzi_' + row + '.prz_prezzo_base_' + row), 
							  eval('document.form_prezzi_' + row + '.prz_prezzo_' + row), 
							  prezzo_var_euro);
		}
	}
	
	function ApplicaVariazioniPrezzo(tag, row){
		var altra_variazione;
		var prezzo_base = eval('document.form_prezzi_' + row + '.prz_prezzo_base_' + row);
		var prezzo_attuale = eval('document.form_prezzi_' + row + '.prz_prezzo_' + row);
		if (tag.name == ('prz_var_euro_' + row)){
			//applica variazione in euro
			altra_variazione = eval('document.form_prezzi_' + row + '.prz_var_sconto_' + row);
			
			//applica variazione in euro al prezzo
			CalcolaPrezzoEuro(prezzo_base, prezzo_attuale, tag)
		}
		else{
			//applica variazione in percentuale
			altra_variazione = eval('document.form_prezzi_' + row + '.prz_var_euro_' + row);
			
			//applica variazione percentuale al prezzo
			CalcolaPrezzo(prezzo_base, prezzo_attuale, tag)
		}
		altra_variazione.value = '0,00';
		tag.value = FormatNumber(tag.value, 2);
		
	}
	
	function RicalcolaVariazioniPrezzo(prezzo_attuale, row){
		var value_euro, value_perc;
		var prezzo_base = eval('document.form_prezzi_' + row + '.prz_prezzo_base_' + row);
		var prezzo_var_euro = eval('document.form_prezzi_' + row + '.prz_var_euro_' + row)
		var prezzo_var_perc = eval('document.form_prezzi_' + row + '.prz_var_sconto_' + row)
		
		value_euro = toNumber(prezzo_var_euro.value);
		value_perc = toNumber(prezzo_var_perc.value);
		
		if (value_euro){
			//presente una variazione espressa in euro
			//azzera variazione in percentuale
			prezzo_var_perc.value = '0,00';
			//ricalcola variazione in euro
			CalcolaDifferenza(prezzo_base, prezzo_attuale, prezzo_var_euro);
		}
		else{
			//presente variazione in percentuale o nessuna variazione preesistente
			//azzera variazione in euro
			prezzo_var_euro.value='0,00';
			//ricalcola variazione in percentuale
			CalcolaVariazione(prezzo_base, prezzo_attuale, prezzo_var_perc);
		}
		
	}
	
	function ResetAll(){
		form_prezzi.reset(); 
		form_impostazioni.reset();
	}
	
//-->
</script>
<div id="content_liquid">
	<% 
	CALL SearchForm_Write(categorie, conn)
	CALL Listino_Scheda(rsl, listino, true) %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px; width:1050px;">
		<%if session("B2B_PREZZI_LISTINI_SEARCHED") then%>
			<%sql = Listino_GetRowQuery(rsl, listino, sql_where)

			CALL Pager.OpenSmartRecordset(conn, rs, sql, ARTICOLI_PER_PAGINA)
			if rs.eof then	%>
				<tr>
					<td class="noRecords">Nessun articolo trovato</td>
				</tr>
			<% else %>
				<tr><th colspan="14">Elenco prezzi del listino</th></tr>
				<tr>
					<td class="label" colspan="14">
						Selezionati n&ordm; <%= Pager.recordcount %> articoli in n&ordm; <%= Pager.PageCount %> pagine
					</td>
				</tr>
				<tr>
					<th class="L2" colspan="2" style="border-bottom:0px;" title="codice interno dell'articolo">ARTICOLO</th>
					<th class="l2_center" colspan="3" style="border-bottom:0px;">PREZZI in &euro;uro</th>
					<th class="l2_center" colspan="<%= IIF(show_scc, 4, 3) %>" style="border-bottom:0px;">VAR. / SCONTI</th>
					<% if rsl("listino_offerte") then %>
						<th class="l2_center" style="border-bottom:0px;" colspan="3">STATO</th>
					<% else %>
						<th class="l2_center" colspan="3" style="border-bottom:0px;">STATO</th>
					<% end if %>
					<th class="l2_center" style="width:4%;" rowspan="2">SALVA</th>
				</tr>
				<tr>
					<th class="L2" title="codice interno dell'articolo">CODICE</th>
					<th class="L2" title="nome dell'articolo">NOME</th>
					<th style="width:3%;" class="l2_center" title="<%= IIF(rsl("listino_Base"), "prezzo base dell'articolo", "prezzo dell'articolo nel listino base") %>">BASE</th>
					<th style="width:3%;" class="l2_center" title="prezzo dell'articolo nel listino corrente">ATTUALE</th>
					<th style="width:4%;" class="l2_center" title="nuovo prezzo per l'articolo nel listino corrente">NUOVO</th>
					<th style="width:3%;" class="l2_center" title="variazione in euro del prezzo listino dal prezzo base dell'articolo">&euro;uro</th>
					<th style="width:2%;" class="l2_center" title="inverti e ricalcola tipo variazione">&nbsp;</th>
					<th style="width:3%;" class="l2_center" title="variazione percentuale del prezzo listino dal prezzo base dell'articolo">%</th>
					<% if show_scc then %>
						<th style="width:5%;" class="l2_center" title="classe di sconto per quantit&agrave;">CLASSE SC.</th>
					<% end if
					if rsl("listino_offerte") then %>
						<th style="width:6%;" class="l2_center" title="articolo visibile tra le offerte speciali.">IN OFFERTA</th>
						<th style="width:8%;" class="l2_center" title="data inizio offerta.">DAL</th>
						<th style="width:8%;" class="l2_center" title="data fine offerta.">AL</th>
					<% else %>
						<th style="width:3%;" class="l2_center" title="articolo disponibile all'acquisto per i clienti che usano questo listino.">VIS.</th>
						<th style="width:2%;" class="l2_center" title="articolo segnalato &quot;in promozione&quot; ai clienti che usano questo listino.">PROMO.</th>
						<th style="width:1%;" class="l2_center" title="articolo in offerta.">&nbsp;&nbsp;</th>
					<% end if %>
				</tr>
				<%rs.AbsolutePage = Pager.PageNo
				while not rs.eof and rs.AbsolutePage = Pager.PageNo
					CALL Listino_StatoRiga(rs, rsl, listino, prezzo_base, prezzo_nuovo, _
										   prezzo_attuale, var_sconto, var_euro, personalizzato, _
										   iva_id, scontoQ_id, visibile, promozione, offerta_dal, offerta_al)
					%>
					<form action="ListiniPrezzi_Riga_Salva.asp?ID=<%= request("ID") %>" target="salva_riga" method="post" id="form_prezzi_<%= rs("prz_variante_id") %>" name="form_prezzi_<%= rs("prz_variante_id") %>">
					<input type="hidden" name="prz_variante_id" value="<%= rs("prz_variante_id") %>">
					<input type="hidden" name="prz_iva_id_<%= rs("prz_variante_id") %>" value="<%= iva_id %>">
					<input type="hidden" name="prz_personalizzato_<%= rs("prz_variante_id") %>" value="<%= personalizzato %>">
					<input type="hidden" name="prz_prezzo_base_<%= rs("prz_variante_id") %>" value="<%= prezzo_base %>">
					<input type="hidden" name="prz_prezzo_attuale_<%= rs("prz_variante_id") %>" value="<%= prezzo_attuale %>">
					<tr>
						<% if prezzo_base <> prezzo_attuale OR personalizzato then
							if prezzo_base <> prezzo_attuale then
								if prezzo_base > prezzo_attuale then	
									'applicato sconto
									style = "#D7EFD7"
								else
									'applicata maggiorazione
									style="#FFFCD8"
								end if
								style = " style=""background-color:" + style + ";"" "
							else
								style = ""
							end if %>
							<td class="content_b" title="articolo con impostazioni e/o prezzi personalizzati">
						<% else 
							style=""%>
							<td class="content" title="articolo senza personalizzazioni nel listino">
						<% end if %>
							<%= rs("rel_cod_int") %>
						</td>
						<td class="content">
							<% CALL ArticoloLink(rs("rel_art_id"), rs("art_nome_it"), rs("rel_cod_int"))
							if rs("art_varianti") then %>
								<%= ListValoriVarianti(conn, rsv, rs("rel_id")) %>
							<% end if %>
						</td>
						<td class="content_right" <%= style %>><%= formatPrice(prezzo_base, 2, true) %></td>
						<td class="content_right" <%= style %>><%= formatPrice(prezzo_attuale, 2, true) %></td>
						<td class="content_right" <%= style %>>
							<input type="text" class="number" name="prz_prezzo_<%= rs("prz_variante_id") %>" value="<%= formatPrice(prezzo_nuovo,2, false) %>" size="7" onchange="RicalcolaVariazioniPrezzo(this, '<%= rs("prz_variante_id") %>')">
						</td>
						<% if GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then %>
							<td class="content_center" <%= style %>>
								<input type="text" class="number transparent" name="prz_var_euro_<%= rs("prz_variante_id") %>" value="<%=formatPrice(var_euro, 2, false)%>" size="2" READONLY>
							</td>
							<td class="content_center" style="vertical-align:baseline;">&nbsp;</td>
							<td class="content_center" <%= style %>>
								<input type="text" class="number transparent" name="prz_var_sconto_<%= rs("prz_variante_id") %>" value="<%=formatPrice(var_sconto, 2, false)%>" size="3" READONLY>
							</td>
						<% else %>
							<td class="content_center" <%= style %>>
								<input type="text" class="number" name="prz_var_euro_<%= rs("prz_variante_id") %>" value="<%=formatPrice(var_euro, 2, false)%>" size="2" onchange="ApplicaVariazioniPrezzo(this, '<%= rs("prz_variante_id") %>')">
							</td>
							<td class="content_center" style="vertical-align:baseline;">
								<a href="javascript:void(0)" onclick="InvertiVariazioni('<%= rs("prz_variante_id") %>')" <%= ACTIVE_STATUS %> title="Inverti tipo e ricalcola variazione della riga">
									<img src="../grafica/Frecce_Scambio.gif" alt="Inverti tipo e ricalcola variazione della riga">
								</a>
							</td>
							<td class="content_center" <%= style %>>
								<input type="text" class="number" name="prz_var_sconto_<%= rs("prz_variante_id") %>" value="<%=formatPrice(var_sconto, 2, false)%>" size="3" onchange="ApplicaVariazioniPrezzo(this, '<%= rs("prz_variante_id") %>')">
							</td>
						<% end if
						if show_scc then %>
							<td class="content" <%= style %>>
								<%CALL DropDownRecordset(rs_scc, "scc_id", "scc_nome", "prz_scontoQ_id_" & rs("prz_variante_id"), scontoQ_id , false, "", LINGUA_ITALIANO)%>
							</td>
						<% end if %>
						<td class="content_center"><input type="checkbox" class="checkbox" name="vis_<%= rs("prz_variante_id") %>" value="<%= rs("prz_variante_id") %>" <%= chk(visibile) %>></td>
						<% if not rsl("listino_offerte") then %>
							<td class="content_center"><input type="checkbox" class="checkbox" name="promo_<%= rs("prz_variante_id") %>" value="<%= rs("prz_variante_id") %>" <%= chk(promozione) %>></td>
							<td class="content_center">
								<% if cInteger(rs("OFFERTA"))>0 then %>
									<span class="Icona Offerte" title="articolo attualmente in offerta speciale">&nbsp;</span>
								<% else %>
									&nbsp;
								<% end if %>
							</td>
						<% else %>
							<td class="content_center" style="padding-right:3px;">
								<% CALL WriteDataPicker_Input("form_prezzi_"&rs("prz_variante_id"), "offerta_dal_"&rs("prz_variante_id"), IIF(isDate(rs("prz_offerta_dal")) AND rs("prz_offerta_dal")<>"", rs("prz_offerta_dal"), "") _
														  , "", "/", true, true, LINGUA_ITALIANO) %>
							</td>
							<td class="content_center">
								<% CALL WriteDataPicker_Input("form_prezzi_"&rs("prz_variante_id"), "offerta_al_"&rs("prz_variante_id"), IIF(isDate(rs("prz_offerta_al")) AND rs("prz_offerta_al")<>"", rs("prz_offerta_al"), "") _
														  , "", "/", true, true, LINGUA_ITALIANO) %>
							</td>
						<% end if %>
						<td class="Content_center">
							<input type="submit" class="button_L2" name="update_row" value="SALVA" onclick="OpenAutoPositionedScrollWindow('', 'salva_riga', 510, 200, true)">
						</td>
					</tr>
					</form>
					<%rs.movenext
				wend%>
				<tr>
					<td class="footer" colspan="14" style="text-align:left;">
						<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
					</td>
				</tr>
			<% end if 
			
			rs_scc.close
			rs.close
		else
			'ricerca sui listini non eseguita
			%>
			<tr>
				<td class="noRecords">Per visualizzare l'elenco degli articoli eseguire prima una ricerca.</td>
			</tr>
		<% end if %>
	</table>
</div>
</body>
</html>
<% 
rsl.close
conn.close 
set rs = nothing
set rsl = nothing
set rsv = nothing
set rs_scc = nothing
set conn = nothing
%>