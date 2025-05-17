<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeOut = 100000 %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE file="ListiniPrezzi_Tools.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	

dim conn, rs, rsl, rsv, sql, listino, sql_where
dim prezzo_base, prezzo_attuale, prezzo_nuovo, iva_id, scontoQ_id, var_sconto, var_euro, visibile, promozione, personalizzato, offerta_dal, offerta_al
dim style
set conn = Server.CreateObject("ADODB.Connection")
conn.CommandTimeout = 120
conn.ConnectionTimeout = 120
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsl = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")

listino = cInteger(request.querystring("ID"))

if request.querystring("goto")<>"" then
	if session("B2B_PREZZI_LISTINI_SEARCHED_ALL") then
		session("B2B_PREZZI_LISTINI_SEARCHED") = false
	end if
	CALL GotoRecord(conn, rs, Session("B2B_LISTINI_SQL"), "listino_id", "ListiniPrezzi_Gruppi.asp")
end if

sql = " SELECT *, (SELECT listino_id FROM gtb_listini WHERE listino_base_attuale=1) AS LB_ATTUALE " + _
	  " FROM gtb_listini WHERE listino_id=" & listino
rsl.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

dim Pager
set Pager = new PageNavigator

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if (request.form("cerca")<>"" or request.form("tutti")<>"") then
		'richiesta ricerca o "vedi tutti"
		Pager.Reset()
		CALL SearchForm_Init()
	elseif request.form("applica")<>"" then
		if VariazioniRiga_Valide(request) then
			'applicazione delle variazioni
			conn.begintrans 
			sql_where = SearchForm_Parse(categorie)
			CALL Listino_OpenRowRecordset(rsl, rs, listino, sql_where)

			while not rs.eof
				CALL Listino_StatoRiga(rs, rsl, listino, prezzo_base, prezzo_nuovo, _
													     prezzo_attuale, var_sconto, var_euro, personalizzato, _
											   			 iva_id, scontoQ_id, visibile, promozione, offerta_dal, offerta_al)

				CALL VariazioniRiga_Applica(rsl, request, prezzo_base, prezzo_nuovo, var_sconto, var_euro, scontoQ_id, visibile, promozione, offerta_dal, offerta_al)
				
				CALL Listino_SalvaRiga(conn, rsl, rsv, listino, rs("prz_variante_id"), var_sconto, var_euro, iva_id, scontoQ_id, visibile, promozione, prezzo_nuovo, personalizzato, offerta_dal, offerta_al)
				rs.movenext
			wend
			rs.close
			
			conn.committrans
			CALL VariazioniRiga_Reset()
			response.redirect "ListiniPrezzi_Gruppi.asp?ID=" & request("ID")
		end if
	elseif request.form("ANTEPRIMA")<>"" then
		'imposta le variabili di sessione con le nuove variazioni da applicare alle righe
		Session("variazione_prezzo") = request("variazione_prezzo")
		Session("variazione_percentuale") = request("variazione_percentuale")
		Session("variazione_applica_da") = request("variazione_applica_da")
		Session("variaizone_classe_sconto") = request("variaizone_classe_sconto")
		Session("stato_visibile") = request("stato_visibile")		
		Session("stato_promozione") = request("stato_promozione")
		Session("offerta_dal") = request("offerta_dal")
		Session("offerta_al") = request("offerta_al")
	elseif request.form("ANNULLA")<>"" then
		'imposta le variabili di sessione annullando le variazioni
		CALL VariazioniRiga_Reset()
	end if
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione listini - gestione per gruppi"
if rsl("listino_importato") AND not IsAdminCurrent(conn) then
	dicitura.puls_new = "INDIETRO;SCHEDA LISTINO;"
	dicitura.link_new = "Listini.asp;ListiniMod.asp?ID=" & listino
else
	dicitura.puls_new = "INDIETRO;SCHEDA LISTINO;gestione prezzi:;RIGA PER RIGA;AVANZATA"
	dicitura.link_new = "Listini.asp;ListiniMod.asp?ID=" & listino & ";;ListiniPrezzi_RigaPerRiga.asp?ID=" & request("ID") & ";ListiniPrezzi_Avanzata.asp?ID=" & request("ID")
end if
dicitura.scrivi_con_sottosez()

sql_where = SearchForm_Parse(categorie)%>
<div id="content_liquid">
	<% CALL SearchForm_Write(categorie, conn)%>
	<script language="JavaScript" type="text/javascript">
		function aggiorna(tag) {
		var value = toNumber(tag.value);
		if (tag.name == "variazione_prezzo" && !isNaN(value)){
			//immesso prezzo
			document.form_impostazioni.variazione_percentuale.value = "";
			document.form_impostazioni.variazione_prezzo.value = FormatNumber(value, 2);
		}
		else if (!isNaN(value)) {
			//immessa variazione percentuale
			document.form_impostazioni.variazione_percentuale.value = FormatNumber(value, 2);
			document.form_impostazioni.variazione_prezzo.value = "";
		}
	}
		
		//variabile utilizzata per il controllo del submit nel form per le variazioni su prezzi
		var ClickApplica = false;
		
		function verifica_intenzioni_applica(){
			if (ClickApplica){
				ClickApplica = false;
				return window.confirm('ATTENZIONE: le modifiche sono DEFINITIVE e NON REVERSIBILI. \n' + 
									  'Applicare le variazioni?')
			}
			else{
				ClickApplica = false;
				return true;
			}
		}
	</script>
	<% if session("B2B_PREZZI_LISTINI_SEARCHED") then
		if not VariazioniRiga_Valide(Session) then %>
			<% if not (rsl("listino_importato") AND not IsAdminCurrent(conn)) then %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:7px; width:735px;">
					<caption class="border">Variazioni da applicare alla selezione:</caption>	
					<form action="" method="post" id="form_impostazioni" name="form_impostazioni" onsubmit="return verifica_intenzioni_applica();">
					<tr>
						<td class="label" style="width:20%;" rowspan="2">variazioni sul prezzo:</td>
						<td class="label" style="width:12%;">di &euro;:</td>
						<td class="content">
							<input type="text" tabindex="1" class="number" name="variazione_prezzo" size="7" value="<%= session("variazione_prezzo") %>" maxlength="7" onchange="aggiorna(this);">
							&euro;
						</td>
						<td class="label" rowspan="2">
							applica variazioni a:
						</td>
						<td class="content">
							<input type="radio" tabindex="2" class="checkbox" name="variazione_applica_da" value="prezzo_base" <%= chk(instr(1, session("variazione_applica_da"), "prezzo_base", vbTextCompare))%>>
							prezzo base
						</td>
					</tr>
					<tr>
						<td class="label">in percentuale:</td>
						<td class="content">
							<input type="text" tabindex="3" class="number" name="variazione_percentuale" size="7" value="<%= session("variazione_percentuale") %>" maxlength="7" onchange="aggiorna(this);">
							%
						</td>
						<td class="content">
							<input type="radio" tabindex="4" class="checkbox" name="variazione_applica_da" value="prezzo_attuale" <%= chk(instr(1, session("variazione_applica_da"), "prezzo_attuale", vbTextCompare))%>>
							prezzo attuale
						</td>
					</tr>
					<% if (session("variazione_prezzo")<>"" OR session("variazione_percentuale")<>"") AND session("variazione_applica_da")="" then %>
						<tr>
							<td colspan="5" class="errore">Selezionare il prezzo al quale applicare la variazione.</td>
						</tr>
					<% elseif session("variazione_prezzo")="" AND session("variazione_percentuale")="" AND session("variazione_applica_da")<>"" then%>
						<tr>
							<td colspan="5" class="errore">Impostare la variazione da applicare al <%= IIF(session("variazione_applica_da")="prezzo_base", "prezzo base", "prezzo attuale") %>.</td>
						</tr>
					<% end if %>
					<tr>
						<td class="label" colspan="2">classe di sconto per quantit&agrave;:</td>
						<td class="content" colspan="4">
							<% sql = "SELECT scc_id, scc_nome FROM gtb_scontiQ_classi ORDER BY scc_nome"
							CALL DropDown(conn, sql, "scc_id", "scc_nome", "variazione_classe_sconto", session("variazione_classe_sconto"), false, "tabindex=""6""", LINGUA_ITALIANO)%>
						</td>
					</tr>
					<tr>
						<td class="label" colspan="2"><%= IIF(rsl("listino_offerte"), "stato dell'articolo:", "visibilit&agrave; dell'articolo a listino:") %></td>
						<td class="content" colspan="3">
							<input type="radio" class="checkbox" name="stato_visibile" value="1" <%= chk(cInteger(session("stato_visibile"))=1) %>>
							<%= IIF(rsl("listino_offerte"), "in offerta", "visibile") %>
							&nbsp;&nbsp;&nbsp;&nbsp;
							<input type="radio" class="checkbox" name="stato_visibile" value="0" <%= chk(cInteger(session("stato_visibile"))=0 AND session("stato_visibile")<>"") %>>
							<%= IIF(rsl("listino_offerte"), "non in offerta", "non visibile") %>
						</td>
					</tr>
					<% if not rsl("listino_offerte") then %>
						<tr>
							<td class="label" colspan="2">stato "articolo in promozione":</td>
							<td class="content" colspan="3">
								<input type="radio" class="checkbox" name="stato_promozione" value="1" <%= chk(cInteger(session("stato_promozione"))=1) %>>
								in promozione
								&nbsp;&nbsp;&nbsp;&nbsp;
								<input type="radio" class="checkbox" name="stato_promozione" value="0" <%= chk(cInteger(session("stato_promozione"))=0 AND session("stato_promozione")<>"") %>>
								non in promozione
							</td>
						</tr>
					<% else %>
						<!--
						<tr>
							<td class="label" rowspan="2">periodo offerta:</td>
							<td class="label">
								dal
							</td>
							<td class="content" colspan="3">
								<% 'CALL WriteDataPicker_Input("form_impostazioni", "offerta_dal", IIF(isDate(session("offerta_dal")) AND session("offerta_dal")<>"", session("offerta_dal"), "") _
									'					  , "", "/", true, true, LINGUA_ITALIANO) 
									%>
							</td>
						</tr>
						<tr>
							<td class="label">
								al
							</td>
							<td class="content" colspan="3">
								<% 'CALL WriteDataPicker_Input("form_impostazioni", "offerta_al", IIF(isDate(session("offerta_al")) AND session("offerta_al")<>"", session("offerta_al"), "") _
									'					  , "", "/", true, true, LINGUA_ITALIANO) 
									%>
							</td>
						</tr>
						-->
					<% end if %>
					<tr>
						<td class="footer" colspan="5" style="padding-right:0px;">
							<% if VariazioniRiga_Valide(Session) then %>
								<input type="submit" class="button" name="annulla" value="ANNULLA" style="width:9%;">
								<input type="submit" class="button" name="anteprima" value="CAMBIA VARIAZIONI" style="width:15%;">
							<% else %>
								<input type="submit" class="button" name="anteprima" value="ANTEPRIMA" style="width:11%;">
								<input type="submit" class="button" name="applica" value="APPLICA VARIAZIONI" onclick="ClickApplica=true;" style="width:17%;">
							<% end if %>
						</td>
					</tr>
					</form>
				</table>
			<% end if %>
		<% else %>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:7px; width:735px;">
				<form action="" method="post" id="form_impostazioni" name="form_impostazioni" onsubmit="return verifica_intenzioni_applica();">
				<input type="hidden" name="variazione_prezzo" value="<%= Session("variazione_prezzo") %>">
				<input type="hidden" name="variazione_percentuale" value="<%= Session("variazione_percentuale") %>">
				<input type="hidden" name="variazione_applica_da" value="<%= Session("variazione_applica_da") %>">
				<input type="hidden" name="variaizone_classe_sconto" value="<%= Session("variaizone_classe_sconto") %>">
				<input type="hidden" name="stato_visibile" value="<%= Session("stato_visibile") %>">
				<input type="hidden" name="stato_promozione" value="<%= Session("stato_promozione") %>">
				<input type="hidden" name="offerta_dal" value="<%= Session("offerta_dal") %>">
				<input type="hidden" name="offerta_al" value="<%= Session("offerta_al") %>">
				<caption class="border">Variazioni applicate in anteprima:</caption>	
				<% if (session("variazione_prezzo")<>"" OR session("variazione_percentuale")<>"") AND session("variazione_applica_da")<>"" then%>
					<tr>
						<td class="label" style="width:20%;">variazioni sul prezzo:</td>
						<% if session("variazione_prezzo")<>"" then%>
							<td class="label" style="width:16%;">di &euro;:</td>
							<td class="content_b"><%= session("variazione_prezzo") %>&euro;</td>
						<% else %>
							<td class="label" style="width:16%;">in percentuale:</td>
							<td class="content_b"><%= session("variazione_percentuale") %>%</td>
						<% end if %>
						<td class="label">applica variazioni a:</td>
						<td class="content_b"><%= IIF(session("variazione_applica_da")="prezzo_base", "prezzo base", "prezzo attuale") %></td>
					</tr>
				<% end if 
				if cInteger(Session("variazione_classe_sconto"))>0 then
					sql = "SELECT scc_nome FROM gtb_scontiQ_classi WHERE scc_id=" & session("variazione_classe_sconto")%>
					<tr>
						<td class="label" colspan="2">classe di sconto per quantit&agrave;:</td>
						<td class="content" colspan="4"><%= GetValueList(conn, rsv, sql) %></td>
					</tr>
				<% end if
				if session("stato_visibile")<>"" then %>
					<tr>
						<td class="label" colspan="2">visibilit&agrave; dell'articolo a listino:</td>
						<td class="content" colspan="3">
							<% if cInteger(session("stato_visibile"))=1 then %>
								<%= IIF(rsl("listino_offerte"), "in offerta", "visibile") %>
							<% else %>
								<%= IIF(rsl("listino_offerte"), "non in offerta", "non visibile") %>
							<% end if %>
						</td>
					</tr>
				<% end if %>
				<% if not rsl("listino_offerte") AND _
					  session("stato_promozione")<>"" then %>
					<tr>
						<td class="label" colspan="2">stato "articolo in promozione":</td>
						<td class="content" colspan="3"><%= IIF(cInteger(session("stato_promozione"))=1, "in promozione", "non in promozione") %></td>
					</tr>
				<% end if %>
				<% if rsl("listino_offerte") AND (Session("offerta_dal")<>"" OR Session("offerta_al")<>"") then %>
					<!--<tr>
						<td class="label" rowspan="2">periodo offerta:</td>
						<td class="label">
							dal
						</td>
						<td class="content" colspan="3">
							<% 'CALL WriteDataPicker_Input("form_impostazioni", "offerta_dal", IIF(isDate(session("offerta_dal")) AND session("offerta_dal")<>"", session("offerta_dal"), "") _
								'					  , "", "/", false, true, LINGUA_ITALIANO) 
													  %>
						</td>
					</tr>
					<tr>
						<td class="label">
							al
						</td>
						<td class="content" colspan="3">
							<% 'CALL WriteDataPicker_Input("form_impostazioni", "offerta_al", IIF(isDate(session("offerta_al")) AND session("offerta_al")<>"", session("offerta_al"), "") _
								'					  , "", "/", false, true, LINGUA_ITALIANO) 
								%>
						</td>
					</tr>
					-->
				<% end if %>
				<tr>
					<td class="content_b" colspan="5">Le variazioni sono state applicate in anteprima all'elenco di articoli sotto visualizzato.</td>
				</tr>
				<tr>
					<td class="footer" colspan="5" style="padding-right:0px;">
						<input type="submit" class="button" name="annulla" value="ANNULLA" style="width:9%;">
						<input type="submit" class="button" name="applica" value="APPLICA DEFINITIVAMENTE LE VARIAZIONI" onclick="ClickApplica=true;" style="width:34%;">
					</td>
				</tr>
				</form>
			</table>
		<%end if

		CALL Listino_Scheda(rsl, listino, true) 
		
		dim stile_variazioni
		if GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then
			stile_variazioni = " note"
		else
			stile_variazioni = ""
		end if
		%>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px; width:835px;">
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
					<th class="l2_center" colspan="3" style="border-bottom:0px;">VAR. / SCONTI</th>
					<% if rsl("listino_offerte") then %>
						<th class="l2_center" style="border-bottom:0px;" colspan="3">STATO</th>
					<% else %>
						<th class="l2_center" colspan="3" style="border-bottom:0px;">STATO</th>
					<% end if %>
				</tr>
				<tr>
					<th class="L2" title="codice interno dell'articolo">CODICE</th>
					<th class="L2" title="nome dell'articolo">NOME</th>
					<th style="width:5%;" class="l2_center" title="<%= IIF(rsl("listino_Base"), "prezzo base dell'articolo", "prezzo dell'articolo nel listino base") %>">BASE</th>
					<th style="width:5%;" class="l2_center" title="prezzo dell'articolo nel listino corrente">ATTUALE</th>
					<th style="width:5%;" class="l2_center" title="nuovo prezzo per l'articolo nel listino corrente">NUOVO</th>
					<th style="width:7%;" class="l2_center" title="variazione in euro del prezzo listino dal prezzo base dell'articolo">&euro;uro</th>
					<th style="width:7%;" class="l2_center" title="variazione percentuale del prezzo listino dal prezzo base dell'articolo">%</th>
					<th style="width:9%;" class="l2_center" title="classe di sconto per quantit&agrave;">CLASSE SC.</th>
					<%if rsl("listino_offerte") then %>
						<th style="width:8%;" class="l2_center" title="articolo visibile tra le offerte speciali.">IN OFFERTA</th>
						<th style="width:8%;" class="l2_center" title="data inizio offerta.">DAL</th>
						<th style="width:8%;" class="l2_center" title="data fine offerta.">AL</th>
					<% else %>
						<th style="width:3%;" class="l2_center" title="articolo disponibile all'acquisto per i clienti che usano questo listino.">VIS.</th>
						<th style="width:3%;" class="l2_center" title="articolo segnalato &quot;in promozione&quot; ai clienti che usano questo listino.">PROMO.</th>
						<th style="width:1%;" class="l2_center" title="articolo in offerta.">&nbsp;&nbsp;</th>
					<% end if %>
				</tr>
				<%rs.AbsolutePage = Pager.PageNo
				while not rs.eof and rs.AbsolutePage = Pager.PageNo
					CALL Listino_StatoRiga(rs, rsl, listino, prezzo_base, prezzo_nuovo, _
										   prezzo_attuale, var_sconto, var_euro, personalizzato, _
										   iva_id, scontoQ_id, visibile, promozione, offerta_dal, offerta_al)
					if VariazioniRiga_Valide(Session) then
						CALL VariazioniRiga_Applica(rsl, Session, prezzo_base, prezzo_nuovo, var_sconto, var_euro, scontoQ_id, visibile, promozione, offerta_dal, offerta_al)
					end if
					%>
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
							<% if prezzo_attuale <> prezzo_nuovo then%>
								<strong style="color:red;" title="prezzo variato"><%=formatPrice(prezzo_nuovo, 2, false)%></strong>
							<% else %>
								<%=formatPrice(prezzo_nuovo, 2, false)%>
							<% end if %>
						</td>
						<td class="content_right<%=stile_variazioni%>" <%= style %>>
							<% if Abs(var_euro)>0 then%>
								<strong><%=formatPrice(var_euro, 2, false)%></strong>
							<% else %>
								<%=formatPrice(var_euro, 2, false)%>
							<% end if %>
						</td>
						<td class="content_right<%=stile_variazioni%>" <%= style %>>
							<% if Abs(var_sconto)>0 then%>
								<strong><%=formatPrice(var_sconto, 2, false)%></strong>
							<% else %>
								<%=formatPrice(var_sconto, 2, false)%>
							<% end if %>
						</td>
						<td class="content" <%= style %>>
							<% if scontoQ_id>0 then 
								sql = "SELECT scc_nome FROM gtb_scontiQ_classi WHERE scc_id=" & scontoQ_id%>
								<%= GetValueList(conn, rsv, sql) %>
							<% end if %>
						</td>
						<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(visibile) %>></td>
						<% if not rsl("listino_offerte") then %>
							<td class="content_center"><input type="checkbox" class="checkbox" disabled<%= chk(promozione) %>></td>
							<td class="content_center">
								<% if cInteger(rs("OFFERTA"))>0 then %>
									<span class="Icona Offerte" title="articolo attualmente in offerta speciale">&nbsp;</span>
								<% else %>
									&nbsp;
								<% end if %>
							</td>
						<% else %>
							<td class="content_center"><%=rs("prz_offerta_dal")%></td>
							<td class="content_center"><%=rs("prz_offerta_al")%></td>
						<% end if %>
					</tr>
					<%rs.movenext
				wend%>
				<tr>
					<td class="footer" colspan="14" style="text-align:left;">
						<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
					</td>
				</tr>
			<% end if 
			rs.close %>
		</table>
	<%else
		'ricerca sui listini non eseguita
		%>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:7px;">
			<tr>
				<td class="noRecords">Per visualizzare l'elenco degli articoli eseguire prima una ricerca.</td>
			</tr>
		</table>
	<% end if %>
</div>
</body>
</html>
<% 
rsl.close
conn.close 
set rs = nothing
set rsl = nothing
set rsv = nothing
set conn = nothing


sub VariazioniRiga_Reset()
	Session("variazione_prezzo") = ""
	Session("variazione_percentuale") = ""
	Session("variazione_applica_da") = ""
	Session("variaizone_classe_sconto") = ""
	Session("stato_visibile") = ""
	Session("stato_promozione") = ""
end sub


function VariazioniRiga_Valide(object)
	if (((object("variazione_prezzo")<>"" AND IsNumeric(object("variazione_prezzo"))) OR _
	    (object("variazione_percentuale")<>"" AND IsNumeric(object("variazione_percentuale")))) AND _
	    object("variazione_applica_da")<>"" ) OR _
	   cInteger(object("variazione_classe_sconto"))>0 OR _
	   object("stato_visibile")<>"" OR _
	   object("stato_promozione")<>"" OR _
	   object("offerta_dal")<>"" OR _
	   object("offerta_al")<>"" then
		VariazioniRiga_Valide = true
	else
		VariazioniRiga_Valide = false
	end if
end function


sub VariazioniRiga_Applica(rsl, object, prezzo_base, byref prezzo_nuovo, _
									   byref var_sconto, _
									   byref var_euro, _
									   byref scontoQ_id, _
									   byref visibile, _
									   byref promozione, _
									   byref offerta_dal, _
									   byref offerta_al _
									   )
	'applica le variazioni ai dati in visualizzazione/anteprima
	if (object("variazione_prezzo")<>"" AND IsNumeric(object("variazione_prezzo"))) OR _
	   (object("variazione_percentuale")<>"" AND IsNumeric(object("variazione_percentuale"))) then
		if instr(1, object("variazione_applica_da"), "prezzo_base", vbTextCompare) then
			if object("variazione_prezzo")<>"" AND IsNumeric(object("variazione_prezzo")) then
				'variazione in euro sul prezzo base
				prezzo_nuovo = prezzo_base + cReal(object("variazione_prezzo"))
				var_sconto = 0
				var_euro = cReal(object("variazione_prezzo"))
			else
				'variazione percentuale sul prezzo base
				prezzo_nuovo = GetPricePercent(prezzo_base, cReal(object("variazione_percentuale")))
				var_euro = 0
				var_sconto = cReal(object("variazione_percentuale"))
			end if
		elseif instr(1, object("variazione_applica_da"), "prezzo_attuale", vbTextCompare) then
			if object("variazione_prezzo")<>"" AND IsNumeric(object("variazione_prezzo")) then
				'variazione in euro sul prezzo attuale
				prezzo_nuovo = prezzo_attuale + cReal(object("variazione_prezzo"))
				var_sconto = 0
				var_euro = prezzo_nuovo - prezzo_base
			else
				'variazione percentuale sul prezzo attuale
				prezzo_nuovo = GetPricePercent(prezzo_attuale, cReal(object("variazione_percentuale")))
				var_euro = 0 
				if prezzo_base = prezzo_attuale then
					var_sconto = cReal(object("variazione_percentuale"))
				else
					var_sconto = GetVarPercent(prezzo_base, prezzo_nuovo)
				end if
			end if
		end if
	end if
						
	if cInteger(object("variazione_classe_sconto"))>0 then
		scontoQ_id = cInteger(object("variazione_classe_sconto"))
	end if
						
	if object("stato_visibile")<>"" then
		visibile = ( cInteger(object("stato_visibile"))>0 )
	end if
	
	if not rsl("listino_offerte") then
		if object("stato_promozione")<>"" then
			promozione = ( cInteger(object("stato_promozione"))>0)
		end if
	else
		if object("offerta_dal")<>"" then
			offerta_dal = object("offerta_dal")
		end if
		if object("offerta_al")<>"" then
			offerta_al = object("offerta_al")
		end if
	end if
end sub
%>