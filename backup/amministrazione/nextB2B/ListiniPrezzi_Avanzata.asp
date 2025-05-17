<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeOut = 100000 %>
<!--#INCLUDE file="ListiniPrezzi_Tools.asp" -->
<% 	

const LIMITE_ARTICOLI = 600

dim conn, rs, rsl, rsv, rs_scc, sql, listino, sql_where
dim prezzo_base, prezzo_attuale, prezzo_nuovo, iva_id, scontoQ_id, var_sconto, var_euro, visibile, promozione, personalizzato, offerta_dal, offerta_al
dim style, show_scc, variante_id
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
	CALL GotoRecord(conn, rs, Session("B2B_LISTINI_SQL"), "listino_id", "ListiniPrezzi_Avanzata.asp")
end if

sql = " SELECT *, (SELECT listino_id FROM gtb_listini WHERE listino_base_attuale=1) AS LB_ATTUALE " + _
	  " FROM gtb_listini WHERE listino_id=" & listino
rsl.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if (request.form("cerca")<>"" or request.form("tutti")<>"") then
		'richiesta ricerca o "vedi tutti"
		CALL SearchForm_Init()
	else
		conn.beginTrans
		
		if request.form("salva")<>"" then
			'salva tutte le righe
			dim ids, i
			ids = split(request.form("prz_variante_id"), ",")
			for i =lBound(ids) to ubound(ids)
				variante_id = Trim(ids(i))
				CALL Listino_SalvaRiga(conn, rsl, rs, listino, variante_id, _
						   			   cReal(request.form("prz_var_sconto_" & variante_id)), _
									   cReal(request.form("prz_var_euro_" & variante_id)), _
									   cInteger(request.form("prz_iva_id_" & variante_id)), _
									   cInteger(request.form("prz_scontoQ_id_" & variante_id)), _
									   request.form("vis_" & variante_id)<>"", _
									   request.form("promo_" & variante_id)<>"", _
									   cReal(request.form("prz_prezzo_" & variante_id)), _
									   request.form("prz_personalizzato_" & variante_id)<>"", _
									   request.form("offerta_dal_" & variante_id), _
									   request.form("offerta_al_" & variante_id))
			next
		else
			'salva la riga singola
			dim var
			for each var in request.form
				if instr(1,var,"update_",vbTextCompare)>0 then
					variante_id = right(var,len(var)-len("update_"))
					CALL Listino_SalvaRiga(conn, rsl, rs, listino, variante_id, _
							   			   cReal(request.form("prz_var_sconto_" & variante_id)), _
										   cReal(request.form("prz_var_euro_" & variante_id)), _
										   cInteger(request.form("prz_iva_id_" & variante_id)), _
										   cInteger(request.form("prz_scontoQ_id_" & variante_id)), _
										   request.form("vis_" & variante_id)<>"", _
										   request.form("promo_" & variante_id)<>"", _
										   cReal(request.form("prz_prezzo_" & variante_id)), _
										   request.form("prz_personalizzato_" & variante_id)<>"", _
										   request.form("offerta_dal_" & variante_id), _
										   request.form("offerta_al_" & variante_id))
				end if
			next 
		end if
		
		conn.committrans
		
		response.redirect "ListiniPrezzi_Avanzata.asp?ID=" & request("ID")
	end if
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione listini - gestione avanzata"
dicitura.puls_new = "INDIETRO;SCHEDA LISTINO;gestione prezzi:;PER GRUPPI;RIGA PER RIGA"
dicitura.link_new = "Listini.asp;ListiniMod.asp?ID=" & listino & ";;ListiniPrezzi_Gruppi.asp?ID=" & request("ID") & ";ListiniPrezzi_RigaPerRiga.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez()  

sql_where = SearchForm_Parse(categorie)%>

<script language="JavaScript" type="text/javascript">
<!--
	//funzioni per il box di ricerca
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
	
	
	function ApplicaVariazioni(nome_prezzo_calcolo){
		var prezzo_base_calcolo, tag_prezzo_base_calcolo, prezzo_new, prezzo_base, row_id;
		var variazione_prz, variazione_perc;
		var prezzo_var_euro, prezzo_var_perc;
						
		variazione_prz = toNumber(document.form_impostazioni.variazione_prezzo.value);
		variazione_perc = toNumber(document.form_impostazioni.variazione_percentuale.value);
		
		if (!(isNaN(variazione_prz) && isNaN(variazione_perc))){
			//scorre tutte le righe del lisino
			for(var i=0; i<form_prezzi.prz_variante_id.length; i++){
				row_id = form_prezzi.prz_variante_id(i).value;
				tag_prezzo_base_calcolo = eval('document.form_prezzi.' + nome_prezzo_calcolo + row_id)
				prezzo_base_calcolo = toNumber(tag_prezzo_base_calcolo.value)
				
				prezzo_base = eval('document.form_prezzi.prz_prezzo_base_' + row_id)
				prezzo_new = eval('document.form_prezzi.prz_prezzo_' + row_id)
				prezzo_var_euro = eval('document.form_prezzi.prz_var_euro_' + row_id)
				prezzo_var_perc = eval('document.form_prezzi.prz_var_sconto_' + row_id)
				//resetta vecchie impostazioni di variazione
				prezzo_var_euro.value = '0,00'
				prezzo_var_perc.value = '0,00'
				
				//aggiorna il prezzo di listino
				if (isNaN(variazione_prz) || document.form_impostazioni.variazione_prezzo.value == ''){
					//applicata variazione percentuale
					prezzo_new.value = prezzo_base_calcolo * (1 + (variazione_perc / 100)) ;
					
					//aggiorna percentuale di sconto calcolata dal prezzo base
					CalcolaVariazione(prezzo_base, prezzo_new, prezzo_var_perc);
				}
				else{
					//applicata variazione di euro
					prezzo_new.value = prezzo_base_calcolo + variazione_prz;
					
					//aggiorna variazione di &euro; calcolato da prezzo base
					CalcolaDifferenza(prezzo_base, prezzo_new, prezzo_var_euro);
					
				}
			}
		}
	}
	
	
	function ApplicaScelta(obj_scelta, tag, displacement){
		//scorre tutte le righe del lisino
		for(var i=0; i<form_prezzi.prz_variante_id.length; i++){
			//aggiorna classe di sconto per quantita'
			var classe = eval('document.form_prezzi.' + tag + form_prezzi.prz_variante_id.value);
			classe.selectedIndex = (obj_scelta.selectedIndex - displacement)
		}
	}
	
	function ApplicaFlag(obj_flag, tag){
		//scorre tutte le righe del lisino
		for(var i=0; i<form_prezzi.prz_variante_id.length; i++){
			//aggiorna il flag indicato dal parametro "tag"
			var obj_tag =  eval('document.form_prezzi.' + tag + form_prezzi.prz_variante_id(i).value);
			obj_tag.checked = obj_flag.checked;
		}
	}
	
	//funzioni per l'area di gestione dei prezzi
	function InvertiVariazioni(row){
		var value_euro, value_perc;
		var prezzo_var_euro = eval('document.form_prezzi.prz_var_euro_' + row);
		var prezzo_var_perc = eval('document.form_prezzi.prz_var_sconto_' + row);
		value_euro = toNumber(prezzo_var_euro.value);
		value_perc = toNumber(prezzo_var_perc.value);
		
		if (value_euro){
			//variazione corrente: in euro
			prezzo_var_euro.value = '0,00'
			
			//cambia la variazione in percentuale
			CalcolaVariazione(eval('document.form_prezzi.prz_prezzo_base_' + row), 
							  eval('document.form_prezzi.prz_prezzo_' + row), 
							  prezzo_var_perc);
		}
		else {
			//variazione corrente: percentuale
			prezzo_var_perc.value = '0,00'
			
			//cambia la variazione in euro
			CalcolaDifferenza(eval('document.form_prezzi.prz_prezzo_base_' + row), 
							  eval('document.form_prezzi.prz_prezzo_' + row), 
							  prezzo_var_euro);
		}
	}
	
	function ApplicaVariazioniPrezzo(tag, row){
		var altra_variazione;
		var prezzo_base = eval('document.form_prezzi.prz_prezzo_base_' + row);
		var prezzo_attuale = eval('document.form_prezzi.prz_prezzo_' + row);
		if (tag.name == ('prz_var_euro_' + row)){
			//applica variazione in euro
			altra_variazione = eval('document.form_prezzi.prz_var_sconto_' + row);
			
			//applica variazione in euro al prezzo
			CalcolaPrezzoEuro(prezzo_base, prezzo_attuale, tag)
		}
		else{
			//applica variazione in percentuale
			altra_variazione = eval('document.form_prezzi.prz_var_euro_' + row);
			
			//applica variazione percentuale al prezzo
			CalcolaPrezzo(prezzo_base, prezzo_attuale, tag)
		}
		altra_variazione.value = '0,00';
		tag.value = FormatNumber(tag.value, 2);
		
	}
	
	function RicalcolaVariazioniPrezzo(prezzo_attuale, row){
		var value_euro, value_perc;
		var prezzo_base = eval('document.form_prezzi.prz_prezzo_base_' + row);
		var prezzo_var_euro = eval('document.form_prezzi.prz_var_euro_' + row)
		var prezzo_var_perc = eval('document.form_prezzi.prz_var_sconto_' + row)
		
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
<%'apre recordset sconti per quantità
sql = "SELECT scc_id, scc_nome FROM gtb_scontiQ_classi ORDER BY scc_nome"
rs_scc.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
show_scc = not rs_scc.eof %>
<div id="content">
	<% 
	CALL SearchForm_Write(categorie, conn)
	CALL Listino_Scheda(rsl, listino, true) 
	%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="4">AGGIORNAMENTO PREZZI DEGLI ARTICOLI SELEZIONATI</th></tr>	
		<form action="" method="post" id="form_impostazioni" name="form_impostazioni">
		<tr>
			<td class="label" style="width:22%;" rowspan="2">variazioni sul prezzo:</td>
			<td class="label" style="width:16%;">di &euro;:</td>
			<td class="content">
				<input type="text" tabindex="1" class="number" name="variazione_prezzo" size="7" maxlength="7" onchange="aggiorna(this);">
				&euro;
			</td>
			<td class="content_right" style="vertical-align:middle;" rowspan="2">
				applica variazioni a:
				<a class="button_L2" href="javascript:void(0);" onclick="ApplicaVariazioni('prz_prezzo_base_')"
				   title="calcola i prezzi del listino applicando le variazioni al prezzo base dell'articolo" <%= ACTIVE_STATUS %>>
					PREZZO BASE
				</a>
				&nbsp;
				<a class="button_L2" href="javascript:void(0);" onclick="ApplicaVariazioni('prz_prezzo_attuale_')"
				   title="calcola i prezzi del listino applicando le variazioni al prezzo attuale del listino" <%= ACTIVE_STATUS %>>
					PREZZO ATTUALE
				</a>
			</td>
		</tr>
		<tr>
			<td class="label">in percentuale:</td>
			<td class="content">
				<input type="text" tabindex="2" class="number" name="variazione_percentuale" size="7" maxlength="7" onchange="aggiorna(this);">
				%
			</td>
		</tr>
		<% if show_scc then %>
			<tr>
				<td class="label" colspan="2">classe di sconto per quantit&agrave;:</td>
				<td class="content">
					<% CALL DropDownRecordset(rs_scc, "scc_id", "scc_nome", "variazione_classe_sconto", "", false, "tabindex=""6""", LINGUA_ITALIANO)%>
				</td>
				<td class="content_right" style="vertical-align:middle;">
					<a class="button_L2" href="javascript:void(0);" onclick="ApplicaScelta(document.form_impostazioni.variazione_classe_sconto, 'prz_scontoQ_id_', 0)"
					   title="sostituisce la classe di sconto di tutti gli articoli selezionati" <%= ACTIVE_STATUS %>>
						APPLICA
					</a>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" colspan="2">visibilit&agrave; dell'articolo a listino:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="stato_visibile">
			</td>
			<td class="content_right" style="vertical-align:middle;">
				<a class="button_L2" href="javascript:void(0);" onclick="ApplicaFlag(document.form_impostazioni.stato_visibile, 'vis_')"
				   title="lo stato di visibilit&agrave; in tutti gli articoli selezionati" <%= ACTIVE_STATUS %>>
					APPLICA
				</a>
			</td>
		</tr>
		<tr>
			<td class="label" colspan="2">stato "articolo in promozione":</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="stato_promozione">
			</td>
			<td class="content_right" style="vertical-align:middle;">
				<a class="button_L2" href="javascript:void(0);" onclick="ApplicaFlag(document.form_impostazioni.stato_promozione, 'promo_')"
				   title="lo stato di promozione in tutti gli articoli selezionati" <%= ACTIVE_STATUS %>>
					APPLICA
				</a>
			</td>
		</tr>
	</form>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px">
		<%if session("B2B_PREZZI_LISTINI_SEARCHED") then
			CALL Listino_OpenRowRecordset(rsl, rs, listino, sql_where)
			if rs.eof then	%>
				<tr>
					<td class="noRecords">Nessun articolo trovato</td>
				</tr>
			<% elseif rs.recordcount > LIMITE_ARTICOLI then %>
				<tr><th>Elenco prezzi del listino</th></tr>
				<tr>
					<td class="label_no_width">
						Selezionati n&ordm; <%= rs.recordcount %> articoli
					</td>
				</tr>
				<tr>
					<td class="noRecords">Raggiunto limite massimo di articoli modificabili.<br>Selezionare un massimo di <%= LIMITE_ARTICOLI %> articoli.</td>
				</tr>
				<tr>
					<td class="note">
						Questa procedura di gestione dei prezzi del listino permette la modifica di un gruppo composto da massimo <%= LIMITE_ARTICOLI %> articoli.<br>
						Per eseguire modifiche a gruppi di articoli pi&ugrave; numerosi utilizzare le procedure di "modifica per gruppi", o "modifica riga per riga".
					</td>
				</tr>
			<% else %>
				<tr><th colspan="14">Elenco prezzi del listino</th></tr>
				<tr>
					<td class="label" colspan="2">
						Selezionati n&ordm; <%= rs.recordcount %> articoli
					</td>
					<td class="content_right" colspan="12">
						<a class="button_L2" href="javascript:void(0);" onclick="ResetAll();"
						   title="annulla tutte le modifiche apportate ai dati ma non ancora salvate" <%= ACTIVE_STATUS %>>
							ANNULLA MODIFICHE
						</a>
					</td>
				</tr>
			<form action="" method="post" id="form_prezzi" name="form_prezzi">
				<tr>
					<th class="L2" colspan="2" style="border-bottom:0px;" title="codice interno dell'articolo">ARTICOLO</th>
					<th class="l2_center" colspan="3" style="border-bottom:0px;">PREZZI in &euro;uro</th>
					<th class="l2_center" colspan="<%= IIF(show_scc, 3, 2) %>" style="border-bottom:0px;">VAR. / SCONTI</th>
					<% if rsl("listino_offerte") then %>
						<th class="l2_center" style="border-bottom:0px;">STATO</th>
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
					<th style="width:3%;" class="l2_center" title="variazione percentuale del prezzo listino dal prezzo base dell'articolo">%</th>
					<% if show_scc then %>
						<th style="width:5%;" class="l2_center" title="classe di sconto per quantit&agrave;">CLASSE SC.</th>
					<% end if
					if rsl("listino_offerte") then %>
						<th style="width:6%;" class="l2_center" title="articolo visibile tra le offerte speciali.">IN OFFERTA</th>
					<% else %>
						<th style="width:3%;" class="l2_center" title="articolo disponibile all'acquisto per i clienti che usano questo listino.">VIS.</th>
						<th style="width:2%;" class="l2_center" title="articolo segnalato &quot;in promozione&quot; ai clienti che usano questo listino.">PROMO.</th>
						<th style="width:1%;" class="l2_center" title="articolo in offerta.">&nbsp;&nbsp;</th>
					<% end if %>
				</tr>
				<% while not rs.eof and rs.absoluteposition<=LIMITE_ARTICOLI
					if rs.absoluteposition mod 500 = 0 then
						response.flush
					end if
					CALL Listino_StatoRiga(rs, rsl, listino, prezzo_base, prezzo_nuovo, _
										   prezzo_attuale, var_sconto, var_euro, personalizzato, _
										   iva_id, scontoQ_id, visibile, promozione, offerta_dal, offerta_al)
					%>
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
							<td class="content_center" <%= style %>>
								<input type="text" class="number transparent" name="prz_var_sconto_<%= rs("prz_variante_id") %>" value="<%=formatPrice(var_sconto, 2, false)%>" size="3" READONLY>
							</td>
						<% else %>
							<td class="content_center" <%= style %>>
								<input type="text" class="number" name="prz_var_euro_<%= rs("prz_variante_id") %>" value="<%=formatPrice(var_euro, 2, false)%>" size="2" onchange="ApplicaVariazioniPrezzo(this, '<%= rs("prz_variante_id") %>')">
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
						<% end if %>
						<td class="Content_center">
							<input type="submit" class="button_L2" name="update_<%= rs("prz_variante_id") %>" value="SALVA">
						</td>
					</tr>
					<%rs.movenext
				wend%>
				<tr>
					<td class="footer" colspan="14">
						<input type="reset" class="button" name="annulla" value="ANNULLA MODIFICHE" onclick="ResetAll();">
						<input type="submit" class="button" name="salva" value="SALVA TUTTI">
					</td>
				</tr>
			</form>
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