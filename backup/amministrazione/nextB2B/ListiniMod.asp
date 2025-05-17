<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->

<%
dim conn, rs, sql, i, listino_tipo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_LISTINI_SQL"), "listino_id", "ListiniMod.asp")
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" AND (request("salva")<>"" OR request("salva_elenco")<>"") then
	Server.Execute("ListiniSalva.asp")
	
elseif request("imposta_listino_base")<>"" then
	'imposta il listino corrente come listino base
	dim ListinoBaseOld, ListinoBaseNew
	
	'recupera id listini in gioco
	ListinoBaseNew = cInteger(request("ID"))
	sql = "SELECT listino_id FROM gtb_listini WHERE listino_Base_attuale=1"
	ListinoBaseOld = cInteger(GetValueList(conn, rs, sql))
	
	if ListinoBaseNew <> ListinoBaseOld then
		conn.begintrans
		
		'aggiorna tutti i prezzi dei listini clienti dipendenti dal listino base: solo se la modalità di aggiornamento listini lo prevede
		if not GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then
			sql = "UPDATE gtb_prezzi SET prz_prezzo = ((SELECT prz_prezzo FROM gtb_prezzi base " + _
				  " WHERE base.prz_variante_id=gtb_prezzi.prz_variante_id AND base.prz_listino_id=" & ListinoBaseNew & ") + " + _
				  " prz_var_euro + ((SELECT prz_prezzo FROM gtb_prezzi base " + _
				  " WHERE base.prz_variante_id=gtb_prezzi.prz_variante_id AND base.prz_listino_id=" & ListinoBaseNew & ") / 100 * prz_var_sconto)) " + _
				  " WHERE prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE listino_base=0)"
			CALL conn.execute(sql, , adexecuteNoRecords)
		end if
		
		'aggiorna collegamento listini con i clienti:
		sql = "UPDATE gtb_rivenditori SET riv_listino_id=" & ListinoBaseNew & " WHERE riv_listino_id=" & ListinoBaseOld
		CALL conn.execute(sql, , adExecuteNoRecords)
		
		sql = "UPDATE gtb_listini SET listino_base_attuale=0"
		CALL conn.execute(sql, , adexecuteNoRecords)
		
		sql = "UPDATE gtb_listini SET listino_Base_attuale=1 WHERE listino_id=" & ListinoBaseNew
		CALL conn.execute(sql, , adexecuteNoRecords)
		
		conn.committrans
	end if
elseif request.querystring("sovrascrivi_prezzi_default")<>"" then
	'sovrascrive prezzi correnti del listino con le variazioni di default, impostando anche quelli mancanti dal listino base attuale
	conn.begintrans
	
	CALL UpdateListinoFromVariazioniDefault(conn, cIntero(request("ID")), true)
	
	conn.committrans
end if
 	

sql = "SELECT * FROM gtb_listini WHERE listino_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione listini / prezzi - modifica"
if rs("listino_importato") AND not IsAdminCurrent(conn) then
	dicitura.puls_new = "INDIETRO;gestione prezzi:;PER GRUPPI;"
	dicitura.link_new = "Listini.asp;;ListiniPrezzi_Gruppi.asp?ID=" & request("ID")
else
	dicitura.puls_new = "INDIETRO;gestione prezzi:;PER GRUPPI;RIGA PER RIGA;AVANZATA"
	dicitura.link_new = "Listini.asp;;ListiniPrezzi_Gruppi.asp?ID=" & request("ID") & ";ListiniPrezzi_RigaPerRiga.asp?ID=" & request("ID") & ";ListiniPrezzi_Avanzata.asp?ID=" & request("ID")
end if
dicitura.scrivi_con_sottosez() 


if rs("listino_base") then
	listino_tipo = "BASE"
elseif rs("listino_offerte") then
	listino_tipo = "OFFERTE SPECIALI"
else
	listino_tipo = ""
end if
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del listino <%= lcase(listino_tipo) %></td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="listino precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="listino successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI DEL LISTINO</th></tr>
		<% if cBoolean(rs("listino_importato"), false) then %>
			<tr>
				<td align="center" class="bundle" colspan="4">
					LISTINO IMPORTATO - ultimo aggiornamento <%= rs("listino_Dataimport")%>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" style="width:20%;">Codice:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_listino_codice" value="<%= rs("listino_codice") %>" maxlength="50" size="50">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" style="width:20%;">Nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_listino_nome_it" value="<%= rs("listino_nome_it") %>" maxlength="505" size="115">
			</td>
		</tr>
		<% select case listino_tipo
			case "BASE" %>
				<input type="hidden" name="chk_listino_base" value="1">
				<input type="hidden" name="chk_listino_offerte" value="">
				<tr>
					<td class="label">listino base:</td>
					<td class="content"><input type="checkbox" class="noborder" checked disabled></td>
					<td class="content_right">
						<% if rs("listino_importato") AND not IsAdminCurrent(conn) then %>
							&nbsp;
						<% else %>
							<input <%= disable(rs("listino_Base_attuale")) %> type="submit" name="imposta_listino_base" tabindex="100" value="IMPOSTA COME LISTINO BASE ATTUALE" class="button" style="width:220px;" title="<%= IIF(rs("listino_Base_attuale"), "listino gi&agrave; impostato come listino base in vigore", "click per impostare il listino corrente come listino attuale") %>">
						<% end if %>
					</td>
				</tr>
			<% case "OFFERTE SPECIALI" %>
				<input type="hidden" name="chk_listino_base" value="">
				<input type="hidden" name="chk_listino_base_attuale" value="">
				<input type="hidden" name="chk_listino_offerte" value="1">
				<input type="hidden" name="chk_listino_b2c" value="">
				<tr>
					<td class="label">listino offerte speciali:</td>
					<td class="content" colspan="2"><input type="checkbox" class="noborder" checked disabled></td>
				</tr>
				<% if IsNull(rs("listino_datacreazione")) AND isNull(rs("listino_datascadenza")) then %>
					<tr>
						<td class="label">periodo validit&agrave; offerte:</td>
						<td class="content_b" colspan="2">
							impostato per ogni articolo
						</td>
					</tr>
				<% else %>
					<tr>
						<td class="label">dal:</td>
						<td class="content">
							<% CALL WriteDataPicker_Input("form1", "tfd_listino_datacreazione", rs("listino_datacreazione"), "", "/", true, true, LINGUA_ITALIANO) %>
						</td>
					</tr>
					<tr>
						<td class="label">fino al:</td>
						<td class="content">
							<% CALL WriteDataPicker_Input("form1", "tfd_listino_datascadenza", rs("listino_datascadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
						</td>
					</tr>
				<% end if
			case else %>
				<input type="hidden" name="chk_listino_base" value="">
				<input type="hidden" name="chk_listino_base_attuale" value="">
				<input type="hidden" name="chk_listino_offerte" value="">
				<input type="hidden" name="listino_ancestor_id" value="<%= rs("listino_ancestor_id") %>">
				<tr><th colspan="3">COMPORTAMENTO LISTINO</th></tr>
				<tr>
					<td class="label">modalit&agrave;:</td>
					<td class="content" colspan="3">
						<% if cInteger(rs("listino_ancestor_id"))>0 then %>
							<strong>Listino DERIVATO</strong> - prezzi derivati da altro listino:<br>
							<% sql = " SELECT listino_id, listino_codice FROM gtb_listini " + _
									 " WHERE listino_offerte=0 AND listino_base=0 AND IsNull(listino_ancestor_id,0)=0 ORDER BY listino_codice"
							CALL dropDown(conn, sql, "listino_id", "listino_codice", "tfn_listino_ancestor_id", rs("listino_ancestor_id"), true, "", LINGUA_ITALIANO)%>
						<% else %>	
							<strong>Listino PRINCIPALE</strong>: prezzi gestiti in modo indipendente.
						<% end if %>
					</td>
				</tr>
				<% if cInteger(rs("listino_ancestor_id"))=0 then %>
					
					<script type="text/javascript" language="JavaScript">
									
						function ApplicaVariazioniPrezzo(tag){
							var altra_variazione;
							if (tag.name == ('tfn_listino_default_var_euro')){
								//applica variazione in euro
								altra_variazione = form1.tfn_listino_default_var_sconto;
							}
							else{
								//applica variazione in percentuale
								altra_variazione = form1.tfn_listino_default_var_euro;
							}
							altra_variazione.value = '0,00';
							tag.value = FormatNumber(tag.value, 2);
							
						}	
						
						function Esegui_aggiornamento_prezzi(){
							if (window.confirm("Applicare le variazioni di default a TUTTI i prezzi presenti nel listino?\n ATTENZIONE: tutte le variazioni esistenti verranno sovvrascritte ed i prezzi ricalcolati.")){
								document.location = "<%=GetPageName()%>?ID=<%=request("ID")%>&sovrascrivi_prezzi_default=1";
							}
						}
					</script>
					<tr>
						<td class="label" rowspan="3">
							Variazioni di default:
						</td>
						<td class="content note" colspan="2">
							La modifica delle variazioni di default non ha effetto sulle variazioni gi&agrave; registrate nel listino, ma solo sui nuovi inserimenti di articoli.
							E' possibile aggiornare le variazioni esistenti attraverso l'apposita procedura:
							<a class="button_L2" href="javascript:void(0);"  style="float:right;" onclick="Esegui_aggiornamento_prezzi()">
							SOVRASCRIVI TUTTE LE VARIAZIONI DEI PREZZI</a>
						</td>
					</tr>
					<tr>
						<td class="label_no_width">
							in euro:
						</td>
						<td class="content">
							<input type="text" class="number" name="tfn_listino_default_var_euro" id="tfn_listino_default_var_euro" value="<%= FormatPrice(rs("listino_default_var_euro") , 2, false) %>" size="4" onchange="ApplicaVariazioniPrezzo(this)"> &euro;
						</td>
					</tr>
					<tr>
						<td class="label_no_width">
							in percentuale:
						</td>
						<td class="content">
							<input type="text" class="number" name="tfn_listino_default_var_sconto" id="tfn_listino_default_var_sconto" value="<%= FormatPrice(rs("listino_default_var_sconto") , 2, false) %>" size="4" onchange="ApplicaVariazioniPrezzo(this)"> %
						</td>
					</tr>
				<% end if
		end select 
		if listino_tipo <> "OFFERTE SPECIALI" then%>
			<tr><th colspan="3">LISTINO AL PUBBLICO</th></tr>
			<tr>
				<td class="label">listino al pubblico:</td>
				<td class="content">
					<input type="checkbox" name="chk_listino_b2c" class="noborder" <%= chk(rs("listino_B2C")) %> onclick="EnableIfChecked(this, document.form1.tfd_listino_datacreazione);EnableIfChecked(this, document.form1.tfd_listino_datascadenza);">
				</td>
				<td class="note">
					Listino visibile agli utenti non registrati e del B2C.
				</td>
			</tr>
			<tr>
				<td class="label" rowspan="2">periodo pubblicazione:</td>
				<td class="label">dal:</td>
				<td class="content">
					<% CALL WriteDataPicker_Input("form1", "tfd_listino_datacreazione", rs("listino_datacreazione"), "", "/", true, true, LINGUA_ITALIANO) %>
				</td>
			</tr>
			<tr>
				<td class="label">fino al:</td>
				<td class="content">
					<% CALL WriteDataPicker_Input("form1", "tfd_listino_datascadenza", rs("listino_datascadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
				</td>
			</tr>
			<script type="text/javascript" language="JavaScript">
				form1.chk_listino_b2c.onclick();
			</script>
		<% end if %>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_listino_note"><%= rs("listino_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				<% if rs("listino_importato") AND not IsAdminCurrent(conn) then %>
					&nbsp;
				<% else %>
					(*) Campi obbligatori.
					<input type="submit" class="button" name="salva" value="SALVA">
					<input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA AD ELENCO" style="width:22%;">
				<% end if %>
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% if cBoolean(rs("listino_importato"), false) AND not IsAdminCurrent(conn) then %>
	<script language="JavaScript" type="text/javascript">
		for (i=0; i<form1.length; i++){
			DisableControl(form1.elements[i], true);
		}
		
		var links = form1.all.tags("A");
		for (i=0; i<links.length; i++){
			if (links[i].title == 'SCEGLI' || links[i].title == 'RESET')
				DisableControl(links[i], true)
		}
		
	</script>
<% end if %>

<%
set rs = nothing
conn.Close
set conn = nothing
%>