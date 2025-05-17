<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ListiniSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione listini - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Listini.asp"
dicitura.scrivi_con_sottosez() 


dim conn, listino_base, sql, listino_tipo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")

sql = "SELECT COUNT(*) FROM gtb_listini WHERE listino_base_attuale=1"
listino_base = cInteger(GetValueList(conn, NULL, sql))

listino_tipo = request("tipo")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_listino_with_child" value="0">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo listino <%= IIF(listino_tipo<>"", lCase(listino_tipo), " clienti") %></caption>
		<tr><th colspan="4">DATI DEL LISTINO <%= listino_tipo %></th></tr>
		<tr>
			<td class="label" style="width:17%;">Codice:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_listino_codice" value="<%= request("tft_listino_codice") %>" maxlength="50" size="75">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" style="width:17%;">Nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_listino_nome_it" value="<%= request("tft_listino_nome_it") %>" maxlength="500" size="115">
			</td>
		</tr>
		<% select case listino_tipo
			case "BASE" %>
				<input type="hidden" name="chk_listino_base" value="1">
				<input type="hidden" name="chk_listino_offerte" value="">
				<% if listino_base = 0 then %>
					<input type="hidden" name="tfn_listino_base_attuale" value="1">
				<% else %>
					<input type="hidden" name="tfn_listino_base_attuale" value="0">
				<% end if %>
				<tr>
					<td class="label">listino base:</td>
					<td class="content" colspan="3"><input type="checkbox" class="noborder" checked disabled></td>
				</tr>
			<% case "OFFERTE SPECIALI" %>
				<input type="hidden" name="chk_listino_base" value="">
				<input type="hidden" name="chk_listino_base_attuale" value="">
				<input type="hidden" name="chk_listino_offerte" value="1">
				<input type="hidden" name="chk_listino_b2c" value="">
				<tr>
					<td class="label">listino offerte speciali:</td>
					<td class="content" colspan="3"><input type="checkbox" class="noborder" checked disabled></td>
				</tr>
				<tr>
					<td class="label" rowspan="3">validit&agrave; offerte:</td>
					<td class="content" colspan="3">
						<input class="checkbox" type="radio" name="gestione_validita" id="gestione_validita_articolo" value="" <%= chk(request("gestione_validita")="") %> onclick="DisablePickerIfChecked(this, document.form1.tfd_listino_datacreazione);DisablePickerIfChecked(this, document.form1.tfd_listino_datascadenza);">
						gestito su ogni articolo
					</td>
				</tr>
				<tr>
					<td class="content" rowspan="2">
						<input class="checkbox" type="radio" name="gestione_validita" value="unico" <%= chk(request("gestione_validita")="unico") %> onclick="EnablePickerIfChecked(this, document.form1.tfd_listino_datacreazione);EnablePickerIfChecked(this, document.form1.tfd_listino_datascadenza);">
						unico per l'interno listino
					</td>
					<td class="label">dal:</td>
					<td class="content">
						<% CALL WriteDataPicker_Input("form1", "tfd_listino_datacreazione", IIF(isDate(request("tfd_listino_datacreazione")) AND request("tfd_listino_datacreazione")<>"", request("tfd_listino_datacreazione"), Date) _
													  , "", "/", true, true, LINGUA_ITALIANO) %>
					</td>
				</tr>
				<tr>
					<td class="label">fino al:</td>
					<td class="content">
						<% CALL WriteDataPicker_Input("form1", "tfd_listino_datascadenza", request("tfd_listino_datascadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
					</td>
				</tr>
				<script type="text/javascript" language="JavaScript">
					form1.gestione_validita_articolo.onclick();
				</script>
			<% case else %>
				<input type="hidden" name="chk_listino_base" value="">
				<input type="hidden" name="chk_listino_base_attuale" value="">
				<input type="hidden" name="chk_listino_offerte" value="">
		<% end select 
		if listino_tipo <> "OFFERTE SPECIALI" then%>
			<tr><th colspan="4">LISTINO AL PUBBLICO</th></tr>
			<tr>
				<td class="label">listino al pubblico:</td>
				<td class="content">
					<input type="checkbox" name="chk_listino_b2c" class="noborder" <%= chk(request("chk_listino_B2C")<>"") %> onclick="EnableIfChecked(this, document.form1.tfd_listino_datacreazione);EnableIfChecked(this, document.form1.tfd_listino_datascadenza);">
				</td>
				<td class="note" colspan="2">
					Listino visibile agli utenti non registrati e del B2C.
				</td>
			</tr>
			<tr>
				<td class="label" rowspan="2">periodo pubblicazione:</td>
				<td class="label">dal:</td>
				<td class="content" colspan="2">
					<% CALL WriteDataPicker_Input("form1", "tfd_listino_datacreazione", IIF(isDate(request("tfd_listino_datacreazione")) AND request("tfd_listino_datacreazione")<>"", request("tfd_listino_datacreazione"), Date) _
												  , "", "/", true, true, LINGUA_ITALIANO) %>
				</td>
			</tr>
			<tr>
				<td class="label">fino al:</td>
				<td class="content" colspan="2">
					<% CALL WriteDataPicker_Input("form1", "tfd_listino_datascadenza", request("tfd_listino_datascadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
				</td>
			</tr>
			<script type="text/javascript" language="JavaScript">
				form1.chk_listino_b2c.onclick();
			</script>
		<% end if 
		select case listino_tipo
			case "BASE" %>
				<tr><th colspan="4">GENERAZIONE INIZIALE LISTINO:</th></tr>
				<tr>
					<td class="label" rowspan="2">copia prezzi da:</td>
					<td class="content" colspan="3">
						<input class="checkbox" type="radio" name="copia_da" <%= chk(request("copia_da")="") %> onclick="DisableIfChecked(this, document.form1.copia_da_listino);">
						prezzo base articolo
					</td>
				</tr>
				<tr>
					<td class="content">
						<input class="checkbox" <%= disable(listino_base = 0) %> type="radio" name="copia_da" id="copia_da_B" value="B" <%= chk(request("copia_da")="B") %> onclick="EnableIfChecked(this, document.form1.copia_da_listino);">
						listino base:
					</td>
					<td class="content" colsopan="3">
						<% sql = " SELECT listino_id, listino_codice FROM gtb_listini " + _
								 " WHERE listino_base=1 ORDER BY listino_codice"
						CALL dropDown(conn, sql, "listino_id", "listino_codice", "copia_da_listino", request("copia_da_listino"), true, "id=""copia_da_listino_B""", LINGUA_ITALIANO)%>
					</td>
				</tr>
			<% case "OFFERTE SPECIALI" %>
				<tr><th colspan="4">GENERAZIONE INIZIALE LISTINO:</th></tr>
				<tr>
					<td class="label" rowspan="2">copia prezzi da:</td>
					<td class="content" colspan="3">
						<input class="checkbox" type="radio" name="copia_da" <%= chk(request("copia_da")="") %> onclick="DisableIfChecked(this, document.form1.copia_da_listino);">
						prezzi listino base
					</td>
				</tr>
				<tr>
					<td class="content">
						<input class="checkbox" type="radio" name="copia_da" id="copia_da_B" value="B" <%= chk(request("copia_da")="B") %> onclick="EnableIfChecked(this, document.form1.copia_da_listino);">
						listino offerte:
					</td>
					<td class="content" colspan="2">
						<% sql = " SELECT listino_id, listino_codice FROM gtb_listini " + _
								 " WHERE listino_offerte=1 ORDER BY listino_codice"
						CALL dropDown(conn, sql, "listino_id", "listino_codice", "copia_da_listino", request("copia_da_listino"), true, "id=""copia_da_listino_B""", LINGUA_ITALIANO)%>
					</td>
				</tr>
			<% case else
				'generazione listino clienti 
				%>
				<tr><th colspan="4">GENERAZIONE E COMPORTAMENTO LISTINO:</th></tr>
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
					
					function SetControlState(){
						var copia_da_L = document.getElementById("copia_da_L");
						EnableIfChecked(copia_da_L, document.form1.tfn_listino_default_var_sconto);
						EnableIfChecked(copia_da_L, document.form1.tfn_listino_default_var_euro);
						var copia_da_B = document.getElementById("copia_da_B");
						EnableIfChecked(copia_da_B, document.form1.copia_da_listino);
						var copia_da_D = document.getElementById("copia_da_D");
						EnableIfChecked(copia_da_D, document.form1.tfn_listino_ancestor_id);
					}
					
				</script>
				<tr>
					<td class="label" rowspan="5">modalit&agrave;:</td>
					
					<td class="content" colspan="2">
						<input class="checkbox" type="radio" name="copia_da" id="copia_da_L" <%= chk(request("copia_da")="") %> onclick="SetControlState();">
						 <strong>copia</strong> prezzi da listino base in vigore
					</td>
					<td class="note" rowspan="4" style="width:44%;">
						PREZZI INDIPENDENTI:<br>
						I prezzi verranno copiati interamente nel nuovo listino e saranno completamente indipendenti dal listino di origine.
					</td>
				</tr>
				<tr>
					<td class="content" rowspan="2" style="padding-left:30px; width:25%;">
						applica variazioni di default:
					</td>
					<td class="label_no_width">
						<input type="text" class="number" name="tfn_listino_default_var_sconto" id="tfn_listino_default_var_sconto" value="<%= FormatPrice(request("tfn_listino_default_var_sconto") , 2, false) %>" size="4" onchange="ApplicaVariazioniPrezzo(this)"> %
					</td>
				</tr>
				<tr>
					<td class="label_no_width">
						<input type="text" class="number" name="tfn_listino_default_var_euro" id="tfn_listino_default_var_euro" value="<%= FormatPrice(request("tfn_listino_default_var_euro") , 2, false) %>" size="4" onchange="ApplicaVariazioniPrezzo(this)"> &euro;
					</td>
				</tr>
				<tr>
					<td class="content" colspan="2">
						<input class="checkbox" type="radio" name="copia_da" id="copia_da_B" value="B" <%= chk(request("copia_da")="B") %> onclick="SetControlState();">
						<strong>copia</strong> prezzi da altro listino:<br>
						<span style="padding-left:22px;">
							<% sql = " SELECT listino_id, listino_codice FROM gtb_listini " + _
									 " WHERE listino_offerte=0 AND listino_base=0 ORDER BY listino_codice"
							CALL dropDown(conn, sql, "listino_id", "listino_codice", "copia_da_listino", request("copia_da_listino"), true, "id=""copia_da_listino_B""", LINGUA_ITALIANO)%>
						</span>
					</td>
				</tr>
				<tr>
					<td class="content" colspan="2">
						<input class="checkbox" type="radio" name="copia_da" id="copia_da_D" value="D" <%= chk(request("copia_da")="D") %> onclick="SetControlState();">
						<strong>deriva</strong> prezzi da altro listino:<br>
						<span style="padding-left:22px;">
							<% sql = " SELECT listino_id, listino_codice FROM gtb_listini " + _
									 " WHERE listino_offerte=0 AND IsNull(listino_ancestor_id,0)=0 ORDER BY listino_codice" 'AND listino_base=0
							CALL dropDown(conn, sql, "listino_id", "listino_codice", "tfn_listino_ancestor_id", request("tfn_listino_ancestor_id"), true, "id=""tfn_listino_ancestor_id_B""", LINGUA_ITALIANO)%>
						</span>
					</td>
					<td class="note">
						PREZZI DERIVATI:<br>
						I prezzi del nuovo listino verranno copiati dal listino di origine, ma rimarranno ad esso collegati permettendo
						l'aggiornamento in caso di variazione del listino di origine finch&egrave; non interverr&agrave; una variazione nel prezzo del listino generato.
					</td>
				</tr>
				<script type="text/javascript" language="JavaScript">
					SetControlState();
				</script>
		<% end select %>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_listino_note"><%= request("tft_listino_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA &gt;&gt;">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
conn.close
set conn = nothing
%>
