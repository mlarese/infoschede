<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<% 
dim conn, rs, rsp, sql, tipo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if isNumeric(request("rel_prezzo")) AND isNumeric(request("rel_sconto")) then
		dim prezzo
		conn.beginTrans
	
		sql = "SELECT * FROM grel_art_valori WHERE rel_id=" & cIntero(request("ID"))
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		prezzo = cReal(request("tfn_rel_prezzo"))
		rs("rel_prezzo") = prezzo
		rs("rel_prezzo_indipendente") = cInteger(request("tfn_rel_prezzo_indipendente"))>0
		if rs("rel_prezzo_indipendente") then
			rs("rel_var_euro") = 0
			rs("rel_var_sconto") = 0
		else
			rs("rel_var_euro") = cReal(request("tfn_rel_var_euro"))
			rs("rel_var_sconto") = cReal(request("tfn_rel_var_sconto"))
		end if
		rs.update
		rs.close
		
		if not GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then
			Select case cInteger(request("aggiornamento_listini"))
				case 0		'aggiornamento prezzi via sconti da prezzo variante (normale)
					CALL AggiornaPrezzoListini(conn, request("ID"))
				case 2 
					'azzera variazioni per listini base
					sql = " UPDATE gtb_prezzi SET prz_var_sconto=0, prz_var_euro=0 WHERE prz_variante_id = " & cIntero(request("ID")) & _
						  " AND prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE listino_base=1) "
					CALL conn.execute(sql, , adExecuteNoRecords)
					
					CALL AggiornaPrezzoListini(conn, request("ID"))
				case 1
					'aggiornamento sconti da prezzo base via prezzo variante
					sql = " SELECT * FROM gtb_prezzi WHERE prz_variante_id = " & cIntero(request("ID")) & _
						  " AND prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE listino_base=1) "
					rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
					while not rs.eof
						if cReal(rs("prz_var_euro"))<>0 then
							rs("prz_var_sconto") = 0
							rs("prz_var_euro") = rs("prz_prezzo") - prezzo
						else
							rs("prz_var_sconto") = GetVarPercent(prezzo, rs("prz_prezzo"))
							rs("prz_var_euro") = 0
						end if
						rs.update
						rs.movenext
					wend
					rs.close
					CALL AggiornaPrezzoListiniDaListinoBase(conn, request("ID"))
			end select
		end if
		conn.commitTrans
		set rsp = nothing%>
		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.close();
		</script>
	<%else
		Session("ERRORE") = "Prezzo immesso non valido."
	end if
end if


sql = "SELECT * FROM gv_articoli WHERE rel_id=" & cIntero(request("ID"))

'response.write sql
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<%'--------------------------------------------------------
sezione_testata = "modifica prezzo della variante"%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>
<SCRIPT LANGUAGE="javascript"  src="Tools_B2B.js" type="text/javascript"></SCRIPT>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<caption>
				Modifica prezzo base della variante
			</caption>
			<tr><th colspan="3">DATI BASE ARTICOLO</th></tr>
			<tr>
				<td class="label" colspan="2" rowspan="2" style="width:27%">articolo</td>
				<td class="content"><%= rs("art_nome_it") %></td>
			</tr>
			<tr>
				<% if rs("art_se_bundle") then %>
					<td class="content bundle">bundle</td>
				<% elseif rs("art_se_confezione") then %>
					<td class="content confezione">confezione</td>
				<% elseif rs("art_varianti") then %>
					<td class="content varianti">articolo con varianti</td>
				<% else %>
					<td class="content">articolo singolo</td>
				<% end if %>
			</tr>
			<tr>
				<td class="label" colspan="2">prezzo base dell'articolo</td>
				<td class="content">
					<input type="hidden" name="prezzo_articolo" value="<%= rs("art_prezzo_base") %>">
					<%= FormatPrice(rs("art_prezzo_base"), 2, true) %>&nbsp;&euro;
				</td>
			</tr>
		</table>
		<script language="JavaScript" type="text/javascript">
			function SetImputState(){
				var stato = document.getElementById("tipo_prezzo_indipendente")
				DisableControl(form1.tfn_rel_var_euro, stato.checked);
				DisableControl(form1.tfn_rel_var_sconto, stato.checked);
			}
			
			function ApplicaVariazioniPrezzo(tag){
				var altra_variazione;
				var prezzo_base = form1.prezzo_articolo;
				var prezzo_attuale = form1.tfn_rel_prezzo;
				if (tag.name == ('tfn_rel_var_euro')){
					//applica variazione in euro
					altra_variazione = form1.tfn_rel_var_sconto;
					
					//applica variazione in euro al prezzo
					CalcolaPrezzoEuro(prezzo_base, prezzo_attuale, tag)
				}
				else{
					//applica variazione in percentuale
					altra_variazione = form1.tfn_rel_var_euro;
					
					//applica variazione percentuale al prezzo
					CalcolaPrezzo(prezzo_base, prezzo_attuale, tag)
				}
				altra_variazione.value = '0,00';
				tag.value = FormatNumber(tag.value, 2);
				
			}
			
			function RicalcolaVariazioniPrezzo(prezzo_attuale){
				var value_euro, value_perc;
				var prezzo_base = form1.prezzo_articolo;
				var prezzo_var_euro = form1.tfn_rel_var_euro;
				var prezzo_var_perc = form1.tfn_rel_var_sconto;
				
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

		</script>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr><th colspan="3">DATI VARIANTE</th></tr>
			<tr>
				<td class="label" colspan="2">codice</td>
				<td class="content"><%= rs("rel_cod_int") %></td>
			</tr>
			<tr>
				<td class="label" colspan="2">variante</td>
				<td class="content"><% CALL TableValoriVarianti(conn, rsp, rs("rel_id"), "content") %></td>
			</tr>
			<tr>
				<td class="label" rowspan="5" style="width:10%;">prezzo:</td>
				<td class="label" rowspan="2" style="width:27%;">
					calcola prezzo variante:
				</td>
				<td class="content">
					<input type="radio" class="checkbox" name="tfn_rel_prezzo_indipendente" id="tipo_prezzo_dipendente" value="0" <%= chk(rs("rel_prezzo_indipendente") OR IsNull(rs("rel_prezzo_indipendente"))) %> onclick="SetImputState()">
					da prezzo articolo sulla base delle variazioni applicate
				</td>
			</tr>
			<tr>
				<td class="content">
					<input type="radio" class="checkbox" name="tfn_rel_prezzo_indipendente" id="tipo_prezzo_indipendente" value="1" <%= chk(rs("rel_prezzo_indipendente")) %> onclick="SetImputState()">
					indipendente dal prezzo articolo
				</td>
			</tr>
			<tr>
				<td class="label">
					prezzo della variante:
				</td>
				<td class="content"><input type="text" class="number" name="tfn_rel_prezzo" value="<%= FormatPrice(cReal(rs("rel_prezzo")) , 2, false) %>" size="7" onchange="RicalcolaVariazioniPrezzo(this)"> &euro;&nbsp;&nbsp;&nbsp;&nbsp;(*)</td>
			</tr>
			<tr>
				<td class="label" rowspan="2">
					variazioni da prezzo base articolo:
				</td>
				<td class="content" style="padding-left:19px"><input type="text" class="number" name="tfn_rel_var_sconto" value="<%= FormatPrice(cReal(rs("rel_var_sconto")) , 2, false) %>" size="4" onchange="ApplicaVariazioniPrezzo(this)"> %</td>
			</tr>
			<tr>
				<td class="content" style="padding-left:19px"><input type="text" class="number" name="tfn_rel_var_euro" value="<%= FormatPrice(cReal(rs("rel_var_euro")) , 2, false) %>" size="4" onchange="ApplicaVariazioniPrezzo(this)"> &euro;</td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<% if not GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then %>
				<tr><th colspan="3">AGGIORNAMENTO LISTINI BASE</th></tr>
				<tr>
					<td class="label" rowspan="3" style="width:10%;">tipo:</td>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_listini" value="0" <%= chk(cInteger(request("aggiornamento_listini"))=0) %>>
					</td>
					<td class="content">
						ricalcola i prezzi<br>
						<span class="note">Il calcolo verr&agrave; effettuato sulla base delle variazioni (in &euro; o %) applicate.</span>
					</td>
				</tr>
				<tr>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_listini" value="1" <%= chk(cInteger(request("aggiornamento_listini"))=1) %>>
					</td>
					<td class="content">
						lascia prezzi inalterati<br>
						<span class="note">il sistema ricalcoler&agrave; automaticamente le variazioni (in &euro; o %) applicate.</span>
					</td>
				</tr>
				<tr>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_listini" value="2" <%= chk(cInteger(request("aggiornamento_listini"))=2) %>>
					</td>
					<td class="content">
						sostituisci i prezzi ed azzera le variazioni (in &euro; o %) applicate.
					</td>
				</tr>
				<tr>
					<td class="note" colspan="3">ATTENZIONE: i listini delle offerte speciali e dei clienti verranno aggiornati solo nei prezzi: nessuna variazione verr&agrave; effettuata sullevariazioni (in &euro; o %) applicate.</td>
				</tr>
			<% else %>
				<tr>
					<td class="note" colspan="3">
						ATTENZIONE: la variazione di prezzo non verr&agrave; propagata nei listini perchè attiva la modalità Listini Indipendenti.
					</td>
				</tr>
			<% end if %>
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
	</form>
		</table>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set rsp = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
	SetImputState();
//-->
</script>