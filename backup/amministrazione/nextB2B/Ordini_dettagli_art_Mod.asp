<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("Ordini_dettagliSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica dettaglio ordine"
testata_show_back = false %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>
<SCRIPT LANGUAGE="javascript" src="Tools_B2B.js" type="text/javascript"></SCRIPT>

<% 
dim conn, rs, rsc, sql, giacenza
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")

sql = " SELECT listino_id, listino_base_attuale FROM gtb_listini " + _
	  " INNER JOIN gtb_rivenditori ON gtb_listini.listino_id = gtb_rivenditori.riv_listino_id " + _
	  " INNER JOIN gtb_ordini ON gtb_rivenditori.riv_id = gtb_ordini.ord_riv_id " + _
	  " INNER JOIN gtb_dettagli_ord ON gtb_ordini.ord_id = gtb_dettagli_ord.det_ord_id " + _
	  " WHERE gtb_dettagli_ord.det_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

sql = " gia_qta, gia_impegnato, rel_qta_min_ord, rel_lotto_riordino, rel_cod_int, rel_cod_alt, rel_cod_pro, det_cod_promozione, det_ord_id, det_art_var_id, " + _
	  " valu_cambio, valu_simbolo, listino_offerte, prz_promozione, art_id, art_ha_accessori, art_se_accessorio, det_qta, rel_id, " + _
	  " prz_prezzo, det_prezzo_unitario, IDElencoIndirizzi, ORD_ID, det_ind_id, art_nome_it, art_raggruppamento_id, art_tipologia_id, " + _
	  " art_se_bundle, art_se_confezione, art_varianti, mar_nome_it, art_disabilitato, art_in_bundle, art_in_confezione, art_se_accessorio, art_NoVenSingola, " + _
	  " iva_valore, prz_scontoQ_id, det_note " 
sql = " SELECT " + sql + _
	  " FROM gv_listino_vendita INNER JOIN gtb_dettagli_ord ON gv_listino_vendita.rel_id = gtb_dettagli_ord.det_art_var_id " + _
	  " INNER JOIN gtb_ordini ON gtb_dettagli_ord.det_ord_id = gtb_ordini.ord_id " + _
 	  " INNER JOIN grel_giacenze ON gv_listino_vendita.rel_id = grel_giacenze.gia_art_var_id AND grel_giacenze.gia_magazzino_id = gtb_ordini.ord_magazzino_id " + _
	  " INNER JOIN gv_rivenditori ON gtb_ordini.ord_riv_id = gv_rivenditori.riv_id " + _
	  " INNER JOIN gtb_marche ON gv_listino_vendita.art_marca_id = gtb_marche.mar_id " + _
	  " WHERE gtb_dettagli_ord.det_id=" & cIntero(request("ID")) & " AND " + RivenditoreListinoCondition(rs("listino_id"), rs("listino_base_attuale"))
rs.close

rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

giacenza = rs("gia_qta") - rs("gia_impegnato")
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1" onsubmit="return ControllaQta(form1.tfn_det_qta.value, <%= CInteger(rs("gia_qta")) %>, <%= CInteger(rs("rel_qta_min_ord")) %>, <%= CInteger(rs("rel_lotto_riordino")) %>, true)">
		<input type="hidden" name="tft_det_cod_promozione" value="<%= rs("det_cod_promozione") %>">
		<input type="hidden" name="tfn_det_ord_id" value="<%= rs("det_ord_id") %>">
		<input type="hidden" name="tfn_det_art_var_id" value="<%= rs("det_art_var_id") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica dettaglio d'ordine</caption>
			<tr><th colspan="7">DATI ARTICOLO</th></tr>
			<% CALL ArticoloScheda (conn, rs, rsc) %>
			<tr>
				<td class="label">prezzo:</td>
				<%CALL ScontiQ(conn, rs, rsc, 2, rs("valu_cambio"), rs("valu_simbolo"))%>
				<td class="content" colspan="4">
					<% if rs("listino_offerte") then %>
						<span class="Icona Offerte" title="prodotto in offerta speciale">&nbsp;</span>
						&nbsp;in offerta
					<% elseif rs("prz_promozione") then %>
						<span class="Icona Promozioni" title="prodotto in promozione">&nbsp;</span>
						&nbsp;in promozione
					<% else %>
						&nbsp;
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="label">giacenza attuale:</td>
				<% if giacenza <= 0  then %>
					<td class="content alert" colspan="6">n.d.</td>
				<% else %>
					<td class="content ok" colspan="6"><%= giacenza %></td>
				<% end if%>
			</tr>
			<% if cInteger(rs("rel_qta_min_ord")) > 1 then %>
				<tr>
					<td class="label">min. ordinabile:</td>
					<td class="content" colspan="6"><%= rs("rel_qta_min_ord") %></td>
				</tr>
			<% end if 
			if cInteger(rs("rel_lotto_riordino"))>0 then%>
				<tr>
					<td class="label">lotto di riordino:</td>
					<td class="content" colspan="6"><%= rs("rel_lotto_riordino") %></td>
				</tr>
			<% end if
			CALL ListaCollegamentiArticolo(conn, rsc, rs("art_id"), rs("art_ha_accessori"), rs("art_se_accessorio")) %>
			<input type="hidden" name="qta_old" value="<%= rs("det_qta") %>">
			<%dim quantita, prezzo_listino_base, prezzo_listino_cliente, sconto_listino_base, sconto_listino_cliente, prezzo_finale
			
			'imposta dati per form impostazione e variazione prezzi e sconti
			prezzo_listino_base = GetPrezzoListinoBase(conn, rsc, rs("rel_id"))
			prezzo_listino_cliente = rs("prz_prezzo")
			if Request.ServerVariables("REQUEST_METHOD")="POST" then
				'recupera dati dal form
				quantita = cInteger(request("tfn_det_qta"))
				prezzo_finale = cReal(request("tfn_det_prezzo_unitario"))
			else
				'recupera dati da base
				quantita = rs("det_qta")
				prezzo_finale = cReal(rs("det_prezzo_unitario"))
			end if
			
			sconto_listino_base = GetVarPercent(prezzo_listino_base, prezzo_finale)
			sconto_listino_cliente = GetVarPercent(prezzo_listino_cliente, prezzo_finale)
			
			%>
			<!--#INCLUDE FILE="Ordini_Dettagli_Include.asp" -->
			<tr>
				<td class="label" nowrap>indirizzo di destinazione</td>
				<td class="content" colspan="6">
					<% 	sql = "SELECT IDElencoIndirizzi, " + _
							  "(NomeOrganizzazioneElencoIndirizzi " + SQL_Concat(conn) + "' - '" + SQL_Concat(conn) + " IndirizzoElencoIndirizzi" + _
							  SQL_Concat(conn) + "' - '" + SQL_Concat(conn) + " CittaElencoIndirizzi) AS NOME " + _
							  "FROM tb_indirizzario WHERE IsNull(isSocieta,0)=1 AND (IDElencoIndirizzi=" & rs("IDElencoIndirizzi") & " OR cntRel="& rs("IDElencoIndirizzi") & ") " & _
							  " AND (IDElencoindirizzi NOT IN (SELECT det_ind_id FROM gtb_dettagli_ord WHERE det_ord_id=" & rs("ORD_ID") & " AND det_art_var_id=" & rs("rel_id") & ") " & _
							  " OR IdElencoIndirizzi=" & rs("det_ind_id") & ") ORDER BY cntRel, CittaElencoIndirizzi"
						'seleziono di default l'ultimo indirizzo immesso
						CALL dropDown(conn, sql, "IDElencoIndirizzi", "NOME", "tfn_det_ind_id", rs("det_ind_id"), true, "", LINGUA_ITALIANO) %>
					(*)
				</td>
			</tr>
			<tr>
				<td class="label">note:</td>
				<td class="content" colspan="6">
					<textarea name="tft_det_note" rows="2" style="width:100%;""><%=rs("det_note")%></textarea>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="7">
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
set rsc = nothing
set conn = nothing %>




<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>