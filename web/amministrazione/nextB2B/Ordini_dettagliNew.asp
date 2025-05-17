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
sezione_testata = "inserimento nuovo dettaglio ordine"
testata_show_back = true %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>
<SCRIPT LANGUAGE="javascript" src="Tools_B2B.js" type="text/javascript"></SCRIPT>

<% 
dim conn, rs, rsc, rso, sql, giacenza
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")
set rso = Server.CreateObject("ADODB.Recordset")

sql = " SELECT * FROM gtb_ordini INNER JOIN gv_rivenditori ON gtb_ordini.ord_riv_id = gv_rivenditori.riv_id " + _
	  " INNER JOIN gtb_listini ON gv_rivenditori.riv_listino_id = gtb_listini.listino_id " + _
	  " LEFT JOIN gtb_lista_codici ON gv_rivenditori.riv_LstCod_id = gtb_lista_codici.lstCod_id " + _
	  " WHERE ord_id=" & cIntero(request("ORD_ID"))
rso.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

'elenca articoli per l'ordine
sql = " SELECT * FROM gv_listini INNER JOIN grel_giacenze ON gv_listini.rel_id = grel_giacenze.gia_art_var_id " + _
	  " INNER JOIN gtb_marche ON gv_listini.art_marca_id = gtb_marche.mar_id " + _
	  " INNER JOIN gtb_listini ON gv_listini.prz_listino_id = gtb_listini.listino_id "
if cInteger(rso("riv_LstCod_id"))>0 then
	'aggiunge codifica dei codici personalizzata per rivenditore
	sql = sql + " LEFT JOIN gtb_codici ON (gv_listini.rel_id = gtb_codici.cod_variante_id AND cod_lista_id=" & rso("riv_LstCod_id") &") "
end if
sql = sql + " WHERE prz_id=" & cIntero(request("PRZ_ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

giacenza = rs("gia_qta") - rs("gia_impegnato")
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1" onsubmit="return ControllaQta(form1.tfn_det_qta.value, <%= CInteger(giacenza) %>, <%= CInteger(rs("rel_qta_min_ord")) %>, <%= CInteger(rs("rel_lotto_riordino")) %>, true)">
		<input type="hidden" name="tfn_det_ord_id" value="<%= request("ORD_ID") %>">
		<input type="hidden" name="tfn_det_art_var_id" value="<%= rs("rel_id") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Selezione nuovo dettaglio</caption>
			<tr><th colspan="7">DATI ARTICOLO</th></tr>
			<% 	CALL ArticoloScheda (conn, rs, rsc) %>
			<tr>
				<td class="label" style="width:20%;">prezzo:</td>
				<%CALL ScontiQ(conn, rs, rsc, 2, rso("valu_cambio"), rso("valu_simbolo"))%>
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
					<td class="content alert" colspan="6">0</td>
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
			if rs("art_se_accessorio") then
				sql = " SELECT * FROM gtb_articoli " + _ 
					  " INNER JOIN grel_art_acc ON gtb_articoli.art_id = grel_Art_acc.aa_art_id " + _
					  " WHERE aa_acc_id=" & rs("art_id")
				rsc.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
				if not rsc.eof then%>
					<tr>
						<td class="label">accessorio di:</td>
						<td colspan="6">
							<% if rsc.recordcount>2 then %> 
								<span class="overflow">
							<% end if  %>
							<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
								<tr>
									<th class="l2_center" width="14%">codice</th>
									<th class="L2">nome</th>
									<th class="L2">tipo</th>
								</tr>
								<% while not rsc.eof %>
									<tr>
										<td class="content"><%= rsc("art_cod_int")%></td>
										<td class="content">
											<%= rsc("art_nome_it")%>
										</td>
										<% if rsc("art_se_bundle") then %>
											<td class="content bundle">bundle di articoli</td>
										<% elseif rsc("art_se_confezione") then %>
											<td class="content confezione">confezione di articoli</td>
										<% elseif rsc("art_varianti") then %>
											<td class="content varianti">articolo con varianti</td>
										<% else %>
											<td class="content">articolo singolo</td>
										<% end if %>
									</tr>
									<%rsc.movenext
								wend %>
							</table>
							<% if rsc.recordcount>2 then %>
								</span>
							<% end if %>
						</td>
					</tr>
				<% end if
				rsc.close
			end if 
			
			dim quantita, prezzo_listino_base, prezzo_listino_cliente, sconto_listino_base, sconto_listino_cliente, prezzo_finale
			'imposta dati per form inserimento
			prezzo_listino_base = GetPrezzoListinoBase(conn, rsc, rs("rel_id"))
			prezzo_listino_cliente = rs("prz_prezzo")
			if Request.ServerVariables("REQUEST_METHOD")="POST" then
				'recupera dati dal form
				quantita = cInteger(request("tfn_det_qta"))
				prezzo_finale = cReal(request("tfn_det_prezzo_unitario"))
			else
				'recupera dati da base
				quantita = IIF(rs("rel_qta_min_ord")>0, rs("rel_qta_min_ord"), IIF(rs("rel_lotto_riordino")>0, rs("rel_lotto_riordino"), 1))
				prezzo_finale = GetPrezzoUnitario(conn, rsc, prezzo_listino_cliente, quantita, rs("prz_scontoQ_id"))
			end if
			
			sconto_listino_base = GetVarPercent(prezzo_listino_base, prezzo_finale)
			sconto_listino_cliente = GetVarPercent(prezzo_listino_cliente, prezzo_finale)
			%>
			<!--#INCLUDE FILE="Ordini_Dettagli_Include.asp" -->
			<tr>
				<td class="label" nowrap>indirizzo di destinazione</td>
				<td class="content" colspan="6">
					<% 	sql = "SELECT IDElencoIndirizzi, " + _
							  "((CASE ISNULL(IsSocieta, 0) WHEN 1 THEN NomeOrganizzazioneElencoIndirizzi ELSE CognomeElencoIndirizzi + ' ' + NomeElencoIndirizzi END) " + SQL_Concat(conn) + "' - '" + SQL_Concat(conn) + " IndirizzoElencoIndirizzi" + _
							  SQL_Concat(conn) + "' - '" + SQL_Concat(conn) + " CittaElencoIndirizzi) AS NOME " + _
							  "FROM tb_indirizzario WHERE (IDElencoIndirizzi=" & rso("IDElencoIndirizzi") & " OR cntRel="& rso("IDElencoIndirizzi") & ") " & _
							  " AND IDElencoindirizzi NOT IN (SELECT det_ind_id FROM gtb_dettagli_ord WHERE det_ord_id=" & cIntero(request("ORD_ID")) & " AND det_art_var_id=" & rs("rel_id") & ")" & _
							  " ORDER BY cntRel, CittaElencoIndirizzi"
						'seleziono di default l'ultimo indirizzo immesso
						CALL dropDown(conn, sql, "IDElencoIndirizzi", "NOME", "tfn_det_ind_id", _
										GetValueList(conn, rsc, "SELECT TOP 1 det_ind_id FROM gtb_dettagli_ord WHERE det_ord_id="& cIntero(request("ORD_ID")) &" AND det_art_var_id<>" & rs("rel_id") & " ORDER BY det_id DESC"), true, "", LINGUA_ITALIANO) %>
					(*)
				</td>
			</tr>
			<tr>
				<td class="label">note:</td>
				<td class="content" colspan="6">
					<textarea name="tft_det_note" rows="2" style="width:100%;""><%=request("tft_det_note")%></textarea>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="7">
					(*) Campi obbligatori.
					<input style="width:15%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:15%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
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