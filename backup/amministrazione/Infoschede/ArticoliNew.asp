<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ArticoliSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione modelli - nuovo "
if request("STANDALONE")<>"true" then
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = "Articoli.asp"
else
	dicitura.puls_new = ""
	dicitura.link_new = ""
end if
dicitura.scrivi_con_sottosez()

dim conn, rs, rsv, sql, i, rs_spe
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")
set rs_spe = Server.CreateObject("ADODB.Recordset")
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_art_applicativo_id" value="<%= INFOSCHEDE %>">
	<input type="hidden" name="tfn_art_spedizione_id" value="<%=GetValueList(conn,NULL,"SELECT TOP 1 spa_id FROM gtb_spese_spedizione_articolo")%>">
	<input type="hidden" name="tfn_art_iva_id" value="<%=GetValueList(conn,NULL,"SELECT TOP 1 iva_id FROM gtb_iva")%>">
	<input type="hidden" name="tfn_art_varianti" value="0">
	<input type="hidden" name="tfn_art_giacenza_min" value="1">
	<input type="hidden" name="tfn_art_lotto_riordino" value="1">
	<input type="hidden" name="tfn_art_qta_min_ord" value="1">
	<input type="hidden" name="tfn_art_ordine" value="0">
	<input type="hidden" name="tfn_art_qta_max_ord" value="0">
	<input type="hidden" name="tfn_art_se_accessorio" value="0">
	<input type="hidden" name="tfn_art_ha_accessori" value="0">
	<input type="hidden" name="tfn_art_in_bundle" value="0">
	<input type="hidden" name="tfn_art_in_confezione" value="0">
	<input type="hidden" name="tfn_art_se_bundle" value="0">
	<input type="hidden" name="tfn_art_se_confezione" value="0">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Inserimento nuovo modello</caption>
		<tr><th colspan="7">DATI PRINCIPALI</th></tr>
		<% sql = "SELECT * FROM gtb_lista_codici WHERE lstCod_sistema=1 ORDER BY lstCod_nome" 
		rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
		<tr>
			<td class="label" style="width:16%;"<% if rsv.recordcount>0 then %>rowspan="<%= 1+rsv.recordcount %>"<% end if %>>codici:</td>
			<td class="label" style="width:8%;">interno:</td>
			<td class="content" style="width:25%;">
				<input type="text" class="text" name="tft_art_cod_int" value="<%= request("tft_art_cod_int") %>" maxlength="50" size="15">
				(*)
			</td>
			<td class="label" style="width:8%;">alternativo:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_art_cod_alt" value="<%= request("tft_art_cod_alt") %>" maxlength="50" size="15">
			</td>
		</tr>
		<% while not rsv.eof %>
			<tr>
				<td class="label_no_width" colspan="2">
					<%= rsv("lstCod_nome") %>
				</td>
				<td class="content" colspan="4">
					<input type="text" class="text" name="codice_articolo_<%= rsv("lstCod_id") %>" value="<%= request("codice_articolo_" & rsv("lstCod_id")) %>" maxlength="50" size="23">
				</td>
			</tr>
			<% rsv.movenext
		wend 
		rsv.close
		for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
				<td class="content" colspan="6">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_art_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_art_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">categoria:</td>
			<td class="content" colspan="6">
				<% if cString(request("CATEGORIA"))="ricambio" then %>
					<%CALL dropDown(conn, cat_ricambi.QueryElenco(true, ""), "tip_id", "NAME", "tfn_art_tipologia_id", request("tfn_art_tipologia_id"), false, " onchange='form1.submit()'", LINGUA_ITALIANO)%>
				<% else %>
					<%CALL dropDown(conn, cat_modelli.QueryElenco(true, ""), "tip_id", "NAME", "tfn_art_tipologia_id", request("tfn_art_tipologia_id"), false, " onchange='form1.submit()'", LINGUA_ITALIANO)%>
				<% end if %>
				(*)
			</td>
		</tr>
		<% sql = "SELECT COUNT(*) FROM gtb_tipologie_raggruppamenti"
		if cIntero(getValueList(conn, rsv, sql))>0 then %>
			<tr>
				<td class="label">&nbsp;</td>
				<td class="label" colspan="2">raggruppamento di pubblicazione:</td>
				<td class="content" colspan="5">
					<% if cInteger(request("tfn_art_tipologia_id"))>0 then
						sql = " SELECT * FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=" & cIntero(request("tfn_art_tipologia_id"))
						rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						if rsv.eof then %>
							<span class="note">Nessun raggruppamento disponibile per questa categoria di prodotti</span>
							<input type="hidden" name="nfn_art_raggruppamento_id" value="NULL">
						<% else
							CALL DropDownRecordset(rsv, "rag_id", "rag_nome_it", "nfn_art_raggruppamento_id", request("nfn_art_raggruppamento_id"), false, "", LINGUA_ITALIANO)
						end if
						rsv.close
					end if %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">marchio / produttore:</td>
			<td class="content" colspan="6">
				<%CALL dropDown(conn, "SELECT mar_id, mar_nome_it FROM gtb_marche ORDER BY mar_nome_it", _
							    "mar_id", "mar_nome_it", "tfn_art_marca_id", request("tfn_art_marca_id"), true, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<% if cString(request("CATEGORIA"))="ricambio" then %>
			<tr>
				<td class="label" colspan="2">prezzo base:</td>
				<td class="content" colspan="5">
					<input type="text" class="number" name="tfn_art_prezzo_base" value="<%= FormatPrice(cReal(request("tfn_art_prezzo_base")), 2, false) %>" size="7"> &euro;
					<span style="padding-left:5px;">(*)</span>
				</td>
			</tr>
		<% else %>
			<input type="hidden" name="tfn_art_prezzo_base" value="0">
		<% end if %>
		<tr><th colspan="7">DATI PER LA GESTIONE</th></tr>
		<tr>
			<td class="label" colspan="2">non a catalogo:</td>
			<td class="content" colspan="5"><input type="checkbox" class="checkbox" name="chk_art_disabilitato" <%= chk(request("chk_art_disabilitato")<>"") %>></td>
		</tr>
		<tr><th colspan="7">DESCRIZIONE </th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="7">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="4" name="tft_art_descr_<%= Application("LINGUE")(i) %>"><%= request("tft_art_descr_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
	</table>
	
	<% sql = " SELECT ct_id FROM gtb_carattech INNER JOIN gtb_tip_ctech ON gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id " & _
			 " WHERE rct_tipologia_id = " & cIntero(request("tfn_art_tipologia_id"))
		if cIntero(GetValueList(conn, NULL, sql)) > 0 then %>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
				<tr><th colspan="7">CARATTERISTICHE TECNICHE</th></tr>
				<% if cInteger(request("tfn_art_tipologia_id"))>0 then 					
					sql = " SELECT *" + _
						  " FROM gtb_carattech"& _
						  " INNER JOIN gtb_tip_ctech ON (gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id AND rct_tipologia_id=" & CInteger(request("tfn_art_tipologia_id")) & ")" + _
						  " LEFT JOIN grel_art_ctech ON (gtb_carattech.ct_id = grel_art_ctech.rel_ctech_id AND grel_art_ctech.rel_art_id=" & CInteger(request("ID")) & ")"& _
						  " LEFT JOIN gtb_carattech_raggruppamenti ON gtb_carattech.ct_raggruppamento_id = gtb_carattech_raggruppamenti.ctr_id " & _
						  " ORDER BY ctr_ordine, ctr_id, rct_ordine"
					CALL DesForm  (conn, sql, "gtb_carattech", "ct_id", "ct_nome_it", "ct_tipo", "ct_unita_it", "", "rel_ctech_", "rel_ctech_", "ctr_titolo_it", cIntero(request("ID")) = 0, 7)
					%>
				<% else %>
					<tr><td class="label" colspan="7">Per descrivere le caratteristiche tecniche dell'articolo selezionare prima la sua categoria.</td></tr>
				<% end if %>
			</table>
	<%	end if %>
	
	<% 	CALL oArticoliFoto.Elenco(request("ID"), "FOTO") %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="7">NOTE INTERNE</th></tr>
		<tr>
			<td class="content" colspan="7">
				<textarea style="width:100%;" rows="3" name="tft_art_note"><%= request("tft_art_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="7">
				(*) Campi obbligatori.
				<input <%= Disable(cInteger(request("tfn_art_tipologia_id"))=0) %> type="submit" class="button" name="salva" value="SALVA &gt;&gt;">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<% if cString(request("STANDALONE"))<>"" then %>
	<script language="JavaScript" type="text/javascript">
	<!--
		FitWindowSize(this);
	//-->
	</script>
<% end if %>

<%
set rs = nothing
set rsv = nothing
conn.Close
set conn = nothing
%>


