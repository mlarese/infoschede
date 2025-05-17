
<% 	
dim conn, rs, rsc, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session(name_session_sql), "dot_ID", "OrdiniRigheTipologieMod.asp")
end if

sql = "SELECT * FROM gtb_dettagli_ord_tipo WHERE dot_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova tipologia</caption>
		<tr><th colspan="2">DATI DELLA TIPOLOGIA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
		<% 	end if %>
			<td class="content">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_dot_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("dot_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next %>
		<tr>
			<td class="label" >codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_dot_codice" value="<%= rs("dot_codice") %>" maxlength="250" size="26">
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">descrizione:</td>
		<% 	end if %>
			<td class="content">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_dot_descrizione_<%= Application("LINGUE")(i) %>" value="<%= rs("dot_descrizione_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
			</td>
		</tr>
		<%next %>
        
		<tr><th colspan="2">INFORMAZIONI PER RIGA D'ORDINE</th></tr>
    	<tr>
    	    <td colspan="2">
    				<%	dim value, disabled
                        
    					sql = " SELECT *, " + _
                              " (SELECT COUNT(*) FROM grel_dettagli_ord_des_value INNER JOIN gtb_dettagli_ord ON grel_dettagli_ord_des_value.rel_des_dett_ord_id = gtb_dettagli_ord.det_id " + _
                              "  AND det_tipo_id = " & cIntero(request("ID")) & " WHERE grel_dettagli_ord_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id) AS N_RIGHE_ORDINI, " + _
                              " (SELECT COUNT(*) FROM grel_dett_cart_des_value INNER JOIN gtb_dett_cart ON grel_dett_cart_des_value.rel_des_dett_cart_id = gtb_dett_cart.dett_id " + _
                              "  AND dett_tipo_id = " & cIntero(request("ID")) & " WHERE grel_dett_cart_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id) AS N_RIGHE_CART " + _
                              " FROM gtb_dettagli_ord_des LEFT JOIN grel_dettagli_ord_tipo_des ON gtb_dettagli_ord_des.dod_id = grel_dettagli_ord_tipo_des.rtd_descrittore_id AND rtd_tipo_id=" & cIntero(request("ID")) & _
                              " ORDER BY dod_nome_it "
    					rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
    					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
    						<% if rsc.eof then %>
    							<tr>
    								<td class="label_no_width">
    									Nessuna informazione inserita
    								</td>
    							</tr>
    						<% else %>
    							<tr>
    								<th class="l2_center" width="6%">associa</th>
    								<th class="l2_center" width="7%">ordine</th>
    								<th class="L2">informazione</th>
    								<th class="L2" width="15%">tipo di dato</th>
    							</tr>
    							<% while not rsc.eof %>
    								<tr>
    									<td class="content_center">
    										<% 	disabled = ( cInteger(rsc("N_RIGHE_ORDINI"))>0 OR cInteger(rsc("N_RIGHE_CART"))>0 )
    											if disabled then
    												value = true%>
    											<input type="checkbox" checked class="checked" id="caratteristiche_associate_<%= rsc("dod_id") %>" disabled onclick="set_state_<%= rsc("dod_id") %>(this)" title="<%= IIF(cInteger(rsc("N_RIGHE_ORDINI"))>0 OR cInteger(rsc("N_RIGHE_CART"))>0, "Sono presenti valori per questa caratteristica negli articoli della categoria.", "Il descrittore &egrave; gestito da un applicativo esterno.") %>">
    											<input type="hidden" name="caratteristiche_associate" value=" <%= rsc("dod_id") %> ">
    										<% 	else
    												value = not IsNull(rsc("rtd_id"))%>
    											<input type="checkbox" name="caratteristiche_associate" id="caratteristiche_associate_<%= rsc("dod_id") %>" value=" <%= rsc("dod_id") %> " <%= chk(value) %> class="<%= IIF(value, "checked", "checkbox") %>" onclick="set_state_<%= rsc("dod_id") %>(this)">
    										<% 	end if %>
    									</td>
    									<td class="content_center"><input <%= disable(not value OR disabled) %> type="text" class="<%= IIF(not value OR disabled, "text_disabled", "text") %>" size="2" name="rel_ordine_<%= rsc("dod_id") %>" value="<%= rsc("rtd_ordine") %>"></td>
    									<td class="content" title="<%= rsc("dod_id") %>">
											<%= rsc("dod_nome_it") %>
											<span class="note">( <%= rsc("dod_codice") %> )</span>
										</td>
    									<td class="content"><%= DesVisTipo(rsc("dod_tipo")) %></td>
    								</tr>
    								<script language="JavaScript" type="text/javascript">
    									function set_state_<%= rsc("dod_id") %>(chk){
    										EnableIfChecked(chk, form1.rel_ordine_<%= rsc("dod_id") %>);
    										if (chk.checked){
    											form1.rel_ordine_<%= rsc("dod_id") %>.title = "Inserisci l'ordine di visualizzazione nella scheda";
    										}
    										else{
    											form1.rel_ordine_<%= rsc("dod_id") %>.title = "Selezionare il flag di associazione prima di inserire l'ordine di visualizzazione nella scheda";
    										}
    									}
    								</script>
    								<% rsc.movenext
    							wend %>
    						<% end if %>
    					</table>
    					<% rsc.close %>
    				</td>
    			</tr>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" style="width:23%;" class="button" name="mod" value="SALVA & TORNA ALL'ELENCO">
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
rs.close
set rs = nothing
set rsc = nothing
conn.Close
set conn = nothing
%>