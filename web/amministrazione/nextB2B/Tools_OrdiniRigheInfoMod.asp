
<% 	
dim conn, rs, rsc, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
set rsc = server.CreateObject("ADODB.recordset")


if request("goto")<>"" then
	CALL GotoRecord(conn, rsc, session(name_session_sql), "dod_id", "OrdiniRigheInfoMod.asp")
end if


sql = "SELECT * FROM gtb_dettagli_ord_des WHERE dod_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati dell'informazione per riga d'ordine</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="caratteristica precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="caratteristica successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI DELLA INFORMAZIONE PER RIGA D'ORDINE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_dod_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("dod_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_dod_codice" value="<%= rs("dod_codice") %>" maxlength="255" size="26">
				<% response.write "(*)" %>
			</td>
		</tr>

		<tr>
			<td class="label">tipo di dato:</td>
			<td class="content">
				<% DesDropTipi "tfn_dod_tipo", "", rs("dod_tipo") %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">unit&agrave; di misura:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_dod_unita_<%= Application("LINGUE")(i) %>" value="<%= rs("dod_unita_"& Application("LINGUE")(i)) %>" maxlength="50" size="50">
				</td>
			</tr>
		<%next %>
		
        <tr><th colspan="2">IMPOSTAZIONI CALCOLO QUANTITA'</th></tr>
		<tr>
			<td class="label">abilita detrazione:</td>
			<td class="content">
				<input type="checkbox" <%= chk(rs("dod_qta_in_detrazione")) %> class="<%= IIF(rs("dod_qta_in_detrazione"), "checked", "checkbox") %>" name="chk_dod_qta_in_detrazione" onclick="set_state_abilita_detraz(this)" title="Abilita la detrazione">
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			function set_state_abilita_detraz(chk){
				EnableIfChecked(chk, form1.tfn_dod_percentuale_detrazione);
				if (chk.checked){
					form1.tfn_dod_percentuale_detrazione.title = "Percentuale di detrazione";
				}
				else{
					form1.tfn_dod_percentuale_detrazione.title = "Selezionare il flag che abilita l'inserimento della percentuale di detrazione";
				}
			}
		</script>
		<tr>
			<td class="label">percentuale di detrazione:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_dod_percentuale_detrazione" <%= IIF(rs("dod_qta_in_detrazione"), "", "disabled") %> value="<%= rs("dod_percentuale_detrazione") %>" size="2" title="<%= IIF(rs("dod_qta_in_detrazione"), "Percentuale di detrazione", "Selezionare prima il flag che abilita l'inserimento della percentuale di detrazione")%>">
				%
			</td>
		</tr>
		
		<tr><th colspan="2">LISTINO A CUI E' ASSOCIATA</th></tr>
		<tr>
			<td class="label">listino:</td>
			<td class="content">
				<%CALL dropDown(conn, "SELECT * FROM gtb_listini ", _
							"listino_id", "listino_codice", "tfn_dod_listino_id", rs("dod_listino_id") , false, " style=""width=250""", LINGUA_ITALIANO)%>
			</td>
		</tr>		
		
		<tr><th colspan="2">TIPOLOGIE DI RIGA A CUI &Egrave; ASSOCIATA</th></tr>
		<tr>
			<td colspan="2">
                <%dim value
                sql = "SELECT *, " + _
                      " (SELECT COUNT(*) FROM grel_dett_cart_des_value INNER JOIN gtb_dett_cart ON grel_dett_cart_des_value.rel_des_dett_cart_id = gtb_dett_cart.dett_id " + _
                      "  WHERE dett_tipo_id= gtb_dettagli_ord_tipo.dot_id AND rel_des_descrittore_id=" & cIntero(request("ID")) & ") AS N_RIGHE_CART, " + _
                      " (SELECT COUNT(*) FROM grel_dettagli_ord_des_value INNER JOIN gtb_dettagli_ord ON grel_dettagli_ord_des_value.rel_des_dett_ord_id = gtb_dettagli_ord.det_id " + _
                      "  WHERE det_tipo_id= gtb_dettagli_ord_tipo.dot_id AND rel_des_descrittore_id=" & cIntero(request("ID")) & ") AS N_RIGHE_ORDINI" + _
                      " FROM gtb_dettagli_ord_tipo LEFT JOIN grel_dettagli_ord_tipo_des ON gtb_dettagli_ord_tipo.dot_id = grel_dettagli_ord_tipo_des.rtd_tipo_id " +_
                      " AND grel_dettagli_ord_tipo_des.rtd_descrittore_id=" & cIntero(request("ID")) & _
                      " ORDER BY dot_nome_it"
					  rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% if rsc.eof then %>
						<tr>
							<td class="label_no_width">
								Nessuna tipologia di righe d'ordine definita
							</td>
						</tr>
					<% else %>
						<tr>
							<th class="l2_center" width="6%">associa</th>
							<th class="l2_center" width="7%">ordine</th>
							<th class="L2">tipologia di riga</th>
						</tr>
						<% while not rsc.eof %>
							<tr>
								<td class="content_center">
									<% if cInteger(rsc("N_RIGHE_CART"))>0 OR cinteger(rsc("N_RIGHE_ORDINI"))>0 then 
										value = true%>
										<input type="checkbox" checked class="checked" id="categorie_associate_<%= rsc("dot_id") %>" disabled onclick="set_state_<%= rsc("dot_id") %>(this)" title="Sono presenti valori nelle righe di questa tipologia.">
										<input type="hidden" name="categorie_associate" value=" <%= rsc("dot_id") %> ">
									<% else 
										value = not IsNull(rsc("rtd_id"))%>
										<input type="checkbox" name="categorie_associate" id="categorie_associate_<%= rsc("dot_id") %>" value=" <%= rsc("dot_id") %> " <%= chk(value) %> class="<%= IIF(value, "checked", "checkbox") %>" onclick="set_state_<%= rsc("dot_id") %>(this)">
									<% end if %>
								</td>
								<td class="content_center"><input <%= disable(not value) %> type="text" class="<%= IIF(not value, "text_disabled", "text") %>" size="2" name="rel_ordine_<%= rsc("dot_id") %>" value="<%= rsc("rtd_ordine") %>"></td>
								<td class="content"><%= rsc("dot_nome_it") %></td>
							</tr>
							<script language="JavaScript" type="text/javascript">
								function set_state_<%= rsc("dot_id") %>(chk){
									EnableIfChecked(chk, form1.rel_ordine_<%= rsc("dot_id") %>);
									if (chk.checked){
										form1.rel_ordine_<%= rsc("dot_id") %>.title = "Inserisci l'ordine di visualizzazione";
									}
									else{
										form1.rel_ordine_<%= rsc("dot_id") %>.title = "Selezionare il flag di associazione prima di inserire l'ordine di visualizzazione";
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
set rs = nothing
set rsc = nothing
conn.Close
set conn = nothing
%>