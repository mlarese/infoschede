
<% 	
dim conn, i, rsc, sql, value
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rsc = server.CreateObject("ADODB.recordset")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova informazione per riga d'ordine</caption>
		<tr><th colspan="2">DATI DELL'INFORMAZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_dod_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_dod_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_dod_codice" value="<%= request("tft_dod_codice") %>" maxlength="255" size="26">
				<% response.write "(*)" %>
			</td>
		</tr>
		
		<tr>
			<td class="label">tipo di dato:</td>
			<td class="content">
				<% DesDropTipi "tfn_dod_tipo", "", request.Form("tfn_dod_tipo") %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">unit&agrave; di misura:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_dod_unita_<%= Application("LINGUE")(i) %>" value="<%= request("tft_dod_unita_"& Application("LINGUE")(i)) %>" maxlength="50" size="50">
				</td>
			</tr>
		<%next %>
		
		<tr><th colspan="2">IMPOSTAZIONI CALCOLO QUANTITA'</th></tr>
		<tr>
			<td class="label">abilita detrazione:</td>
			<td class="content">
				<input type="checkbox" <%= chk(request("chk_dod_qta_in_detrazione")<>"") %> class="<%= IIF(request("chk_dod_qta_in_detrazione")<>"", "checked", "checkbox") %>" name="chk_dod_qta_in_detrazione" onclick="set_state_abilita_detraz(this)" title="Abilita la detrazione">
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
				<input type="text" class="text" name="tfn_dod_percentuale_detrazione" <%= IIF(request("chk_dod_qta_in_detrazione")<>"", "", "disabled") %> value="<%= request("tfn_dod_percentuale_detrazione") %>" size="2" title="<%= IIF(request("chk_dod_qta_in_detrazione")<>"", "Percentuale di detrazione", "Selezionare prima il flag che abilita l'inserimento della percentuale di detrazione")%>">
				%
			</td>
		</tr>		
		
		<tr><th colspan="2">LISTINO A CUI E' ASSOCIATA</th></tr>
		<tr>
			<td class="label">listino:</td>
			<td class="content">						
				<%CALL dropDown(conn, "SELECT * FROM gtb_listini ", _
							"listino_id", "listino_codice", "tfn_dod_listino_id", "" , false, " style=""width=250""", LINGUA_ITALIANO)%>

			</td>
		</tr>	
		
		<tr><th colspan="2">TIPOLOGIE DI RIGA A CUI &Egrave; ASSOCIATA</th></tr>
		<tr>
			<td colspan="2">
				<%sql = "SELECT *, (0) AS N_ARTICOLI FROM gtb_dettagli_ord_tipo ORDER BY dot_nome_it"
				rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% if rsc.eof then %>
						<tr>
							<td class="label_no_width">
								Nessuna tipologia di riga definita.
							</td>
						</tr>
					<% else %>
						<tr>
							<th class="l2_center" width="6%">associa</th>
							<th class="l2_center" width="7%">ordine</th>
							<th class="L2">categoria</th>
						</tr>
						<% while not rsc.eof %>
							<tr>
								<td class="content_center">
									<% if cInteger(rsc("N_ARTICOLI"))>0 then 
										value = true%>
										<input type="checkbox" checked class="checked" id="categorie_associate_<%= rsc("dot_id") %>" disabled onclick="set_state_<%= rsc("dot_id") %>(this)" title="Sono presenti valori nelle righe di questa tipologia.">
										<input type="hidden" name="categorie_associate" value=" <%= rsc("dot_id") %> ">
									<% else 
										value = instr(1, request("categorie_associate"), " " & rsc("dot_id") & " ", vbTextCompare)>0%>
										<input type="checkbox" name="categorie_associate" id="categorie_associate_<%= rsc("dot_id") %>" value=" <%= rsc("dot_id") %> " <%= chk(value) %> class="<%= IIF(value, "checked", "checkbox") %>" onclick="set_state_<%= rsc("dot_id") %>(this)">
									<% end if %>
								</td>
								<td class="content_center"><input <%= disable(not value) %> type="text" class="<%= IIF(not value, "text_disabled", "text") %>" size="2" name="rel_ordine_<%= rsc("dot_id") %>" value="<%= request("rel_ordine_" & rsc("dot_id")) %>"></td>
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
set rsc = nothing
conn.Close
set conn = nothing
%>