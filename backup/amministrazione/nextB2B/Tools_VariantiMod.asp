
<% 	
dim conn, rs, rsd, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session(name_session_sql), "var_ID", "VariantiMod.asp")
end if

sql = "SELECT * FROM gtb_varianti WHERE var_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica della variante</td>
					<td align="right" style="font-size: 1px;">
						<% if CBoolean(from_tour, false) then %>
							&nbsp;
						<% else %>
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="variante precedente" <%= ACTIVE_STATUS %>>
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="variante successiva" <%= ACTIVE_STATUS %>>
								SUCCESSIVA &gt;&gt;
							</a>
						<% end if %>
					</td>
				</tr>
			</table>
		</caption>
		<% if not CBoolean(from_tour, false) then %>
			<tr><th colspan="2">DATI DELLA VARIANTE</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
					<td class="content">
						<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="tft_var_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("var_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
						<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
					</td>
				</tr>
			<%next %>
			<tr>
				<td class="label">ordine:</td>
				<td class="content">
					<input type="text" class="text" name="tfn_var_ordine" value="<%= rs("var_ordine") %>" size="4" maxlength="2">
					(*)
				</td>
			</tr>
			<tr><th colspan="2">DESCRIZIONE</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="2">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="2" name="tft_var_descr_<%= Application("LINGUE")(i) %>"><%= rs("var_descr_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			<%next %>
		<% end if %>
		<tr><th colspan="2">DEFINIZIONE DEI VALORI</th></tr>
		<tr>
			<td colspan="2">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% sql = " SELECT *, (SELECT COUNT(*) FROM grel_art_vv WHERE rvv_val_id = gtb_valori.val_id) AS N_ART " & _
						  " FROM gtb_valori WHERE val_var_id=" & rs("var_id") & " ORDER BY val_ordine"
					rsd.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText%>
					<tr>
						<td class="label" colspan="5" style="width:80%;">
							<% if rsd.eof then %>
								nessun valore definito per la variante.
							<% else %>
								n&ordm; <%= rsd.recordcount %> valori per la variante.
							<% end if %>
						</td>
						<td colspan="2" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedWindow('Varianti_valoriNew.asp?EXTID=<%= request("ID") %>', 'ValoriVarianti', 510, 355)"
							   title="apre la finestra per l'inserimento di un nuovo valore" <%= ACTIVE_STATUS %>>
								NUOVO VALORE
							</a>
						</td>
					</tr>
					<% if not rsd.eof then %>
						<tr>
							<th class="L2">CODICE INTERNO</th>
							<th class="L2">CODICE PRODUTTORE</th>
							<th class="L2">VALORE</th>
							<th class="l2_center">ICONA</th>
							<th class="l2_center" width="10%">ORDINE</th>
							<th class="l2_center" width="16%" colspan="2">OPERAZIONI</th>
						</tr>
						<%while not rsd.eof %>
							<tr>
								<td class="content" width="20%"><%= rsd("val_cod_int") %></td>
								<td class="content" width="20%"><%= rsd("val_cod_pro") %></td>
								<td class="content"><%= rsd("val_nome_it") %></td>
								<td class="content_center">
									<% 	if CString(rsd("val_icona")) <> "" then %>
										<img height="15" src="http://<%= Application("IMAGE_SERVER") &"/"& Application("AZ_ID") &"/images/"& rsd("val_icona") %>" alt="logo" border="0">
									<% 	else %>
										&nbsp;
									<% 	end if %>
								</td>
								<td class="content_center"><%= rsd("val_ordine") %></td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedWindow('Varianti_valoriMod.asp?ID=<%= rsd("val_id") %>', 'ValoriVarianti', 510, 355)">
										MODIFICA
									</a>
								</td>
								<td class="content_center">
									<% if rsd("N_ART")>0 then %>
										<a class="button_L2_disabled" href="javascript:void(0);" title="Inpossibile cancellare il valore perch&egrave; associato ad almeno un articolo" <%= ACTIVE_STATUS %>>
											CANCELLA
										</a>
									<% else %>
										<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('VALORI','<%= rsd("val_id") %>');">
											CANCELLA
										</a>
									<% end if %>
								</td>
							</tr>
							<%rsd.MoveNext
						wend
					end if
					rsd.close%>
				</table>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
				<% if CBoolean(from_tour, false) then %>
					&nbsp;
				<% else %>
					(*) Campi obbligatori.
					<input type="submit" style="width:23%;" class="button" name="mod" value="SALVA & TORNA ALL'ELENCO">
					<input type="submit" class="button" name="salva" value="SALVA">
				<% end if %>
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
set rsd = nothing
conn.Close
set conn = nothing
%>