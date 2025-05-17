<% 
'.................................................................................................................
'		procedura che controlla i permessi dell'utente e scrive e genera la parte del form
'		conn				connessione al database aperta
'		rs					oggetto recordset chiuso
'		amministrazione		indica se le applicazioni sono dell'amministrazione o dell'area riservata
'		utente				id dell'utente da gestire (0 se nuovo utente)
'		permessi_per_riga	numero di scelte permesso per riga
'.................................................................................................................
sub write_permessi(conn, rs, amministrazione, utente, permessi_per_riga)
	dim sql, rsP, checked, i, est, estAbilita, estNome, conta, max_num_permessi, lim_inf, lim_sup, num_col, fine
	set rsP = Server.CreateObject("ADODB.Recordset")
	max_num_permessi = 9
	
	sql = " SELECT * FROM tb_siti"
	if amministrazione then
		sql = sql & " WHERE " & SQL_IsTrue(conn, "sito_Amministrazione")
	else
		sql = sql & " WHERE not" & SQL_IsTrue(conn, "sito_Amministrazione")
	end if
	sql = sql & " ORDER BY sito_prmEsterni_admin, Sito_nome"
	
	if cIntero(permessi_per_riga) = 0 OR cIntero(permessi_per_riga) > max_num_permessi then
		permessi_per_riga = max_num_permessi
	end if
	
	
	est = ""
	estAbilita = (CInt(GetValueList(conn, rs, "SELECT COUNT(*) FROM tb_siti WHERE NOT "& SQL_IsNull(conn, "sito_prmEsterni_admin") &" AND sito_prmEsterni_admin <> ''")) > 0)
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText %>
	<script type="text/javascript">
		function apriGestione(siti, pagina, id) {
			for(i = 0; i < siti.length; i++)
				if (siti[i].defaultChecked != siti[i].checked) {
					alert("Non e' possibile accedere ai permessi aggiuntivi dopo una modifica. Salvare prima i dati.")
					return
				}
			
			OpenPositionedScrollWindow(pagina +"?ID="+ id, "permessi", 0, 0, 750, 300, true)
		}
	</script>
	<table cellpadding="0" cellspacing="1" width="100%">
		<% while not rs.eof %>
			<%
			fine = false
			lim_inf = 1
			lim_sup = permessi_per_riga
			for conta=0 to cIntero(max_num_permessi/permessi_per_riga) %>
				<tr>
					<td nowrap class="content" width="20%" title="<%= rs("id_sito") %>">
						<% if conta = 0 then %>
							<%= Replace(rs("sito_nome"), "[", "<br>[") %>
						<% else %>
							&nbsp;
						<% end if %>
					</td>
					<% num_col = 1 %>
					<%for i=lim_inf to lim_sup
						if CString(rs("sito_p"& i)) <> "" then%>
							<td class="content" nowrap
					<% 		if i < max_num_permessi then
								if CString(rs("sito_p"& i+1)) = "" then 
									fine = true %>
									colspan="<%= (permessi_per_riga - num_col)+1 %>"
					<%			end if
							else %>
								colspan="<%= (permessi_per_riga - num_col)+1 %>"
					<%		end if %>
							>
								<% if (rs("id_sito") <> cInteger(Session("ID_SITO"))) OR i>1 OR _	
									  Session("PASS_ADMIN")<>"" then 		'salta il permesso di amministratore dell'applicazione
									checked = "class=""checkbox"" "
									if utente>0 then
										if amministrazione then
											sql = "SELECT * FROM rel_admin_sito WHERE admin_id=" & utente &_
												  " AND sito_id=" & rs("id_sito") & " AND rel_as_permesso=" & i
										else
											sql = "SELECT * FROM rel_utenti_sito WHERE rel_ut_id=" & utente &_
												  " AND rel_sito_id=" & rs("id_sito") & " AND rel_permesso=" & i
										end if
											 
										 rsp.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
										 if not rsp.eof then
											checked = " checked class=""checked"" "
										 end if
										 rsp.close
									end if%>
									<input type="checkbox" name="chk_perm" id="chk_<%= rs("id_sito") %>" value="<%=rs("id_sito")%>,<%=i%>" <%= checked %>>
									&nbsp;<%=rs("sito_p" & i)%>
								<%end if %>
							</td>
						<%else
							exit for
						end if
						num_col = num_col + 1
					next
					
					if estAbilita AND request("ID") <> "" then
						if CString(rs("sito_prmEsterni_admin")) = "" then %>
							<td class="content_center">&nbsp;</td>
						<% elseif CString(rs("sito_prmEsterni_admin")) <> est then
							est = CString(rs("sito_prmEsterni_admin"))
							sql = " SELECT id_sito FROM tb_siti"& _
								  " WHERE sito_prmEsterni_admin = '"& rs("sito_prmEsterni_admin") &"'"
							estNome = Replace(GetValueList(conn, NULL, sql), " ", "") %>
							<td class="content_center visibile" rowspan="<%= Count(estNome, ",")+1 %>" style="vertical-align: middle; width: 17%;">
								<input type="hidden" name="btn_<%= rs("id_sito") %>" id="hdn_<%= estNome &"," %>" value="<%= estNome &"," %>">
								<a href="javascript:void(0);" name="btn_<%= rs("id_sito") %>" id="btn_<%= rs("id_sito") %>"
								   class="button_L2" 
								   onclick="apriGestione(new Array(<%= "form1.chk_" & Replace(estNome, ",", ",form1.chk_") %>),
														 '<%= rs("sito_prmEsterni_admin") %>', <%= request("ID") %>)">
									PERMESSI AGGIUNTIVI
								</a>
							</td>
						<% end if
					elseif estAbilita then %>
						<!--<td class="content_center">ciao</td>-->
						<%
					end if %>
				</tr>
				<% 
				lim_inf = (lim_inf + permessi_per_riga)
				lim_sup = lim_sup + permessi_per_riga
				if cIntero(lim_sup)>max_num_permessi then lim_sup=max_num_permessi 
				if cIntero(lim_inf) > max_num_permessi OR fine then exit for
			next	
			%>
			<%rs.MoveNext
		wend%>
	</table>
	<%rs.close
	set rsP = nothing
end sub

'.................................................................................................................
'		procedura che salva i permessi dell'utente
'			conn:			connessione aperta a database
'			rs:				recordset chiuso
'			amministrazione	flag che indica se l'utente di cui si devono salvare i permessi fa parte 
'							dell'area amministrativa o dell'area riservata.
'			utente			ID dell'utente
'.................................................................................................................
sub save_permessi(conn, rs, amministrazione, utente)
	dim permessi, i
	'cancella permessi precedenti se in modifica
	if amministrazione then
		sql = "DELETE FROM rel_admin_sito WHERE admin_id=" & utente
		if Session("PASS_ADMIN")="" then
			'non cancella l'eventuale permesso da amministratore
			sql = sql & " AND (sito_id<>" & Session("ID_SITO") & " OR rel_as_permesso<>1)"
		end if
	else
		sql = "DELETE FROM rel_utenti_sito WHERE rel_ut_id=" & utente
	end if
	CALL conn.execute(sql, 0, adExecuteNoRecords)
	
	'inserisce nuovi permessi
	if request("chk_perm")<>"" then	
		permessi = split(request.form("chk_perm"),",")
		i = 0
		while i<=ubound(permessi)
			if (i mod 2) = 0 then
				if amministrazione then
					sql = "INSERT INTO rel_admin_sito(admin_id, sito_id, rel_as_permesso) " 
				else
					sql = "INSERT INTO rel_utenti_sito(rel_ut_id, rel_sito_id, rel_permesso) " 
				end if
				sql = sql & " VALUES(" & utente & ", " & permessi(i) & "," & permessi(i+1) & ")"
				CALL conn.execute(sql, 0, adExecuteNoRecords)
			end if
			i = i + 2
		wend
	end if	
	
	if not amministrazione then
		'collega il contatto alla rubrica degli utenti dell'applicazione
		
		'cancela relazioni non piu attive
		sql = "DELETE FROM rel_rub_ind WHERE id_indirizzo IN (SELECT ut_nextCom_ID FROM tb_utenti WHERE ut_ID=" & utente & ")" & _
			  " AND id_rubrica IN (SELECT sito_rubrica_area_riservata FROM tb_siti)"
		CALL conn.execute(sql, 0, adExecuteNoRecords)
		
		'inserisce l'utente nelle rubriche presenti
		sql = "INSERT INTO rel_rub_ind (id_indirizzo, id_rubrica) " &_
			  " SELECT (SELECT ut_nextCom_ID FROM tb_utenti WHERE ut_ID=" & utente & "), sito_rubrica_area_riservata " &_
			  " FROM tb_siti WHERE id_sito IN (SELECT rel_sito_id FROM rel_utenti_sito WHERE rel_ut_id=" & utente & ")"
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	
end sub


'.................................................................................................................
'		Funzione che controlla il login: cerca tra gli altri utenti della sezione se c'e' gia' il login
'		in caso il login non fosse valido o fosse gia' usato da un'altro utente segnala il problema nalla variabile
'		di sessione "ERRORE"
'		Amministrazione:	flag che indica se l'ambito del login e' da considerare l'area amministrativa o l'area riservata
'		ID:					eventuale id del record dell'utente (solo in caso di modifica del login)
'		Value:				Valore di login da controllare
'.................................................................................................................
sub Check_login(conn, rs, Amministrazione, ID, Value)
	dim sql
	'controlla se il login e' corretto
	if Value<>"" AND CheckChar(Value, LOGIN_VALID_CHARSET) then
		
		'controlla se login gia' presente
		if Amministrazione then
			'utente dell'area amministrativa
			sql = "SELECT id_admin FROM tb_admin WHERE admin_login LIKE '" & Value & "'"
			if cInteger(ID)>0 then		'se l'utente e' gia' inserito lo confronta solo con gli altri
				sql = sql & " AND id_admin<>" & ID
			end if
		else
			'utente dell'area riservata.
			sql = "SELECT ut_id FROM tb_utenti WHERE ut_login LIKE '" & Value & "'"
			if cInteger(ID)>0 then		'se l'utente e' gia' inserito lo confronta solo con gli altri
				sql = sql & " AND ut_id<>" & ID
			end if
		end if
		
		'se trova piu qualche altro utente che ha lo stesso login blocca il salvataggio
		rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		if not rs.eof then
			Session("ERRORE") = "Login gi&agrave; utilizzato per un altro utente!"
		end if
		rs.close
	else
		Session("ERRORE") = "Login mancante o non valido! Utilizzare solo caratteri alfanumerici o &quot;_&quot;"
	end if
end sub
%>