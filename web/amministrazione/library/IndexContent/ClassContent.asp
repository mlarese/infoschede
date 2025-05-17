<%
'array di caratteri sostituire
const CODIFICA_SOSTITUISCI 		= " '""./"
'carattere sostituto
dim CODIFICA_SOSTITUTO
if cString(Application("CHAR_REPLACE_URL")) <> "" then
	CODIFICA_SOSTITUTO	= cString(Application("CHAR_REPLACE_URL"))
else
	CODIFICA_SOSTITUTO	= "_"
end if


'CLASSE DEL CONTENUTO DELL'INDICE
Class ObjContent

	'connessione DB
	Public conn
	
	'dictionary contenente i dati da salvare
	Public dizionario
	
	'dati caratterizzanti il contenuto
	Public co_F_table_id
	Public co_F_key_id
		
	
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'COSTRUTTORI CLASSE
'******************************************************************************************************************************************
	
	Private Sub Class_Initialize()
		set dizionario = request.form
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	Function GetID(co_F_table_id, co_F_key_id)
		if CIntero(co_F_table_id) = 0 then
			co_F_table_id = index.GetTable(co_F_table_id)
		end if
		GetID = GetValueList(conn, NULL, " SELECT co_id FROM tb_contents"& _
										 " WHERE co_F_table_id = "& cIntero(co_F_table_id) & _
										 " AND co_F_key_id = "& CIntero(co_F_key_id))
	End Function
    
	
	'restituisce il nome dato l'ID del contenuto
	Private Function Nome(ID)
		Nome = GetValueList(conn, NULL, "SELECT co_titolo_it FROM tb_contents WHERE co_id="& cIntero(ID))
	End Function
	
	
	'.................................................................................................
	'..		scrive il sistema di input per la selezione di una categoria
	'..		FormName		Nome del form in cui viene generato l'input
	'..		InputName 		Nome dell'input generato
	'..		InputValue		Valore/categoria selezionata
	'..		DisplayReduced	Indica se viene visualizzato un input 
	'.						ridotto (TRUE, per selezione nei motori di ricerca) o esteso (con link testuali)
	'..		disabled		disabilita l'input ma non l'hidden
	'.................................................................................................
	Public Sub WritePicker(FormName, co_F_table_id, co_F_key_id, InputName, InputValue, DisplayReduced, InputSize, disabled)
		dim ViewName, ViewValue
		
		if cInteger(InputValue)>0 then
			'recupera valore dell'input
			ViewValue = Nome(InputValue)
		else
			ViewValue = ""
		end if
		ViewName = "view_" & InputName %>
		<input type="hidden" name="<%= InputName %>" id="<%= InputName %>" value="<%= InputValue %>">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td <%= IIF(DisplayReduced, " colspan=""2"" ", " style=""padding-top:2px;"" ") %>>
					<input <%= DisableClass(disabled, "") %> READONLY type="text" name="<%= ViewName %>" value="<%= ViewValue %>" style="padding-left:3px;" size="<%= InputSize %>" onmouseover="<%= FormName %>_<%= InputName %>_UpdateTitle(this)" onclick="<%= FormName %>_<%= InputName %>_ApriFinestra('selezione')">
				</td>
			<% 	if DisplayReduced then %>	
				</tr>
				<tr>
			<% 	end if %>
			<% 	if not disabled then %>
				<td style="<%= IIF(DisplayReduced, "width:68%; padding-bottom:2px;", "padding-top:1px;") %>" nowrap>
					<a id="link_scegli_<%= InputName %>" class="button_input" href="javascript:void(0)" onclick="<%= FormName %>_<%= InputName %>_ApriFinestra('selezione')" title="Apre l'elenco delle voci indicizzate per selezionarne una." <%= ACTIVE_STATUS %>><%= IIF(DisplayReduced, "SCELGI CONTENUTO", "SCEGLI") %></a>
				</td>
                <td style="<%= IIF(DisplayReduced, "width:68%; padding-bottom:2px;", "padding-top:1px;") %>" nowrap>
                    <a id="link_reset_<%= InputName %>" class="button_input" href="javascript:void(0)" onclick="<%= FormName %>_<%= InputName %>_ApriFinestra('gestione')" title="Apre la scheda del contenuto per l'eventuale modifica dei dati." <%= ACTIVE_STATUS %>>COMPLETA I DATI</a>
                </td>
			<% 	end if %>
			</tr>
		</table>
        <script language="JavaScript" type="text/javascript">
            function <%= FormName %>_<%= InputName %>_ApriFinestra(tipo){
				var oInput = document.getElementById('<%= InputName %>');
				var href = '';
				if (tipo == 'gestione'){
					if (oInput.value != ''){
						href = '&tipo=gestione';
						if (oInput.disabled){
							href += '&selection_disabled=1';
						}
					}
				}
				else{
					if (!oInput.disabled){
						href = '&tipo=selezione';
					}
				}
				if (href != '')
					OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>library/IndexContent/ContentSeleziona.asp?' + 
												   'co_F_key_id=<%= co_F_key_id %>&co_F_table_id=<%= co_F_table_id %>&formname=<%= FormName %>&inputname=<%= InputName %>' + 
												   '&selected=' + oInput.value + href, tipo + '_contenuto', 760, 450, true);
												   
                <%= FormName %>_<%= InputName %>_UpdateTitle(<%= FormName %>.<%= ViewName %>)
            }
            
            function <%= FormName %>_<%= InputName %>_UpdateTitle(viewInput){
                viewInput.title = viewInput.value;
            }
            
            <%= FormName %>_<%= InputName %>_UpdateTitle(<%= FormName %>.<%= ViewName %>)
        </script>
    <%	
    End Sub


    '.................................................................................................
    '..		scrive il pulsante per l'apertura della finestra di cancellazione del contenuto
    '.................................................................................................
    Public Sub WriteDeleteButton(CssClass, co_id) %>
        <a class="button<%= CssClass %>" href="javascript:void(0);" 
           onclick="OpenAutoPositionedScrollWindow('<%= GetLibraryPath() %>IndexContent/DeleteIndexContenuto.asp?ID=<%= co_id %>', 'delete', 500, 300, false);">
            CANCELLA
        </a>
    <% end sub


    'verifica se il record del contenuto &egrave; un raggruppamento
    Public Function IsRaggruppamento(tab_name)
        IsRaggruppamento = ( instr(1, tab_name, tabRaggruppamentoTable, vbtextCompare)>0 )
    end function
	
	
	'verifica se il contenuto passato come parametro è un raggruppamento
	Public Function IsRaggruppamentoById(co_id)
		dim sql
		sql = " SELECT tab_name FROM tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + _
			  " WHERE co_id=" & co_id
		IsRaggruppamentoById = IsRaggruppamento(GetValueList(conn, NULL, sql))
	end function
	
	
	'verifica se il record del contenuto &egrave; un sito
    Public Function IsSito(tab_name)
        IsSito = ( instr(1, tab_name, tabSitoTable, vbtextCompare)>0 )
    end function
	
	
	'verifica se il record del contenuto &egrave; un sito
    Public Function IsPagina(tab_name)
        IsPagina = ( instr(1, tab_name, tabPagineTable, vbtextCompare)>0 )
    end function
	
	
	'lista di valodi del tb_web messa in cache per non essere ricalcolata ogni volta.
	Private cachedWebId  				'ID del sito per il quale e' stata calcolata e messa in cache l'home page e la lingua iniziale
	Private cachedWebHomePageId			'ID dell'home page
	Private cachedWebLinguaDefault 		'Lingua di default per il sito
	
	
	Private sub TbWebsCacheValues(webId)
		
		if cString(cachedWebLinguaDefault) = "" OR _
		   cIntero(cachedWebId) = 0 OR _
		   cIntero(cachedWebHomePageId) = 0 OR _
		   cIntero(cachedWebId) <> cIntero(webId) then
			dim wrs, sql
			set wrs = server.createObject("ADODB.Recordset")
			
			sql = "SELECT id_home_page, lingua_iniziale, id_webs FROM tb_webs WHERE id_webs = "& CIntero(webId)
			wrs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			
			if not wrs.eof then
				cachedWebLinguaDefault = lcase(cString(wrs("lingua_iniziale")))
				cachedWebHomePageId = cIntero(wrs("id_home_page"))
				cachedWebId = cIntero(wrs("id_webs"))
			else
				cachedWebLinguaDefault = ""
				cachedWebHomePageId = 0
				cachedWebId = 0
			end if
			
			wrs.close
			set wrs = nothing
		end if
	end sub
	
	
	'restituisce true se il record e l'home page
	Public Function IsHomePage(rs)
		if IsPagina(rs("tab_name")) then
			
			CALL TbWebsCacheValues(rs("idx_webs_id"))
			
			IsHomePage = ( rs("co_F_key_id") = cachedWebHomePageId )
			
		else
			IsHomePage = false
		end if
	End Function
	
	
	'restituisce la lingua inziale
	Public Function GetLinguaDefault(webId)
		CALL TbWebsCacheValues(webId)
		
		GetLinguaDefault = cachedWebLinguaDefault
	end function
	
    
    'verifica se il contenuto presente nel record esiste ed &egrave; stato creato correttamente
    Public function IsValid(ID) 
        dim rs, sql
        set rs = Server.CreateObject("ADODB.Recordset")
        
        sql = " SELECT * FROM tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + _
              " WHERE co_id = " & cIntero(ID)
        rs.open sql, conn, adOpenStatic, adLockOptimistic
        
        if not rs.eof then
            sql = "SELECT COUNT(*) FROM " & rs("tab_from_sql") & SQL_AddOperator(rs("tab_from_sql"), "AND") & rs("tab_field_chiave") & "=" & cIntero(rs("co_F_key_id"))
            rs.close
            IsValid = cIntero(GetValueList(conn, rs, sql))>0
        else
            rs.close
            IsValid = false
        end if
        
    end function
    
    
	' scrive il nome del contenuto piu il tipo
	Public Function WriteNomeETipo(rs)
		response.write rs("co_titolo_it") &"&nbsp;"
		CALL WriteTipoRS(rs)
	End Function
	
    
	'scrive il nome del tipo dato il recordset che lo contiene
	Sub WriteTipoRS(rs)
		if CString(rs("tab_colore")) <> "" then
			response.write "<span style=""color: "& rs("tab_colore") &";"">("& rs("tab_titolo") &")</span>"
		else
			response.write "("& rs("tab_titolo") &")"
		end if
	End Sub
	
    
	'scrive il nome del tipo dato l'ID
	Sub WriteTipo(ID)
		WriteTipoRS(conn.execute("SELECT * FROM tb_siti_tabelle WHERE tab_id = "& cIntero(ID)))
	End Sub
	
	
	'true se tutti i campi del recordset sono pieni
	Function IsAllFull(rs, prefisso)
		dim field
		IsAllFull = true
		for each field in rs.fields
			if Left(field.name, Len(prefisso)) = prefisso AND CString(rs(field.name)) = "" then
				IsAllFull = false
				exit for
			end if
		next
	End Function
	
	
	'drop down dei tipi dei contenuti (siti_tabelle)
	Public Sub DropDownTipi(selectNome, SelectSqlCondition, valSelected)
		dim sql, aux
		sql = " SELECT * FROM tb_siti_tabelle " + _
              IIF(SelectSqlCondition<>"", " WHERE " + SelectSqlCondition, "") + _
              " ORDER BY tab_sito_id, tab_titolo"
		set aux = conn.Execute(sql) %>
		<select name="<%= selectNome %>" style="width: 100%;">
	<%	if aux.eof then %>
			<option value="">Nessun tipo trovato</option>
	<%	else %>
			<option value="">scegli...</option>
	<%		while not aux.eof %>
			<option value="<%= aux("tab_id") %>"
					<%= IIF(CIntero(valSelected) = aux("tab_id"), " selected ", "") %>
					<%= IIF(CString(aux("tab_colore")) <> "", " style=""color: "& aux("tab_colore") &";"" ", "") %>>
				<%= aux("tab_titolo") %>
			</option>
	<%			aux.movenext
			wend
		end if %>
		</select>
	<%	set aux = nothing
	End Sub
	
	
	'restituisce true se il link e vincolato.
	'setta i campi link dell'rs (tabella tb_contents)
	Public Function LinkPrecalcola(rs, calcola)
		if CString(rs("tab_field_url_it")) <> "" then			'url vincolato
			LinkPrecalcola = true
		else
			LinkPrecalcola = false
			
			if calcola then 									'url precalcolato
				dim sql, aux, lingua
				sql = "SELECT TOP 1 * FROM tb_contents_index WHERE idx_content_id = "& cIntero(rs("co_id"))
				set aux = conn.Execute(sql)
				if not aux.eof then
					rs("co_link_tipo") = aux("idx_link_tipo")
					rs("co_link_pagina_id") = aux("idx_link_pagina_id")
					for each lingua in Application("LINGUE")
						rs("co_link_url_"& lingua) = aux("idx_link_url_"& lingua)
						rs("co_link_url_rw_"& lingua) = aux("idx_link_url_rw_"& lingua)
	 				next
	 			end if
				set aux = nothing
			end if
		end if
	End Function
	
	
    '.................................................................................................
    '..		funzioni per la gestione dei tag
    '.................................................................................................
	Public function GetTagsList(co_id, lingua)
		dim sql
		
		sql = "SELECT * FROM v_tags WHERE rct_content_id=" & co_id
		if cString(lingua)<>"" then
			sql = sql & " AND tag_lingua LIKE '" & lingua & "'"
		end if
		sql = sql & " ORDER BY tag_value "
		
		GetTagsList = GetValueList(conn, NULL, sql)
		
	end function
	
	
    '.................................................................................................
    '..		fumzioni per la gestione dei form
    '.................................................................................................
	
	'scrive il form per l'inserimento dati (inserimento/modifica).
	'l'oggetto index deve essere creato all'esterno della classe.
	Public Sub Modifica(ID)
		dim tab, rs, sql, i, LockIcon, value
		
		ID = CIntero(ID)
		if ID > 0 AND request.servervariables("REQUEST_METHOD") <> "POST" then
			set rs = conn.execute("SELECT * FROM tb_contents WHERE co_id = "& cIntero(ID))
			set dizionario = rs
			co_F_table_id = rs("co_F_table_id")
		end if
		
		set tab = server.createobject("adodb.recordset")
		sql = " SELECT * FROM tb_siti_tabelle INNER JOIN tb_siti ON tb_siti_Tabelle.tab_sito_id = tb_siti.id_sito " + _
			  " WHERE tab_id = "& cIntero(co_F_table_id)
		tab.open sql, conn, adOpenStatic, adLockOptimistic
		
		LockIcon = "<img src=""" & GetAmministrazionePath() & "grafica/padlock.gif"" style=""border:2px solid #f4F4F4; margin-left:2px; margin-right:2px;""" + _
				   "alt=""Valore modificabile esclusivamente dall'applicativo &ldquo;" + GetApplicationShortName(tab("sito_nome")) + _
				   "&rdquo; dalla sezione che gestisce i contenuti di tipo &ldquo;" & tab("tab_titolo") & "&rdquo;."">"
		
		if ID = 0 OR tab("tab_id") = index.GetTable(tabRaggruppamentoTable) then
			tab.AddNew		'se in inserimento o gestione tabella indice tolgo tutti i disabled
		end if %>
		<div id="content_ridotto">
		
			<form action="" method="post" id="form1" name="form1">
			<% 	if CString(tab("tab_field_titolo_it")) <> "" then %>
					<input type="hidden" name="co_titolo_it" value="<%= dizionario("co_titolo_it") %>">
			<% 	end if %>
			<% CALL WriteAdminIndexDataViewMode(true) %>
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<caption>
					Modifica dati complementari del contenuto di tipo &ldquo;<span style="color:<%= tab("tab_colore") %>;"><%= tab("tab_titolo") %></span>&rdquo;
				</caption>
				<tr><th colspan="4">DATI DEL CONTENUTO</th></tr>
				<tr>
					<td class="label_no_width" style="width:16%;">tipo di contenuto:</td>
					<td class="content" colspan="3" style="width:85%;"><% WriteTipoRS(tab) %></td>
				</tr>
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
					<% 	if i = 0 then %>
						<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
					<% 	end if %>
						<td class="content" colspan="3">
							<table width="100%" cellpadding="0" cellspacing="0">
								<tr>
									<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
									<% if CString(tab("tab_field_titolo_" + Application("LINGUE")(i))) <> "" then
										'campo sincronizzato: non modificabile 
										%>
										<td width="1%" valign="top"><%= LockIcon %></td>
										<td class="content_disabled">
											<%= TextHtmlEncode(dizionario("co_titolo_"& Application("LINGUE")(i))) %>
										</td>
									<% else %>
										<td class="content">
											<input type="text" class="text" name="co_titolo_<%= Application("LINGUE")(i) %>" value="<%= dizionario("co_titolo_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:95%;">
											<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "<span id=""obblN"">(*)</span>" end if %>
										</td>
									<% end if %>
								</tr>
							</table>
						</td>
					</tr>
				<% next
				for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
					<% 	if i = 0 then %>
						<td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">codice univoco:</td>
					<% 	end if %>
						<td class="content" colspan="3">
							<table width="100%" cellpadding="0" cellspacing="0">
								<tr>
									<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
									<% if CString(tab("tab_field_codice_" + Application("LINGUE")(i))) <> "" then
										'campo sincronizzato: non modificabile 
										%>
										<td width="1%" valign="top"><%= LockIcon %></td>
										<td class="content_disabled">
											<%= TextHtmlEncode(dizionario("co_chiave_"& Application("LINGUE")(i))) %>
										</td>
									<% else %>
									<td class="content">
										<input type="text" class="text" name="co_chiave_<%= Application("LINGUE")(i) %>" value="<%= dizionario("co_chiave_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:95%;">
										</td>
									<% end if %>
								</tr>
							</table>
						</td>
					</tr>
				<% next %>
				<tr><th class="L2" colspan="4">DESCRIZIONE</th></tr>
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="4">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<% if CString(tab("tab_field_descrizione_" + Application("LINGUE")(i))) <> "" then
									'campo sincronizzato: non modificabile 
									%>
									<td width="1%" valign="top"><%= LockIcon %></td>
									<td class="content_disabled">
										<%= TextHtmlEncode(dizionario("co_descrizione_"& Application("LINGUE")(i))) %>
									</td>
								<% else %>
									<td class="content">
										<textarea style="width:100%;" rows="3" name="co_descrizione_<%= Application("LINGUE")(i) %>"><%= dizionario("co_descrizione_" & Application("LINGUE")(i)) %></textarea>
									</td>
								<% end if %>
							</tr>
						</table>
					</td>
				</tr>
				<%next %>
				<tr><th class="L2" colspan="4">DATI DI PUBBLICAZIONE</th></tr>
				
				<tr>
					<td class="label_no_width">visibile:</td>
					<% if CString(tab("tab_field_visibile")) <> "" then  %>
						<td class="content_disabled">
							<%= LockIcon %>
							<%= IIF(dizionario("co_visibile") = "" OR CBool(dizionario("co_visibile")), "si", "no") %>
						</td>
					<% else %>
						<td class="content">
							<input type="radio" class="checkbox" value="1" name="co_visibile" <%= chk(dizionario("co_visibile") = "" OR CBool(dizionario("co_visibile"))) %>>
							si
							<input type="radio" class="checkbox" value="0" name="co_visibile" <%= chk(dizionario("co_visibile") <> "" AND CIntero(dizionario("co_visibile")) = 0 AND NOT CBool(dizionario("co_visibile"))) %>>
							no
						</td>
					<% end if %>
					
					<td class="label_no_width" style="width:15%;">ordine:</td>
					<% if CString(tab("tab_field_ordine")) <> "" then  %>
						<td class="content_disabled" style="width:42%;">
							<%= LockIcon %>
							<%= dizionario("co_ordine") %>&nbsp;
						</td>
					<% else %>
						<td class="content" style="width:42%;">
							<input type="text" class="text" name="co_ordine" value="<%= dizionario("co_ordine") %>" maxlength="<%= index.OrdineLenght %>" size="3">
						</td>
					<% end if %>
					</td>
				</tr>
				<tr>
					<td class="label_no_width">data pubblicazione:</td>
					<% if CString(tab("tab_field_data_pubblicazione")) <> "" then %>
						<td class="content_disabled">
							<%= LockIcon %>
							<%= dizionario("co_data_pubblicazione") %>&nbsp;
						</td>
					<% else %>
						<td class="content">
							<% CALL WriteDataPicker_Input("form1", "co_data_pubblicazione", dizionario("co_data_pubblicazione"), "", "/", true, false, LINGUA_ITALIANO) %>
						</td>
					<% end if %>
					<td class="label_no_width">data scadenza:</td>
					<% if CString(tab("tab_field_data_scadenza")) <> "" then %>
						<td class="content_disabled">
							<%= LockIcon %>
							<%= dizionario("co_data_scadenza") %>&nbsp;
						</td>
					<% else %>
						<td class="content">
							<% CALL WriteDataPicker_Input_Ex("form1", "co_data_scadenza", dizionario("co_data_scadenza"), "", "/", true, false, LINGUA_ITALIANO, "co_data_pubblicazione") %>
						</td>
					<% end if %>
				</tr>
				<%
				for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
					<% 	if i = 0 then %>
						<td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo alternativo:</td>
					<% 	end if %>
						<td class="content" colspan="3">
							<table width="100%" cellpadding="0" cellspacing="0">
								<tr>
									<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
									<% if CString(tab("tab_field_titolo_alt_" + Application("LINGUE")(i))) <> "" then
										'campo sincronizzato: non modificabile 
										%>
										<td width="1%" valign="top"><%= LockIcon %></td>
										<td class="content_disabled">
											<%= TextHtmlEncode(dizionario("co_alt_"& Application("LINGUE")(i))) %>
										</td>
									<% else %>
										<td class="content">
											<input type="text" class="text" name="co_alt_<%= Application("LINGUE")(i) %>" value="<%= dizionario("co_alt_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:95%;">
										</td>
									<% end if %>
								</tr>
							</table>
						</td>
					</tr>
				<% next %>
				<tr><th class="L2" colspan="4">FOTO</th></tr>
				<tr>
					<td class="label_no_width">thumbnail:</td>
					<td class="content" colspan="3">
						<% if CString(tab("tab_field_foto_thumb")) <> "" then %>
							<table cellpadding="0" cellspacing="0">
								<tr>
									<td width="1%" valign="top"><%= LockIcon %></td>
									<td><% CALL FileLink(GetUrlImage(dizionario("co_foto_thumb"), Application("AZ_ID")))  %></td>
								</tr>
							</table>
						<% else
							CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "co_foto_thumb", dizionario("co_foto_thumb"), "width:425px", false)
						end if %>
					</td>
				</tr>
				<tr>
					<td class="label_no_width">zoom:</td>
					<td class="content" colspan="3">
						<% if CString(tab("tab_field_foto_zoom")) <> "" then %>
							<table cellpadding="0" cellspacing="0">
								<tr>
									<td width="1%" valign="top"><%= LockIcon %></td>
									<td><% CALL FileLink(GetUrlImage(dizionario("co_foto_zoom"), Application("AZ_ID")))  %></td>
								</tr>
							</table>
						<% else
							CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "co_foto_zoom", dizionario("co_foto_zoom"), "width:425px", false)
						end if %>
					</td>
				</tr>
				<% if CString(tab("tab_field_url_it")) <> "" then %>
				    <tr><th colspan="4">LINK DEL CONTENUTO</th></tr>
				    <% if CIntero(dizionario("co_link_pagina_id")) > 0 AND CIntero(dizionario("co_link_tipo")) = lnk_interno then
					    for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
				            <tr>
					            <% if i = 0 then %>
					                <td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">pagina interna:</td>
					            <% end if %>
					            <td class="content" colspan="3">
                                    <img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
									<%= LockIcon %>
									<% CALL WritePageLink(conn, NULL, dizionario("co_link_pagina_id"), Application("LINGUE")(i))  %>
                                </td>
                            </tr>
                        <% next
					else
					    for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
				            <tr>
					            <% if i = 0 then %>
					                <td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">link esterno:</td>
					            <% end if %>
					            <td class="content" colspan="3">
                                    <img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
									<%= LockIcon %>
									<a href="<%= dizionario("co_link_url_"& Application("LINGUE")(i)) %>" title="apri il link &ldquo;<%= dizionario("co_link_url_"& Application("LINGUE")(i)) %>&rdquo; in una nuova finestra" target="_blank">
										<%= dizionario("co_link_url_"& Application("LINGUE")(i)) %>
									</a>
                                </td>
                            </tr>
                        <% next
					end if
				end if %>
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %>><th colspan="4">META TAG</th></tr>
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %>><th colspan="4" class="L2">KEYWORDS</th></tr>
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
						<td class="content" colspan="4">
							<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
								<tr>
									<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
									<% if CString(tab("tab_field_meta_keywords_" + Application("LINGUE")(i))) <> "" then
										'campo sincronizzato: non modificabile 
										%>
										<td width="1%" valign="top"><%= LockIcon %></td>
										<td class="content_disabled">
											<%= TextHtmlEncode(dizionario("co_meta_keywords_"& Application("LINGUE")(i))) %>&nbsp;
										</td>
									<% else %>
										<td class="content">
											<textarea style="width:100%;" rows="2" name="co_meta_keywords_<%= Application("LINGUE")(i) %>"><%= dizionario("co_meta_keywords_" & Application("LINGUE")(i)) %></textarea>
										</td>
									<% end if %>
								</tr>
							</table>
						</td>
					</tr>
				<%next %>
				
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %>><th colspan="4" class="L2">DESCRIPTION</th></tr>
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
						<td class="content" colspan="4">
							<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
								<tr>
									<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
									<% if CString(tab("tab_field_meta_description_" + Application("LINGUE")(i))) <> "" then
										'campo sincronizzato: non modificabile 
										%>
										<td width="1%" valign="top"><%= LockIcon %></td>
										<td class="content_disabled">
											<%= TextHtmlEncode(dizionario("co_meta_description_"& Application("LINGUE")(i))) %>&nbsp;
										</td>
									<% else %>
										<td class="content">
											<textarea style="width:100%;" rows="2" name="co_meta_description_<%= Application("LINGUE")(i) %>"><%= dizionario("co_meta_description_" & Application("LINGUE")(i)) %></textarea>
										</td>
									<% end if %>
								</tr>
							</table>
						</td>
					</tr>
				<%next %>
				<tr>
					<td class="footer" colspan="4">
						(*) Campi obbligatori.
						<% 	if IsAllFull(tab, "tab_") then %>
							<input type="button" class="button_link_like" onclick="javascript:history.go(-1)" name="salva" value="INDIETRO">
						<% 	else %>
							<input type="submit" class="button" name="salva" value="SALVA">
						<% 	end if %>
					</td>
				</tr>
			</table>
			&nbsp;
			</form>
		</div>
	</body>
	</html>
<%		set tab = nothing
	End Sub
	
	
	'funzione che recupera l'id del tag ed eventualmente lo aggiunge
	function GetTagId(tagValue, lingua)
		dim rs, sql
		
		set rs = server.CreateObject("ADODB.Recordset")
		
		sql = "SELECT * FROM tb_contents_tags WHERE tag_value LIKE '" & ParseSql(Trim(cString(tagValue)), adChar) & "' AND tag_lingua LIKE '" & ParseSql(Trim(cString(lingua)), adChar) & "'"	
		rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdtext
		if rs.eof then
			rs.addNew
			rs("tag_value") = Trim(cString(tagValue))
			rs("tag_lingua") = Trim(cString(lingua))
			rs.update
		end if
		GetTagId = rs("tag_id")
		rs.close
		
		set rs = nothing
	end function
	
	
	'associa il contenuto al tag
	Sub TaggaContenuto(tag_id, co_id, autogenerati)
		dim rs, sql
		
		set rs = server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM rel_contents_tags WHERE rct_content_id = " & co_id & " AND rct_tag_id = " & tag_id
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext

		if rs.eof then
			rs.AddNew
			rs("rct_content_id") = co_id
			rs("rct_tag_id") = tag_id
			rs("rct_autogenerato") = autogenerati
			rs.update
		end if
		
		rs.close
		
		set rs = nothing
	end sub
	
	
	'associa il conenuto alla lista di tag, eventualmente aggiungendo i tag
	Sub SaveTags(co_id, tags, lingua, autogenerati, charSeparator)
		dim taglist, tagValue, tagId
		
		taglist = split(RemoveByCharset(tags, TAGS_INVALID_CHARSET, true), charSeparator)
			
		for each tagValue in tagList
			if trim(tagValue)<>"" then
				'recupera id del tag, eventualmente lo aggiunge
				tagId = GetTagId(tagValue, lingua)
'response.write tagValue & "<br>"	
				'associa contenuto al tag.
				CALL TaggaContenuto(tagid, co_id, autogenerati)
			end if
		next
		
	end sub
	
	
	'rimuove l'associazione del contenuto dal tag.
	Sub RemoveTag(co_id, tag_id)
		dim sql, rs
		
		'rimuove associazione del tag
		sql = "DELETE FROM rel_contents_tags WHERE rct_content_id=" & co_id & " AND rct_tag_id=" & tag_id
		CALL conn.execute(Sql)
		
		'rimuove tag non associati ad almeno un contenuto
		CALL RemoveNotUsedTags()
		
	end sub
	
	
	'rimuove i tags di un contenuto
	Sub RemoveTags(co_id, soloAutogenerati)
		dim sql
		sql = "DELETE FROM rel_contents_tags WHERE rct_content_id=" & co_id
		if soloAutogenerati then
			sql = sql + " AND " & SQL_IsTrue(conn, "rct_autogenerato")
		end if
		CALL conn.execute(Sql)
		
		'rimuove tag non associati ad almeno un contenuto
		CALL RemoveNotUsedTags()
		
	end sub
	
	
	'rimuove tag non associati ad almeno un contenuto
	Sub RemoveNotUsedTags()
		dim sql
		sql = "DELETE FROM tb_contents_tags WHERE NOT EXISTS(SELECT * FROM rel_contents_tags WHERE rct_tag_id = tb_contents_tags.tag_id)"
		CALL conn.execute(sql)
		
	end sub
	
	
	'scrive il form per l'inserimento dei tag riguardanti il contenuto
	Sub Tags()
		dim rsc, rsi, sql, lingua, value, ClearedQueryString
		set rsc = server.createobject("adodb.recordset")
		set rsi = server.CreateObject("ADODB.recordset")
		
		ClearedQueryString = replace(request.serverVariables("QUERY_STRING"), "&rimuovitag=" & request.querystring("rimuovitag") & "&tagcontent_id=" & request.querystring("tagcontent_id"), "")
		
		if request.ServerVariables("REQUEST_METHOD") = "POST" AND _
		   cintero(request("co_id"))>0 AND _
		   request("tagsInput")<>"" then
			
			CALL SaveTags(request("co_id"), request("tagsInput"), request("lingua"), false, ",")
			
		elseif cIntero(request.querystring("rimuovitag"))<>0 AND cIntero(request.querystring("tagcontent_id"))<>0 then
			CALL RemoveTag(request.querystring("tagcontent_id"), request.querystring("rimuovitag"))
		end if
		%>
		<div id="content_ridotto">
			<%  'recupera tutti i contenuti con la stesso id provenienti dalla stessa tabella
			sql = " SELECT * FROM tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " & _
		  	  	  " WHERE co_F_key_id = "& cIntero(co_F_key_id) & " AND tb_siti_tabelle.tab_name LIKE (SELECT tab_name FROM tb_siti_tabelle WHERE tab_id=" & cIntero(co_F_table_id) & ")"
			rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
			
			while not rsc.eof %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
					<caption>Gestione tags del contenuto di tipo &ldquo;<span style="color:<%= rsc("tab_colore") %>;"><%= rsc("tab_titolo") %></span>&rdquo; </caption>
					<tr>
						<th colspan="5">DATI DEL CONTENUTO</th>
					</tr>
					<tr>
						<td class="label_no_width" style="width:8%;">contenuto:</td>
						<td class="content" colspan="3">
		                	<%= rsc("co_titolo_it") %>&nbsp;<% WriteTipoRS(rsc) %>
		               	</td>
						<td style="width:15%;" class="content_center" rowspan="2">
							<a class="button_block" style="padding-top:9px;padding-bottom:9px;"
							   href="ContentGestione.asp?FROM=tags&co_F_key_id=<%= rsc("co_F_key_id") %>&co_F_table_id=<%= rsc("co_F_table_id") %>" title="Modifica e completa i dati del contenuto." <%= ACTIVE_STATUS %>>
								COMPLETA I DATI
							</a>
						</td>
					</tr>
					<tr>
						<td class="label_no_width">visibile:</td>
						<td class="content">
							<input class="checkbox" type="checkbox" name="co_visibile" value="1" disabled <%= Chk(rsc("co_visibile")) %>>
						</td>
						<td class="label">data di pubblicazione:</td>
						<td class="content"><%= rsc("co_data_pubblicazione") %></td>
					</tr>
				</table>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
					<% 
					sql = " SELECT * FROM v_indice " & _
				  	  		 " WHERE co_F_table_id = "& cIntero(rsc("co_F_table_id")) & _
					  		 " AND co_F_key_id = "& cIntero(rsc("co_F_key_id")) & _
					  		 " ORDER BY idx_ordine_assoluto"
					rsi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
					<tr>
						<th colspan="5">COLLEGAMENTI ALL'INDICE DEL CONTENUTO</th>
					</tr>
					<% if rsi.eof then %>
						
					<% else %>
						<tr>
							<th class="l2">nome</th>
							<th class="l2_center" style="width:4%;">pubblicato</th>
							<th class="l2_center" style="width:5%;">principale</th>
							<th class="l2_center" style="width:16%;">modifica</th>
						</tr>
						<% 
						while not rsi.eof %>
							<tr>
								<td class="content"><% CALL index.WriteNodeLink(rsi, "", LINGUA_ITALIANO) %></td>
								<td class="content_center">
									<input class="checkbox" type="checkbox" name="visibile" value="1" disabled <%= Chk(rsi("visibile_assoluto")) %>>
								</td>
								<td class="content_center" <%= IIF(rsi("idx_principale"), "title=""Url principale di navigazione.""", "") %>>
									<input class="checkbox" type="checkbox" name="principale" value="1" disabled <%= Chk(rsi("idx_principale")) %>>
								</td>
								<% if rsi.absoluteposition = 1 then %>
									<td class="content_center" rowspan="<%= rsi.recordcount %>">
										<a class="button_block" <%= IIF(rsi.recordcount>1, "style=""padding-top:" & (9 * (rsi.recordcount - 1)) & "px; padding-bottom:" & (9 * (rsi.recordcount - 1)) & "px;""", "") %>
							   			   href="Indicizza.asp?FROM=tags&co_F_key_id=<%= rsc("co_F_key_id") %>&co_F_table_id=<%= rsc("co_F_table_id") %>" 
										   title="Gestione dei collegamenti all'indice del contenuto." <%= ACTIVE_STATUS %>>
										   COLLEGAMENTI
										</a>
									</td>
								<% end if %>
							</tr>
							<% rsi.movenext
						wend 
					end if
					
					rsi.close %>
				</table>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px;">
					<tr>
						<th colspan="4">TAGS ASSOCIATI</th>
					</tr>
					<tr>
						<td class="content notes" colspan="4">I tag devono essere inseriti separati da ",".</td>
					</tr>
					<% for each lingua in Application("LINGUE")
						sql = "SELECT * FROM v_tags WHERE rct_content_id=" & rsc("co_id") & " AND tag_lingua LIKE '" & lingua & "'" & " ORDER BY tag_value "
						rsi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
						<tr>
							<th class="L2" colspan="4">tag i lingua <%= GetNomeLingua(lingua) %></th>
						</tr>
						<tr>
							<td class="content_center" style="width:5%" rowspan="3">
								<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= lingua %>.jpg" alt="" border="0">
							</td>
							<% if rsi.eof then %>
								<td class="content note" colspan="3">
									nessun tag definito.
								</td>
							<% else %>
								<td class="content" colspan="3">
									<% while not rsi.eof %>									
										<span id="tag_<%= rsi("tag_id") %>_<%= rsc("co_id") %>" style="white-space:nowrap;">
											<%= rsi("tag_value") %>
											<% if not CBoolean(rsi("rct_autogenerato"), false) then%>
												<a href="?<%= ClearedQueryString %>&rimuovitag=<%= rsi("tag_id") %>&tagcontent_id=<%= rsc("co_id") %>" title="Rimuovi il tag." <%= ACTIVE_STATUS %>>
													<img src="../../grafica/rimuovi.gif" alt=""></a>
											<% end if %>	
										</span>
										<% rsi.movenext
										if not rsi.eof then %>
											,&nbsp;&nbsp;
										<% end if
									wend %>
								</td>
							<% end if %>
						</tr>
						<form action="?<%= ClearedQueryString %>" method="post" id="form_<%= rsc("co_id") %>_<%= lingua %>" name="form_<%= rsc("co_id") %>_<%= lingua %>">
							<input type="hidden" name="co_id" id="co_id_<%= rsc("co_id") %>_<%= lingua %>" value="<%= rsc("co_id") %>">
							<input type="hidden" name="lingua" id="lingua_<%= rsc("co_id") %>_<%= lingua %>" value="<%= lingua %>">
							<tr>
								<td class="label">aggiungi tags</td>
								<td class="content">
									<input type="text" name="tagsInput" id="tagsInput_<%= rsc("co_id") %>_<%= lingua %>" value="" class="text" style="width:100%;">
									<% 
									CALL Lightbox_Autocomplete_DIV("tagsInput_" & rsc("co_id") & "_" & lingua, "GestioneTags/autocompletamento_tags.asp")
									%>
								</td>
								<td class="content_center" style="width:15%; vertical-align:middle;">
									<input type="submit" class="button" name="tagsAdd" id="tagsAdd_<%= rsc("co_id") %>_<%= lingua %>" value="AGGIUNGI" style="width:100%;">
								</td>
							</tr>
						</form>
						<% rsi.close
					next %>
				</table>
				<% rsc.movenext
			wend 
			rsc.close
			%>
		</div>
		
		<%
	end sub
	
	
	'scrive il form per l'inserimento di una nuova associazione tra contenuto e indice
	Sub Associazioni() 
		dim tag_it, tag_en, tag_fr, tag_de, tag_es, tag_ru, tag_cn, tag_pt, i, lingua
		dim rsc, rsi, sql, value
		dim nCollegamenti, contUtenti, contCrawler, contAltro, contTotale, contGeneraleUtenti, contGeneraleCrawler, contGeneraleAltro, contGeneraleTotale
		
		set rsc = server.createobject("adodb.recordset")
		set rsi = server.CreateObject("ADODB.recordset")
		%>
		
		<div id="content_ridotto">
			<form action="" method="post" id="form1" name="form1">
			<%  'recupera tutti i contenuti con la stesso id provenienti dalla stessa tabella
			sql = " SELECT tab_colore, tab_titolo, co_titolo_it, co_F_key_id, co_F_table_id, co_visibile, co_data_pubblicazione, co_id, " & _
				  " tab_pagina_default_id, co_visibile " & _
				  " FROM tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " & _
			  	  " WHERE co_F_key_id = "& cIntero(co_F_key_id) & " AND tb_siti_tabelle.tab_name LIKE " & _
																				" (SELECT tab_name FROM tb_siti_tabelle WHERE tab_id=" & cIntero(co_F_table_id) & ")" & _
				  " ORDER BY "&SQL_IfIsNull(conn, "tab_priorita_base", "0")&" DESC, tab_id " 
			rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 

			nCollegamenti = 0
			contGeneraleUtenti = 0
			contGeneraleCrawler = 0
			contGeneraleAltro = 0
			contGeneraleTotale = 0
		
			if rsc.eof then %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px;">
					<caption>Collegamenti e pubblicazioni del contenuto</caption>
					<tr>
						<td colspan="5" class="label_no_width">Nessun contenuto trovato</td>
					</tr>
				</table>
				<%
			else
				'gestione pubblicazione "semi-automatica"
				if cIntero(rsc("tab_pagina_default_id")) > 0 AND cString(request("MODE"))<>"standard" AND rsc.RecordCount = 1 then 'AND cIntero(rsc("co_visibile")) = 1 
					sql = "SELECT idx_id, idx_autopubblicato, idx_padre_id FROM tb_contents_index WHERE idx_content_id = " & rsc("co_id")
					rsi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
					if rsi.eof then 'se non è già pubblicato sull'indice
						response.redirect GetAmministrazionePath() & "library/IndexContent/IndexPubblicaContenuto.asp" & _
										"?co_F_table_id="&rsc("co_F_table_id")&"&co_F_key_id="&rsc("co_F_key_id")&"&tab_pagina_default_id="&rsc("tab_pagina_default_id")
						response.end
					end if
					rsi.close
				end if
			end if		
			
			
			while not rsc.eof %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px;">
					<caption>Collegamenti e pubblicazioni del contenuto di tipo &ldquo;<span style="color:<%= rsc("tab_colore") %>;"><%= rsc("tab_titolo") %></span>&rdquo; </caption>
					<tr>
						<th colspan="5">DATI DEL CONTENUTO</th>
					</tr>
					<tr>
						<td class="label_no_width" style="width:8%;">contenuto:</td>
						<td class="content" colspan="3">
		                	<%= rsc("co_titolo_it") %>&nbsp;<% WriteTipoRS(rsc) %>
		               	</td>
						<td style="width:15%;" class="content_center" rowspan="2">
							<a class="button_block" style="padding-top:9px;padding-bottom:9px;"
							   href="ContentGestione.asp?FROM=Associazioni&co_F_key_id=<%= rsc("co_F_key_id") %>&co_F_table_id=<%= rsc("co_F_table_id") %>&MODE=<%= request("MODE") %>" title="Modifica e completa i dati del contenuto." <%= ACTIVE_STATUS %>>
								COMPLETA I DATI
							</a>
						</td>
					</tr>
					<tr>
						<td class="label_no_width">visibile:</td>
						<td class="content">
							<input class="checkbox" type="checkbox" name="co_visibile" value="1" disabled <%= Chk(rsc("co_visibile")) %>>
						</td>
						<td class="label">data di pubblicazione:</td>
						<td class="content"><%= rsc("co_data_pubblicazione") %></td>
					</tr>
					<% if TagAbilitati(conn) then %>
						<tr>
							<th colspan="5">TAGS DEL CONTENUTO</th>
						</tr>
						<%	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
							<tr>
								<% if i = 0 then %>
							         <td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">tags:</td>
					            <% end if %>
								<td class="content" colspan="3">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
											<% 
											sql = "SELECT tag_value FROM v_tags WHERE rct_content_id = " & rsc("co_id") & " AND tag_lingua LIKE'" & Application("LINGUE")(i) & "'"
											value = GetValueList(conn, rsi, sql)
											if value <> "" then%>
												<td class="content"><%= value %></td>
											<% else %>
												<td class="content notes"> tags non definiti </td>
											<% end if %>
										</tr>
									</table>
								</td>
								<% if i = 0 then %>
							         <td class="content_center" rowspan="<%= ubound(Application("LINGUE"))+1 %>">
										<a class="button_block" <%= IIF(ubound(Application("LINGUE"))>0, "style=""padding-top:" & (9 * ubound(Application("LINGUE"))) & "px; padding-bottom:" & (9 * ubound(Application("LINGUE"))) & "px;""", "") %>
							   			   href="Tagga.asp?FROM=Associazioni&co_F_key_id=<%= rsc("co_F_key_id") %>&co_F_table_id=<%= rsc("co_F_table_id") %>" 
										   title="Tagga il contenuto." <%= ACTIVE_STATUS %>>
										   GESTIONE TAGS
										</a>
									 </td>
					            <% end if %>
							</tr>
						<% next
					end if
					
					sql = ""
					for each lingua in Application("LINGUE")
						sql = sql & " idx_link_url_"& lingua & ","
					next
					sql = " SELECT "&sql&" idx_contUtenti, idx_contCrawler, idx_contAltro, idx_contatore, idx_ContRes, idx_principale, " & _
						  " idx_autopubblicato, idx_id, co_titolo_it, * " & _
						  " FROM v_indice " & _
					  	  " WHERE co_F_table_id = "& cIntero(rsc("co_F_table_id")) & _
						  " AND co_F_key_id = "& cIntero(rsc("co_F_key_id")) & _
						  " ORDER BY idx_ordine_assoluto"
					session("IDX_SQL") = sql
					rsi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
					<tr>
						<th colspan="5">DATI DEI COLLEGAMENTI ALL'INDICE</th>
					</tr>
					<tr>
						<td colspan="5">
							<table cellpadding="0" cellspacing="1" width="100%">
								<tr>
									<td class="label_no_width">
										<% if rsi.eof then %>
											Nessun collegamento inserito.
										<% else 
											nCollegamenti = nCollegamenti + rsc.recordcount%>
											Trovati n&ordm; <%= rsc.recordcount %> collegamenti
										<% end if %>
									</td>
									<td class="content_right" colspan="7">
										<a class="button_L2" title="apre in una nuova finestra l'elenco delle possibili associazioni" <%= ACTIVE_STATUS %>
										   href="IndicizzaAssocia.asp?co_F_table_id=<%= rsc("co_F_table_id") %>&co_F_key_id=<%= rsc("co_F_key_id") %>&MODE=<%= request("MODE") %>">
											NUOVO COLLEGAMENTO
										</a>
									</td>
								</tr>
								<tr>
									<th class="l2" rowspan="2">nome</th>
									<th class="l2_center" colspan="4" style="border-bottom: 0px;">statistiche accessi</th>
									<th class="l2_center" rowspan="2" style="width:4%;">pubblicato</th>
									<th class="l2_center" rowspan="2" style="width:5%;">principale</th>
									<th class="l2_center" rowspan="2" style="width:16%;">operazioni</th>
								</tr>
								<tr>
									<th class="l2_center" style="width:4%">utenti</th>
									<th class="l2_center" style="width:4%">crawler</th>
									<th class="l2_center" style="width:4%">altro</th>
									<th class="l2_center" style="width:4%">totale</th>
								</tr>
								<%
								contUtenti = 0
								contCrawler = 0
								contAltro = 0
								contTotale = 0
								while not rsi.eof
									contUtenti = contUtenti + CIntero(rsi("idx_contUtenti"))
									contCrawler = contCrawler + CIntero(rsi("idx_contCrawler"))
									contAltro = contAltro + CIntero(rsi("idx_contAltro"))
									contTotale = contTotale + CIntero(rsi("idx_contatore")) %>
									<tr>
										<td class="content"><% CALL index.WriteNodeLink(rsi, "", LINGUA_ITALIANO) %></td>
										<td class="content_right<%= IIF(rsi("idx_contUtenti") = 0, "_disabled", "") %>" title="a partire dal <%= DateTimeIta(rsi("idx_ContRes")) %>"><%= rsi("idx_contUtenti") %></td>
										<td class="content_right<%= IIF(rsi("idx_contCrawler") = 0, "_disabled", "") %>" title="a partire dal <%= DateTimeIta(rsi("idx_ContRes")) %>"><%= rsi("idx_contCrawler") %></td>
										<td class="content_right<%= IIF(rsi("idx_contAltro") = 0, "_disabled", "") %>" title="a partire dal <%= DateTimeIta(rsi("idx_ContRes")) %>"><%= rsi("idx_contAltro") %></td>
										<td class="content_right<%= IIF(rsi("idx_contatore") = 0, "_disabled", "") %>" title="a partire dal <%= DateTimeIta(rsi("idx_ContRes")) %>"><%= rsi("idx_contatore") %></td>
										<td class="content_center">
											<input class="checkbox" type="checkbox" name="visibile" value="1" disabled <%= Chk(rsi("visibile_assoluto")) %>>
										</td>
										<td class="content_center" <%= IIF(rsi("idx_principale"), "title=""Url principale di navigazione.""", "") %>>
											<input class="checkbox" type="checkbox" name="principale" value="1" disabled <%= Chk(rsi("idx_principale")) %>>
										</td>
										<td class="content_center" style="font-size:1px;">
											<a class="button_l2" href="IndicizzaAssocia.asp?co_F_table_id=<%= rsc("co_F_table_id") %>&co_F_key_id=<%= rsc("co_F_key_id") %>&ID=<%= rsi("idx_id") %>">
												MODIFICA
											</a>
											&nbsp;
                                    		<% if rsi("idx_autopubblicato") then %>
		                                        <a class="button_l2_disabled" href="javascript:void(0);" title="Impossibile cancellare la voce perch&egrave; fa parte delle seguenti pubblicazioni automatiche:<%= vbCrLF & index.GetPubblicazioniLockers(rsi("idx_id")) %>." <%= ACTIVE_STATUS %>>
													CANCELLA
												</a>
		                                    <% else 
		                                        CALL index.WriteDeleteButton("_L2", rsi("idx_id"))
		                                    end if %>
										</td>
									</tr>
									<% 
									rsi.movenext
								wend
								
								contGeneraleUtenti = contGeneraleUtenti + contUtenti
								contGeneraleCrawler = contGeneraleCrawler + contCrawler
								contGeneraleAltro = contGeneraleAltro + contAltro
								contGeneraleTotale = contGeneraleTotale + contTotale
								%>
								<tr>
									<td class="label_right" style="width: 50%;">totali:</td>
									<td class="content_right<%= IIF(contUtenti = 0, "_disabled", "") %>"><%= contUtenti %></td>
									<td class="content_right<%= IIF(contCrawler = 0, "_disabled", "") %>"><%= contCrawler %></td>
									<td class="content_right<%= IIF(contAltro = 0, "_disabled", "") %>"><%= contAltro %></td>
									<td class="content_right<%= IIF(contTotale = 0, "_disabled", "") %>"><%= contTotale %></td>
									<td class="content_center" colspan="3">&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<% rsi.close %>
				</table>
				<% rsc.movenext
			wend 
			
			'visualizza le statistiche globali uguali alla somma di tutte le statistiche di tutti i contenuti evidenziati.
			if rsc.recordcount>1 AND nCollegamenti > 0 then%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
					<caption class="border">Statistiche globali</caption>
					<tr>
						<th class="l2" style="width:51%;">statistiche generali accessi a tutti i contenuti</th>
						<th class="l2_center" style="width:4%">utenti</th>
						<th class="l2_center" style="width:4%">crawler</th>
						<th class="l2_center" style="width:4%">altro</th>
						<th class="l2_center" style="width:4%">totale</th>
						<th class="l2_center" style="width:27%;">&nbsp;</th>
					</tr>
					<tr>
						<td class="label_right">totali generali:</td>
						<td class="content_right<%= IIF(contGeneraleUtenti = 0, "_disabled", "") %>"><%= contGeneraleUtenti %></td>
						<td class="content_right<%= IIF(contGeneraleCrawler = 0, "_disabled", "") %>"><%= contGeneraleCrawler %></td>
						<td class="content_right<%= IIF(contGeneraleAltro = 0, "_disabled", "") %>"><%= contGeneraleAltro %></td>
						<td class="content_right<%= IIF(contGeneraleTotale = 0, "_disabled", "") %>"><%= contGeneraleTotale %></td>
						<td class="content_center">&nbsp;</td>
					</tr>
				</table>
			<% end if %>
			
			<script type="text/javascript">
				if (opener && opener.document) {
					var linkCollegamento = opener.document.getElementById("indicizza_<%= request("co_F_key_id") %>");
					if (linkCollegamento) {
						<% if nCollegamenti > 0 then  %>
							if (linkCollegamento.innerHTML != 'INDICE'){
								linkCollegamento.innerHTML = 'COLLEGATO ALL\'INDICE';
							}
							//linkCollegamento.style.color = '#444444';
							linkCollegamento.className = linkCollegamento.className.replace('DaIndicizzare','Indicizzato');
						<% else %>
							if (linkCollegamento.innerHTML != 'INDICE'){
								linkCollegamento.innerHTML = 'COLLEGA ALL\'INDICE';
							}
							//linkCollegamento.style.color = 'red';
							linkCollegamento.className =  linkCollegamento.className.replace('Indicizzato','DaIndicizzare');
						<% end if %>
					}
				}
			</script>
			
			<%
			rsc.close
			set rsc = nothing
			set rsi = nothing
			%>
			</form>
		</div>
	<% End Sub
	
	
	'controlla la validatà dei dati del dictionary.
	'ritorna false se errore e imposta session("ERRORE")
	Public Function ChkContent()
		if session("ERRORE") = "" then
			if CIntero(co_F_table_id) = 0 then
				session("ERRORE") = "Errore di sistema."
			elseif CString(dizionario("co_titolo_it")) = "" then
				session("ERRORE") = "Titolo italiano obbligatorio."
			end if
		end if
		ChkContent = (session("ERRORE") = "")
	End Function
	
	
	'esegue il chk e salva il contenuto.
	'restituisce l'ID, se vuoto ricerca per co_F_key_id e co_F_table_id in dizionario.
	Public Function Salva(ID)
		if ChkContent() then
			dim rs, campo, sql
			set rs = server.createobject("adodb.recordset")
			
			sql = "SELECT * FROM tb_contents WHERE "
			if CIntero(ID) > 0 then
				sql = sql &"co_id="& cIntero(ID)
			else
				sql = sql &"co_F_table_id = "& cIntero(co_F_table_id) &" AND co_F_key_id = "& CIntero(co_F_key_id)
			end if
			rs.open sql, conn, adOpenKeySet, adLockOptimistic
			if rs.eof then
				rs.addnew
				rs("co_F_table_id") = co_F_table_id
				rs("co_F_key_id") = co_F_key_id
				CALL SetUpdateParamsRS(rs, "co_", true)
			else
				co_F_table_id = rs("co_F_table_id")
				co_F_key_id = rs("co_F_key_id")
				CALL SetUpdateParamsRS(rs, "co_", false)
			end if

			for each campo in dizionario
				if FieldExists(rs, campo) AND LCase(campo) <> "co_f_key_id" then
					'response.write "campo:" & campo & "-->" & dizionario(campo) & "<br>"
					if dizionario(campo) = "" then
						rs(campo) = null
					else
						if instr(1, campo, "co_data_", vbTextCompare)>0 then
							rs(campo) = ConvertForSave_Date(dizionario(campo))
						else
							if rs(campo).type = 202 then
								rs(campo) = left(cString(dizionario(campo)), rs(campo).definedsize)
							else
								'response.write "campo:" & campo & "<br>"
								rs(campo) = dizionario(campo)
							end if
						end if
					end if
				end if
			next

			rs.update

			Salva = rs("co_id")
			'indico a Seleziona() il contenuto da scegliere
			Session("prmCo_SELECTED") = rs("co_id")
			
			rs.close
	
			'aggiorna i dati degli indici associati
			sql = " SELECT * FROM tb_contents_index WHERE idx_content_id = "& cIntero(Salva)
			rs.open sql, conn, adOpenStatic, adLockReadOnly
			while not rs.eof
				CALL index.operazioni_ricorsive_tipologia(rs("idx_id"))
				rs.movenext
			wend
			rs.close

			set rs = nothing
		end if
	End Function
	
	
	'salva solamente l'URL dell'indice prendendosi i dati del contenuto
	Sub SalvaIndexLink(contentID, rs, parametro)
		dim sql
		sql = "SELECT * FROM tb_contents_index WHERE idx_content_id = "& cIntero(contentID)
		rs.open sql, conn, adOpenStatic, adLockOptimistic
		
		if not rs.eof then
			index.SetIndexFromContent()
			while not rs.eof
				CALL LinkCalculate(conn, "idx", rs, index.dizionario, "idx_link_pagina_id", "idx_link_url_", parametro)
				rs.movenext
			wend
		end if
		
		rs.close
	End Sub
	
	
	'procedura che cancella il contenuto, tutte le relative pubblicazioni e le loro sottovoci.
	Sub Delete(ID)
		dim rs, sql
        
        'recupera pubblicazioni sull'indice del contenuto e loro voci discendenti
        sql = "SELECT idx_id FROM tb_contents_index WHERE idx_content_id = "& cIntero(ID)
        set rs = index.DiscendentiOrder(GetValueList(conn, rs, sql), "idx_livello DESC")
        
        'cancella pubblicazioni del contenuto e loro voci discendenti
        while not rs.eof
            call index.Delete(rs("idx_id"))
            rs.movenext
        wend
        rs.close
        
        'cancella contenuto
        sql = "DELETE FROM tb_contents WHERE co_id=" & cIntero(ID)
		CALL conn.execute(sql, ,adExecuteNoRecords)
        
		set rs = nothing
	End Sub
	
    
	'procedura che cancella il contenuto e tutte le sue indicizzazioni tramite nome tabella sorgente ed id del record.
	sub DeleteAll(F_table, F_key_id)
		dim rs, sql
		set rs = Server.CreateObject("ADODB.recordset")
        
        'recupera id dei contenuti collegati al record di origine
		sql = " SELECT DISTINCT co_id FROM tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " & _
			  " WHERE tab_name LIKE '"& ParseSQL(F_table, adChar) &"' AND co_F_key_id = "& cIntero(F_key_id)
        rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
        
        while not rs.eof
            'cancella contenuto, pubblicazioni e sotto-voci
            CALL Delete(rs("co_id"))
            rs.movenext
        wend
        
        rs.close
        
        set rs = nothing
    end sub
	
	
	'restituisce l'SQL da concatenare alla query per verificare i permessi (se non c'è la join con l'index)
	'la tabella tb_contents NON deve avere alias
	Function SQLPermessi()
		if index.ChkPrm(prm_indice_trasparente, 0) then
			SQLPermessi = " (1=1)"
		else
			'SQLPermessi = " (EXISTS ("& index.QueryElenco(false, "TIP_L0.idx_content_id = tb_contents.co_id")
			''tolgo l'order by
			'SQLPermessi = Left(SQLPermessi, InStr(1, SQLPermessi, "ORDER BY", vbTextCompare)-1) &")"
			SQLPermessi = " (EXISTS (SELECT 1 FROM tb_contents_index WHERE idx_content_id = tb_contents.co_id)"
			
			'aggiungo i permessi di visualizzazione dei contenuti non indicizzati ma creati dall'admin corrente
			SQLPermessi = SQLPermessi &" OR tb_contents.co_insAdmin_id = "& session("ID_ADMIN") & _
									   "    AND NOT EXISTS (SELECT 1 FROM tb_contents_index WHERE idx_content_id = tb_contents.co_id))"
		end if
	End Function
	
	
	'True se l'utente corrente ha i permessi di modifica sul contenuto
	Function ChkPrm(ID)
		if index.ChkPrm(prm_indice_trasparente, 0) then
			ChkPrm = true
		elseif CIntero(ID) = 0 then
			ChkPrm = false
		else
			dim sql
			sql = " SELECT COUNT(*) FROM tb_contents"& _
				  " WHERE co_id = "& cIntero(ID) &" AND "& SQLPermessi()
			
			ChkPrm = (CIntero(GetValueList(conn, NULL, sql)) > 0)
		end if
	End Function
	
	
	'True se l'utente corrente ha i permessi di modifica sul contenuto
	'	tab: 	nome della tabella co_F_table_id
	'	ID: 	co_F_key_id
	Function ChkPrmF(tab, ID)
		ChkPrmF = ChkPrm(GetValueList(conn, NULL, " SELECT co_id FROM tb_contents"& _
												  " WHERE co_F_key_id = "& cIntero(ID) & _
												  " AND co_F_table_id = "& CIntero(index.GetTable(tab))))
	End Function
	
	
'.............................................................................................................................
'.............................................................................................................................
'FUNZIONI PER GESTIONE SINCRONIZZATA
'.............................................................................................................................

'.............................................................................................................................
'Funzione che genera/aggiorna il contenuto dalla sorgente
'.............................................................................................................................
	Public sub GeneraDaTabella(byRef ContentId, TableRs, KeyValue)
		dim rsql, field, rsc, i, Column, k
		set rsc = Server.CreateObject("ADODB.Recordset")
		if not TableRs.eof then
			'recupera dati dalla sorgente
			rsql = ""
			k = 1
			for each field in TableRs.fields
				if instr(1, field.name, "_field", vbTextCompare)>0 then
					if cString(field.value)<>"" then
	                    if left(trim(field.value), 1) = "'" AND DB_Type(conn) = DB_sql then	
							Column = " (N" + field.value + ")"
						else
							Column = " (" + field.value + ")"
						end if
						Column = Column + " AS [" + field.name + "] "
						
	                    if instr(1, rsql, Column, vbTextCompare)<1 then
	                        rsql = rsql + IIF(rsql<>"", ", ", "") + Column
	                    end if
					end if
				end if
			next

			rsql = "SELECT TOP 1 " + rsql + " FROM " & TableRs("tab_from_sql") & _
				   SQL_AddOperator(TableRs("tab_from_sql"), "AND") & TableRs("tab_field_chiave") & " = " & cIntero(KeyValue)
			rsc.open rsql, conn, adOpenDynamic, adLockOptimistic
'response.write rsql & "<br>"
			if not rsc.eof then
				set dizionario = Server.CreateObject("Scripting.Dictionary")
				dizionario.compareMode = vbTextCompare
				
				for each i in Application("LINGUE")
					CALL GeneraDaTabella_AddField(rsc, "co_chiave_" + i, TableRs("tab_field_codice_" + i), "")
					dizionario("co_chiave_" + i) = Codifica(dizionario("co_chiave_" + i))
'response.end					
					CALL GeneraDaTabella_AddField(rsc, "co_titolo_" + i, TableRs("tab_field_titolo_" + i), "")
					CALL GeneraDaTabella_AddField(rsc, "co_descrizione_" + i, TableRs("tab_field_descrizione_" + i), "")
					CALL GeneraDaTabella_AddField(rsc, "co_meta_keywords_" + i, TableRs("tab_field_meta_keywords_" + i), "")
					CALL GeneraDaTabella_AddField(rsc, "co_meta_description_" + i, TableRs("tab_field_meta_description_" + i), "")
					CALL GeneraDaTabella_AddField(rsc, "co_alt_" + i, TableRs("tab_field_titolo_alt_" + i), "")
				next
				
				CALL GeneraDaTabella_AddField(rsc, "co_ordine", TableRs("tab_field_ordine"), "")
				CALL GeneraDaTabella_AddField(rsc, "co_foto_thumb", TableRs("tab_field_foto_thumb"), TableRs("tab_default_foto_thumb"))
'response.end
				CALL GeneraDaTabella_AddField(rsc, "co_foto_zoom", TableRs("tab_field_foto_zoom"), TableRs("tab_default_foto_zoom"))
		
				CALL GeneraDaTabella_AddField(rsc, "co_visibile", TableRs("tab_field_visibile"), "")
				'imposto valore di default se in inserimento
				if CIntero(contentId) = 0 AND NOT dizionario.Exists("co_visibile") then
					dizionario("co_visibile") = true
				end if
				CALL GeneraDaTabella_AddField(rsc, "co_data_pubblicazione", TableRs("tab_field_data_pubblicazione"), "")
				CALL GeneraDaTabella_AddField(rsc, "co_data_scadenza", TableRs("tab_field_data_scadenza"), "")
			
				'gestione link vincolato
				dim linkVincolato
				linkVincolato = false
				for each i in Application("LINGUE")
					if CString(tableRs("tab_field_url_" + i).value) <> "" then
						linkVincolato = true
						exit for
					end if
				next
				if linkVincolato then
					CALL LinkCalculate(conn, "co", dizionario, rsc, "tab_field_url_it", "tab_field_url_", _
									   tableRs("tab_parametro").value)
				end if
				
				co_F_table_id = cInteger(TableRs("tab_id"))
				co_F_key_id = KeyValue
				ContentId = Salva(ContentId)
			end if
			rsc.close
			
			if ContentId > 0 then
				'Giacomo 15/10/2012
				'Eliminio gli eventuali tag HTML dal campo descrizione
				rsql = ""
				for each i in Application("LINGUE")
					rsql = rsql + " co_descrizione_" + i + ","
				next
				rsql = Left(rsql, Len(rsql) - 1)
				rsql = "SELECT " & rsql & " FROM tb_contents WHERE co_id = " & ContentId
				rsc.open rsql, conn, adOpenDynamic, adLockOptimistic
				dim descr
				for each i in Application("LINGUE")
					descr = cString(rsc("co_descrizione_" + i))
					descr = RemoveHtmlTags(descr, " ")
					descr = Replace(descr, "  ", " ")
					descr = Trim(descr)
					if instr(descr, "  ") > 0 then
						descr = Replace(descr, "  ", " ")
					end if
					rsc("co_descrizione_" + i) = descr
				next
				
				rsc.update
				rsc.close
			end if
		end if
		set rsc = nothing
	end sub
	
	
'.............................................................................................................................
'Funzione che genera/aggiorna il contenuto dalla sorgente
'.............................................................................................................................
	private sub GeneraDaTabella_AddField(rsc, ContentField, ByRef TableField, defaultValue)
		dim val
		ContentField = cString(ContentField)
'response.write ContentField & ":" & defaultValue &".<br>"
		if CString(TableField.value) <> "" then
			if cString(rsc(tableField.name))="" AND cString(defaultValue)<>"" then
				val = defaultValue
			else
				val = rsc(tableField.name)
			end if
			dizionario.Add ContentField, val
			
		elseif cString(defaultValue)<>"" then
			dizionario.Add ContentField, defaultValue
		end if
	end sub
	
	
'.............................................................................................................................
'Funzione che codifica la stringa in ingresso per essere usata come codice
'.............................................................................................................................
	Public Function Codifica(str)
		if cString(str)<>"" then
			'riporta la string in minuscolo
			str = lcase(cstring(str))
			dim c, i
			
			'esegue codiifca per caratteri cirillici ed estesi
			str  = ExtendedUTFToBaseLatin(str)
			
			for i=1 to len(str)
				c = Mid(str, i, 1)
				if Server.HtmlEncode(c) <> c then
					Codifica = Codifica & CharToAsii(c, CODIFICA_SOSTITUTO)
				elseif instr(1, CODIFICA_SOSTITUISCI, c, vbTextCompare)>0 then
					Codifica = Codifica & CODIFICA_SOSTITUTO
				elseif instr(1, ALPHANUMERIC_CHARSET + "_", c, vbTextCompare)>0 then
					Codifica = Codifica & c
				end if
			next
			
			'ripulisce dai doppioni del carattere sostitutivo
			while instr(1, Codifica, CODIFICA_SOSTITUTO + CODIFICA_SOSTITUTO, vbTextCompare)>0
				Codifica = replace(codifica, CODIFICA_SOSTITUTO + CODIFICA_SOSTITUTO, CODIFICA_SOSTITUTO)
			wend
			
			'tolgo il carattere CODIFICA_SOSTITUTO dall'inizio e dalla fine
			if Codifica <> "" then
				if Left(Codifica, 1) = CODIFICA_SOSTITUTO then
					Codifica = Right(Codifica, Len(Codifica) - 1)
				end if
				if Right(Codifica, 1) = CODIFICA_SOSTITUTO then
					Codifica = Left(Codifica, Len(Codifica) - 1)
				end if
			end if
		end if
	End Function
	
	
'.............................................................................................................................
'Funzione che genera/aggiorna i tag del contenuto
'	id_contents				id del contenuto
'	stringa					stringa con i vari tag separati dal carattere "charSplit"
'	lingua					lingua dei tag
'	charSplit				carattere che decide la separazione dei tag in "stringa"
'.............................................................................................................................

	public sub generaTag(id_contents, stringa, lingua, charSplit)
		dim conn, sql, rs_tags, rs_rel, tagsAry, i, lengthTag
		set conn = Server.CreateObject("ADODB.Connection")
		set rs_tags = Server.CreateObject("ADODB.RecordSet")
		set rs_rel = Server.CreateObject("ADODB.RecordSet")
		set conn = Index.conn
		
		tagsAry = Split((RemoveInvalidChar(stringa, CHARSET)), charSplit)

		for i = 0 to Ubound(tagsAry)
			tagsAry(i) = Trim(tagsAry(i))
			lengthTag = Len(tagsAry(i)) 
			if lengthTag > 2 then
				sql = "SELECT * FROM tb_contents_tags WHERE tag_value LIKE '" & ParseSQL(tagsAry(i), adChar) & "'"
				rs_tags.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdtext
				if rs_tags.eof then
					rs_tags.addNew
					rs_tags("tag_value") = tagsAry(i)
					rs_tags("tag_lingua") = lingua
					rs_tags.Update
				end if
				sql = "SELECT * FROM rel_contents_tags WHERE rct_tag_id =" & rs_tags("tag_id") & " AND rct_content_id =" & id_contents
				rs_rel.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdtext
				if rs_rel.eof then
					rs_rel.addNew
					rs_rel("rct_content_id") = id_contents
					rs_rel("rct_tag_id") = rs_tags("tag_id")
					rs_rel("rct_autogenerato") = true
					rs_rel.Update
				end if
				rs_tags.close
				rs_rel.close
			end if
		next
	end sub

	
	
	
End Class


'.............................................................................................................................
'procedura che genera la porzione di interfaccia per nascondere / visualizzare i campi "avanzati" del contenuto e dell'indice
'.............................................................................................................................
const IndedDataViewMode_SIMPLE = "SIMPLE"
const IndedDataViewMode_ADVANCED = "FULL"

Sub WriteAdminIndexDataViewMode(IsContenuto)
	dim url, value
	if request.querystring("IndedDataViewMode")<>"" then
		Session("IndedDataViewMode") = request.querystring("IndedDataViewMode")
	elseif Session("IndedDataViewMode") = "" then
		Session("IndedDataViewMode") = IndedDataViewMode_SIMPLE
	end if
	
	value = "IndedDataViewMode=" & request.querystring("IndedDataViewMode")
	url = GetCurrentUrl() & "?" & request.serverVariables("QUERY_STRING")
	url = replace(url, "&" + value, "")
	url = replace(url, "?" + value, "")
	url = url + IIF(request.serverVariables("QUERY_STRING") <> "", "&", "?") + "IndedDataViewMode="
	
	if request.ServerVariables("REQUEST_METHOD")<>"POST" then
		%>
		
			<table cellpadding="0" cellspacing="1" class="tabella_madre" style="border-top-width:1px; margin-bottom:10px;">
			<tr>
				<td class="content notes">
					<%
					if Session("IndedDataViewMode") = IndedDataViewMode_SIMPLE then
						response.write IIF(IsContenuto, _
										   "Abilita la visualizzazione avanzata dei meta tag e dati aggiuntivi del contenuto.", _
										   "Abilita la visualizzazione avanzata dei meta tag, dei permessi di gestione e dei dati aggiuntivi del nodo.")
					else
						response.write IIF(IsContenuto, _
										   "Nascondi la visualizzazione avanzata dei meta tag e dati aggiuntivi del contenuto.", _
										   "Nascondi la visualizzazione avanzata dei meta tag, dei permessi di gestione e dei dati aggiuntivi del nodo.")
					end if 
					%>
				</td>
				<td class="content_right" style="width:10%;">
					<a class="button_L2_block" href="<%= url & IIF(Session("IndedDataViewMode") = IndedDataViewMode_SIMPLE, IndedDataViewMode_ADVANCED, IndedDataViewMode_SIMPLE)%>" 
					   style="width:200px;">
						<% ' IIF(Session("IndedDataViewMode") = IndedDataViewMode_SIMPLE, "VISUALIZZA META TAG", "NASCONDI META TAG") & IIF(IsContenuto, "", " E PERMESSI")
						response.write IIF(Session("IndedDataViewMode") = IndedDataViewMode_SIMPLE, "VISUALIZZA", "NASCONDI") & " DATI AGGIUNTIVI" %>
					</a>
				</td>
			</tr>
		</table>
		
	<% end if
	
end sub


function AdminIndexDataViewMode_CssStyle(writeStyle)
	if Session("IndedDataViewMode") = IndedDataViewMode_SIMPLE then
		AdminIndexDataViewMode_CssStyle = "display:none;"
		if writeStyle then
			AdminIndexDataViewMode_CssStyle = " style=""" & AdminIndexDataViewMode_CssStyle & """ "
		end if
	end if
end function

%>