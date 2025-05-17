<!--#INCLUDE FILE="../library/editorHTML/ckeditor/Tools_CKEditor.asp" -->
<%
'.................................................................................................
'..			SERIE DI FUNZIONI PER LA GESTIONE DI TABELLE TIPO DESCRITTORI
'.................................................................................................
dim tipiDescrittori(16, 3)
'testo normale
tipiDescrittori(0, 0) = adVarChar			'200
tipiDescrittori(0, 1) = "Testo"
tipiDescrittori(0, 2) = "Testo"
tipiDescrittori(0, 3) = "Text"

'numero
tipiDescrittori(1, 0) = adNumeric			'131
tipiDescrittori(1, 1) = "Numerico"
tipiDescrittori(1, 2) = "Numerico"
tipiDescrittori(1, 3) = "Numeric"

'valuta
tipiDescrittori(2, 0) = adCurrency			'6
tipiDescrittori(2, 1) = "Valuta"
tipiDescrittori(2, 2) = "Valuta"
tipiDescrittori(2, 3) = "Currency"

'valore true/false
tipiDescrittori(3, 0) = adBoolean			'11
tipiDescrittori(3, 1) = "Si/No"
tipiDescrittori(3, 2) = "boolean"
tipiDescrittori(3, 3) = "Yes/No"

'valore data
tipiDescrittori(4, 0) = adDate				'7
tipiDescrittori(4, 1) = "Data"
tipiDescrittori(4, 2) = "Data"
tipiDescrittori(4, 3) = "Data"

'link ad un file
tipiDescrittori(5, 0) = adBinary			'128
tipiDescrittori(5, 1) = "File"
tipiDescrittori(5, 2) = "Link"
tipiDescrittori(5, 3) = "File"

'testo lungo
tipiDescrittori(6, 0) = adLongVarChar		'201
tipiDescrittori(6, 1) = "Testo lungo"
tipiDescrittori(6, 2) = "TestoLungo"
tipiDescrittori(6, 3) = "Long text"

'link ad una risorsa esterna
tipiDescrittori(7, 0) = adUserDefined		'132
tipiDescrittori(7, 1) = "Link ad una risorsa esterna"
tipiDescrittori(7, 2) = "Link"
tipiDescrittori(7, 3) = "External link"

'anagrafiche selezionate dal NEXT-com
tipiDescrittori(8, 0) = adIUnknown			'13
tipiDescrittori(8, 1) = "Elenco anagrafiche"
tipiDescrittori(8, 2) = "Link"
tipiDescrittori(8, 3) = "Registries list"

'pagine selezionate dal NEXT-web
tipiDescrittori(9, 0) = adGUID				'72
tipiDescrittori(9, 1) = "Pagine NEXT-web"
tipiDescrittori(9, 2) = "Link"
tipiDescrittori(9, 3) = "NEXT-web pages"

'valori numerici min/max
tipiDescrittori(10, 0) = adDouble			'5
tipiDescrittori(10, 1) = "Min/Max"
tipiDescrittori(10, 2) = "doppio"
tipiDescrittori(10, 3) = "Min/Max"

'colore valido
tipiDescrittori(11, 0) = adPropVariant		'138		'ricerca non implementata
tipiDescrittori(11, 1) = "Colore HTML"
tipiDescrittori(11, 2) = "Colore"
tipiDescrittori(11, 3) = "HTML Color"

'collegamento all'indice
tipiDescrittori(12, 0) = adChapter			'136		'ricerca non implementata
tipiDescrittori(12, 1) = "Collegamento alla voce"
tipiDescrittori(12, 2) = "indice"
tipiDescrittori(12, 3) = "Link to item"

'rubrica
tipiDescrittori(13, 0) = adIDispatch		'9
tipiDescrittori(13, 1) = "Rubrica"
tipiDescrittori(13, 2) = "Rubrica"
tipiDescrittori(13, 3) = "Address book"

'amministratore
tipiDescrittori(14, 0) = adSingle			'4
tipiDescrittori(14, 1) = "Amministratore"
tipiDescrittori(14, 2) = "Amministratore"
tipiDescrittori(14, 3) = "Administrator"

'link ad un file protetto da password
tipiDescrittori(15, 0) = adChar			'129
tipiDescrittori(15, 1) = "File Protetto"
tipiDescrittori(15, 2) = "Link"
tipiDescrittori(15, 3) = "ProtectedFile"

'link ad una directory
tipiDescrittori(15, 0) = adVarBinary			'204
tipiDescrittori(15, 1) = "Directory"
tipiDescrittori(15, 2) = "Directory"
tipiDescrittori(15, 3) = "Direcotry"

'Codice HTML editabile con CKEditor
tipiDescrittori(16, 0) = adWChar				'130
tipiDescrittori(16, 1) = "Codice HTML"
tipiDescrittori(16, 2) = "Codice HTML"
tipiDescrittori(16, 3) = "HTML Code"


dim tipiDisable, UseSingleLanguage, FormName, AutocompletamentoValori, AutoCompletionListSource


'impostazione di default
AutocompletamentoValori = false
AutoCompletionListSource = ""


function AutoCompleteHTML(HtmlFieldId, lingua)
	if AutocompletamentoValori AND AutoCompletionListSource <>"" then
		CALL Lightbox_Autocomplete_DIV(HtmlFieldId, AutoCompletionListSource & "?lingua=" & IIF(cString(lingua)<>"", lingua, LINGUA_ITALIANO))
	end if
end function


function AutocompleteList(relTableDes, relCampoDesId, relCampoValore)
	dim inputField, valueToSearch, des_id, des_tipo, lingua, sql
	inputField = request.querystring("input")
	lingua = request.querystring("lingua")
	valueToSearch = request.form(inputField)
	if instr(1, inputField, "descr", vbTextCompare)>0 then
		des_id = right(inputField, len(inputField) - (instr(inputField, "descr") + 4))
		des_id = right(des_id, len(des_id) - (instr(des_id, "_")))
	else
		des_id = right(inputField, len(inputField)-instr(inputField, "_"))
	end if
	des_tipo = Right(inputField, Len(inputField)-3)
	des_tipo = Left(des_tipo, Instr(des_tipo, "_")-1)
	
	sql = " SELECT DISTINCT CAST(" & relCampoValore & lingua & " AS nvarchar(300)) " & _
		  " FROM " & relTableDes & _
		  " WHERE " & relCampoDesId & "=" & des_id & " AND " & relCampoValore & lingua & " LIKE '" & valueToSearch & "%' "
	CALL Lightbox_Autocomplete_QUERY(sql, "")
	
end function


FormName = "form1"
tipiDisable = replace(session("DES_TIPI_DISABLE"), " ", "") &","
if Session("DES_USE_SINGLE_LANGUAGE") then
	UseSingleLanguage = true
else
	UseSingleLanguage = false
end if


'dropdown con l'elenco dei tipi
Sub DesDropTipi(nome, stile, valore)
	CALL DesAdvancedDropTipi(nome, stile, valore, true)
End Sub


Sub DesAdvancedDropTipi(nome, stile, valore, mandatory)
	dim i 
	%>
	<select <%= stile %> name="<%= nome %>">
	
		<% if not mandatory then %>
			<option value="" <% if valore=0 then %> selected <% end if %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "scegli...", "choose...", "", "", "", "", "", "")%></option>
		<% end if
		for i = 0 to UBound(tipiDescrittori, 1)	
			if InStr(tipiDisable, tipiDescrittori(i, 0) &",") < 1 then 
				if Session("LINGUA") = "en" then %>
					<option value="<%= tipiDescrittori(i, 0) %>" <%= IIF(CIntero(valore) = tipiDescrittori(i, 0), "selected", "") %>><%= tipiDescrittori(i, 3) %></option>
				<% else %>
					<option value="<%= tipiDescrittori(i, 0) %>" <%= IIF(CIntero(valore) = tipiDescrittori(i, 0), "selected", "") %>><%= tipiDescrittori(i, 1) %></option>
				<% end if %>
			<% end if
		next %>
	</select>
<% end sub


'visualizza il tipo in formato testo
Function DesVisTipo(tipo)
	dim i
	for i = 0 to UBound(tipiDescrittori, 1)
		if tipo = tipiDescrittori(i, 0) then
			if Session("LINGUA") = "en" then
				DesVisTipo = tipiDescrittori(i, 3)
					exit for
			else
				DesVisTipo = tipiDescrittori(i, 1)
				exit for
			end if
		end if
	next
End Function


'righe con l'elenco dei descrittori per comporre il form di inserimento e modifica
'inInserimento: TRUE se il form è di inserimento
'colspan: numero colonne nella tabella
Sub DesElenco(conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, relCampoValore, inInserimento, colspan)
	CALL DesElenco_EXT(conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, relCampoValore, relCampoValore, inInserimento, colspan)
end sub

Sub DesElenco_EXT(conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, relCampoValore, relCampoValoreLungo, inInserimento, colspan)
	CALL DesForm(conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, "", relCampoValore, relCampoValoreLungo, "", inInserimento, colspan)
End Sub

Sub DesForm(conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, desLocked, relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan)
	CALL DesFormConn(NULL, conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, desLocked, relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan)
End Sub

'connContent:		connessione per recupero dati delle rubriche, admin, pagine, ecc. (vedi import parametri)
Sub DesFormConn(connContent, conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, desLocked, relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan)
	CALL DesFullFormConn(connContent, conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, "", desLocked, "", relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan)
end sub

'desCampoCodice:	campo codice da visualizzare
sub DesFullFormConn(connContent, conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM,  desCampoNote, desLocked, desCampoCodice, relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan)
	CALL DesFullFormComplete(connContent, conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM,  desCampoNote, desLocked, desCampoCodice, relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan, false, "")	
end sub	
	
'typeThBig		decide se il th è L2 o no
'widthColumn    decide la larghezza, in percentuale, della colonna contenente il nome dei descrittori
sub DesFullFormComplete(connContent, conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM,  desCampoNote, desLocked, desCampoCodice, relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan, typeThBig, widthColumn)	
	CALL DesFullFormComplete2(connContent, conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM,  desCampoNote, desLocked, desCampoCodice, relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan, typeThBig, widthColumn, "")
end sub

'typeThBig		decide se il th è L2 o no
'widthColumn    decide la larghezza, in percentuale, della colonna contenente il nome dei descrittori
sub DesFullFormComplete2(connContent, conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM,  desCampoNote, desLocked, desCampoCodice, relCampoValore, relCampoValoreLungo, grpNome, inInserimento, colspan, typeThBig, widthColumn, desCampoNumRowsTextarea)
	dim rs, i, unita_misura, postfisso_lingua, gruppo, disabled
	dim nascondi_des, almeno_un_valore, des_valorizzato, rowsTextarea, only_italian
	nascondi_des = false
	almeno_un_valore = false
	des_valorizzato = false
	only_italian = false
	
	if UseSingleLanguage then
		postfisso_lingua = ""
	else
		postfisso_lingua = "IT"
	end if
	gruppo = ""
	
	'modifica, Giacomo 18/12/2013, se ho fatto il post della pagina mi comporto come in inserimento, ovvero vado a leggere request
	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		inInserimento = true
	end if
	'-------------------------------------------------------
	
	if LBound(Application("LINGUE")) = UBound(Application("LINGUE")) then
		only_italian = true
	end if
	
	'set rs = conn.execute(sql)
	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	
	if rs.eof then %>
		<tr>
			<td class="content" colspan="<%= colspan %>"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Nessuna caratteristica trovata.", "No feature found.", "", "", "", "", "", "")%></td>
		</tr>
	<%else
		if IsNull(connContent) then
			set connContent = conn
		end if
		
		while not rs.eof
			
			'imposto il numero di righe delle textarea
			if desCampoNumRowsTextarea <> "" then
				rowsTextarea = cIntero(rs(desCampoNumRowsTextarea))
				if rowsTextarea = 0 then rowsTextarea = 5 end if
			else
				rowsTextarea = 5
			end if
	
			'gestione descrittori nel caso in cui si cambi la categoria e vi siano dei valori di descrittori associati alla vecchia categoria
			dim class_err
			if FieldExists(rs, "rct_tipologia_id") then
				if cIntero(rs("rct_tipologia_id")) = 0 then
					nascondi_des = true
				else
					nascondi_des = false
				end if
			else
				nascondi_des = false
			end if
			
			if FieldExists(rs, "rel_id") then
				if cIntero(rs("rel_id")) > 0 then
					des_valorizzato = true
				else
					des_valorizzato = false
				end if
			else
				des_valorizzato = false
			end if
			
			if nascondi_des AND des_valorizzato then
				class_err = "alert_descrittore"
				almeno_un_valore = true
			else
				class_err = ""
			end if

			
			unita_misura = ""
			if desCampoUnitaM <> "" then
				if cString(rs(desCampoUnitaM))<>"" then
					unita_misura = "&nbsp;" & rs(desCampoUnitaM)
				end if
			end if
			
			'gestione disabled
			if desLocked <> "" then
				if rs(desLocked) then
					disabled = "disabled"
				else
					disabled = ""
				end if
			end if
			
			'gestione gruppo
			dim valore, soloBooleanDesc, valoreOriginale
			if grpNome <> "" AND (not nascondi_des OR des_valorizzato) then
				if gruppo <> CString(rs(grpNome)) then
					gruppo = rs(grpNome)
					%>
					<tr>
						<th <%=IIF(typeThBig,"","class=""L2""")%> colspan="<%= colspan %>"><%= gruppo %></th>
					</tr>
					<%
					'------------ Verifico se tutti i descrittori di questo raggruppamento sono Boolean. Se è così allora imposto una diversa visualizzazione (Giacomo 30/05/2013)
					dim stepBack, tipoDesc
					stepBack = 0
					tipoDesc = ""
					soloBooleanDesc = true
					if not rs.eof then
						Do While not rs.eof
							If rs(grpNome) <> gruppo Then Exit Do
							If rs(desCampoTipo) <> adBoolean then
								soloBooleanDesc = false
								Exit Do
							end if
							stepBack = stepBack + 1
							rs.moveNext
						Loop
						rs.Move -(stepBack)
					end if
					'------------
				end if
			end if 
			

			if not nascondi_des OR des_valorizzato then
				
				'------------ come sopra, Giacomo 30/05/2013
				if cBoolean(soloBooleanDesc, false) then
					%>
					<tr>
						<td colspan="<%= colspan %>">
							<table cellspacing="1" cellpadding="0" class="tabella_madre" style="width:100%; border:0px;">
								<% 
								dim nColonne
								nColonne = 4
								For i = 1 To stepBack
									if ((i) mod nColonne) = 1 then
									%>
									<tr>
									<%
									end if
										%>
										<td class="content" style="width:1%;">
											<input type="checkbox" <%=  DisableClass(disabled<>"", "checkbox") %> name="<%= "des"& adBoolean &"_"& rs(desCampoID) %>" value="1" <%= Chk(IIF(inInserimento, request.form("des"& adBoolean &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua))<>"") %>>
										</td>
										<td class="label nomedescrittore <%=class_err%>">
											<%= CBLE(rs, desCampoNome, Session("LINGUA")) %>
											<% if desCampoCodice<>"" then %>
												<span class="notes codicedescrittore">( <%= rs(desCampoCodice) %> )</span>
											<% end if %>
										</td>
										<%
										if (i = (stepBack)) then
											dim colspanBool
											colspanBool = (i mod nColonne)
											colspanBool = nColonne - colspanBool
											if colspanBool < nColonne then
												%>
												<td class="content" colspan="<%=colspanBool*2%>">&nbsp;</td>
												<%
											end if
										end if										
									if ((i) mod nColonne) = 0 OR (i = (stepBack)) then
									%>
									</tr>
									<%
									end if
									rs.moveNext
								Next
								rs.movePrevious
								%>
							</table>
						</td>
					</tr>
					<%
				else
				'------------
				
				%>
				<tr>
					<td class="label nomedescrittore <%=class_err%>" <% if cString(desCampoNote)<>"" then %> <%if cString(rs(desCampoNote))<>"" then %> rowspan="2" <%end if%> <% end if%> <%=IIF(widthColumn<>"","style=""width:"&widthColumn&"%;""","")%>>
						<%= CBLE(rs, desCampoNome, Session("LINGUA")) %>
						<% if desCampoCodice<>"" then %>
							<span class="notes codicedescrittore">( <%= rs(desCampoCodice) %> )</span>
						<% end if %>
					</td>
					<td class="content <%=class_err%>" colspan="<%= colspan-1 %>">
						<% SELECT CASE rs(desCampoTipo) 
							CASE adNumeric
						'		response.write "valore db=" & rs(relCampoValore & postfisso_lingua) & "<br>"
							'	response.write "inInserimento=" & inInserimento & "<br>"
							'	response.write "iif=" & IIF(inInserimento, request.form("des"& adNumeric &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)) & "<br>"
								valoreOriginale = IIF(inInserimento, request.form("des"& adNumeric &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua))
							'	response.write "valoreOriginale=" & valoreOriginale & "<br>"
								valore = cReal(valoreOriginale)
							'	response.write "valore=" & valore & "<br>"
								%>

								<input type="text" <%= DisableClass(disabled<>"", "text") %> name="des<%= adNumeric %>_<%= rs(desCampoID) %>" id="des<%= adNumeric %>_<%= rs(desCampoID) %>" size="5" maxlength="20" value="<%= IIF(valore=0,"",valoreOriginale) %>"><%= unita_misura %>
								<% CALL AutoCompleteHTML("des" & adNumeric & "_" & rs(desCampoID), postfisso_lingua) %>
								<% 	
											
							CASE adCurrency 
								valoreOriginale = IIF(inInserimento, request.form("des"& adCurrency &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua))
								valore = cIntero(valore)
								
								if valore=0 then
									valore = valoreOriginale
									unita_misura = ""
								else
									valore = FormatPrice(valoreOriginale, 2, TRUE)
									if unita_misura = "" then
										unita_misura = "&nbsp;&euro;"
									end if
								end if
							%>
								<input type="text" <%=  DisableClass(disabled<>"", "text") %> name="des<%= adCurrency %>_<%= rs(desCampoID) %>" id="des<%= adCurrency %>_<%= rs(desCampoID) %>" size="11" maxlength="20" value="<%= valore %>"><%= unita_misura %>
								<% CALL AutoCompleteHTML("des" & adCurrency & "_" & rs(desCampoID), postfisso_lingua) %>
						<% 	CASE adBoolean %>
								<input type="checkbox" <%=  DisableClass(disabled<>"", "checkbox") %> name="<%= "des"& adBoolean &"_"& rs(desCampoID) %>" value="1" <%= Chk(IIF(inInserimento, request.form("des"& adBoolean &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua))<>"") %>>
						<% 	CASE adBinary, adChar, adVarBinary
								if UseSingleLanguage then
									if disabled = "" then
										CALL WriteFileSystemPicker_Input(Application("AZ_ID"), _
																		 IIF(rs(desCampoTipo) = adVarBinary, FILE_SYSTEM_DIRECTORY, FILE_SYSTEM_FILE), _
																		 "images", _
																		 "", _
																		 FormName, _
																		 "des"& rs(desCampoTipo) &"_"& rs(desCampoID), _
																		 IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), _
																		 "", _
																		 FALSE, _
																		 FALSE) 
									else %>
										<input type="text" disabled class="disabled" style="width: 95%;" name="<%= "des"& rs(desCampoTipo) &"_"& rs(desCampoID) %>" value="<%= IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)) %>">
						<%			end if
								else %>
									<table cellpadding="0" cellspacing="0" width="100%">
										<%for i = LBound(Application("LINGUE")) to UBound(Application("LINGUE")) %>
											<tr>
												<% if not only_italian then %>
													<td class="content" width="1%">
														<img style="vertical-align:top;" src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" alt="" border="0">&nbsp;
													</td>
												<% end if %>
												<td class="content" <%=IIF(only_italian, "colspan=""2"" style=""width:100%;""", "style=""width:99%;""") %>>
												<% 	if disabled = "" then
													CALL WriteFileSystemPicker_Input(Application("AZ_ID"), _
																					 IIF(rs(desCampoTipo) = adVarBinary, FILE_SYSTEM_DIRECTORY, FILE_SYSTEM_FILE), _
																					 "images", _
																					 "", _
																					 FormName, _
																					 IIF(	Application("LINGUE")(i) = LINGUA_ITALIANO, _
																							"des"& rs(desCampoTipo) &"_"& rs(desCampoID), _
																							"deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i) _
																						), _
																					 IIF(Application("LINGUE")(i) = LINGUA_ITALIANO, _
																						 IIF(inInserimento, _
																							 request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID)), _
																							 rs(relCampoValore & postfisso_lingua) _
																						 ), _
																						 IIF(inInserimento, _
																							 request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)), _
																							 rs(relCampoValore & Application("LINGUE")(i)) _
																						 ) _
																						), _
																					 "width:436px;", _
																					 FALSE, _
																					 FALSE)
														'if Application("LINGUE")(i) = LINGUA_ITALIANO then
														'	CALL WriteFilePicker_Input(Application("AZ_ID"), "images", FormName, "des"& rs(desCampoTipo) &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), "width:436px;", FALSE) 
														'else
														'	CALL WriteFilePicker_Input(Application("AZ_ID"), "images", FormName, "deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i), IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)), rs(relCampoValore & Application("LINGUE")(i))), "width:436px;", FALSE) 
														'end if
													else %>
														<input type="text" style="width: 95%;" class="text disabled" disabled name="<%= "deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i) %>" value="<%= IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)), rs(relCampoValore & Application("LINGUE")(i))) %>">
												<%	end if %>
												</td>
											</tr>
										<% next %>
									</table>
						<% 		end if
							CASE adDate
								if disabled = "" then
									CALL WriteDataPicker_Input(FormName, "des"& adDate &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& adDate &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), "", "/", TRUE, TRUE, Session("LINGUA"))
								else %>
									<input type="text" class="text disabled" disabled name="<%= "des"& adDate &"_"& rs(desCampoID) %>" value="<%= IIF(inInserimento, request.form("des"& adDate &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)) %>">
						<%		end if
							CASE adUserDefined
								if UseSingleLanguage then
									if disabled = "" then
										CALL WriteLinkBox("des"& adUserDefined &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID)), rs(relCampoValore)), FormName)
									else %>
										<input type="text" style="width: 95%;" class="text disabled" disabled name="<%= "des"& adUserDefined &"_"& rs(desCampoID) %>" value="<%= IIF(inInserimento, request.form("des"& adUserDefined &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)) %>">
						<%			end if
								else %>
									<table cellpadding="0" cellspacing="0" width="100%">
										<%for i = LBound(Application("LINGUE")) to UBound(Application("LINGUE")) %>
											<tr>
												<% if not only_italian then %>
													<td class="content" width="1%">
														<img style="vertical-align:top;" src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" alt="" border="0">&nbsp;
													</td>
												<% end if %>
												<td class="content" <%=IIF(only_italian, "colspan=""2"" style=""width:100%;""", "style=""width:99%;""") %>>
													<% 	if disabled = "" then
															if Application("LINGUE")(i) = LINGUA_ITALIANO then
																CALL WriteLinkBox("des"& rs(desCampoTipo) &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& adUserDefined &"_"& rs(desCampoID)), rs(relCampoValore & Application("LINGUE")(i))), FormName)
															else
																CALL WriteLinkBox("deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i), IIF(inInserimento, request.form("deL"& adUserDefined &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)), rs(relCampoValore & Application("LINGUE")(i))), FormName)
															end if
														else %>
														<input type="text" style="width: 95%;" class="text disabled" disabled name="<%= "deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i) %>" value="<%= IIF(inInserimento, request.form("deL"& adUserDefined &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)), rs(relCampoValore & Application("LINGUE")(i))) %>">
													<%	end if %>
												</td>
											</tr>
										<% next %>
									</table>
								<%end if
							CASE adGUID
								CALL DropDownPages(NULL, "form1", IIF(disabled = "", "505", "505"" "& disabled &" title="""), 0, _
												   "des"& rs(desCampoTipo) &"_"& rs(desCampoID), _
												   IIF(inInserimento, request.form("des"& adGUID &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), _
												   false, false)
							CASE adIUnknown
									if cIntero(desTabella)=0 then
										'gestioen standard: ricerca le rubriche sincronizzate e generate dai descrittori speciali
										sql = "SELECT id_rubrica FROM tb_rubriche WHERE syncroFilterTable LIKE '" & ParseSql(desTabella, adChar) & "' AND syncroFilterKey=" & cIntero(rs(desCampoID))
										Session("form1_des"& adIUnknown &"_"& rs(desCampoID) & "_contatti_rubriche") = GetValueList(conn, NULL, sql)
									else
										'gestione speciale: filtra direttamente per l'id indicato nel nome tabella: 
										'viene usato nel next-gallery per permettere la selezione solo dei contatti indicati in una specifica rubrica.
										Session("form1_des"& adIUnknown &"_"& rs(desCampoID) & "_contatti_rubriche") = desTabella
									end if
									CALL WriteContactPicker_Input(connContent, NULL, "", "", "form1", "des"& adIUnknown &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& adIUnknown &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), "", true, false, disabled <> "", "")
									
							CASE adDouble %>
									min <input type="text" <%= DisableClass(disabled<>"", "text") %> size="5" name="<%= "des"& adDouble &"_"& rs(desCampoID) %>" value="<%= IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)) %>"><%= unita_misura %>
									&nbsp;&nbsp;&nbsp;
									max <input type="text" <%= DisableClass(disabled<>"", "text") %> size="5" name="<%= "d2"& adDouble &"_"& rs(desCampoID) %>" value="<%= IIF(inInserimento, request.form("d2"& rs(desCampoTipo) &"_"& rs(desCampoID)), rs(relCampoValoreLungo & postfisso_lingua)) %>"><%= unita_misura %>
						<%	CASE adLongVarChar, adVarChar
								if UseSingleLanguage then
									Select case rs(desCampoTipo)
										case adLongVarChar
											'testo lungo
											%>
											<textarea style="width:95%;" <%= DisableClass(disabled<>"", "") %> rows="<%=rowsTextarea%>" name="<%= "des"& adLongVarChar &"_"& rs(desCampoID) %>"><%= IIF(inInserimento, request.form("des"& adLongVarChar &"_"& rs(desCampoID)), rs(relCampoValoreLungo)) %></textarea>
										<%case else 
											'campo di testo normale
											%>
											<input <%= DisableClass(disabled<>"", "text") %> maxlength="250" style="width:94%;" type="Text" name="<%= "des"& rs(desCampoTipo) &"_"& rs(desCampoID) %>" id="<%= "des"& rs(desCampoTipo) &"_"& rs(desCampoID) %>" value="<%= left(cString(IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID)), rs(relCampoValore))), 250) %>"><%= unita_misura %>
											<% CALL AutoCompleteHTML("des"& rs(desCampoTipo) &"_"& rs(desCampoID), postfisso_lingua)
										end select
								else%>
									<table cellpadding="0" cellspacing="0" width="100%">
										<%for i = LBound(Application("LINGUE")) to UBound(Application("LINGUE")) %>
											<tr>
												<% if not only_italian then %>
													<td class="content" width="1%">
														<img style="vertical-align:top;" src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" alt="" border="0">&nbsp;
													</td>
												<% end if %>
												<td class="content" <%=IIF(only_italian, "colspan=""2"" style=""width:100%;""", "style=""width:99%;""") %>>
													<% Select case rs(desCampoTipo)
														case adLongVarChar
															'testo lungo
															if Application("LINGUE")(i) = LINGUA_ITALIANO then %>
																<textarea style="width:95%;" <%= DisableClass(disabled<>"", "") %> rows="<%=rowsTextarea%>" name="<%= "des"& adLongVarChar &"_"& rs(desCampoID) %>"><%= IIF(inInserimento, request.form("des"& adLongVarChar &"_"& rs(desCampoID)), rs(relCampoValoreLungo & Application("LINGUE")(i))) %></textarea>
															<% else %>
																<textarea style="width:95%;" <%= DisableClass(disabled<>"", "") %> rows="<%=rowsTextarea%>" name="<%= "deL"& adLongVarChar &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i) %>"><%= cString(IIF(inInserimento, request.form("deL"& adLongVarChar &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)), rs(relCampoValoreLungo & Application("LINGUE")(i)))) %></textarea>
															<% end if
														case else 
															'campo di testo normale o link esterno
															%>
															<% if Application("LINGUE")(i) = LINGUA_ITALIANO then %>
																<input <%= DisableClass(disabled<>"", "text") %> maxlength="250" style="width:94%;" type="Text" name="<%= "des"& rs(desCampoTipo) &"_"& rs(desCampoID) %>" id="<%= "des"& rs(desCampoTipo) &"_"& rs(desCampoID) %>" value="<%= IIF(inInserimento, request.form("des"& rs(desCampoTipo) &"_"& rs(desCampoID)), rs(relCampoValore & Application("LINGUE")(i))) %>"><%= unita_misura %>
																<% CALL AutoCompleteHTML("des"& rs(desCampoTipo) &"_"& rs(desCampoID), Application("LINGUE")(i))
															else %>
																<input <%= DisableClass(disabled<>"", "text") %> maxlength="250" style="width:94%;" type="Text" id="<%= "deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i) %>" name="<%= "deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i) %>" value="<%= IIF(inInserimento, request.form("deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)), rs(relCampoValore & Application("LINGUE")(i))) %>"><%= unita_misura %>
																<% CALL AutoCompleteHTML("deL"& rs(desCampoTipo) &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i), Application("LINGUE")(i))
															end if
													end select %>
												</td>
											</tr>
										<% next %>
									</table>
								<% end if
							case adPropVariant 
								'colore in HTML
								CALL WriteColorPicker_Input_Disable(FormName, "des"& adPropVariant &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& adPropVariant &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), IIF(disabled <> "", """ disabled title=""", ""), false, false, "", disabled <> "")
							case adChapter
								'voce dell'indice
								CALL index.WritePicker("", "", FormName, "des"& adChapter &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& adChapter &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), 0, false, false, 91, disabled<>"", false)
							case adIDispatch
								'rubrica
								sql = "SELECT * FROM tb_rubriche WHERE id_rubrica IN ("& GetList_Rubriche(conn, NULL) &")"
								CALL DropDown(connContent, sql, "id_rubrica", "nome_rubrica", "des"& adIDispatch &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& adIDispatch &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), false, "", LINGUA_ITALIANO)
							case adSingle
								'amministratore
								sql = "SELECT *, admin_cognome "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) &" admin_nome AS NOME FROM tb_admin"
								CALL DropDown(connContent, sql, "id_admin", "NOME", "des"& adSingle &"_"& rs(desCampoID), IIF(inInserimento, request.form("des"& adSingle &"_"& rs(desCampoID)), rs(relCampoValore & postfisso_lingua)), false, "", LINGUA_ITALIANO)
							case adWChar
								if UseSingleLanguage then
									%>
									<textarea style="width:95%;" <%= DisableClass(disabled<>"", "") %> rows="<%=rowsTextarea%>" name="<%= "des"& adWChar &"_"& rs(desCampoID) %>"><%= IIF(inInserimento, request.form("des"& adWChar &"_"& rs(desCampoID)), rs(relCampoValoreLungo)) %></textarea>
									<%  CALL activateCKEditor("des"& adWChar &"_"& rs(desCampoID)) 
								else%>
									<table cellpadding="0" cellspacing="0" width="100%">
										<%for i = LBound(Application("LINGUE")) to UBound(Application("LINGUE")) %>
											<tr>
												<% if not only_italian then %>
													<td class="content" width="1%">
														<img style="vertical-align:top;" src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" alt="" border="0">&nbsp;
													</td>
												<% end if %>
												<td class="content" <%=IIF(only_italian, "colspan=""2"" style=""width:100%;""", "style=""width:99%;""") %>>
													<% if Application("LINGUE")(i) = LINGUA_ITALIANO then %>
														<textarea style="width:95%;" <%= DisableClass(disabled<>"", "") %> rows="<%=rowsTextarea%>" name="<%= "des"& adWChar &"_"& rs(desCampoID) %>"><%= IIF(inInserimento, request.form("des"& adWChar &"_"& rs(desCampoID)), MakeAbsoluteLink(rs(relCampoValoreLungo & Application("LINGUE")(i)))) %></textarea>
														<% CALL activateCKEditor("des"& adWChar &"_"& rs(desCampoID)) %>
													<% else %>
														<textarea style="width:95%;" <%= DisableClass(disabled<>"", "") %> rows="<%=rowsTextarea%>" name="<%= "deL"& adWChar &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i) %>"><%= cString(IIF(inInserimento, request.form("deL"& adWChar &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)), rs(relCampoValoreLungo & Application("LINGUE")(i)))) %></textarea>
														<% CALL activateCKEditor("deL"& adWChar &"_"& rs(desCampoID) &"_"& Application("LINGUE")(i)) %>
													<% end if %>
												</td>
											</tr>
										<% next %>
									</table>
								<% end if
						end select %>
					</td>
				</tr>
				<% end if %>
				
				<% if cSTring(desCampoNote)<>"" then
					if cString(rs(desCampoNote))<>"" then %> 
						<tr>
							<td class="content note"><%= rs(desCampoNote) %></td>
						</tr>
					<%end if
				end if
			
			end if
			
			
			rs.movenext
		wend
	end if
	
	if almeno_un_valore then
		%>
		<tr>
			<th class="warning" colspan="<%= colspan %>">ATTENZIONE! Le caratteristiche evidenziate non saranno visibili perchè non sono collegate alla categoria attuale.</th>
		</tr>
		<%
	end if
	
	set rs = nothing
End Sub

'salva nel DB la relazione gestendo i dati del form
'gestisce gli errori di tipo con la transazione e impostando session("ERRORE")
Sub DesSalva(conn, ID, relTab, relCampoValore, relCampoExt, relCampoDes)
	CALL DesSalva_EXT(conn, ID, relTab, relCampoValore, relCampoValore, relCampoExt, relCampoDes)
end sub


Sub DesSalva_EXT(conn, ID, relTab, relCampoValore, relCampoValoreLungo, relCampoExt, relCampoDes)
	CALL DesSave(conn, ID, relTab, relCampoValore, relCampoValoreLungo, relCampoExt, relCampoDes, "")
end sub

sub DesSave(conn, ID, relTab, relCampoValore, relCampoValoreLungo, relCampoExt, relCampoDes, filtroDelete)
	dim campo, tipo, dID, desVuoto, sql, i, postfisso_lingua, doubleVal
	if UseSingleLanguage then
		postfisso_lingua = ""
	else
		postfisso_lingua = "IT"
	end if
	
	if request("ID") <> "" then						'se sono in modifica cancello prima le relazioni esistenti
		sql = "DELETE FROM "& relTab &" WHERE "& relCampoExt &"="& cIntero(request("ID")) &" "& filtroDelete
		conn.execute(sql)
	end if
	
	for each campo in request.form
		if lcase(Left(campo, 3)) = "des" then
			tipo = Right(campo, Len(campo)-3)
			tipo = Left(tipo, Instr(tipo, "_")-1)
			dID = right(campo, len(campo)-instr(campo, "_"))
			doubleVal = request.form("d2"& tipo &"_"& dID)
			
			'controllo che i campi non siano vuoti
			desVuoto = TRUE
			if request.form(campo) = "" AND doubleVal = "" AND not UseSingleLanguage then		'se il campo IT Ã¨ vuoto controllo le altre lingue
				for i = LBound(Application("LINGUE")) to UBound(Application("LINGUE"))
					if request.form("deL"& tipo &"_"& dID &"_"& Application("LINGUE")(i)) <> "" _
					   OR request.form("d2"& tipo &"_"& dID &"_"& Application("LINGUE")(i)) <> "" then
						desVuoto = FALSE
						exit for
					end if
				next
			elseif request.form(campo) <> "" OR doubleVal <> "" then
				desVuoto = FALSE
			end if
			
			if NOT desVuoto then
				'controllo sul tipo
				SELECT CASE CInt(tipo)
				CASE adNumeric, adCurrency, adIDispatch, adSingle
					if NOT IsNumeric(request.form(campo)) then
						conn.RollBackTrans
						Session("errore") = "Errore dei dati nel campo numerico!"
						exit for
					end if
				CASE adDouble
					if NOT IsNumeric(request.form(campo)) AND NOT IsNumeric(doubleVal) then
						conn.RollBackTrans
						Session("errore") = "Errore dei dati nel campo "& tipiDescrittori(10, 1) &"!"
						exit for
					end if
				END SELECT
				
				sql = "INSERT INTO "& relTab &"("
				if UseSingleLanguage then
					sql = sql & IIF(CInt(tipo) = adLongVarChar or CInt(tipo) = adWChar, relCampoValoreLungo, relCampoValore) & ", "
				else
					for i = LBound(Application("LINGUE")) to UBound(Application("LINGUE"))
						if doubleVal <> "" then
							sql = sql & relCampoValore & Application("LINGUE")(i) &", "& relCampoValoreLungo & Application("LINGUE")(i) &", "
						else
							sql = sql & IIF(CInt(tipo) = adLongVarChar or CInt(tipo) = adWChar, relCampoValoreLungo, relCampoValore) & Application("LINGUE")(i) &", "
						end if
					next
				end if
				
				sql = sql & relCampoExt &", "& relCampoDes &") VALUES ("
				
				dim valore
				if UseSingleLanguage then
					sql = sql &"'"& ParseSQL(request.form(campo), adChar) &"', "
				else
					for i = LBound(Application("LINGUE")) to UBound(Application("LINGUE"))
						if UCase(Application("LINGUE")(i)) = "IT" then
							valore = ParseSQL(request.form(campo), adChar)
							if CInt(tipo) = adWChar then  'in caso di descrittore HTML rendo i link relativi 
								valore = MakeRelativeLink(valore)
							end if
							sql = sql &"'"& valore &"', "
							if doubleVal <> "" then
								sql = sql &"'"& ParseSQL(doubleVal, adChar) &"', "
							end if
						else
							valore = ParseSQL(request.form("deL"& tipo &"_"& dID &"_"& Application("LINGUE")(i)), adChar)
							if CInt(tipo) = adWChar then 'in caso di descrittore HTML rendo i link relativi 
								valore = MakeRelativeLink(valore)
							end if
							sql = sql&IIF(DB_Type(conn)=DB_SQL," N'"," '")& valore &"', "
							if doubleVal <> "" then
								sql = sql&IIF(DB_Type(conn)=DB_SQL," N'"," '")& ParseSQL(request.form("d2"& tipo &"_"& dID &"_"& Application("LINGUE")(i)), adChar) &"', "
							end if
						end if
					next
				end if
				sql = sql & ID &", "& dID &")"
'response.write sql
'response.end
				conn.execute(sql)
			end if
			
		end if
	next
End Sub


'scrive le righe per il box della ricerca con gli input per i descrittori
'prefisso:		prefisso per i nomi degli input (es.: "adv_im")
Sub DesRicerca(conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, prefisso)
	CALL DesRicercaEX(conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, prefisso, prefisso)
end sub


Sub DesRicercaEX(conn, sql, desTabella, desCampoID, desCampoNome, desCampoTipo, desCampoUnitaM, prefissoForm, prefissoSessione)
	dim rsd, unita_misura
	set rsd = conn.execute(sql)
	if not rsd.eof then %>
		<% while not rsd.eof
				unita_misura = ""
				if desCampoUnitaM <> "" then
					if cString(rsd(desCampoUnitaM))<>"" then
						unita_misura = "&nbsp;" & rsd(desCampoUnitaM)
					end if
				end if%>
			<tr><td class="label"><%= rsd(desCampoNome) %>:</td></tr>
			<tr>
				<td class="content">
					<%	SELECT CASE rsd(desCampoTipo)
							CASE adDate
								CALL WriteDataPicker_Input("form1", prefissoForm &"_descr"& adDate &"_"& rsd(desCampoID), Session(prefissoSessione &"_descr"& adDate &"_"& rsd(desCampoID)), "", "/", true, true, LINGUA_ITALIANO)
							CASE adNumeric %>
							<input class="text" maxlength="50" size="10" type="Text" name="<%= prefissoForm &"_descr"& adNumeric &"_"& rsd(desCampoID) %>" value="<%= Session(prefissoSessione &"_descr"& adNumeric &"_"& rsd(desCampoID)) %>"><%= unita_misura %>
							<% CALL AutoCompleteHTML(prefissoForm &"_descr"& adNumeric &"_"& rsd(desCampoID), "")
						CASE adCurrency
								if unita_misura = "" then
									unita_misura = "&nbsp;&euro;"
								end if %>
							<input class="text" maxlength="50" size="7" type="Text" name="<%= prefissoForm &"_descr"& adCurrency &"_"& rsd(desCampoID) %>" value="<%= IIF(Session(prefissoSessione &"_descr"& adCurrency &"_"& rsd(desCampoID)) = "", "", FormatPrice(Session(prefissoSessione &"_descr"& adCurrency &"_"& rsd(desCampoID)), 2, TRUE)) %>"><%= unita_misura %>&nbsp;&nbsp;&nbsp;(es.: 100,00)
						<% 	CASE adBoolean %>
							Si<input type="checkbox" class="checkbox" name="<%= prefissoForm &"_descr"& adBoolean &"_"& rsd(desCampoID) %>" value="1" <%= Chk(InStr(Session(prefissoSessione &"_descr"& adBoolean &"_"& rsd(desCampoID)), "1") > 0) %>>
							No<input type="checkbox" class="checkbox" name="<%= prefissoForm &"_descr"& adBoolean &"_"& rsd(desCampoID) %>" value="0" <%= Chk(InStr(Session(prefissoSessione &"_descr"& adBoolean &"_"& rsd(desCampoID)), "0") > 0) %>>
						<% 	CASE adLongVarChar %>
							<textarea style="width:100%;" rows="5" name="<%= prefissoForm &"_descr"& adVarChar &"_"& rsd(desCampoID) %>"><%= Session(prefissoSessione &"_descr"& adVarChar &"_"& rsd(desCampoID)) %></textarea>
						<% 	CASE adIUnknown%>
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td>
                                        <% sql = "syncroFilterTable LIKE '" & desTabella & "'"
                                        CALL WriteContactPicker_Input(conn, NULL, "", sql, "form1", prefissoForm &"_descr"& adIUnknown &"_"& rsd(desCampoID), Session(prefissoSessione &"_descr"& adIUnknown &"_"& rsd(desCampoID)), "", true, false, false, "") %>
                                    </td>
                                </tr>
								<tr>
									<td>
										<table width="100%" cellpadding="0" cellspacing="0">
											<tr>
												<td width="5%">
													<input class="checkbox" type="radio" name="<%= prefissoForm &"_descr"& adIUnknown &"_"& rsd(desCampoID) %>_ANDOR" value="AND" <%= chk(Session(prefissoSessione &"_descr"& adIUnknown &"_"& rsd(desCampoID) &"_ANDOR")="AND" OR Session(prefissoForm &"_descr"& adIUnknown &"_"& rsd(desCampoID) &"_ANDOR") = "") %>>
												</td>
												<td class="content">associato a tutte le anagrafiche selezionate</td>
											</tr>
											<tr>
												<td>
													<input class="checkbox" type="radio" name="<%= prefissoForm &"_descr"& adIUnknown &"_"& rsd(desCampoID) %>_ANDOR" value="OR" <%= chk(Session(prefissoSessione &"_descr"& adIUnknown &"_"& rsd(desCampoID) &"_ANDOR")="OR") %>>
												</td>
												<td class="content">associato ad almeno una anagrafica selezionata</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						<%	CASE adGUID
								CALL DropDownPages(NULL, "form1", "300", Application("AZ_ID"), _
												   prefissoForm &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID), _
												   Session(prefissoSessione &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID)), _
												   false, false)
							CASE adIDispatch
								'rubrica
								sql = "SELECT * FROM tb_rubriche WHERE id_rubrica IN ("& GetList_Rubriche(conn, NULL) &")"
								CALL DropDown(conn, sql, "id_rubrica", "nome_rubrica", prefissoForm &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID), Session(prefissoSessione &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID)), true, "", LINGUA_ITALIANO)
							CASE adIDispatch
								'amministratore
								sql = "SELECT *, admin_cognome "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) &" admin_nome AS NOME FROM tb_admin"
								CALL DropDown(conn, sql, "id_admin", "NOME", prefissoForm &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID), Session(prefissoSessione &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID)), true, "", LINGUA_ITALIANO)
						 	CASE ELSE %>
							<input class="text" maxlength="255" style="width:100%;" type="Text" name="<%= prefissoForm &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID) %>" value="<%= Session(prefissoSessione &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID)) %>">
							<% CALL AutoCompleteHTML(prefissoForm &"_descr"& rsd(desCampoTipo) &"_"& rsd(desCampoID), "") %>
					<% 	END SELECT %>
				</td>
			</tr>
			<%rsd.movenext
		wend
	end if
	set rsd = nothing
End Sub

'imposta la query di ricerca ed il testo
'campoID:		nome del campo ID della tabella esterna
'relCampoID:	nome del campo ID del descrittore nella relazione con la tabella esterna
Sub DesRicercaQuery(ByRef sql, ByRef testo, desTabella, desCampoID, desCampoNome, desCampoUnitaM, campoID, relTab, relCampoID, relCampoExt, relCampoValore, prefisso)
dim var, id, tipo, ind, vettore, lingua, unita_misura, aux, postfisso_lingua
	prefisso = prefisso &"_descr"
	if UseSingleLanguage then
		postfisso_lingua = ""
	else
		postfisso_lingua = "IT"
	end if
	for each var in session.contents
		'response.write left(var, len(prefisso)) & "<br>"
		if left(var, len(prefisso)) = prefisso AND InStr(var, "visContatti") = 0 AND InStr(var, "ANDOR") = 0 AND IsArray(session(var)) then 'session(var) <> "" dava errore su memo 2 di bibliotecasanfr. 10/09/2014
			id = right(var, len(var) - InStrRev(var, "_", len(var)))
			tipo = right(var, len(var) - len(prefisso))
			tipo = cIntero(left(tipo, len(tipo) - len(id) -1))
			
			SELECT CASE tipo
				CASE adBoolean
					if len(session(var)) = 1 then		'non ho selezionato sia si che no
						sql = sql &" AND "& campoID &" IN ( SELECT "& relCampoExt & _
														  " FROM "& relTab &" WHERE "& relCampoID &"="& cIntero(id) &" AND "& relCampoValore & postfisso_lingua &" LIKE '"& session(var) &"')"
						if session(var) = "1" then
							testo = testo & "<tr><td class=""label"">" & GetValueList(conn, NULL, "SELECT "& desCampoNome &" FROM "& desTabella &" WHERE "& desCampoID &"="& cIntero(ID)) &":</td></tr><tr><td class=""content_right"">Si</td></tr>"
						else
							testo = testo & "<tr><td class=""label"">" & GetValueList(conn, NULL, "SELECT "& desCampoNome &" FROM "& desTabella &" WHERE "& desCampoID &"="& cIntero(ID)) &":</td></tr><tr><td class=""content_right"">No</td></tr>"
						end if
					end if
				CASE adIUnknown
					vettore = replace(session(var), " ", "")
					vettore = left(vettore, len(vettore)-1)
					sql = sql &" AND "& campoID & _
							   " IN ( SELECT "& relCampoExt & _
							   " FROM "& relTab &" WHERE "& relCampoID &"="& cIntero(id) &" AND ("
					for each ind in split(vettore, ";")
						sql = sql &" "& relCampoValore & postfisso_lingua &" LIKE '% "& ParseSql(ind, adChar) &";%' "& Session(var &"_ANDOR")
					next
					sql = left(sql, len(sql) - len(Session(var &"_ANDOR"))) &"))"
					testo = testo & "<tr><td class=""label"">" & GetValueList(conn, NULL, "SELECT "& desCampoNome &" FROM "& desTabella &" WHERE "& desCampoID &"="& cIntero(ID)) &":</td></tr><tr><td class=""content_right"">"& Session(var &"_visContatti") &"</td></tr>"
				CASE adNumeric, adCurrency
					sql = sql &" AND "& campoID &" IN ( SELECT "& relCampoExt & _
													  " FROM "& relTab &" WHERE "& relCampoID &"="& cIntero(id) &" AND "& relCampoValore & postfisso_lingua &" LIKE '"& session(var) &"')"
					unita_misura = ""
					set aux = conn.execute("SELECT * FROM "& desTabella &" WHERE "& desCampoID &"="& cIntero(ID))
					if desCampoUnitaM <> "" then
						if cString(aux(desCampoUnitaM))<>"" then
							unita_misura = "&nbsp;" & aux(desCampoUnitaM)
						end if
					end if
					if tipo = adCurrency AND unita_misura = "" then
						unita_misura = "&nbsp;&euro;"
					end if
					testo = testo & "<tr><td class=""label"">" & aux(desCampoNome) &":</td></tr><tr><td class=""content_right"">"& Session(var) &" "& unita_misura &"</td></tr>"
				CASE adDate
					sql = sql &" AND "& campoID &" IN ( SELECT "& relCampoExt & _
													  " FROM "& relTab &" WHERE "& relCampoID &"="& cIntero(id) &" AND ("
					sql = sql & relCampoValore & postfisso_lingua &" LIKE '%"& session(var) &"%'))"
					testo = testo & "<tr><td class=""label"">" & GetValueList(conn, NULL, "SELECT "& desCampoNome &" FROM "& desTabella &" WHERE "& desCampoID &"="& cIntero(ID)) &":</td></tr><tr><td class=""content_right"">"& Session(var) &"</td></tr>"
				CASE adDouble
					session(var) = CIntero(session(var))
					sql = sql &" AND "& campoID &" IN ( SELECT "& relCampoExt & _
													  " FROM "& relTab &" WHERE "& relCampoID &"="& cIntero(id) & _
													  " AND "& SQL_If(conn, SQL_IsTrue(conn, "ISNUMERIC("& relCampoValore &")"), _
																	  "("& SQL_Numeric(conn, relCampoValore) &" <= "& session(var), "(1=1)") & _
													  " AND "& SQL_If(conn, SQL_IsTrue(conn, "ISNUMERIC("& relCampoValoreLungo &")"), _
																	  "("& SQL_Numeric(conn, relCampoValoreLungo) &" >= "& session(var), "(1=1)") & _
													  ")"
					
					unita_misura = ""
					set aux = conn.execute("SELECT * FROM "& desTabella &" WHERE "& desCampoID &"="& cIntero(ID))
					if desCampoUnitaM <> "" then
						if cString(aux(desCampoUnitaM))<>"" then
							unita_misura = "&nbsp;" & aux(desCampoUnitaM)
						end if
					end if
					if tipo = adCurrency AND unita_misura = "" then
						unita_misura = "&nbsp;&euro;"
					end if
					testo = testo & "<tr><td class=""label"">" & aux(desCampoNome) &":</td></tr><tr><td class=""content_right"">"& Session(var) &" "& unita_misura &"</td></tr>"
				CASE ELSE
					sql = sql &" AND "& campoID &" IN ( SELECT "& relCampoExt & _
													  " FROM "& relTab &" WHERE "& relCampoID &"="& cIntero(id) &" AND ("
					if UseSingleLanguage then
						sql = sql & relCampoValore &" LIKE '%"& session(var) &"%'))"
					else
						sql = sql &" 1=0 "
						for each lingua in application("LINGUE")
							sql = sql &" OR "& relCampoValore & lingua &" LIKE '%"& session(var) &"%'"
						next
						sql = sql &"))"
					end if
					testo = testo & "<tr><td class=""label"">" & GetValueList(conn, NULL, "SELECT "& desCampoNome &" FROM "& desTabella &" WHERE "& desCampoID &"="& cIntero(ID)) &":</td></tr><tr><td class=""content_right"">"& Session(var) &"</td></tr>"
					
			END SELECT
		end if
	next
End Sub


'formatta value in base al tipo
Function DesFormat(tipo, value, stile_link, max_length, unita)
	SELECT CASE tipo
		CASE adBinary
			value = CString(value)
			DesFormat = "<a "& stile_link &" href=""http://"& Application("IMAGE_SERVER") &"/"& Application("AZ_ID") &"/images/"& value &""">"& right(value, len(value)-InStrRev(value, "/")) &"</a>"
		CASE adCurrency
			DesFormat = FormatPrice(value, 2, true) & IIF(cString(unita)<>"", " " & unita, "")
		CASE adBoolean
			value = CIntero(value)
			DesFormat = IIF(value, ChooseValueByAllLanguages(session("LINGUA"), "S&igrave;", "Yes", "Ja", "Oui", "S&iacute;", "быть", "是", "Sì"), ChooseValueByAllLanguages(session("LINGUA"), "No", "No", "Nein", "No", "No", "нет", "否", "Não"))
		CASE adUserDefined
			value = cString(value)
			DesFormat = "<a "& stile_link &" href="""
			if instr(1, value, "@", vbTextCompare)>0 then
				DesFormat = DesFormat + "mailto:" & value
			elseif instr(1, value, "http", vbTextCompare)<1 then
				DesFormat = DesFormat + "http://" & value & """  target=""_blank"
			else
				DesFormat = DesFormat + value
			end if
			DesFormat = DesFormat + """>"& value &"</a>"
		CASE adGUID
			dim connL, rsL
			set connL = server.createobject("adodb.connection")
			connL.open Application("L_conn_ConnectionString")
			set rsL = server.createobject("adodb.recordset")
            
			DesFormat = "<a "& stile_link &" href=""" & GetPageSiteUrl(connL, value, Session("LINGUA")) & _
						""">"& GetValueList(connL, rsL, "SELECT nome_ps_"& session("LINGUA") &" FROM tb_pagineSito WHERE id_pagineSito="& CIntero(value)) &"</a>"

			set rsL = nothing
			connL.close
			set connL = nothing
		CASE adIUnknown
			DesFormat = ConvertiID(conn, value, InStr(Request.ServerVariables("NEXTcom"), "amministrazione") > 0)
		case adChapter
			'recupera valore dell'indice
			if cInteger(value)>0 then
				DesFormat = index.NomeCompleto(value)
			else
				DesFormat = ""
			end if
		CASE adIDispatch
			'rubrica
			DesFormat = GetValueList(conn, NULL, "SELECT nome_rubrica FROM tb_rubriche WHERE id_rubrica = "& CIntero(value))
			if cIntero(max_length) > 0 then
				DesFormat = Sintesi(DesFormat, cIntero(max_length), "...")
			end if
		CASE adSingle
			'amministratore
			DesFormat = GetValueList(conn, NULL, "SELECT admin_cognome "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) &" admin_nome AS NOME FROM tb_admin WHERE id_admin = "& CIntero(value))
			if cIntero(max_length) > 0 then
				DesFormat = Sintesi(DesFormat, cIntero(max_length), "...")
			end if
		CASE ELSE
			if cIntero(max_length) = 0 then
				DesFormat = TextEncode(value)
			else
				DesFormat = TextEncode(Sintesi(value, cIntero(max_length), "..."))
			end if
	END SELECT
End Function


'restituisce il valore per il descrittore
Function DesFormatValue(tipo, value_testo, value_memo, stile_link, max_length, unita)
	dim value
	if tipo = adDouble then
		if CString(value_testo) <> "" then
			desFormatValue = "Min. "& value_testo & IIF(cString(unita)<>"", " " & unita, "")
		end if
		if CString(value_memo) <> "" then
			if desFormatValue <> "" then
				desFormatValue = desFormatValue &"&nbsp;&nbsp;&nbsp;"
			end if
			desFormatValue = desFormatValue &"Max. "& value_memo & IIF(cString(unita)<>"", " " & unita, "")
		end if
	else
		if tipo = adLongVarChar or tipo = adWChar then
			value = value_memo
		else
			value = value_testo
		end if
		DesFormatValue = DesFormat(tipo, value, stile_link, max_length, unita)
	end if
end function


'restituisce lo stile da applicare al valore
function DesStyleValue(tipo)
	dim i
	for i = lbound(TipiDescrittori) to Ubound(TipiDescrittori)
		if TipiDescrittori(i,0) = tipo then
			DesStyleValue = TipiDescrittori(i, 2)
		end if
	next
end function


'restituisce il valore del descrittore tipizzato, non esegue il CBL
Function DesValue(rs, nomeCampoTipo, prefissoRelazione, lingua)
	dim valore, memo
	valore = rs(prefissoRelazione &"valore_"& lingua)
	memo = rs(prefissoRelazione &"memo_"& lingua)
	
	DesValue = DesValore(rs(nomeCampoTipo), valore, memo)
End Function
'.................................................................................................
'..			FINE GESTIONE DESCRITTORI
'.................................................................................................


'.................................................................................................
'funzione che converte gli ID del tipo anagrafiche in nomi
'richiede che sia incluso anche Tools_contatti
'link: 			se true restiuisce il nome con il link alla scheda
'.................................................................................................
Function ConvertiID(conn, stringa, link)
dim aux
	'converto gli ID in nomi
	if cString(stringa) <> "" then
		set aux = conn.execute(" SELECT IDElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, CognomeElencoIndirizzi, NomeElencoIndirizzi, isSocieta " & _
							   " FROM tb_indirizzario WHERE IDElencoIndirizzi IN ("& replace(left(stringa, len(stringa)-1), ";", ",") &")")
		while not aux.eof
			if link then
				ContactLinkedName(aux)
				response.write "; "
			else
				if aux("isSocieta") then
					ConvertiID = ConvertiID &" "& aux("NomeOrganizzazioneElencoIndirizzi") &";"
				else
					ConvertiID = ConvertiID &" "& aux("CognomeElencoIndirizzi") &" "& aux("NomeElencoIndirizzi") &";"
				end if
			end if
			aux.movenext
		wend
		ConvertiID = replace(ConvertiID, "^", "")
		ConvertiID = replace(ConvertiID, "'", "")
		ConvertiID = replace(ConvertiID, """", "")
		ConvertiID = replace(ConvertiID, "(", "")
		ConvertiID = replace(ConvertiID, ")", "")
	end if
End Function


'.................................................................................................
'funzione che disegna un box di testo con pulsante per aprire il link inserito
'.................................................................................................
sub WriteLinkBox(FieldName, FieldValue, FormName) %>
	<script language="JavaScript" type="text/javascript">
	<!--
		function <%= FieldName %>_onClick(){
			var value = <%= FormName %>.<%= FieldName %>.value;
			if (value != ""){
				OpenWindow(value, '', '');
			}
		}
	//-->
	</script>
	<table cellspacing="0" cellpadding="0" width="95%">
		<tr>
			<td>
				<input class="text" maxlength="250" style="width:100%;" type="Text" name="<%= FieldName %>" value="<%= FieldValue %>">
			</td>
			<td width="2%">	
				<a href="javascript:void(0);" class="button_input" id="<%= FieldName %>_link" onclick="<%= FieldName %>_onClick()" title="Apri link in una nuova pagina" <%= ACTIVE_STATUS %>>APRI</a>
			</td>
		</tr>
	</table>
<%end sub


%>
