<% 

'*******************************************************************************************************************
'CLASSE CHE GESTISCE LE GALLERIE DI IMMAGINI

class ClassPhotoGallery

'******************************************************************************************************************************************
'******************************************************************************************************************************************
'VARIABILI E PROPRIETA' PUBBLICHE
'******************************************************************************************************************************************

	'impostazione dati tabella
	Public TableName							'nome tabella delle foto.
	Public FieldPrefix							'prefisso dei campi della tabella.
	Public FieldForeignKey						'nome del campo chiave esterna.
	Public FilePrefix							'prefisso dei nomi dei files.
	Public DeleteKey							'chiave per l'apertura
	Public FieldRiservata						'nome del campo che decide se la foto è riservata
	
	'descrizione elemento a cui sono associate le immagini (Usa la proprietà per poter impostare correttamente anche la variabile di dichiarazione nell'indice)
	Private elementDatabaseTableName
	
	Public Property Let ElementTableName(value)
		elementDatabaseTableName = value
		if cString(ElementIndexTableName) = "" then
			ElementIndexTableName = elementDatabaseTableName
		end if
	End Property
	
	Public Property Get ElementTableName(  )
		ElementTableName = elementDatabaseTableName
	End Property

	Public ElementIndexTableName				'Nome della tabella "padre" registrato nell'indice.
	Public ElementFieldPrefix					'prefisso dei campi
	Public ElementUpdateParams					'se true va ad aggiornare date di modifica dell'elemento.
	
	'tipizzazione delle foto
	Public TipiFotoTableName					'nome della tabella di tipizzazione delle foto.
	
	'consigli sui formati delle immagini: visualizzati nei form di inserimento/modifica come consiglio/nota per i formati
	Public NoteFormatoThumbnail
	Public NoteFormatoZoom
	
	'impostazione vincoli di inserimento
	Public MaxFotoPerRecord						'numero massimo di foto per record.
	
	'Impostazioni di comportamento
	Public MandatoryThumb						'se true imposta il campo thumbnail obbligatorio
	Public MandatoryZoom						'se true imposta il campo zoom obbligatorio
	Public FotoSingola							'se true visualizza solo una foto (zoom) e non thumb e zoom.
	'proprieta per la modalita di salvataggio dei dati dal file manager
	Public ElementID							'se > 0 indica l'ID del record della tabella "padre" cui salvare l'immagine
	Public ElementThumb							'thumb da inserire
	Public ElementZoom							'zoom da inserire
	
	'abilitazioni dell'esecuzione delle estensioni dichiarate esternamente
	Public Abilita_FormGestioneAddOn			'Abilita l'esecuzione della funzione FormGestione_ADDON(conn) presente nel form (sopra la didascalia) 
												'e dichiarata nel file civetta di gestione
	Public Abilita_ElencoAddOn 					'Abilita l'esecuzione della funzione ElencoTestata_ADDON(conn) ed ElencoRow_ADDON(conn, rs) presente nell'elenco ed eseguita in ogni riga.
												'e dichiarata nel file civetta di gestione

	Public conn									'connessione a database (se non impostata se ne crea una).
	
	
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'VARIABILI E PROPRIETA' PRIVATE
'******************************************************************************************************************************************

	Private oIndex								'oggetto per la gestione dei contenuti e dell'indice generale.
	Private ImageWebId							'id del sito da cui prendere le immagini.
	Private ImageBaseUrl						'url di base della directory immagini indicata.

	
	'......................................................................................................................................
	'Oggetto di gestione dell'indice
	'......................................................................................................................................
	Public Property Get Index()
	    set Index = oIndex
	End Property
	
	Public Property Let Index(obj)
	    set oIndex = obj
	    set conn = oIndex.conn
	End Property
	
	
	'......................................................................................................................................
	'Sito web dove "risiedono" le immagini
	'......................................................................................................................................
	Public Property Get WebId()
		WebId = ImageWebId
	End Property
	
	Public Property Let WebId(id)
	    ImageWebId = id
		ImageBaseUrl = Application("IMAGE_SERVER") & "/" & ImageWebId & "/images/"
		ImageBaseUrl = "http://" & replace(ImageBaseUrl, "//", "/")
	End Property
	
	
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'COSTRUTTORI CLASSE
'******************************************************************************************************************************************
	
    Private Sub Class_Initialize()
	
		'crea connessione di default
		Set conn = Server.CreateObject("ADODB.connection")
		conn.open Application("DATA_ConnectionString")
    	
		'imposta limiti generali
		MaxFotoPerRecord = 0
		
		'Altre impostazioni di default
		DeleteKey = "FOTO"
		FotoSingola = false
		MandatoryThumb = true
		MandatoryZoom = true
		Abilita_ElencoAddOn = false
		Abilita_FormGestioneAddOn = false
		TipiFotoTableName = ""
		
		'impostazione salvataggio immagine da file manager
		ElementID = CIntero(request("RS_ID"))
		ElementThumb = request("RS_THUMB")
		ElementZoom = request("RS_ZOOM")
	End Sub
    
    
    Private Sub Class_Terminate()
        
		
    End Sub
	
	
	
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'METODI PUBBLICI
'******************************************************************************************************************************************
	
	Public Sub Elenco(ForeignKeyId, SectionLabel)
		CALL ElencoExt(ForeignKeyId, SectionLabel, "", false)
	end sub

	'......................................................................................................................................
	'GESTIONE ELENCO FOTO
	'ID:				id del record a cui le foto sono associate
	'SectionLabel		Titolo della sezione di form
	'additionalPath		path della cartella nella quale aggiungere il/i file
	'directUpload		se true, e se "additionalPath" non è vuota, si passa direttamente all'upload dei file senza passare per la scelta della cartella
	'......................................................................................................................................
	Public Sub ElencoExt(ForeignKeyId, SectionLabel, additionalPath, directUpload)
		dim rs, sql, AbilitaNuoveFoto, AbilitaTipoFoto
		AbilitaNuoveFoto = false
		if (TipiFotoTableName <> "") then
			if (GetValueList(conn, NULL, "SELECT COUNT(*) FROM " & TipiFotoTableName) > 1) then
				AbilitaTipoFoto = true
			else
				AbilitaTipoFoto = false
			end if
		else
			AbilitaTipoFoto = false
		end if
		%>
		
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr><th colspan="<%= IIF(Abilita_ElencoAddOn, "8", "7") + IIF(AbilitaTipoFoto, "1", "0") %>"><%= SectionLabel %></th></tr>
			<% if cIntero(ForeignKeyId) = 0 then %>
				<tr>
					<td class="label_no_width" colspan="<%= IIF(Abilita_ElencoAddOn, "8", "7") %>">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "E' possibile inserire le foto solo dopo aver salvato il record.", "The picture can be entered only after the record is saved.", "", "", "", "", "", "")%>
					</td>
				</tr>
			<%else
				set rs = server.CreateObject("ADODB.Recordset")
				sql = " SELECT * FROM " + TableName 
				if AbilitaTipoFoto then
					sql = sql + " INNER JOIN " + TipiFotoTableName + " ON " + TableName + "." + FieldPrefix + "_tipo_id = " + TipiFotoTableName + ".ft_id"
				end if
				sql = sql + " WHERE " + FieldForeignKey + "=" & cIntero(ForeignKeyId) & " ORDER BY " + FieldPrefix + "_ordine"
				rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
				Session("RE_FOTO_SQL") = sql
				
				if cInteger(MaxFotoPerRecord)=0 OR _
		   		   cInteger(MaxFotoPerRecord) > rs.recordcount then
		   			AbilitaNuoveFoto = true
				end if 
				
				dim colspan
				colspan = 4
				if FieldRiservata <> "" then
					colspan = colspan + 1
				end if
				if AbilitaTipoFoto then
					colspan = colspan + 1
				end if
				%>
				<tr>
					<td class="label" colspan="<%= colspan %>" style="width:74%">
						<% if rs.eof then %>
							<%= ChooseValueByAllLanguages(Session("LINGUA"), "Nessuna foto inserita.", "No picture found.", "", "", "", "", "", "")%>
						<% else %>
							<%= ChooseValueByAllLanguages(Session("LINGUA"), "Trovate n&ordm; " & rs.recordcount & " foto", rs.recordcount & " pictures found", "", "", "", "", "", "")%>
						<% end if %>
					</td>
					<td colspan="<%= IIF(Abilita_ElencoAddOn, "4", "3") %>" class="content_right" style="padding-right:0px;padding-left:0px;">
					<% 	if AbilitaNuoveFoto then
							dim tabId
							sql = " SELECT TOP 1 tab_id FROM tb_siti_tabelle t"& _
							      " INNER JOIN rel_immaginiFormati r ON t.tab_id = r.rif_tab_id"& _
								  " WHERE tab_name LIKE '"& ElementIndexTableName &"'"
							'Giacomo - 30/10/2012 -------------------------
							'modifica: ora controllo tutte le tabelle riguardanti il contenuto
							sql = " SELECT tab_id, tab_field_chiave, tab_from_sql FROM tb_siti_tabelle t"& _
							      " INNER JOIN rel_immaginiFormati r ON t.tab_id = r.rif_tab_id"& _
								  " WHERE tab_name LIKE '"& ElementIndexTableName &"'"
							dim sqlVer, rsTab
							set rsTab = conn.execute(sql)
							tabId = 0
							do while (not rsTab.EOF)
								sqlVer = " SELECT "&rsTab("tab_id")&" FROM "&rsTab("tab_from_sql")& SQL_AddOperator(rsTab("tab_from_sql"), " AND ") & _
										 rsTab("tab_field_chiave")&"="&ForeignKeyId
								tabId = cIntero(GetValueList(conn, NULL, sqlVer))
								if tabId > 0 then
									exit do
								end if
								rsTab.moveNext
							loop
							rsTab.close
							set sqlVer = nothing
							set rsTab = nothing
							'----------------------------------------------
							'tabId = CIntero(GetValueList(conn, NULL, sql))

							if tabId > 0 then
								dim fotoUrl
								sql = " SELECT COUNT(*) FROM tb_immaginiFormati f"& _
									  " INNER JOIN rel_immaginiFormati r ON f.imf_id = r.rif_imf_id"& _
									  " WHERE "& SQL_IsNULL(conn, "imf_dir") &" OR imf_dir = ''"
								if CIntero(GetValueList(conn, NULL, sql)) = 0 then
									fotoUrl = "../../amministrazione2/filemanager/FileMultiUpload.aspx?RS_URL="& GetCurrentBaseUrl() &"/"& FilePrefix &"FotoGestione.asp&RS_TAB="& tabId &"&RS_ID="& ForeignKeyId
								
								elseif cBoolean(cString(directUpload),false) AND cString(additionalPath) <> "" then
									'creo directory immobile, se non è già stata creata, e vado direttamente alla finestra di upload 
									'senza passare per la scelta della cartella di destinazione dei file
									dim fso, dir, segment, segmentPath
									set fso = Server.CreateObject("Scripting.FileSystemObject")						
									
									additionalPath = Replace(additionalPath, "/", "\")
									additionalPath = Replace(additionalPath, "\\", "\")
									
									segmentPath = Split(additionalPath, "\")
									dir = Application("IMAGE_PATH") & Application("AZ_ID") & "\images\"
									for each segment in segmentPath
										dir = dir & "\" & segment
										dir = Replace(dir, "\\", "\")
										if NOT fso.FolderExists(dir) then 'se non c'è la directory, la creo
											fso.CreateFolder(dir)
										end if
									next

									set fso = nothing
									
									fotoUrl = "\images\" & additionalPath
									fotoUrl = "../../amministrazione2/filemanager/FileMultiUpload.aspx?PATH=" & Server.UrlEncode(fotoUrl) & "&RS_URL="& GetCurrentBaseUrl() &"/"& FilePrefix &"FotoGestione.asp&RS_TAB="& tabId &"&RS_ID="& ForeignKeyId
								else
									fotoUrl = "../library/filemanager.asp?STANDALONE=1&OBJECT_TYPE="& FILE_SYSTEM_DIRECTORY &"&FILEMAN_AZ_ID="& IIF(CIntero(session("AZ_ID")) > 0, session("AZ_ID"), Application("AZ_ID")) &"&RS_URL="& GetCurrentBaseUrl() &"/"& FilePrefix &"FotoGestione.asp&RS_TAB="& tabId &"&RS_ID="& ForeignKeyId
								end if 
								%>
							<a class="button_L2" href="javascript:void(0)" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apre in una nuova finestra l'inserimento di un elenco di foto", "it opens a new window to insert a group of pictures", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedScrollWindow('<%= fotoUrl %>', 'importa_foto_<%= ForeignKeyId %>', 770, 600, true)">
								<%= ChooseValueByAllLanguages(Session("LINGUA"), "IMPORTA FOTO", "IMPORT PHOTO", "", "", "", "", "", "")%>
							</a>&nbsp;
					<% 		end if %>
							<a class="button_L2" href="javascript:void(0)" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apre in una nuova finestra l'inserimento di una foto", "it opens a new window to insert a new picture", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedScrollWindow('<%= FilePrefix %>FotoGestione.asp?FKID=<%= ForeignKeyId %>', 'nuova_foto_<%= ForeignKeyId %>', 600, 400, true)">
								<%= ChooseValueByAllLanguages(Session("LINGUA"), "NUOVA FOTO", "NEW PHOTO", "", "", "", "", "", "")%>
							</a>
					<% 	else %>
							<a class="button_L2_disabled" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "raggiunto limite massimo di foto", "maximum number of pictures reached", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
								<%= ChooseValueByAllLanguages(Session("LINGUA"), "NUOVA FOTO", "NEW PHOTO", "", "", "", "", "", "")%>
							</a>
					<% 	end if %>
					</td>
				</tr>
				<% if not rs.eof then %>
					<tr>
						<th class="l2_center" width="5%"><%= ChooseValueByAllLanguages(Session("LINGUA"), "ORDINE", "ORDER", "", "", "", "", "", "")%></th>
						<% if not FotoSingola then %>
							<th class="l2_center" width="20%">THUMBNAIL</th>
						<% end if %>
						<th class="l2_center" width="20%"><%= IIF(FotoSingola, "IMMAGINE", "ZOOM") %></th>
						<th class="l2_center"><%= ChooseValueByAllLanguages(Session("LINGUA"), "DIDASCALIA", "CAPTION", "", "", "", "", "", "")%></th>
						
						<% if AbilitaTipoFoto then %>
							<th class="l2_center"><%= ChooseValueByAllLanguages(Session("LINGUA"), "TIPO", "TYPE", "", "", "", "", "", "")%></th>
						<% end if %>
						
						<th class="l2_center" width="5%"><%= ChooseValueByAllLanguages(Session("LINGUA"), "VISIBILE", "VISIBLE", "", "", "", "", "", "")%></th>
						
						<% if FieldRiservata <> "" then %>
							<th class="l2_center" width="5%"><%= ChooseValueByAllLanguages(Session("LINGUA"), "PROTETTA", "PROTECTED", "", "", "", "", "", "")%></th>
						<% end if %>
						
						<% 
	'......................................................................................................................................
	'funzione dichiarata esternamente
	'......................................................................................................................................
					if Abilita_ElencoAddOn then
						CALL ElencoTestata_ADDON(conn)
					end if
	'......................................................................................................................................
						%>
						<th class="l2_center" width="16%" colspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "OPERAZIONI", "OPERATIONS", "", "", "", "", "", "")%></th>
					</tr>
					<%while not rs.eof %>
						<tr>
							<td class="content_center"><%= rs(FieldPrefix + "_ordine") %></td>
							<% if not FotoSingola then %>
								<td class="content"><% if CString(rs(FieldPrefix + "_thumb")) <> "" then FileLink(ImageBaseUrl + rs(FieldPrefix + "_thumb")) end if %></td>
							<% end if %>
							<td class="content"><% if CString(rs(FieldPrefix + "_zoom")) <> "" then FileLink(ImageBaseUrl + rs(FieldPrefix + "_zoom")) end if %></td>
							<td class="content"><%= Sintesi(CBLL(rs, FieldPrefix + "_didascalia", Session("LINGUA")), 30, "...") %></td>							
							<% if AbilitaTipoFoto then %>
								<td class="content"><%= CString(rs("ft_nome")) %></td>
							<% end if %>							
							<td class="content_center"><% if rs(FieldPrefix + "_visibile") then %><input type="checkbox" class="checkbox" disabled checked><% else %>&nbsp;<% end if %></td>
							<% if FieldRiservata <> "" then %>
								<td class="content_center"><% if rs(FieldRiservata) then %><input type="checkbox" class="checkbox" disabled checked><% else %>&nbsp;<% end if %></td>
							<% end if %>
					
							<% 
	'......................................................................................................................................
	'funzione dichiarata esternamente
	'......................................................................................................................................
							if Abilita_ElencoAddOn then
								CALL ElencoRow_ADDON(conn, rs)
							end if
	'......................................................................................................................................
						%>
							<td class="content_center">
								<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('<%= FilePrefix %>FotoGestione.asp?FKID=<%= ForeignKeyId %>&ID=<%= rs(FieldPrefix + "_id") %>', 'import_foto_<%= ForeignKeyId %>_<%= rs(FieldPrefix + "_id") %>', 600, 400, true)">
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "MODIFICA", "MODIFY", "", "", "", "", "", "")%>
								</a>
							</td>
							<td class="content_center">
								<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('<%= DeleteKey %>','<%= rs(FieldPrefix + "_id") %>');">
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANCELLA", "DELETE", "", "", "", "", "", "")%>
								</a>
							</td>
						</tr>
						<% rs.movenext
					wend
				end if
				rs.close
				set rs = nothing
			end if %>
		</table>
	<%
	end sub
	
	
	
	'......................................................................................................................................
	'Gestione form di inserimento e modifica nuova immagine
	'......................................................................................................................................
	Public Sub FormGestione()
		if elementID = 0 then
			if Request.ServerVariables("REQUEST_METHOD")="POST" then
				
				'Salvataggio della foto
				dim oSalva
				Set oSalva = New OBJ_Salva
				with oSalva
					.ConnectionString		= Application("DATA_ConnectionString")
					.Requested_Fields_List	= IIF(MandatoryThumb, "tft_" + FieldPrefix + "_thumb", "") + _
										      IIF(MandatoryThumb AND MandatoryZoom, ";", "") + _
											  IIF(MandatoryZoom, "tft_" + FieldPrefix + "_zoom", "")
					.Checkbox_Fields_List 	= ""
					.Page_Ins_Form			= ""
					.Page_Mod_Form			= ""
					.Next_Page				= ""
					.Next_Page_ID			= FALSE
					.Table_Name				= TableName
					.id_Field				= FieldPrefix + "_id"
					.Read_New_ID			= TRUE
					.isReport 				= TRUE
					.Gestione_Relazioni 	= TRUE
					
					.Salva()
				end with
				
				
				if Session("ERRORE") = "" then 
					%>
					<script language="JavaScript" type="text/javascript">
						opener.location.reload(true);
						<% if request("salva_elenco") <> "" then %>
							window.close();
						<% else %>
							window.reload(true);
						<% end if %>
					</script>
					<%
				end if
			end if
			
			dim rsr
			Set rsr = Server.CreateObject("ADODB.Recordset")
			if request("goto")<>"" then
				CALL GotoRecord(conn, rsr, Session("RE_FOTO_SQL"), FieldPrefix + "_id", FilePrefix&"FotoGestione.asp?FKID=" & Request("FKID"))
			end if
			Set rsr = nothing
			
			'----------------------------------------------------- 
			if request("ID")<>"" then
				sezione_testata = ChooseValueByAllLanguages(Session("LINGUA"), "modifica foto", "modify picture", "", "", "", "", "", "")
			else
				sezione_testata = ChooseValueByAllLanguages(Session("LINGUA"), "inserimento nuova foto", "insert new picture", "", "", "", "", "", "") 
			end if
			%>
			<!--#INCLUDE FILE="../library/Intestazione_Ridotta_include.asp" -->
			<%'----------------------------------------------------- 
			
			dim rs, sql, AbilitaTipoFoto
			if cIntero(request("ID")) > 0 then
				Set rs = Server.CreateObject("ADODB.Recordset")
				
				sql = " SELECT * FROM " + TableName + " WHERE " + FieldPrefix + "_id=" & cIntero(request("ID"))
				rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
			else
				rs = null
			end if
			
			if (TipiFotoTableName <> "") then
				if (GetValueList(conn, NULL, "SELECT COUNT(*) FROM " & TipiFotoTableName) > 1) then
					AbilitaTipoFoto = true
				else
					AbilitaTipoFoto = false
				end if
			else
				AbilitaTipoFoto = false
			end if
			%>
			<div id="content_ridotto">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
					<form action="" method="post" id="form1" name="form1"<% if FotoSingola then %> onsubmit="ImpostaThumbDaZoom()" <% end if %>>
					<input type="hidden" name="tfn_<%= FieldForeignKey %>" value="<%= cIntero(request("FKID")) %>">
					<caption>
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td class="caption">
									<% if request("ID")<>"" then %>
										<%=ChooseValueByAllLanguages(Session("LINGUA"), "Modifica foto", "Modify picture", "", "", "", "", "", "")%>
									<% else %>
										<%=ChooseValueByAllLanguages(Session("LINGUA"), "Inserimento nuova foto", "Insert new picture", "", "", "", "", "", "")%>
									<% end if %>
								</td>
								<td align="right" style="font-size: 1px;">
									<a class="button" href="?FKID=<%= request("FKID") %>&ID=<%= request("ID") %>&goto=PREVIOUS" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "vai alla precedente", "previous", "", "", "", "", "", "")%>">
										&lt;&lt; <%= ChooseValueByAllLanguages(Session("LINGUA"), "PRECEDENTE", "PREVIOUS", "", "", "", "", "", "")%>
									</a>
									&nbsp;
									<a class="button" href="?FKID=<%= request("FKID") %>&ID=<%= request("ID") %>&goto=NEXT" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "vai alla successiva", "next", "", "", "", "", "", "")%>">
										<%= ChooseValueByAllLanguages(Session("LINGUA"), "SUCCESSIVA", "NEXT", "", "", "", "", "", "")%> &gt;&gt;
									</a>
								</td>
							</tr>
						</table>
					</caption>
					<tr><th colspan="3"><%=ChooseValueByAllLanguages(Session("LINGUA"), "DATI DELLA FOTO", "PICTURE INFO", "", "", "", "", "", "")%></th></tr>
					<% if not FotoSingola then %>
						<tr>
							<td class="label_no_width" rowspan="2"><%=ChooseValueByAllLanguages(Session("LINGUA"), "immagine", "image", "", "", "", "", "", "")%></td>
							<td class="label_no_width">thumbnail</td>
							<td class="content">
								<% CALL WriteFileSystemPicker_Input(ImageWebId, FILE_SYSTEM_FILE, "images", EXTENSION_IMAGES, "form1", "tft_" + FieldPrefix + "_thumb", CBR(rs, FieldPrefix + "_thumb", "tft_"), "width:320px", true, MandatoryThumb)
								
								if NoteFormatoThumbnail<>"" then %>
									<div class="note"><%= NoteFormatoThumbnail %></div>
								<% end if %>
							</td>
						</tr>
						<tr>
							<td class="label_no_width">zoom</td>
							<td class="content">
								<% CALL WriteFileSystemPicker_Input(ImageWebId, FILE_SYSTEM_FILE, "images", EXTENSION_IMAGES, "form1", "tft_" + FieldPrefix + "_zoom", CBR(rs, FieldPrefix + "_zoom", "tft_"), "width:320px", true, MandatoryZoom)
								
								if NoteFormatoZoom<>"" then %>
									<div class="note"><%= NoteFormatoZoom %></div>
								<% end if %>
							</td>
						</tr>
					<% else %>
						<tr>
							<td class="label_no_width"><%=ChooseValueByAllLanguages(Session("LINGUA"), "immagine", "image", "", "", "", "", "", "")%></td>
							<td class="content" colspan="2">
								<% CALL WriteFileSystemPicker_Input(ImageWebId, FILE_SYSTEM_FILE, "images", EXTENSION_IMAGES, "form1", "tft_" + FieldPrefix + "_zoom", CBR(rs, FieldPrefix + "_zoom", "tft_"), "width:370px", true, true) %>
								<input type="hidden" name="tft_<%= FieldPrefix %>_thumb" value="" />
							</td>
						</tr>
						<script language="JavaScript" type="text/javascript">
							function ImpostaThumbDaZoom(){
								form1.tft_<%= FieldPrefix %>_thumb.value = form1.tft_<%= FieldPrefix %>_zoom.value;
							}
						</script>
					<% end if %>
					<tr>
						<td class="label_no_width" style="width:13%;" rowspan="2"><%=ChooseValueByAllLanguages(Session("LINGUA"), "pubblicazione", "publication", "", "", "", "", "", "")%></td>
						<td class="label_no_width" style="width:8%;"><%=ChooseValueByAllLanguages(Session("LINGUA"), "ordine", "order", "", "", "", "", "", "")%></td>
						<td class="content">
							<input type="text" class="text" name="tfn_<%= FieldPrefix %>_ordine" value="<%= CBR(rs, FieldPrefix + "_ordine", "tfn_") %>" maxlength="10" size="4">
						</td>
					</tr>
					
					<% dim visibile
					if request.servervariables("REQUEST_METHOD")<>"POST" AND not IsNull(rs) then
						visibile = rs(FieldPrefix + "_visibile")
					elseif request("tfn_" & FieldPrefix & "_visibile")<>"" then
						visibile = cIntero(request("tfn_" + FieldPrefix + "_visibile"))>0
					else
						visibile = true
					end if
					 %>
					<tr>
						<td class="label_no_width"><%=ChooseValueByAllLanguages(Session("LINGUA"), "visibile:", "visible:", "", "", "", "", "", "")%></td>
						<td class="content">
							<input type="radio" class="checkbox" value="1" name="tfn_<%= FieldPrefix %>_Visibile" <%= chk(visibile) %>>
							<%=ChooseValueByAllLanguages(Session("LINGUA"), "si", "yes", "", "", "", "", "", "")%>
							<input type="radio" class="checkbox" value="0" name="tfn_<%= FieldPrefix %>_Visibile" <%= chk(not visibile) %>>
							no
						</td>
					</tr>
					<% if FieldRiservata <> "" then %>
						<% dim riservata
						if request.servervariables("REQUEST_METHOD")<>"POST" AND not IsNull(rs) then
							riservata = cBoolean(rs(FieldRiservata), false)
						elseif request("tfn_" & FieldRiservata)<>"" then
							riservata = cIntero(request("tfn_" + FieldRiservata))>0
						else
							riservata = false
						end if
						 %>
						<tr>
							<td class="label_no_width"><%=ChooseValueByAllLanguages(Session("LINGUA"), "riservata:", "reserved:", "", "", "", "", "", "")%></td>
							<td class="content" colspan="2">
								<input type="radio" class="checkbox" value="1" name="tfn_<%= FieldRiservata %>" <%= chk(riservata) %>>
								<%=ChooseValueByAllLanguages(Session("LINGUA"), "si", "yes", "", "", "", "", "", "")%>
								<input type="radio" class="checkbox" value="0" name="tfn_<%= FieldRiservata %>" <%= chk(not riservata) %>>
								no
							</td>
						</tr>
					<% end if %>
					<% if AbilitaTipoFoto then %>
						<tr>					
							<td class="label_no_width" style="width:8%;"><%=ChooseValueByAllLanguages(Session("LINGUA"), "tipo", "type", "", "", "", "", "", "")%></td>
							<td class="label_no_width" style="width:8%;" colspan="2">
								<%  sql = " SELECT * FROM " + TipiFotoTableName
									CALL dropDown(conn, sql, "ft_id", "ft_nome", "tfn_" + FieldPrefix + "_tipo_id", CBR(rs, FieldPrefix + "_tipo_id", "tfn_"), true, "", "") %>
							</td>
						</tr>
					<% else %>
						<% if cString(TipiFotoTableName)<>"" then %>
							<input type="hidden" name="tfn_<%=FieldPrefix%>_tipo_id" value="<%=GetValueList(conn, NULL, "SELECT TOP 1 ft_id FROM " + TipiFotoTableName) %>">
						<% end if %>
					<% end if
					
	'......................................................................................................................................
	'funzione dichiarata esternamente
	'......................................................................................................................................
					if Abilita_FormGestioneAddOn then
						CALL FormGestione_ADDON(conn, rs)
					end if
	'......................................................................................................................................
					%>
					<tr><th colspan="3"><%=ChooseValueByAllLanguages(Session("LINGUA"), "DIDASCALIA", "CAPTION", "", "", "", "", "", "")%></th></tr>
					<%dim i
					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
						<tr>
							<td class="content" colspan="4">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
									<tr>
										<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" alt="" border="0"></td>
										<td><textarea style="width:100%;" rows="2" name="tft_<%= FieldPrefix %>_didascalia_<%= Application("LINGUE")(i) %>"><%= CBR(rs, FieldPrefix + "_didascalia_" & Application("LINGUE")(i), "tft_") %></textarea></td>
									</tr>
								</table>
							</td>
						</tr>
					<%next %>
					<tr>
						<td class="footer" colspan="4">
							<%= ChooseValueByAllLanguages(Session("LINGUA"), "(*) Campi obbligatori.", "(*) Mandatory fields.", "", "", "", "", "", "")%>
							<input style="width:12%;" type="submit" class="button" name="salva" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "SALVA", "SAVE", "", "", "", "", "", "")%>">
							<input style="width:30%;" type="submit" class="button" name="salva_elenco" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "SALVA & TORNA ALL'ELENCO", "SAVE & RETURN TO THE LIST", "", "", "", "", "", "")%>">
							<input style="width:12%;" type="button" class="button" name="annulla" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "ANNULLA", "CANCEL", "", "", "", "", "", "")%>" onclick="window.close();">
						</td>
					</tr>
					</form>
				</table>
			</div>
			</body>
			</html>
			<script language="JavaScript" type="text/javascript">
				FitWindowSize(this);
			</script>
			
<% 		else
			
			conn.begintrans
			
			set rs = Server.CreateObject("ADODB.RecordSet")
			CALL Gestione_Relazioni_foto(conn, rs, 0)
			set rs = nothing
			
			conn.committrans
		end if
	end sub
	
	
	'......................................................................................................................................
	'Gestione form di import immagini da directory
	'......................................................................................................................................
	Public Sub Import()
		dim dir, i, AbilitaTipoFoto
		dir = request("DIR")
		
		if (TipiFotoTableName <> "") then
			AbilitaTipoFoto = true
		else
			AbilitaTipoFoto = false
		end if
			
		'salvo
		if request.form("salva") <> "" then
			
			dim rs, rsAux, sql, field
			set rs = server.createobject("adodb.recordset")
			set rsAux = server.createobject("adodb.recordset")
			rs.open "SELECT TOP 1 * FROM " + tableName, conn, adOpenKeyset, adLockOptimistic
			conn.BeginTrans
			
			for each field in request.form
				if Left(field, 5) = "file_" then
					i = right(field, len(field) - 5)
					rs.AddNew
						rs(FieldPrefix + "_zoom") = dir + "/" + Replace(request.form(field), " ", "_")
						rs(FieldPrefix + "_thumb") = rs(FieldPrefix + "_zoom")
						rs(FieldPrefix + "_visibile") = true
						rs(FieldPrefix + "_ordine") = CIntero(request.form(i + "_ordine"))
						rs(fieldForeignKey) = cIntero(request("FKID"))
						if AbilitaTipoFoto then
							rs(FieldPrefix + "_tipo_id") = CIntero(GetValueList(conn, NULL, "SELECT TOP 1 ft_id FROM " + TipiFotoTableName ))
						end if
					rs.Update
					
					'gestione relazioni
					CALL Gestione_Relazioni_record(conn, rsAux, rs(fieldPrefix + "_id"))
				end if
			next
			
			rs.close
			CALL Gestione_Relazioni_finale(conn, rsAux)
			set rs = nothing
			set rsAux = nothing
			conn.CommitTrans
			
			if Session("ERRORE") = "" then 
				%>
				<script language="JavaScript" type="text/javascript">
					opener.location.reload(true);
					window.close();
				</script>
				<%
			end if
		end if
		
		'----------------------------------------------------- 
		sezione_testata = "import foto da una directory" %>
		<!--#INCLUDE FILE="../library/Intestazione_Ridotta_include.asp" -->
		<%'----------------------------------------------------- 
		%>
		<div id="content_ridotto">
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
				<form action="" method="post" id="form1" name="form1">
				<input type="hidden" name="tfn_<%= FieldForeignKey %>" value="<%= cIntero(request("FKID")) %>">
				<caption>Import foto da una directory - passo <%= IIF(dir = "", 1, 2) %> di 2</caption>
				<% if dir = "" then %>
					<tr><th colspan="2">SELEZIONE DIRECTORY</th></tr>
					<tr>
						<td rowspan="2" class="label">directory</td>
						<td class="content">
							<% CALL WriteFileSystemPicker_Input(ImageWebId, FILE_SYSTEM_DIRECTORY, "images", "", "form1", "dir", dir, "width:370px", true, true) %>
						</td>
					</tr>
					<tr>
						<td class="content notes">
							Selezionare la directory che contiene i file da importare.
						</td>
					</tr>
					<tr>
						<td class="footer" colspan="2">
							(*) Campi obbligatori.
							<input type="submit" class="button" name="avanti" value="AVANTI &raquo;">
						</td>
					</tr>
				<% else %>
					<input type="hidden" name="dir" value="<%= dir %>">
					<tr><th colspan="4">DIRECTORY SELEZIONATA</th></tr>
					<tr>
						<td class="label">directory</td>
						<td class="content_b" colspan="3">
							<%= dir %>
						</td>
					</tr>
					<tr><th colspan="4">FILE CONTENUTI</th></tr>
					<%
					dim d, fso, f, f1, extension
					set d = new directory
					d.RelativeDirPath = "/images" + dir
					
					set fso = CreateObject("Scripting.FileSystemObject")
					set f = fso.GetFolder(d.DIRPath)
					if f.Files.Count = 0 then %>
						<tr><td class="label" colspan="4">Nessun file trovato.</td></tr>
					<% else %>
						<tr>
							<td colspan="4">
								<table cellpadding="0" cellspacing="1" style="width: 100%;">
									<tr>
										<td class="label_no_width" colspan="6">Seleziona i file da importare ed immetti l'ordine di visualizzazione.</td>
									</tr>
									<tr>
										<th class="l2_center" width="6%">SEL.</th>
										<th class="l2_center" width="11%">ORDINE</th>
										<th class="L2">NOME</th>
										<th class="L2" width="18%">TIPO</th>
										<th class="l2_center" width="15%">DIMENSIONE</th>
									</tr>
									<% for each f1 in f.Files
										extension = File_Extension( f1.name )
										if instr(1, EXTENSION_IMAGES, extension, vbTextCompare)>0 then
											i = i + 1 %>
											<tr>
												<td class="content_center">
													<input type="checkbox" name="file_<%= i %>" value="<%= f1.name %>" checked class="checkbox">
												</td>
												<td class="content_center">
													<input type="text" class="text" name="<%= i %>_ordine" value="<%= i %>" size="3">
												</td>
												<td class="content"><% FileLink(ImageBaseUrl + d.RelativeURL(f1.name)) %></td>
												<td class="content"><%= File_Type( Extension ) %></td>
												<td class="content"><%= File_Dimension(f1.size)  %></td>
											</tr>
										<% end if
									next %>
								</table>
							</td>
						</tr>
					<% end if
					set fso = nothing %>
					<tr>
						<td class="footer" colspan="4">
							(*) Campi obbligatori.
							<input type="submit" class="button" name="avanti" value="&lt;&lt; INDIETRO">
							<input type="submit" class="button" name="salva" value="IMPORTA FOTO SELEZIONATE">
							
						</td>
					</tr>
				<% end if %>
				</form>
			</table>
		</div>
		</body>
		</html>
		<script language="JavaScript" type="text/javascript">
			FitWindowSize(this);
		</script>
<%	End Sub
	
	
	'......................................................................................................................................
	'Parte di codice standard eseguito nel salvataggio delle foto. (da richiamare in gestione_relazioni_record su file civetta)
	'......................................................................................................................................
	Public Sub Gestione_Relazioni_foto(conn, rs, ID)
		dim chiaveEsterna, sql
		if elementID = 0 then
			chiaveEsterna = request("tfn_" + FieldForeignKey)
		else
			chiaveEsterna = elementID
	
			'inserimento relazione
			sql = " SELECT * FROM "& tableName & _
				  " WHERE " + FieldForeignKey + " = " & chiaveEsterna
			if DB_Type(conn) = DB_Access then
				sql = sql + " AND ( "& fieldPrefix &"_thumb LIKE '"& elementThumb &"'"& _
								  " OR "& fieldPrefix &"_zoom LIKE '"& elementZoom &"' ) "
			else
				sql = sql + " AND ( Replace("& fieldPrefix &"_thumb, '//', '/') LIKE '"& elementThumb &"'"& _
								  " OR Replace("& fieldPrefix &"_zoom, '//', '/') LIKE '"& elementZoom &"' ) "
			end if

			'calcola ordine foto
			dim MaxOrdine 
			'commentato l'ordine massimo per fare in modo che venga ordinato in ordine alfabetico
			'MaxOrdine = cIntero(getValueList(conn, rs, " SELECT TOP 1 " & fieldPrefix & "_ordine " + _
			'																  " FROM " & tableName & _
			'																  " WHERE " + FieldForeignKey + " = " & chiaveEsterna & " ORDER BY " & fieldPrefix & "_ordine DESC ")) + 10
			dim fileName
			'mette l'ordine alfabetico delle prime due lettere.
			fileName = uCase(File_name(elementZoom))
			MaxOrdine = cString(Asc(left(fileName, 1))) & FixLenght(cString(Asc(mid(fileName, 2, 1))), "0", 3)
	
			rs.open sql, conn, adOpenDynamic, adLockOptimistic
			if rs.eof then
				rs.addnew
				rs(fieldPrefix &"_ordine") = cIntero(MaxOrdine)
			end if
			rs(fieldPrefix &"_visibile") = true
			rs(fieldPrefix &"_thumb") = elementThumb
			rs(fieldPrefix &"_zoom") = elementZoom
			rs(FieldForeignKey) = chiaveEsterna
			if TipiFotoTableName <> "" then
				rs(FieldPrefix + "_tipo_id") = CIntero(GetValueList(conn, NULL, "SELECT TOP 1 ft_id FROM " + TipiFotoTableName ))
			end if
			rs.update
			rs.close
		end if

		'..............................................................................
		'sincronizzazione con i contenuti e l'indice
		CALL Index_UpdateItem(conn, ElementIndexTableName, chiaveEsterna, false)
		'..............................................................................

		if ElementUpdateParams then
			'imposta date aggiornamento elemento "padre" a cui sono collegate le immagini
		    CALL UpdateParams(conn, elementDatabaseTableName, ElementFieldPrefix + "_", ElementFieldPrefix + "_id", chiaveEsterna, false)
		end if
		
		response.clear()
		response.write "OK"
	end Sub
	
	
	
	'......................................................................................................................................
	'Impostazioni della Classe di cancellazione dell'elemento.
	'......................................................................................................................................
	Public Sub SetDeleteSettings(ClassDelete, Caption)
		dim rs, sql
		with ClassDelete
			.Index = oIndex
			.Message = ChooseValueByAllLanguages(Session("LINGUA"), "Cancellare la foto ?", "Delete picture ?", "", "", "", "", "", "")
			.Name_Field = FieldPrefix + "_thumb"
			.ID_Field = FieldPrefix + "_ID"
			.Table = TableName
			.Caption = Caption
			.AfterDelete = TRUE
			.DeleteRelations = TRUE
			
			
			'aggiunge flag per la cancellazione dei file.
			set rs = server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM " & .Table & " WHERE " + .ID_Field + " = " & cIntero(.ID_Value)
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
			
			if not rs.eof then
				
				if cString(rs(FieldPrefix + "_thumb"))<>"" then
					if FileCanBeRemoved(conn, NULL, NULL, FILE_TYPE_IMAGE, Application("AZ_ID"), rs(FieldPrefix + "_thumb")) then
					
						if FileUsedCount(rs(FieldPrefix + "_thumb"), .ID_Value)>0 then
							.Note = .Note + ChooseValueByAllLanguages(Session("LINGUA"), " Il file immagine thumbnail &quot;<b>", " Impossible delete thumbnail image &quot;<b>", "", "", "", "", "", "") + rs(FieldPrefix + "_thumb") & ChooseValueByAllLanguages(Session("LINGUA"), "</b>&quot; non pu&ograve; essere cancellato perch&egrave; collegato anche ad altri record.<br>", "</b>&quot; because it is linked to other records.<br>", "", "", "", "", "", "")
						else
							'aggiunge richiesta per cancellazione file "thumb"
							.AddOption "delete_files_thumb" , ChooseValueByAllLanguages(Session("LINGUA"), "cancella anche il file immagine thumbnail &quot;<b>", "delete also thumbnail image files &quot;<b>", "", "", "", "", "", "") + rs(FieldPrefix + "_thumb") & "</b>&quot;", false, ""
						end if
					else
						.Note = .Note + ChooseValueByAllLanguages(Session("LINGUA"), " Il file immagine thumbnail &quot;<b>", " Impossible delete thumbnail image &quot;<b>", "", "", "", "", "", "") + rs(FieldPrefix + "_thumb") + rs(FieldPrefix + "_thumb") & ChooseValueByAllLanguages(Session("LINGUA"), "</b>&quot; non pu&ograve; essere cancellato perch&egrave; utilizzato nella costruzione delle pagine.<br>", "</b>&quot; because it is used to build some pages.<br>", "", "", "", "", "", "")
					end if
				end if
				if not FotoSingola AND cString(rs(FieldPrefix + "_zoom"))<>"" then
					if FileCanBeRemoved(conn, NULL, NULL, FILE_TYPE_IMAGE, Application("AZ_ID"), rs(FieldPrefix + "_zoom")) then
						if FileUsedCount(rs(FieldPrefix + "_zoom"), .ID_Value)>0 then
							.Note = .Note + " Il file immagine zoom &quot;<b>" + rs(FieldPrefix + "_zoom") & "</b>&quot; non pu&ograve; essere cancellato perch&egrave; collegato anche ad altri record.<br>"
						else
							'aggiunge richiesta per cancellazione file "zoom"
							.AddOption "delete_files_zoom" , ChooseValueByAllLanguages(Session("LINGUA"), "cancella anche il file immagine zoom &quot;<b>", "delete also zoom image files &quot;<b>", "", "", "", "", "", "") + rs(FieldPrefix + "_zoom") & "</b>&quot;", false, ""
						end if
					else
						.Note = .Note + ChooseValueByAllLanguages(Session("LINGUA"), " Il file immagine zoom &quot;<b>", " Impossible delete zoom image &quot;<b>", "", "", "", "", "", "") + rs(FieldPrefix + "_zoom") & ChooseValueByAllLanguages(Session("LINGUA"), "</b>&quot; non pu&ograve; essere cancellato perch&egrave; utilizzato nella costruzione delle pagine.<br>", "</b>&quot; because it is used to build some pages.<br>", "", "", "", "", "", "")
					end if
				end if
			end if
			rs.close
		end with
	end sub
	
	
	
	'......................................................................................................................................
	'Funzione applicata all'evento di cancellazione.
	'Verifica se l'impostazione di rimozione dei file e' impostata e predispone per la cancellazione definitiva
	'......................................................................................................................................
	Public Sub OnDeleteRelazioni( conn, ID )
		if request("delete_files_thumb")<>"" OR _
   		   request("delete_files_zoom")<>"" then
		   	
			dim sql, rs
			set rs = server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM " & TableName & " WHERE " + FieldPrefix + "_id = " & cIntero(ID)
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
			
			if request("delete_files_thumb")<>"" then
				Session("delete_files_thumb_" & ID) = rs(FieldPrefix & "_thumb") 
			end if
			if NOT FotoSingola AND request("delete_files_zoom")<>"" then
				Session("delete_files_zoom_" & ID) = rs(FieldPrefix & "_zoom")
			end if
			
			rs.close
			
		end if
	end sub
	
	
	'......................................................................................................................................
	'Funzione applicata all'evento di cancellazione conlcusa.
	'Eventualmente rimuove i file fisici
	'......................................................................................................................................
	Public Sub OnAfterDelete(conn, ID)
		
		'cancella files.
		if Session("delete_files_thumb_" & ID)<>"" then
			CALL FileDelete(Session("delete_files_thumb_" & ID))
		end if
		
		if NOT FotoSingola AND Session("delete_files_zoom_" & ID)<>"" then
			CALL FileDelete(Session("delete_files_zoom_" & ID))
		end if
		
	end sub
	
    
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'FUNZIONI E METODI PRIVATI
'******************************************************************************************************************************************
	
	
	'......................................................................................................................................
	'funzione che restituisce il numero di volte che un file e' stato utilizzato nella struttura corrente.
	'	FileName:		nome del file da cancellare
	'	ExcludePhotoId:	lista di id della foto da escludere dal conteggio
	'......................................................................................................................................
	private function FileUsedCount(FileName, ExcludePhotoId)
		dim sql
		
		sql = "SELECT COUNT(*) FROM " + TableName + _
			  " WHERE ( " + FieldPrefix + "_thumb LIKE '" + ParseSql(FileName, adChar) + "' OR " + FieldPrefix + "_zoom LIKE '" + ParseSql(FileName, adChar) + "' ) "
		if cString(ExcludePhotoId)<>"" then
			sql = sql + " AND " + FieldPrefix + "_id NOT IN (" & ExcludePhotoId & ") "
		end if
		FileUsedCount = cIntero(GetValueList(conn, NULL, sql))
		
	end function
	
	
	'......................................................................................................................................
	' procedura che cancella il file indicato
	'......................................................................................................................................
	private sub FileDelete(FileName)
		dim fso
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		
		CALL FileRemove(fso, Application("IMAGE_PATH") & "\" & ImageWebId & "\" + FILE_TYPE_IMAGE, replace(FileName, "/", "\"), false)
		
		set fso = nothing
	end sub
	
	
end class
 %>