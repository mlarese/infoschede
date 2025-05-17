<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 2147483647 %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<!--#INCLUDE FILE="Tools4Issuu.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva

	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_doc_titolo_it;tfd_doc_pubblicazione;"
	if cBoolean(Session("CATEGORIE_NEXTMEMO2_ABILITATE"), false) then
		Classe.Requested_Fields_List =  Classe.Requested_Fields_List & ";tfn_doc_categoria_id"
	end if
	Classe.Checkbox_Fields_List 	= "doc_visibile;doc_protetto;chk_doc_catalogo_sfogliabile"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "mtb_documenti"
	Classe.id_Field					= "doc_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE


'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	'..............................................................................
	'sincronizzazione con i contenuti e l'indice
	CALL Index_UpdateItem(conn, Classe.Table_Name, ID, false)
	'..............................................................................
	
	CALL UpdateParams(conn, "mtb_documenti", "doc_", "doc_id", ID, cBoolean(cIntero(request("ID")) = 0, false))

	dim val, sql, ut_id, url_catalogo
	
	'inserimento relazioni tra profili e documento
	sql = "DELETE FROM mrel_doc_profili WHERE rdp_doc_id = " & ID
	conn.Execute(sql)
	for each val in Split(request.form("profili_associati"), ",")
		if CIntero(val) > 0 then
			sql = " INSERT INTO mrel_doc_profili(rdp_doc_id, rdp_profilo_id)"& _
				  " VALUES (" & ID & ", " & val & ")"
			conn.Execute(sql)
		end if
	next
	val = ""

	'inserimento relazioni tra admin e documento
	sql = "DELETE FROM mrel_doc_admin WHERE rda_doc_id = " & ID
	conn.Execute(sql)
	for each val in Split(request.form("admin_associati"), ";")
		if CIntero(val) > 0 then
			sql = " INSERT INTO mrel_doc_admin(rda_doc_id, rda_admin_id)"& _
				  " VALUES (" & ID & ", " & val & ")"
			conn.Execute(sql)
		end if
	next
	val = ""
		
	'inserimento relazioni utenti e documento
	sql = "DELETE FROM mrel_doc_utenti WHERE rdu_doc_id = " & ID
	conn.Execute(sql)
	for each val in Split(request.form("utenti_associati"), ";")
		if CIntero(val) > 0 then
			ut_id = GetValueList(conn, null, "SELECT ut_ID FROM tb_utenti WHERE ut_NextCom_id = " & val)
			sql = " INSERT INTO mrel_doc_utenti(rdu_doc_id, rdu_utenti_id)"& _
				  " VALUES (" & ID & ", " & ut_id & ")"
			'response.write sql
			conn.Execute(sql)
		end if
	next
	val = ""

	'catalogo sfogliabile
	if cString(request("chk_doc_catalogo_sfogliabile"))<>"" then
		dim execute_path, file_path, url, i, file_data, catalogo_last_mod, doc_file, x, lingua
	
		for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			doc_file = cString(request("tft_doc_file_" & lingua))
			
			if doc_file <> "" AND File_Extension(doc_file) = FILE_TYPE_PDF then
				'controllo data file pdf da convertire con data catalogo scritta su db
				file_path = application("IMAGE_PATH")&application("AZ_ID")&"\images\"&doc_file
				file_data = DateTimeISO(File_Date(file_path))
				if cString(request("doc_data_modifica_catalogo_" & lingua)) <> "" then
					catalogo_last_mod = DateTimeIso(cString(request("doc_data_modifica_catalogo_" & lingua)))
				else
					catalogo_last_mod = file_data
				end if

				if i > 0 then
					for x=lbound(Application("LINGUE")) to i
						if doc_file = cString(request("tft_doc_file_"&Application("LINGUE")(x))) AND x<>i then
							sql = "SELECT doc_url_catalogo_"&Application("LINGUE")(x)&" FROM mtb_documenti WHERE doc_id = " & ID
							url_catalogo = cString(GetValueList(conn, NULL, sql))
						end if
					next
				end if
				
				if url_catalogo = "" AND (file_data >= catalogo_last_mod OR cString(request("old_doc_file_" & lingua)) <> doc_file) then
					
					if cBool(Application("CatalogoSfogliabileNonIssuu")) = true then
						
						url = "http://"&Request.ServerVariables("SERVER_NAME") & "/amministrazione2/nextMemo2/CreaCatalogo.aspx"
						execute_path = "?FILEPATH=\"&replace(doc_file,"/","\") & "&LINGUA="&Application("LINGUE")(i)&"&ID="&ID
						url = url & execute_path 
						
						'response.write "eseguo URL: " & url & "<br>"
						CALL WriteLogAdmin(conn, "mtb_documenti", ID, "GENERAZIONE_CAT_SFOGLIABILE","Url Generazione: " + url)
						
						url_catalogo = ExecuteHttpUrl(url)
						
						url_catalogo = replace(url_catalogo," ","")
						url_catalogo = Trim(url_catalogo)
						url_catalogo = Left(url_catalogo,inStr(url_catalogo, ".html")+4)
						url_catalogo = ParseSQL(url_catalogo,adChar)
						'response.write "restituisce: -" & url_catalogo & "--fine--<br><br>"
					else
						
						'creo il catalogo issuu
						Dim postData, httpRequest, oXmlResponse, signatureData, tmpResponse
						
						Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
						httpRequest.Open "POST", ISSUU_url, False
						httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		
						CALL AddIssuuPost("access", ISSUU_access, false, postData, signatureData)
						CALL AddIssuuPost("action", ISSUU_action, false, postData, signatureData)
						CALL AddIssuuPost("apiKey", ISSUU_apiKey, false, postData, signatureData)
						CALL AddIssuuPost("category", ISSUU_category, false, postData, signatureData)		'travel & events
						CALL AddIssuuPost("commentsAllowed", ISSUU_commentsAllowed, false, postData, signatureData)
						CALL AddIssuuPost("description", ChooseValueByAllLanguages(lingua, request("tft_doc_estratto_it"), _
																						   request("tft_doc_estratto_en"), _
																						   request("tft_doc_estratto_de"), _
																						   request("tft_doc_estratto_fr"), _
																						   request("tft_doc_estratto_es"), _
																						   request("tft_doc_estratto_ru"), _
																						   request("tft_doc_estratto_cn"), _
																						   request("tft_doc_estratto_pt")), _
																						   true, postData, signatureData)
						CALL AddIssuuPost("downloadable", ISSUU_downloadable, false, postData, signatureData)
						CALL AddIssuuPost("explicit", "false", false, postData, signatureData)
						CALL AddIssuuPost("format", "xml", false, postData, signatureData)
						CALL AddIssuuPost("infoLink", GetSiteUrl(conn, 0, 0), true, postData, signatureData)
						CALL AddIssuuPost("language", lingua, false, postData, signatureData)
						CALL AddIssuuPost("publishDate", DateISO(ConvertForSave_Date(request("tfd_doc_pubblicazione"))), false, postData, signatureData)
						CALL AddIssuuPost("ratingsAllowed", ISSUU_ratingsAllowed, false, postData, signatureData)
						CALL AddIssuuPost("slurpUrl", GetUrlImage(doc_file, 0), true, postData, signatureData)	'Server.UrlEncode(GetUrlImage(doc_file, 0)), postData, signatureData)
						CALL AddIssuuPost("tags", ISSUU_tags, true, postData, signatureData)
						CALL AddIssuuPost("title", ChooseValueByAllLanguages(lingua, request("tft_doc_titolo_it"), _
																					 request("tft_doc_titolo_en"), _
																					 request("tft_doc_titolo_de"), _
																					 request("tft_doc_titolo_fr"), _
																					 request("tft_doc_titolo_es"), _
																					 request("tft_doc_titolo_ru"), _
																					 request("tft_doc_titolo_cn"), _
																					 request("tft_doc_titolo_pt")), _
																					 true, postData, signatureData)
						CALL AddIssuuPost("type", ISSUU_type, false, postData, signatureData)		'catalog
						CALL AddIssuuSignature(postData, signatureData)
						
						httpRequest.Send postData
						tmpResponse = httpRequest.ResponseXML.xml
						
						if instr(1, tmpResponse, "stat=""ok""", vbTextCompare)>0 then
							'conversione OK - salva dati
							dim docName, attr
							set oXmlResponse = Server.CreateObject("MSXML2.DOMDocument")
							oXmlResponse.loadXML(tmpResponse)
							for each attr in oXmlResponse.SelectSingleNode("rsp/document").Attributes
								'Response.Write attr.Name & "=" & attr.Value & "<br>"
								if attr.name = "name" then
									docName = attr.value
								end if
							next
							
							url_catalogo = ISSUU_baseurl & docName & ISSUU_parameters
							
							CALL SendEmailSupportEX("Issuu: "& Request.ServerVariables("SERVER_NAME") & " - conversione pdf to ISSUU", _
													"conversione file riuscita: " & url_catalogo & vbCrLF & _
													"Risultato conversione:" & vbCrLF & _
													tmpResponse & vbCrLF & _
													"POST Inviato:" & vbCrLf & _
													postData)
							'response.write "docName=" & docName & "<br>"
							'response.write "url_catalogo=" & url_catalogo & "<br>"
						else
							'conversione fallita - manda avviso e scrive log
							CALL SendEmailSupportEX("Errore: "& Request.ServerVariables("SERVER_NAME") & " - conversione pdf to ISSUU", _
													"Errore nella conversione del file: " & doc_file & vbCrLF & _
													"Risultato conversione:" & vbCrLF & _
													tmpResponse & vbCrLF & _
													"POST Inviato:" & vbCrLf & _
													postData)
						end if
					
					end if
					
					if cString(url_catalogo) <> "" then
						sql = " UPDATE mtb_documenti SET doc_url_catalogo_"&Application("LINGUE")(i)&" = '"&url_catalogo&"', " & _
							  " doc_data_modifica_catalogo_"&Application("LINGUE")(i)&" = "&SQL_date(conn, Now())&" WHERE doc_id = " & ID
						'response.write "update: "&sql & "<br>"
						'response.end
						conn.Execute(sql)
						url_catalogo = ""
					end if
					
					'Response.Write "<pre>"
					'response.write "POST:<br>" + postData + "<br>"
					'response.write "RESPONSE:<br>" + Server.HtmlEncode(tmpResponse) + "<br>"
					'Response.Write httpRequest.Status & " " & httpRequest.StatusText & vbCrLf
					'Response.Write "GetAllResponseHeaders:" + vbcrlf + httpRequest.GetAllResponseHeaders & vbCrLf & vbCrLf & vbCrLf
					'Response.Write "</pre>"
										
				
					'response.end
				end if
			
			
'			execute_path = ""
'			url_catalogo = ""
'			url = "http://"&Request.ServerVariables("SERVER_NAME") & "/amministrazione2/nextMemo2/CreaCatalogo.aspx"
'			doc_file = cString(request("tft_doc_file_"&Application("LINGUE")(i)))
'			
'			if doc_file <> "" AND File_Extension(doc_file) = FILE_TYPE_PDF then
'			
'				'controllo data file pdf da convertire con data catalogo scritta su db
'				file_path = application("IMAGE_PATH")&application("AZ_ID")&"\images\"&doc_file
'				file_data = DateTimeISO(File_Date(file_path))
'				if cString(request("doc_data_modifica_catalogo_"&Application("LINGUE")(i))) <> "" then
'					catalogo_last_mod = DateTimeIso(cString(request("doc_data_modifica_catalogo_"&Application("LINGUE")(i))))
'				else
'					catalogo_last_mod = file_data
'				end if
'				'response.write "file_data:"&file_data & "<br>catalogo_last_mod:" & catalogo_last_mod & "<br>"
'				'response.write file_data > catalogo_last_mod
'
'				if i > 0 then
'					for x=lbound(Application("LINGUE")) to i
'						if doc_file = cString(request("tft_doc_file_"&Application("LINGUE")(x))) AND x<>i then
'							sql = "SELECT doc_url_catalogo_"&Application("LINGUE")(x)&" FROM mtb_documenti WHERE doc_id = " & ID
'							url_catalogo = cString(GetValueList(conn, NULL, sql))
'						end if
'					next
'				end if
'				
'				
'				if url_catalogo = "" AND (file_data >= catalogo_last_mod OR cString(request("old_doc_file_"&Application("LINGUE")(i))) <> doc_file) then
'					
'					execute_path = "?FILEPATH=\"&replace(doc_file,"/","\") & "&LINGUA="&Application("LINGUE")(i)&"&ID="&ID
'					url = url & execute_path 
'					
'					'response.write "eseguo URL: " & url & "<br>"
'					CALL WriteLogAdmin(conn, "mtb_documenti", ID, "GENERAZIONE_CAT_SFOGLIABILE","Url Generazione: " + url)
'					
'					url_catalogo = ExecuteHttpUrl(url)
'					
'					url_catalogo = replace(url_catalogo," ","")
'					url_catalogo = Trim(url_catalogo)
'					url_catalogo = Left(url_catalogo,inStr(url_catalogo, ".html")+4)
'					url_catalogo = ParseSQL(url_catalogo,adChar)
'					'response.write "restituisce: -" & url_catalogo & "--fine--<br><br>"
'				end if
'				
'				if cString(url_catalogo) <> "" then
'					sql = " UPDATE mtb_documenti SET doc_url_catal	ogo_"&Application("LINGUE")(i)&" = '"&url_catalogo&"', " & _
'						  " doc_data_modifica_catalogo_"&Application("LINGUE")(i)&" = "&SQL_date(conn, Now())&" WHERE doc_id = " & ID
'					'response.write "update: "&sql & "<br>"
'					'response.end
'					conn.Execute(sql)
'					url_catalogo = ""
'				end if
'				
			else
				sql = " UPDATE mtb_documenti SET doc_url_catalogo_"&Application("LINGUE")(i)&" = '', " & _
					  " doc_data_modifica_catalogo_"&Application("LINGUE")(i)&" = NULL WHERE doc_id = " & ID
				conn.Execute(sql)
			end if
		next
	else
		for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			sql = " UPDATE mtb_documenti SET doc_url_catalogo_"&Application("LINGUE")(i)&" = '', " & _
				  " doc_data_modifica_catalogo_"&Application("LINGUE")(i)&" = NULL WHERE doc_id = " & ID
			conn.Execute(sql)
		next
	end if
	
	'gestione caratteristiche
	CALL DesSalva(conn, ID, "mrel_doc_ctech", "rdc_valore_", "rdc_doc_id", "rdc_ctech_id")

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if cString(request("chk_doc_catalogo_sfogliabile"))<>"" then
		Classe.Next_Page = "DocumentiMod.asp?ID="&ID
	else
		Classe.Next_Page = "Documenti.asp"
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
	

	
'compone le stringhe per il post e la stringa per la firma del post
sub AddIssuuPost(postField, postValue, encode, byref postData, byref signatureData)
	
	if cString(postData)<> "" then
		postData = postData + "&"
	end if
	
	postData = postData + postField + "=" + IIF(Encode, Server.UrlEncode(postValue), postValue)
	signatureData = signatureData + postField + postValue
	
end sub

%>














