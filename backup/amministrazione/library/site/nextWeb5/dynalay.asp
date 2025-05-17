<!--#INCLUDE FILE="page_tools.asp"-->
<% 
'**************************************************************************************************************************************
'VERIFICHE INIZIALI DEI PARAMETRI
'**************************************************************************************************************************************

'****************************************************
'verifica validita' della richiesta HTTP
if not Security_RequestIsValid() then
	
	'richiesta non valida:
	if Session("AZ_ERRORE_PS_ID")>0 then
		'rimanda alla pagina di errore
		response.redirect "dynalay.asp?PAGINA=" & session("PAGINE")(Session("AZ_ERRORE_PS_ID"))
	else
		'rimanda all'home page.
		response.redirect "default.asp"
	end if
end if
'****************************************************

'cambio della lingua (se necessario)
if request.Querystring("LINGUA")<>"" then
	CALL Applicazione_RESET(request.Querystring("LINGUA"))
end if

dim conn, rs_pagina, rs_layers, rs_paginasito, sql
dim pagina, i, META_aggiuntivi


' ***************************************
' Per vedere i bordi degli oggetti
' ***************************************
dim debug
if not debug then
	debug = false
end if


' ***************************************
' Anteprima da area amministrativa
' ***************************************
dim isPreviewAmministrazione
if instr(1, request.servervariables("SCRIPT_NAME"), "amministrazione/NextWeb", vbTextCompare)>0 then
	isPreviewAmministrazione = true
else
	isPreviewAmministrazione = false
end if

set conn = server.createObject("ADODB.Connection")
set rs_pagina = server.createObject("ADODB.recordset")
set rs_paginasito = server.createObject("ADODB.recordset")
set rs_layers = server.createObject("ADODB.recordset")
conn.Open Application("l_conn_ConnectionString"),"",""

'verifica stato applicazione
CALL Applicatione_CHECK(conn, rs_pagina)

'id della pagina
if CIntero(request.querystring("pagina")) > 0 then
	pagina = cInteger(request.querystring("pagina"))
elseif CIntero(request.querystring("PS")) > 0 then
	pagina = session("PAGINE")(CIntero(request.querystring("PS")))
else
	pagina = session("PAGINE")(session("AZ_HOME_PS_ID"))
end if

'in area amministrativa bypassa controlli congruenza
if not isPreviewAmministrazione then
	'tipo di pagina (ASPX o ASP)
	sql = " SELECT COUNT(*) FROM tb_objects o"& _
		  " INNER JOIN tb_layers l ON o.id_objects = l.id_objects"& _
		  " WHERE (identif_objects LIKE '%.ascx' OR obj_type = 'class')"& _
		  " AND (id_pag = "& cIntero(pagina) & _
		  "		 OR id_pag = (SELECT id_template FROM tb_pages WHERE id_page = "& cIntero(pagina) &"))"
	if CIntero(GetValueList(conn, rs_pagina, sql)) > 0 then
		response.redirect "default.aspx?PAGINA="& pagina
	end if
	
	'stato del sito
	session("AZ_AGGIORNAMENTO") = GetValueList(conn, rs_paginaSito, "SELECT sito_in_aggiornamento FROM tb_webs WHERE id_webs = "& cIntero(session("AZ_ID")))
	if Session("AZ_AGGIORNAMENTO") then
		if CIntero(session("AZ_AGGIORNAMENTO_PS_ID")) > 0 then
			pagina = session("PAGINE")(session("AZ_AGGIORNAMENTO_PS_ID"))
		else
			pagina = session("PAGINE")(session("AZ_HOME_PS_ID"))
		end if
	elseif Session("AZ_COSTRUZIONE") AND CIntero(Session("AZ_COSTRUZIONE_PS_ID")) > 0 then
		pagina = session("PAGINE")(session("AZ_COSTRUZIONE_PS_ID"))
	end if
end if

'**************************************************************************************************************************************
'APERTURA RECORDSET PAGINA
'**************************************************************************************************************************************
sql = " SELECT tb_pages.*, tb_webs.*, (tb_pages.id_webs) AS ID_SITO, ps.*, " + _
	  " (tb_templates.sfondoColore) AS TEMPLATE_COLORE, (tb_templates.SfondoImmagine) AS TEMPLATE_IMMAGINE, " + _
	  " tb_templates.semplificata AS TEMPLATE_SEMPLIFICATO," + _
	  " (SELECT MAX(x + largo) FROM tb_layers WHERE tb_layers.id_pag=tb_pages.id_page OR tb_layers.id_pag=tb_pages.id_template) AS PAGE_MAX_WIDTH " + _
	  " FROM ((tb_pages INNER JOIN tb_webs ON tb_pages.id_webs=tb_webs.id_webs) " + _
	  " LEFT JOIN tb_pagineSito ps ON (tb_pages.id_page = ps.id_pagDyn_it OR tb_pages.id_page = ps.id_pagDyn_en OR tb_pages.id_page = ps.id_pagDyn_fr OR tb_pages.id_page = ps.id_pagDyn_es OR tb_pages.id_page = ps.id_pagDyn_de))" + _
	  " LEFT JOIN tb_pages tb_templates ON tb_pages.id_template=tb_templates.id_page " + _
	  " WHERE tb_pages.id_page="
rs_pagina.open sql & cIntero(pagina), conn, adOpenForwardOnly, adLockOptimistic, adCmdText

if rs_pagina.eof then														'errore
	if CIntero(Session("AZ_ERRORE_PS_ID")) > 0 then
		rs_pagina.close()
		rs_pagina.open sql & cIntero(session("PAGINE")(Session("AZ_ERRORE_PS_ID"))), conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		pagina = session("PAGINE")(Session("AZ_ERRORE_PS_ID"))
	end if
elseif not isPreviewAmministrazione then
	if rs_pagina("riservata") AND Session("LOGIN_4_LOG")="" then			'non autenticato
		if CIntero(Session("AZ_LOGIN_PS_ID")) > 0 then
			rs_pagina.close()
			rs_pagina.open sql & cIntero(session("PAGINE")(Session("AZ_LOGIN_PS_ID"))), conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			pagina = session("PAGINE")(Session("AZ_LOGIN_PS_ID"))
		else
			pagina = 0
		end if
	end if
end if

if not rs_pagina.eof AND pagina > 0 then
	if not isPreviewAmministrazione then
		'**************************************************************************************************************************************
		'VERIFICA COERENZA PARAMETRI PAGINA CON SESSIONE SE NON IN PREVIEW DI AREA AMMINISTRATIVA
		'**************************************************************************************************************************************
		'verifica id del sito corrente
		if cInteger(Session("AZ_ID")) <> cInteger(rs_pagina("ID_SITO")) then
			'sito della pagina richiesta diverso da quello corrente: re-inizializza l'applicazione per il nuovo sito
			response.redirect GetPageURL(conn, request("PAGINA"))
		end if
		
		'verifica lingua della pagina
		if cString(rs_pagina("lingua"))<>"" then
			if lcase(rs_pagina("lingua")) <> lCase(Session("lingua")) then
				'lingua della pagina diversa da quella corrente: re-inizializza l'applicazione per la nuova lingua
				CALL Applicazione_RESET(rs_pagina("lingua"))	'resetta lingua e array pagine
				CALL Applicatione_CHECK(conn, rs_paginasito)		'reinizializza l'applicazione per la nuova lingua
			end if
		end if
	end if
else
	'**************************************************************************************************************************************
	'ERRORE NELLA PAGINA O NEI PARAMETRI
	'**************************************************************************************************************************************
	%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html>
	<head>
		<title>ERRORE NELL'APPLICAZIONE - APPLICATION ERROR</title>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<meta name="robots" content="noindex,nofollow">
	</head>
	<body>
		<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr>
				<td>
					ERRORE NELL'APPLICAZIONE<br>
					Ritornare alla <a href="default.asp">HOME PAGE</a> e riprovare.<br>
					<br>
					Se il problema persiste segnalarlo al webmaster.
				</td>
				<td>
					APPLICATION ERROR<br>
					go to <a href="default.asp">HOME PAGE</a> and then retry your request.<br><br>
					
					If the problem persists, advise the webmaster.
				</td>
			</tr>
		</table>
	</body>
</html>
	<%response.end
end if

'**************************************************************************************************************************************
'**************************************************************************************************************************************
'se l'esecuzione arriva qui vuol dire che la pagina e' correttamente aperta
'**************************************************************************************************************************************
'**************************************************************************************************************************************
dim template, sfondo_immagine, sfondo_colore, path_immagini, lingua, path_upload
dim title, keywords, description, author, style, DivTitle, DivStyle

'**************************************************************************************************************************************
'gestisce impostazione di base della pagina
'**************************************************************************************************************************************
lingua = Session("lingua")
'salva eventuale id template
template = cInteger(rs_pagina("id_template"))
	
if template > 0 then
	sfondo_immagine = cString(rs_pagina("template_immagine"))
	sfondo_colore = cString(rs_pagina("template_colore"))
else
	sfondo_immagine = cString(rs_pagina("SfondoImmagine"))
	sfondo_colore = cString(rs_pagina("sfondoColore"))
end if

'compone path per le immagini
if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
	'gestisce inidirizzo SSL
	path_immagini = "https://" 
	if Application("SECURE_IMAGE_SERVER")<>"" then
		path_immagini = path_immagini & Application("SECURE_IMAGE_SERVER")
	else
		path_immagini = path_immagini & Application("IMAGE_SERVER")
	end if
	if CString(session("SERVER_NAME")) = "" then
		session("SERVER_NAME") = "https://"& Application("SERVER_NAME")
	end if
else
	'indirizzo semplice
	path_immagini = "http://" & Application("IMAGE_SERVER")
	if CString(session("SERVER_NAME")) = "" then
		session("SERVER_NAME") = "http://"& Application("SERVER_NAME")
	end if
end if
path_upload = path_immagini
path_immagini = path_immagini & "/" & Session("AZ_ID")  & "/images/"

'imposta stile del tag body
style = ""
if sfondo_colore <> "" then
	style = style & "background-color:" & sfondo_colore & ";"
end if
if sfondo_immagine <> "" AND replace(sfondo_immagine, "nosfondo", "")<>"" then
	style = style & "background-image: url(" & path_immagini & sfondo_immagine & ");"
end if
if style <> "" then
	style = " style=""" & style & """ "
end if

'**************************************************************************************************************************************
'recupera dati della paginasito corrispondente per recuperare nomi pagina e metatag
'**************************************************************************************************************************************
sql = " SELECT * FROM tb_paginesito WHERE id_web=" & cIntero(rs_pagina("ID_SITO")) & _
	  " AND (id_pagDyn_" & lingua & "=" & cIntero(pagina) & " OR id_pagStage_" & lingua & "=" & cIntero(pagina) & ") "
rs_paginasito.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdtext

'**************************************************************************************************************************************
Session("CURRENT_PAGINA") = rs_pagina("id_page")
if not rs_paginasito.eof then
	Session("CURRENT_PAGINASITO") = rs_paginasito("id_pagineSito")
end if
dim templateSemplificato
templateSemplificato = CBoolean(rs_pagina("semplificata") OR rs_pagina("TEMPLATE_SEMPLIFICATO"), false) OR session("SITO_MOBILE")
'**************************************************************************************************************************************

'gestione del titolo 
title = rs_pagina("nomepage")
if cString(rs_pagina("titolo_" & lingua))<>"" then
	title = title & " - " & rs_pagina("titolo_" & lingua)
end if

'gestione keywords e description della pagina
keywords = ""
description = ""
if not rs_paginasito.eof then
	keywords = cString(rs_paginasito("PAGE_keywords_" & lingua))
	description = cString(rs_paginasito("PAGE_description_" & lingua))
end if
if keywords = "" then
	'keywords della pagina non presenti: mette quelle generali del sito
	keywords = rs_pagina("META_keywords_" & lingua)
end if

if description = "" then
	'description della pagina non presente: mette quella generale del sito
	description = rs_pagina("META_description_" & lingua)
end if
'Evito la memorizzazione in cache/proxy della pagina
%>
 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html>
	<head>
		<title><%= title %></title>
		<meta name="description" content="<%= description %>">
		<meta name="keywords" content="<%= keywords %>">
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<meta name="Language" content="<%= lingua %>">
		<meta name="author" content="<%= rs_pagina("META_Author") %>">
		<meta name="generator" content="NEXT-web 5 by NEXT-AIM">
		<meta name="copyright" content="&copy;<%= Year(Date) %> - NEXT-AIM">
		<meta name="distribution" content="Global">
		<%= META_aggiuntivi %>
		<% sql = "SELECT MAX(x + largo) " %>
		<style type="text/css" media="screen">
			div.dynalay_container{
				position:relative; 
				margin:0px; 
				padding:0px;
				width:<%= rs_pagina("PAGE_MAX_WIDTH") %>px;
			}
			div.dynamail {display:none;visibility:hidden}
		</style>
<% 	'caricamento stili
	dim fso, CssPath, cssDir, file, CssFile
	set fso = Server.CreateObject("scripting.filesystemobject")
	CssPath = request.ServerVariables("APPL_PHYSICAL_PATH") & "App_Themes/Default"
	if fso.FolderExists(CssPath) then
		set cssDir = fso.GetFolder(cssPath)
		for each file in cssDir.Files
			if LCase(Right(file.name, 4)) = ".css" then
			
				if request.querystring("HTML_FOR_EMAIL")<>"" then 		'pagina generata per email
					set CssFile = fso.OpenTextFile(file.path, 1, false) %>
		<style type="text/css">
			<%= CssFile.ReadAll %>
		</style>
<%					CssFile.close
					set CssFile = nothing %>
<%				else													'pagina con stili in link 
%>
		<link rel="stylesheet" type="text/css" href="<%= session("SERVER_NAME") &"/App_Themes/Default/"& file.name %>">			
<%				end if

			end if
		next
		set cssDir = nothing
	end if
	set fso = nothing
	
	
	'caricamento javascript
	if request.querystring("HTML_FOR_EMAIL")<>"" AND not isPreviewAmministrazione then %>
		<SCRIPT LANGUAGE="javascript" type="text/javascript">
		<!--#INCLUDE FILE="../../tools_JS.asp" -->
		</SCRIPT>
<%	else %>
		<link rel="stylesheet" type="text/css" href="<%= session("SERVER_NAME") %>/amministrazione/library/site/nextweb5/standard.css" />
		<link rel="stylesheet" type="text/css" href="<%= path_upload &"/"& session("AZ_ID") %>/css/stili_testo.css" />
		<SCRIPT LANGUAGE="javascript" src="<%= IIF(isPreviewAmministrazione, "..", "amministrazione")%>/library/Tools_JS.asp" type="text/javascript"></SCRIPT>
		<script type="text/javascript">
		// <![CDATA[
		   var baseURL = "<%= Session("SERVER_NAME") %>/";
		   var imageURL = "<%= path_immagini %>";
		   var imageAlt = "Doppio click per chiudere la finestra.";
		// ]]>
		</script>
		<script src="<%= IIF(isPreviewAmministrazione, "..", "amministrazione")%>/library/Utils4Dynalay.js" type="text/javascript"></script><script type="text/vbscript">
		// <![CDATA[
		     Function VBGetSwfVer(i)
		         on error resume next
		         Dim swControl, swVersion
		         swVersion = 0
		         set swControl = CreateObject("ShockwaveFlash.ShockwaveFlash." + CStr(i))
		         if (IsObject(swControl)) then
		             swVersion = swControl.GetVariable("$version")
		         end if
		         VBGetSwfVer = swVersion
		     End Function
		// ]]>
		</script>
<% 	end if %>
	</head>
	<body <%= style %> dir="ltr" lang="<%= lingua %>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0">
		<form method="post" action="" id="NextForm">
		<div class="dynalay_container">
		<% if templateSemplificato then
				dim url
				url = completeURL() 
				response.write "<div class=""dynamail"">"
				response.write ChooseValueByAllLanguages(Session("LINGUA"), _
												"Se il messaggio non viene visualizzato in modo corretto segui <a href="""+ url +""">questo link</a>", _
												"follow <a href="""+ url +""">this link</a> if you do not see the message correctly", _
												"follow <a href="""+ url +""">this link</a> if you do not see the message correctly", _
												"follow <a href="""+ url +""">this link</a> if you do not see the message correctly", _
												"follow <a href="""+ url +""">this link</a> if you do not see the message correctly", _
												"follow <a href="""+ url +""">this link</a> if you do not see the message correctly", _
												"follow <a href="""+ url +""">this link</a> if you do not see the message correctly", _
												"follow <a href="""+ url +""">this link</a> if you do not see the message correctly") 
				response.write "</div>"
			end if %>
		<%'**************************************************************************************************************************************
		'ciclo di generazione dei layers
		'**************************************************************************************************************************************
		 'query per recuperare layers del template seguiti dai layers della pagina
		sql = " SELECT * FROM tb_layers l"& _
		 	  " LEFT JOIN tb_objects o ON l.id_objects = o.id_objects"& _
			  " WHERE (id_pag=" & cIntero(pagina) & " OR id_pag=" & cIntero(template) & ") AND "& SQL_IsTrue(conn, "visibile")
		if templateSemplificato then
			sql = sql + " ORDER BY y, z_order "
		else
			sql = Sql + " ORDER BY tipo_contenuto DESC, z_order "
		end if
		rs_layers.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
		
		'calcola l'eventuale allineamento del contenuto
		dim isContenuto, alignLeft, alignTop, f
		isContenuto = CString(request.querystring("HTML_FOR_CONTENTS")) <> ""
		if isContenuto then
			alignLeft = 0
			alignTop = 0
			while not rs_layers.eof
				if UCase(CString(rs_layers("tipo_contenuto"))) = "C" then
	                f = CReal(Replace(rs_layers("em_x"), ".", ","))
	                if (alignLeft > f OR alignLeft = 0) then
	                    alignLeft = f
					end if
	                f = CReal(Replace(rs_layers("em_y"), ".", ","))
	                if (alignTop > f OR alignTop = 0) then
	                    alignTop = f
					end if
				end if
				rs_layers.movenext
			wend
			isContenuto = alignTop <> 0
		end if
		
		'ciclo generazione layers
		dim l, t, cssClass
		if rs_layers.recordCount > 0 then
			rs_layers.MoveFirst
		end if
		while not rs_layers.eof
			if NOT isContenuto OR UCase(CString(rs_layers("tipo_contenuto"))) = "C" then
				if isContenuto then
					l = CReal(Replace(rs_layers("em_x"), ".", ",")) - alignLeft
					t = CReal(Replace(rs_layers("em_y"), ".", ",")) - alignTop
				else
					l = rs_layers("em_x")
					t = rs_layers("em_y")
				end if
				
				if NOT templateSemplificato then
					 DivStyle = "position: absolute; " + _
					 			"width: " & rs_layers("em_largo") & "em; " + _
								"height: " & rs_layers("em_alto") & "em; " + _
								"z-index: " & rs_layers("z_order") & "; " + _
								"left: " & Replace(l, ",", ".") & "em; " + _
								"top: " & Replace(t, ",", ".") & "em;"
				else
					DivStyle = ""
				end if
				
				if IsPreviewAmministrazione then
					DivTitle = "........................................................................" & vbCrLF + _
               				   "left:" & rs_layers("em_x") & "em ( " & rs_layers("x") & "px ); " + VBcrlf + _
							   "top:" & rs_layers("em_y") & "em ( " & rs_layers("y") & "px ); " + VBcrlf + _
							   "width:" & rs_layers("em_largo") & "em ( " & rs_layers("largo") & "px ); " + VBcrlf + _
							   "height:" & rs_layers("em_alto") & "em ( " & rs_layers("alto") & "px ); " + VBcrlf + _
							   "........................................................................"
				end if
				
				SELECT CASE rs_layers("id_tipo")
					CASE 1												'testo normale
						cssClass = "layers_text"
					CASE 2												'immagine
						cssClass = "layers_image"
					CASE 3												'flash
						cssClass = "layers_flash"
					CASE 4												'oggetto
						cssClass = "layers_object "& LCase(rs_layers("name_objects"))
						if IsPreviewAmministrazione then
							DivTitle = DivTitle +_
									   "........................................................................" & vbCrLF + _
	                       			   "OGGETTO:" & rs_layers("name_objects") & ";" & VBcrlf & _
									   "ASPCODE:" & rs_layers("aspcode") & ";" & VBcrlf & _
									   "PROPRIETA:" & VBcrlf & _
									   rs_layers("testo") & VBcrlf & _
									   "........................................................................"
							DivStyle = DivStyle + _
									   "border:1px dashed #000;" + _
									   "background:#CCC;" + _
									   "layer-background-color:#CCC;" + _
									   "-moz-opacity:70%;" + _
									   "filter:Alpha(Opacity=70);"
						end if
					CASE 5												'testo strutturato
						cssClass = "layers_text_s"
				END SELECT 
				%>
<div id="lay_<%= rs_layers("id_lay") %>" class="<%= cssClass %>" <% if DivStyle<>"" then %> style="<%= DivStyle %>" <% end if %> <% if DivTitle<>"" then %> title="<%= DivTitle %>" <% end if %>>
				<% 'gestione tipo layer
				if rs_layers("id_tipo") = 4 then
					CALL MakeObject(l, t)
				else
					response.write Replace(Replace(rs_layers("html"), "<@PATH_RES>", path_upload &"/"& session("AZ_ID")), _
										   "<@PATH>", session("SERVER_NAME") & "\dynalay.asp")
				end if
				
				response.write "</div>"
			end if
			
			rs_layers.MoveNext
		wend
		
		rs_layers.close %>
		</div>
		</form>
	</body>
</html>
<%
rs_paginasito.close
rs_pagina.close


'**************************************************************************************************************************************
'log visite pagina
'**************************************************************************************************************************************
CALL LogVisit_Page(conn)

conn.close
set rs_pagina = nothing
set rs_paginasito = nothing
set rs_layers = nothing


'**************************************************************************************************************************************
'funzioni di gestione dei layers
'**************************************************************************************************************************************

'******************************************************************************
'LAYER OGGETTO
sub MakeObject(l, t)
	dim confstr, obj
	confstr = replace(rs_layers("testo"), vbCRLF, "")
	
	if instr(1, Request.ServerVariables("SERVER_NAME"), ".local", vbTextCompare)>0 OR isPreviewAmministrazione then %>
		<!-- 
			OGGETTO: <%= rs_layers("nome") %>
			ASPCODE: <%= rs_layers("aspcode") %>
			TESTO:........................................................................
			<%= rs_layers("testo") %>
			CONFSTR:......................................................................
			<%= confstr %>
		-->
<%	end if
	
	if isPreviewAmministrazione then
		'dynalay in area amministrativa: genera il segnaposto
		if Session("LOGIN_4_LOG") = "NEXTAIM" then %>
        	<div style="width:100%; height:100%; overflow:hidden; font-size:11px;">
            	<%= TextEncode(DivTitle) %> 
	        </div>
    	<% else %>
			<!--
			<%= TextEncode(DivTitle) %>
			-->
		<% end if
	else
		'dynalay in area pubblica: genera il plugin
		if debug then
			response.write "<div style=""border:1px dotted red"">"
		end if
		'creazione oggetto con pagina ASP eseguita
		Session("LAYER_ID") = rs_layers("id_lay")
		Session("LAYER_NAME") = rs_layers("nome")
		Session("LAYER_LEFT") = l
		Session("LAYER_TOP") = t
		Session("LAYER_WIDTH") = rs_layers("largo")
		Session("LAYER_HEIGHT") = rs_layers("alto")
		Session("LAYER_Z_ORDER") = rs_layers.absoluteposition
		Session("CONFSTR") = confstr
		if instr(1,rs_layers("aspcode"),"amministrazione/",vbTextCompare)>0 then
			Server.Execute(rs_layers("aspcode") )
		else
			Server.Execute("plugin/" & rs_layers("aspcode") )
		end if
		if debug then
			response.write "</div>"
			response.write "<div style=""font-size:9px;color:red"">"+ rs_layers("aspcode") +"</div>"
		end if
	end if
end sub


function CompleteURL()
	dim url
	url = "http"
	if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
		url = url+"s"
	end if
	url = url+ "://" + request.ServerVariables("HTTP_HOST") + request.ServerVariables("SCRIPT_NAME") + "?" + request.ServerVariables("QUERY_STRING")
	CompleteURL = url
end function
%>