<!--#INCLUDE FILE="page_tools.asp"-->
<% 
'**************************************************************************************************************************************
'VERIFICHE INIZIALI DEI PARAMETRI
'**************************************************************************************************************************************

'****************************************************
'verifica validita' della richiesta HTTP
if not Security_RequestIsValid() then
	'richiesta non valida: rimanda all'home page.
	response.redirect "default.asp"
end if
'****************************************************

'cambio della lingua (se necessario)
if request.Querystring("LINGUA")<>"" then
	CALL Applicazione_RESET(request.Querystring("LINGUA"))
end if


dim conn, rs_pagina, rs_layers, rs_paginasito, sql
dim pagina, QueryString


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
if instr(1, request.servervariables("SCRIPT_NAME"), "amministrazione", vbTextCompare)>0 then
	isPreviewAmministrazione = true
else
	isPreviewAmministrazione = false
end if


dim META_aggiuntivi


set conn = server.createObject("ADODB.Connection")
set rs_pagina = server.createObject("ADODB.recordset")
set rs_paginasito = server.createObject("ADODB.recordset")
set rs_layers = server.createObject("ADODB.recordset")
conn.Open Application("l_conn_ConnectionString"),"",""

if request.querystring("pagina")<>"" AND IsNumeric(request.querystring("pagina")) then
	'verifica stato applicazione
	CALL Applicatione_CHECK(conn, rs_pagina)
	
	'recupera pagina
	pagina = cIntero(request.querystring("pagina"))
	
elseif request("PS")<>"" AND isNumeric(request("PS")) then
	'manda alla pagina indicata dalla conversione del parametro 
	pagina = DecodePaginaSito(conn, rs_pagina, cLng(request("PS")))
	QueryString = replace(request.ServerVariables("QUERY_STRING"), "PS=" & request("PS"), "")
	response.redirect "dynalay.asp?PAGINA=" & pagina & QueryString
	
else
	'nessun parametro che indichi una pagina: carica l'home page.
	pagina = Applicazione_INIT(conn, rs_pagina)
	pagina = cIntero(Session("PAGINE")(pagina))
end if


'**************************************************************************************************************************************
'APERTURA RECORDSET PAGINA
'**************************************************************************************************************************************
sql = " SELECT tb_pages.*, tb_webs.*, (tb_pages.id_webs) AS ID_SITO, " + _
	  " (tb_templates.sfondoColore) AS TEMPLATE_COLORE, (tb_templates.SfondoImmagine) AS TEMPLATE_IMMAGINE, " + _
	  " tb_templates.semplificata AS TEMPLATE_SEMPLIFICATO," + _
	  " (SELECT MAX(x + largo) FROM tb_layers WHERE tb_layers.id_pag=tb_pages.id_page OR tb_layers.id_pag=tb_pages.id_template) AS PAGE_MAX_WIDTH " + _
	  " FROM (tb_pages INNER JOIN tb_webs ON tb_pages.id_webs=tb_webs.id_webs) " + _
	  " LEFT JOIN tb_pages tb_templates ON tb_pages.id_template=tb_templates.id_page " + _
	  " WHERE tb_pages.id_page=" & cIntero(pagina)
rs_pagina.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

if not rs_pagina.eof then
	if not isPreviewAmministrazione then

		'**************************************************************************************************************************************
		'VERIFICA COERENZA PARAMETRI PAGINA CON SESSIONE
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
	'**************************************************************************************************************************************%>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
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
dim template, sfondo_immagine, sfondo_colore, path_immagini, lingua
dim title, keywords, description, style

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
	'gestisce inidirizzo SSL per le immagini
	path_immagini = "https://" 
	if Application("SECURE_IMAGE_SERVER")<>"" then
		path_immagini = path_immagini & Application("SECURE_IMAGE_SERVER")
	else
		path_immagini = path_immagini & Application("IMAGE_SERVER")
	end if
else
	'indirizzo semplice per le immagini
	path_immagini = "http://" & Application("IMAGE_SERVER")
end if
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
	  " AND (id_pagDyn_" & lingua & "=" & pagina & " OR id_pagStage_" & lingua & "=" & pagina & ") "
rs_paginasito.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdtext


'**************************************************************************************************************************************
Session("CURRENT_PAGINA") = rs_pagina("id_page")
if not rs_paginasito.eof then
	Session("CURRENT_PAGINASITO") = rs_paginasito("id_pagineSito")
end if
dim templateSemplificato
templateSemplificato = CBoolean(rs_pagina("semplificata") OR rs_pagina("TEMPLATE_SEMPLIFICATO"), false)
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
 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title><%= title %></title>
		<meta name="description" content="<%= description %>">
		<meta name="keywords" content="<%= keywords %>">
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<meta name="Language" content="<%= lingua %>">
		<meta name="author" content="<%= rs_pagina("META_Author") %>">
		<meta name="generator" content="NEXT-web 4 by NEXT-AIM">
		<meta name="copyright" content="&copy;<%= Year(Date) %> - NEXT-AIM">
		<% if not rs_paginasito.eof then
			if cString(rs_pagina("google_webmaster_tools_verify_code"))<>"" AND _
				  cIntero(rs_pagina("id_home_page")) =  cIntero(rs_paginasito("id_paginesito")) then %>
				<meta name="verify-v1" content="<%= rs_pagina("google_webmaster_tools_verify_code") %>" />
			<% end if
		end if%>
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
			div.dynamail {
				display:none;
				visibility:hidden
			}
		</style>
	<% if request.querystring("HTML_FOR_EMAIL")<>"" then 
		'pagina generata per email
		CALL WriteStili()
		
		'commentato il 25/06/2012 da Nicola: nelle email il javascript è inutile
		'<SCRIPT LANGUAGE="javascript" type="text/javascript">
		'	<!--#INCLUDE FILE="../../tools_JS.asp" -->
		'</SCRIPT>
		 
	else %>
		<link rel="stylesheet" type="text/css" href="<%= IIF(isPreviewAmministrazione, "../../", "")%>stili.css">
		<SCRIPT LANGUAGE="javascript" src="<%= IIF(isPreviewAmministrazione, "..", "amministrazione")%>/library/Tools_JS.asp" type="text/javascript"></SCRIPT>
	<% end if %>
	</head>
	<body <%= style %> dir="ltr" lang="<%= lingua %>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" class="<%=IIF(request.querystring("HTML_FOR_EMAIL")<>"", "email", "page")%>"
	
	>
	<!--  -->
		<div class="dynalay_container">
		<% if request.querystring("HTML_FOR_EMAIL")<>"" then
			dim url
			url = completeURL() 
			response.write "<div class=""dynamail"">"
			response.write ChooseValueByAllLanguages(Session("LINGUA"),"Se il messaggio non viene visualizzato in modo corretto segui <a href="""+ url +""">questo link</a>", _
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
		sql = "SELECT * FROM tb_layers WHERE (id_pag=" & cIntero(pagina) & " OR id_pag=" & cIntero(template) & ") AND visibile=True " 
		if template < pagina then
			sql = sql & " ORDER BY id_pag, z_order "
		else
			sql = sql & " ORDER BY id_pag DESC, z_order "
		end if 
		rs_layers.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
		
		while not rs_layers.eof
			select case rs_layers("id_tipo")
				case 1
					CALL MakeText("layers_text", 		"testo_" & rs_layers.AbsolutePosition)
				case 2
					CALL MakeImage("layers_image", 		"immagine_" & rs_layers.AbsolutePosition)	
				case 4
					CALL MakeObject("layers_object",	rs_layers("nome"))
			end select
			Response.Write vbCRLF & vbCRLF
			
			rs_layers.MoveNext
		wend
		
		rs_layers.close %>
		</div>
		<% if cString(rs_pagina("google_analytics_code"))<>"" _
			  AND request.querystring("HTML_FOR_EMAIL")="" then %>
			<script type="text/javascript">
				// <![CDATA[
				var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
				document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
				// ]]>
			</script>
			<script type="text/javascript">
				// <![CDATA[
				try {
				   	var pageTracker = _gat._getTracker("<%= rs_pagina("google_analytics_code") %>");
				   	pageTracker._initData();
				   	pageTracker._trackPageview();
				} 
				catch (e) { /* javascript semi-disabilitato */ }
				// ]]>
			</script>
		<% end if %>
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

function DIV_Start_Tag(classe, l_name, isPlugin)
	dim DivStyle, DivTitle
	
	if request.querystring("HTML_FOR_EMAIL")<>"" then
		classe = classe + " " + l_name
	end if
	DIV_Start_Tag = "<div class=""" & classe & """ id=""" & l_name & """ "
	
	if NOT templateSemplificato then
		DivStyle = "position:absolute; left:" & rs_layers("x") & "px;" & _
				   " top:" & rs_layers("y") & "px;" & _
				   " width:" & rs_layers("largo") & +"px;" & _
				   " z-index:" & rs_layers.absoluteposition & "; " & _
				   " height:" & rs_layers("alto") & "px;"
	else
		DivStyle = " width: 100%;"
	end if
	 
	
	if IsPreviewAmministrazione then
		
		if IsPlugin then
			DivTitle = DivTitle +_
					   "........................................................................" & vbCrLF + _
	                   "OGGETTO:" & rs_layers("nome") & ";" & VBcrlf & _
					   "ASPCODE:" & rs_layers("aspcode") & ";" & VBcrlf & _
					   "PROPRIETA:" & VBcrlf & _
					   replace(rs_layers("testo"), ";", vbCrLf) & VBcrlf & _
					   "........................................................................"
					   
			DivStyle = DivStyle + _
					   "border:1px dashed #000;" + _
					   "background:#CCC;" + _
					   "layer-background-color:#CCC;" + _
					   "-moz-opacity:70%;" + _
					   "filter:Alpha(Opacity=70);"
		end if
		
	end if
	
	DIV_Start_Tag = DIV_Start_Tag + _
					" style=""" + DivStyle + """ "
	
	if DivTitle<>"" then
		DIV_Start_Tag = DIV_Start_Tag + _
						" title=""" + DivTitle + """ "
	end if
	
	DIV_Start_Tag = DIV_Start_Tag + ">"
	
	if isPreviewAmministrazione then
		'dynalay in area amministrativa: genera il segnaposto
		if Session("LOGIN_4_LOG") = "NEXTAIM" AND IsPlugin then
			DIV_Start_Tag = DIV_Start_Tag + vbCrLF + _	
							"<div style=""width:100%; height:100%; overflow:hidden; font-size:11px;"">" + vbCrLF + vbCrLF + _
							TextEncode(DivTitle) + vbCrLF + vbCrLF + _
							"</div>" + vbCrLF
		else
			DIV_Start_Tag = DIV_Start_Tag + vbCrLF + _	
							"<!--" + vbCrLF + _
							TextEncode(DivTitle) + vbCrLF + _
							"-->" + vbCrLF
		end if
	end if
end function


'******************************************************************************
'LAYER DI TESTO
sub MakeText(classe, lname)
	response.write DIV_Start_Tag( classe, lname, false ) &_
				   rs_layers("html") & _
				   "</div>"
end sub


'******************************************************************************
'LAYER OGGETTO
sub MakeObject(classe, lname)
	dim confstr, obj
	
	confstr = replace(rs_layers("testo"), vbCRLF, "")
	if instr(1, Request.ServerVariables("SERVER_NAME"), ".local", vbTextCompare)>0  OR isPreviewAmministrazione then %>
		<!-- 
			OGGETTO: <%= lname %>
			ASPCODE: <%= rs_layers("aspcode") %>
			TESTO:........................................................................
			<%= rs_layers("testo") %>
			CONFSTR:......................................................................
			<%= confstr %>
		-->
	<%end if
	
	response.write DIV_Start_Tag( classe, lname, true )
	if not isPreviewAmministrazione then
		if debug then
			response.write "<div style=""border:1px dotted red"">"
		end if
		if instr(1,rs_layers("aspcode"),".asp", vbTextCompare)>0 then 
			'creazione oggetto con pagina ASP eseguita
			Session("LAYER_ID") = rs_layers("id_lay")
			Session("LAYER_NAME") = lname
			Session("LAYER_LEFT") = rs_layers("x")
			Session("LAYER_TOP") = rs_layers("y")
			Session("LAYER_WIDTH") = rs_layers("largo")
			Session("LAYER_HEIGHT") = rs_layers("alto")
			Session("LAYER_Z_ORDER") = rs_layers.absoluteposition
			Session("CONFSTR") = confstr
			if instr(1,rs_layers("aspcode"),"amministrazione/",vbTextCompare)>0 then
				Server.Execute(rs_layers("aspcode") )
			else
				Server.Execute("plugin/" & rs_layers("aspcode") )
			end if
			
		else
			'creazione oggetto da oggetto COM
			set obj = Server.CreateObject(rs_layers("aspcode"))
			obj.Configura(confstr)
			obj.settaLayer(rs_layers("id_lay"))
			obj.disegna()
			set obj =  nothing
		end if
		if debug then
			response.write "</div>"
			response.write "<div style=""font-size:9px;color:red"">"+ rs_layers("aspcode") +"</div>"
		end if
	end if
	response.Write "</div>"
	
end sub


'******************************************************************************
'LAYER IMMAGINI
sub MakeImage(classe, lname)
	with response
		.write DIV_Start_Tag( classe, lname, false )
		if rs_layers("format")<>"[]" then
			'compone link su immagine
			.write "<A HREF="+extractLink(rs_layers("format"))
			if isBlank( rs_layers("format") ) then
				.write " target=""_blank"" "
			end if
			.write ">" 
		end if
		'scrive immagine
		.write "<img src=""" & path_immagini & rs_layers("nome") & """ border=""0"" alt="""">"
		if rs_layers("format")<>"[]" then
			.write "</a>"
		end if
		.write "</div>"
	end with
END SUB

'[[1, 1, "javascript:OpenImage('panorama_venice.jpg', '','')", 0]]
'[[1, 1, "http://www.turismovenezia.it", 1]]
function extractLink( format )
	dim QuoteBegin, QuoteEnd
	QuoteBegin = instr(1, format, """", vbTextCompare)
	QuoteEnd = instrrev(format, """", vbTrue, vbTextCompare) + 1
	extractLink = mid(format, QuoteBegin, (QuoteEnd - QuoteBegin))
end function


function isBlank(format)
	format = right(format, len(format) - instrrev(format, ",", vbTrue, vbTextCompare))
	isBlank = (instr(1, format, "1", vbTextCompare) > 0)
end function


function CompleteURL()
	dim var,url
	url = "http"
	if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
		url = url+"s"
	end if
	url = url+ "://" + request.ServerVariables("HTTP_HOST") + request.ServerVariables("SCRIPT_NAME") + "?" + request.ServerVariables("QUERY_STRING")
	CompleteURL = url
end function

%>