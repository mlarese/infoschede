<!--#INCLUDE FILE="../../tools.asp"-->
<!--#INCLUDE FILE="../site_tools.asp"-->
<%
Response.CodePage = 65001 
Response.CharSet = "utf-8"
Response.ContentType = "text/html"
Response.CacheControl = "no-cache"
Response.Expires = -1
Response.ExpiresAbsolute=MonthName(Month(Now())) & " " & Day(Now()) & "," & Year(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now())

'****************************************************
'verifica validita' della richiesta HTTP
if not Security_RequestIsValid() then
	'richiesta non valida: rimanda all'home page.
	response.redirect "default.asp"
end if
'****************************************************

'****************************************************
'cambio della lingua
if request.Querystring("LINGUA")<>"" then
	Session.Contents.Remove("PAGINE")
	CALL SetSessionLingua(request.Querystring("LINGUA"))
end if
'****************************************************


' ***************************************
' Per vedere i bordi degli oggetti
' ***************************************
dim debug
if not debug then
	debug = false
end if


'inizializza variabile per impostazione livello dynalay rispetto a root sito (sottodirectory)
'viene impostata nel file di inclusione sulla root del sito
dim PathSubLevels, PathExecute
if PathSubLevels<>"" then
    PathExecute = "/"
else
    PathExecute = ""
end if
dim conn, rs, sql, pag
dim sfondo, colore_sfondo, path_azienda
dim nome_pagina, id_template

set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.recordset") 
conn.Open Application("l_conn_ConnectionString"),"",""
	
pag = cInt("0" & request.querystring("PAGINA"))
if pag>=1 then
	'Verifica della lingua della pagina
	sql = "SELECT lingua FROM tb_pages WHERE id_page = " & pag
	rs.Open sql,conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if Session("LINGUA")<>rs(0) then
		Session.Contents.Remove("PAGINE")
		CALL SetSessionLingua(rs(0))
	end if
	rs.close
end if
' inizializza il vettore delle pagine attive
if IsEmpty(Session("PAGINE")) then
	Server.Execute(PathExecute + "amministrazione/library/site/nextweb3/SiteInit.asp")
end if





if pag < 1 then
	'pagina non indicata: cerca home page
	sql = "SELECT id_home_page FROM tb_webs WHERE id_webs=" & cIntero(Session("AZ_ID"))
	rs.Open sql,conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	pag = Session("PAGINE")(cInt(rs("id_home_page")))
	rs.close
end if




sql = "SELECT tb_pages.*, (tb_templates.sfondo) AS SFONDO_TEMPLATE, " + _
	  " tb_webs.titolo_" & Session("LINGUA") & ", tb_webs.google_webmaster_tools_verify_code,tb_webs.google_analytics_code,tb_webs.id_home_page, " &_
	  " (SELECT id_pagineSito FROM tb_paginesito WHERE id_web = tb_pages.id_webs AND (id_pagDyn_" & Session("LINGUA") & "=" & cIntero(pag) & " OR id_pagStage_" & Session("LINGUA") & "=" & cIntero(pag) & ")) AS PAGINASITO " + _
	  " FROM (tb_pages INNER JOIN tb_webs ON tb_pages.id_webs=tb_webs.id_webs) " &_
	  " LEFT JOIN tb_pages tb_templates ON tb_pages.id_template=tb_templates.id_page " & _
	  " WHERE tb_pages.id_page=" & cIntero(pag)

rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText



if not rs.eof then
	
	'**************************************************************************************************************************************
	Session("CURRENT_PAGINA") = rs("id_page")
	Session("CURRENT_PAGINASITO") = rs("PAGINASITO")
	'**************************************************************************************************************************************

	'salva eventuale id template
	id_template = cInt("0" & rs("id_template"))
		
	'imposta lo sfondo della pagina
	if ("" & rs("SFONDO_TEMPLATE"))<>"" then
		sfondo = rs("SFONDO_TEMPLATE") & ""		'sfondo del template
	else
		sfondo = rs("sfondo") & ""					'sfondo della pagina
	end if

else
	response.write "Pagina non trovata"
	Response.end
end if
	
'imposta colore di sfondo
if instr(1,sfondo,"#", vbTextCompare)>0 then
	colore_sfondo = sfondo
	sfondo = ""
else
	colore_sfondo = "#FFFFFF"
end if

nome_pagina = rs("nomepage")
if rs("titolo_" & Session("Lingua"))<>"" then
	nome_pagina = nome_pagina & ", " & rs("titolo_" & Session("Lingua"))
end if
	
	


'imposta http sicuro
if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
	path_azienda = "https://"
else
	path_azienda = "http://"
end if	
path_azienda = path_azienda & Application("IMAGE_SERVER")+"/" & Session("AZ_ID")  & "/"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>

<title><%= nome_pagina %></title>
<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
<%
	CALL DeterminaMetaTag( conn,Session("CURRENT_PAGINASITO"), Session("META_TAG_DEFINITIONS") )
%>
<META NAME="expires" CONTENT="none">
<META NAME="rating" CONTENT="General">
<META NAME="Language" CONTENT="<%= setLanguage(Session("LINGUA")) %>">
<META NAME="author" CONTENT="<%= Session("META_Author") %>">
<META NAME="copyright" CONTENT="Copyright &copy;<%= Year(Date) %> - <%= Session("META_Author") %>">
<META NAME="distribution" CONTENT="Global">
<META NAME="robots" CONTENT="INDEX,FOLLOW">
<META NAME="revisit-after" CONTENT="10 Days">
<META NAME="rating" CONTENT="General">
<% if not rs.eof then
			if cString(rs("google_webmaster_tools_verify_code"))<>"" AND _
				  cIntero(rs("id_home_page")) =  cIntero(rs("PAGINASITO")) then %>
				<meta name="verify-v1" content="<%= rs("google_webmaster_tools_verify_code") %>" />
			<% end if
	end if
	dim ga_code
	ga_code = rs("google_analytics_code")
	rs.close
%>
<% if request.querystring("HTML_FOR_EMAIL")<>"" then 
    'pagina generata per email
	CALL WriteStili()
else %>
		<link rel="stylesheet" type="text/css" href="<%= PathSubLevels %>stili.css">
<% end if %>
<SCRIPT LANGUAGE="javascript" src="<%= PathSubLevels %>amministrazione/library/Tools_JS.asp" type="text/javascript"></SCRIPT>

</head>
<body <% if colore_sfondo<>"" then %> bgcolor="<%=colore_sfondo %>"<% end if %> <% if sfondo <> "" then  %>background="http://<%= Application("IMAGE_SERVER")+"/" & Session("AZ_ID") %>/images/<%= sfondo %>"<% end if %>>
	<%	
	if id_template>0 then
		'query per recuperare layers del template seguiti dai layers della pagina
		sql = "SELECT * FROM tb_layers WHERE (id_pag=" & cIntero(pag) & " OR id_pag=" & cIntero(id_template) & ") AND visibile " 
		if cIntero(id_template) < cIntero(pag) then
			sql = sql & " ORDER BY id_pag, z_order "
		else
			sql = sql & " ORDER BY id_pag DESC, z_order "
		end if
	else
		'query per recuperare layers della pagina
		sql = "SELECT * FROM tb_layers WHERE id_pag=" & cIntero(pag) & " AND visibile ORDER BY z_order"
	end if
	
	rs.Open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
	while not rs.eof
		select case rs("id_tipo")
			case 1
				CALL MakeLayer( "lay_" & rs.AbsolutePosition, rs.AbsolutePosition)
			case 2
				CALL MakeImage( "lay_" & rs.AbsolutePosition, rs.AbsolutePosition)	
			case 4
				CALL MakeObject( "lay_" & rs.AbsolutePosition, rs.AbsolutePosition)	
		end select
		rs.MoveNext
		Response.Write vbCRLF & vbCRLF
	wend
	' Google analitycs
	if cString(ga_code)<>"" then %>
			<script type="text/javascript">
				// <![CDATA[
				var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
				document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
				// ]]>
			</script>
			<script type="text/javascript">
				// <![CDATA[
				try {
				   	var pageTracker = _gat._getTracker("<%= ga_code %>");
				   	pageTracker._initData();
				   	pageTracker._trackPageview();
				} 
				catch (e) { /* javascript semi-disabilitato */ }
				// ]]>
			</script>
	<% end if 
	rs.close
	
	'*************************************************************************************
	'gestione statistiche n° visitatori della pagina
	sql = "UPDATE tb_pages set contatore=contatore + 1 WHERE id_page=" & cIntero(pag)
	CALL conn.execute(sql, 0, adExecuteNoRecords)

	conn.close
	set rs = nothing
	set conn = nothing
	
	%>
	
</body>

</html>
<%
FUNCTION DIV_Start_Tag(l_name, l_left, l_top, l_width, l_height, l_z_order )
	DIV_Start_Tag = "<div id=""" & l_name & """ " & _
					"style=""position:absolute; left:" & l_left & "px;" & _
										 	  " top:" & l_top & "px;" & _
										 	  " width:" & l_width & +"px;" & _
										 	  " height:" & l_height & "px;" & _
											  IIF(debug, "border:1px dotted red;", "") & _
										 	  " z-index:" & l_z_order & """" &_
					">"
END FUNCTION


SUB MakeLayer( lname , z_order)
	response.write DIV_Start_Tag(lname, rs("x"),rs("y"),rs("largo"),rs("alto"),z_order) &_
				   rs("html") & _
				   "</div>"
END SUB

'...........................
'Makeobject per area visibile
SUB MakeObject(byref lname, z_order)
	dim confstr,obj
	
	confstr = replace(rs("testo"),vbCRLF,"")
	response.write DIV_Start_Tag(lname, rs("x"),rs("y"),rs("largo"),rs("alto"),z_order)
		
	if instr(1,rs("aspcode"),".asp", vbTextCompare)>0 then 
		'creazione oggetto con pagina ASP eseguita
        Session("LAYER_ID") = rs("id_lay")
		Session("LAYER_NAME") = lname
		Session("LAYER_WIDTH") = rs("largo")
		Session("LAYER_HEIGHT") = rs("alto")
		Session("LAYER_Z_ORDER") = z_order
		Session("CONFSTR") = confstr
		if debug then
			response.write vbCrLF + "<!-- aspcode: " & PathExecute & "plugin/" & rs("aspcode") & "//-->" + vbCrLF + _
							"<!-- config: " & confstr & "//-->" + vbCrLF
		end if
        if instr(1,rs("aspcode"),"amministrazione/",vbTextCompare)>0 then
			Server.Execute(PathExecute & rs("aspcode"))
		else
			Server.Execute(PathExecute & "plugin/" & rs("aspcode"))
		end if
	else
		'creazione oggetto da oggetto COM
		set obj = Server.CreateObject(rs("aspcode"))
		obj.Configura(confstr)
		obj.settaLayer(rs("id_lay"))
		obj.disegna()
		set obj =  nothing
	end if
	
	response.Write "</div>"
	
END SUB

'...........................
'Makeobject per area amministrativa
'SUB MakeObject(byref lname, z_order)
'	with response
'		.write "<div id="""+lname+""" "
'		.write divstyle( rs("x"),rs("y"),rs("largo"),rs("alto"),z_order)
'		.write "><img src="""+path_azienda+"objects/"+rs("nome")+""" border=""0""></div>"
'	end with
'END SUB

SUB MakeImage(byref lname, z_order)
	with response
		.write DIV_Start_Tag(lname, rs("x"),rs("y"),rs("largo"),rs("alto"),z_order)
		if rs("format")<>"[]" then
			'compone link su immagine
			.write "<A HREF="+extractLink(rs("format"))
			if isBlank( rs("format") ) then
				.write " target=""_blank"" "
			end if
			.write ">" 
		end if
		'scrive immagine
		.write "<img src="""+path_azienda+"images/"+rs("nome")+""" border=""0"" alt="""">"
		if rs("format")<>"[]" then
			.write "</a>"
		end if
		.write "</div>"
	end with
END SUB

'[[1, 1, "javascript:OpenImage('panorama_venice.jpg', '','')", 0]]
'[[1, 1, "http://www.turismovenezia.it", 1]]
'["1 1 link dynalay.asp?PAGINA=892"] 		-- vecchia versione (Editor 2.X)
function extractLink( format )
	if instr(1, format, "[[", vbTextCompare)>0 then
		'nuova gestione link con target
		dim QuoteBegin, QuoteEnd
		QuoteBegin = instr(1, format, """", vbTextCompare)
		QuoteEnd = instrrev(format, """", vbTrue, vbTextCompare) + 1
		extractLink = mid(format, QuoteBegin, (QuoteEnd - QuoteBegin))
	else
		'vecchia gestione link (Editor 2.x)
		dim a
		a = split( mid(format,3,len(format)-4),"," )
		if ubound(a) > 0 then
			extractLink = a(2)
		else
			'extractLink = "#"
			a = split( format," " )
			extractLink = left( a(3), len(a(3))-2 )
		end if
	end if
end function

function isBlank(format)
	if instr(1, format, "[[", vbTextCompare)>0 then
		format = right(format, len(format) - instrrev(format, ",", vbTrue, vbTextCompare))
		isBlank = (instr(1, format, "1", vbTextCompare) > 0)
	else
		isBlank = false
	end if
end function

function setLanguage( sigla )
	dim s
	select case lcase(sigla)
	case "it"
		s = "italian"
	case "en"
		s = "english"
	case "fr"
		s = "french"
	case else
		s = sigla
	end select
	setLanguage = s
end function

SUB DeterminaMetaTag( conn, pagina_sito, webs_meta_defs )
	dim sql,rs,lingua,metadefs
	' 
	sql = "SELECT * FROM tb_paginesito WHERE id_pagineSito=" & pagina_sito
	set rs = server.createObject("ADODB.recordset")
	rs.Open sql,conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	metadefs = ""
	lingua = Session("LINGUA")
	'meta tag per le description
	if rs("PAGE_description_" & lingua)<>"" then
		metadefs = metadefs & _
			"<META NAME=""description"" CONTENT=""" & rs("PAGE_description_" & lingua) & """>" & vbcrlf
	end if
	'meta tag per le keywords
	if rs("PAGE_keywords_" & lingua)<>"" then
		metadefs = metadefs & _
			"<META NAME=""keywords"" CONTENT=""" & rs("PAGE_keywords_" & lingua) & """>" & vbcrlf
	end if
	if metadefs = "" then
		metadefs = webs_meta_defs
	end if
	response.write metadefs + vbCRLF
	rs.close
END SUB
%>
