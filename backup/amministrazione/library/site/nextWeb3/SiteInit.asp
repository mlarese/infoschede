<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../../tools.asp"-->
<%
dim sql, conn, rs, a, virtual_dir, lingua

if request("lingua")<>"" then
	CALL SetSessionLingua(request("lingua"))
elseif Session("LINGUA")="" then
	Session("LINGUA") = "it"
end if

set conn = server.createObject("ADODB.Connection")
conn.Open Application("l_conn_ConnectionString"),"",""
set rs = server.createObject("ADODB.recordset") 

'*************************************************************************************
'recupera dati del sito
sql = "SELECT * FROM tb_webs WHERE "

if Application("AZ_ID")="" and Session("AZ_ID")="" then
	'calcolo directory virtuale
	'sql = sql & " nome_webs='" & virtual_dir & "' " &_
	'	  		" OR nome_stage1_web='" & virtual_dir & "' " &_
	'	  		" OR nome_stage2_web='" & virtual_dir & "'"
	sql = sql & " nome_webs<>'' "
else
	if Session("AZ_ID")<>"" then
		'utilizzato solo nell'area amministrativa
		sql = sql & " id_webs=" & cIntero(Session("AZ_ID"))
	else
		sql = sql & " id_webs=" & cIntero(Application("AZ_ID"))
	end if
end if

'recupera dati sito
rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
Session("AZ_ID") = rs("id_webs")

'legge meta tag per il sito
Session("META_Author") = rs("META_Author")
Session("META_TAG_DEFINITIONS") = ""
lingua = Session("LINGUA")
'for each lingua in Application("LINGUE")
	'meta tag per la descrizione
	if rs("META_description_" & lingua)<>"" then
		Session("META_TAG_DEFINITIONS") = Session("META_TAG_DEFINITIONS") & _
			"<META NAME=""description"" CONTENT=""" & rs("META_description_" & lingua) & """>" & vbcrlf
	end if
'next
'for each lingua in Application("LINGUE")
	'meta tag per le keywords
	if rs("META_keywords_" & lingua)<>"" then
		Session("META_TAG_DEFINITIONS") = Session("META_TAG_DEFINITIONS") & _
			"<META NAME=""keywords"" CONTENT=""" & rs("META_keywords_" & lingua) & """>" & vbcrlf
	end if
'next

rs.close


'*************************************************************************************
'crea vettore pagine
a = array()
sql = "SELECT id_pagineSito, (id_pagDyn_" & Session("LINGUA") & ") AS PAGE_ID " &_
	  " FROM tb_pagineSito WHERE id_web=" & cIntero(Session("AZ_ID")) & _
	  " ORDER BY id_pagineSito DESC"

rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adAsyncFetch
if not rs.EOF then
	'imposta dimensione array pagine
	redim a(rs("id_paginesito")+1)
	
	'salva nell'array l'elenco delle pagine attive del sito
	while not rs.EOF
		'Response.Write "a(" & rs("id_paginesito") & ") = " & rs("PAGE_ID")
		a(rs("id_paginesito")) = rs("PAGE_ID")
		rs.MoveNext
	wend
end if 
rs.Close
'imposta vettore di sessione pagine
Session("PAGINE") = a
Session("LANGUAGE") = setLanguage( lingua )

'*************************************************************************************
'se viene passato il parametro PS converte la paginaSito richiesta nella corrispettiva pagina di tb_pages corretta
if request("PS")<>"" AND isNumeric(request("PS")) AND request("PAGINA")="" then
	response.redirect "dynalay.asp?PAGINA=" & Session("PAGINE")(cLng(request("PS")))
end if


'*************************************************************************************
'gestione statistiche n° visitatori singoli del sito
if not Session("VISIT_LOGGED") AND instr(1, Request.ServerVariables("SCRIPT_NAME"), "nextWeb", vbTextCompare)<1 then
	sql = "UPDATE tb_webs set contatore=contatore + 1 WHERE id_webs=" & cIntero(Session("AZ_ID"))
	CALL conn.execute(sql, 0, adExecuteNoRecords)

	Session("VISIT_LOGGED") = true
end if

conn.Close
set rs = nothing
set conn = nothing


'**********************************************************************************************
'FUNZIONI INTERNE
function GetPageName()
	GetPageName = Right(Request.ServerVariables("SCRIPT_NAME"), (Len(Request.ServerVariables("SCRIPT_NAME")) - instrRev(Request.ServerVariables("SCRIPT_NAME"), "/")))
end function


'**********************************************************************************************
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
%>
