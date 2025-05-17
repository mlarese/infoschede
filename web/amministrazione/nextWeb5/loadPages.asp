<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<!--#INCLUDE FILE="../library/TOOLS4ADMIN.ASP" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
dim page_id
dim conn, rs, sql
dim lingua, base_sql
set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

page_id = cInteger(request("PAGINA"))

'recupera dati pagina
sql = "SELECT * FROM tb_pages WHERE id_page=" & page_id
rs.open sql, conn, adOpenForwardOnly, adLockReadonly, adCmdText
if rs("template") then
	'la pagina e' un template: carico la lista completa di pagine per il sito corrente
	sql = ""
	'query di base per ogni lingua
	base_sql = " SELECT id_page, ('<LINGUA> - ' "& SQL_Concat(conn) & SQL_PaginaSitoNome(conn, "nomepage") + ") AS nome_page " &_
			   " FROM (tb_pages INNER JOIN tb_PagineSito ON tb_pages.id_page = tb_PagineSito.id_pagDyn_<LINGUA>) " &_
			   " INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " & _
			   " WHERE tb_webs.id_webs=" & rs("id_webs") & _
					 " AND not "& SQL_IsTrue(conn, "template")
    if uBound(Application("LINGUE"))>0 then
    	for each lingua in Application("LINGUE")
			sql = sql & " (" & replace(base_sql, "<LINGUA>", lingua)
			if lingua <> LINGUA_ITALIANO then
				sql = sql & " AND " & SQL_IsTrue(conn, "tb_webs.lingua_" & lingua)
			end if
			sql = sql & ") UNION "
	    next
    	sql = left(sql, len(sql)-6) & " ORDER BY nome_page"
    else
        sql = replace(base_sql, "<LINGUA>", LINGUA_ITALIANO) & " ORDER BY nomepage"
    end if
else
	'la pagina non e' un template: carico solo le pagine della lingua corrente e del sito corrente
	rs.close
	
	'recupera lingua della pagina corrente
	sql = "SELECT * FROM tb_pagineSito WHERE "
    for each lingua in Application("LINGUE")
		sql = sql & "id_pagDyn_" & lingua & "=" & page_id & " OR id_pagStage_" & lingua & "=" & page_id & " OR "
	next
	sql = left(sql, len(sql)-3)
	rs.open sql, conn, adOpenForwardOnly, adLockReadonly, adCmdText
	if not rs.eof then
		for each lingua in Application("LINGUE")
			if cInteger(rs("id_pagDyn_" & lingua)) = cInteger(page_id) OR _
			   cInteger(rs("id_pagStage_" & lingua)) = cInteger(page_id) then
				exit for
			end if
		next
	
		'compone query per lista pagine
		sql = " SELECT id_page, ('" & lingua & " - ' "& SQL_Concat(conn) & _ 
                                 SQL_PaginaSitoNome(conn, SQL_If(conn, SQL_IsNull(conn, "nomepage"), "nome_ps_IT", "nomepage")) + ") AS nome_page " + _
			  " FROM tb_pages INNER JOIN tb_PagineSito ON tb_pages.id_page = tb_PagineSito.id_pagDyn_" & lingua &_
			  " WHERE id_webs=" & rs("id_web") & " AND not "& SQL_IsTrue(conn, "template") &" ORDER BY nomepage"
	else
		'la pagina non appartiene ad alcun sito: nextMAIL
		sql = ""
		for each lingua in Application("LINGUE")
			sql = sql & " SELECT id_page, ('" & lingua & " - ' "& SQL_Concat(conn) & SQL_PaginaSitoNome(conn, "nomepage") + ") AS nome_page " &_
			  	        " FROM tb_pages INNER JOIN tb_PagineSito ON tb_pages.id_page = tb_PagineSito.id_pagDyn_" & lingua &_
			  	  		" WHERE id_webs=" & Application("AZ_ID") & " AND not "& SQL_IsTrue(conn, "template") &_
				  		" UNION "
		next
		sql = left(sql, len(sql)-6) + " ORDER BY nome_page"
	end if
end if
rs.close

'apre recorset delle pagine
response.write sql
rs.open sql, conn, adOpenForwardOnly, adLockReadonly, adCmdText

dim ListaPageId, ListaPageName

ListaPageId = ""
ListaPageName = ""

while not rs.eof
	ListaPageId = ListaPageId & rs("id_page")
	ListaPageName = ListaPageName & Server.HtmlEncode(rs("nome_page"))
	rs.movenext
	if not rs.eof then
		ListaPageId = ListaPageId & vbCrLf
		ListaPageName = ListaPageName & vbCrLF
	end if
wend
rs.close
conn.close
set rs = nothing
set conn = nothing

response.ContentType = "text/xml"
response.clear
%><?xml version="1.0" encoding="UTF-8"?>
<xml>
<pages>
<ids><%= ListaPageId %></ids>
<names><%= ListaPageName %></names>
</pages>
</xml>