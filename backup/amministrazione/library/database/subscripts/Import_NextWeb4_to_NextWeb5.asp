<!--#INCLUDE FILE="../../Tools.asp" -->
<!--#INCLUDE FILE="../../Tools4Admin.asp" -->
<% 
dim conn_4, conn_5, rs4, rs5, sql

Set conn_4 = Server.CreateObject("ADODB.Connection")
Set conn_5 = Server.CreateObject("ADODB.Connection")

Set rs4 = Server.CreateObject("ADODB.RecordSet")
Set rs5 = Server.CreateObject("ADODB.RecordSet")

conn_4.open Application("L_conn_ConnectionString"),"",""
conn_5.open Application("DATA_ConnectionString"),"",""

'..............................................................................
conn_5.begintrans
'..............................................................................

'copia tb_webs 
'..............................................................................
CALL OpenTable("tb_webs", "tb_webs", rs4, rs5, false)
while not rs4.eof
	response.write rs4("id_webs") & "<br>"
	rs5.addnew
	rs5("id_webs") = rs4("id_webs")
	rs5("nome_webs") = rs4("nome_webs")
	rs5("URL_base") = rs4("nome_webs")
	rs5("id_home_page") = rs4("id_home_page")
	rs5("lingua_iniziale") = rs4("lingua_iniziale")
	rs5("lingua_EN") = rs4("lingua_EN")
	rs5("lingua_FR") = rs4("lingua_FR")
	rs5("lingua_DE") = rs4("lingua_DE")
	rs5("lingua_ES") = rs4("lingua_ES")
	rs5("titolo_IT") = rs4("titolo_IT")
	rs5("titolo_EN") = rs4("titolo_EN")
	rs5("titolo_FR") = rs4("titolo_FR")
	rs5("titolo_DE") = rs4("titolo_DE")
	rs5("titolo_ES") = rs4("titolo_ES")
	rs5("META_Author") = rs4("META_Author")
	rs5("META_keywords_IT") = rs4("META_keywords_IT")
	rs5("META_keywords_EN") = rs4("META_keywords_EN")
	rs5("META_keywords_FR") = rs4("META_keywords_FR")
	rs5("META_keywords_DE") = rs4("META_keywords_DE")
	rs5("META_keywords_ES") = rs4("META_keywords_ES")
	rs5("META_description_IT") = rs4("META_description_IT")
	rs5("META_description_EN") = rs4("META_description_EN")
	rs5("META_description_FR") = rs4("META_description_FR")
	rs5("META_description_DE") = rs4("META_description_DE")
	rs5("META_description_ES") = rs4("META_description_ES")
	rs5("contatore") = rs4("contatore")
	rs5("contRes") = rs4("contRes")
	rs5("contUtenti") = rs4("contUtenti")
	rs5("contCrawler") = rs4("contCrawler")
	rs5("contAltro") = rs4("contAltro")
	rs5("google_analytics_code") = rs4("google_analytics_code")
	rs5("google_webmaster_tools_verify_code") = rs4("google_webmaster_tools_verify_code")
	rs5("sito_in_aggiornamento") = 0
	rs5("sito_in_costruzione") = 0
	rs5("webs_insData") = NOW()
	rs5("webs_modData") = NOW()
	rs5("webs_modData_pagine") = NOW()
	rs5("sito_accessibile") = 0
	rs5("editor_guide_visibili") = 0
	rs5("editor_guide_colore") = "#000000"
	rs5("editor_guide_posizioni_visibili") = 0
	rs5("editor_help_attivo") = 0
	rs5("URL_rewriting_attivo") = 0
	rs5.update
	
	rs4.movenext
wend
CALL CloseTable("tb_webs", rs4, rs5, false)

'..............................................................................
CALL OpenTable("tb_objects", "tb_objects", rs4, rs5, true)
while not rs4.eof
	
	response.write rs4("id_objects") & "<br>"
	rs5.addnew
	rs5("id_objects") = rs4("id_objects")
	rs5("id_webs") = rs4("id_webs")
	rs5("name_objects") = rs4("name_objects")
	rs5("identif_objects") = rs4("identif_objects")
	rs5("param_list") = rs4("param_list")
	rs5("obj_insData") = NOW()
	rs5("obj_modData") = Now()
	rs5("obj_type") = "ascx"
	
	rs5.update
	
	rs4.movenext
wend
CALL CloseTable("tb_objects", rs4, rs5, true)

'..............................................................................
CALL OpenTable("tb_paginesito", "tb_paginesito", rs4, rs5, true)
while not rs4.eof
	
	response.write rs4("id_pagineSito") & "<br>"
	rs5.addnew
	
	rs5("id_pagineSito") = rs4("id_pagineSito")
	rs5("id_web") = rs4("id_web")
	rs5("id_pagDyn_IT") = rs4("id_pagDyn_IT")
	rs5("id_pagDyn_EN") = rs4("id_pagDyn_EN")
	rs5("id_pagDyn_FR") = rs4("id_pagDyn_FR")
	rs5("id_pagDyn_DE") = rs4("id_pagDyn_DE")
	rs5("id_pagDyn_ES") = rs4("id_pagDyn_ES")
	rs5("id_pagStage_IT") = rs4("id_pagStage_IT")
	rs5("id_pagStage_EN") = rs4("id_pagStage_EN")
	rs5("id_pagStage_FR") = rs4("id_pagStage_FR")
	rs5("id_pagStage_DE") = rs4("id_pagStage_DE")
	rs5("id_pagStage_ES") = rs4("id_pagStage_ES")
	rs5("nome_ps_IT") = rs4("nome_ps_IT")
	rs5("nome_ps_EN") = rs4("nome_ps_EN")
	rs5("nome_ps_FR") = rs4("nome_ps_FR")
	rs5("nome_ps_DE") = rs4("nome_ps_DE")
	rs5("nome_ps_ES") = rs4("nome_ps_ES")
	rs5("PAGE_keywords_IT") = rs4("PAGE_keywords_IT")
	rs5("PAGE_keywords_EN") = rs4("PAGE_keywords_EN")
	rs5("PAGE_keywords_FR") = rs4("PAGE_keywords_FR")
	rs5("PAGE_keywords_DE") = rs4("PAGE_keywords_DE")
	rs5("PAGE_keywords_ES") = rs4("PAGE_keywords_ES")
	rs5("PAGE_description_IT") = rs4("PAGE_description_IT")
	rs5("PAGE_description_EN") = rs4("PAGE_description_EN")
	rs5("PAGE_description_FR") = rs4("PAGE_description_FR")
	rs5("PAGE_description_DE") = rs4("PAGE_description_DE")
	rs5("PAGE_description_ES") = rs4("PAGE_description_ES")
	rs5("archiviata") = 0
	rs5("riservata") = 0
	rs5("ps_insData") = rs4("DataModifica")
	rs5("ps_modData") = rs4("DataModifica")
	
	rs5.update
	
	rs4.movenext
wend
CALL CloseTable("tb_paginesito", rs4, rs5, true)

'..............................................................................
CALL OpenTable(" tb_pages p LEFT JOIN tb_pagineSito ps ON " + _
			   " ( p.id_page = ps.id_pagDyn_IT OR p.id_page = ps.id_pagDyn_EN OR p.id_page = ps.id_pagDyn_FR OR " + _
			   "   p.id_page = ps.id_pagDyn_DE OR p.id_page = ps.id_pagDyn_ES OR p.id_page = ps.id_pagStage_IT OR p.id_page = ps.id_pagStage_EN OR " + _
			   "   p.id_page = ps.id_pagStage_FR OR p.id_page = ps.id_pagStage_DE OR p.id_page = ps.id_pagStage_ES) ", "tb_pages", rs4, rs5, true)
while not rs4.eof
	
	response.write rs4("id_page") & "<br>"
	rs5.addnew
	rs5("id_page") = rs4("id_page")
	rs5("id_webs") = rs4("id_webs")
	rs5("id_template") = rs4("id_template")
	rs5("nomepage") = cString(rs4("nomepage"))
	rs5("template") = rs4("template")
	rs5("SfondoColore") = rs4("SfondoColore")
	rs5("SfondoImmagine") = rs4("SfondoImmagine")
	rs5("lingua") = rs4("lingua")
	if rs4("template") then
		rs5("Contatore") = 0
		rs5("ContRes") = NOW()
		rs5("contUtenti") = 0
		rs5("contCrawler") = 0
		rs5("contAltro") = 0
	else
		rs5("Contatore") = cIntero(rs4("Contatore"))
		rs5("ContRes") = rs4("ContRes")
		rs5("contUtenti") = cIntero(rs4("contUtenti"))
		rs5("contCrawler") = cIntero(rs4("contCrawler"))
		rs5("contAltro") = cIntero(rs4("contAltro"))
	end if
	rs5("page_insData") = Now()
	rs5("page_modData") = Now()
	rs5("id_PaginaSito") = rs4("id_pagineSito")
	rs5("semplificata") = 0
	
	rs5.update
	
	rs4.movenext
wend
CALL CloseTable("tb_pages", rs4, rs5, true)

'..............................................................................
CALL OpenTable("tb_layers", "tb_layers", rs4, rs5, true)
while not rs4.eof
	
	response.write rs4("id_lay") & "<br>"
	rs5.addnew
	
	rs5("id_lay") = rs4("id_lay")
	rs5("id_pag") = rs4("id_pag")
	rs5("id_tipo") = rs4("id_tipo")
	rs5("id_objects") = rs4("id_objects")
	rs5("z_order") = rs4("z_order")
	rs5("nome") = rs4("nome")
	rs5("visibile") = rs4("visibile")
	rs5("x") = rs4("x")
	rs5("y") = rs4("y")
	rs5("largo") = rs4("largo")
	rs5("alto") = rs4("alto")
	rs5("em_x") = PxToEm(rs4("x"))
	rs5("em_y") = PxToEm(rs4("y"))
	rs5("em_largo") = PxToEm(rs4("largo"))
	rs5("em_alto") = PxToEm(rs4("alto"))
	rs5("html") = rs4("html")
	rs5("format") = rs4("format")
	rs5("testo") = rs4("testo")
	rs5("aspcode") = rs4("aspcode")
	rs5("tipo_contenuto") = "A"
	rs5("rtf") = "RTF"
	rs5("checksum_stili") = ""
	rs5.update
	
	rs4.movenext
wend
CALL CloseTable("tb_layers", rs4, rs5, true)

'..............................................................................
CALL OpenTable("tb_links", "tb_menu", rs4, rs5, true)
while not rs4.eof
	
	response.write rs4("id") & "<br>"
	rs5.addnew
	rs5("m_id") = rs4("id")
	rs5("m_id_webs") = rs4("id_webs")
	rs5("m_nome_it") = rs4("nomelink_it")
	rs5("m_nome_en") = rs4("nomelink_en")
	rs5("m_nome_fr") = rs4("nomelink_de")
	rs5("m_nome_de") = rs4("nomelink_fr")
	rs5("m_nome_es") = rs4("nomelink_es")
	
	rs5.update
	
	rs4.movenext
wend
CALL CloseTable("tb_menu", rs4, rs5, true)

'..............................................................................
CALL OpenTable("tb_menuitem", "tb_menuitem", rs4, rs5, true)
while not rs4.eof
	
	response.write rs4("id_menuItem") & "<br>"
	rs5.addnew
	rs5("mi_id") = rs4("id_menuItem")
	rs5("mi_ordine") = rs4("ordine_menuItem")
	rs5("mi_menu_id") = rs4("id_link")
	rs5("mi_attivo") = rs4("attivo_mi")
	rs5("mi_target") = rs4("link_target")
	rs5("mi_figli") = 0
	rs5("mi_titolo_it") = rs4("titolo_menuItem_IT")
	rs5("mi_titolo_fr") = rs4("titolo_menuItem_fr")
	rs5("mi_titolo_de") = rs4("titolo_menuItem_de")
	rs5("mi_titolo_es") = rs4("titolo_menuItem_es")
	rs5("mi_titolo_en") = rs4("titolo_menuItem_en")
	rs5("mi_link_it") = rs4("link_menuItem_IT")
	rs5("mi_link_en") = rs4("link_menuItem_en")
	rs5("mi_link_fr") = rs4("link_menuItem_fr")
	rs5("mi_link_de") = rs4("link_menuItem_de")
	rs5("mi_link_es") = rs4("link_menuItem_es")
	rs5("mi_image_it") = rs4("image_menuItem_it")
	rs5("mi_image_fr") = rs4("image_menuItem_fr")
	rs5("mi_image_en") = rs4("image_menuItem_en")
	rs5("mi_image_de") = rs4("image_menuItem_de")
	rs5("mi_image_es") = rs4("image_menuItem_es")
	rs5("mi_tag_title_it") = rs4("tag_title_IT")
	rs5("mi_tag_title_en") = rs4("tag_title_en")
	rs5("mi_tag_title_fr") = rs4("tag_title_fr")
	rs5("mi_tag_title_de") = rs4("tag_title_de")
	rs5("mi_tag_title_es") = rs4("tag_title_es")
	
	rs5.update
	
	rs4.movenext
wend
CALL CloseTable("tb_menuitem", rs4, rs5, true)
	
	
	
'..............................................................................
conn_5.committrans
'..............................................................................

conn_4.close()
conn_5.close()

set rs4 = nothing
set rs5 = nothing

set conn_4 = nothing
set conn_5 = nothing
%>

<h1>COPIA ESEGUITA CORRETTAMENTE</h1>

<%
'............................................................................................................................................................
sub OpenTable(SourceTable, DestTable, rs4, rs5, SetIdentity)
	%>
	INIZIO COPIA <strong><%= SourceTable %></strong> in <strong><%= DestTable %></strong><br>
	<%
	dim sql
	
	if SetIdentity then
		sql = " SET IDENTITY_INSERT " + DestTable + " ON "
		CALL conn_5.execute(sql, ,adCmdText)
	end if
	
	sql = "SELECT * FROM " + SourceTable 
	rs4.open sql, conn_4, adOpenStatic, adLockOptimistic
	
	sql = "SELECT * FROM " + DestTable 
	rs5.open sql, conn_5, adOpenStatic, adLockOptimistic
	
end sub


sub CloseTable(TableName, rs4, rs5, SetIdentity)
	dim sql
	
	rs4.close()
	rs5.close()
	
	if SetIdentity then
		sql = " SET IDENTITY_INSERT " + TableName + " OFF "
    	CALL conn_5.execute(sql, ,adCmdText)
	end if
	
	%>
	FINE COPIA <strong><%= TableName %></strong><br>
	<hr>
	<%
end sub
%>