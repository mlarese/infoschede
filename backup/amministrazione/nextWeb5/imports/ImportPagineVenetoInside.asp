<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout=200000 %>
<!--#INCLUDE FILE="../../library/Tools.asp" -->
<!--#INCLUDE FILE="../../library/Tools4Admin.asp" -->
<!--#include file="../Tools_NextWeb5.asp"-->

INIZIO<BR>
<%


dim idxPadreBase
dim ID_paginaDaCopiare
idxPadreBase = 5785
ID_paginaDaCopiare = 735

dim conn, sql, rs, rsp, rs_pag, ID_paginaCreata, value, codice, ContentId, codicePadre, idxPadre
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")
set rs_pag = Server.CreateObject("ADODB.RecordSet")
set conn = Index.conn

conn.begintrans

sql = " SELECT * FROM PagineDaImportare " + _
	  " WHERE IsNull([TAG Title],'') <> '' " + _
	  " ORDER BY carattere, categoria, area "
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext 

while not rs.eof 
	response.write rs.Absoluteposition & " / " & rs.recordcount & "<br>"
	codice = "matriceLink:idCat=" & cIntero(rs("id categoria")) & ";idArea=" & cIntero(rs("id area")) & ";idCar=" &cIntero(rs("id carattere"))
	response.write "codice:" & codice & "<br>"
	
	idxPadre = 0
	
	if cIntero(rs("id carattere"))<>0 then
		codicePadre = "matriceLink:idCat=" & cIntero(rs("id categoria")) & ";idArea=" & cIntero(rs("id area")) & ";idCar=0"
		idxPadre = getIdxPadre(codicePadre)
		if idxPadre = 0 then
			codicePadre = "matriceLink:idCat=" & cIntero(rs("id categoria")) & ";idArea=" & cIntero(rs("id_area_padre")) & ";idCar=0"
			idxPadre = getIdxPadre(codicePadre)
		end if
	end if
	
	if idxPadre = 0 AND cIntero(rs("id categoria"))<>0 then
		codicePadre = "matriceLink:idCat=0;idArea=" & cIntero(rs("id area")) & ";idCar=0"
		idxPadre = getIdxPadre(codicePadre)
		if idxPadre = 0 then
			codicePadre = "matriceLink:idCat=0;idArea=" & cIntero(rs("id_area_padre")) & ";idCar=0"
			idxPadre = getIdxPadre(codicePadre)
		end if
	end if
	
	if idxPadre = 0 then
		codicePadre = "matriceLink:idCat=0;idArea=" & cIntero(rs("id_area_padre")) & ";idCar=0"
		idxPadre = getIdxPadre(codicePadre)
	end if
	
	
	if idxPadre = 0 then
		codicePadre = "matriceLink"
		idxPadre = idxPadreBase
	end if
		
	sql = "SELECT * FROM tb_pagineSito WHERE nome_ps_interno like '" & codice & "'"
	rsp.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText

	if not rsp.eof then
		response.write "pagina esistente:" & ID_paginaCreata & "<br>"
		'riga commentata per non sovrascrivere le modifiche alle pagine fatte dal primo import.
		'ID_paginaCreata = CopiaPaginaSito(conn, ID_paginaDaCopiare, rsp("id_pagineSito"), true)
		ID_paginaCreata = rsp("id_pagineSito")
	else
		response.write "pagina creata:" & ID_paginaCreata & "<br>"
		ID_paginaCreata = CopiaPaginaSito(conn, ID_paginaDaCopiare, 0, true)
		rsp.close
		
		sql = "SELECT * FROM tb_pagineSito WHERE id_pagineSito=" & ID_paginaCreata
response.write sql
		rsp.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText

		rsp("nome_ps_IT") = rs("Titolo pagina")
		rsp("nome_ps_EN") = rsp("nome_ps_IT")
		rsp("nome_ps_FR") = rsp("nome_ps_IT")
		rsp("nome_ps_DE") = rsp("nome_ps_IT")
		rsp("nome_ps_ES") = rsp("nome_ps_IT")
		rsp("PAGE_description_IT") = rs("Description")
		rsp("PAGE_keywords_IT") = rs("Keywords")
		rsp("nome_ps_interno") = codice
		rsp.update
		
	end if
	
	CALL PaginaSitoUpdatePages(conn, ID_paginaCreata)
	
	sql = " SELECT * FROM tb_pages INNER JOIN tb_layers ON tb_pages.id_page = tb_layers.id_pag " + _
		  " WHERE tb_pages.id_PaginaSito=" & ID_paginaCreata
	rs_pag.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	
	while not rs_pag.eof
		value = cString(rs_pag("testo"))
		value = replace(value, "<valuecategoriabaseid>", cIntero(rs("id categoria")))
		value = replace(value, "<valueindicebaseid>", cIntero(rs("id area")))
		value = replace(value, "<valuecarattereid>", cIntero(rs("id carattere")))
		rs_pag("testo") = value
		rs_pag.update
		rs_pag.moveNext
	wend
	
	rs_pag.close
	rsp.close
	
	response.write "padre:" & idxPadre & "<br>"
		
	ContentId = Index_UpdateItem(conn, "tb_pagineSito", ID_paginaCreata, true)
	index.dizionario("idx_content_id") = ContentId
	index.dizionario("idx_link_tipo") = lnk_interno
	index.dizionario("idx_link_pagina_id") = ID_paginaCreata
	index.dizionario("idx_padre_id") = idxPadre
	index.dizionario("idx_principale") = true
	index.dizionario("idx_visibile") = false
	index.dizionario("idx_alt_it") = rs("tag title")
	index.dizionario("idx_alt_en") = index.dizionario("idx_alt_it")
	index.dizionario("idx_alt_fr") = index.dizionario("idx_alt_it")
	index.dizionario("idx_alt_de") = index.dizionario("idx_alt_it")
	index.dizionario("idx_alt_es") = index.dizionario("idx_alt_it")
	sql = "SELECT idx_id FROM tb_contents_index WHERE idx_content_id=" & ContentId & " AND idx_padre_id=" & idxPadre
	CALL index.Salva(cIntero(GetValueList(index.conn, NULL, sql)))
	Session("errore") = ""
	response.write "creato conteuto:" & ContentId & "<br><br>"
	
	rs.moveNext
wend
rs.close
set rs = nothing
set rsp = nothing
set rs_pag = nothing

conn.committrans
'conn.rollbacktrans


function getIdxPadre(codice)
	dim sql
	sql = " SELECT * FROM v_indice INNER JOIN tb_paginesito ON v_indice.co_link_pagina_id = tb_paginesito.id_paginesito " + _
		  " WHERE nome_ps_interno LIKE '" + codice + "'"
	response.write sql & "<br>"
	getIdxPadre = cIntero(GetValueList(conn, rsp, sql))
	response.write "getIdxPadre = " & getIdxPadre & "<br>"
end function
%>

FINE