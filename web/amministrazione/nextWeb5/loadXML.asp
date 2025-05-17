<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<!--#INCLUDE FILE="../library/TOOLS4ADMIN.ASP" -->
<%

dim page_id
dim conn, rs, sql
set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

if cInteger(request("PAGINA"))>0 then
	page_id = cInteger(request("PAGINA"))
else
	page_id = cInteger(request("PAG"))
end if

'recupera dati pagina
sql = " SELECT id_template, (SELECT MAX(z_order) FROM tb_layers WHERE id_pag=tb_pages.id_template) AS MAX_LAYER" &_
	  " FROM tb_pages WHERE id_page=" & page_id
rs.open sql, conn, adOpenForwardOnly, adLockReadonly, adCmdText

'compone query per recupero layers
if cInteger(rs("id_template"))>0 then
	'pagina con template
	sql = "SELECT id_lay, (" & page_id & ") AS id_pag, id_tipo, " & _
		  SQL_If(conn, "id_pag=" & page_id, "z_order + " & cInteger(rs("MAX_LAYER")), "z_order") &" AS z_order, " & _
		  " nome, visibile, x, y, largo, alto, testo, format, html, aspcode, " & _
		  SQL_If(conn, "id_pag=" & page_id, "'U'", "'L'") &" AS Stato, "& _
		  " tipo_contenuto, RTF, checksum_stili " &_
		  " FROM tb_layers WHERE id_pag=" & page_id	& " OR id_pag=" & rs("id_template")
else
	'pagina senza template
	sql = " SELECT id_lay, id_pag, id_tipo, z_order, nome, visibile, x, y, largo, alto, testo, format, " &_
		  " html, aspcode, 'U' AS Stato, tipo_contenuto, RTF, checksum_stili " &_
		  " FROM tb_layers WHERE id_pag=" & page_id
end if
rs.close


'apre recordset su layer della pagina
rs.CursorLocation = adUseClient
rs.open sql, conn, adOpenForwardOnly, adLockReadonly, adCmdText

if rs.eof then
	'nessun layer per la pagina: restituisce lo schema vuoto
	response.ContentType = "text/plain"
	response.write "vuoto"
else
	'ordina i layers per z_order definitivo
	rs.Sort = "z_order"
	
	'restituisce XML completo del recordset
	CALL Export_XML(rs, false)
end if
rs.close
conn.close
set rs = nothing
set conn = nothing
%>

