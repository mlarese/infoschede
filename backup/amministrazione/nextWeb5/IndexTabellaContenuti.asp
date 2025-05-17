<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit
response.charset = "UTF-8"
response.buffer = true
Server.ScriptTimeout = 1073741824 %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/ExportTools.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="IndexMetaTag_TOOLS.asp" -->
<%
Server.ScriptTimeout = 360

dim rs, rsi, sql, i, maxLivello, WebsBaseUrl, WebsId, IdxUrl
set rs = Server.CreateObject("ADODB.RecordSet")
set rsi = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT MAX(idx_livello) FROM tb_contents_index"
maxLivello = cIntero(GetValueList(index.conn, rs, sql))

'sql = index.QueryElenco(false, "")
sql = Session("IDX_SQL")
rs.open sql, index.conn, adOpenStatic, adLockOptimistic, adCmdText 
%>
<table cellpadding="0" cellspacing="0" border="1">
	<tr>
		<th>n</th>
		<th>ID nodo</th>
		<% for i = 0 to maxLivello %>
			<th>livello <%= i %></th>
		<% next %>
		<th>visibile</th>
		<th>tipo contenuto (applicativo)</th>
		<th>pagina nextweb (id pagina)</th>
		<th>url (ITALIANO)</th>
	</tr>
	<%  i = 0
	websId = 0
	while not rs.eof 
		sql = " SELECT * FROM v_indice INNER JOIN tb_siti ON v_indice.tab_sito_id = tb_siti.id_sito " + _
			  " LEFT JOIN tb_pagineSito ON v_indice.idx_link_pagina_id = tb_paginesito.id_paginesito " + _
			  " WHERE idx_id=" & rs("idx_id")
		rsi.open sql, index.conn, adOpenStatic, adLockOptimistic, adCmdText 
		if websId <> cIntero(rsi("idx_webs_id")) then
			websId = WebsId
			WebsBaseUrl = GetSiteUrl(index.conn, rsi("idx_webs_id"), 5)
			if instr(1, WebsBaseUrl, "http", vbTextCompare)>0 then
				WebsBaseUrl = "http://" + WebsBaseUrl
			end if
		end if %>
		<tr>
			<td><%= i %></td>
			<td><%= rs("idx_id") %></td>
			<% if cIntero(rs("idx_livello"))>0 then %>
				<td <% if cIntero(rs("idx_livello"))>1 then %>colspan="<%= cIntero(rs("idx_livello")) %>"<% end if %>>&nbsp;</td> 
			<% end if %>
			<td <% if cIntero(rs("idx_livello")) <= maxLivello then %>colspan="<%= (maxLivello + 1) - cIntero(rs("idx_livello")) %>"<% end if %> style="color:<%= rsi("tab_colore") %>;">
				<%= rs("co_titolo_it") %>
			</td>
			<td><%= IIF(rsi("visibile_assoluto"), "SI", "NO") %></td>
			<td style="color:<%= rsi("tab_colore") %>;">
				<%= rsi("tab_titolo") %> ( <%= GetApplicationShortName(rsi("sito_nome")) %> )
			</td>
			<td>
				<% if cIntero(rsi("idx_link_pagina_id"))>0 then %>
					<%= rsi("nome_ps_it") %> (<%=rsi("idx_link_pagina_id")  %>)
				<% end if %>
			</td>
			<td>
				<% IdxUrl = index.GetNodeUrl(rsi, LINGUA_ITALIANO) %>
				<a href="<%= WebsBaseUrl & IdxUrl %>">
					<%= IdxUrl %>
				</a>
			</td>
		</tr>
		<% rs.movenext
		i = i + 1
		rsi.close
	wend %>
</table>
<%rs.close

set rs = nothing
set rsi = nothing
 %>


