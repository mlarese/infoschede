<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../Tools.asp"-->
<!--#INCLUDE FILE="../Tools4Plugin.asp"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<% 
dim config
set config = new Configuration
'impostazione delle proprieta' di default
Config.AddDefault "NessunaNews_IT", "Nessuna news disponibile."
Config.AddDefault "NessunaNews_EN", "No news availables."
Config.AddDefault "NessunaNews_FR", ""
Config.AddDefault "NessunaNews_DE", ""
Config.AddDefault "NessunaNews_ES", ""

dim conn, rs, sql
Set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.recordset")

sql = " SELECT * FROM tb_News "& _
	  " WHERE "& SQL_CompareDateTime(conn, "news_dataPubbl", adCompareLessThan, Date) & _
	  " AND "& SQL_CompareDateTime(conn, "news_dataScad", adCompareGreaterThan, Date) & _
	  " AND "& SQL_IsTrue(conn, "news_elenco") & _
	  " ORDER BY news_DataPubbl DESC"
rs.open sql, conn, adOpenstatic, adLockOptimistic, adCmdText
if not rs.eof then
	while not rs.eof%>
		<div class="news">
			<h1 class="news">
				<%if cInteger(rs("news_pagina")) > 0 then %>
					<a class="news" href="?PAGINA=<%= session("PAGINE")(rs("news_pagina")) %>"><%= CBL(rs, "news_titolo") %></a>
				<% elseif cString(rs("news_url"))<>"" then %>
					<a class="news" href="<%= rs("news_url") %>" title="<%= rs("news_url") %>" target="news"><%= CBL(rs, "news_titolo") %></a>
				<% else %>
					<%= CBL(rs, "news_titolo") %>
				<% end if %>
			</h1>
			<% 	if CString(rs("news_img")) <> "" then %>
				<table cellpadding="0" cellspacing="0" align="left">
					<tr>
						<td>
							<img src="<%= Config.ImageURL & rs("news_img") %>" border="0" class="news">
						</td>
					</tr>
				</table>
			<% 	end if %>
			<p class="news_data"><%= DataEstesa(rs("news_dataPubbl"), Config.lingua) %></p>
			<p class="news_text"><%= TextEncode(CBL(rs, "news_estratto")) %></p>
			<% 	if CString(rs("news_doc")) <> "" then %>
				<p class="news_doc"><a class="news_doc" target="_blank" href="<%= config.imageURL & rs("news_doc") %>">approfondisci</a></p>
			<% 	end if %>
		</div>
		<% 	rs.movenext
	wend
else %>
	<div class="NoRecords"><%= CBL(Config, "NessunaNews") %></div>
<% end if
set rs = nothing
conn.close
set conn = nothing
%>