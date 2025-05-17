<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#include file="Intestazione.asp"-->
<%

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Import pagine e template"
dicitura.scrivi_con_sottosez()
%>


<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
	<form action="" method="post" id="form1" name="form1">
		<caption>Import paginesito: selezione database</caption>
        <tr><th colspan="2">Database da cui importare</th></tr>
        <% if request("source_import")="" AND request("conn_import")="" then %>
			<tr>
				<td class="label" style="width:18%;">file da cui importare:</td>
				<td class="content">
					<% CALL WriteFilePicker_Input(Session("AZ_ID"), "images", "form1", "source_import", request("source_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file dal quale vengono importate le pagine.</span>
				</td>
			</tr>
			<tr>
				<td class="label" style="width:18%;">connessione da cui importare:</td>
				<td class="content">
					<input type="text" name="conn_import" value="" class="text" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					<input style="width:20%;" type="submit" class="button" name="importa" value="AVANTI &gt;&gt;">
				</td>
			</tr>
		<% else 
			dim dbPath, sql
			dim conn, sconn, rs, srs
			
			dbPath = Application("IMAGE_PATH") & Session("AZ_ID") & "\images\" & replace(request("source_import"), "/", "\")
			
			set rs = Server.CreateObject("ADODB.RecordSet")
			set srs = Server.CreateObject("ADODB.RecordSet")
			set conn = Server.CreateObject("ADODB.Connection")
			set sconn = Server.CreateObject("ADODB.Connection")
			conn.open Application("DATA_ConnectionString")
			if request("conn_import")<>"" then
				sconn.open request("conn_import")
			else
				sconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";"
			end if
			set Session("nw5_import_connection") = sconn
			%>
			<tr>
				<td class="label" style="width:18%;">file da cui importare:</td>
				<td class="content">
					<%= sconn.connectionString %>
				</td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<tr><th colspan="4">Sezione plugin</th></tr>
			
			<% sql = "SELECT * FROM tb_webs ORDER BY nome_webs" 
			srs.open sql, sconn, adOpenStatic, adLockOptimistic, adCmdText %>
			<tr>
				<th class="L2" colspan="3">SITO</th>
				<th class="L2">PLUGIN</th>
			</tr>
			<% while not srs.eof %>
				<tr>
					<td class="content" colspan="3"><%= srs("nome_webs") %></td>
					<td class="content_center">
						<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ImportPlugins.asp?ID_WEB=<%= srs("id_webs") %>', 'PluginsImporta', 680, 405, true)" title="Import plugin sito" <%= ACTIVE_STATUS %>>IMPORTA</a>
					</td>
				</tr>
			<% srs.movenext
			wend
			srs.close
			%>

			<tr><th colspan="4">Elenco templates</th></tr>
			<% sql = "SELECT * FROM tb_webs INNER JOIN tb_pages ON tb_webs.id_webs = tb_pages.id_webs WHERE "  & SQL_IsTrue(sconn, "template") & " ORDER BY nome_webs, nomepage"
			srs.open sql, sconn, adOpenStatic, adLockOptimistic, adCmdText %>
			<tr><td class="content" colspan="4">trovate n&deg; <%= srs.recordcount %> template</td></tr>
			<tr>
				<th class="L2">SITO</th>
				<th class="L2">PAGINA</th>
				<th class="l2_center">NUMERO</th>
				<th class="l2_center">IMPORTA</th>
			</tr>
			<% while not srs.eof %>
				<tr>
					<td class="content"><%= srs("nome_webs") %></td>
					<td class="content"><%= srs("nomepage") %></td>
					<td class="content"><%= srs("id_page") %></td>
					<td class="content_center">
						<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ImportTemplate.asp?ID=<%= srs("id_page") %>&ID_WEB=<%= srs("id_webs") %>', 'TemplateImporta', 680, 405, true)" title="Import template sito" <%= ACTIVE_STATUS %>>IMPORTA</a>
					</td>
				</tr>
				<% srs.movenext
			wend
			
			srs.close
			%>
		
			
			<tr><th colspan="4">Elenco pagine</th></tr>
			<% sql = "SELECT * FROM tb_paginesito INNER JOIN tb_webs ON tb_paginesito.id_web = tb_webs.id_webs ORDER BY nome_webs, nome_ps_IT" 
			srs.open sql, sconn, adOpenStatic, adLockOptimistic, adCmdText %>
			<tr><td class="content" colspan="4">trovate n&deg; <%= srs.recordcount %> pagine</td></tr>
			<tr>
				<th class="L2">SITO</th>
				<th class="L2">PAGINA</th>
				<th class="l2_center">NUMERO</th>
				<th class="l2_center">IMPORTA</th>
			</tr>
			<% while not srs.eof %>
				<tr>
					<td class="content"><%= srs("nome_webs") %></td>
					<td class="content"><%= srs("nome_ps_it") %></td>
					<td class="content" title="<%= GetPageNumbers(srs) %>"><%= srs("id_pagineSito") %></td>
					<td class="content_center">
						<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ImportPaginaSito.asp?ID=<%= srs("id_pagineSito") %>&ID_WEB=<%= srs("id_webs") %>', 'PaginaImporta', 680, 405, true)" title="Import pagina sito" <%= ACTIVE_STATUS %>>IMPORTA</a>
					</td>
				</tr>
				<% srs.movenext
			wend
			
			srs.close
			%>
		<% end if %>
	</table>
	</form>
</div>