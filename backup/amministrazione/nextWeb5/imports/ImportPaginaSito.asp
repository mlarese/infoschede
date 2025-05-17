<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="ImportPagineTools.asp" -->
<!--#INCLUDE FILE="../Tools_NextWeb5.asp" -->
<%
'--------------------------------------------------------
sezione_testata = "Import pagina sito" %>
<!--#INCLUDE FILE="../../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim dbPath, sql
dim conn, sconn, rs, srs, rsp, srsp, rsl, srsl, web_id, new_id_template
web_id = cIntero(request("ID_WEB"))
if web_id = 0 then
	web_id = cIntero(Session("AZ_ID"))
end if
dbPath = Application("IMAGE_PATH") & web_id & "\images\" & replace(request("source_import"), "/", "\")

set rs = Server.CreateObject("ADODB.RecordSet")
set srs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")
set srsp = Server.CreateObject("ADODB.RecordSet")
set rsl = Server.CreateObject("ADODB.RecordSet")
set srsl = Server.CreateObject("ADODB.RecordSet")
set conn = Server.CreateObject("ADODB.Connection")
'set sconn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set sconn = Session("nw5_import_connection")
'sconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";"

sql = "SELECT * FROM tb_paginesito WHERE id_paginesito = " & cIntero(request("ID"))
srs.open sql, sconn, adOpenStatic, adLockOptimistic, adCmdText
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Copia pagina sito</caption>
		<tr><th colspan="4">PAGINA SORGENTE</th></tr>
		<tr>
			<td class="label">
				pagina:
			</td>
			<td class="content" colspan="2">
				<%= srs("nome_ps_IT") %>
			</td>
		</tr>
		<tr><th colspan="4">PAGINA DESTINAZIONE</th></tr>
		<% if cIntero(request("IDDest")) = 0 AND request("tipoDest") <> "N" then 
			'scelta della pagina di destinazione
			%>
			<tr>
				<td class="label" rowspan="2">
					pagina:
				</td>
				<td class="content" colspan="2">
					<input type="radio" value="N" name="tipoDest" class="checkbox">
					Nuova
				</td>
			</tr>
			<tr>
				<td class="content">
					<input type="radio" value="E" name="tipoDest" class="checkbox">
					Esistente: 
				</td>
				<td class="content">
					<% CALL DropDownPages(conn, "form1", "380", web_id, "IDDest", request("IDDest"), false, false) %>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					<input type="submit" class="button" name="mod" value="IMPORTA PAGINA">
				</td>
			</tr>
		<% else 
			dim PageDest
			conn.begintrans
			
			if request("tipoDest") = "N" then
				'inserimento nuova pagina
				
				sql = "SELECT * FROM tb_paginesito"
				rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText

				rs.AddNew
				rs("id_web") = web_id
				CALL RecordsetCopyFields(srs, rs, "id_pagineSito, id_web, " & replace(FieldLanguageList("id_pagDyn_;id_pagStage_"), ";", ","))			
				rs.Update
				sql = "SELECT * FROM tb_paginesito WHERE id_pagineSito="&rs("id_pagineSito")
				rs.close
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				CALL Ceck_page_exists(conn, rs)
				rs.Update
				
				PageDest = rs("id_paginesito")
				
			else
				PageDest = cIntero(request("IDdest"))
				
				sql = "SELECT * FROM tb_paginesito WHERE id_paginesito = " & PageDest
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				CALL RecordsetCopyFields(srs, rs, "id_pagineSito, id_web, " & replace(FieldLanguageList("id_pagDyn_;id_pagStage_"), ";", ","))
				rs.update
				
			end if
			
			%>
			<tr>
				<td class="label">
					pagina:
				</td>
				<td class="content" colspan="2">
					<%= rs("nome_ps_IT") %>
				</td>
			</tr>
			<tr>
				<td class="label">Pagine importate:</td>
				<td>
					<table cellpadding="0" cellspacing="1" width="100%">
						<% 
						dim field
						for each field in split(FieldLanguageList("id_pagStage_"), ";")
							if cIntero(rs(field))>0 then
								'se la pagina esiste: copia ogni pagina di stage
								sql = "SELECT * FROM tb_pages WHERE id_page="
								rsp.open sql & rs(field), conn, adOpenStatic, adLockOptimistic, adCmdText
								srsp.open sql & srs(field), sconn, adOpenStatic, adLockOptimistic, adCmdText
								
								CALL RecordsetCopyFields(srsp, rsp, "id_page, id_webs, id_template, id_PaginaSito")
								
								sql = "SELECT nomepage FROM tb_pages WHERE id_page="&cIntero(srsp("id_template")) 'recupero il nome del template della pagina dal sito sorgente		
								sql = "SELECT id_page FROM tb_pages WHERE "&SQL_IsTrue(sconn, "template")&" AND id_webs="&web_id&" AND nomepage LIKE '"&GetValueList(sconn,NULL,sql)&"'" 'recupero il corrispondente template nel sito di destinazione					
								new_id_template = cIntero(GetValueList(conn,NULL,sql))
								if new_id_template > 0 then
									rsp("id_template") = new_id_template 'setto l'id template
								end if
								rsp.update%>
								<tr>
									<td class="content"><%= rsp("id_page") %></td>
									<td class="content"><%= rsp("lingua") %></td>
									<td class="content"><%= rsp("nomepage") %></td>
									<%
									'cancella layers di destinazione
									sql = "DELETE FROM tb_layers WHERE id_pag=" & rs(field)
									CALL conn.execute(sql)
									
									'importa layers
									sql = "SELECT * FROM tb_layers WHERE id_pag="
									rsl.open sql & rs(field), conn, adOpenStatic, adLockOptimistic, adCmdText
									srsl.open sql & srs(field), sconn, adOpenStatic, adLockOptimistic, adCmdText
									
									while not srsl.eof
										rsl.AddNew
										CALL RecordsetCopyFields(srsl, rsl, "id_lay,id_pag")
										rsl("id_pag") = rs(field)
										
										if cIntero(srsl("id_objects"))>0 then 'se il layer che sto copiando è un oggetto
											sql = "SELECT id_objects FROM tb_objects WHERE id_webs="&web_id&" AND name_objects LIKE '"&srsl("nome")&"' "
											rsl("id_objects") = cIntero(GetValueList(conn, NULL, sql))
										end if
										
										rsl.update
										srsl.movenext
									wend
									%>
									<td class="content">layers: <%= srsl.recordcount %></td>
								</tr>
								<% rsl.close
								srsl.close
								rsp.close
								srsp.close
							end if
						next
						%>
					</table>
				</td>
			</tr>
			<tr>
				<td class="content_center ok" colspan="3">Import eseguito correttamente</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					<input type="button" class="button" name="chiudi" onclick="window.close()" value="CHIUDI">
				</td>
			</tr>
			<% 
			rs.close
			conn.committrans
		end if %>
	</table>
	</form>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>
