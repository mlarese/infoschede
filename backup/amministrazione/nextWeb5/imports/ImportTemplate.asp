<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="ImportPagineTools.asp" -->
<!--#INCLUDE FILE="../Tools_NextWeb5.asp" -->
<%
'--------------------------------------------------------
sezione_testata = "Import template" %>
<!--#INCLUDE FILE="../../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim dbPath, sql
dim conn, sconn, rs, srs, rsp, srsp, rsl, srsl, web_id
web_id = cIntero(request("ID_WEB"))
if web_id = 0 then
	web_id = cIntero(Session("AZ_ID"))
end if
dbPath = Application("IMAGE_PATH") & web_id & "\images\" & replace(request("source_import"), "/", "\")

set rs = Server.CreateObject("ADODB.RecordSet")
set srs = Server.CreateObject("ADODB.RecordSet")
set rsl = Server.CreateObject("ADODB.RecordSet")
set srsl = Server.CreateObject("ADODB.RecordSet")
set conn = Server.CreateObject("ADODB.Connection")
'set sconn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
'sconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";"
set sconn = Session("nw5_import_connection")

sql = "SELECT * FROM tb_pages WHERE " & SQL_IsTrue(sconn, "template") & " AND id_page = " & cIntero(request("ID"))
srs.open sql, sconn, adOpenStatic, adLockOptimistic, adCmdText
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Copia template</caption>
		<tr><th colspan="4">TEMPLATE SORGENTE</th></tr>
		<tr>
			<td class="label">
				template:
			</td>
			<td class="content" colspan="2">
				<%= srs("nomepage") %>
			</td>
		</tr>
		<tr><th colspan="4">TEMPLATE DESTINAZIONE</th></tr>
		<% if cIntero(request("IDDest")) = 0 AND request("tipoDest") <> "N" then 
			'scelta del template di destinazione
			%>
			<tr>
				<td class="label" rowspan="2">
					template:
				</td>
				<td class="content" colspan="2">
					<input type="radio" value="N" name="tipoDest" class="checkbox">
					Nuovo
				</td>
			</tr>
			<tr>
				<td class="content" style="width:100px;">
					<input type="radio" value="E" name="tipoDest" class="checkbox">
					Esistente: 
				</td>
				<td class="content">
					<%'sql = "SELECT ('N') AS id_page, ('NUOVO TEMPLATE') AS NAME, (0) AS ordine FROM tb_pages UNION " &_
						'	QryElencoTemplate("", true)
					sql = " SELECT * FROM tb_pages WHERE "& SQL_IsTrue(conn, "template")&" AND id_webs="&web_id
						
					CALL dropDown(conn, sql, "id_page", "nomepage", "IDDest", request("IDDest"), TRUE, " style=""95%"" ", LINGUA_ITALIANO)%>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					<input type="submit" class="button" name="mod" value="IMPORTA TEMPLATE">
				</td>
			</tr>
		<% else 
			dim PageDest
			conn.begintrans
			
			if cString(request("tipoDest")) = "N" then
				'inserimento nuovo template
				
				sql = "SELECT * FROM tb_pages"
				rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
				
				rs.AddNew
				rs("id_webs") = web_id
				CALL RecordsetCopyFields(srs, rs, "id_page")
				rs.Update
				
				PageDest = cIntero(rs("id_page"))
				
			else
				PageDest = cIntero(request("IDdest"))
				
				sql = "SELECT * FROM tb_pages WHERE id_page = " & PageDest
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				CALL RecordsetCopyFields(srs, rs, "id_page")
				rs.update
				
			end if
			
			%>
			<tr>
				<td class="label">
					template:
				</td>
				<td class="content" colspan="2">
					<%= rs("nomepage") %>
				</td>
			</tr>
			<tr>
				<td class="label">template importato:</td>
				<td>
					<table cellpadding="0" cellspacing="1" width="100%">
						<tr>
							<td class="content"><%= rs("id_page") %></td>
							<td class="content"><%= rs("lingua") %></td>
							<td class="content"><%= rs("nomepage") %></td>
							<%
							'cancella layers di destinazione
							sql = "DELETE FROM tb_layers WHERE id_pag=" & PageDest
							CALL conn.execute(sql)
							
							'importa layers
							sql = "SELECT * FROM tb_layers WHERE id_pag="
							rsl.open sql & rs("id_page"), conn, adOpenStatic, adLockOptimistic, adCmdText
							srsl.open sql & srs("id_page"), sconn, adOpenStatic, adLockOptimistic, adCmdText
							
							while not srsl.eof
								rsl.AddNew
								CALL RecordsetCopyFields(srsl, rsl, "id_lay,id_pag")
								rsl("id_pag") = rs("id_page")
								
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
