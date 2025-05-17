<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="ImportPagineTools.asp" -->
<!--#INCLUDE FILE="../Tools_NextWeb5.asp" -->
<%
'--------------------------------------------------------
sezione_testata = "Import plugin sito" %>
<!--#INCLUDE FILE="../../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim dbPath, sql
dim conn, sconn, rs, srs, rsp, srsp, rsl, srsl, i, lingua

dbPath = Application("IMAGE_PATH") & Session("AZ_ID") & "\images\" & replace(request("source_import"), "/", "\")

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

sql = "SELECT * FROM tb_paginesito where id_paginesito = " & cIntero(request("ID"))
srs.open sql, sconn, adOpenStatic, adLockOptimistic, adCmdText
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Copia plugin sito</caption>
		<tr><th colspan="4">COPIA PLUGIN</th></tr>
		<% if cString(request("mod")) = "" then %>
			<tr>
				<td class="content" colspan="3">
					Se vuoi allineare la situazione dei plugin clicca ALLINEA.
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					<input type="submit" class="button" name="mod" value="ALLINEA">
				</td>
			</tr>
		<% else
			
			if cIntero(request("ID_WEB"))>0 then
				conn.beginTrans
				sql = "SELECT * FROM tb_objects WHERE id_webs = "&cIntero(request("ID_WEB"))
				srsp.open sql, sconn, adOpenStatic, adLockOptimistic, adCmdText

				while not srsp.eof
					sql = "SELECT * FROM tb_objects WHERE id_webs = "&srsp("id_webs")&" AND name_objects LIKE '"&srsp("name_objects")&"'"
					rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
					if rs.eof then
						rs.AddNew
					end if
					rs("id_webs") = srsp("id_webs")
					rs("name_objects") = srsp("name_objects")
					rs("identif_objects") = srsp("identif_objects")
					rs("param_list") = srsp("param_list")
					rs("obj_insData") = srsp("obj_insData")
					rs("obj_insAdmin_id") = srsp("obj_insAdmin_id")
					rs("obj_modData") = srsp("obj_modData")
					rs("obj_modAdmin_id") = srsp("obj_modAdmin_id")
					rs("obj_type") = srsp("obj_type")
					
					''fare ciclo sulle lingue
					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
						lingua = Application("LINGUE")(i)
						rs("obj_html_"&lingua) = srsp("obj_html_"&lingua)
					next
	  
					rs.Update
					rs.close
					
					srsp.moveNext
				wend
				
				%>
				<tr>
					<td class="content_center ok" colspan="3">Import eseguito correttamente</td>
				</tr>
				<%
				conn.commitTrans
			else
				%>
				<tr>
					<td class="content_center ok" colspan="3">WEB_ID = 0</td>
				</tr>
				<%
			end if
			
			%>
			
			<tr>
				<td class="footer" colspan="3">
					<input type="button" class="button" name="chiudi" onclick="window.close()" value="CHIUDI">
				</td>
			</tr>
			<% 
		end if %>
	</table>
	</form>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>
