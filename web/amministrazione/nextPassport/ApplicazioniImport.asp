<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="ToolsApplicazioni.asp" -->
<%'----------------------------------------------------- 

if session("PASS_ADMIN") = "" then %>
<script type="text/javascript">this.window.close();</script>
<%
end if 

dim connSorg, rs, sql, importMode

if request("MODE") = "EXPORT" then
	importMode = false
	sezione_testata = "Export applicazioni" 
else
	importMode = true
	sezione_testata = "Import applicazioni" 
end if

%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<div id="content_ridotto">
<%
set connSorg = Server.CreateObject("ADODB.Connection")

if importMode then
	connSorg.open GetConfigurationConnectionstring()
else
	connSorg.open Application("DATA_ConnectionString")
end if

set rs = Server.CreateObject("ADODB.RecordSet")



if request("importa") <> "" then
	if request.form("selezione") <> "" then
		dim connDest, rsDest, rsa, sitiEsistenti
		set connDest = Server.CreateObject("ADODB.Connection")
		if importMode then
			connDest.open Application("DATA_ConnectionString")
		else
			connDest.open GetConfigurationConnectionstring()
		end if	
		set rsDest = Server.CreateObject("ADODB.RecordSet")
		connDest.BeginTrans
		
		sitiEsistenti = GetValueList(connDest, rsDest, "SELECT id_sito FROM tb_siti")
		
		'import siti
		sql = "SELECT * FROM tb_siti WHERE id_sito IN ("& ParseSQL(request.form("selezione"), adChar) &")"
		CALL CopyTables(connSorg, rs, sql, connDest, Array("tb_siti"), Array("id_sito"), Array(false), null, Array("id_sito"), false)
		
		'import siti tabelle
		if importMode then
			sql = " SELECT * FROM tb_siti_tabelle"& _
				  " WHERE tab_sito_id IN ("& ParseSQL(request.form("selezione"), adChar) &")"& _
				  " AND (tab_db_tipo = 0"& _
				  " OR tab_db_tipo = "& _
				  IIF(DB_Type(connDest) = DB_Access, "1", "2") &")"
		else
			sql = " SELECT *, " & IIF(DB_Type(connSorg) = DB_Access, "1", "2") & _
				  " AS tab_db_tipo FROM tb_siti_tabelle" & _
				  " WHERE tab_sito_id IN ("& ParseSQL(request.form("selezione"), adChar) & ")"				  
		end if
		CALL CopyTables(connSorg, rs, sql, connDest, Array("tb_siti_tabelle"), Array("tab_id"), null, null, Array("tab_titolo"), false)
		
		'import descrittori raggruppamenti
		CALL ConfigurationImport(connSorg, connDest, request.form("selezione"))
		
		
		'creazione rubriche per siti nuovi
		if importMode then
			sql = " SELECT * FROM tb_siti_descrittori d"& _
				  " INNER JOIN rel_siti_descrittori r ON d.sid_id = r.rsd_descrittore_id"& _
				  " WHERE NOT rsd_sito_id IN ("& sitiEsistenti &")"& _
				  " AND sid_tipo = "& adIDispatch
			rsDest.open sql, connDest, adOpenDynamic, adLockOptimistic
			rs.open "SELECT * FROM tb_rubriche WHERE 1=0", connDest, adOpenKeySet, adLockOptimistic
			while not rsDest.eof
				rs.AddNew
				rs("nome_rubrica") = rsDest("sid_nome_it")
				rs("locked_rubrica") = true
				rs("rubrica_esterna") = false
				rs.Update
				
				rsDest("rsd_valore_it") = rs("id_rubrica")
				rsDest.update
				rsDest.movenext
			wend
			rsDest.close
			rs.close
		end if
		
		connDest.CommitTrans
		set rs = nothing
		set rsDest = nothing
		connDest.close
		set connDest = nothing
	end if
	connSorg.close
	set connSorg = nothing %>
<script type="text/javascript">
	opener.window.location.reload();
	this.window.close();
</script>
<%
end if

sql = " SELECT * FROM tb_siti ORDER BY sito_nome"
rs.open sql, connSorg, adOpenStatic, adLockReadOnly
%>
<form action="" method="post" name="form1">
    <table cellspacing="1" cellpadding="0" class="tabella_madre">
	    <caption>Trovati n&ordm; <%= rs.recordcount %> applicazioni</caption>
        <% 	if not rs.eof then %>
		    <tr>
			    <th class="center" style="width:5%;">SCEGLI</th>
   				<th>APPLICAZIONE</th>
			</tr>
       	<% 		while not rs.eof %>
			<tr>
				<td class="content_center">
	        		<input class="checkbox" name="selezione" type="checkbox" value="<%= rs("id_sito") %>" />
				</td>
				<td class="content"><%= rs("sito_nome") %></td>
			</tr>
        <% 			rs.moveNext
				wend%>
			<tr>
				<td class="footer" colspan="2">
					<input type="submit" class="button" name="importa" value="<%=IIF(importMode, "IMPORTA", "ESPORTA")%>">
					<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
        <% 	else%>
		    <tr><td class="noRecords">Nessun record trovato</th></tr>
		<% 	end if %>
	</table>
</form>
</body>
</html>

<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>

<%
rs.close
connSorg.close
set rs = nothing
set connSorg = nothing
%>
