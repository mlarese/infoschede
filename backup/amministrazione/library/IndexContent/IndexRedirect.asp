<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<%
'check dei permessi dell'utente
if NOT index.ChkPrm(prm_indice_accesso, 0) then %>
<script type="text/javascript">
	window.close()
</script>
<%
end if

'--------------------------------------------------------
sezione_testata = "gestione dei collegamenti all'indice - elenco indirizzi alternativi" %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, sql, ID, var, id_url_redir, nome_url_redir
ID = CIntero(request("ID"))

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


'SALVA
if request.form("salva") <> "" then
	if request.form("extt_riu_url") = "" then
		session("ERRORE") = "Valore obbligatorio."
	else
		conn.beginTrans
		
		id_url_redir = 0
		sql = "SELECT * FROM rel_index_url_redirect"
		id_url_redir = SalvaCampiEsterniAdvanced(conn, rs, sql, "riu_id", ID, "riu_idx_id", cIntero(request("IDX")), "", "riu_")
		
		if cIntero(id_url_redir) > 0 then
			sql = "SELECT riu_url FROM rel_index_url_redirect WHERE riu_id = " & id_url_redir
			nome_url_redir = GetValueList(conn, rs, sql)
			if inStr(nome_url_redir, ".") = 0 AND inStr(nome_url_redir, "?") = 0 AND inStr(Right(nome_url_redir,1),"/") = 0 then
				nome_url_redir = nome_url_redir & "/"
				sql = "UPDATE rel_index_url_redirect SET riu_url = '" & ParseSQL(nome_url_redir, adChar) & "' WHERE riu_id = " & id_url_redir
				response.write sql
				CALL conn.execute(sql, 0, adExecuteNoRecords)
			end if
		end if
		
		conn.commitTrans
		
		if session("ERRORE") = "" then
			response.redirect "IndexRedirect.asp?IDX="& request("IDX")
		end if
	end if
end if


sql = " SELECT * FROM rel_index_url_redirect"& _
	  " WHERE riu_idx_id = "& CIntero(request("IDX"))
rs.Open sql, conn, AdOpenStatic, adLockReadOnly, adCmdText
%>
<div id="content">
<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco indirizzi alternativi della voce "<%= GetValueList(conn, NULL, "SELECT co_titolo_it FROM v_indice WHERE idx_id = "& cIntero(request("IDX"))) %>"</caption>
		<tr>
			<th>URL</th>
			<th style="text-align: center; width: 15%;">LINGUA</th>
			<th class="center" style="width:20%;">OPERAZIONI</th>
		</tr>
		<%	while not rs.eof
				if ID = rs("riu_id") then 			'modifica 
		%>
		<tr>
			<td class="content">
				<input type="text" class="text" name="extt_riu_url" value="<%= rs("riu_url") %>" maxlength="<% if DB_Type(conn) = DB_ACCESS then response.write "255" else response.write "500" %>" style="width:95%;">&nbsp;(*)
			</td>
			<td class="content_center">
				<% CALL DropLingue(conn, NULL, "extt_riu_lingua", rs("riu_lingua"), true, false, "width:100px;") %>
			</td>
			<td class="Content_center" style="vertical-align: middle;">
				<input style="width:70px;" type="submit" class="button" name="salva" value="SALVA">
				<input style="width:70px;" type="button" class="button" name="annulla" value="ANNULLA" onclick="document.location='<%= GetPageName() & "?FROM="&request("FROM")&"&co_F_key_id="&request("co_F_key_id")&"&co_F_table_id="&request("co_F_table_id")&"&IDX="&request("IDX")%>';">
			</td>
		</tr>
		<% 		else 								'visualizza 
		%>
		<tr>
			<td class="content"><%= rs("riu_url") %></td>
			<td class="content_center"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= rs("riu_lingua") %>.jpg" width="26" height="15" alt="" border="0"></td>
			<td class="Content_center" style="font-size:1px;">
				<a class="button" href="?FROM=<%= request("FROM") %>&co_F_key_id=<%= request("co_F_key_id") %>&co_F_table_id=<%= request("co_F_table_id") %>&IDX=<%= request("IDX") %>&ID=<%= rs("riu_id") %>">
					MODIFICA
				</a>
				&nbsp;
				<a class="button" href="javascript:void(0);" onclick="OpenAutoPositionedScrollWindow('<%= GetLibraryPath() %>IndexContent/DeleteIndexRedirect.asp?ID=<%= rs("riu_id") %>', 'delete', 500, 300, false);" >
				    CANCELLA
				</a>
			</td>
		</tr>
		<% 		end if
				rs.moveNext
			wend %>
		<%	if ID = 0 then							'nuovo 
		%>
		<tr>
			<td class="content">
				<input type="text" class="text" name="extt_riu_url" value="<%= request("extt_riu_url") %>" maxlength="<% if DB_Type(conn) = DB_ACCESS then response.write "255" else response.write "500" %>" style="width:95%;">&nbsp;(*)
			</td>
			<td class="content_center">
				<% CALL DropLingue(conn, NULL, "extt_riu_lingua", request("extt_riu_lingua"), true, false, "width:100px;") %>
			</td>
			<td class="Content_center">
				<input style="width:70px;" type="submit" class="button" name="salva" value="AGGIUNGI">
				<input style="width:70px;" type="button" class="button" name="annulla" value="ANNULLA" onclick="document.form1.reset();">
			</td>
		</tr>
		<% 	end if %>
		<tr><td colspan="5" class="footer">(*) Campi obbligatori.</td></tr>
	</table>
</form>
</div>
</body>
</html>
<%
rs.close
conn.close
set rs = nothing
set conn = nothing %>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>