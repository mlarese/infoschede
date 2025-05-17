<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
'check dei permessi
if NOT index.content.ChkPrmF("tb_pagineSito", request.Querystring("ID")) then %>
	<script language="JavaScript">
		window.close()
	</script>
<% end if

dim conn, sql, rsp, rs

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - indice delle pagine - elenco contenuti pubblicati" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------

sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & cIntero(request.Querystring("PAGINA"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

sql = " SELECT * FROM v_indice WHERE NOT (tab_name LIKE 'tb_paginesito') AND " + _
      " ( co_link_pagina_id=" & rs("id_pagineSito") & " OR idx_link_pagina_id = " & rs("id_pagineSito") & ")"
rsp.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>


<div id="content_ridotto">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
		<caption>
			Elenco dei nodi dell'indice e dei relativi contenuti pubblicati per mezzo della pagina
		</caption>
		<tr>
			<th>Contenuto pubblicato</th>
			<th class="center">Visibile</th>
			<th class="center">Principale</th>
		</tr>
		<% while not rsp.eof %>
			<tr>
				<td class="content">
					<% CALL index.WriteNodeLink(rsp, "", LINGUA_ITALIANO) %>
					<% CALL index.content.WriteTipoRS(rsp) %>
				</td>
				<td class="content_center">
					<input class="checkbox" type="checkbox" name="visibile" value="1" disabled <%= Chk(rsp("idx_visibile_assoluto")) %>>
				</td>
				<td class="content_center" <%= IIF(rsp("idx_principale"), "title=""Url principale di navigazione.""", "") %>>
					<input class="checkbox" type="checkbox" name="principale" value="1" disabled <%= Chk(rsp("idx_principale")) %>>
				</td>
			</tr>
			<% rsp.movenext
		wend %>
		<tr>
			<td class="footer" colspan="3">
				<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
			</td>
		</tr>
	</table>
</div>
</body>
</html>

<% rs.close
rsp.close
conn.close
set rs = nothing
set rsp = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>