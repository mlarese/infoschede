<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<% Server.ScriptTimeout = 2147483647 %>
<% '--------------------------------------------------------
sezione_testata = "gestione dei collegamenti all'indice" 
if request("from") = "tags" then
	testata_elenco_pulsanti = "INDIETRO"
	testata_elenco_href = "Tagga.asp?from=associazioni&co_F_key_id="& request("co_F_key_id") &"&co_F_table_id="& request("co_F_table_id")
end if


%>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<% 
'check dei permessi dell'utente
if NOT index.content.ChkPrm(index.content.GetID(request("co_F_table_id"), request("co_F_key_id"))) then %>
<script type="text/javascript">
	window.close()
</script>
<%
end if
'----------------------------------------------------- 


index.conn.BeginTrans()

CALL Index_UpdateItem(index.conn, index.GetTableName(request("co_F_table_id")), request("co_F_key_id"), true)

index.conn.CommitTrans()
index.content.co_F_table_id = request("co_F_table_id")
index.content.co_F_key_id = request("co_F_key_id")


CALL index.content.Associazioni()
%>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>