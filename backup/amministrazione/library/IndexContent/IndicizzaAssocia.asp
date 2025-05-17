<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<%

'--------------------------------------------------------
sezione_testata = "gestione dei collegamenti all'indice - nuovo"
testata_show_back = false
testata_elenco_pulsanti = "INDIETRO"
testata_elenco_href = "Indicizza.asp?co_F_key_id="& request("co_F_key_id") &"&co_F_table_id="& request("co_F_table_id") & "&MODE=" & request("MODE") %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------

'check dei permessi dell'utente
if NOT index.content.ChkPrm(index.content.GetID(request("co_F_table_id"), request("co_F_key_id"))) then %>
<script type="text/javascript">
	window.close()
</script>
<%
end if

dim rs
set rs = server.createobject("adodb.recordset")
if request("goto")<>"" then
	CALL GotoRecord(index.conn, rs, session("IDX_SQL"), "idx_id", "IndicizzaAssocia.asp?co_F_table_id="& request("co_F_table_id") &"&co_F_key_id="& request("co_F_key_id"))
end if

if request.form("salva") <> "" then
	index.conn.BeginTrans
	CALL index.Salva(request("ID"))
	if session("ERRORE") = "" then
		index.conn.CommitTrans
		response.redirect "Indicizza.asp?co_F_table_id="& request("co_F_table_id") &"&co_F_key_id="& request("co_F_key_id")
	else
		index.conn.RollbackTrans
	end if
end if
%>

<div id="content_ridotto">
	<%
	index.Modifica(request("ID"))
	set index = nothing
	%>
</div>
</body>
</html>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>