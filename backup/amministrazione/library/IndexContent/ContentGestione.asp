<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<%
'check dei permessi dell'utente
if NOT index.content.ChkPrm(index.content.GetID(request("co_F_table_id"), request("co_F_key_id"))) then %>
<script type="text/javascript">
	window.close()
</script>
<%
end if

index.content.co_F_table_id = request("co_F_table_id")
index.content.co_F_key_id = CIntero(request("co_F_key_id"))
if request.form("salva") <> "" then

	index.content.conn.BeginTrans
	
	CALL index.content.Salva(request("ID"))
		
	index.content.conn.CommitTrans
	
	if session("ERRORE") = "" then
        if lcase(request("from")) = "selezione" AND request.querystring("selection_disabled") = "" then
            response.redirect "ContentSeleziona.asp?co_F_key_id="& index.content.co_F_key_id &"&co_F_table_id="& index.content.co_F_table_id
        elseif lcase(request("from")) = "associazioni" then
    		response.redirect "Indicizza.asp?co_F_key_id="& index.content.co_F_key_id &"&co_F_table_id="& index.content.co_F_table_id & "&MODE=" & request("MODE")
		else %>
			<script type="text/javascript">
				window.close()
			</script>
        <% end if
	end if
end if


'--------------------------------------------------------
testata_show_back = false
if request.querystring("selection_disabled") = "" then
	testata_elenco_pulsanti = "INDIETRO"
end if
if request("from") = "selezione" then
    sezione_testata = "modifica contenuto"
    testata_elenco_href = "ContentSeleziona.asp?co_F_key_id="& index.content.co_F_key_id &"&co_F_table_id="& index.content.co_F_table_id
else
	sezione_testata = "gestione dei collegamenti all'indice - modifica"
	if request("from") = "tags" then
		testata_elenco_href = "Tagga.asp?co_F_key_id="& index.content.co_F_key_id &"&co_F_table_id="& index.content.co_F_table_id
	else
	    testata_elenco_href = "Indicizza.asp?co_F_key_id="& index.content.co_F_key_id &"&co_F_table_id="& index.content.co_F_table_id & "&MODE=" & request("MODE")
	end if
end if
%>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

if cInteger(request("ID"))>0 then
    CALL index.content.Modifica(request("ID"))
else
    CALL index.content.Modifica(index.content.GetID(index.content.co_F_table_id, index.content.co_F_key_id))
end if

%>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>