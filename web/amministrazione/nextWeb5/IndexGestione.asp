<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_indice_accesso, 0))

'eventuale redirect a raggruppamento
if request("ID") <> "" then
	if CIntero(GetValueList(index.conn, NULL, " SELECT COUNT(*) FROM (tb_contents c"& _
											  " INNER JOIN tb_contents_index i ON c.co_id = i.idx_content_id)"& _
											  " INNER JOIN tb_siti_tabelle t ON c.co_F_table_id = t.tab_id"& _
											  " WHERE tab_titolo = '"& tabRaggruppamento &"' AND tab_name = '" & tabRaggruppamentoTable & "'"& _
											  " AND idx_id = "& cIntero(request("ID")))) = 1 then
		response.redirect "IndexRaggruppamentoGestione.asp?ID="& request("ID") &"&FROM="& request("FROM")
	end if
end if

if request.form("salva") <> "" then
	index.conn.BeginTrans
	CALL index.Salva(request("ID"))
	
	if session("ERRORE") = "" then
		index.conn.CommitTrans
		
		if request.querystring("SOTTO") <> "" then		'vengo da voci collegate 
		%>
<script type="text/javascript">
	if (opener) {
		opener.window.location.reload()
		window.close()
	}
</script>
<%		else
			response.redirect IIF(request("FROM")=FROM_ALBERO, "IndexAlbero.asp", "IndexGenerale.asp")
		end if
	else
		index.conn.RollbackTrans
	end if
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
if request("ID") = "" then
	dicitura.sezione = "Indice generale - nuova voce"
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = IIF(request("FROM")=FROM_ALBERO, "IndexAlbero.asp", "IndexGenerale.asp")
else
	dicitura.sezione = "Indice generale - modifica voce"
	dicitura.puls_new = "INDIETRO;VOCI COLLEGATE"
	dicitura.link_new = IIF(request("FROM")=FROM_ALBERO, "IndexAlbero.asp", "IndexGenerale.asp") &";IndexSottosezioni.asp?FROM="& request("FROM") &"&ID="& request("ID")
end if
dicitura.scrivi_con_sottosez()
%>

<div id="content">
	<%
	index.Modifica(request("ID"))
	set index = nothing
	%>
</div>
</body>
<% if request("OPEN")<>"" then %>
<script language="JavaScript" type="text/javascript"> 
<!--
	FitWindowSize(this);
	PageOnLoad_FocusSet();
//-->
</script>
<% end if %>
</html>