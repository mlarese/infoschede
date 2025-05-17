<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
dim conn, rs, rsp, ID_DYN, ID_STAGE, sql, i, lingua, PaginaSito

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

PaginaSito = cInteger(request("PAGINA"))

'check dei permessi dell'utente
if PaginaSito>0 then
	if NOT index.content.ChkPrmF("tb_pagineSito", PaginaSito) then
		session("ERRORE") = "Non si possiedono i permessi per modificare la pagina." %>
		<script language="JavaScript">window.close()</script>
	<% end if
else
	if NOT index.ChkPrm(prm_pagine_altera, 0) then
		conn.close
		set conn = nothing %>
		<script language="JavaScript">window.close()</script>
	<% end if
end if

if request("conferma") <> "" then
	conn.BeginTrans
	
	sql = " SELECT * FROM tb_PagineSito "
	if PaginaSito>0 then
		sql = sql & " WHERE id_pagineSito=" & PaginaSito
	else
		sql = sql & " WHERE id_web=" & Session("AZ_ID")
	end if
	sql = sql + " ORDER BY nome_ps_IT, nome_ps_interno "
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
	
	while not rs.eof
		for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then
				ID_STAGE = rs("id_pagStage_" & lingua)
				'crea pagina visibile se mancante
				if cInteger(rs("id_pagDyn_" & lingua))<1 then
					ID_DYN = Create_page(conn, rs("nome_ps_" & lingua), rs("id_web"), rs("id_pagineSito"), lingua)
					rs("id_pagDyn_" & lingua) = ID_DYN
					rs.update
				else
					ID_DYN = rs("id_pagDyn_" & lingua)
				end if
				
				'copia dei dati della pagina e dei layers
				CALL Copy_page(conn, ID_STAGE, ID_DYN, false)
			end if
		next
		rs.movenext
	wend
	
	rs.close
end if
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - indice delle pagine - " & IIF(PaginaSito>0, "pubblica tutte le lingue", "pubblica tutte le pagine") %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">
				Pubblicazione di tutte le <%= IIF(PaginaSito>0, "lingue della pagina", "pagine del sito") %>
			</caption>
			<% 	if request("conferma") = "" then %>
			<tr>
				<td class="content_b" colspan="2">
					Attenzione: l'operazione sar&agrave; irreversibile
				</td>
			</tr>
			<tr>
				<td class="content" colspan="2">
					<table cellpadding="4" cellspacing="0" width="100%">
						<tr>
							<td class="label" colspan="2">Conferma l'operazione?</td>
						</tr>
						<tr>
							<td class="content_center" width="50%">
								<input type="submit" class="button" name="conferma" value="CONFERMA" tabindex="1" id="primo_elemento">
							</td>
							<td class="content_center">
								<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close()" tabindex="2">
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<% 	else 'copia a buon fine %>
			<%		if err.number=0 then
						conn.CommitTrans%>
			<tr>
				<td colspan="2" class="content_b">
					Pubblicazione eseguita correttamente
				</td>
			</tr>
			<tr>
				<td colspan="2" class="note">
					Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.
				</td>
			</tr>
			<script language="JavaScript">
				opener.location.reload(true);
				window.setTimeout("close();", 5000);
			</script>
			<%		else
						conn.RollBackTrans %>
			<tr>
				<td colspan="2" class="content_b">
					Errore nella copia: trasferimento annullato<br>
					<%= Err.Number & " - " & Err.Description %>
				</td>
			</tr>
			<%		end if %>
			<% 	end if %>
			<tr>
				<td class="footer" colspan="2">
					<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
				</td>
			</tr>
		</form>
	</table>
</div>
</body>
</html>

<%
conn.close
set rs = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
	PageOnLoad_FocusSet();
//-->
</script>