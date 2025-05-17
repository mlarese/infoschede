<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Gestione siti - indice delle pagine - strumenti" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------
%>
<%
dim conn, rs, sql, lingua, page_id_STAGE, page_id_DYN
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

lingua = request("LINGUA")

if request("PAGINA") <> "" then			'se sono una pagina
	sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & cIntero(request("PAGINA"))
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	page_id_STAGE = cInteger(rs("id_pagStage_" & lingua))
	page_id_DYN = cInteger(rs("id_pagDyn_" & lingua))
	
	'check dei permessi dell'utente
	if NOT index.content.ChkPrmF("tb_pagineSito", request("PAGINA")) then
		conn.close
		set conn = nothing 
		%>
<script language="JavaScript">
	window.close()
</script>
<%	end if
else									'sono un template
	page_id_stage = request("template")
	'check dei permessi dell'utente
	if NOT index.ChkPrm(prm_template_accesso, 0) then
		conn.close
		set conn = nothing %>
<script language="JavaScript">
	window.close()
</script>
<%	end if
end if

if request("conferma") <> "" AND request("op") <> "" then
	SELECT CASE request("op")
		CASE "RESET"
			response.redirect "SitoPagineCopia.asp?ID_S="& page_id_DYN &"&ID_D="& page_id_STAGE &"&lingua="& lingua &"&nome_lingua="& GetNomeLingua(lingua) &"&azione=RESET&conferma=true"
		CASE "COPIA"
			if request("PAGINA") <> "" then
				response.redirect "SitoPagineCopia.asp?ID_S="& rs("id_pagStage_it") &"&ID_D="& page_id_STAGE &"&lingua="& lingua &"&nome_lingua="& GetNomeLingua(lingua) &"&azione=COPIA&conferma=true"
			else
				response.redirect "SitoPagineCopia.asp?template=true&ID_S=&ID_D="& page_id_STAGE &"&lingua="& lingua &"&nome_lingua="& GetNomeLingua(lingua) &"&azione=COPIA&conferma=true"
			end if
		CASE "RIPULISCI"
			response.redirect "SitoPagineClear.asp?PAGINA="& page_id_STAGE
	END SELECT
end if
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">
			<% 	if request("PAGINA") <> "" then %>
				Pagina "<%= PaginaSitoNome(rs, request("LINGUA")) %>"
			<% 	else %>
				Template
			<% 	end if %>
			</caption>
			<% 	if request("op") <> "" then %>
			<input type="hidden" name="op" value="<%= request("op") %>">
			<tr>
				<td class="content_b" colspan="2">
					Attenzione: l'operazione sar&agrave; irreversibile
				</td>
			</tr>
			<tr>
				<td class="content" colspan="2">
					<table cellpadding="4" cellspacing="0" width="100%">
						<tr>
							<td class="label">Conferma l'operazione?</td>
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
			<% 	else %>
			<tr>
				<td class="label" style="width:80%;">
					Annulla le modifiche apportate alla pagina di lavoro dopo l'ultima pubblicazione.
				</td>
				<td class="content_center">
					<%if page_id_DYN>0 then%>
						<input type="submit" class="button" name="op" value="RESET" style="width:80%;">
					<%else%>
						<input type="submit" class="button" name="op" value="RESET" disabled style="width:80%;">
					<%end if%>
				</td>
			</tr>
			</form>
			<form action="SitoPagineCopia.asp?template=<%= request("template") %>&azione=COPIA&ID_D=<%= page_id_stage %>&PAGINASITO=<%= request("PAGINA") %>" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:80%;">
					Copia la pagina di lavoro da scegliere sostituendo la pagina di lavoro corrente.
				</td>
				<td class="content_center">
					<input type="submit" class="button" name="op" value="COPIA" style="width:80%;">
				</td>
			</tr>
			</form>
			<form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:80%;">
					Ripulisce la pagina da caratteri non validi che creano problemi nella visualizzazione o nella modifica della pagina.
				</td>
				<td class="content_center">
					<input type="submit" class="button" name="op" value="RIPULISCI" style="width:80%;">
				</td>
			</tr>
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