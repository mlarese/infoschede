<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<% Server.ScriptTimeout = 10000000 %>
<%
'check dei permessi
if NOT index.content.ChkPrmF("tb_pagineSito", request.Querystring("ID")) then %>
	<script language="JavaScript">
		window.close()
	</script>
<% end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoPagineMetaTagSalva.asp")
end if

dim conn, sql, rs, i, campo, lingua

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - indice delle pagine - descrizione e meta tag della pagina" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------

sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & cIntero(request.Querystring("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
			<caption>
				Modifica descrizione e meta tag della pagina &ldquo;<%= PaginaSitoNome(rs, "") %>&rdquo;:
			</caption>
			<tr><th colspan="2">DESCRIZIONE</th></tr>
			<tr>
				<td class="content notes" colspan="2">
					La descrizione della pagina verr&agrave; utilizzata anche come meta tag "description" per i motori di ricerca.
				</td>
			</tr>
			<% for each lingua in Application("LINGUE")
				if Session("lingua_" & lingua) then%>
					<tr>
						<td class="label_no_width" style="width:4%;"><img src="../grafica/flag_<%= lingua %>.jpg"></td>
						<td class="content"><textarea class="codice" rows="6" name="tft_page_description_<%= lingua %>"><%= rs("page_description_"& lingua) %></textarea></td>
					</tr>
				<% end if
			next %>
			<tr><th colspan="2">KEYWORDS</th></tr>
			<% for each lingua in Application("LINGUE")
				if Session("lingua_" & lingua) then%>
					<tr>
						<td class="label_no_width"><img src="../grafica/flag_<%= lingua %>.jpg"></td>
						<td class="content"><textarea class="codice" rows="4" name="tft_page_keywords_<%= lingua %>"><%= rs("page_keywords_"& lingua) %></textarea></td>
					</tr>
				<% end if
			next %>
			<tr>
				<td class="footer" colspan="2">
					<input type="button" class="button" name="chiudi" value="ANNULLA" onclick="window.close();">
					<input type="submit" class="button" name="salva" value="SALVA">
				</td>
			</tr>
		</form>
	</table>
</div>
</body>
</html>

<% rs.close
conn.close
set rs = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>