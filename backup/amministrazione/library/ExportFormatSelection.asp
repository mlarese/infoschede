<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="ExportTools.asp" -->
<%'--------------------------------------------------------
sezione_testata = "selezione formato di export dei dati" %>
<!--#INCLUDE FILE="Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim ConnString, SessionQueryName, format, SessionContattiQueryName
ConnString = request("conn")
SessionQueryName = request("query")
SessionContattiQueryName = SessionQueryName & "_contatti"
format = request("format")
%>

<div id="content_ridotto">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Selezione del formato di export dei dati</caption>
		<tr><th colspan="2">FORMATI DISPONIBILI</th></tr>
		<% 'export dei dati in formato EXCEL XP o successivo
		if format="" OR instr(1, format, FORMAT_ACCESS, vbTextCompare) then %>
			<tr>
				<td class="content">
					In formato Access 97 o successivo
				</td>
				<td class="content_center" style="width:30%;">
					<%CALL WRITE_EXPORT_LINK("ESPORTA", ConnString, SessionQueryName, FORMAT_ACCESS, true) %>
				</td>
			</tr>
		<% end if
		'export dei dati in formato EXCEL XP o successivo
		if format="" OR instr(1, format, FORMAT_EXCEL_XML, vbTextCompare) then %>
			<tr>
				<td class="content">
					In formato Excel XP / 2003 o successivo
				</td>
				<td class="content_center" style="width:30%;">
					<%CALL WRITE_EXPORT_LINK("ESPORTA", ConnString, SessionQueryName, FORMAT_EXCEL_XML, true) %>
				</td>
			</tr>
		<% end if 
		'export dei dati in formato EXCEL 97 / 2000
		if format="" OR instr(1, format, FORMAT_EXCEL_FILE, vbTextCompare) then %>
			<tr>
				<td class="content">
					In formato Excel 97 / 2000
				</td>
				<td class="content_center" style="width:30%;">
					<%CALL WRITE_EXPORT_LINK("ESPORTA", ConnString, SessionQueryName, FORMAT_EXCEL_FILE, true) %>
				</td>
			</tr>
		<% end if 
		'export dei dati in formato HTML
		if format="" OR instr(1, format, FORMAT_HTML, vbTextCompare) then %>
			<tr>
				<td class="content">
					In una tabella HTML
				</td>
				<td class="content_center" style="width:30%;">
					<%CALL WRITE_EXPORT_LINK("ESPORTA", ConnString, SessionQueryName, FORMAT_HTML, true) %>
				</td>
			</tr>
		<% end if
		'export dei dati in formato TXT
		if format="" OR instr(1, format, FORMAT_TXT, vbTextCompare) then %>
			<tr>
				<td class="content">
					In formato testuale con le colonne separate da "<strong>;</strong>"
				</td>
				<td class="content_center" style="width:30%;">
					<%CALL WRITE_EXPORT_LINK("ESPORTA", ConnString, SessionQueryName, FORMAT_TXT, true) %>
				</td>
			</tr>
		<% end if
		'export dei dati in formato XML
		if format="" OR instr(1, format, FORMAT_XML, vbTextCompare) then %>
			<tr>
				<td class="content">
					In formato XML con descrizione estesa.
				</td>
				<td class="content_center" style="width:30%;">
					<%CALL WRITE_EXPORT_LINK("ESPORTA", ConnString, SessionQueryName, FORMAT_XML, true) %>
				</td>
			</tr>
		<% end if
		'export dei dati (contatti) in una rubrica)
		if SessionContattiQueryName <> "" then %>
			<tr>
				<td class="content">
					Esporta la selezione di contatti come rubrica
				</td>
				<td class="content_center" style="width:30%;">
					<a style="width:100%; text-align:center; line-height:12px;" class="button"
					   title="Esporta la selezione di contatti come rubrica" <%= ACTIVE_STATUS %>
					   href="../NextCom/RubricheExport.asp?sql=<%= SessionContattiQueryName %>">
			   			SALVA RUBRICA
					</a>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="2"><a class="button" href="javascript:window.close();" title="annulla l'export dati" <%= ACTIVE_STATUS %>>ANNULLA</a></td>
		</tr>
	</table>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this)
</script>