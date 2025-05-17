<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Titolo_sezione = "Tipi di import dati disponibili"%>
<!--#include file="Intestazione.asp"-->
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Import documenti Memo 2</caption>
        
        <tr><th colspan="2">Import documenti da una directory</th></tr>
        <tr>
			<td class="note">
				Scelta la directory di partenza vengono importate le sottocartelle come categorie e i file come documenti.
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="" href="ImportDocumentsFromFolder.asp" <%= ACTIVE_STATUS %>>IMPORTA DATI</a>
			</td>
		</tr>
		
	</table>
</div>