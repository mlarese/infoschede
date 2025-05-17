<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Titolo_sezione = "Tipi di import dati disponibili"%>
<!--#include file="Intestazione.asp"-->
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Import dati dei contatti</caption>
        <!-- 
		<tr><th colspan="2">Import in formato vCard</th></tr>
		<tr>
			<td class="note">
				La procedura importa tutte le vcard presenti nella cartella scelta e li inserisce direttamente nella rubrica selezionata.
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Import contatti in formato vCard" href="ImportVCards.asp" <%= ACTIVE_STATUS %>>IMPORT vCard</a>
			</td>
		</tr> -->
        
        <tr><th colspan="2">Import da file esportato dal NEXT-doc</th></tr>
        <tr>
			<td class="note">
				La procedura importa tutti i record utenti area riservata derivati da una esportazione da un altro applicativo NEXT-passport, controlla l'esistenza dello stesso utente in base alla chiave unica nel next-com (ID)  e li inserisce direttamente nella rubrica selezionata, abilita inoltre un utente di area riservata con i dati.
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Import contatti da file in formato NEXT-doc" href="ImportNextUser.asp" <%= ACTIVE_STATUS %>>IMPORTA DATI</a>
			</td>
		</tr>
	</table>
</div>