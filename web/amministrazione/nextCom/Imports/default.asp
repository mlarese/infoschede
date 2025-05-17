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
		</tr>
		
		-->
        
        <tr><th colspan="2">Import da file esportato dal NEXT-doc</th></tr>
        <tr>
			<td class="note">
				La procedura importa tutti i record derivati da una esportazione da un altro applicativo NEXT-doc o NEXT-com e li inserisce direttamente
                nella rubrica selezionata.
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Import contatti da file in formato NEXT-doc" href="ImportNextCom.asp" <%= ACTIVE_STATUS %>>IMPORTA DATI</a>
			</td>
		</tr>
		
		<tr><th colspan="2">Import (per aggiornamento) da file esportato dal NEXT-doc</th></tr>
        <tr>
			<td class="note">
				La procedura importa tutti i record derivati da una esportazione da un altro applicativo NEXT-doc o NEXT-com e li inserisce direttamente
                nella rubrica selezionata. In questo import viene confrontata la colonna PartitaIva del file excel con le partite iva gi&agrave; inserite tra i contatti,
				aggiornando i dati presenti in caso di corrispondenza o, in caso contrario, inserendo il nuovo contatto.<br>
				Script utilizzato per il sito CRM. Controllarlo prima di riutilizzarlo con altri db.
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Import per aggiornamento contatti da file in formato NEXT-doc" href="AggiornaNextCom.asp" <%= ACTIVE_STATUS %>>IMPORTA DATI</a>
			</td>
		</tr>
		
		<tr><th colspan="2">Import da tabella SQL server</th></tr>
        <tr>
			<td class="note">
				La procedura importa tutti i record presenti in una tabella SQL server.
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Import contatti da tabella SQL" href="ImportSQLCom.asp" <%= ACTIVE_STATUS %>>IMPORTA DATI</a>
			</td>
		</tr>
		<tr><th colspan="2">Import dati nextCom e nextPassport</th></tr>
        <tr>
			<td class="note">
				La procedura importa tutti i record presenti nei due applicativi da una connessione specifica
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Import" href="ImportNextComConnection.asp" <%= ACTIVE_STATUS %>>IMPORTA DATI</a>
			</td>
		</tr>
	</table>
</div>