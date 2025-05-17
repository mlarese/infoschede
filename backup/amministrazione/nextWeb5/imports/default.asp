<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#include file="Intestazione.asp"-->
<%

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Import pagine e template"
dicitura.scrivi_con_sottosez()
%>


<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Import dati del NEXT-web</caption>
        <tr><th colspan="2">Import da altro database</th></tr>
        <tr>
			<td class="note">
				La procedura permette di scegliere le pagine di origine da copiare e di destinazione in cui copiare
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Import da altro database dbLayer" href="ImportPaginesito.asp" <%= ACTIVE_STATUS %>>IMPORTA PAGINE</a>
			</td>
		</tr>
		<!-- importa storico indice -->
		<tr>
			<td class="note">
				La procedura permette di importare lo storico dell'indice da una tabella access o un file excel
			</td>
			<td class="content_right" style="width:25%;">
				<a class="button" title="Import da altro database dbLayer" href="ImportStoricoIndice.asp" <%= ACTIVE_STATUS %>>IMPORTA STORICO INDICE</a>
			</td>
		</tr>
		<tr>
			<td class="note">
				Import Pagine VenetoInside
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Import da altro database dbLayer" href="ImportPagineVenetoInside.asp" <%= ACTIVE_STATUS %>>ESEGUI IMPORT</a>
			</td>
		</tr>
        <tr><th colspan="2">Import da stesso database</th></tr>
        <tr>
			<td class="note">
				La procedura permette di duplicare un sito, partendo dal sito di origine, impostando il sito di destinazione
			</td>
			<td class="content_right" style="width:19%;">
				<a class="button" title="Duplica sito" href="DuplicaSito.asp" <%= ACTIVE_STATUS %>>DUPLICA SITO</a>
			</td>
		</tr>
	</table>
</div>