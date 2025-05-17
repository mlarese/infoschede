<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim sql, conn

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tabelle di parametri"
dicitura.scrivi_con_sottosez() 
%> 
<div id="content">
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Tabelle di gestione delle schede</caption>
    <tr><th colspan="2">STATI DI LAVORAZIONE</th></tr>
    <tr>
        <td class="note">
            Gestione degli stati di lavorazione di una scheda.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione stati di lavorazione" href="SchedeStatiLavorazione.asp" <%= ACTIVE_STATUS %>>STATI DI LAVORAZIONE</a>
        </td>
    </tr>
	<tr><th colspan="2">ACCESSORI</th></tr>
    <tr>
        <td class="note">
            Gestione degli accessori che vengono consegnati assieme ai modelli da riparare.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione accessori" href="Accessori.asp" <%= ACTIVE_STATUS %>>ACCESSORI</a>
        </td>
    </tr>
	<tr><th colspan="2">ESITI INTERVENTI</th></tr>
    <tr>
        <td class="note">
            Gestione degli esiti degli interventi.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione esiti" href="Esiti.asp" <%= ACTIVE_STATUS %>>Esiti</a>
        </td>
    </tr>
	<tr><th colspan="2">CONTROLLI RIPARAZIONI</th></tr>
    <tr>
        <td class="note">
            Gestione dei controlli fatti durante le riparazioni delle schede.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione controlli" href="Descrittori.asp" <%= ACTIVE_STATUS %>>CONTROLLI</a>
        </td>
    </tr>
	<!--
    <tr><th colspan="2">CAUSALI DDT</th></tr>
    <tr>
        <td class="note">
            Gestione delle causali dei DDT.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione causali ddt" href="OrdiniStatiLavorazione.asp" <%= ACTIVE_STATUS %>>CAUSALI</a>
        </td>
    </tr>
	-->
</table>

<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Tabelle di gestione dei modelli</caption>
	<tr><th colspan="2">GUASTI</th></tr>
	<tr>
        <td class="note">
            Inserimento, modifica e cancellazione dei guasti.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione dei guasti" href="Problemi.asp" <%= ACTIVE_STATUS %>>GUASTI</a>
        </td>
    </tr>
	<tr><th colspan="2">MODELLI</th></tr>
    <tr>
        <td class="note">
            Inserimento, modifica e cancellazione dei modelli.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione modelli" href="Articoli.asp" <%= ACTIVE_STATUS %>>MODELLI</a>
        </td>
    </tr>
    <tr><th colspan="2">CATEGORIE E SOTTOCATEGORIE</th></tr>
    <tr>
        <td class="note">
            Definizione delle categorie ed eventuali livelli multipli di sottocategorie.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione categorie di modelli" href="Categorie.asp" <%= ACTIVE_STATUS %>>CATEGORIE</a>
        </td>
    </tr>
	<!--
    <tr><th colspan="2">CARATTERISTICHE TECNICHE</th></tr>
    <tr>
        <td class="note">
            Definizione delle caratteristiche dei modelli. Attraverso l'associazione delle caratteristiche alle categorie di modelli sar&agrave; possibile
            descrivere i modelli di ogni categoria con le proprie caratteristiche.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione caratteristiche tecniche" href="Caratteristiche.asp" <%= ACTIVE_STATUS %>>CARATTERISTICHE</a>
        </td>
    </tr>
    <tr><th colspan="2">GRUPPI DI CARATTERISTICHE TECNICHE</th></tr>
    <tr>
        <td class="note">
            Definizione dei gruppi di caratteristiche tecniche. Rappresentano le sezioni e le tipologie di caratteristiche.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione caratteristiche tecniche" href="CaratteristicheGruppi.asp" <%= ACTIVE_STATUS %>>GRUPPI DI CARATTERISTICHE</a>
        </td>
    </tr>
	-->
    <tr><th colspan="2">MARCHI / PRODUTTORI</th></tr>
    <tr>
        <td class="note">
            Definizione dei marchi o produttori dei modelli.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione marchi/produttori dei modelli" href="Marchi.asp" <%= ACTIVE_STATUS %>>MARCHI</a>
        </td>
    </tr>
</table>

<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Tabelle di gestione delle anagrafiche</caption>
	<tr><th colspan="2">CENTRI ASSISTENZA</th></tr>
    <tr>
        <td class="note">
            Gestione dei centri assistenza.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione centri assistenza" href="Agenti.asp" <%= ACTIVE_STATUS %>>CENTRI ASSISTENZa</a>
        </td>
    </tr>
    <tr><th colspan="2">TRASPORTATORI</th></tr>
    <tr>
        <td class="note">
            Gestione dei trasportatori.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione trasportatori" href="Clienti.asp?PROFILO=trasportatori" <%= ACTIVE_STATUS %>>TRASPORTATORI</a>
        </td>
    </tr>
	<tr><th colspan="2">COSTRUTTORI</th></tr>
    <tr>
        <td class="note">
            Gestione dei costruttori.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione costruttori" href="Clienti.asp?PROFILO=costruttori" <%= ACTIVE_STATUS %>>COSTRUTTORI</a>
        </td>
    </tr>	
</table>


<% if IsNextAim() then %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Strumenti di import dati: solo per utente NEXT-AIM</caption>
		<tr><th colspan="2">CLIENTI</th></tr>
		<tr>
			<td class="note">
				Import dei clienti
			</td>
			<td class="content_right" style="width:25%;">
				<a class="button" title="Import clienti" href="ImportClienti.asp" <%= ACTIVE_STATUS %>>IMPORT CLIENTI</a>
			</td>
		</tr>	
		<tr><th colspan="2">MODELLI</th></tr>
		<tr>
			<td class="note">
				Import dei modelli
			</td>
			<td class="content_right" style="width:25%;">
				<a class="button" title="Import modelli" href="ImportModelli.asp" <%= ACTIVE_STATUS %>>IMPORT MODELLI</a>
			</td>
		</tr>
		
		<tr><th colspan="2">SCHEDE</th></tr>
		<tr>
			<td class="note">
				Import schede di lavorazione
			</td>
			<td class="content_right" style="width:25%;">
				<a class="button" title="Import schede" href="ImportSchede.asp" <%= ACTIVE_STATUS %>>IMPORT Schede</a>
			</td>
		</tr>
		
		<tr><th colspan="2">RICAMBI</th></tr>
		<tr>
			<td class="note">
				Import dei ricambi
			</td>
			<td class="content_right" style="width:25%;">
				<a class="button" title="Import ricambi" href="ImportRicambi.asp" <%= ACTIVE_STATUS %>>IMPORT RICAMBI</a>
			</td>
		</tr>
		
		<tr><th colspan="2">DDT</th></tr>
		<tr>
			<td class="note">
				Import DDT
			</td>
			<td class="content_right" style="width:25%;">
				<a class="button" title="Import DDT" href="ImportDDT.asp" <%= ACTIVE_STATUS %>>IMPORT DDT</a>
			</td>
		</tr>
			
	</table>
<% end if %>		

</div>