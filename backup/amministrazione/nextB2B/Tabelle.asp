<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tabelle di parametri"
dicitura.scrivi_con_sottosez() 
%> 
<div id="content">
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Tabelle di gestione dei prodotti</caption>
    <tr><th colspan="2">CATEGORIE E SOTTOCATEGORIE</th></tr>
    <tr>
        <td class="note">
            Definizione delle categorie ed eventuali livelli multipli di sottocategorie.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione categorie di prodotti" href="Categorie.asp" <%= ACTIVE_STATUS %>>CATEGORIE</a>
        </td>
    </tr>
    <tr><th colspan="2">CARATTERISTICHE TECNICHE</th></tr>
    <tr>
        <td class="note">
            Definizione delle caratteristiche dei prodotti. Attraverso l'associazione delle caratteristiche alle categorie di prodotti sar&agrave; possibile
            descrivere i prodotti di ogni categoria con le proprie caratteristiche.
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
    <tr><th colspan="2">MARCHI / PRODUTTORI</th></tr>
    <tr>
        <td class="note">
            Definizione dei marchi o produttori dei prodotti.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione marchi/produttori dei prodotti" href="Marchi.asp" <%= ACTIVE_STATUS %>>MARCHI</a>
        </td>
    </tr>
    <tr><th colspan="2">VARIANTI</th></tr>
    <tr>
        <td class="note">
            Definizione delle varianti per tutti i prodotti messi a catalogo.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione varianti" href="Varianti.asp" <%= ACTIVE_STATUS %>>VARIANTI</a>
        </td>
    </tr>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Tabelle di gestione dei prezzi</caption>
    <tr><th colspan="2">VALUTE</th></tr>
    <tr>
        <td class="note">
            Gestione delle valute dei clienti. La registrazione della valuta permette di far visualizzare al proprio cliente i prezzi nella sua valuta.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione valute" href="Valute.asp" <%= ACTIVE_STATUS %>>VALUTE</a>
        </td>
    </tr>
    <tr><th colspan="2">CATEGORIE I.V.A.</th></tr>
    <tr>
        <td class="note">
            Gestione delle categorie I.V.A. e relative percentuali da associare agli articoli.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione categorie i.v.a." href="CategorieIva.asp" <%= ACTIVE_STATUS %>>CATEGORIE I.V.A.</a>
        </td>
    </tr>
    <tr><th colspan="2">CLASSI DI SCONTO PER QUANTIT&Agrave;</th></tr>
    <tr>
        <td class="note">
            Gestione delle classi di sconto per quantit&agrave; acquistata. Tale sconto verr&agrave; applicato in cascata, 
            in fase di vendita e sulla base di quanto specificato nel listino o, in mancanza di specifiche, di quanto specificato nel prodotto.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione sconti per quantit&agrave;" href="ScontiQ.asp" <%= ACTIVE_STATUS %>>SCONTI PER QUANTIT&Agrave;</a>
        </td>
    </tr>
	<tr><th colspan="2">PROMOZIONI</th></tr>
	<tr>
        <td class="note">
            Elenco delle promozioni (all'interno cliccando sulla singola promozione si possono scegliere gli articoli da associare).
        </td>
        <td class="content_right">
            <a class="button" title="Promozioni" href="Promozioni.asp" <%= ACTIVE_STATUS %>>PROMOZIONI</a>
        </td>
    </tr>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Tabelle di gestione degli ordini</caption>
    <tr><th colspan="2">METODI DI CONSEGNA MERCE</th></tr>
	<tr>
        <td class="note">
            Gestione dei tipi di porti (metodi di consegna della merce).
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione porti - metodo consegna merce" href="Porti.asp" <%= ACTIVE_STATUS %>>PORTI</a>
        </td>
    </tr>
	<tr>
        <td class="note">
            Gestione dei tipi consegna (consegna parziale)
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione tipi consegna" href="TipiConsegna.asp" <%= ACTIVE_STATUS %>>TIPI CONSEGNA</a>
        </td>
    </tr>
	<tr><th colspan="2">TRASPORTATORI</th></tr>
	<tr>
        <td class="note">
            Gestione dei trasportatori.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione trasportatori" href="Trasportatori.asp" <%= ACTIVE_STATUS %>>TRASPORTATORI</a>
        </td>
    </tr>
	<tr><th colspan="2">STATI DI LAVORAZIONE DEGLI ORDINI</th></tr>
    <tr>
        <td class="note">
            Gestione degli stati di lavorazione degli ordini. Descrivono lo stato di lavorazione di ogni ordine cliente.
        </td>
        <td class="content_right" style="width:25%;">
            <a class="button" title="Gestione stati di lavorazione degli ordini" href="OrdiniStatiLavorazione.asp" <%= ACTIVE_STATUS %>>STATI DI LAVORAZIONE</a>
        </td>
    </tr>
	<tr><th colspan="2">SPESE DI SPEZIONE ORDINE</th></tr>
    <tr>
        <td class="note">
            Specifica le spese di spedizione in base all'area geografica di consegna della merce.
        </td>
        <td class="content_right">
            <a class="button" title="Specifica le spese di spedizione in base all'area geografica di consegna della merce." href="SpeseSpedizione.asp" <%= ACTIVE_STATUS %>>SPESE SPEDIZIONE ORDINE</a>
        </td>
    </tr>
	<tr><th colspan="2">METODI SPEDIZIONE ARTICOLO</th></tr>
    <tr>
        <td class="note">
            Specifica l'importo massimo e la quantit&agrave; massima oltre il quale le spese di spedizione dell'articolo vengono azzerate.
        </td>
        <td class="content_right">
            <a class="button" title="Specifica l'importo massimo e la quantit&agrave; massima oltre il quale le spese di spedizione dell'articolo vengono azzerate." href="SpeseSpedizioneArticolo.asp" <%= ACTIVE_STATUS %>>METODO SPEDIZIONE</a>
        </td>
    </tr>
	<tr><th colspan="2">MODALIT&Agrave; DI PAGAMENTO</th></tr>
    <tr>
        <td class="note">
            Specifica le diverse modalit&agrave; di pagamento attive nel sistema.
        </td>
        <td class="content_right">
            <a class="button" title="Specifica le modalità di pagamento e le eventuali spese aggiuntive." href="MetodiPagamento.asp" <%= ACTIVE_STATUS %>>MODALIT&Agrave; DI PAGAMENTO</a>
        </td>
    </tr>
    <tr><th colspan="2">TIPOLOGIE DI RIGA D'ORDINE</th></tr>
    <tr>
        <td class="note">
            Gestione delle tipologie di righe d'ordine per il collegamento delle informazioni aggiuntive appropriate ad ogni riga.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione tipologie di righe d'ordine" href="OrdiniRigheTipologie.asp" <%= ACTIVE_STATUS %>>TIPOLOGIE DI RIGA</a>
        </td>
    </tr>
    <tr><th colspan="2">INFORMAZIONI AGGIUNTIVE PER RIGA D'ORDINE</th></tr>
    <tr>
        <td class="note">
            Gestione delle informazioni aggiuntive per riga d'ordine, con relativa associazione alle tipologie di righe.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione informazioni aggiuntive per riga d'ordine" href="OrdiniRigheInfo.asp" <%= ACTIVE_STATUS %>>INFORMAZIONI PER RIGA</a>
        </td>
    </tr>
	<tr><th colspan="2">TIPOLOGIE DI FATTURAZIONE</th></tr>
    <tr>
        <td class="note">
            Gestione delle tipologie di fatturazione da applicare ad ogni riga.
        </td>
        <td class="content_right">
            <a class="button" title="Gestione delle tipologie di fatturazione" href="Fatturazioni.asp" <%= ACTIVE_STATUS %>>TIPOLOGIE DI FATTURAZIONE</a>
        </td>
    </tr>
    
</table>
</div>