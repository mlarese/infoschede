<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
Imposta_Proprieta_Sito("ID")
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_strumenti_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Strumenti di analisi del sito"
dicitura.puls_new = "INDIETRO A SITI"
dicitura.link_new = "Siti.asp"
dicitura.scrivi_con_sottosez() 


dim conn, rs, sql
set rs = server.CreateObject("ADODB.recordset")
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")

sql = "SELECT * FROM tb_webs WHERE id_webs="& session("AZ_ID")
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext

%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Analisi statistica degli accessi</caption>
		<tr>
			<th colspan="2">STATISTICHE ACCESSI DELL'INDICE</th>
		</tr>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Visualizza l'indice generale con il numero di visite dei navigatori per ogni nodo.
			</td>
			<td class="content_center">
				<% if rs("statistiche_attive") then %>
					<a class="button_L2_block" href="SitoAnalisiStatIndice.asp" title="Visualizzazione statistiche attuali dell'indice." <%= ACTIVE_STATUS %>>
						STATISTICHE ATTUALI
					</a>
				<% else %>
					<a class="button_L2_block_disabled" title="Registrazione statistiche non attiva." <%= ACTIVE_STATUS %>>
						STATISTICHE NON ATTIVE
					</a>
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Visualizza lo storico delle statistiche di accesso all'indice generale.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="SitoAnalisiStoricoIndice.asp" title="Visualizzazione statistiche storiche dell'indice." <%= ACTIVE_STATUS %>>
					STORICO
				</a>
			</td>
		</tr>
		<tr>
			<th colspan="2">STATISTICHE ACCESSI DELLE PAGINE</th>
		</tr>
		<tr>
			<td class="label_no_width">
				Visualizza un elenco delle pagine con il relativo numero di visite dei navigatori per ogni lingua.
			</td>
			<td class="content_center">
				<% if rs("statistiche_attive") then %>
					<a class="button_L2_block" href="SitoAnalisiStatPagine.asp" title="Visualizzazione statistiche attuali delle pagine." <%= ACTIVE_STATUS %>>
						STATISTICHE ATTUALI
					</a>
				<% else %>
					<a class="button_L2_block_disabled" title="Registrazione statistiche non attiva." <%= ACTIVE_STATUS %>>
						STATISTICHE NON ATTIVE
					</a>
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Visualizza lo storico delle statistiche di visualizzazione delle pagine.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="SitoAnalisiStoricoPagine.asp" title="Visualizzazione statistiche storiche dell'indice." <%= ACTIVE_STATUS %>>
					STORICO
				</a>
			</td>
		</tr>
	</table>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption class="border">Analisi stato delle pagine</caption>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Esegue una analisi dello stato delle pagine del sito segnalando se le pagine sono state create correttamente, 
				se sono correttamente pubblicate o se necessitano di manutenzione.<br>
				ATTENZIONE: Se il numero delle pagine &egrave; alto l'analisi completa potrebbe richiedere alcuni minuti.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="SitoAnalisiPubblicazione.asp" title="Analisi stato creazione e pubblicazione delle pagine." <%= ACTIVE_STATUS %>>
					ANALISI STATO
				</a>
			</td>
		</tr>
	</table>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption class="border">Analisi utilizzo immagini</caption>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Esegue una analisi dell'utilizzo di tutte le immagini caricate, visualizzando, per ogni immagine, se &egrave; 
				utilizzata in almeno una pagina e, se richiesto, le pagine in cui viene utilizzata.<br>
				ATTENZIONE: Se il numero dei file &egrave; grande l'analisi potrebbe richiedere alcuni minuti.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="SitoAnalisiImmagini.asp" title="Analisi di utilizzo delle immagini." <%= ACTIVE_STATUS %>>
					ANALISI IMMAGINI
				</a>
			</td>
		</tr>
	</table>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption class="border">Analisi Ricerche</caption>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Esegue una analisi delle parole chiave utilizzate nelle ricerche.<br>
				ATTENZIONE: Se il numero dei file &egrave; grande l'analisi potrebbe richiedere alcuni minuti.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="SitoAnalisiRicerche.asp" title="Analisi di utilizzo delle immagini." <%= ACTIVE_STATUS %>>
					ANALISI RICERCHE
				</a>
			</td>
		</tr>
	</table>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Ottimizzazione dei motori di ricerca</caption>
		<tr>
			<th colspan="2">GESTIONE META TAG DELL'INDICE</th>
		</tr>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Gestione semplificata dei meta tag per i motori di ricerca applicati all'indice.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="IndexMetaTag.asp?FROM=strumenti" title="Gestione dei meta tag dell'indice." <%= ACTIVE_STATUS %>>
					GESTIONE META TAG
				</a>
			</td>
		</tr>
	</table>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Impostazioni generali</caption>
		<tr>
			<th colspan="2">FILTRI SUI CONTATORI</th>
		</tr>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Imposta i filtri generali di esclusione dai conteggi della navigazione, click ed attivit&agrave; su tutti i siti.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="SitoAnalisiFiltri.asp" title="Impostazioni dei filtri di esclusione da tutti i log di navigazione e click." <%= ACTIVE_STATUS %>>
					FILTRI DI ESCLUSIONE LOG
				</a>
			</td>
		</tr>
	</table>
	
	<%if instr(1, Session("LOGIN_4_LOG"), "NEXTAIM", vbTextCompare)>0 then %>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<caption>Import dati </caption>
			<tr>
				<th colspan="2">IMPORT DATI</th>
			</tr>
			<tr>
				<td class="label_no_width" style="width:79%;">
					Vai all'area di import
				</td>
				<td class="content_center">
					<a class="button_L2_block" href="imports/" title="" <%= ACTIVE_STATUS %>>
						AREA DI IMPORT
					</a>
				</td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<caption>Gestione link su pagine</caption>
			<tr>
				<th colspan="2">SOSTITUZIONE LINK</th>
			</tr>
			<tr>
				<td class="label_no_width" style="width:79%;">
					Apri la procedura di sostutizione dei link
				</td>
				<td class="content_center">
					<a class="button_L2_block" href="SostituisciLink.asp" title="" <%= ACTIVE_STATUS %>>
						SOSTITUISCI LINK
					</a>
				</td>
			</tr>
		</table>
	<% end if %>
</div>
</body>
</html>
<% 
rs.close
conn.close
set rs = nothing
set conn = nothing
%>