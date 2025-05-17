<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
Reset_Proprieta_Sito()

'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoSalva.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Siti.asp"
dicitura.scrivi_con_sottosez() 

dim conn, sql, i, lingua
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfd_contRes" value="NOW">
	<input type="hidden" name="tfd_webs_modData_pagine" value="NOW">
	<input type="hidden" name="tfd_webs_modData_parametri" value="NOW">
	<input type="hidden" name="tfd_webs_modData_plugin" value="NOW">
	<input type="hidden" name="tfd_webs_modData_tabelle" value="NOW">
	<input type="hidden" name="tfn_contatore" value="0">
	<input type="hidden" name="tfn_contUtenti" value="0">
	<input type="hidden" name="tfn_contCrawler" value="0">
	<input type="hidden" name="tfn_contAltro" value="0">
    <input type="hidden" name="tfn_editor_guide_visibili" value="1">
    <input type="hidden" name="tft_editor_guide_colore" value="#000000">
    <input type="hidden" name="tfn_editor_guide_posizioni_visibili" value="1">
    <input type="hidden" name="tfn_editor_help_attivo" value="1">

	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Inserimento nuovo sito</caption>
		<tr><th colspan="4">DATI DEL SITO</th></tr>
		<tr>
			<td class="label_no_width" style="width:19%;" colspan="2">ID del sito:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tfn_id_webs" value="<%= IIF(CIntero(request("tfn_id_webs")) > 0, request("tfn_id_webs"), CIntero(GetValueList(conn, NULL, "SELECT MAX(id_webs) FROM tb_webs")) + 1) %>" maxlength="10" size="3">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label_no_width" colspan="2">Nome del sito:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_nome_webs" value="<%= request("tft_nome_webs") %>" maxlength="50" size="100">
				(*)
			</td>
		</tr>
		<tr>
			<input type="hidden" name="tfn_sito_mobile" value="<%= IIF(request("sito_mobile"), "1", "0") %>">
			<td class="label_no_width" colspan="2" rowspan="2">tipo di sito:</td>
			<td class="content" colspan="2">
				<input type="radio" name="sito_mobile" id="tipo_mobile" class="checkbox" <%= chk(request("sito_mobile")) %> onclick="ImpostaTipo()">
				sito per dispositivi mobili
				<img src="../grafica/mobile_icon.png" border="0" alt="Sito per dispositivi mobili.">
			</td>			
		</tr>
		<tr>			
			<td class="content" colspan="2">
				<input type="radio" name="sito_mobile" id="tipo_normale" class="checkbox" <%= chk(not request("sito_mobile")) %> onclick="ImpostaTipo()">
				sito normale
			</td>
		</tr>
		<tr>
			<input type="hidden" name="tfn_sito_in_costruzione" value="<%= IIF(request("tfn_sito_in_costruzione"), "1", "0") %>">
			<input type="hidden" name="tfn_sito_in_aggiornamento" value="<%= IIF(request("tfn_sito_in_aggiornamento"), "1", "0") %>">
			<td class="label_no_width" colspan="2" rowspan="3">stato del sito:</td>
			<td class="content" colspan="2">
				<input type="radio" name="stato_sito" id="stato_attivo" class="checkbox" <%= chk(request("tfn_sito_in_costruzione")<>"" AND cintero(request("tfn_sito_in_costruzione"))=0 AND cintero(request("tfn_sito_in_aggiornamento"))=0) %> onclick="ImpostaStato()">
				sito attivo
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="radio" name="stato_sito" id="stato_costruzione" class="checkbox" <%= chk(cintero(request("tfn_sito_in_costruzione"))>0 OR request("tfn_sito_in_costruzione")="") %> onclick="ImpostaStato()">
				sito in costruzione
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="radio" name="stato_sito" id="stato_aggiornamento" class="checkbox" <%= chk(cintero(request("tfn_sito_in_aggiornamento"))>0) %> onclick="ImpostaStato()">
				sito in aggiornamento
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			function ImpostaStato(){
				var stato_attivo = document.getElementById("stato_attivo");
				var stato_costruzione = document.getElementById("stato_costruzione");
				var stato_aggiornamento = document.getElementById("stato_aggiornamento");
				
				if (stato_attivo.checked){
					form1.tfn_sito_in_costruzione.value = 0;
					form1.tfn_sito_in_aggiornamento.value = 0;
				}
				else if (stato_costruzione.checked){
					form1.tfn_sito_in_costruzione.value = 1;
					form1.tfn_sito_in_aggiornamento.value = 0;
				}
				else if (stato_aggiornamento.checked){
					form1.tfn_sito_in_costruzione.value = 0;
					form1.tfn_sito_in_aggiornamento.value = 1;
				}
			}
			
			function ImpostaTipo(){
				var tipo_mobile = document.getElementById("tipo_mobile");
				var tipo_normale = document.getElementById("tipo_normale");				
				
				if (tipo_mobile.checked){					
					form1.tfn_sito_mobile.value = 1;
				}
				else if (tipo_normale.checked){					
					form1.tfn_sito_mobile.value = 0;
				}
			}
		</script>
        <tr>
			<td class="label_no_width" colspan="2" rowspan="2">accessibilit&agrave;:</td>
			<td class="content">
				<input type="radio" name="tfn_sito_accessibile" class="checkbox" <%= chk(cIntero(request("tfn_sito_accessibile"))>0) %> value="1">
				sito accessibile
			</td>
            <td class="content notes" style="width:66%;" rowspan="2"> 
                Permette di rendere aderente agli standard sull'accessibilit&agrave;
                definiti dal <a href="http://www.w3.org/WAI/" title="Web Accessibility Initiative (WAI)" target="_blank">WAI</a> con le normative 
                <a href="http://www.w3.org/TR/WAI-WEBCONTENT/" title="Web Content Accessibility Guidelines 1.0" target="_blank">WCAG</a> e di rispondere ai requisiti 
                definiti dalla <a href="http://www.pubbliaccesso.gov.it" title="" target="_blank">legge 04/2004 ( legge Stanca ) e normative seguenti</a> per l'accessibilit&agrave; dei servizi informatici.
            </td>
		</tr>
        <tr>
            <td class="content">
				<input type="radio" name="tfn_sito_accessibile" class="checkbox" <%= chk(cIntero(request("tfn_sito_accessibile"))=0) %> value="0">
				sito non accessibile
			</td>
		</tr>
		<tr>
			<td class="label_no_width" colspan="2" rowspan="2">statistiche interne:</td>
			<td class="content">
				<input type="radio" name="tfn_statistiche_attive" value="1" class="checkbox" <%= chk(request("tfn_statistiche_attive")="" OR cIntero(request("tfn_statistiche_attive"))=1) %>>
				registra contatori
			</td>
			<td class="content notes" rowspan="2">
				Attiva i contatori interni per la registrazione delle visite sulle pagine e sull'indice.
			</td>
		</tr>
		<tr>
			<td class="content">
				<input type="radio" name="tfn_statistiche_attive" value="0" class="checkbox" <%= chk(request("tfn_statistiche_attive")<>"" OR cIntero(request("tfn_statistiche_attive"))=0) %>>
				senza contatori
			</td>
		</tr>
		 <tr>
			<td class="label_no_width" colspan="2" rowspan="2">indicizzazione:</td>
			<td class="content">
				<input type="radio" name="tfn_sito_indicizzabile" class="checkbox" <%= chk(cIntero(request("tfn_sito_indicizzabile"))>0 OR CString(request("tfn_sito_indicizzabile"))="") %> value="1">
				sito indicizzabile
			</td>
            <td class="content notes" style="width:66%;" rowspan="2"> 
               Permette ai motori di ricerca (ad es. Google, Bing, Yahoo) di indicizzare i contenuti del sito.
			</td>
		</tr>
        <tr>
            <td class="content">
				<input type="radio" name="tfn_sito_indicizzabile" class="checkbox" <%= chk(cIntero(request("tfn_sito_indicizzabile"))=0 AND NOT CString(request("tfn_sito_indicizzabile"))="") %> value="0">
				sito non indicizzabile
			</td>
		</tr>
        <tr><th colspan="4">URL DEL SITO</th></tr>
		<tr>
            <td class="label_no_width" rowspan="3">URL:</td>
			<td class="label_no_width">principale:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_URL_base" value="<%= IIF(request.ServerVariables("REQUEST_METHOD") = "POST", request("tft_URL_base"), "http://") %>" maxlength="255" size="100">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label_no_width">sicuro:</td>
			<td class="content notes" colspan="2">
				<input type="text" class="text" name="tft_URL_secure" value="<%= request("tft_URL_secure") %>" maxlength="255" size="100"><br>
                Indirizzo HTTPS utilizzato solo nelle transazioni che trattano dati sensibili.
			</td>
		</tr>
		<tr>
			<td class="label_no_width">alternativo:</td>
			<td class="content notes" colspan="2">
				<input type="text" class="text" name="tft_URL_alternativo" value="<%= request("URL_alternativo")%>" maxlength="255" size="100"><br>
				Indirizzo alternativo per il passaggio tra la versione normale a quella mobile.
			</td>
		</tr>
		 <tr>
			<td class="label_no_width" style="width:24%;" colspan="2" rowspan="2">gestione url rewriting:</td>
			<td class="content">
				<input type="radio" name="tfn_URL_rewriting_attivo" class="checkbox" <%= chk(cIntero(request("tfn_URL_rewriting_attivo"))>0) %> value="1">
				attiva url &ldquo;statici&rdquo;
			</td>
            <td class="content notes" style="width:66%;" rowspan="2"> 
                L'attivazione degli url statici permette di ottenere degli indirizzi semplici ed ottimizzati per i motori di ricerca, ad esempio:<br>
				http:// &lt;nome dominio&gt; / &lt; nome sezione &gt; / &lt; nome pagina &gt;
            </td>
		</tr>
        <tr>
            <td class="content">
				<input type="radio" name="tfn_URL_rewriting_attivo" class="checkbox" <%= chk(cIntero(request("tfn_URL_rewriting_attivo"))=0) %> value="0">
				mantieni url &ldquo;dinamici&rdquo;
			</td>
		</tr>
    </table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="3" class="l2">URL ALTERNATIVI</th></tr>
		<tr>
			<td colspan="3">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label_no_width">
							L'associazione degli altri url &egrave; possibile dopo aver salvato.
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr><th colspan="3">GESTIONE LINGUE</th></tr>
		<tr>
			<td colspan="2" class="label_no_width" style="width:16%;">lingua iniziale:</td>
			<td class="content">
				<select name="tft_lingua_iniziale" id="tft_lingua_iniziale">
				<% 	for i = 0 to uBound(application("LINGUE")) %>
					<option value="<%= application("LINGUE")(i) %>" <%= IIF(request.form("tft_lingua_iniziale")=application("LINGUE")(i), "selected", "") %>>
						<%= GetNomeLingua(application("LINGUE")(i) )  %>
					</option>
				<% 	next %>
				</select>
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			function attiva_lingua(lingua, obj){
				var obj_titolo = eval('form1.tft_titolo_' + lingua)
				var obj_keywords = eval('form1.tft_meta_keywords_' + lingua)
				var obj_description = eval('form1.tft_meta_description_' + lingua)
				DisableControl(obj_titolo, !(obj.checked));
				DisableControl(obj_keywords, !(obj.checked));
				DisableControl(obj_description, !(obj.checked));
			}
		</script>
		<% for each lingua in Application("LINGUE")%>
			<tr>
				<td class="label_no_width" rowspan="2" style="width:4%;"><img src="../grafica/flag_<%= lingua %>.jpg" width="26" height="15" alt="" border="0"></td>
				<td class="label_no_width">attiva lingua:</td>
				<td class="content">
					<% if lingua <> LINGUA_ITALIANO then %>
					<input <%= chk(request.form("chk_lingua_"& lingua)<>"") %> class="checkbox" type="checkbox" name="chk_lingua_<%= lingua %>" onclick="attiva_lingua('<%= lingua %>', this)">
					<% else %>
					<input class="checkbox" type="checkbox" name="lingua_it" value="1" checked disabled>
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width">titolo pagine:</td>
				<td class="content"><input type="text" name="tft_titolo_<%= lingua %>" size="100" maxlength="255" class="text" value="<%= request("tft_titolo_"& lingua) %>"></td>
			</tr>
		<% next %>
		<tr>
			<td class="footer" colspan="3">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
conn.Close
set conn = nothing
%>