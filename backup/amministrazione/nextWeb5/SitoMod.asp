<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
Imposta_Proprieta_Sito("ID")
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoSalva.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Siti.asp"
dicitura.scrivi_con_sottosez() 

dim conn, connData, rs, sql, i, lingua, rsp
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set connData = Server.CreateObject("ADODB.Connection")
connData.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
set rsp = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("WEB_SITI_SQL"), "id_webs", "SitoMod.asp")
end if

sql = "SELECT * FROM tb_webs WHERE id_webs="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<%if cString(rs("webs_modData_pagine"))="" then %>
		<input type="hidden" name="tfd_webs_modData_pagine" value="NOW">
	<% end if %>
	<%if cString(rs("webs_modData_parametri"))="" then %>
		<input type="hidden" name="tfd_webs_modData_parametri" value="NOW">
	<% end if %>
	<%if cString(rs("webs_modData_plugin"))="" then %>
		<input type="hidden" name="tfd_webs_modData_plugin" value="NOW">
	<% end if %>
	<%if cString(rs("webs_modData_tabelle"))="" then %>
		<input type="hidden" name="tfd_webs_modData_tabelle" value="NOW">
	<% end if %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del sito</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="sito precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="sito successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label_no_width" style="width:20%;" colspan="2">nome del sito:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_nome_webs" value="<%= rs("nome_webs") %>" maxlength="50" size="50">
				(*)
			</td>
		</tr>
		<tr>
			<input type="hidden" name="tfn_sito_mobile" value="<%= IIF(rs("sito_mobile"), "1", "0") %>">
			<td class="label_no_width" colspan="2" rowspan="2">tipo di sito:</td>
			<td class="content" colspan="2">
				<input type="radio" name="sito_mobile" id="tipo_mobile" class="checkbox" <%= chk(rs("sito_mobile")) %> onclick="ImpostaTipo()">
				sito per dispositivi mobili
				<img src="../grafica/mobile_icon.png" border="0" alt="Sito per dispositivi mobili.">
			</td>
					
		</tr>
		<tr>			
			<td class="content" colspan="2">
				<input type="radio" name="sito_mobile" id="tipo_normale" class="checkbox" <%= chk(not rs("sito_mobile")) %> onclick="ImpostaTipo()">
				sito normale
			</td>
		</tr>
		<tr>
			<input type="hidden" name="tfn_sito_in_costruzione" value="<%= IIF(rs("sito_in_costruzione"), "1", "0") %>">
			<input type="hidden" name="tfn_sito_in_aggiornamento" value="<%= IIF(rs("sito_in_aggiornamento"), "1", "0") %>">
			<td class="label_no_width" colspan="2" rowspan="3">stato del sito:</td>
			<td class="content" colspan="2">
				<input type="radio" name="stato_sito" id="stato_attivo" class="checkbox" <%= chk(not rs("sito_in_costruzione") AND not rs("sito_in_aggiornamento")) %> onclick="ImpostaStato()">
				sito attivo
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="radio" name="stato_sito" id="stato_costruzione" class="checkbox" <%= chk(rs("sito_in_costruzione")) %> onclick="ImpostaStato()">
				sito in costruzione
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="radio" name="stato_sito" id="stato_aggiornamento" class="checkbox" <%= chk(rs("sito_in_aggiornamento")) %> onclick="ImpostaStato()">
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
			<td class="content" style="width:20%;">
				<input type="radio" name="tfn_sito_accessibile" class="checkbox" <%= chk(rs("sito_accessibile")) %> value="1">
				sito accessibile
			</td>
            <td class="content notes" style="width:70%;" rowspan="2"> 
                Permette di rendere aderente agli standard sull'accessibilit&agrave;
                definiti dal <a href="http://www.w3.org/WAI/" title="Web Accessibility Initiative (WAI)" target="_blank">WAI</a> con le normative 
                <a href="http://www.w3.org/TR/WAI-WEBCONTENT/" title="Web Content Accessibility Guidelines 1.0" target="_blank">WCAG</a> e di rispondere ai requisiti 
                definiti dalla <a href="http://www.pubbliaccesso.gov.it" title="" target="_blank">legge 04/2004 ( legge Stanca ) e normative seguenti</a> per l'accessibilit&agrave; dei servizi informatici.
            </td>
		</tr>
        <tr>
            <td class="content">
				<input type="radio" name="tfn_sito_accessibile" class="checkbox" <%= chk(not rs("sito_accessibile")) %> value="0">
				sito non accessibile
			</td>
		</tr>
		<input type="hidden" name="old_statistiche_attive" value="<%=IIF(rs("statistiche_attive"), 1, 0)%>">
		<tr>
			<td class="label_no_width" colspan="2" rowspan="2">statistiche interne:</td>
			<td class="content">
				<input type="radio" name="tfn_statistiche_attive" value="1" class="checkbox" <%= chk(rs("statistiche_attive")) %>>
				registra contatori
			</td>
			<td class="content notes" rowspan="2">
				Attiva i contatori interni per la registrazione delle visite sulle pagine e sull'indice.
			</td>
		</tr>
		<tr>
			<td class="content">
				<input type="radio" name="tfn_statistiche_attive" value="0" class="checkbox" <%= chk(not rs("statistiche_attive")) %>>
				senza contatori
			</td>
		</tr>
		<tr>
			<td class="label_no_width" colspan="2" rowspan="2">indicizzazione:</td>
			<td class="content">
				<input type="radio" name="tfn_sito_indicizzabile" class="checkbox" <%= chk(rs("sito_indicizzabile")) %> value="1">
				sito indicizzabile
			</td>
            <td class="content notes" style="width:70%;" rowspan="2"> 
				Permette ai motori di ricerca (ad es. Google, Bing, Yahoo) di indicizzare i contenuti del sito.
            </td>
		</tr>
        <tr>
            <td class="content">
				<input type="radio" name="tfn_sito_indicizzabile" class="checkbox" <%= chk(not rs("sito_indicizzabile")) %> value="0">
				sito non indicizzabile
			</td>
		</tr>
        <tr><th colspan="4">URL PRINCIPALI</th></tr>
		<tr>
             <td class="label_no_width" rowspan="5">URL:</td>
			<td class="label_no_width">principale:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_URL_base" value="<%= IIF(cString(rs("URL_base"))<>"", rs("URL_base"), "http://") %>" maxlength="255" size="100">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label_no_width" rowspan="2">sicuro:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_URL_secure" value="<%= rs("URL_secure")%>" maxlength="255" size="100"><br>
			</td>
		</tr>
		<tr>
			<td class="note" colspan="2">Indirizzo HTTPS utilizzato solo nelle transazioni che trattano dati sensibili.</td>
		</tr>
		<tr>
			<td class="label_no_width" rowspan="2">alternativo:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_URL_alternativo" value="<%= rs("URL_alternativo")%>" maxlength="255" size="100"><br>
			</td>
		</tr>
		<tr>
			<td class="note" colspan="2">Indirizzo alternativo per il passaggio tra la versione normale a quella mobile.</td>
		</tr>
		<tr>
			<td class="label_no_width" colspan="2" rowspan="2">gestione url rewriting:</td>
			<td class="content">
				<input type="radio" name="tfn_URL_rewriting_attivo" class="checkbox" <%= chk(rs("URL_rewriting_attivo")) %> value="1">
				attiva url &ldquo;statici&rdquo;
			</td>
            <td class="content notes" style="width:66%;" rowspan="2"> 
                L'attivazione degli url statici permette di ottenere degli indirizzi semplici ed ottimizzati per i motori di ricerca, ad esempio:<br>
				http:// &lt;nome dominio&gt; / &lt; nome sezione &gt; / &lt; nome pagina &gt;
            </td>
		</tr>
        <tr>
            <td class="content">
				<input type="radio" name="tfn_URL_rewriting_attivo" class="checkbox" <%= chk(NOT rs("URL_rewriting_attivo")) %> value="0">
				mantieni url &ldquo;dinamici&rdquo;
			</td>
		</tr>
		<tr><th colspan="4">URL ALTERNATIVI</th></tr>
		<tr>
			<td colspan="4">
				<% sql = " SELECT * FROM tb_webs_directories WHERE dir_web_id=" & cIntero(request("ID")) & _
						 " ORDER BY dir_url "
				rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label_no_width" style="width:74%">
							<% if rsp.eof then %>
								Nessun url inserito.
							<% else %>
								Trovati n&ordm; <%= rsp.recordcount %> record
							<% end if %>
						</td>
						<td colspan="3" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra l'inserimento di un url alternativo per il sito" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedWindow('SitoUrlNew.asp?WEB_ID=<%= request("ID") %>', 'Url_nuovo', 640, 430)">
								NUOVO URL
							</a>
						</td>
					</tr>
					<% if not rsp.eof then %>
						<tr>
							<th class="L2">URL COMPLETO</th>
							<th class="l2_center" width="18%">SERVIZI ATTIVATI</th>
							<th class="l2_center" width="16%" colspan="2">OPERAZIONI</th>
						</tr>
						<%while not rsp.eof %>
							<tr>
								<td class="content">
									<%= rsp("dir_url") %>
								</td>
								<td class="content_center">
									<% if rsp("dir_google_maps_key")<>"" then %>
										<span class="note" title="chiave: <%= rsp("dir_google_maps_key") %>">&nbsp;Google Maps</span>
									<% end if %>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica dei dati dell'url." <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedScrollWindow('SitoUrlMod.asp?ID=<%= rsp("dir_id") %>&WEB_ID=<%= request("ID") %>', 'Url_<%= rsp("dir_id") %>', 640, 430, true)">
										MODIFICA
									</a>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione dell'url alternativo"
									   onclick="OpenDeleteWindow('URL','<%= rsp("dir_id") %>');">
										CANCELLA
									</a>
								</td>
							</tr>
							<%rsp.MoveNext
						wend
					end if%>
				</table>
			<% rsp.close %>
			</td>
		</tr>
		
		<tr><th colspan="4">DOMINI AGGIUNTIVI PER SITO MULTIDOMINIO</th></tr>
		<tr>
			<td colspan="4">
				<% sql = " SELECT * FROM tb_webs_domini WHERE dom_web_id = " & cIntero(request("ID")) & _
						 " ORDER BY dom_ordine, dom_url "
				rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label_no_width" colspan="7">
							<% if rsp.eof then %>
								Nessun dominio aggiuntivo trovato.
							<% else %>
								Trovati n&ordm; <%= rsp.recordcount %> record
							<% end if %>
							<span style="float:right;">
								<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra l'inserimento di un dominio aggiuntivo per il sito" <%= ACTIVE_STATUS %>
								   onclick="OpenAutoPositionedWindow('SitoDominioNew.asp?WEB_ID=<%= request("ID") %>', 'Dominio_nuovo', 640, 430)">
									NUOVO DOMINIO AGGIUNTIVO
								</a>
							</span>
						</td>
						
					</tr>
					<% if not rsp.eof then %>
						<tr>
							<th class="l2_center" width="4%">ORDINE</th>
							<th class="l2_center" width="4%">LINGUA</th>
							<th class="L2">DOMINIO</th>
							<th class="l2_center" width="10%">NOME</th>
							<th class="l2_center" width="7%">HREF_LANG</th>
							<th class="l2_center" width="16%" colspan="2">OPERAZIONI</th>
						</tr>
						<%while not rsp.eof %>
							<tr>
								<td class="content_center"><%= rsp("dom_ordine") %></td>
								<td class="content_center">
									<img src="../grafica/flag_mini_<%= rsp("dom_lingua") %>.jpg">
								</td>
								<td class="content"><a href="<%= rsp("dom_url") %>"><%= rsp("dom_url") %></a></td>
								<td class="content_center"><%= rsp("dom_name") %></td>
								<td class="content_center"><%= rsp("dom_href_lang") %></td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica dei dati del dominio aggiuntivo." <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedScrollWindow('SitoDominioMod.asp?ID=<%= rsp("dom_id") %>&WEB_ID=<%= request("ID") %>', 'Dominio_<%= rsp("dom_id") %>', 640, 430, true)">
										MODIFICA
									</a>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione del dominio aggiuntivo"
									   onclick="OpenDeleteWindow('DOMINIO','<%= rsp("dom_id") %>');">
										CANCELLA
									</a>
								</td>
							</tr>
							<%rsp.MoveNext
						wend
					end if%>
				</table>
			<% rsp.close %>
			</td>
		</tr>
		
		<tr><th colspan="4">PAGINE DI SISTEMA</th></tr>
		<tr>
			<td class="label_no_width" colspan="2" rowspan="3">home page</td>
			<td class="label_no_width">per sito attivo:</td>
			<td class="content">
				<% CALL DropDownPages(conn, "form1", "", Session("AZ_ID"), "tfn_id_home_page", rs("id_home_page"), false, false) %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width">per sito in aggiornamento:</td>
			<td class="content">
				<% CALL DropDownPages(conn, "form1", "", Session("AZ_ID"), "nfn_sito_in_aggiornamento_pagina", rs("sito_in_aggiornamento_pagina"), false, false) %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width">per sito in costruzione:</td>
			<td class="content">
				<% CALL DropDownPages(conn, "form1", "", Session("AZ_ID"), "nfn_sito_in_costruzione_pagina", rs("sito_in_costruzione_pagina"), false, false) %>
			</td>
		</tr>
        <% if IsAreaRiservataActive(conn) then %>
		    <tr>
			    <td class="label_no_width" colspan="2" rowspan="4">area riservata:</td>
			    <td class="label_no_width">home area riservata</td>
				<td class="content">
					<% CALL DropDownPages(conn, "form1", "", Session("AZ_ID"), "tfn_id_home_page_riservata", rs("id_home_page_riservata"), false, false) %>
				</td>
            </tr>
            <tr>
			    <td class="label_no_width">pagina di login</td>
				<td class="content">
					<% CALL DropDownPages(conn, "form1", "", Session("AZ_ID"), "tfn_id_login_page_riservata", rs("id_login_page_riservata"), false, false) %>
				</td>
            </tr>
			<td class="label_no_width">pagina di logout</td>
				<td class="content">
					<% CALL DropDownPages(conn, "form1", "", Session("AZ_ID"), "tfn_id_logout_page_riservata", rs("id_logout_page_riservata"), false, false) %>
				</td>
            </tr>
            <tr>
			    <td class="label_no_width">pagina di registrazione</td>
				<td class="content">
					<% CALL DropDownPages(conn, "form1", "", Session("AZ_ID"), "tfn_id_registrazione_page_riservata", rs("id_registrazione_page_riservata"), false, false) %>
				</td>
            </tr>
        <% end if %>
		<tr>
			<td colspan="2" class="label_no_width">pagina di errore:</td>
			<td class="content" colspan="2">
				<% CALL DropDownPages(conn, "form1", "", Session("AZ_ID"), "nfn_errore_pagina", rs("errore_pagina"), false, false) %>
			</td>
		</tr>
	</table>
	
	<input type="hidden" name="tft_favicon" value="favicon.ico">
	<!--
	gestione commentata il 24/01/2010 perchè c'è un bug di explorer che ha bisogno della favicon.ico sempre in root e sempre con quel nome.
	il file dell'icona va quindi messo in root.
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="3">ALTRE INFORMAZIONI</th></tr>
		<tr>
			<td colspan="2" class="label_no_width" style="width:17%;">icona del sito:</td>
			<td class="content">
				<% 
				'CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_favicon", rs("favicon") , "width:320px;", false) 
				%>
			</td>
		</tr>
	</table>
	-->
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="4">GESTIONE LINGUE</th></tr>
		<tr>
			<td colspan="2" class="label_no_width" style="width:17%;">lingua iniziale:</td>
			<td class="content" colspan="2">
				<% if uBound(application("LINGUE"))=0 then %>
					<input type="hidden" name="tft_lingua_iniziale" value="<%= application("LINGUE")(0) %>">
				<% else %>
					<select name="tft_lingua_iniziale" id="tft_lingua_iniziale" <%= disable(uBound(application("LINGUE"))=0) %>>
						<option value="" <%= IIF(rs("lingua_iniziale")="", "selected", "") %>>
							Usa lingua del browser dell'utente
						</option>
					<% 	for i = 0 to uBound(application("LINGUE")) %>
						<option value="<%= application("LINGUE")(i) %>" <%= IIF(rs("lingua_iniziale")=application("LINGUE")(i), "selected", "") %>>
							<%= GetNomeLingua(application("LINGUE")(i) ) %>
						</option>
					<% 	next %>
					</select>
				<% end if %>
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			function attiva_lingua(lingua, obj){
				var obj_titolo = eval('form1.tft_titolo_' + lingua)
				var obj_keywords = eval('form1.tft_meta_keywords_' + lingua)
				var obj_description = eval('form1.tft_meta_description_' + lingua)
				DisableControl(obj_titolo, !(obj.checked))
				DisableControl(obj_keywords, !(obj.checked))
				DisableControl(obj_description, !(obj.checked))
			}
		</script>
		<% for each lingua in Application("LINGUE")%>
			<tr>
				<td class="label_no_width" style="width:4%;" rowspan="2">
					<img src="../grafica/flag_<%= lingua %>.jpg"></td>
				<td class="label_no_width">attiva lingua:</td>
				<td class="content" colspan="2">
					<% if lingua <> LINGUA_ITALIANO then 
					response.write lingua%>
						<input <%= chk(rs("lingua_"& lingua)) %> <% if chk(rs("lingua_"& lingua)) = "" then %> disabled <% end if %>
							class="checkbox" type="checkbox" name="chk_lingua_<%= lingua %>" onclick="attiva_lingua('<%= lingua %>', this)">
					<% else %>
						<input class="checkbox" type="checkbox" name="lingua_it" value="1" checked disabled>
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width">titolo pagine:</td>
				<td class="content" <% if lingua = LINGUA_ITALIANO then %> colspan="2" <% end if %>>
					<input type="text" name="tft_titolo_<%= lingua %>" size="90" maxlength="255" class="text" value="<%= rs("titolo_"& lingua) %>">
				</td>
				<% if lingua <> LINGUA_ITALIANO then %>
					<td class="content_right">
						<a class="button_L2" style="width:80px; text-align:center;" href="javascript:void(0)" title="apre in una nuova finestra l'attivazione della lingua corrispondente, con la possibilit&agrave; di copiare le pagine da italiano o inglese" <%= ACTIVE_STATUS %>
							onclick="OpenAutoPositionedWindow('AttivaLingua.asp?WEB_ID=<%= request("ID") %>&LINGUA_D=<%= lingua %>', 'Attiva_lingua', 600, 400)">
							ATTIVA LINGUA
						</a>
					</td>
				<% end if %>
			</tr>
		<%next %>
		<tr><th colspan="4">GESTIONE META TAG PER MOTORI DI RICERCA</th></tr>
		<tr>
			<td class="label_no_width" colspan="2">autore:</td>
			<td class="content" colspan="2"><input type="text" name="tft_meta_author" size="50" maxlength="255" class="text" value="<%= rs("meta_Author") %>"></td>
		</tr>
		<% for each lingua in Application("LINGUE")%>
			<tr>
				<td class="label_no_width" rowspan="2"><img src="../grafica/flag_<%= lingua %>.jpg"></td>
				<td class="label_no_width">keywords:</td>
				<td class="content" colspan="3"><textarea class="codice" rows="3" name="tft_meta_keywords_<%= lingua %>"><%= rs("meta_keywords_"& lingua) %></textarea></td>
			</tr>
			<tr>
				<td class="label_no_width">description:</td>
				<td class="content" colspan="2"><textarea class="codice" rows="3" name="tft_meta_description_<%= lingua %>"><%= rs("meta_description_"& lingua) %></textarea></td>
			</tr>
			<% if lingua<> LINGUA_ITALIANO then %>
				<script language="JavaScript" type="text/javascript">
					attiva_lingua('<%= lingua %>', form1.chk_lingua_<%=lingua %>);
				</script>
			<%end if
		next %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr>
			<th style="width:11%;"><a href="http://www.google.it" target="_blank"><img src="../grafica/Google/Logo_25wht.gif" width="75" height="32" alt="Google" border="1"></a></th>
			<th style="vertical-align:middle;">INTEGRATION</th>
			<th style="vertical-align:middle;">&nbsp;</th>
		</tr>
		<tr>
			<td class="label_no_width" style="width:20%;"><a href="http://www.google.it/analytics" target="_blank" title="apri www.google.it/analytics in una nuova finestra">Google Analytics</a></td>
			<td class="label_no_width" style="width:20%;">Account ID di monitoraggio:</td>
			<td class="content">
				<input type="text" class="text" name="tft_google_analytics_code" value="<%= rs("google_analytics_code") %>" maxlength="250" size="70">
			</td>
		</tr>
		<tr>
			<td class="label_no_width" style="width:20%;"><a href="http://www.google.com/webmasters/tools/" target="_blank" title="apri http://www.google.com/webmasters/tools/ in una nuova finestra">Google Webmaster Tools</a></td>
			<td class="label_no_width" style="width:20%;">Codice di verifica:</td>
			<td class="content">
				<input type="text" class="text" name="tft_google_webmaster_tools_verify_code" value="<%= rs("google_webmaster_tools_verify_code") %>" maxlength="250" size="70">
			</td>
		</tr>
		

		<tr><th colspan="4">METATAG AGGIUNTIVI</th></tr>
		<tr>
			<td colspan="4">
				<% sql = " SELECT * FROM tb_webs_metatag WHERE meta_web_id=" & cIntero(request("ID")) & _
						 " ORDER BY meta_name "
				rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label_no_width" style="width:30%">
							<% if rsp.eof then %>
								Nessun url inserito.
							<% else %>
								Trovati n&ordm; <%= rsp.recordcount %> record
							<% end if %>
						</td>
						<td colspan="3" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra l'inserimento di un metatag aggiuntivo" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedWindow('SitoMetaNew.asp?WEB_ID=<%= request("ID") %>', 'Url_nuovo', 640, 430)">
								NUOVO METATAG
							</a>
						</td>
					</tr>
					<% if not rsp.eof then %>
						<tr>
							<th class="L2">NAME</th>
							<th class="L2">CONTENT</th>
							<th class="l2_center" width="16%" colspan="2">OPERAZIONI</th>
						</tr>
						<%while not rsp.eof %>
							<tr>
								<td class="content"><%= rsp("meta_name") %></td>
								<td class="content"><%= rsp("meta_content") %></td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica dei dati del metatag." <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedScrollWindow('SitoMetaMod.asp?ID=<%= rsp("meta_id") %>&WEB_ID=<%= request("ID") %>', 'Url_<%= rsp("meta_id") %>', 640, 430, true)">
										MODIFICA
									</a>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione del metatag"
									   onclick="OpenDeleteWindow('META','<%= rsp("meta_id") %>');">
										CANCELLA
									</a>
								</td>
							</tr>
							<%rsp.MoveNext
						wend
					end if%>
				</table>
			<% rsp.close %>
			</td>
		</tr>
		
		<tr><th colspan="4">FILES RSS</th></tr>
		<tr>
			<td colspan="4">
				<% sql = " SELECT * FROM tb_rss WHERE rss_web_id=" & cIntero(request("ID")) & _
						 " ORDER BY rss_titolo "
				rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td colspan="4" class="label_no_width" style="width:74%">
							<% if rsp.eof then %>
								Nessun rss inserito.
							<% else %>
								Trovati n&ordm; <%= rsp.recordcount %> record
							<% end if %>
						</td>
						<td colspan="2" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra l'inserimento di un rss per il sito" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedWindow('SitoRssNew.asp?WEB_ID=<%= request("ID") %>', 'Rss_nuovo', 640, 430)">
								NUOVO RSS
							</a>
						</td>
					</tr>
					<% if not rsp.eof then %>
						<tr>
							<th class="L2">TITOLO</th>
							<th class="L2" width="20%">FILE</th>
							<th class="l2_center" width="7%">ABILITATO</th>
							<th class="l2_center" width="7%">METATAG</th>
							<th class="l2_center" width="16%" colspan="2">OPERAZIONI</th>
						</tr>
						<%while not rsp.eof %>
							<tr>
								<td class="content"><%= rsp("rss_titolo") %></td>
								<td class="content">
									<% if rsp("rss_abilitato") then %>
										<a target="_blank" href="<%= rs("URL_base") %>/<%= rsp("rss_file") %>"><%= rsp("rss_file") %></a>
									<% else %>
										<%= rsp("rss_file") %>
									<% end if %>
								</td>
								<td class="content_center"><input <%= chk(rsp("rss_abilitato")) %> type="checkbox" name="attivo" value="1" disabled class="checkbox"></td>
								<td class="content_center"><input <%= chk(rsp("rss_metatag")) %> type="checkbox" name="attivo" value="1" disabled class="checkbox"></td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica dei dati dell'RSS." <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedScrollWindow('SitoRssMod.asp?ID=<%= rsp("rss_id") %>&WEB_ID=<%= request("ID") %>', 'Rss_<%= rsp("rss_id") %>', 640, 430, true)">
										MODIFICA
									</a>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione dell'RSS"
									   onclick="OpenDeleteWindow('RSS','<%= rsp("rss_id") %>');">
										CANCELLA
									</a>
								</td>
							</tr>
							<%rsp.MoveNext
						wend
					end if%>
				</table>
			<% rsp.close %>
			</td>
		</tr>
		
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<tr><th colspan="4">CODICE DI APERTURA DELLA PAGINA</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea class="codice" rows="5" name="tft_pagehead_script"><%= rs("pagehead_script") %></textarea>
			</td>
		</tr>
		<tr><th colspan="4">CODICE DI CHIUSURA DELLA PAGINA</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea class="codice" rows="5" name="tft_pagefooter_script"><%= rs("pagefooter_script") %></textarea>
			</td>
		</tr>
		<%
		CALL Form_DatiModifica_EX(conn, rs, "webs_", "Dati di modifica", "") 
		sql = "SELECT COUNT(*) FROM tb_contents_index WHERE idx_webs_id=" & rs("id_webs")
		dim value
		value = cIntero(getValueList(conn, NULL, sql))
		if value > 1500 then %>
			<tr>
				<td class="footer" colspan="4">
					<input type="checkbox" class="checkbox" name="aggiorna_indice">
					Esegui anche aggiornamento completo dell'indice (presenti n&ordm; <%=value%> nodi)
				</td>
			</tr>
		<% else %>
			<input type="hidden" name="aggiorna_indice" value="1">
		<% end if %>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="mod" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
set rs = nothing
set rsp = nothing
conn.Close
set conn = nothing
connData.Close
set connData = nothing
%>