<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->

<%
'-------------------------------------------------------------------
function GetTipoNewsletter()
	dim tipo_newsletter
	tipo_newsletter = ComunicazioniNew_Wizard_Session_GetField(messageType, "tipo")
	tipo_newsletter = Replace(tipo_newsletter, EMAIL_TYPE_NEWSLETTER&"_", "")
	GetTipoNewsletter = cIntero(tipo_newsletter)
end function

function GetIdPaginaNewsletter(conn, tipo_newsletter)
	dim sql, id_pagina, lingua
	sql = "SELECT nl_pagina_id FROM tb_newsletters WHERE nl_id = " & tipo_newsletter
	id_pagina = cIntero(GetValueList(conn, NULL, sql))
	
	sql = "SELECT nl_lingua FROM tb_newsletters WHERE nl_id = " & tipo_newsletter
	lingua = Trim(GetValueList(conn, NULL, sql))
	if lingua="" then
		lingua = "it"
	end if
			
	sql = "SELECT id_pagDyn_"&lingua&" FROM tb_pagineSito WHERE id_pagineSito = "&id_pagina
	id_pagina = cIntero(GetValueList(conn, NULL, sql))
	
	GetIdPaginaNewsletter = id_pagina
end function

function GetUrlPaginaNewsletter(conn, id_pagina, tipo_newsletter)
	'GetUrlPaginaNewsletter = GetSiteUrl(conn, Application("AZ_ID"), 0)&"/default.aspx?PAGINA="&id_pagina&"&TIPO_NEWSLETTER="&tipo_newsletter&"&HTML_FOR_EMAIL=1&Generatedtime="
	GetUrlPaginaNewsletter = replace(GetSiteBaseUrl(conn, GetPaginaSitoIdByPaginaId(conn, id_pagina))&"/default.aspx?PAGINA="&id_pagina&"&TIPO_NEWSLETTER="&tipo_newsletter&"&HTML_FOR_EMAIL=1&Generatedtime="&Now(), "https", "http")
	
end function
'-------------------------------------------------------------------



dim conn, rs, sql, MessageType, MittenteValido, ConfigurazioneValida
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

MessageType = cIntero(request("type"))

dim tipo_newsletter, id_pagina, lingua, url

tipo_newsletter = GetTipoNewsletter()
id_pagina = GetIdPaginaNewsletter(conn, tipo_newsletter)
CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, EMAIL_TYPE_NEWSLETTER, id_pagina)

lingua = ComunicazioniNew_Wizard_Session_GetField(messageType, "newsletter_lingua")
if lingua="" then
	lingua = "it"
end if

url = GetUrlPaginaNewsletter(conn, id_pagina, tipo_newsletter)
CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "newsletter_url", url)



if request.ServerVariables("REQUEST_METHOD") = "POST" OR _
	ComunicazioniNew_Wizard_Session_GetField(MessageType, "newsletter_scelta_contenuti") = "" then
	'dopo aver scelto i contenuti dell newsletter, oppure nel caso di newsletter senza la scelta dei contenuti
	
	dim html, html_to_replace
	url = ComunicazioniNew_Wizard_Session_GetField(messageType, "newsletter_url")
	html = ExecuteHttpUrl(EncodeCssInlinedUrl(URL))
'response.write URL & "<br>"
'response.write EncodeCssInlinedUrl(URL)
'response.end
	'ripulisco il codice html e lo preparo per CKEditor	
	html = CleanHtmlForCKEditor(html)
	
	CALL ComunicazioniNew_Wizard_Session_AddField(MessageType, "tft_email_text", html)
'response.write ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_text")
'response.end	
	'Salva messaggio in file html temporaneo
	CALL WriteBozzaHtml("")

	if request.form("indietro") <> "" then
		response.redirect "ComunicazioniNew_Wizard_1.asp?type=" & MessageType
	else
		response.redirect "ComunicazioniNew_Wizard_2.asp?type=" & MessageType
	end if
end if



'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = ComunicazioniNew_Wizard_Titolo("Comunicazioni in uscita - ", 1, ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo"))
'Indirizzo pagina per link su sezione 
	HREF = "Comunicazioni.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->
<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<div id="content">
<form action="<%= GetPageName() %>?type=<%= MessageType %>" method="post" id="form1" name="form1">
	<input type="hidden" name="email_tipo" value="<%= ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") %>">
	<input type="hidden" name="tfd_email_Data" value="NOW">
	<input type="hidden" name="tft_email_dipgenera" value="<%= Session("ID_ADMIN") %>">
	<input type="hidden" name="tfn_email_isBozza" value="<%= IIF(ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = EMAIL_TYPE_BOZZA OR ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = FAX_TYPE_BOZZA OR ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo") = SMS_TYPE_BOZZA, "1", "0") %>">
	<script language="JavaScript" type="text/javascript">
		function OpenLoadShock(operazione){
			var width, height;
			width = 650;
			height = 300;
			if (operazione == 'anteprima'){
				var url = parent.email_newsletter_view.document.location;
				 window.open(url);
			}
			else if (operazione == 'modifica'){
				OpenAutoPositionedScrollWindow('ComunicazioniNew_Wizard_2_newsletter_contents.asp?type=<%= MessageType %>&TIPO_NEWSLETTER=<%=tipo_newsletter%>&PAGINA=<%= id_pagina %>', 
												operazione, width, height, true);
			}
			else{
				OpenAutoPositionedScrollWindow('ComunicazioniNew_Wizard_2_newsletter_lingua.asp?type=<%= MessageType %>&PAGINA=' + document.form1.email_newsletter.value + '&operazione=' + operazione, 
												operazione, width, height, true);
			}
		}

	</script>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<% CALL Comunicazioni_Icona(MessageType) %>&nbsp;
			<%= ComunicazioniNew_Wizard_Titolo("", 1, ComunicazioniNew_Wizard_Session_GetField(MessageType, "tipo")) %> -
			Scelta contenuti newsletter
		</caption>
		<tr><th colspan="2">CORPO DEL MESSAGGIO</th></tr> 
		<input type="hidden" name="email_newsletter" id="email_newsletter" value="<%= id_pagina %>">
		<tr>
			<td class="label" colspan="2">
				<span style="float:left; line-height:25px;">Scegli i contenuti che devono apparire nella newsletter: &nbsp;</span>
				<a class="button_input_bottom" style="float:right; width:250px; line-height:25px;" href="javascript:void(0);" onclick="OpenLoadShock('modifica')">
					CAMBIA I CONTENUTI NELLA NEWSLETTER
				</a>
				&nbsp;
				<!--<a class="button_L2" href="javascript:void(0);" onclick="OpenLoadShock('anteprima')">
					VEDI ANTEPRIMA NEWSLETTER
				</a>-->
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<iframe src="" name="email_newsletter_view" width="100%" height="600" id="email_newsletter_view"></iframe>
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			function SetPreview(pagina){
				if (pagina > 0) {
					var now = new Date();
					//imposta visualizzazione della pagina nel frame e numero di pagina nell'input
					document.form1.email_newsletter.value = pagina;
					//alert("<%=GetSiteUrl(conn, Application("AZ_ID"), 0)%>/default.aspx?PAGINA=" + pagina + "&TIPO_NEWSLETTER=<%=tipo_newsletter%>&HTML_FOR_EMAIL=1&Generatedtime=" + now.getMilliseconds());
					parent.email_newsletter_view.document.location = "<%=GetSiteUrl(conn, Application("AZ_ID"), 0)%>/default.aspx?PAGINA=" + pagina + "&TIPO_NEWSLETTER=<%=tipo_newsletter%>&HTML_FOR_EMAIL=1&Generatedtime=" + now.getMilliseconds();
					//document.email_newsletter_view.document.location ="<%= GetLibraryPath() %>site/PageView.asp?PAGINA=" + pagina + "&TIPO_NEWSLETTER=<%=tipo_newsletter%>";
				} else
					parent.email_newsletter_view.document.location.reload(true);

			}
		</script>
		<script language="JavaScript" type="text/javascript">
			//imposta visualizzazione della pagina nel frame e numero di pagina nell'input
			SetPreview( <%= id_pagina %>);
		</script>
		<tr>
			<td class="footer" colspan="2">
				<input style="width:10%;" type="submit" class="button" name="indietro" value="&laquo; INDIETRO" title="torna alla selezione del formato.">
				<input style="width:10%;" type="submit" class="button" name="avanti" value="AVANTI &raquo;" title="vai alla scelta destinatari e composizione del corpo del messaggio.">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
conn.close 
set rs = nothing
set conn = nothing%>