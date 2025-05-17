<% 
'variabile impostata nel file chiamante
dim sezione_testata, testata_show_back
dim testata_elenco_pulsanti, testata_elenco_href, testata_index, body_attributes


if Session("LOGIN_4_LOG")="" then
	'utente non loggato
	%>
	<script language="JavaScript" type="text/javascript">
		//esegue reload della finestra padre
		try { opener.location.reload(true);}
		catch(e){/*istruzione messa solo per sintassi*/}

		//chiude la finestra corrente
		window.close();
	</script>
	<%response.end
end if%>
<!DOCTYPE html>
<html>
<head>
	<title><%= Session("NOME_APPLICAZIONE") %></title>
	<link rel="stylesheet" type="text/css" href="<%= GetLibraryPath() %>stili.css">
	<SCRIPT LANGUAGE="javascript" src="<%= GetLibraryPath() %>utils.js" type="text/javascript"></SCRIPT>
	<% if not IsHttpsActive() then %>
		<script src="http://script.aculo.us/prototype.js" type="text/javascript"></script>
		<script src="http://script.aculo.us/scriptaculous.js" type="text/javascript"></script>
		<!--<script src="<%= GetLibraryPath() %>lightbox/js/lightbox.js" type="text/javascript"></script>-->
	<% end if %>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" onLoad="window.focus()" <%=body_attributes%>>
	<div id="testata_ridotto">
		<table width="100%" cellspacing="0" cellpadding="0">
			<caption class="menu" style="text-align:right;">
				<a href="javascript:window.close();" class="menu" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Chiudi la finestra", "Close the window", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%>
				</a>
			</caption>
			<tr>
				<td><img src="<%= GetAmministrazionePath() %>grafica/<%= Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA") %>barra_ridotta_left.jpg" alt=""></td>
				<td width="90%" style="font-size: 1px; background-image: url(<%= GetAmministrazionePath() %>grafica/<%= Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA") %>barra_ridotta_center.jpg); background-repeat: repeat;">&nbsp;</td>
				<td><img src="<%= GetAmministrazionePath() %>grafica/<%= Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA") %>barra_ridotta_right.jpg" alt=""></td>
			<tr>
			<tr>
				<td colspan="3" style="padding-left:5px; padding-top:1px;">
					<table width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<% if Session("ERRORE")<>"" then %>
								<td class="errore"><%= Session("ERRORE") %></td>
								<% Session("ERRORE") = ""
							else %>
								<td style="font-weight:bold; font-size:12px;">
									<span style="font-size:10px;">
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "sezione:", "section:", "", "", "", "", "", "")%>
									</span>
									<%= lCase(sezione_testata) %>
								</td>
							<% end if %>
							<td align="right" style="padding-right:1px;">
								<% if testata_show_back then %>
									<a class="button" href="javascript:history.go(-1);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Indietro", "Back", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "INDIETRO", "BACK", "", "", "", "", "", "")%></a>
								<% end if 
								if testata_elenco_pulsanti<>"" AND testata_elenco_href<>"" then
									testata_elenco_pulsanti = split(testata_elenco_pulsanti, ",")
									testata_elenco_href = split(testata_elenco_href, ",")
									if ubound(testata_elenco_href) = ubound(testata_elenco_pulsanti) then
										for testata_index=lbound(testata_elenco_href) to ubound(testata_elenco_href)%>
											<a class="button" href="<%= testata_elenco_href(testata_index) %>" title="<%= testata_elenco_pulsanti(testata_index) %>" <%= ACTIVE_STATUS %>>
												<%= testata_elenco_pulsanti(testata_index) %></a>
										<%next
									end if
								end if %>
								<a class="button" href="javascript:window.close();" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Chiudi la finestra", "Close the window", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%></a>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>