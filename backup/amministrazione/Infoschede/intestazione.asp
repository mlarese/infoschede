<% Server.ScriptTimeout = 1073741824 %>
<% CALL CheckAutentication((Session("INFOSCHEDE_ADMIN"))<>"" OR (Session("INFOSCHEDE_CENTRO_ASSISTENZA"))<>"" OR (Session("INFOSCHEDE_OFFICINA"))<>"") 
%>

<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->

<html>
	<head>
		<title><%= Session("NOME_APPLICAZIONE") %></title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		
		<style type="text/css"> 
			.riscontrato {
			  background-color: #d0e4f5 !important;
			}
		
			.segnalato {
			  background-color: #fce0c7 !important;
			  color: #8e4302 !important;
			}
		</style>
		
		<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
		<!--<SCRIPT LANGUAGE="javascript"  src="../nextB2B/Tools_B2b.js" type="text/javascript"></SCRIPT>-->
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	</head>
<body>
<!-- barra alta -->
<div id="Layer0" style="position:absolute; left:0px; top:0px; width:740px; height:400px; z-index:0">
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption class="menu">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
				<% if Session("INFOSCHEDE_ADMIN") <> "" then %>
					<td><a href="Schede.asp?ASSEGNATA=false" class="menu" title="gestione delle schede da assegnare ad un centro assistenza" <%= ACTIVE_STATUS %>>RICHIESTE ASSISTENZA</a></td>	
				<% end if %>
				<td><a href="Schede.asp?ASSEGNATA=true" class="menu" title="gestione delle schede assegnate ad un centro assistenza" <%= ACTIVE_STATUS %>>SCHEDE ASSISTENZA</a></td>				
				<td><a href="Clienti.asp?PROFILO=anagrafiche_clienti" class="menu" title="gestione clienti privati, rivenditori e supervisori negozi" <%= ACTIVE_STATUS %>>ANAGRAFICHE CLIENTI</a></td>
				<% if Session("INFOSCHEDE_ADMIN") <> "" then %>
					<td><a href="Ritiri.asp" class="menu" title="sezione ritiri" <%= ACTIVE_STATUS %>>RITIRI</a>&nbsp;&nbsp;</td>
					<td><a href="Spedizioni.asp" class="menu" title="sezione spedizioni" <%= ACTIVE_STATUS %>>SPEDIZIONI</a>&nbsp;</td>
					<td><a href="Tabelle.asp" class="menu" title="gestione dei parametri generali" <%= ACTIVE_STATUS %>>TABELLE</a></td>
				<% end if %>
				<td class="logout"><% CALL WriteLogoutLink(Session("NOME_APPLICAZIONE")) %></td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione("") %>
