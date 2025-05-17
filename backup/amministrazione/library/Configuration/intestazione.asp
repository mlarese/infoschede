<%
'verifica dei permessi
CALL VerificaPermessiUtente(true)
%>

<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../class_testata.asp" -->
<!--#INCLUDE FILE="../database/Tools4Database.asp" -->

<html>
	<head>
		<title>Gestione applicativi</title>
		<link rel="stylesheet" type="text/css" href="../stili.css">
		<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
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
	  			<td width="10">&nbsp;</td>
				<td><a href="Applicazioni.asp" class="menu" title="gestione delle applicazioni e loro impostazioni" <%= ACTIVE_STATUS %>>APPLICAZIONI</a></td>
				<td class="logout"><a href="#" onclick="javascript:window.close()">CHIUDI</a></td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione("barra_nextPassport.jpg") %>
