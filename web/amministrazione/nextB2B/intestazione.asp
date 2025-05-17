
<%
CALL CheckAutentication(Session("B2B_ADMIN") <> "" )
%>

<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->

<%

'configurazione autocompletamento descrittori
AutocompletamentoValori = true
AutoCompletionListSource = "CaratteristicheAutocompletamento.asp"

%>
<!DOCTYPE html>
<html>
	<head>
		<title><%= Session("NOME_APPLICAZIONE") %></title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<script src="http://script.aculo.us/prototype.js" type="text/javascript"></script>
		<script src="http://script.aculo.us/scriptaculous.js" type="text/javascript"></script>
		<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
		<SCRIPT LANGUAGE="javascript"  src="Tools_B2b.js" type="text/javascript"></SCRIPT>
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
				<td><a href="Articoli.asp" class="menu" title="gestione articoli" <%= ACTIVE_STATUS %>>ARTICOLI</a></td>
				<td><a href="Listini.asp" class="menu" title="gestione dei listini" <%= ACTIVE_STATUS %>>LISTINI</a></td>
				<td><a href="ListeCodici.asp" class="menu" title="gestione delle liste codici alternativi" <%= ACTIVE_STATUS %>>LISTE CODICI</a></td>
				<td><a href="Magazzini.asp" class="menu" title="gestione delle quantit&agrave; di magazzino" <%= ACTIVE_STATUS %>>MAGAZZINO</a></td>
				<td><a href="Ordini.asp" class="menu" title="gestione degli ordini" <%= ACTIVE_STATUS %>>ORDINI</a></td>
				<td><a href="Clienti.asp" class="menu" title="gestione dei clienti" <%= ACTIVE_STATUS %>>ANAGRAFICHE</a></td>
				<td><a href="Agenti.asp" class="menu" title="gestione degli agenti" <%= ACTIVE_STATUS %>>AGENTI</a></td>
				<% if cBoolean(cString(Session("ATTIVA_FATTURE")), false) then %>
					<td><a href="Fatture.asp" class="menu" title="gestione fatture" <%= ACTIVE_STATUS %>>FATTURE</a></td>
				<% end if %>
				<td><a href="Tabelle.asp" class="menu" title="gestione dei parametri generali: categorie, varianti, valute, classi di sconto..." <%= ACTIVE_STATUS %>>TABELLE</a></td>
				<td class="logout"><% CALL WriteLogoutLink(Session("NOME_APPLICAZIONE")) %></td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione("barra_nextb2b.jpg") %>
