
<%
CALL CheckAutentication((Session("PASS_ADMIN") & Session("PASS_AMMINISTRATORI") & Session("PASS_UTENTI")) <>"")
%>

<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!DOCTYPE html>
<html>
	<head>
		<title><%= Session("NOME_APPLICAZIONE") %></title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
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
				<% if Session("PASS_ADMIN")<>"" OR Session("PASS_AMMINISTRATORI")<>"" then %>
					<td><a href="Applicazioni.asp" class="menu" title="gestione delle applicazioni e loro impostazioni" <%= ACTIVE_STATUS %>>APPLICAZIONI</a></td>
				<% end if %>
				<td><a href="Alert.asp" class="menu" title="gestione degli eventi delle applicazioni" <%= ACTIVE_STATUS %>>ALERT&nbsp;</a></td>
				<% if (Session("PASS_ADMIN")<>"") OR (Session("PASS_AMMINISTRATORI")<>"") then %>
					<td><a href="Amministratori.asp" class="menu" title="gestione profili di accesso degli utenti dell'area amministrativa" <%= ACTIVE_STATUS %>>UTENTI AREA AMMINISTRATIVA</a></td>
					<% if cInteger(Application("NextCom_DefaultWorkGroup"))=0 then %>
						<td><a href="Gruppi.asp" class="menu" title="gestione dei gruppi di lavoro e loro composizione degli utenti dell'area amministrativa" <%= ACTIVE_STATUS %>>GRUPPI DI LAVORO</a></td>
					<% end if
				end if
				
				if (Session("PASS_ADMIN")<>"") OR (Session("PASS_UTENTI")<>"") then %>
					<td><a href="Utenti.asp" class="menu" title="gestione profili di accesso degli utenti dell'area riservata" <%= ACTIVE_STATUS %>>UTENTI AREA RISERVATA</a></td>
				<% end if %>
				<% if Session("PASS_ADMIN")<>""then %>
					<td><a href="Strumenti.asp" class="menu" title="strumenti per la gestione del sito" <%= ACTIVE_STATUS %>>STRUMENTI</a></td>
				<% end if %>
				<td class="logout"><% CALL WriteLogoutLink(Session("NOME_APPLICAZIONE")) %></td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione("barra_nextPassport.jpg") %>
