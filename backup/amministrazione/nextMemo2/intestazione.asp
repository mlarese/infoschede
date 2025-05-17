
<%
CALL CheckAutentication(Session("MEMO2_ADMIN")<>"" OR Session("MEMO2_DOWNLOAD")<>"")
%>

<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<!--#INCLUDE FILE="Tools_Memo2.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!DOCTYPE html>
<html>
	<head>
		<title><%= Session("NOME_APPLICAZIONE") %></title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
<body>
<!-- barra alta -->
<div id="Layer0" style="position:absolute; left:0px; top:0px; width:740px; height:400px; z-index:0">
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption class="menu">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
	  			<td width="10">&nbsp;</td>
				<% if Session("MEMO2_ADMIN")<>"" then%>
					<td><a href="Documenti.asp" class="menu" title="gestione documenti e circolari" <%= ACTIVE_STATUS %>>GESTIONE DOCUMENTI</a></td>
					<% if cBoolean(Session("CATEGORIE_NEXTMEMO2_ABILITATE"), false) then %>
						<td><a href="Categorie.asp" class="menu" title="gestione categorie di documenti" <%= ACTIVE_STATUS %>>CATEGORIE</a></td>
					<% end if %>
					<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) then %>
						<td><a href="Profili.asp" class="menu" title="gestione profili" <%= ACTIVE_STATUS %>>PROFILI</a></td>
					<% end if %>	
					<% if cBoolean(Session("CONDIVISIONE_INTERNA"), false) then %>
						<td><a href="Amministratori.asp" class="menu" title="gestione utenti dell'area amministrativa" <%= ACTIVE_STATUS %>>UTENTI AREA AMMINISTRATIVA</a></td>
					<% end if %>
					<% if cBoolean(Session("CONDIVISIONE_PUBBLICA"), false) then %>
						<td><a href="Utenti.asp" class="menu" title="gestione utenti dell'area riservata" <%= ACTIVE_STATUS %>>UTENTI AREA RISERVATA</a></td>
					<% end if %>
					<% if cBoolean(Session("AGENDA_MEMO2_ATTIVA"), false) then %>
						<td><a href="Impegni.asp" class="menu" title="gestione impegni/appuntamenti" <%= ACTIVE_STATUS %>>AGENDA</a></td>
					<% end if %>
				<% end if
				if cBoolean(Session("CONDIVISIONE_INTERNA"), false) AND Session("MEMO2_DOWNLOAD")<>"" then%>
					<td><a href="Download.asp" class="menu" title="lettura e download di documenti e circolari" <%= ACTIVE_STATUS %>>DOWNLOAD</a></td>
				<% end if %>
				<td class="logout"><% CALL WriteLogoutLink(Session("NOME_APPLICAZIONE")) %></td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione("barra_nextMemo.jpg") %>
