<% 
CALL CheckAutentication(Session("COM_USER") <> "" OR Session("COM_ADMIN") <> "" OR Session("COM_POWER") <> "")

%>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!DOCTYPE HTML>
<html>
<head>
	<title><%= Session("NOME_APPLICAZIONE") %></title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin=0 topmargin=0>
<!-- barra alta -->
<div id="Layer0" style="position:absolute; left:0px; top:0px; width:740px; height:400px; z-index:0">
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<% if request("MODE") <> "iframe" then %>
		<caption class="menu">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
			<tr>
				<td width="10">&nbsp;</td>
				<td><a href="Contatti.asp" class="menu" title="gestione anagrafiche dei contatti" <%= ACTIVE_STATUS %>>ANAGRAFICHE</a></td>
				<td><a href="Comunicazioni.asp" class="menu" title="gestione di tutte le comunicazioni in uscita" <%= ACTIVE_STATUS %>>COMUNICAZIONI</a></td>
				<% if Session("COM_ADMIN")<>"" OR Session("COM_POWER")<>"" then %>
					<td><a href="Rubriche.asp" class="menu" title="gestione rubriche di classificazione dei contatti" <%= ACTIVE_STATUS %>>RUBRICHE</a></td>
				<%end if %>
				<% if Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") AND (Session("COM_ADMIN")="" OR Session("COM_POWER")="") then %>
					<td><a href="Campagne.asp" class="menu" title="gestione campagne marketing" <%= ACTIVE_STATUS %>>CAMPAGNE MARKETING</a></td>
				<%end if %>
				<% If Application("NextCrm") then %>
					<td><a href="Pratiche.asp?all=1" class="menu" title="gestione completa delle pratiche" <%= ACTIVE_STATUS %>>PRATICHE</a></td>
					<td><a href="Attivita.asp?all=si" class="menu" title="gestione completa delle attivit&agrave;" <%= ACTIVE_STATUS %>>ATTIVIT&Agrave;</a></td>
					<td><a href="Documenti.asp?all=si" class="menu" title="gestione completa dei documenti" <%= ACTIVE_STATUS %>>DOCUMENTI</a></td>
				<% End If %>
				<% if Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE") then %>
					<td><a href="ContattiCategorie.asp" class="menu" title="gestione delle categorie di gallery" <%= ACTIVE_STATUS %>>CATEGORIE</a></td>
				<% end if %>
				<td class="logout"><% CALL WriteLogoutLink(Session("NOME_APPLICAZIONE")) %></td>
			</tr>
		</table>
		</caption>
		<% CALL WriteChiusuraIntestazione(IIF(Application("NextCrm"), "barra_nextDoc.jpg", "barra_nextCom.jpg")) %>
	<% end if %>	

<% 	dim header, SSezioniText, SSezioniLink, NumSottosez
	set header = New testata 
	header.sezione = Titolo_sezione
	header.puls_new = action
	header.link_new = HREF
	if SSezioniText<>"" then
		NumSottosez = split(";" & SSezioniText, ";")
		header.iniz_sottosez(UBound(NumSottosez))
		header.sottosezioni = NumSottosez
		header.links = split(";" & SSezioniLink, ";")
	else
		header.iniz_sottosez(0)
	end if
	header.scrivi_con_sottosez() 

%>