<!--#include file="../../library/Tools.asp"-->
<!--#include file="../../library/Tools4Admin.asp"-->
<!--#INCLUDE FILE="../../library/class_testata.asp" -->
<!--#include file="../../library/ClassIndirizzarioLock.asp"-->
<!--#INCLUDE FILE="../../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../../library/Categorie/ClassCategorie.asp" -->
<%
dim import_no_login
if not cBoolean(import_no_login, false) then
	'mette filtro: solo utente NEXT-AIM puo' entrare qui dentro
	if uCase(Session("MEMO2_ADMIN")) <> "NEXTAIM" then
		response.redirect "../default.asp"
	end if
end if
%>
<html>
<head>
	<title>Import dati NEXT-memo 2</title>
	<link rel="stylesheet" type="text/css" href="../../library/stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../../library/utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin=0 topmargin=0>
<!-- barra alta -->
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption style="border:0px;">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
	  			<td style="text-align:right; padding-right:10px;">
					<a href="../default.asp" class="menu" title="esci dall'import dati e torna al NEXT-memo 2" <%= ACTIVE_STATUS %>> torna al NEXT-memo 2</a>
				</td>
	  		</tr>
		</table>
	</caption>
  	<tr><td style="font-size:1px;">&nbsp;</td></tr>
  	<tr>
		<td style="background-image: url(../../grafica/barra_nextcom.jpg);" class="barra_menu">
			<% if cString(Application("DISABLE_NEXTAIM_LINKS"))="" then %>
				<a href="http://www.next-aim.com" target="_blank" title="supporto clienti su www.next-aim.com" <%= ACTIVE_STATUS %>>
			<% end if %>
				<img src="../../grafica/transp.gif" width="64" height="27" border="0" title="supporto clienti su www.next-aim.com" alt="supporto clienti su www.next-aim.com">
			<% if cString(Application("DISABLE_NEXTAIM_LINKS"))="" then %>
				</a>
			<% end if %>
			<br>
			<%=DataEstesa(Date(), LINGUA_ITALIANO)%>
		</td>
  	</tr>
	<% 	dim header, SSezioniText, SSezioniLink, Titolo_sezione, Sezione, HREF, Action
	set header = New testata 
	header.sezione = Titolo_sezione
	header.puls_new = action
	header.link_new = HREF
	if SSezioniText<>"" then
		header.sottosezioni = split(";" & SSezioniText, ";")
		header.links = split(";" & SSezioniLink, ";")
	else
		header.iniz_sottosez(0)
	end if
	header.scrivi_con_sottosez() 

%>
</table>