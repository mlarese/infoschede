<!--#include file="../../library/Tools.asp"-->
<!--#include file="../../library/Tools4Admin.asp"-->
<!--#INCLUDE FILE="../../library/class_testata.asp" -->
<!--#INCLUDE FILE="../Tools_NextWeb5.asp" -->
<%CALL CheckAutentication(Session("WEB_ADMIN")<>"")
%>
<html>
<head>
	<title>Import dati NEXT-web 5.0</title>
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
					<a href="../default.asp" class="menu" title="esci dall'import dati e torna al NEXT-web" <%= ACTIVE_STATUS %>> torna al NEXT-web</a>
				</td>
	  		</tr>
		</table>
	</caption>
  	<% CALL WriteChiusuraIntestazione("barra_nextweb.jpg") %>