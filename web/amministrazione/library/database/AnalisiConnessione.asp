<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<% 
'*****************************************************************************************************************
'verifica dei permessi
CALL VerificaPermessiUtente(true)
'*****************************************************************************************************************
%>
<html>
<head>
	<title>Amministrazione aggiornamenti database</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();">
<% 
'imposta elenco di schemi da visualizzare
dim Conn, prop
set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application(request("ConnString")), "", ""
%>
<table width="100%" cellspacing="1" cellpadding="0" style="margin-bottom:10px;">
    <caption style="border:0px;">
        <a style="float:right;" href="javascript:close();" class="menu" name="top">CHIUDI</a>
	</caption>
</table>
<table cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td style="padding-left:2px;">
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<caption>dettagli connessione "<%= request("ConnString") %>"</caption>
				<tr>
					<th>Nome propriet&agrave;</th>
					<th>Valore</th>
				</tr>
				<% for each prop in conn.Properties %>
				<tr>
					<td class="content"><%= prop.name %></td>
					<td class="content"><%= prop.value %></td>
				</tr>
				<% next %>
			</table>
		</td>
	</tr>
</table>
<br>
<br>
<br>
</body>
</html>

<% 
conn.close
set conn = nothing
%>
