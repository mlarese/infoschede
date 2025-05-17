<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 1073741824 %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../../nextPassport/ToolsApplicazioni.asp" -->
<!--#INCLUDE FILE="../ClassCryptography.asp"-->
<!--#INCLUDE FILE="../Tools_4_CreditCard.asp"-->

<% 
'*****************************************************************************************************************
'verifica dei permessi
CALL VerificaPermessiUtente(true)
'*****************************************************************************************************************
%>
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	<title>Amministrazione aggiornamenti database</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();">
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption style="border:0px;">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
				<td align="right" style="padding-right:10px;">
					<a href="javascript:close();" class="menu" name="top">CHIUDI</a>
				</td>
	  		</tr>
		</table>
	</caption>
</table>

<div id="content" style="top:20px;">
<% 
dim conn, sql, rs, rsT, DB
dim fso, path, folder, SubFolder, Field
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsT = Server.CreateObject("ADODB.RecordSet")
Conn.Open Application(request("ConnString")), "", ""
'*******************************************************************************************
set DB = new UpdateDabase
conn.beginTrans
DB.Init (conn)
'*******************************************************************************************
%>