<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools.asp" -->
<!--#INCLUDE FILE="Tools4Admin.asp" -->
<%
CALL ResetSession()
response.redirect "../default.asp"
%>
<html>
<head>
	<title>Logout</title>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>

<body>



</body>
</html>
