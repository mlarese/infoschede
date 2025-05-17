<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="Tools.asp" -->
<!--#INCLUDE FILE="Tools4Admin.asp" -->
<!--#INCLUDE FILE="ClassCryptography.asp" -->
<%
dim stringToEncrypt
stringToEncrypt = cString(request("string"))
if stringToEncrypt <> "" then
	response.write EncryptPassword(stringToEncrypt)
end if
%>