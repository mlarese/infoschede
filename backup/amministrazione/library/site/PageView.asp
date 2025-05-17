<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<%

'calcola url effettivo della pagina richiesta
dim URL
if cIntero(request("PAGINA"))>0 then
	URL = GetPageUrl(NULL, request("PAGINA"))
elseif cIntero(request("PS"))>0 then
	URL = GetPageSiteUrl(NULL, request("PS"), IIF(request("lingua")<>"", request("lingua"), LINGUA_ITALIANO))
else
	URL = ""
end if



if url<>"" then
	response.redirect URL
else
	response.redirect "http://" & Application("SERVER_NAME")
end if
%>