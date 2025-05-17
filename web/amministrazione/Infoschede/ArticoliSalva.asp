<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede.asp" -->
<!--#INCLUDE FILE="../nextB2B/Tools4Save_B2B.asp" -->

<% dim conn, sql
if cString(request("CATEGORIA"))="ricambio" then 
	redirect = "../Infoschede/SchedeDettagliNew.asp?IDSCH="&request("IDSCH")
end if
%>

<!--#INCLUDE FILE="../nextB2B/ArticoliSalva_Tools.asp" -->