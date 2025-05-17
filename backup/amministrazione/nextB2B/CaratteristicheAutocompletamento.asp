<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->

<% 
	CALL AutocompleteList("grel_art_ctech", "rel_ctech_id", "rel_ctech_")
 %>