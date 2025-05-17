<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
'--------------------------------------------------------
sezione_testata = "selezione categoria" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

categorie.Seleziona()
set categorie = nothing
%>