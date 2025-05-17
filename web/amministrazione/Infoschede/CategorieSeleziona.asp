<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Const.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Categorie.asp" -->

<%
'--------------------------------------------------------
sezione_testata = "selezione categoria" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
dim conn, sql
if Session("TYPE")="R" then
	cat_ricambi.Seleziona()
elseif Session("TYPE")="M" then
	cat_modelli.Seleziona()
else
	cat_articoli.Seleziona()
end if
	
'set cat_modelli = nothing
%>