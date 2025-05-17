<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<% 	
dim Titolo_sezione, action, HREF
Titolo_sezione = "Gestione categorie - elenco"
HREF = "ContattiCategorieNew.asp"
Action = "NUOVA CATEGORIA"
SSezioniText = "CARATTERISTICHE;"
SSezioniLink = "ContattiDescrittori.asp;"
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/ClassJsTree.asp" -->
<%
CatContatti.Albero()
set CatContatti = nothing
%>