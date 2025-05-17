<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<% 
CALL CheckAutentication(Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE"))

dim Titolo_sezione, action, HREF
Titolo_sezione = "Gestione categorie - elenco"
HREF = "ContattiCategorieNew.asp"
Action = "NUOVA CATEGORIA"
SSezioniText = "CARATTERISTICHE;RAGGRUPPAMENTI"
SSezioniLink = "ContattiDescrittori.asp;ContattiDescrittoriGruppi.asp"
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%

CatContatti.Elenco()
set CatContatti = nothing
%>