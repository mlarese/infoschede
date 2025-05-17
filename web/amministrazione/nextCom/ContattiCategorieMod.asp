<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
CALL CheckAutentication(Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE"))

dim Titolo_sezione, action, HREF
Titolo_sezione = "Gestione categorie - elenco"
HREF = "ContattiCategorie.asp"
Action = "INDIETRO"
SSezioniText = "CARATTERISTICHE;"
SSezioniLink = "ContattiDescrittori.asp;"
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%
dim noObject
set noObject = nothing
CatContatti.Modifica(noObject)
set CatContatti = nothing
%>