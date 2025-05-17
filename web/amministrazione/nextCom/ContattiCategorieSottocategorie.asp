<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<% 
CALL CheckAutentication(Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE"))

dim Titolo_sezione, action, HREF
Titolo_sezione = "Gestione categorie - modifica sottocategorie"
HREF = "ContattiCategorie.asp;ContattiCategorieMod.asp?ID=" & request("ID") & "&FROM=" & request("FROM")
Action = "INDIETRO;SCHEDA"
SSezioniText = "CATEGORIE;CARATTERISTICHE;"
SSezioniLink = "ContattiCategorie.asp;ContattiDescrittori.asp;"
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%

CatContatti.ElencoSottoCategorie()
set CatContatti = nothing
%>