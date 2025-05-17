<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione categorie - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA CATEGORIA"
dicitura.link_new = "Tabelle.asp;CategorieNew.asp?FROM=" & FROM_ELENCO
dicitura.scrivi_con_sottosez()

categorie.Elenco()
set categorie = nothing
%>