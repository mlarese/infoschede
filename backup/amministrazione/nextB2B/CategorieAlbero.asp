<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassJsTree.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione categorie - albero"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA CATEGORIA"
dicitura.link_new = "Tabelle.asp;CategorieNew.asp?FROM=" & FROM_ALBERO
dicitura.scrivi_con_sottosez()

categorie.Albero()
set categorie = nothing
%>