<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 

dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "CARATTERISTICHE"
dicitura.links(1) = "Caratteristiche.asp"


dicitura.sezione = "Gestione categorie - elenco"
dicitura.puls_new = "NUOVA CATEGORIA"
dicitura.link_new = "CategorieNew.asp?FROM=" & FROM_ELENCO
dicitura.scrivi_con_sottosez()

categorie.Elenco()
set categorie = nothing
%>