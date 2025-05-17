<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione categorie - modifica sottocategorie"
dicitura.puls_new = "INDIETRO;SCHEDA"
dicitura.link_new = "Categorie.asp;CategorieMod.asp?ID=" & request("ID") & "&FROM=" & request("FROM")
dicitura.scrivi_con_sottosez()

categorie.ElencoSottoCategorie()
set categorie = nothing
%>