<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, conn, sql
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione categorie - nuova"

categorie.Nuova(dicitura)
set categorie = nothing
%>