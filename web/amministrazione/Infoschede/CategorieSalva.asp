<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede.asp" -->
<%
dim conn, sql

categorie.Salva()
set categorie = nothing

Sub Gestione_Relazioni_record(ByRef conn, rs, ID)
	CALL categorie.Gestione_Relazioni_record(rs, ID)
End Sub
%>