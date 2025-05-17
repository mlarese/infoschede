<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<%
categorie.Salva()

Sub Gestione_Relazioni_record(ByRef conn, rs, ID)
	CALL categorie.Gestione_Relazioni_record(rs, ID)
End Sub

set categorie = nothing
%>