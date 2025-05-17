<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="../library/tools.asp" -->
<!--#INCLUDE FILE="../library/tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="Tools_b2b.asp" -->

<%
CALL oArticoliFoto.FormGestione()


Sub Gestione_Relazioni_record(ByRef conn, rs, ID)
	CALL oArticoliFoto.Gestione_Relazioni_foto(conn, rs, ID)
End Sub
%>