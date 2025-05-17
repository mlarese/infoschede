<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim rs, sql, conn
set rs = server.createobject("ADODB.recordset")
sql =" SELECT *, (SELECT COUNT(*) FROM gtb_articoli WHERE art_tipologia_id=t.tip_id) AS N_ART, " & _
	  " (SELECT COUNT(*) FROM gtb_tipologie WHERE tip_padre_id=t.tip_id) AS N_FIGLI, " & _
	  " (SELECT COUNT(*) FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=t.tip_id) AS N_GRUPPI " & _
	  " FROM gtb_tipologie t WHERE tip_id="& cIntero(request("ID"))
rs.open sql, categorie.conn, adOpenForwardOnly, adLockReadOnly, adCmdText

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione categorie - modifica"
if rs("tip_foglia") then
	dicitura.puls_new = dicitura.puls_new + "RAGGRUPPAMENTI;"
	dicitura.link_new = dicitura.link_new + "CategorieRaggruppamenti.asp?ID=" & rs("tip_id") & ";"
end if
if cInteger(rs("N_ART")) = 0 AND cInteger(rs("N_GRUPPI"))=0 then
	dicitura.puls_new = dicitura.puls_new + "SOTTOCATEGORIE"
	dicitura.link_new = dicitura.link_new + "CategorieSottocategorie.asp?ID=" & rs("tip_id") & ";"
end if
rs.close
set rs = nothing

categorie.Modifica(dicitura)
set categorie = nothing
%>