<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	

dim rs, sql
set rs = server.CreateObject("ADODB.recordset")

sql =" SELECT * FROM gtb_tipologie t WHERE tip_id="& cIntero(request("ID"))
rs.open sql, categorie.conn, adOpenForwardOnly, adLockReadOnly, adCmdText

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione categorie - modifica sottocategorie"
dicitura.puls_new = "INDIETRO;SCHEDA"
dicitura.link_new = "Categorie.asp;CategorieMod.asp?FROM=" & request("FROM") & "&ID=" & rs("tip_id")
if rs("tip_foglia") then
	dicitura.puls_new = dicitura.puls_new + ";RAGGRUPPAMENTI"
	dicitura.link_new = dicitura.link_new + ";CategorieRaggruppamenti.asp?ID=" & rs("tip_id")
end if
dicitura.scrivi_con_sottosez() 

categorie.ElencoSottoCategorie()
set categorie = nothing

set rs = nothing
%>