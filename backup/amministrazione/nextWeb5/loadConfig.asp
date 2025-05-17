<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<% response.ContentType = "text/xml" %>
<% response.buffer = true %>
<%
dim conn, rs, sql
dim BgImage, BgColor, IsTemplate, WebId, PageId
dim stnGrigliaAttiva, stnGuideVisibili, stnGuideColore, stnGuidePosVisibili, stnHelpAttivo

set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

PageId = cInteger(request("PAGINA"))

'apre recordset della pagina ed eventuale template
sql = " SELECT *, (tb_templates.id_page) AS TEMPLATE_ID, " + _
	  " (tb_pages.SfondoColore) AS PAGE_BgColor, (tb_pages.SfondoImmagine) AS PAGE_BgImage, " + _
	  " (tb_templates.SfondoColore) AS TEMPLATE_BgColor, (tb_templates.SfondoImmagine) AS TEMPLATE_BgImage, (tb_webs.id_webs) AS IdWeb " + _
	  " FROM ( tb_pages INNER JOIN tb_webs ON tb_pages.id_webs = tb_webs.id_webs ) " + _
      " LEFT JOIN tb_pages tb_templates ON tb_pages.id_template=tb_templates.id_page " + _
	  " WHERE tb_pages.id_page=" & PageId
rs.open sql, conn, adOpenForwardOnly, adLockReadonly, adCmdText

'imposta id del sito
WebId = rs("IdWeb")

'imposta parametro se lo sfondo e' modificabile
if cInteger(rs("TEMPLATE_ID"))>0 then
	'la pagina e' associata ad un template: sfondo non modificabile
	IsTemplate = 0
	
	'imposta sfondi del template
	BgImage = IIF(cString(rs("TEMPLATE_BgImage"))<>"", rs("TEMPLATE_BgImage"), "nosfondo")
	BgColor = IIF(cString(rs("TEMPLATE_BgColor"))<>"", rs("TEMPLATE_BgColor"), "#FFFFFF")
else
	'la pagina non e' associata ad alcun template o e' un template: sfondo modificabile
	IsTemplate = 1
	
	'imposta sfondi della pagina
	BgImage = IIF(cString(rs("PAGE_BgImage"))<>"", rs("PAGE_BgImage"), "nosfondo")
	BgColor = IIF(cString(rs("PAGE_BgColor"))<>"", rs("PAGE_BgColor"), "#FFFFFF")
end if

'imposta settaggi dell'editor
stnGrigliaAttiva = IIF(rs("sito_accessibile"), 1, 0)
stnGuideVisibili = IIF(rs("editor_guide_visibili"), 1, 0)
stnGuideColore = IIF(cString(rs("editor_guide_colore"))="", "#000000", rs("editor_guide_colore"))
stnGuidePosVisibili = IIF(rs("editor_guide_posizioni_visibili"), 1, 0)
stnHelpAttivo = IIF(rs("editor_help_attivo"), 1, 0)

rs.close
conn.close
set rs = nothing
set conn = nothing

'xml restituito dalla pagina
response.clear
%><?xml version="1.0" encoding="UTF-8"?>
<xml>
	<addressfiles>http://<%=Application("IMAGE_SERVER") %>/<%= WebId %></addressfiles>
	<sfondo><%= BgImage %></sfondo>
	<sfondoColore><%= BgColor %></sfondoColore>
	<setemplate><%= IsTemplate %></setemplate>
    <se_griglia_attiva><%= stnGrigliaAttiva %></se_griglia_attiva>
    <se_guide_visibili><%= stnGuideVisibili %></se_guide_visibili>
    <se_label_pos_visibili><%= stnGuidePosVisibili %></se_label_pos_visibili>
    <col_guide_ind><%= stnGuideColore %></col_guide_ind>
    <help_contestuale><%= stnHelpAttivo %></help_contestuale>
</xml>