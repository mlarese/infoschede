<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<% 

dim conn, rss, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rss = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM (tb_webs INNER JOIN tb_css_groups ON tb_webs.id_webs = tb_css_groups.grp_id_webs) " + _
	  " INNER JOIN tb_css_styles ON tb_css_groups.grp_id = tb_css_styles.style_grp_id " + _
	  " WHERE style_id=" & cIntero(request("ID"))
rss.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>
<!DOCTYPE html>
<html>
	<head>
		<title>Anteprima stili</title>
		<link rel="stylesheet" type="text/css" href="../../stili.css">
		<link rel="stylesheet" type="text/css" href="../library/site/nextweb5/standard.css">
		<style>
			a{
				cursor:pointer;
			}
			body
			{
				background-color: #FFFFFF;
			}
			<%= rss("style_class") %> {
				margin: 0px;
			}
			<%= vbCrLF & replace(Session("TMP_STILI_TESTO_" & session("AZ_ID")), "A:", "A.", 1, -1, vbTextCompare) %>
		</style>
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	</head>
<body>
	<div class="<%= rss("grp_name_class") %>" id="testo_1" style="position:absolute; left:2px; top:2px; width:100%;">
		<% if instr(1, rss("style_class"), "A", vbTextCompare)>0 then %>
			<p>
				<<%= rss("style_class") %> id="esempio" <% if cString(rss("style_pseudoclass"))<>"" then%> class="<%= replace(rss("style_pseudoclass"), ":", "") %>" <% end if %> title="link di esempio su testo di un paragrafo">
					Testo di esempio
				</<%= rss("style_class") %>>
			</p>
		<% else %>
			<<%= rss("style_class") %> id="esempio">Testo di esempio</<%= rss("style_class") %>>
		<% end if %>
	</div>
</body>
</html>
<% 

'rimuove variabile di sessione temporanea degli stili solo se e' l'ultima esecuzione del iframe
if request("last")<>"" then
	Session.contents.remove("TMP_STILI_TESTO_" & session("AZ_ID"))
end if

rss.close
conn.close 
set rss = nothing
set conn = nothing
%>
