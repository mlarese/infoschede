<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
Imposta_Proprieta_Sito("ID")
%>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<% 
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_stili_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - stili di testo"
dicitura.puls_new = "INDIETRO A SITI"
dicitura.link_new = "Siti.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rsg, rss, sql, cssO
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rsg = Server.CreateObject("ADODB.RecordSet")
set rss = Server.CreateObject("ADODB.RecordSet")

'genera stili una volta sola anche per tutti gli iframe che mostrano gli stili.
'La variabile di sessione eviene rimossa nell'ultima esecuzione dell'iframe quando c'e' il parametro LAST=1
set cssO = new CssManager
Session("TMP_STILI_TESTO_" & session("AZ_ID")) = cssO.GenerateCss(conn, session("AZ_ID"), false)

Session("STILI_SQL") = " SELECT * FROM tb_css_groups INNER JOIN tb_css_styles ON tb_css_groups.grp_id = tb_css_styles.style_grp_id" + _
				 	   " WHERE tb_css_groups.grp_id_webs = " & Session("AZ_ID") & _
					   " ORDER BY grp_name, style_id"

sql = "SELECT * FROM tb_css_groups WHERE grp_id_webs = " & Session("AZ_ID") & " ORDER BY grp_name"
rsg.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText  %>
<div id="content">
	<% while not rsg.eof %>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<caption class="border" title="class name=<%= rsg("grp_name_class") %><%= vbCrLf %>id=<%= rsg("grp_id") %><%= vbCrLf %>checksum=<%= rsg("grp_checksum") %>">
				Stili per: <%= rsg("grp_name") %>
			</caption>
			<tr>
				<td class="label" style="width:22%;">Data ultima modifica</td>
				<td class="content" colspan="2"><%= rsg("grp_modData") %></td>
			</tr>
			<tr>
				<td class="label" style="width:22%;">Utente ultima modifica</td>
				<td class="content" colspan="2"><%= GetAdminName(conn, rsg("grp_modAdmin_id")) %></td>
			</tr>
			<tr>
				<th class="L2">applicato a</th>
				<th class="l2_center">esempio</th>
				<th class="l2_center" width="10%">operazioni</th>
			</tr>
			<% sql = "SELECT * FROM tb_css_styles WHERE style_grp_id=" & rsg("grp_id") & " ORDER BY style_id"
			rss.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
			
			while not rss.eof %>
				<tr>
					<td class="content" title="<%= rss("style_class") & rss("style_pseudoclass") %>">
						<%= rss("style_description") %>
						<% if instr(1, rss("style_class"), "A", vbTextCompare)>0 then %>
							<span class="note">
								<br>Link su un testo di paragrafo
							</span>
						<% end if %>
					</td>
					<td class="content_center" style="height:44px; vertical-align:middle;">
						<iframe src="SitoStiliPreview.asp?ID=<%= rss("style_id") %><%= IIF(rss.recordcount = rss.absoluteposition AND rsg.recordcount = rsg.Absoluteposition, "&last=1", "")%>" frameborder="0" scrolling="No" 
								style="width:99%; height:40px;">
						</iframe>
					</td>
					<td class="content_center">
						<a class="button_l2" href="SitoStiliMod.asp?ID=<%= rss("style_id") %>" title="Personalizza lo stile" <%= ACTIVE_STATUS %>>
							MODIFICA
						</a>
					</td>
				</tr>
				<% rss.movenext
			wend
			
			rss.close%>
		</table>
		<% rsg.movenext
	wend %>
	<br>
</div>
</body>
</html>
<% 
rsg.close
conn.close 
set rsg = nothing
set conn = nothing
set cssO = nothing%>

