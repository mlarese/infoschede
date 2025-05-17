<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>
<% response.Buffer = false %>

<!--#INCLUDE FILE="Intestazione.asp" -->
<!--#INCLUDE FILE="../nextCom/Imports/Tools_Import.asp" -->

<% 

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Sostituzione link"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoAnalisi.asp"
dicitura.scrivi_con_sottosez()  


dim conn, rs, rsd, sql, old_link, new_link, new_string
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")

old_link = ParseSQL(request("link_da_sostituire"), adChar)
new_link = ParseSQL(request("link_nuovo"), adChar)
sql = " SELECT html, format FROM tb_layers WHERE id_pag IN (SELECT id_page FROM tb_pages WHERE id_webs = "&Session("AZ_ID")&") AND " & _
	  " 	html LIKE '%"&old_link&"%' AND format LIKE '%"&old_link&"%' "

%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Sostituzione dei link presenti nei layer del sito</caption>
        <tr><th colspan="3">PARAMETRI DI SOSTITUZIONE</th></tr>
		<%
		if request("sostituisci")<>"" AND old_link<>"" AND new_link<>"" then
			rsd.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
			conn.beginTrans
			while not rsd.eof
				new_string = Replace(rsd("html"), old_link, new_link)
				rsd("html") = new_string
				
				new_string = Replace(rsd("format"), old_link, new_link)
				rsd("format") = new_string
				
				rsd.Update
				rsd.moveNext
			wend
			conn.commitTrans
			
			%>
					<tr>
						<td class="content_b" colspan="3">SOSTITUZIONE COMPLETATA</td>
					</tr>
					<tr>
					<td class="content" colspan="3">
						Hai sostituito la stringa "<b><%=old_link%></b>" con la stringa "<b><%=new_link%></b>"			
					</td>
				</tr>
					<tr>
						<td class="footer" colspan="3" style="border-bottom:1px solid #999999;">
							<a class="button" href="SitoAnalisi.asp">FINE</a>
						</td>
					</tr>
				</table>
			</div>
			<%
			response.end
		end if
		
		%>
		
		
		<% if not (request("continua")<>"" AND request("link_da_sostituire")<>"" AND request("link_nuovo")<>"") then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:22%;">link da sostituire:</td>
				<td class="content" colspan="2">
					<input type="text" class="text" name="link_da_sostituire" value="<%= request("link_da_sostituire") %>" maxlength="255" size="105">
					&nbsp;ad es. "http://rallo2011.next-aim.local"
				</td>
			</tr>
			<tr>
				<td class="label" style="width:22%;">link nuovo:</td>
				<td class="content" colspan="2">
					<input type="text" class="text" name="link_nuovo" value="<%= request("link_nuovo") %>" maxlength="255" size="105">
					&nbsp;ad es. "http://www.agenziarallo.it"
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:20%;" type="submit" class="button" name="continua" value="AVANTI &gt;&gt;">
				</td>
			</tr>
			</form>
        <% else %>
			<form action="" method="post" id="form1" name="form1">
			<% 
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText  %>
			<input type="hidden" name="link_da_sostituire" value="<%= old_link %>">
			<input type="hidden" name="link_nuovo" value="<%= new_link %>">
			<tr>
				<td class="content" colspan="3">
					Vuoi sostituire la stringa "<b><%=old_link%></b>" con la stringa "<b><%=new_link%></b>"
					<!--<br><%= sql%><br>-->				
				</td>
			</tr>
            <tr>
				<td class="label" style="width:18%;">n. link da sostituire:</td>
				<td class="content">
					<%= rs.recordCount %>
				</td>
			</tr>
            <tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:20%;" type="submit" class="button" name="sostituisci" value="SOSTITUISCI &gt;&gt;">
				</td>
			</tr>
			</form>
			<% rs.close %>
        <% end if %>
	</table>
</div>

<%
conn.close
set rs = nothing
set rsd = nothing
set conn = nothing
%>