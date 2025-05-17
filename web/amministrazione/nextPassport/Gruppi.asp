<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<%
dim dicitura
set dicitura = New testata  
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gruppi di lavoro"
dicitura.puls_new = "NUOVO GRUPPO"
dicitura.link_new = "GruppiNew.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_gruppi ORDER BY nome_gruppo"
rs.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco gruppi di lavoro - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<%%>
		<tr>
			<th style="text-align:center; width:5%;">ID</th>
			<th>NOME DEL GRUPPO</th>
			<th colspan="2" style="width: 20%; text-align: center;">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("id_Gruppo") %></td>
				<td class="content"><%= rs("nome_gruppo") %></td>
				<td class="Content_center">
					<a class="button" href="GruppiMod.asp?ID=<%= rs("id_gruppo") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('GRUPPI','<%= rs("id_gruppo") %>');" >
						CANCELLA
					</a>
				</td>
			</tr>
			<%rs.movenext
		wend%>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing%>