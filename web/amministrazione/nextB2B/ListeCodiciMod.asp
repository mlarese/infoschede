<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ListeCodiciSalva.asp")
end if

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_LSTCOD_SQL"), "LstCod_id", "ListeCodiciMod.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione liste codici - modifica"
dicitura.puls_new = "INDIETRO;CODICI"
dicitura.link_new = "ListeCodici.asp;ListeCodiciCodici.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez() 


sql = "SELECT * FROM gtb_lista_codici WHERE lstCod_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table align="right" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="lista codici precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="lista codici successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
			Modifica dati della lista codici
		</caption>
		<tr><th colspan="2">DATI DELLA LISTA DI CODICI</th></tr>
		<tr>
			<td class="label">Codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_LstCod_cod" value="<%= rs("LstCod_cod") %>" maxlength="50" size="40">
			</td>
		</tr>
		<tr>
			<td class="label">Nome:</td>
			<td class="content">
				<input type="text" class="text" name="tft_LstCod_nome" value="<%= rs("LstCod_nome") %>" maxlength="50" size="75">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2">tipo di lista:</td>
			<td class="content" colspan="2">
				<input class="checkbox" type="radio" name="tfn_lstCod_sistema" value="0" <%= chk(NOT rs("lstCod_sistema")) %>> dei clienti
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input class="checkbox" type="radio" name="tfn_lstCod_sistema" value="1" <%= chk(rs("lstCod_sistema")) %>> di sistema
			</td>
		</tr>
		<tr><th colspan="2">NOTE</th></tr>
		<tr>
			<td class="content" colspan="2">
				<textarea style="width:100%;" rows="3" name="tft_lstCod_note"><%= rs("lstCod_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
set rs = nothing
conn.Close
set conn = nothing
%>