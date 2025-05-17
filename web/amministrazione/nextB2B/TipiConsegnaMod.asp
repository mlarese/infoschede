<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("TipiConsegnaSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tipi consegna - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "TipiConsegna.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_TIPI_CONSEGNA_SQL"), "tco_id", "TipiConsegnaMod.asp")
end if

sql = "SELECT * FROM gtb_tipo_consegna WHERE tco_id = " & cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del tipi consegna</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="tipo consegna precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="tipo consegna successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="3">DATI DEL TIPO CONSEGNA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_tco_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("tco_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="60">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_tco_ordine" value="<%= rs("tco_ordine") %>" maxlength="10" size="3">
			</td>
		</tr>
		<tr><th colspan="2">DESCRIZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_tco_descrizione_<%= Application("LINGUE")(i) %>"><%= rs("tco_descrizione_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="footer" colspan="3">
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