<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("PortiSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione porti - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Porti.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_PORTI_SQL"), "prt_id", "PortiMod.asp")
end if

sql = "SELECT * FROM gtb_porti WHERE prt_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del porto</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="porto precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="porto successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="3">DATI DEL PORTO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_prt_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("prt_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="60">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">attivo:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_prt_attivo" <%= chk(rs("prt_attivo")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_prt_attivo" <%= chk(not rs("prt_attivo")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_prt_codice" value="<%= rs("prt_codice") %>" maxlength="250" size="20">
			</td>
		</tr>
		<tr>
			<td class="label" style="width:15%;">con spese:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_prt_con_spese" <%= chk(rs("prt_con_spese")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_prt_con_spese" <%= chk(not rs("prt_con_spese")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">con vettore:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_prt_con_vettore" <%= chk(rs("prt_con_vettore")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_prt_con_vettore" <%= chk(not rs("prt_con_vettore")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">scelta sede:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_prt_scelta_sede" <%= chk(rs("prt_scelta_sede")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_prt_scelta_sede" <%= chk(not rs("prt_scelta_sede")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">scelta modalit&agrave; di consegna:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_prt_scelta_modalita_consegna" <%= chk(rs("prt_scelta_modalita_consegna")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_prt_scelta_modalita_consegna" <%= chk(not rs("prt_scelta_modalita_consegna")) %>>
				no
			</td>
		</tr>
		<tr><th colspan="2">DESCRIZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_prt_descrizione_<%= Application("LINGUE")(i) %>"><%= rs("prt_descrizione_" & Application("LINGUE")(i)) %></textarea></td>
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