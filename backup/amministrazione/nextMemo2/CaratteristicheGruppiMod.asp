<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("CaratteristicheGruppiSalva.asp")
end if

dim i, conn, rs, rsc, sql, disabled
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gruppi di caratteristiche- modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "CaratteristicheGruppi.asp"
dicitura.scrivi_con_sottosez() 

sql = "SELECT * FROM mtb_carattech_raggruppamenti WHERE ctr_ID=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Modifica dati gruppo di caratteristiche</caption>
		<tr><th colspan="4">DATI DEL GRUPPO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
				<% end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" <%= disabled %> class="text" name="tft_ctr_titolo_<%= Application("LINGUE")(i) %>" value="<%= rs("ctr_titolo_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then
						response.write "(*)"
						if disabled <> "" then %>
					<input type="hidden" name="tft_ctr_titolo_it" value="<%= rs("ctr_titolo_it") %>" maxlength="250" size="75">
				<%		end if
					end if %>
				</td>
			</tr>
		<%next %>
		</tr>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="3">
				<input type="text" <%= disabled %> class="text" name="tfn_ctr_ordine" size="3" value="<%= rs("ctr_ordine") %>">
			</td>
		</tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content" style="width:35%;">
				<input type="text" <%= disabled %> class="text" name="tft_ctr_codice" value="<%= rs("ctr_codice") %>" maxlength="50" >
			</td>
			<td class="label">Interno al sitema:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_ctr_di_sistema" <%= chk(rs("ctr_di_sistema")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_ctr_di_sistema" <%= chk(not rs("ctr_di_sistema")) %>>
				no
			</td>
		</tr>
		<tr><th colspan="4">CARATTERISTICHE</th></tr>
		<tr>
			<td colspan="4">
			<% 	'seleziono solo descrittori associati a raggruppamenti non locked
				sql = " SELECT *,"& _
					  " ("& SQL_If(conn, SQL_IsNull(conn, "ctr_id"), "ct_nome_it", "(ct_nome_it "& SQL_Concat(conn) &"' (gruppo: '"& SQL_Concat(conn) &" ctr_titolo_it "& SQL_Concat(conn) &"')')") &") AS nome,"& _
					  " ("& SQL_If(conn, "ct_raggruppamento_id = "& cIntero(request("ID")), "1", "NULL") &") AS rel"& _
					  " FROM mtb_carattech d"& _
					  " LEFT JOIN mtb_carattech_raggruppamenti r ON d.ct_raggruppamento_id = r.ctr_id"& _
					  " ORDER BY ct_nome_it"
				CALL Write_Relations_Checker(conn, rsc, sql, 2, "ct_id", "nome", "rel", "car") %>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
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
rs.close
conn.close
set rs = nothing
set rsc = nothing
set conn = nothing
%>