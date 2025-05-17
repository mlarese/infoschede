<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
'controllo permessi
CALL CheckAutentication(session("PASS_ADMIN") <> "")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniParamsGruppiSalva.asp")
end if

dim i, conn, rs, rsc, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(3)
dicitura.sottosezioni(1) = "APPLICAZIONI"
dicitura.links(1) = "Applicazioni.asp"
dicitura.sottosezioni(2) = "PARAMETRI"
dicitura.links(2) = "ApplicazioniParams.asp"
dicitura.sottosezioni(3) = "GRUPPI DI PARAMETRI"
dicitura.links(3) = "ApplicazioniParamsGruppi.asp"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "ApplicazioniParamsGruppi.asp"
dicitura.sezione = "Gestione gruppi di parametri - modifica"
dicitura.scrivi_con_sottosez()

sql = "SELECT * FROM tb_siti_descrittori_raggruppamenti WHERE sdr_ID=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Modifica dati gruppo di parametri delle applicazioni</caption>
		<tr><th colspan="4">DATI DEL GRUPPO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
				<% end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_sdr_titolo_<%= Application("LINGUE")(i) %>" value="<%= rs("sdr_titolo_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then
						response.write "(*)"
					end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tfn_sdr_ordine" size="3" value="<%= rs("sdr_ordine") %>">
			</td>
		</tr>
		<tr><th colspan="4">PARAMETRI</th></tr>
		<tr>
			<td colspan="4">
			<% 	'seleziono solo descrittori associati a raggruppamenti non locked
				sql = " SELECT *,"& _
					  " ("& SQL_If(conn, SQL_IsNull(conn, "sdr_id"), "sid_nome_it", "(sid_nome_it "& SQL_Concat(conn) &"' (gruppo: '"& SQL_Concat(conn) &" sdr_titolo_it "& SQL_Concat(conn) &"')')") &") AS nome,"& _
					  " ("& SQL_If(conn, "sid_raggruppamento_id = "& cIntero(request("ID")), "1", "NULL") &") AS rel"& _
					  " FROM tb_siti_descrittori d"& _
					  " LEFT JOIN tb_siti_descrittori_raggruppamenti r ON d.sid_raggruppamento_id = r.sdr_id"& _
					  " ORDER BY sid_nome_it"
				CALL Write_Relations_Checker(conn, rsc, sql, 2, "sid_id", "nome", "rel", "car") %>
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