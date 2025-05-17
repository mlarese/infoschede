<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ContattiDescrittoriGruppiSalva.asp")
end if
%>
<% 
CALL CheckAutentication(Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE"))

dim Titolo_sezione, action, HREF
Titolo_sezione = "Gruppi di caratteristiche - nuovo"
HREF = "ContattiDescrittoriGruppi.asp"
Action = "INDIETRO"
SSezioniText = ""
SSezioniLink = ""
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo gruppo di caratteristiche</caption>
		<tr><th colspan="4">DATI DEL GRUPPO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
				<% end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_icr_titolo_<%= Application("LINGUE")(i) %>" value="<%= request.form("tft_icr_titolo_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tfn_icr_ordine" size="3"
					   value="<%= IIF(request.servervariables("REQUEST_METHOD")<>"POST", CIntero(GetValueList(conn, NULL, "SELECT MAX(icr_ordine) FROM tb_indirizzario_carattech_raggruppamenti")) + 1, request.form("tfn_icr_ordine")) %>">
			</td>
		</tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content" style="width:35%;">
				<input type="text" class="text" name="tft_icr_codice" value="<%= request("tft_icr_codice") %>" maxlength="50" >
			</td>
			<td class="label">Interno al sitema:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_icr_di_sistema" <%= chk(cInteger(request("tfn_icr_di_sistema"))>0) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_icr_di_sistema" <%= chk(cInteger(request("tfn_icr_di_sistema"))=0) %>>
				no
			</td>
		</tr>
		<tr><th colspan="4">CARATTERISTICHE ASSOCIATE</th></tr>
		<tr>
			<td colspan="4">
			<% 	'seleziono solo descrittori associati a raggruppamenti non locked
				sql = " SELECT *,"& _
					  " ("& SQL_If(conn, SQL_IsNull(conn, "icr_id"), "ict_nome_it", "(ict_nome_it "& SQL_Concat(conn) &"' (gruppo: '"& SQL_Concat(conn) &"icr_titolo_it "& SQL_Concat(conn) &"')')") &") AS nome,"& _
					  " (NULL) AS rel"& _
					  " FROM tb_indirizzario_carattech d"& _
					  " LEFT JOIN tb_indirizzario_carattech_raggruppamenti r ON d.ict_raggruppamento_id = r.icr_id"& _
					  " ORDER BY ict_nome_it"
				CALL Write_Relations_Checker(conn, rs, sql, 2, "ict_id", "nome", "rel", "car") %>
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
conn.close
set conn = nothing
%>