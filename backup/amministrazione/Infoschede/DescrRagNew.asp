<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("DescrRagSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura, i
set dicitura = New testata 
dicitura.iniz_sottosez(2)
dicitura.sottosezioni(1) = "CONTROLLI RIPARAZIONI"
dicitura.links(1) = "Descrittori.asp"
dicitura.sottosezioni(2) = "GRUPPI"
dicitura.links(2) = "DescrRag.asp"
dicitura.sezione = "Gruppi controlli per riparazioni - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "DescrRag.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo gruppo di controlli per riparazioni</caption>
		<tr><th colspan="4">DATI DEL GRUPPO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
				<% end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_rag_titolo_<%= Application("LINGUE")(i) %>" value="<%= request.form("tft_rag_titolo_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" width="35%">
				<input type="text" class="text" name="tfn_rag_ordine" size="3"
					   value="<%= IIF(request.servervariables("REQUEST_METHOD")<>"POST", CIntero(GetValueList(conn, NULL, "SELECT MAX(rag_ordine) FROM sgtb_descrittori_raggruppamenti")) + 10, request.form("tfn_rag_ordine")) %>">
			</td>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_rag_codice" value="<%= request("tft_rag_codice") %>" maxlength="50" >
			</td>
		</tr>
		<tr><th colspan="4">CARATTERISTICHE ASSOCIATE</th></tr>
		<tr>
			<td colspan="4">
			<% 	'seleziono solo descrittori associati a raggruppamenti non locked
				sql = " SELECT *,"& _
					  " ("& SQL_If(conn, SQL_IsNull(conn, "rag_id"), "des_nome_it", "(des_nome_it "& SQL_Concat(conn) &"' (gruppo: '"& SQL_Concat(conn) &" rag_titolo_it "& SQL_Concat(conn) &"')')") &") AS nome,"& _
					  " (NULL) AS rel"& _
					  " FROM sgtb_descrittori d"& _
					  " LEFT JOIN sgtb_descrittori_raggruppamenti r ON d.des_raggruppamento_id = r.rag_id"& _
					  " ORDER BY des_nome_it"
				CALL Write_Relations_Checker(conn, rs, sql, 2, "des_id", "nome", "rel", "car") %>
			</td>
		</tr>
		<tr><th colspan="4">NOTE INTERNE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_rag_note"><%=request("tft_rag_note")%></textarea>
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