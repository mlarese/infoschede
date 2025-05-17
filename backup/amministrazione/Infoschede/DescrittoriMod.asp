<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("DescrittoriSalva.asp")
end if

dim i, conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(2)
dicitura.sottosezioni(1) = "CONTROLLI RIPARAZIONI"
dicitura.links(1) = "Descrittori.asp"
dicitura.sottosezioni(2) = "GRUPPI"
dicitura.links(2) = "DescrRag.asp"
dicitura.sezione = "Gestione controlli per riparazioni - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Descrittori.asp"
dicitura.scrivi_con_sottosez() 

sql = "SELECT * FROM sgtb_descrittori WHERE des_ID=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Modifica dati caratteristica</caption>
		<tr><th colspan="4">DATI DELLA CARATTERISTICA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
		<% 	end if %>
			<td class="content" colspan="3">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_des_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("des_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
			<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next%>
		<tr>
			<td class="label">tipo:</td>
			<td class="content">
				<% CALL DesDropTipi("tfn_des_tipo", "disabled", rs("des_tipo")) %>
			</td>
			<td class="label" nowrap>caratteristica principale:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_des_principale" <%= Chk(rs("des_principale")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_des_principale" <%= Chk(NOT rs("des_principale")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">immagine:</td>
			<td class="content" colspan="3">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_des_img", rs("des_img"), """ class=""scheda1", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">gruppo:</td>
			<td class="content" colspan="3">
				<% CALL dropDown(conn, "SELECT * FROM sgtb_descrittori_raggruppamenti ORDER BY rag_titolo_it", "rag_id", "rag_titolo_it", "tfn_des_raggruppamento_id", rs("des_raggruppamento_id"), false, "", LINGUA_ITALIANO) %>
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