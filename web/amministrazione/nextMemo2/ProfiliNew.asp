<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ProfiliSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, data
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione profili - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Profili.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo profilo</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_pro_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_pro_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					<% 	if i = 0 then %>(*)<% end if %>
				</td>
			</tr>
		<% next %>

		<tr><th colspan="4">ANAGRAFICHE ASSOCIATE</th></tr>
		<% if cBoolean(Session("CONDIVISIONE_PUBBLICA"), false) then %>
		<tr>
			<td class="label">utenti area riservata:</td>
			<td class="content" colspan="3">
				<% CALL WriteContactPicker_Input(conn, NULL, " ut_ID IN (SELECT rel_ut_id FROM rel_utenti_sito WHERE rel_permesso = 1 OR rel_permesso = 2) ", "", "form1", "utenti_associati", "", "LOGINMANDATORY", true, false, false, "")  %>
			</td>
		</tr>
		<% end if %>
		<% if cBoolean(Session("CONDIVISIONE_INTERNA"), false) then %>
		<tr>
			<td class="label">utenti area amministrativa:</td>
			<td class="content" colspan="3">
				<% CALL WriteAdminPicker_Input(conn, NULL, " ID_admin IN (SELECT admin_id FROM rel_admin_sito WHERE sito_id = 36) ", "form1", "admin_associati", "", "", true, false, false, "")  %>
			</td>
		</tr>
		<% end if %>
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
