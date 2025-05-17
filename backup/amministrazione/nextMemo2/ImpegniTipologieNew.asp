<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ImpegniTipologieSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<% 	
dim dicitura, data
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tipologie impegni/appuntamenti - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "ImpegniTipologie.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova tipologia</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_tim_nome_<%= Application("LINGUE")(i) %>" 
								value="<%= request("tft_tim_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					<% 	if i = 0 then %>(*)<% end if %>
				</td>
			</tr>
		<% next %>
		<tr>
			<td class="label_no_width">colore dei contentuti</td>
			<td class="content">
				<% CALL WriteColorPicker_Input("form1", "tft_tim_colore", request("tft_tim_colore"), "", false, true, "") %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">descrizione:</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="16" alt="" border="0" style="vertical-align: top;">
					<textarea name="tft_tim_descrizione_<%= Application("LINGUE")(i) %>" 
								style="width:94.5%;" rows="4"><%= request("tft_tim_descrizione_"& Application("LINGUE")(i)) %></textarea>
				</td>
			</tr>
		<%next %>
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
