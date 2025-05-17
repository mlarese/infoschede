<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("TrasportatoriSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, i
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione trasportatori - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Trasportatori.asp"
dicitura.scrivi_con_sottosez()  

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo trasportatore</caption>
		<tr><th colspan="2">DATI DEL TRASPORTATORE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_tra_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tra_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="60">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_tra_codice" value="<%= request("tft_tra_codice") %>" maxlength="250" size="20">
			</td>
		</tr>
		<tr>
			<td class="label">attivo:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_tra_attivo" <%= chk(cInteger(request("tfn_tra_attivo"))>0) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_tra_attivo" <%= chk(cInteger(request("tfn_tra_attivo"))=0) %>>
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
							<td><textarea style="width:100%;" rows="4" name="tft_tra_descrizione_<%= Application("LINGUE")(i) %>"><%= request("tft_tra_descrizione_" & Application("LINGUE")(i)) %></textarea></td>
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