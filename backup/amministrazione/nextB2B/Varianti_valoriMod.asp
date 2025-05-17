<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("Varianti_valoriSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica dati valore" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM gtb_valori WHERE val_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>modifica dati valore</caption>
			<tr><th colspan="4">DATI DEL VALORE</th></tr>
			<tr>
				<td class="label" rowspan="3">codici del valore:</td>
				<td class="label">interno:</td>
				<td class="content">
					<input type="text" class="text" name="tft_val_cod_int" value="<%= rs("val_cod_int") %>" maxlength="50" size="10">
				</td>
				<td class="note" style="vertical-align:top;" rowspan="3">
					Parte di codice proprio del valore che andr&agrave; a comporre il codice finale di ogni variate dell'articolo.
				</td>
			</tr>
			<tr>
				<td class="label">alternativo:</td>
				<td class="content">
					<input type="text" class="text" name="tft_val_cod_alt" value="<%= rs("val_cod_alt") %>" maxlength="50" size="10">
				</td>
			</tr>
			<tr>
				<td class="label">produttore:</td>
				<td class="content">
					<input type="text" class="text" name="tft_val_cod_pro" value="<%= rs("val_cod_pro") %>" maxlength="50" size="10">
				</td>
			</tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_val_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("val_nome_"& Application("LINGUE")(i)) %>" maxlength="100" size="50">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
			<%next %>
			<tr>
				<td class="label">ordine:</td>
				<td class="content" colspan="3">
					<input type="text" class="text" name="tfn_val_ordine" value="<%= rs("val_ordine") %>" size="4" maxlength="4">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label">icona:</td>
				<td class="content" colspan="3">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_val_icona", rs("val_icona"), "", FALSE) %>
				</td>
			</tr>
			<tr><th colspan="4">DESCRIZIONE</th></tr>
			<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="4">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="2" name="tft_val_descr_<%= Application("LINGUE")(i) %>"><%= rs("val_descr_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			<% next %>
			<tr>
				<td class="footer" colspan="4">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
		</table>
	</form>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>