<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
'controllo accesso
if Session("COM_ADMIN")="" AND Session("COM_POWER")="" then
	response.redirect "Contatti.asp"
end if
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("RubricheSalva.asp")
end if

dim conn, rs, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action, i
'Titolo della pagina
	Titolo_sezione = "Rubriche - nuova"
'Indirizzo pagina per link su sezione 
		HREF = "Rubriche.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova rubrica</caption>
		<tr><th colspan="2">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label" style="width:22%;">nome rubrica:</td>
			<td class="content">
				<input type="text" class="text" name="tft_nome_rubrica" value="<%= request("tft_nome_rubrica") %>" maxlength="250" size="75">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<% 	if IsAdminCurrent(conn) then %>
		<tr>
			<td class="label">rubrica utilizzata dal sistema:</td>
			<td class="content">
				<input type="checkbox" class="noborder" name="chk_locked_rubrica" <%if request("chk_locked_rubrica") <>"" then%>checked <% end if %>>
			</td>
		</tr>
		<tr>
			<td class="label">rubrica esterna:</td>
			<td class="content">
				<input type="checkbox" class="noborder" name="chk_rubrica_esterna" <%if request("chk_rubrica_esterna") <>"" then%>checked <% end if %>>
			</td>
		</tr>
		<% 	end if %>
		<% if Cinteger(Application("NextCom_DefaultWorkGroup"))=0 then %>
			<% sql = "SELECT *, (NULL) AS id_rel_grupprub FROM tb_gruppi" & _
							 " ORDER BY nome_gruppo" 
			dim rs_group
			set rs_group = Server.CreateObject("ADODB.RecordSet")
			rs_group.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
			
			if rs_group.recordcount > 1 then %>
				<tr><th colspan="2">GRUPPI DI LAVORO COMPETENTI (*)</th></tr>
				<tr>
					<td colspan="2">
						<% CALL Write_Relations_Checker(conn, rs, sql, 2, "id_gruppo", "nome_Gruppo", "id_rel_grupprub", "gruppi") %>
					</td>
				</tr>
			<% else %>
				<input type="hidden" name="gruppi" value="<%= rs_group("id_gruppo") %>"
			<% end if 
			rs_group.close 
			set rs_group = nothing
			%>
		<% else %>
			<input type="hidden" name="gruppi" value="<%= Application("NextCom_DefaultWorkGroup") %>">
		<% end if %>
		<tr><th colspan="2">CONTATTI ASSOCIATI</th></tr>
		<tr>
			<td class="label">singoli contatti:</td>
			<td class="content">
                <% CALL WriteContactPicker_Input(conn, rs, "", "", "form1", "contatti", request("contatti"), "", true, false, false, "") %>
			</td>
		</tr>
		<tr><th colspan="2">DATI VISIBILI NELLA PARTE PUBBLICA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:18%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome pubblico:</td>
			<% 	end if %>
				<td class="content" colspan="5">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_nome_pubblico_rubrica_<%= Application("LINGUE")(i) %>" value="<%= request("tft_nome_pubblico_rubrica_"& Application("LINGUE")(i)) %>" maxlength="500" size="75">
				</td>
			</tr>
		<% next %>
		<tr><th colspan="2">NOTE</th></tr>
		<tr>
			<td class="content" colspan="2">
				<textarea style="width:100%;" rows="5" name="tft_Note_rubrica"><%=request("tft_Note_rubrica")%></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva_avanti" value="SALVA">
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
set rs = nothing
set conn = nothing%>