<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("AmministratoriOrarioSalva.asp")
end if

dim i, conn, rs, rsp, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")


%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="Tools_Passport.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione utenti area amministrativa - modifica orario"
dicitura.puls_new = "INDIETRO;NUOVO ORARIO"
dicitura.link_new = "Amministratori.asp;AmministratoriOrarioMod.asp?ID_ADMIN=" & cIntero(request("ID_ADMIN")) & "&NEW=1"
dicitura.scrivi_con_sottosez()  

sql = "SELECT * FROM tb_admin_orario WHERE ao_id_admin=" & cIntero(request("ID_ADMIN"))
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			Orario di lavoro di <%=GetAdminName(conn, cIntero(request("ID_ADMIN")))%>
		</caption>
		<tr>
			<th rowspan="2" class="center">DATA DAL</th>
			<th rowspan="2" class="center">DATA AL</th>
			<th colspan="7" class="center">MINUTI DI LAVORO</th>
			<th rowspan="2" colspan="2" class="center">OPERAZIONI</th>
		</tr>
		<tr>
			<th class="l2_center">LUNEDI'</th>
			<th class="l2_center">MARTEDI'</th>
			<th class="l2_center">MERCOLEDI'</th>
			<th class="l2_center">GIOVEDI'</th>
			<th class="l2_center">VENERDI'</th>
			<th class="l2_center">SABATO</th>
			<th class="l2_center">DOMENICA</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<%if cInteger(request("ID")) = rs("ao_id") then%>
					<td class="content">
						<% CALL WriteDataPicker_Input_Manuale2("form1", "tfd_ao_data_dal", rs("ao_data_dal"), "", "/", true, true, LINGUA_ITALIANO, "", false, "", "") %>
					</td>
					<td class="content">
						<% CALL WriteDataPicker_Input_Manuale2("form1", "tfd_ao_data_al", rs("ao_data_al"), "", "/", true, true, LINGUA_ITALIANO, "", false, "", "") %>
					</td>
					<td class="content">
						<input type="text" class="number" name="tfn_ao_min_lav_lun" value="<%= rs("ao_min_lav_lun") %>" maxlength="3" size="5">
					</td>
					<td class="content">
						<input type="text" class="number" name="tfn_ao_min_lav_mar" value="<%= rs("ao_min_lav_mar") %>" maxlength="3" size="5">
					</td>
					<td class="content">
						<input type="text" class="number" name="tfn_ao_min_lav_mer" value="<%= rs("ao_min_lav_mer") %>" maxlength="3" size="5">
					</td>
					<td class="content">
						<input type="text" class="number" name="tfn_ao_min_lav_gio" value="<%= rs("ao_min_lav_gio") %>" maxlength="3" size="5">
					</td>
					<td class="content">
						<input type="text" class="number" name="tfn_ao_min_lav_ven" value="<%= rs("ao_min_lav_ven") %>" maxlength="3" size="5">
					</td>
					<td class="content">
						<input type="text" class="number" name="tfn_ao_min_lav_sab" value="<%= rs("ao_min_lav_sab") %>" maxlength="3" size="5">
					</td>
					<td class="content">
						<input type="text" class="number" name="tfn_ao_min_lav_dom" value="<%= rs("ao_min_lav_dom") %>" maxlength="3" size="5">
					</td>
					<td class="content_right" style="vertical-align:middle;">
						<input type="submit" class="button" name="salva" value="SALVA">
					</td>
					<td class="content_right" style="vertical-align:middle;">
						<input type="button" class="button" name="annulla" value="ANNULLA" onclick="document.location='AmministratoriOrarioMod.asp?ID_ADMIN=<%=cIntero(request("ID_ADMIN"))%>';">
					</td>
				<% else %>
					<td class="content_center"><%= rs("ao_data_dal") %></td>
					<td class="content_center"><%= rs("ao_data_al") %></td>
					<td class="content_center"><%= rs("ao_min_lav_lun") %></td>
					<td class="content_center"><%= rs("ao_min_lav_mar") %></td>
					<td class="content_center"><%= rs("ao_min_lav_mer") %></td>
					<td class="content_center"><%= rs("ao_min_lav_gio") %></td>
					<td class="content_center"><%= rs("ao_min_lav_ven") %></td>
					<td class="content_center"><%= rs("ao_min_lav_sab") %></td>
					<td class="content_center"><%= rs("ao_min_lav_dom") %></td>
					<td class="content_center">
						<a class="button" href="AmministratoriOrarioMod.asp?ID_ADMIN=<%=cIntero(request("ID_ADMIN"))%>&ID=<%= rs("ao_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="content_center">
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('AMMINISTRATORI_ORARIO','<%= rs("ao_id") %>');" >
							CANCELLA
						</a>
					</td>
				<% end if %>
			</tr>
			<%rs.movenext
		wend
		if request("NEW")<>"" then%>
			<tr>
				<input type="hidden" name="tfn_ao_id_admin" value="<%=cIntero(request("ID_ADMIN"))%>">
				<td class="content">
					<% CALL WriteDataPicker_Input_Manuale2("form1", "tfd_ao_data_dal", request("tfd_ao_data_dal"), "", "/", false, true, LINGUA_ITALIANO, "", false, "", "") %>
				</td>
				<td class="content">
					<% CALL WriteDataPicker_Input_Manuale2("form1", "tfd_ao_data_al", request("tfd_ao_data_al"), "", "/", false, true, LINGUA_ITALIANO, "", false, "", "") %>
				</td>
				<td class="content">
					<input type="text" class="number" name="tfn_ao_min_lav_lun" value="<%= request("tfn_ao_min_lav_lun") %>" maxlength="3" size="5">
				</td>
				<td class="content">
					<input type="text" class="number" name="tfn_ao_min_lav_mar" value="<%= request("tfn_ao_min_lav_mar") %>" maxlength="3" size="5">
				</td>
				<td class="content">
					<input type="text" class="number" name="tfn_ao_min_lav_mer" value="<%= request("tfn_ao_min_lav_mer") %>" maxlength="3" size="5">
				</td>
				<td class="content">
					<input type="text" class="number" name="tfn_ao_min_lav_gio" value="<%= request("tfn_ao_min_lav_gio") %>" maxlength="3" size="5">
				</td>
				<td class="content">
					<input type="text" class="number" name="tfn_ao_min_lav_ven" value="<%= request("tfn_ao_min_lav_ven") %>" maxlength="3" size="5">
				</td>
				<td class="content">
					<input type="text" class="number" name="tfn_ao_min_lav_sab" value="<%= request("tfn_ao_min_lav_sab") %>" maxlength="3" size="5">
				</td>
				<td class="content">
					<input type="text" class="number" name="tfn_ao_min_lav_dom" value="<%= request("tfn_ao_min_lav_dom") %>" maxlength="3" size="5">
				</td>
				<td class="content_right" style="vertical-align:middle;">
					<input type="submit" class="button" name="salva" value="SALVA">
				</td>
				<td class="content_right" style="vertical-align:middle;">
					<input type="button" class="button" name="annulla" value="ANNULLA" onclick="document.location='AmministratoriOrarioMod.asp?ID_ADMIN=<%=cIntero(request("ID_ADMIN"))%>';">
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="11">
				(*) Campi obbligatori.
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set rsp = nothing
set conn = nothing
%>