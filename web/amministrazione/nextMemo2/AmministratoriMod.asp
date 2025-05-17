<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("AmministratoriSalva.asp")
end if

dim i, conn, rs, rsp, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_AMMINISTRATORI"), "id_admin", "AmministratoriMod.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->

<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione utenti area amministrativa - modifica utente"
dicitura.puls_new = "INDIETRO;ACCESSI"
dicitura.link_new = "Amministratori.asp;AmministratoriAccessi.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez() 


sql = "SELECT * FROM tb_admin WHERE ID_admin=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati dell'utente</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="utente precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="utente successiva">
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">Cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_cognome" value="<%= rs("admin_cognome") %>" maxlength="50" size="50">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" style="width:18%;">Nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_nome" value="<%= rs("admin_nome") %>" maxlength="50" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">Email in uscita:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_email" value="<%= rs("admin_email") %>" maxlength="50" size="50">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">Telefono:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_telefono" value="<%= rs("admin_telefono") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">Cellulare:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_cell" value="<%= rs("admin_cell") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">Fax:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_fax" value="<%= rs("admin_fax") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr><th colspan="4">ACCOUNT DI ACCESSO</th></tr>
		<tr>
			<td class="label">Login:</td>
			<td class="content">
				<input type="hidden" name="old_admin_login" value="<%= rs("admin_login") %>">
				<input type="text" class="text" name="tft_admin_login" value="<%= rs("admin_login") %>" maxlength="50" size="20">
				(*)
			</td>
			<td class="label">Scadenza:</td>
			<td class="content">
				<% CALL WriteDataPicker_Input("form1", "tfd_admin_scadenza", rs("admin_scadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">Password:</td>
			<td class="content" colspan="3" style="padding:2px;">
				<a href="javascript:void(0);" class="button_form" onclick="OpenAutoPositionedWindow('AmministratoriPassword.asp?ID=<%= request("ID") %>', 'pws', 402, 240);">
					MODIFICA LA PASSWORD
				</a>
			</td>
		</tr>
		
		
		<!-- assegnazione del permesso MEMO2_DOWNLOAD all'utente -->
		<input type="hidden" value="<%=NEXTMEMO2%>,2" name="chk_perm">

	
		<% if cInteger(Application("NextCom_DefaultWorkGroup"))=0 then %>
			<tr><th colspan="4">GRUPPO DI LAVORO</th></tr>
			<tr>
				<td colspan="4">
					<% sql = " SELECT tb_gruppi.id_gruppo, nome_gruppo, id_rel_dipgruppi FROM tb_gruppi " & _
							 " LEFT JOIN tb_rel_dipgruppi ON (tb_gruppi.id_gruppo=tb_rel_dipgruppi.id_Gruppo " + _
							 " AND tb_rel_dipgruppi.id_impiegato=" & cIntero(request("ID")) & ")" &_
							 " ORDER BY nome_gruppo"
					CALL Write_Relations_Checker(conn, rsp, sql, 2, "id_gruppo", "nome_gruppo", "id_rel_dipgruppi", "gruppi_di_lavoro") %>
				</td>
			</tr>
		<% else %>
			<input type="hidden" name="gruppi_di_lavoro" value="<%= Application("NextCom_DefaultWorkGroup") %>">
		<% end if %>
		<tr><th colspan="4">ACCESSO AI FILES</th></tr>
		<tr>
			<td class="label" rowspan="2">Directory di partenza:</td>
			<td class="content" colspan="3">
				<% CALL WriteFileSystemPicker_Input(application("AZ_ID"), FILE_SYSTEM_DIRECTORY, "images", "", "form1", "tft_admin_dir", rs("admin_dir"), "width:420px;", false, false) %>
			</td>
		</tr>
		<tr>
			<td class="content notes" colspan="3">
				Permette di limitare la gestione dei files da parte dell'utente all'interno della cartella selezionata.
			</td>
		</tr>
		<% sql = "SELECT pro_id FROM mtb_profili" %>
		<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) _
					AND cString(GetValueList(conn, NULL, sql)) <> "" then %>
			<tr><th colspan="4">PROFILI PER LA VISUALIZZAZIONE DEI DOCUMENTI / CIRCOLARI</th></tr>
			<tr>
				<td class="content" colspan="4">
					<% sql = " SELECT * FROM mtb_profili LEFT JOIN mrel_profili_admin " + _
							 " ON (mtb_profili.pro_id = mrel_profili_admin.rpa_profilo_id AND mrel_profili_admin.rpa_admin_id = " & cIntero(request("ID")) & ")" + _
							 " ORDER BY pro_nome_it"
					   CALL Write_Relations_Checker(conn, rsp, sql, 4, "pro_id", "pro_nome_it", "rpa_id", "profili_associati")%>
				</td>
			</tr>
		<% end if %>
		
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
                <input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA AD ELENCO">
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