<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("AlertConfSalva.asp")
end if

dim conn, sql, rs, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM rel_siti_eventi WHERE rse_id = "& cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly

'--------------------------------------------------------
sezione_testata = "Gestione alert - configurazione - modifica" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'------------------------------------------------------%>

<div id="content_ridotto">
<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="width: 695px;">
		<caption>Modifica configurazione</caption>
		<% 	sql = "SELECT sev_multisito FROM tb_siti_eventi WHERE sev_id = "& rs("rse_evento_id")
			if CBoolean(GetValueList(conn, NULL, sql), false) then %>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">sito:</td>
			<td class="content" colspan="3">
				<%	sql = "SELECT * FROM tb_webs ORDER BY id_webs"
					CALL dropDown(conn, sql, "id_webs", "nome_webs", "tfn_rse_web_id", rs("rse_web_id"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<% 	end if %>
		
		<tr><th colspan="4">E-MAIL</th></tr>
		<tr>
			<td class="label">abilitato:</td>
			<td class="content" colspan="3">
				<input type="checkbox" class="checkbox" name="chk_rse_email_abilitato" value="1" onclick="EmailAbilita(this.checked)"
				<%= Chk(rs("rse_email_abilitato")) %>>
			</td>
		</tr>
		<tr>
			<td class="label">invio all'amministratore:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_rse_email_admin_invio" value="1" <%= Chk(rs("rse_email_admin_invio")) %>>
			</td>
			<td class="label">invio agli utenti:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_rse_email_utenti_invio" value="1" <%= Chk(rs("rse_email_utenti_invio")) %>>
			</td>
		</tr>
		<tr>
			<td class="label" style="white-space: nowrap;">amministratore mittente:</td>
			<td class="content" colspan="3">
				<%	sql = " SELECT *, (admin_cognome "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) & SQL_IfIsNull(conn, "admin_nome", "''") &") AS NOME"& _
						  " FROM tb_admin ORDER BY admin_cognome"
					CALL dropDown(conn, sql, "id_admin", "NOME", "tfn_rse_email_admin_id", rs("rse_email_admin_id"), true, " style=""width:100%;""", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 		if i=lbound(Application("LINGUE")) then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">oggetto:</td>
		<% 		end if %>
			<td class="content" colspan="3">
				<table width="95%" border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td width="5%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
						<td>
							<input type="text" class="text" size="90" maxlength="250" name="tft_rse_email_oggetto_<%= Application("LINGUE")(i) %>" value="<%= rs("rse_email_oggetto_"& application("LINGUE")(i)) %>">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<%next %>
		<tr>
			<td class="label">pagina da inviare:</td>
			<td class="content" colspan="3">
				<% CALL DropDownPages(NULL, "form1", "429", 0, "tfn_rse_email_paginaId", rs("rse_email_paginaId"), true, false) %>
				(*)
			</td>
		</tr>
		<% 	CALL WriteDestinatari("email", "EMAILMANDATORY") %>
		<script type="text/javascript">
			function EmailAbilita(v) {
				var o
				o = document.getElementById("chk_rse_email_admin_invio")
				DisableControl(o, !v)
				o = document.getElementById("chk_rse_email_utenti_invio")
				DisableControl(o, !v)
				o = document.getElementById("tfn_rse_email_admin_id")
				DisableControl(o, !v)
				o = document.getElementById("tfn_rse_email_paginaId")
				DisableControl(o, !v)
				o = document.getElementById("email_contatti")
				DisableControl(o, !v)
				o = document.getElementById("email_admin")
				DisableControl(o, !v)
				
				<%	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				o = document.getElementById("tft_rse_email_oggetto_<%= Application("LINGUE")(i) %>")
				DisableControl(o, !v)
				<%  next %>
			}
			EmailAbilita(document.getElementById('chk_rse_email_abilitato').checked);
		</script>
		
		<% 	'--------------------------------------------------------------------------------------------------FAX
			if FaxAbilitati(conn) then %>
		<tr><th colspan="4">FAX</th></tr>
		<tr>
			<td class="label">abilitato:</td>
			<td class="content" colspan="3">
				<input type="checkbox" class="checkbox" name="chk_rse_fax_abilitato" value="1" onclick="FaxAbilita(this.checked)"
				<%= Chk(rs("rse_fax_abilitato")) %>>
			</td>
		</tr>
		<tr>
			<td class="label">invio all'amministratore:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_rse_fax_admin_invio" value="1" <%= Chk(rs("rse_fax_admin_invio")) %>>
			</td>
			<td class="label">invio agli utenti:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_rse_fax_utenti_invio" value="1" <%= Chk(rs("rse_fax_utenti_invio")) %>>
			</td>
		</tr>
		<tr>
			<td class="label" style="white-space: nowrap;">amministratore mittente:</td>
			<td class="content" colspan="3">
				<%	sql = " SELECT *, (admin_cognome "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) & SQL_IfIsNull(conn, "admin_nome", "''") &") AS NOME"& _
						  " FROM tb_admin ORDER BY admin_cognome"
					CALL dropDown(conn, sql, "id_admin", "NOME", "tfn_rse_fax_admin_id", rs("rse_fax_admin_id"), true, " style=""width:100%;""", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 		if i=lbound(Application("LINGUE")) then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">oggetto:</td>
		<% 		end if %>
			<td class="content" colspan="3">
				<table width="95%" border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td width="5%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
						<td>
							<input type="text" class="text" size="90" maxlength="250" name="tft_rse_fax_oggetto_<%= Application("LINGUE")(i) %>" value="<%= rs("rse_fax_oggetto_"& application("LINGUE")(i)) %>">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<%next %>
		<tr>
			<td class="label">pagina da inviare:</td>
			<td class="content" colspan="3">
				<% CALL DropDownPages(NULL, "form1", "429", 0, "tfn_rse_fax_paginaId", rs("rse_fax_paginaId"), true, false) %>
				(*)
			</td>
		</tr>
		<% 		CALL WriteDestinatari("fax", "FAXMANDATORY") %>
		<script type="text/javascript">
			function FaxAbilita(v) {
				var o
				o = document.getElementById("chk_rse_fax_admin_invio")
				DisableControl(o, !v)
				o = document.getElementById("chk_rse_fax_utenti_invio")
				DisableControl(o, !v)
				o = document.getElementById("tfn_rse_fax_admin_id")
				DisableControl(o, !v)
				o = document.getElementById("tfn_rse_fax_paginaId")
				DisableControl(o, !v)
				o = document.getElementById("fax_contatti")
				DisableControl(o, !v)
				o = document.getElementById("fax_admin")
				DisableControl(o, !v)
			}
			FaxAbilita(document.getElementById('chk_rse_fax_abilitato').checked);
		</script>
		<% 	end if %>
		
		<% 	'--------------------------------------------------------------------------------------------------SMS
			if SMSAbilitati(conn) then %>
		<tr><th colspan="4">SMS</th></tr>
		<tr>
			<td class="label">abilitato:</td>
			<td class="content" colspan="3">
				<input type="checkbox" class="checkbox" name="chk_rse_sms_abilitato" value="1" onclick="SMSAbilita(this.checked)"
				<%= Chk(rs("rse_sms_abilitato")) %>>
			</td>
		</tr>
		<tr>
			<td class="label">invio all'amministratore:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_rse_sms_admin_invio" value="1" <%= Chk(rs("rse_sms_admin_invio")) %>>
			</td>
			<td class="label">invio agli utenti:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_rse_sms_utenti_invio" value="1" <%= Chk(rs("rse_sms_utenti_invio")) %>>
			</td>
		</tr>
		<tr>
			<td class="label" style="white-space: nowrap;">amministratore mittente:</td>
			<td class="content" colspan="3">
				<%	sql = " SELECT *, (admin_cognome "& SQL_concat(conn) &" ' ' "& SQL_concat(conn) & SQL_IfIsNull(conn, "admin_nome", "''") &") AS NOME"& _
						  " FROM tb_admin ORDER BY admin_cognome"
					CALL dropDown(conn, sql, "id_admin", "NOME", "tfn_rse_sms_admin_id", rs("rse_sms_admin_id"), true, " style=""width:100%;""", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 		if i=lbound(Application("LINGUE")) then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">testo:</td>
		<% 		end if %>
			<td class="content" colspan="3">
				<table width="95%" border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td width="5%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
						<td>
							<input type="text" class="text" size="90" maxlength="160" name="tft_rse_sms_testo_<%= Application("LINGUE")(i) %>" value="<%= rs("rse_sms_testo_"& application("LINGUE")(i)) %>">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<%next %>
		<% 		CALL WriteDestinatari("sms", "SMSMANDATORY") %>
		<script type="text/javascript">
			function SMSAbilita(v) {
				var o
				o = document.getElementById("chk_rse_sms_admin_invio")
				DisableControl(o, !v)
				o = document.getElementById("chk_rse_sms_utenti_invio")
				DisableControl(o, !v)
				o = document.getElementById("tfn_rse_sms_admin_id")
				DisableControl(o, !v)
				o = document.getElementById("sms_contatti")
				DisableControl(o, !v)
				o = document.getElementById("sms_admin")
				DisableControl(o, !v)
						
				<%	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				o = document.getElementById("tft_rse_sms_testo_<%= Application("LINGUE")(i) %>")
				DisableControl(o, !v)
				<%  next %>
			}
			SMSAbilita(document.getElementById('chk_rse_sms_abilitato').checked);
		</script>
		<% 	end if %>
		
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori per le rispettive sezioni.
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
</form>
</div>
</body>
</html>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>

<%
conn.close
set rs = nothing
set conn = nothing


'scrive la sezione destinatari
Sub WriteDestinatari(sezione, tipo)
	dim contatti, admin, sezioneId, sql
	select case sezione
		case "email"
			sezioneId = MSG_EMAIL
		case "fax"
			sezioneId = MSG_FAX
		case "sms"
			sezioneId = MSG_SMS
	end select
	
	sql = " SELECT rec_contatto_id FROM rel_siti_eventi_contatti"& _
		  " WHERE rec_tipo_messaggio_id = "& sezioneId &" AND rec_sitoevento_id = "& cIntero(request("ID"))
	contatti = GetValueList(conn, NULL, sql)
	sql = " SELECT rea_admin_id FROM rel_siti_eventi_admin"& _
		  " WHERE rea_tipo_messaggio_id = "& sezioneId &" AND rea_sitoevento_id = "& cIntero(request("ID"))
	admin = GetValueList(conn, NULL, sql) %>
		<tr><th colspan="4" class="L2">DESTINATARI AGGIUNTIVI</th></tr>
		<tr>
			<td class="label">contatti:</td>
			<td class="content" colspan="3">
				<% CALL WriteContactPicker_Input(conn, NULL, "", "", "form1", sezione &"_contatti", contatti, tipo, true, false, false, "")  %>
			</td>
		</tr>
		<tr>
			<td class="label">amministratori:</td>
			<td class="content" colspan="3">
				<% CALL WriteAdminPicker_Input(conn, NULL, "", "form1", sezione &"_admin", admin, tipo, true, false, false, "")  %>
			</td>
		</tr>
<%
End Sub
%>