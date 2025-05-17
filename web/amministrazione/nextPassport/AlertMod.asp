<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("AlertSalva.asp")
end if

dim i, conn, rs, rsv, rsd, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_ALERT"), "sev_id", "AlertMod.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione alert - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Alert.asp"
dicitura.scrivi_con_sottosez()


sql = "SELECT * FROM tb_siti_eventi WHERE sev_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati dell'alert</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="alert precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="alert successivo">
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_sev_nome_IT" value="<%= Server.HtmlEncode(cString(rs("sev_nome_IT"))) %>" maxlength="250" size="75">
			</td>
		</tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_sev_codice" value="<%= rs("sev_codice") %>" maxlength="50" size="50">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">applicazione:</td>
			<td class="content" colspan="3">
				<%	sql = "SELECT * FROM tb_siti WHERE " & SQL_IsTrue(conn, "sito_Amministrazione") & " ORDER BY sito_nome"
					CALL DropDown(conn, sql, "id_sito", "sito_nome", "tfn_sev_sito_id", rs("sev_sito_id"), true, " style=""width:100%;""", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">abilitato:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_sev_abilitato" value="1" <%= Chk(rs("sev_abilitato")) %>>
			</td>
			<td class="label">multisito:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_sev_multisito" value="1" <%= Chk(rs("sev_multisito")) %>>
			</td>
		</tr>
		<tr><th colspan="4">CONFIGURAZIONI</th></tr>
		<tr>
			<td colspan="4">
				<% sql = " SELECT * " + _
					     " FROM rel_siti_eventi e"& _
						 " LEFT JOIN tb_webs w ON e.rse_web_id = w.id_webs"& _
                         " WHERE rse_evento_id = "& cIntero(request("ID")) & _
                         " ORDER BY id_webs"
                rsv.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label" colspan="5" style="width:74%">
							<% if rsv.eof then %>
								Nessuna configurazione inserita.
							<% else %>
								Trovati n&ordm; <%= rsv.recordcount %> record
							<% end if %>
						</td>
						<% 	if session("PASS_ADMIN") <> "" then %>
						<td colspan="2" class="content_right" style="white-space: nowrap;padding-right: 5px;">
							<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra l'inserimento di una nuova configurazione" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedWindow('AlertConfNew.asp?ALERT_ID=<%= request("ID") %>', 'Configurazione', 700, 600)">
								NUOVA CONFIGURAZIONE
							</a>
						</td>
						<% 	end if %>
					</tr>
					<% 	if not rsv.eof then %>
                        <tr>
							<th class="L2" width="20%">APPLICATIVO</th>
							<th class="L2">DESTINATARI</th>
						    <th class="l2_center" width="5%">EMAIL</th>
							<th class="l2_center" width="5%">FAX</th>
							<th class="l2_center" width="5%">SMS</th>
			        		<th class="l2_center" width="12%" colspan="2">OPERAZIONI</th>
    					</tr>
					    <% 	while not rsv.eof %>
                        <tr>
					        <td class="content">
								<% if CString(rsv("nome_webs")) <> "" then %>
									<%= rsv("nome_webs")%>
								<% else 
									sql = "SELECT nome_webs FROM tb_webs ORDER BY nome_webs " %>
									<%= getvaluelist(conn, null, sql) %>
								<% end if %>
							</td>
							<td>
								<table cellspacing="1" width="100%">
									<% if rsv("rse_email_abilitato") then 
										CALL WriteDestinatari(conn, rsd, "email", rsv("rse_id"))
									end if 
									if rsv("rse_fax_abilitato") then 
										CALL WriteDestinatari(conn, rsd, "fax", rsv("rse_id"))
									end if 
									if rsv("rse_sms_abilitato") then 
										CALL WriteDestinatari(conn, rsd, "sms", rsv("rse_id"))
									end if 
									%>
								</table>
							</td>
							<td class="content_center">
								<input type="checkbox" class="checkbox" disabled name="email" value="1" <%= Chk(rsv("rse_email_abilitato")) %>>
							</td>
							<td class="content_center">
								<input type="checkbox" class="checkbox" disabled name="fax" value="1" <%= Chk(rsv("rse_fax_abilitato")) %>>
							</td>
							<td class="content_center">
								<input type="checkbox" class="checkbox" disabled name="sms" value="1" <%= Chk(rsv("rse_sms_abilitato")) %>>
							</td>
							<td class="content_center">
								<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica dei dati della configurazione" <%= ACTIVE_STATUS %>
								   onclick="OpenAutoPositionedScrollWindow('AlertConfMod.asp?ID=<%= rsv("rse_id") %>', 'Configurazione', 700, 600, true)">
									MODIFICA
								</a>
							</td>
							<td class="content_center">
							<% 	if session("PASS_ADMIN") = "" then %>
								<a class="button_disabled" title="Configurazione non cancellabile: non si hanno i permessi necessari.">
									CANCELLA
								</a>
							<% 	else %>
								<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione della configurazione"
								   onclick="OpenDeleteWindow('ALERT_CONF','<%= rsv("rse_id") %>');">
									CANCELLA
								</a>
							<% 	end if %>
							</td>
						</tr>
					<%			rsv.MoveNext
							wend
						end if %>
				</table>
				<% 	rsv.close %>
			</td>
		</tr>
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
<% conn.close
set rs = nothing
set rsv = nothing
set rsd = nothing
set conn = nothing



'scrive la sezione destinatari
Sub WriteDestinatari(conn, rs, sezione, eventoId)
	dim sezioneId, sql
	select case sezione
		case "email"
			sezioneId = MSG_EMAIL
		case "fax"
			sezioneId = MSG_FAX
		case "sms"
			sezioneId = MSG_SMS
	end select %>
	<tr>
		<td class="label" rowspan="2"><%= sezione %>:</td>
		<td class="label"> contatti:</td>
		<td class="content" colspan="3">
			<%
			sql = " SELECT * FROM rel_siti_eventi_contatti INNER JOIN v_indirizzario ON rel_siti_eventi_contatti.rec_contatto_id = v_indirizzario.Idelencoindirizzi "& _
				  " WHERE rec_tipo_messaggio_id = "& sezioneId &" AND rec_sitoevento_id = "& eventoId & " ORDER BY ModoRegistra"
			rs.open sql, conn, adOpenStatic, AdLockOptimistic, adCmdText
			if rs.eof then %>
				--
			<% else 
				while not rs.eof%>
					<%=ContactFullName(rs)%><br>
					<% rs.movenext
				wend
			end if
			rs.close
			%>
		</td>
	</tr>
	<tr>
		<td class="label">amministratori:</td>
		<td class="content" colspan="3">
			<%
			sql = " SELECT * FROM rel_siti_eventi_admin INNER JOIN tb_admin ON rel_siti_eventi_admin.rea_admin_id = tb_admin.id_admin "& _
				  " WHERE rea_tipo_messaggio_id = "& sezioneId &" AND rea_sitoevento_id = "& eventoId & " ORDER BY admin_cognome, admin_nome "
			rs.open sql, conn, adOpenStatic, AdLockOptimistic, adCmdText
			if rs.eof then %>
				--
			<% else 
				while not rs.eof%>
					<%=rs("admin_cognome")%>&nbsp;<%=rs("admin_nome")%><br>
					<% rs.movenext
				wend
			end if
			rs.close
			%>
		</td>
	</tr>
<% End Sub

%>