<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<% Server.ScriptTimeout = 1073741824 %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<% 
'*****************************************************************************************************************
'verifica dei permessi
if request("STAND_ALONE") = "" then
	CALL VerificaPermessiUtente(true)
end if
'*****************************************************************************************************************

'url da chiaamre direttametne:
'/amministrazione/library/database/archivio__IMPORT_log_framework.asp?STAND_ALONE=1
%>
<html>
<head>
	<title>Archiviazione log_framework</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();">

<table width="100%" cellspacing="0" cellpadding="0" border="0">
	<caption style="border:0px;">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
				<td align="right" style="padding-right:10px;">
					<a href="javascript:close();" class="menu" name="top">CHIUDI</a>
				</td>
	  		</tr>
		</table>
	</caption>
</table>
<br>
<!--

--spostamento manuale dei record

INSERT INTO infogross_archivio.dbo.log_framework(log_id, log_table_nome, log_record_id, log_codice, log_descrizione, log_data, log_admin_id, log_user_id, log_http_request, log_application_id, log_admin_name, log_user_name, log_application_name)
SELECT TOP 1 log_id, log_table_nome, log_record_id, log_codice, log_descrizione, log_data, log_admin_id, log_user_id, log_http_request, log_application_id
, (SELECT admin_nome + ' ' + admin_cognome FROM tb_admin WHERE ID_ADMIN = log_admin_id) AS log_admin_name, (SELECT NomeElencoIndirizzi + ' ' + CognomeElencoIndirizzi + ' - ' + NomeOrganizzazioneElencoIndirizzi FROM tb_Indirizzario WHERE IDElencoIndirizzi IN (SELECT ut_NextCom_ID FROM tb_Utenti WHERE ut_ID = log_user_id)) AS log_user_name, (SELECT sito_nome FROM tb_siti WHERE id_sito = log_application_id) AS log_application_name 
FROM log_framework 
WHERE log_data < CONVERT(DATETIME, '2013-07-05 00:00:00',102)
AND log_id NOT IN (SELECT log_id FROM infogross_archivio.dbo.log_framework)
ORDER BY log_data 


delete from log_framework where log_id in
(SELECT log_id FROM infogross_archivio.dbo.log_framework)
-->
			<%
			
			dim conn, Aconn, rs, Ars, sql, field, contatore
			
			set conn = server.CreateObject("ADODB.Connection")
			set Aconn = server.CreateObject("ADODB.Connection")
			
			set rs = Server.CreateObject("ADODB.Recordset")
			set Ars = Server.CreateObject("ADODB.Recordset")
			
			conn.Open Application("DATA_ConnectionString"), "", ""
			Aconn.Open Application("DATA_ARCHIVE_ConnectionString"), "", "" 
			
			conn.commandtimeout = 120
			aconn.commandtimeout = 120
			%>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
				<caption class="border">
					<table cellspacing="0" cellpadding="0" align="right">
						<tr>
							<td style="font-size: 1px;">
								<a class="button" href="javascript:void(0);" onclick="window.close();">
									CHIUDI
								</a>
							</td>
						</tr>
					</table>
					ARCHIVIAZIONE LOG_FRAMEWORK
				</caption>
				<% 
				'********************************************************************************
				conn.BeginTrans
				Aconn.BeginTrans
				contatore = 300
				'********************************************************************************
				
				dim data, id_lista_delete
				id_lista_delete = "0"
				data = Now()
				data = DateAdd("m",-1, data)
				
				sql = " SELECT TOP 301 *, " & vbCrLF &  _
					  " (SELECT admin_nome + ' ' + admin_cognome FROM tb_admin WHERE ID_ADMIN = log_admin_id) AS log_admin_name, " & vbCrLF &  _
					  " (SELECT NomeElencoIndirizzi + ' ' + CognomeElencoIndirizzi + ' - ' + NomeOrganizzazioneElencoIndirizzi FROM tb_Indirizzario " & vbCrLF &  _
					  " WHERE IDElencoIndirizzi IN (SELECT ut_NextCom_ID FROM tb_Utenti WHERE ut_ID = log_user_id)) AS log_user_name, " & vbCrLF &  _
					  " (SELECT sito_nome FROM tb_siti WHERE id_sito = log_application_id) AS log_application_name " & vbCrLF &  _
					  " FROM log_framework " & vbCrLF &  _
					  " WHERE log_data < " & SQL_date(conn, data) & vbCrLF &  _
					  " ORDER BY log_data "
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
				<tr>
					<td class="label_no_width" colspan="5">TROVATI n&ordm; <%= rs.recordcount %> LOG DA ARCHIVIARE</td>
				</tr>
				<tr>
					<th class="center" style="width:5%;">ID</th>
					<th style="width:10%;">DATA</th>
					<th style="width:13%;">CODICE</th>
					<th style="width:60%;">DESCRIZIONE</th>
					<th style="width:12%;">ESITO</th>
				</tr>
			</table>
			<% while (contatore>0) and not rs.eof 
				if rs.AbsolutePosition MOD 100 = 0 then
'****************************************************************************************************************************************************************
					'restituisce al browser client il risultato
					response.flush()
'****************************************************************************************************************************************************************
				end if%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
					<tr>
						<td style="width:5%;" class="content_center"><%= rs("log_id") %></td>
						<td style="width:10%;" class="content_center"><%= DateTimeIta(rs("log_data")) %></td>
						<td style="width:13%;" class="content"><%= rs("log_codice") %></td>
						<td style="width:60%;" class="content"><%= rs("log_descrizione") %></td>
						<%'apre recordset da archiviare
						sql = "SELECT * FROM log_framework WHERE log_id=" & rs("log_id")
						Ars.open sql, Aconn, adOpenStatic, adLockOptimistic, adCmdText
						if Ars.eof then
							Ars.addNew
							for each field in rs.fields

								Ars(field.name) = rs(field.name)
							next
							'Ars("log_admin_name") = rs("log_admin_name")
							id_lista_delete = id_lista_delete & "," & rs("log_id")
							Ars.update
							
							%>
							<td style="width:12%;" class="content_center ok">ARCHIVIATA</td>
						<% else %>
							<td style="width:12%;" class="content_center error" title="log gia' presente nel DB di archivio: verificare correttezza.">ERRORE</td>
						<% end if
						Ars.close %>
					</tr>
				</table>
				<% 
				contatore = contatore -1
				rs.movenext
			wend
			if not rs.eof then %>
				<script language="JavaScript" type="text/javascript">
					document.location.reload(true);
				</script>
			<% end if
			rs.close
			
			sql = "DELETE FROM log_framework WHERE log_data < " & SQL_date(conn, data) & " AND log_id IN ("&id_lista_delete&")"
			CALL conn.execute(sql, , adExecuteNoRecords)
			
			'********************************************************************************
			conn.CommitTrans
			Aconn.CommitTrans
			'********************************************************************************
			
			conn.close
			Aconn.close
			set rs = nothing
			set Ars = nothing
			set conn = nothing
			set Aconn = nothing
			%>
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<tr>
					<td class="footer" colspan="2">
						<a class="button" href="javascript:void(0);" onclick="window.close();">
							CHIUDI
						</a>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<br>
<br>
</body>
</html>