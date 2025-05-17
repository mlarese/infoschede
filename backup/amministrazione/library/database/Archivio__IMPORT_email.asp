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
'/amministrazione/library/database/archivio__import_email.asp?STAND_ALONE=1
%>
<html>
<head>
	<title>Archiviazione email</title>
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
					ARCHIVIAZIONE EMAIL
				</caption>
				<% 
				'********************************************************************************
				conn.BeginTrans
				Aconn.BeginTrans
				contatore = 200
				'********************************************************************************
				
				sql = " SELECT * FROM tb_email"& _
					  " WHERE NOT(" & SQL_IsTrue(conn, "email_archiviata") & ") " + _
					  " AND NOT " + SQL_IsTrue(conn, "email_isBozza") + _
					  " ORDER BY email_data "
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
				<tr>
					<td class="label_no_width" colspan="4">TROVATE n&ordm; <%= rs.recordcount %> EMAIL DA ARCHIVIARE</td>
				</tr>
				<tr>
					<th class="center" style="width:5%;">ID</th>
					<th style="width:13%;">DATA</th>
					<th style="width:70%;">OGGETTO</th>
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
						<td style="width:5%;" class="content_center"><%= rs("email_id") %></td>
						<td style="width:13%;" class="content"><%= DateTimeIta(rs("email_data")) %></td>
						<td style="width:70%;" class="content"><%= rs("email_object") %></td>
						<%'apre recordset da archiviare
						sql = "SELECT * FROM tb_email WHERE email_id=" & rs("email_id")
						Ars.open sql, Aconn, adOpenStatic, adLockOptimistic, adCmdText
						if Ars.eof then
							Ars.addNew
							for each field in rs.fields
								Ars(field.name) = rs(field.name)
							next
							rs("email_archiviata") = true
							rs("email_archiviata_il") = Now()
							Ars.update
						
							'aggiorna stato dell'email
							rs("email_archiviata") = true
							rs("email_archiviata_il") = Now()
							rs("email_text") = NULL
							rs("email_name_database") = cString(Aconn.DefaultDatabase)
							
							rs.update %>
							<td style="width:12%;" class="content_center ok">ARCHIVIATA</td>
						<% else %>
							<td style="width:12%;" class="content_center error" title="email gia' presente nello storico: verificare correttezza.">ERRORE</td>
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