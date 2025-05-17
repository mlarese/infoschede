<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
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
'/amministrazione/library/database/CompactDatabase.asp?STAND_ALONE=1&CONFERMA=1&ConnString=DATA_ConnectionString

%>
<html>
<head>
	<title>Amministrazione aggiornamenti database</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();">
<%
dim conn
set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application(request("ConnString")), "", ""
%>
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
	<form action="" method="post" id="form1" name="form1">
	<tr>
		<td style="padding-top:4px;">
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<caption>connessione "<%= request("ConnString") %>"</caption>
				<tr>
					<th>COMPRESSIONE DATABASE</th>
				</tr>
				<tr>
					<td class="content">
						Dimensione attuale: <%=DatabaseSize(conn)%>
					</td>
				</tr>
				<% if request("conferma")="" then %>
					<tr>
						<td class="content">
							<table cellpadding="4" cellspacing="0" width="100%">
								<tr>
									<td class="label" colspan="2">Conferma l'operazione?</td>
								</tr>
								<tr>
									<td class="content_center" width="50%">
										<input type="submit" class="button" name="conferma" value="CONFERMA" tabindex="1" id="primo_elemento">
									</td>
									<td class="content_center">
										<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close()" tabindex="2">
									</td>
								</tr>
							</table>
						</td>
					</tr>
				<% else 
					'compatta il database
					CALL CompactDatabase(conn) %>
					<tr>
						<td class="content_b">
							Compressione database eseguita correttamente.
						</td>
					</tr>
					<tr>
						<td class="content">
							Dimensione file dopo compressione: <%=DatabaseSize(conn)%>
						</td>
					</tr>
				<% end if %>
				<tr>
					<td class="footer" colspan="2">
						<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
					</td>
				</tr>
			</form>
			</table>
		</td>
	</tr>
</table>
</form>
</body>
</html>

<% 
conn.close
set conn = nothing


if request("STAND_ALONE") = "" then %>
	<script language="JavaScript" type="text/javascript">
	<!--
		FitWindowSize(this);
		PageOnLoad_FocusSet();
	//-->
	</script>
<% end if %>